const {
    default: makeWASocket,
    useMultiFileAuthState,
    DisconnectReason,
    fetchLatestBaileysVersion
} = require('@whiskeysockets/baileys');
const express = require('express');
const fs = require('fs');
const http = require('http');
const multer = require('multer');
const path = require('path');
const pino = require('pino');
const QRCode = require('qrcode');
const qrcodeTerminal = require('qrcode-terminal');
const { Server } = require('socket.io');
const { spawn } = require('child_process');

const app = express();
const server = http.createServer(app);
const io = new Server(server);
const port = 2602;
const SEND_TIMEOUT_MS = 20000;
const RECONNECT_DELAY_MS = 3000;
const MAX_AUTOMATION_LOGS = 180;
const MAX_AUTOMATION_ARTIFACTS = 30;
const AUTOMATION_SCRIPT_PATH = path.join(__dirname, 'baixaRel.py');
const UPLOADS_DIR = path.join(__dirname, 'uploads');
const PUBLIC_DIR = path.join(__dirname, 'public');
const AUTH_INFO_DIR = path.join(__dirname, 'auth_info_baileys');
const PYTHON_EXECUTABLE = fs.existsSync(path.join(__dirname, '.venv', 'Scripts', 'python.exe'))
    ? path.join(__dirname, '.venv', 'Scripts', 'python.exe')
    : 'python';

const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const dir = UPLOADS_DIR;
        if (!fs.existsSync(dir)) fs.mkdirSync(dir);
        cb(null, dir);
    },
    filename: (req, file, cb) => {
        cb(null, `${Date.now()}-${file.originalname}`);
    }
});

const upload = multer({ storage });

app.use(express.json({ limit: '1mb' }));
app.use(express.static(PUBLIC_DIR));

let sock;
let connected = false;
let qrCodeImage = null;
let isConnecting = false;
let reconnectTimer = null;
let logoutRequested = false;
let automationProcess = null;
let automationStopRequested = false;

function createAutomationState(overrides = {}) {
    return {
        runId: null,
        status: 'idle',
        stage: 'Aguardando automacao',
        message: 'Nenhuma execucao iniciada.',
        startedAt: null,
        finishedAt: null,
        lastUpdatedAt: new Date().toISOString(),
        metrics: {
            messagesSent: 0,
            filesSent: 0,
            cleanedFiles: 0,
            totalClients: 0,
            processedClients: 0,
            pendingCompanies: 0,
            representatives: 0
        },
        artifacts: [],
        logs: [],
        ...overrides
    };
}

let automationState = createAutomationState();

function getPanelState() {
    return {
        connectionStatus: connected ? 'connected' : 'disconnected',
        qrCodeImage,
        automation: automationState,
        serverTime: new Date().toISOString()
    };
}

function emitPanelState(target = io) {
    target.emit('panel-state', getPanelState());
}

function emitConnectionStatus(status) {
    connected = status === 'connected';
    io.emit('status', { status });
    emitPanelState();
}

function emitAutomationState() {
    io.emit('automation-state', automationState);
    emitPanelState();
}

function isSocketReady() {
    return Boolean(sock && connected && sock.ws?.isOpen);
}

function isAutomationProcessRunning() {
    return Boolean(automationProcess && automationProcess.exitCode === null && !automationProcess.killed);
}

function isAutomationBusy() {
    return isAutomationProcessRunning() || ['starting', 'running'].includes(automationState.status);
}

function stopAutomationProcess() {
    const child = automationProcess;

    if (!child || child.exitCode !== null || child.killed) {
        return false;
    }

    automationStopRequested = true;

    if (process.platform === 'win32' && child.pid) {
        const killer = spawn('taskkill', ['/pid', String(child.pid), '/T', '/F'], {
            windowsHide: true,
            stdio: 'ignore'
        });

        killer.on('error', (error) => {
            console.error('Erro ao encerrar processo com taskkill:', error.message);
            try {
                child.kill();
            } catch (killError) {
                console.error('Erro no fallback de encerramento:', killError.message);
            }
        });

        return true;
    }

    try {
        return child.kill('SIGTERM');
    } catch (error) {
        console.error('Erro ao encerrar processo:', error.message);
        return false;
    }
}

function withTimeout(promise, timeoutMs, errorMessage) {
    let timer;

    return Promise.race([
        promise,
        new Promise((_, reject) => {
            timer = setTimeout(() => reject(new Error(errorMessage)), timeoutMs);
        })
    ]).finally(() => clearTimeout(timer));
}

function sanitizePhoneNumber(number) {
    const digits = String(number || '').replace(/\D/g, '');

    if (digits.length < 12 || digits.length > 15) {
        throw new Error('Numero invalido. Use DDI + DDD + numero.');
    }

    return digits;
}

function mergeAutomationMetrics(metrics = {}) {
    for (const [key, value] of Object.entries(metrics)) {
        const parsedValue = Number(value);
        if (Number.isFinite(parsedValue)) {
            automationState.metrics[key] = parsedValue;
        }
    }
}

function incrementAutomationMetrics(delta = {}) {
    for (const [key, value] of Object.entries(delta)) {
        const parsedValue = Number(value);
        if (!Number.isFinite(parsedValue)) continue;

        const currentValue = Number(automationState.metrics[key] || 0);
        automationState.metrics[key] = currentValue + parsedValue;
    }
}

function addAutomationLog({ message, level = 'info', stage }) {
    if (!message) return;

    automationState.logs.unshift({
        id: `${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
        message,
        level,
        stage: stage || automationState.stage,
        timestamp: new Date().toISOString()
    });

    automationState.logs = automationState.logs.slice(0, MAX_AUTOMATION_LOGS);
}

function addAutomationArtifact(artifact) {
    if (!artifact?.label) return;

    automationState.artifacts.unshift({
        id: `${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
        type: artifact.type || 'file',
        label: artifact.label,
        path: artifact.path || '',
        timestamp: new Date().toISOString()
    });

    automationState.artifacts = automationState.artifacts.slice(0, MAX_AUTOMATION_ARTIFACTS);
}

function resetAutomationState(payload = {}) {
    automationStopRequested = false;
    automationState = createAutomationState({
        runId: payload.runId || null,
        status: payload.status || 'starting',
        stage: payload.stage || 'Preparacao',
        message: payload.message || 'Nova execucao iniciada.',
        startedAt: payload.startedAt || new Date().toISOString(),
        finishedAt: null,
        lastUpdatedAt: new Date().toISOString()
    });

    if (payload.metrics) {
        mergeAutomationMetrics(payload.metrics);
    }

    addAutomationLog({
        message: automationState.message,
        level: payload.level || 'info',
        stage: automationState.stage
    });

    emitAutomationState();
}

function applyAutomationEvent(payload = {}) {
    if (payload.runId) automationState.runId = payload.runId;
    if (payload.status) automationState.status = payload.status;
    if (payload.stage) automationState.stage = payload.stage;
    if (payload.message) automationState.message = payload.message;
    if (payload.metrics) mergeAutomationMetrics(payload.metrics);
    if (payload.incrementMetrics) incrementAutomationMetrics(payload.incrementMetrics);
    if (payload.artifact) addAutomationArtifact(payload.artifact);

    automationState.lastUpdatedAt = new Date().toISOString();

    if (automationState.status === 'completed' || automationState.status === 'error') {
        automationState.finishedAt = automationState.lastUpdatedAt;
    } else if (!automationState.startedAt) {
        automationState.startedAt = automationState.lastUpdatedAt;
    }

    if (payload.message && payload.log !== false) {
        addAutomationLog({
            message: payload.message,
            level: payload.level || 'info',
            stage: payload.stage || automationState.stage
        });
    }

    emitAutomationState();
}

async function resolveWhatsAppJid(number) {
    const sanitizedNumber = sanitizePhoneNumber(number);
    const results = await withTimeout(
        sock.onWhatsApp(sanitizedNumber),
        SEND_TIMEOUT_MS,
        'Tempo limite ao validar o numero no WhatsApp.'
    );

    const [contact] = results || [];
    if (!contact?.exists || !contact?.jid) {
        throw new Error('Numero nao encontrado no WhatsApp.');
    }

    return contact.jid;
}

function scheduleReconnect(delayMs = RECONNECT_DELAY_MS) {
    if (reconnectTimer) return;

    reconnectTimer = setTimeout(() => {
        reconnectTimer = null;
        connectToWhatsApp().catch((error) => {
            console.error('Erro ao reconectar:', error.message);
            scheduleReconnect();
        });
    }, delayMs);
}

function clearAuthFolder() {
    if (fs.existsSync(AUTH_INFO_DIR)) {
        fs.rmSync(AUTH_INFO_DIR, { recursive: true, force: true });
    }
}

function pipeAutomationOutput(stream, logger) {
    if (!stream) return;

    let buffer = '';
    stream.setEncoding('utf8');

    stream.on('data', (chunk) => {
        buffer += chunk;
        const lines = buffer.split(/\r?\n/);
        buffer = lines.pop() || '';

        for (const line of lines) {
            const trimmedLine = line.trim();
            if (trimmedLine) {
                logger(trimmedLine);
            }
        }
    });

    stream.on('end', () => {
        const trimmedLine = buffer.trim();
        if (trimmedLine) {
            logger(trimmedLine);
        }
    });
}

function startAutomationProcess({ usuario, senha }) {
    automationStopRequested = false;
    const child = spawn(PYTHON_EXECUTABLE, ['-u', AUTOMATION_SCRIPT_PATH], {
        cwd: __dirname,
        env: {
            ...process.env,
            AUTOMACAO_SERVER_URL: process.env.AUTOMACAO_SERVER_URL || `http://localhost:${port}`,
            AUTOMACAO_USUARIO: usuario,
            AUTOMACAO_SENHA: senha
        },
        windowsHide: true,
        stdio: ['ignore', 'pipe', 'pipe']
    });

    automationProcess = child;

    pipeAutomationOutput(child.stdout, (line) => {
        console.log(`[automacao] ${line}`);
    });

    pipeAutomationOutput(child.stderr, (line) => {
        console.error(`[automacao][erro] ${line}`);
    });

    child.on('error', (error) => {
        if (automationProcess === child) {
            automationProcess = null;
        }

        console.error('Erro ao iniciar automacao:', error.message);
        applyAutomationEvent({
            status: 'error',
            stage: 'Inicializacao',
            message: `Falha ao iniciar o processo: ${error.message}`,
            level: 'error'
        });
    });

    child.on('close', (code, signal) => {
        if (automationProcess === child) {
            automationProcess = null;
        }

        const exitDescription = signal ? `sinal ${signal}` : `codigo ${code}`;
        console.log(`Processo de automacao finalizado com ${exitDescription}.`);

        if (automationStopRequested) {
            automationStopRequested = false;
            applyAutomationEvent({
                status: 'stopped',
                stage: 'Interrompido',
                message: 'Processo interrompido manualmente.',
                level: 'warning'
            });
        } else if (!['completed', 'error', 'stopped'].includes(automationState.status)) {
            const success = code === 0;
            applyAutomationEvent({
                status: success ? 'completed' : 'error',
                stage: success ? 'Concluido' : 'Falha',
                message: success
                    ? 'Processo finalizado.'
                    : `Processo encerrado inesperadamente (${exitDescription}).`,
                level: success ? 'success' : 'error'
            });
        } else {
            emitPanelState();
        }
    });

    return child;
}

async function connectToWhatsApp() {
    if (isConnecting) return;
    isConnecting = true;

    try {
        const { state, saveCreds } = await useMultiFileAuthState(AUTH_INFO_DIR);
        const { version, isLatest } = await fetchLatestBaileysVersion();
        console.log(`Usando versao v${version.join('.')}, isLatest: ${isLatest}`);

        const newSock = makeWASocket({
            version,
            auth: state,
            logger: pino({ level: 'silent' }),
            browser: ['Ubuntu', 'Chrome', '110.0.5563.147'],
            syncFullHistory: false,
            markOnlineOnConnect: true,
            defaultQueryTimeoutMs: SEND_TIMEOUT_MS
        });

        sock = newSock;
        newSock.ev.on('creds.update', saveCreds);

        newSock.ev.on('connection.update', async (update) => {
            if (sock !== newSock) return;

            const { connection, lastDisconnect, qr } = update;

            if (qr) {
                qrCodeImage = await QRCode.toDataURL(qr);
                io.emit('qr', qrCodeImage);
                emitPanelState();
                console.log('Novo QR Code gerado.');
                qrcodeTerminal.generate(qr, { small: true });
            }

            if (connection === 'close') {
                const statusCode = lastDisconnect?.error?.output?.statusCode;
                const shouldReconnect = statusCode !== DisconnectReason.loggedOut;

                emitConnectionStatus('disconnected');
                console.log(`Conexao fechada. Codigo: ${statusCode ?? 'desconhecido'}`);

                if (shouldReconnect) {
                    scheduleReconnect();
                } else if (logoutRequested || !fs.existsSync(AUTH_INFO_DIR)) {
                    clearAuthFolder();
                    scheduleReconnect(1000);
                }

                logoutRequested = false;
            } else if (connection === 'open') {
                qrCodeImage = null;
                logoutRequested = false;
                emitConnectionStatus('connected');
                console.log('Conexao estabelecida.');
            }
        });
    } finally {
        isConnecting = false;
    }
}

io.on('connection', (socket) => {
    socket.emit('status', { status: connected ? 'connected' : 'disconnected' });
    socket.emit('automation-state', automationState);
    if (qrCodeImage) socket.emit('qr', qrCodeImage);
    emitPanelState(socket);

    socket.on('reconnect', () => {
        if (sock?.ws?.isOpen) {
            sock.end(new Error('Reconexao solicitada manualmente'));
        } else {
            logoutRequested = true;
            clearAuthFolder();
            scheduleReconnect(0);
        }
    });

    socket.on('logout', async () => {
        if (sock) {
            logoutRequested = true;
            await sock.logout();
        }
    });
});

app.get('/panel-state', (req, res) => {
    res.json(getPanelState());
});

app.get('/automation-state', (req, res) => {
    res.json(automationState);
});

app.post('/automation/reset', (req, res) => {
    resetAutomationState(req.body || {});
    res.json({ success: true });
});

app.post('/automation/event', (req, res) => {
    applyAutomationEvent(req.body || {});
    res.json({ success: true });
});

app.post('/automation/clear-events', (req, res) => {
    const cleared = automationState.logs.length;
    automationState.logs = [];
    emitAutomationState();
    res.json({ success: true, cleared });
});

app.post('/automation/start', (req, res) => {
    const usuario = String(req.body?.usuario || '').trim();
    const senha = String(req.body?.senha || '');

    if (!usuario) {
        return res.status(400).json({ error: 'Usuario obrigatorio' });
    }

    if (!senha) {
        return res.status(400).json({ error: 'Senha obrigatoria' });
    }

    if (!fs.existsSync(AUTOMATION_SCRIPT_PATH)) {
        return res.status(500).json({ error: 'Script da automacao nao encontrado' });
    }

    if (isAutomationBusy()) {
        return res.status(409).json({ error: 'Ja existe uma automacao em andamento' });
    }

    resetAutomationState({
        runId: `web-${Date.now()}`,
        status: 'starting',
        stage: 'Inicializacao',
        message: 'Processo solicitado pelo painel. Preparando automacao...'
    });

    try {
        startAutomationProcess({ usuario, senha });
        return res.status(202).json({ success: true });
    } catch (error) {
        applyAutomationEvent({
            status: 'error',
            stage: 'Inicializacao',
            message: `Falha ao iniciar o processo: ${error.message}`,
            level: 'error'
        });
        return res.status(500).json({ error: error.message || 'Falha ao iniciar o processo' });
    }
});

app.post('/automation/stop', (req, res) => {
    if (!isAutomationProcessRunning() || !['starting', 'running'].includes(automationState.status)) {
        return res.status(409).json({ error: 'Nenhuma automacao em andamento' });
    }

    if (!stopAutomationProcess()) {
        return res.status(500).json({ error: 'Nao foi possivel interromper o processo' });
    }

    applyAutomationEvent({
        stage: 'Interrompendo',
        message: 'Solicitacao de parada enviada. Encerrando automacao...',
        level: 'warning'
    });

    return res.status(202).json({ success: true });
});

app.post('/send-message', async (req, res) => {
    const { number, message } = req.body;

    if (!message || !String(message).trim()) {
        return res.status(400).json({ error: 'Mensagem obrigatoria' });
    }

    if (!isSocketReady()) {
        emitConnectionStatus('disconnected');
        return res.status(503).json({ error: 'WhatsApp nao conectado' });
    }

    try {
        const jid = await resolveWhatsAppJid(number);
        await withTimeout(
            sock.sendMessage(jid, { text: String(message).trim() }),
            SEND_TIMEOUT_MS,
            'Tempo limite ao enviar a mensagem.'
        );

        applyAutomationEvent({
            message: `Mensagem enviada para ${sanitizePhoneNumber(number)}.`,
            stage: automationState.stage || 'Envio no WhatsApp',
            status: automationState.status === 'idle' ? 'running' : automationState.status,
            incrementMetrics: { messagesSent: 1 },
            level: 'success'
        });

        res.json({ success: true });
    } catch (error) {
        console.error('Erro ao enviar mensagem:', error.message);
        if (!sock?.ws?.isOpen) emitConnectionStatus('disconnected');
        res.status(500).json({ error: error.message || 'Erro interno ao enviar mensagem' });
    }
});

app.post('/send-media', upload.single('documento'), async (req, res) => {
    const { number, caption } = req.body;
    const file = req.file;

    if (!file) {
        return res.status(400).json({ error: 'Arquivo obrigatorio' });
    }

    if (!isSocketReady()) {
        emitConnectionStatus('disconnected');
        return res.status(503).json({ error: 'WhatsApp nao conectado' });
    }

    try {
        const jid = await resolveWhatsAppJid(number);
        await withTimeout(
            sock.sendMessage(jid, {
                document: fs.readFileSync(file.path),
                fileName: file.originalname,
                mimetype: file.mimetype,
                caption: caption || ''
            }),
            SEND_TIMEOUT_MS,
            'Tempo limite ao enviar o arquivo.'
        );

        applyAutomationEvent({
            message: `Arquivo enviado: ${file.originalname}.`,
            stage: automationState.stage || 'Envio no WhatsApp',
            status: automationState.status === 'idle' ? 'running' : automationState.status,
            incrementMetrics: { filesSent: 1 },
            level: 'success',
            artifact: {
                type: 'attachment',
                label: file.originalname,
                path: file.path
            }
        });

        res.json({ success: true });
    } catch (error) {
        console.error('Erro ao enviar arquivo:', error.message);
        if (!sock?.ws?.isOpen) emitConnectionStatus('disconnected');
        res.status(500).json({ error: error.message || 'Erro interno ao enviar arquivo' });
    } finally {
        if (file?.path && fs.existsSync(file.path)) {
            fs.unlinkSync(file.path);
        }
    }
});

server.listen(port, () => {
    console.log(`Dashboard rodando em http://localhost:${port}`);
    connectToWhatsApp().catch((error) => {
        console.error('Erro ao iniciar o WhatsApp:', error.message);
        scheduleReconnect();
    });
});
