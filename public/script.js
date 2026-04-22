const socket = io({
    transports: ['websocket', 'polling'],
    reconnection: true,
    reconnectionDelay: 1000,
    reconnectionDelayMax: 5000,
    timeout: 10000
});

const PANEL_POLL_INTERVAL_MS = 2500;

const state = {
    connectionStatus: 'disconnected',
    qrImage: null,
    automation: null,
    socketConnected: false,
    startRequestPending: false,
    stopRequestPending: false
};

const statusEl = document.getElementById('connection-status');
const statusDot = document.querySelector('.status-dot');
const waChip = document.getElementById('wa-chip');
const automationChip = document.getElementById('automation-chip');
const automationStatusLabel = document.getElementById('automation-status-label');
const automationStageEl = document.getElementById('automation-stage');
const automationMessageEl = document.getElementById('automation-message');
const automationRunIdEl = document.getElementById('automation-run-id');
const lastUpdatedEl = document.getElementById('last-updated');
const syncStatusEl = document.getElementById('panel-sync-status');
const qrImageEl = document.getElementById('qr-image');
const qrPlaceholderEl = document.getElementById('qr-placeholder');
const qrCaptionEl = document.getElementById('qr-caption');
const logListEl = document.getElementById('automation-log-list');
const artifactsListEl = document.getElementById('artifacts-list');
const sendTestResultEl = document.getElementById('send-test-result');
const clearEventsButton = document.getElementById('btn-clear-events');
const clearEventsResultEl = document.getElementById('clear-events-result');
const startProcessButton = document.getElementById('btn-start-process');
const stopProcessButton = document.getElementById('btn-stop-process');
const startProcessResultEl = document.getElementById('start-process-result');
const processUserEl = document.getElementById('process-user');
const processPasswordEl = document.getElementById('process-password');

const countMsg = document.getElementById('count-msg');
const countFiles = document.getElementById('count-files');
const countCleaned = document.getElementById('count-cleaned');
const countPending = document.getElementById('count-pending');
const countClients = document.getElementById('count-clients');
const countReps = document.getElementById('count-reps');

function formatTimestamp(value) {
    if (!value) return '-';

    try {
        return new Intl.DateTimeFormat('pt-BR', {
            dateStyle: 'short',
            timeStyle: 'medium'
        }).format(new Date(value));
    } catch (error) {
        return value;
    }
}

function statusLabelFromCode(status) {
    const labels = {
        idle: 'Idle',
        starting: 'Inicializando',
        running: 'Em andamento',
        completed: 'Concluido',
        stopped: 'Interrompido',
        error: 'Erro'
    };

    return labels[status] || 'Idle';
}

function isAutomationBusy(status = state.automation?.status) {
    return ['starting', 'running'].includes(status);
}

function updateProcessControls() {
    if (!startProcessButton || !stopProcessButton) return;

    if (state.startRequestPending) {
        startProcessButton.disabled = true;
        stopProcessButton.disabled = true;
        startProcessButton.textContent = 'Iniciando...';
        stopProcessButton.textContent = 'Parar processo';
        return;
    }

    if (state.stopRequestPending) {
        startProcessButton.disabled = true;
        stopProcessButton.disabled = true;
        startProcessButton.textContent = 'Processo em andamento';
        stopProcessButton.textContent = 'Parando...';
        return;
    }

    if (isAutomationBusy()) {
        startProcessButton.disabled = true;
        startProcessButton.textContent = 'Processo em andamento';
        stopProcessButton.disabled = false;
        stopProcessButton.textContent = 'Parar processo';
        return;
    }

    startProcessButton.disabled = false;
    startProcessButton.textContent = 'Iniciar processo';
    stopProcessButton.disabled = true;
    stopProcessButton.textContent = 'Parar processo';
}

function updateDocumentTitle() {
    const stage = state.automation?.stage || 'Painel de Operacao';
    const prefix = state.socketConnected ? 'Ao vivo' : 'Sincronizando';
    document.title = `${prefix} | ${stage}`;
}

function setSyncStatus(label) {
    syncStatusEl.textContent = label;
    updateDocumentTitle();
}

function updateConnectionStatus() {
    const connected = state.connectionStatus === 'connected';
    statusEl.textContent = connected ? 'Conectado' : 'Desconectado';
    statusDot.classList.toggle('online', connected);
    waChip.classList.toggle('is-online', connected);

    if (connected) {
        state.qrImage = null;
        qrPlaceholderEl.style.display = 'none';
        qrImageEl.style.display = 'none';
        qrCaptionEl.textContent = 'WhatsApp conectado e pronto para o envio automatico.';
        return;
    }

    if (state.qrImage) {
        qrPlaceholderEl.style.display = 'none';
        qrImageEl.style.display = 'block';
        qrImageEl.src = state.qrImage;
        qrCaptionEl.textContent = 'Escaneie o QR Code para reconectar o WhatsApp do servidor.';
    } else {
        qrImageEl.style.display = 'none';
        qrPlaceholderEl.style.display = 'flex';
        qrCaptionEl.textContent = 'Aguardando um novo QR Code ou uma reconexao manual.';
    }
}

function renderMetrics(metrics = {}) {
    countMsg.textContent = metrics.messagesSent || 0;
    countFiles.textContent = metrics.filesSent || 0;
    countCleaned.textContent = metrics.cleanedFiles || 0;
    countPending.textContent = metrics.pendingCompanies || 0;
    countReps.textContent = metrics.representatives || 0;
    countClients.textContent = `${metrics.processedClients || 0} / ${metrics.totalClients || 0}`;
}

function renderLogs(logs = []) {
    if (!logs.length) {
        logListEl.innerHTML = '<div class="empty-state">Nenhum evento recebido ainda.</div>';
        return;
    }

    logListEl.innerHTML = logs.map((item) => `
        <article class="log-item log-${item.level || 'info'}">
            <div class="log-meta">
                <span class="log-stage">${item.stage || 'Automacao'}</span>
                <span class="log-time">${formatTimestamp(item.timestamp)}</span>
            </div>
            <p>${item.message}</p>
        </article>
    `).join('');
}

function renderArtifacts(artifacts = []) {
    if (!artifacts.length) {
        artifactsListEl.innerHTML = '<div class="empty-state">Nenhum arquivo gerado nesta execucao.</div>';
        return;
    }

    artifactsListEl.innerHTML = artifacts.map((item) => `
        <article class="artifact-item">
            <div>
                <span class="artifact-type">${item.type || 'arquivo'}</span>
                <strong>${item.label}</strong>
                <p>${item.path || ''}</p>
            </div>
            <span class="artifact-time">${formatTimestamp(item.timestamp)}</span>
        </article>
    `).join('');
}

function renderAutomationState(automation) {
    state.automation = automation || null;

    if (!automation) {
        renderMetrics();
        renderLogs();
        renderArtifacts();
        updateProcessControls();
        updateDocumentTitle();
        return;
    }

    automationStageEl.textContent = automation.stage || 'Aguardando automacao';
    automationMessageEl.textContent = automation.message || 'Nenhuma execucao iniciada.';
    automationRunIdEl.textContent = automation.runId || '-';
    automationStatusLabel.textContent = statusLabelFromCode(automation.status);
    lastUpdatedEl.textContent = formatTimestamp(automation.lastUpdatedAt);

    automationChip.className = `status-chip status-${automation.status || 'idle'}`;

    renderMetrics(automation.metrics || {});
    renderLogs(automation.logs || []);
    renderArtifacts(automation.artifacts || []);
    updateProcessControls();
    updateDocumentTitle();
}

function applyPanelState(panelState) {
    if (!panelState) return;

    if (typeof panelState.connectionStatus === 'string') {
        state.connectionStatus = panelState.connectionStatus;
    }

    if (Object.prototype.hasOwnProperty.call(panelState, 'qrCodeImage')) {
        state.qrImage = panelState.qrCodeImage || null;
    }

    updateConnectionStatus();

    if (panelState.automation) {
        renderAutomationState(panelState.automation);
    }
}

async function fetchPanelState(reason = 'polling') {
    try {
        const response = await fetch(`/panel-state?ts=${Date.now()}`, { cache: 'no-store' });
        if (!response.ok) throw new Error('Falha ao buscar o estado do painel');

        const data = await response.json();
        applyPanelState(data);

        if (state.socketConnected) {
            setSyncStatus('Tempo real ativo');
        } else if (reason === 'polling' || reason === 'startup') {
            setSyncStatus('Sincronizando por consulta');
        }
    } catch (error) {
        if (!state.socketConnected) {
            setSyncStatus('Aguardando servidor');
        }
    }
}

socket.on('connect', () => {
    state.socketConnected = true;
    setSyncStatus('Tempo real ativo');
    fetchPanelState('socket-connect');
});

socket.on('disconnect', () => {
    state.socketConnected = false;
    setSyncStatus('Reconectando, usando consulta');
    fetchPanelState('socket-disconnect');
});

socket.on('connect_error', () => {
    state.socketConnected = false;
    setSyncStatus('Falha no socket, usando consulta');
});

socket.io.on('reconnect_attempt', () => {
    if (!state.socketConnected) {
        setSyncStatus('Tentando reconectar');
    }
});

socket.on('status', (data) => {
    state.connectionStatus = data.status || 'disconnected';
    updateConnectionStatus();
});

socket.on('qr', (qrBase64) => {
    state.qrImage = qrBase64 || null;
    updateConnectionStatus();
});

socket.on('automation-state', (automation) => {
    renderAutomationState(automation);
});

socket.on('panel-state', (panelState) => {
    state.socketConnected = true;
    setSyncStatus('Tempo real ativo');
    applyPanelState(panelState);
});

document.getElementById('btn-reconnect').addEventListener('click', () => {
    socket.emit('reconnect');
});

document.getElementById('btn-logout').addEventListener('click', () => {
    if (confirm('Tem certeza que deseja desconectar?')) {
        socket.emit('logout');
    }
});

document.getElementById('btn-send-test').addEventListener('click', async () => {
    const number = document.getElementById('test-number').value.trim();
    const message = document.getElementById('test-message').value.trim();
    const button = document.getElementById('btn-send-test');

    if (!number || !message) {
        sendTestResultEl.textContent = 'Preencha numero e mensagem para testar.';
        sendTestResultEl.className = 'helper-text error';
        return;
    }

    button.disabled = true;
    button.textContent = 'Enviando...';
    sendTestResultEl.textContent = 'Enviando teste...';
    sendTestResultEl.className = 'helper-text';

    try {
        const response = await fetch('/send-message', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ number, message })
        });

        const data = await response.json().catch(() => ({ error: 'Resposta invalida do servidor' }));

        if (response.ok && data.success) {
            sendTestResultEl.textContent = 'Mensagem de teste enviada com sucesso.';
            sendTestResultEl.className = 'helper-text success';
            fetchPanelState('after-test');
        } else {
            sendTestResultEl.textContent = `Erro: ${data.error || 'Falha ao enviar mensagem'}`;
            sendTestResultEl.className = 'helper-text error';
        }
    } catch (error) {
        sendTestResultEl.textContent = 'Erro ao conectar com o servidor.';
        sendTestResultEl.className = 'helper-text error';
    } finally {
        button.disabled = false;
        button.textContent = 'Enviar teste';
    }
});

startProcessButton.addEventListener('click', async () => {
    const usuario = processUserEl.value.trim();
    const senha = processPasswordEl.value;

    if (!usuario || !senha) {
        startProcessResultEl.textContent = 'Preencha usuario e senha para iniciar o processo.';
        startProcessResultEl.className = 'helper-text error';
        return;
    }

    if (isAutomationBusy()) {
        startProcessResultEl.textContent = 'Ja existe uma automacao em andamento neste momento.';
        startProcessResultEl.className = 'helper-text error';
        return;
    }

    state.startRequestPending = true;
    updateProcessControls();
    startProcessResultEl.textContent = 'Disparando automacao...';
    startProcessResultEl.className = 'helper-text';

    try {
        const response = await fetch('/automation/start', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ usuario, senha })
        });

        const data = await response.json().catch(() => ({ error: 'Resposta invalida do servidor' }));

        if (response.ok && data.success) {
            startProcessResultEl.textContent = 'Processo iniciado. O painel vai acompanhar a execucao em tempo real.';
            startProcessResultEl.className = 'helper-text success';
            processPasswordEl.value = '';
            fetchPanelState('after-start');
        } else {
            startProcessResultEl.textContent = `Erro: ${data.error || 'Falha ao iniciar o processo'}`;
            startProcessResultEl.className = 'helper-text error';
        }
    } catch (error) {
        startProcessResultEl.textContent = 'Erro ao conectar com o servidor para iniciar o processo.';
        startProcessResultEl.className = 'helper-text error';
    } finally {
        state.startRequestPending = false;
        updateProcessControls();
    }
});

stopProcessButton.addEventListener('click', async () => {
    if (!isAutomationBusy()) {
        startProcessResultEl.textContent = 'Nenhuma automacao em andamento para parar.';
        startProcessResultEl.className = 'helper-text error';
        return;
    }

    if (!confirm('Deseja interromper o processo atual?')) {
        return;
    }

    state.stopRequestPending = true;
    updateProcessControls();
    startProcessResultEl.textContent = 'Solicitando parada da automacao...';
    startProcessResultEl.className = 'helper-text';

    try {
        const response = await fetch('/automation/stop', {
            method: 'POST'
        });

        const data = await response.json().catch(() => ({ error: 'Resposta invalida do servidor' }));

        if (response.ok && data.success) {
            startProcessResultEl.textContent = 'Parada solicitada. Aguarde a confirmacao no painel.';
            startProcessResultEl.className = 'helper-text success';
            fetchPanelState('after-stop');
        } else {
            startProcessResultEl.textContent = `Erro: ${data.error || 'Falha ao parar o processo'}`;
            startProcessResultEl.className = 'helper-text error';
        }
    } catch (error) {
        startProcessResultEl.textContent = 'Erro ao conectar com o servidor para parar o processo.';
        startProcessResultEl.className = 'helper-text error';
    } finally {
        state.stopRequestPending = false;
        updateProcessControls();
    }
});

clearEventsButton.addEventListener('click', async () => {
    if (!confirm('Deseja limpar os eventos exibidos no painel?')) {
        return;
    }

    clearEventsButton.disabled = true;
    clearEventsButton.textContent = 'Limpando...';
    clearEventsResultEl.textContent = 'Limpando eventos do painel...';
    clearEventsResultEl.className = 'helper-text card-helper';

    try {
        const response = await fetch('/automation/clear-events', {
            method: 'POST'
        });

        const data = await response.json().catch(() => ({ error: 'Resposta invalida do servidor' }));

        if (response.ok && data.success) {
            clearEventsResultEl.textContent = `${data.cleared || 0} evento(s) removido(s) do painel.`;
            clearEventsResultEl.className = 'helper-text card-helper success';
            renderLogs([]);
            fetchPanelState('after-clear-events');
        } else {
            clearEventsResultEl.textContent = `Erro: ${data.error || 'Falha ao limpar os eventos'}`;
            clearEventsResultEl.className = 'helper-text card-helper error';
        }
    } catch (error) {
        clearEventsResultEl.textContent = 'Erro ao conectar com o servidor para limpar os eventos.';
        clearEventsResultEl.className = 'helper-text card-helper error';
    } finally {
        clearEventsButton.disabled = false;
        clearEventsButton.textContent = 'Limpar eventos';
    }
});

updateConnectionStatus();
renderAutomationState(null);
updateProcessControls();
setSyncStatus('Conectando...');
fetchPanelState('startup');
setInterval(() => fetchPanelState('polling'), PANEL_POLL_INTERVAL_MS);
window.addEventListener('focus', () => fetchPanelState('focus'));
document.addEventListener('visibilitychange', () => {
    if (!document.hidden) {
        fetchPanelState('visible');
    }
});
