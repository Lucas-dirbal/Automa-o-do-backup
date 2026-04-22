from glob import glob
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from urllib import request as urllib_request
import json
import msvcrt
import os
import pdfplumber
import re
import sys
import time
import unicodedata
import pandas as pd
from datetime import datetime


DATA_HOJE = datetime.now().strftime("%d/%m/%Y")
DIRETORIO_SCRIPT = os.path.dirname(os.path.abspath(__file__))
PASTA_RELATORIOS = os.path.join(DIRETORIO_SCRIPT, "Relatorios")
PASTA_RELATORIOS_LEGADO = os.path.join(
    os.path.expanduser("~"),
    "Desktop",
    "automacao 27",
    "Automacao backup",
    "Relatorios",
)
PASTA_RELATORIOS_LEGADO_ALT = os.path.join(
    os.path.expanduser("~"),
    "Desktop",
    "automacao 27",
    "Automa\u00e7\u00e3o backup",
    "Relatorios",
)
PASTA_PLANILHAS_SEPARADAS = os.path.join(DIRETORIO_SCRIPT, "representantes_separados")
PASTA_UPLOADS = os.path.join(DIRETORIO_SCRIPT, "uploads")
CAMINHO_MAPEAMENTO = os.path.join(DIRETORIO_SCRIPT, "mapeamento_representantes.xlsx")
URL_SISTEMA = "http://192.168.1.11:8078"
SERVIDOR_AUTOMACAO_URL = os.environ.get("AUTOMACAO_SERVER_URL", "http://192.168.1.11:2602").rstrip("/")
URL_MSG = f"{SERVIDOR_AUTOMACAO_URL}/send-message"
URL_MEDIA = f"{SERVIDOR_AUTOMACAO_URL}/send-media"
URL_AUTOMATION_RESET = f"{SERVIDOR_AUTOMACAO_URL}/automation/reset"
URL_AUTOMATION_EVENT = f"{SERVIDOR_AUTOMACAO_URL}/automation/event"
MENSAGEM_WHATSAPP = (
    "Olá, tudo bem? Me chamo Lucas e sou da Interface Sistemas. Estou entrando em contato para solicitar, por gentileza, "
    "que seja verificada, assim que possível, a planilha em anexo, referente "
    "às empresas que estão sem backup recente."
)
ROTULOS_SITUACAO_EMPRESA = (
    "situacao",
    "situacao cadastro",
    "status",
    "status cadastro",
)
TERMOS_EMPRESA_ATIVA = {
    "a",
    "at",
    "ativo",
    "ativa",
    "habilitado",
    "habilitada",
    "normal",
}
TERMOS_EMPRESA_INATIVA = {
    "i",
    "in",
    "inativo",
    "inativa",
    "cancelado",
    "cancelada",
    "encerrado",
    "encerrada",
    "baixado",
    "baixada",
    "desativado",
    "desativada",
    "bloqueado",
    "bloqueada",
    "suspenso",
    "suspensa",
}

EXECUCAO_ID = None


def garantir_pasta(caminho):
    if not os.path.exists(caminho):
        os.makedirs(caminho)
    return caminho


def normalizar_nome_arquivo(texto):
    return "".join(
        caractere for caractere in str(texto) if caractere.isalnum() or caractere in (" ", "_", "-")
    ).strip().replace(" ", "_")


def normalizar_texto_busca(texto):
    texto = unicodedata.normalize("NFD", str(texto or ""))
    texto = "".join(caractere for caractere in texto if unicodedata.category(caractere) != "Mn")
    return re.sub(r"\s+", " ", texto).strip().lower()


def postar_json_dashboard(url, payload, timeout=2):
    data = json.dumps(payload).encode("utf-8")
    requisicao = urllib_request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    try:
        with urllib_request.urlopen(requisicao, timeout=timeout) as resposta:
            resposta.read()
        return True
    except Exception:
        return False


def resetar_dashboard(arquivos_limpados):
    payload = {
        "runId": EXECUCAO_ID,
        "status": "starting",
        "stage": "Preparacao",
        "message": "Nova execucao iniciada.",
        "metrics": {
            "messagesSent": 0,
            "filesSent": 0,
            "cleanedFiles": arquivos_limpados,
            "totalClients": 0,
            "processedClients": 0,
            "pendingCompanies": 0,
            "representatives": 0,
        },
    }
    postar_json_dashboard(URL_AUTOMATION_RESET, payload)


def enviar_evento_dashboard(
    message=None,
    level="info",
    stage=None,
    status=None,
    metrics=None,
    increment_metrics=None,
    artifact=None,
    log=True,
):
    payload = {"runId": EXECUCAO_ID, "level": level, "log": log}

    if message is not None:
        payload["message"] = message
    if stage is not None:
        payload["stage"] = stage
    if status is not None:
        payload["status"] = status
    if metrics:
        payload["metrics"] = metrics
    if increment_metrics:
        payload["incrementMetrics"] = increment_metrics
    if artifact:
        payload["artifact"] = artifact

    postar_json_dashboard(URL_AUTOMATION_EVENT, payload)


def registrar_evento(
    mensagem,
    level="info",
    stage=None,
    status=None,
    metrics=None,
    increment_metrics=None,
    artifact=None,
    log=True,
):
    print(mensagem)
    enviar_evento_dashboard(
        message=mensagem,
        level=level,
        stage=stage,
        status=status,
        metrics=metrics,
        increment_metrics=increment_metrics,
        artifact=artifact,
        log=log,
    )


def limpar_arquivos_execucao_anterior():
    garantir_pasta(PASTA_RELATORIOS)
    garantir_pasta(PASTA_PLANILHAS_SEPARADAS)
    garantir_pasta(PASTA_UPLOADS)

    arquivos_removidos = []
    arquivos_vistos = set()
    alvos = [
        (PASTA_RELATORIOS, ("*.pdf", "*.xlsx")),
        (PASTA_RELATORIOS_LEGADO, ("*.pdf", "*.xlsx")),
        (PASTA_RELATORIOS_LEGADO_ALT, ("*.pdf", "*.xlsx")),
        (PASTA_PLANILHAS_SEPARADAS, ("*.xlsx",)),
        (PASTA_UPLOADS, ("*",)),
        (DIRETORIO_SCRIPT, ("PENDENCIAS*.xlsx",)),
    ]

    for pasta, padroes in alvos:
        if not os.path.isdir(pasta):
            continue

        for padrao in padroes:
            for caminho in glob(os.path.join(pasta, padrao)):
                if not os.path.isfile(caminho):
                    continue

                caminho_normalizado = os.path.normcase(os.path.abspath(caminho))
                if caminho_normalizado in arquivos_vistos:
                    continue

                try:
                    os.remove(caminho)
                    arquivos_vistos.add(caminho_normalizado)
                    arquivos_removidos.append(caminho)
                except PermissionError:
                    print(f"Arquivo em uso e nao removido: {caminho}")

    return arquivos_removidos


def localizar_pdf_baixado(pastas_monitoradas, arquivos_antes, inicio_download, timeout=40):
    registrar_evento("Aguardando download do PDF...", stage="Relatorio", status="running", log=False)

    for _ in range(timeout):
        time.sleep(1)

        for pasta in pastas_monitoradas:
            atuais = set(os.listdir(pasta))
            novos = []

            for nome in atuais - arquivos_antes.get(pasta, set()):
                nome_lower = nome.lower()
                if nome_lower.endswith(".pdf") and not nome_lower.endswith((".crdownload", ".tmp")):
                    novos.append(nome)

            if novos:
                novos.sort(key=lambda nome: os.path.getmtime(os.path.join(pasta, nome)))
                return pasta, novos[-1]

            recentes = []
            for nome in atuais:
                nome_lower = nome.lower()
                caminho = os.path.join(pasta, nome)
                if (
                    nome_lower.endswith(".pdf")
                    and os.path.isfile(caminho)
                    and os.path.getmtime(caminho) >= inicio_download - 2
                ):
                    recentes.append(nome)

            if recentes:
                recentes.sort(key=lambda nome: os.path.getmtime(os.path.join(pasta, nome)))
                return pasta, recentes[-1]

    return None, None


def digitar_senha(prompt="Senha: "):
    print(prompt, end="", flush=True)
    senha = ""
    while True:
        tecla = msvcrt.getch()
        if tecla in (b"\r", b"\n"):
            print()
            break
        elif tecla == b"\x08":
            if senha:
                senha = senha[:-1]
                sys.stdout.write("\b \b")
                sys.stdout.flush()
        elif tecla == b"\x03":
            raise KeyboardInterrupt
        else:
            senha += tecla.decode("utf-8", errors="ignore")
            sys.stdout.write("*")
            sys.stdout.flush()
    return senha


def obter_credenciais_iniciais():
    usuario_ambiente = os.environ.get("AUTOMACAO_USUARIO", "").strip()
    senha_ambiente = os.environ.get("AUTOMACAO_SENHA", "")

    if usuario_ambiente and senha_ambiente:
        registrar_evento(
            "Credenciais recebidas pelo painel web.",
            stage="Login",
            status="starting",
            log=False,
        )
        return usuario_ambiente, senha_ambiente, True

    print("Informe suas credenciais:")
    usuario = input("Usuario: ").strip()
    senha = digitar_senha("Senha: ")
    return usuario, senha, False


def fazer_login(driver, usuario, senha):
    wait = WebDriverWait(driver, 15)
    try:
        wait.until(EC.visibility_of_element_located((By.ID, "O18_id")))
        campo_usuario = wait.until(EC.element_to_be_clickable((By.ID, "O24_id-inputEl")))
        campo_usuario.clear()
        campo_usuario.send_keys(usuario)

        campo_senha = wait.until(EC.element_to_be_clickable((By.ID, "O20_id-inputEl")))
        campo_senha.clear()
        campo_senha.send_keys(senha)

        wait.until(EC.element_to_be_clickable((By.ID, "O34_id"))).click()
        time.sleep(2)

        msg_box_login = driver.find_elements(By.ID, "messagebox-1001")
        if msg_box_login and any(elemento.is_displayed() for elemento in msg_box_login):
            registrar_evento(
                "Mensagem de login invalido detectada (messagebox-1001).",
                level="warning",
                stage="Login",
                status="running",
            )
            return False

        msg_erro = driver.find_elements(By.XPATH, "//*[contains(text(), 'Login ou senha')]")
        if msg_erro and any(elemento.is_displayed() for elemento in msg_erro):
            registrar_evento("Login ou senha invalidos.", level="warning", stage="Login", status="running")
            return False

        registrar_evento("Login realizado com sucesso.", level="success", stage="Login", status="running")
        return True
    except Exception:
        return False


def salvar_codigos_clientes_extraidos(client_data, excel_path, origem):
    if not client_data:
        registrar_evento(
            f"Nenhum codigo de cliente encontrado na {origem}.",
            level="warning",
            stage="Extracao",
            status="error",
        )
        return False

    df = pd.DataFrame(client_data)
    total_encontrado = len(df)
    df = df.drop_duplicates(subset=["Codigo Cliente"], keep="first")
    total_unico = len(df)
    if "Assinatura Linha" in df.columns:
        df = df.drop(columns=["Assinatura Linha"])

    try:
        df.to_excel(excel_path, index=False)
    except PermissionError:
        excel_path = excel_path.replace(".xlsx", f"_{int(time.time())}.xlsx")
        df.to_excel(excel_path, index=False)

    registrar_evento(
        f"Planilha com codigos gerada: {excel_path}",
        level="success",
        stage="Extracao",
        status="running",
        artifact={
            "type": "xlsx",
            "label": os.path.basename(excel_path),
            "path": excel_path,
        },
    )
    registrar_evento(
        f"Resumo da extracao pela {origem}: {total_encontrado} codigos encontrados e {total_unico} unicos.",
        level="info",
        stage="Extracao",
        status="running",
    )
    return excel_path


def extrair_codigos_clientes(pdf_path, excel_path):
    registrar_evento("Extraindo codigos dos clientes do PDF...", stage="Extracao", status="running")
    try:
        pattern = r"(\d{2}/\d{2}/\d{4}\s\d{2}:\d{2}:\d{2})\s+(\d+)"
        client_data = []

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                matches = re.findall(pattern, text)
                for data_hora, codigo in matches:
                    client_data.append({
                        "Data/Hora": data_hora,
                        "Codigo Cliente": codigo,
                    })

        return salvar_codigos_clientes_extraidos(client_data, excel_path, "PDF")
    except Exception as error:
        registrar_evento(
            f"Erro na extracao: {error}",
            level="error",
            stage="Extracao",
            status="error",
        )
        return False


def extrair_registros_visiveis_relatorio(driver):
    script = """
    const isVisible = (element) => {
      if (!element) return false;
      const style = window.getComputedStyle(element);
      if (!style || style.display === 'none' || style.visibility === 'hidden') return false;
      const rect = element.getBoundingClientRect();
      return rect.width > 0 && rect.height > 0;
    };

    const getText = (element) => (element?.innerText || element?.textContent || '')
      .replace(/\\s+/g, ' ')
      .trim();

    const rows = [];
    const seen = new Set();
    for (const row of document.querySelectorAll('.x-grid-item, tr.x-grid-row, tr.x-grid-data-row')) {
      if (!isVisible(row)) continue;

      const cells = Array.from(row.querySelectorAll('.x-grid-cell-inner'))
        .map(getText)
        .filter(Boolean);

      if (!cells.length) continue;

      const signature = cells.join(' | ');
      if (seen.has(signature)) continue;
      seen.add(signature);
      rows.push(cells);
    }

    return rows;
    """

    try:
        linhas_grade = driver.execute_script(script) or []
    except Exception:
        linhas_grade = []

    registros = []
    for linha in linhas_grade:
        celulas = [str(celula).strip() for celula in linha if str(celula).strip()]
        if not celulas:
            continue

        data = next((celula for celula in celulas if re.fullmatch(r"\d{2}/\d{2}/\d{4}", celula)), "")
        hora = next((celula for celula in celulas if re.fullmatch(r"\d{2}:\d{2}(?::\d{2})?", celula)), "")
        codigo = ""

        indice_inicio_busca = 0
        for indice, celula in enumerate(celulas):
            if re.fullmatch(r"\d{2}/\d{2}/\d{4}", celula) or re.fullmatch(r"\d{2}:\d{2}(?::\d{2})?", celula):
                indice_inicio_busca = indice + 1
                continue
            break

        for celula in celulas[indice_inicio_busca:]:
            if re.fullmatch(r"\d{1,10}", celula):
                codigo = celula
                break

        if not codigo:
            codigo = next((celula for celula in celulas if re.fullmatch(r"\d{1,10}", celula)), "")

        if not codigo:
            continue

        registros.append({
            "Data/Hora": " ".join(parte for parte in [data, hora] if parte).strip(),
            "Codigo Cliente": codigo,
            "Assinatura Linha": " | ".join(celulas),
        })

    return registros


def localizar_botao_proxima_pagina_relatorio(driver):
    try:
        return driver.find_element(
            By.XPATH,
            "//a[not(contains(@class,'x-btn-disabled')) and not(contains(@class,'x-item-disabled'))]"
            "[.//span[contains(@class,'x-tbar-page-next')]]",
        )
    except Exception:
        return None


def extrair_codigos_clientes_da_grade(driver, excel_path, max_paginas=20):
    registrar_evento(
        "Extraindo codigos dos clientes diretamente da grade do relatorio...",
        stage="Extracao",
        status="running",
    )

    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Quantidade de Acessos:')]"))
        )
    except Exception:
        pass

    client_data = []
    paginas_vistas = set()

    for indice_pagina in range(max_paginas):
        time.sleep(1.2)
        registros_pagina = extrair_registros_visiveis_relatorio(driver)

        if not registros_pagina:
            break

        assinatura_pagina = tuple(
            registro.get("Assinatura Linha") or f"{registro.get('Data/Hora', '')}|{registro['Codigo Cliente']}"
            for registro in registros_pagina[:10]
        )
        if assinatura_pagina in paginas_vistas:
            break

        paginas_vistas.add(assinatura_pagina)
        client_data.extend(registros_pagina)

        botao_proxima = localizar_botao_proxima_pagina_relatorio(driver)
        if not botao_proxima:
            break

        try:
            driver.execute_script("arguments[0].click();", botao_proxima)
        except Exception:
            ActionChains(driver).move_to_element(botao_proxima).click().perform()

        try:
            WebDriverWait(driver, 10).until(
                lambda current_driver: tuple(
                    registro.get("Assinatura Linha") or f"{registro.get('Data/Hora', '')}|{registro['Codigo Cliente']}"
                    for registro in extrair_registros_visiveis_relatorio(current_driver)[:10]
                ) != assinatura_pagina
            )
        except Exception:
            if indice_pagina > 0:
                break

    return salvar_codigos_clientes_extraidos(client_data, excel_path, "grade do relatorio")


def acionar_download_pdf_no_visualizador(driver, handles_antes, timeout=8):
    janela_original = driver.current_window_handle
    fim_espera = time.time() + timeout

    def clicar_botao_download():
        seletores = [
            (By.ID, "open-button"),
            (By.ID, "download"),
            (By.CSS_SELECTOR, "cr-icon-button#download"),
            (By.CSS_SELECTOR, "button[aria-label*='Download']"),
            (By.CSS_SELECTOR, "button[title*='Download']"),
        ]

        for by, seletor in seletores:
            for elemento in driver.find_elements(by, seletor):
                try:
                    if elemento.is_displayed() and elemento.is_enabled():
                        driver.execute_script("arguments[0].click();", elemento)
                        return True
                except Exception:
                    continue

        return False

    while time.time() < fim_espera:
        handles_atuais = driver.window_handles
        handles_novos = [handle for handle in handles_atuais if handle not in handles_antes]
        handles_busca = handles_novos + [handle for handle in handles_atuais if handle not in handles_novos]

        for handle in handles_busca:
            try:
                driver.switch_to.window(handle)
                driver.switch_to.default_content()

                if clicar_botao_download():
                    try:
                        driver.switch_to.window(janela_original)
                        driver.switch_to.default_content()
                    except Exception:
                        pass
                    return True

                frames = driver.find_elements(By.CSS_SELECTOR, "iframe, frame")
                for frame in frames:
                    try:
                        driver.switch_to.default_content()
                        driver.switch_to.frame(frame)
                        if clicar_botao_download():
                            try:
                                driver.switch_to.default_content()
                                driver.switch_to.window(janela_original)
                                driver.switch_to.default_content()
                            except Exception:
                                pass
                            return True
                    except Exception:
                        continue
                    finally:
                        driver.switch_to.default_content()
            except Exception:
                continue

        time.sleep(1)

    try:
        driver.switch_to.window(janela_original)
        driver.switch_to.default_content()
    except Exception:
        pass

    return False


def acessar_relatorio_e_salvar_pdf(driver):
    wait = WebDriverWait(driver, 25)
    registrar_evento(
        "Acessando relatorio 'Backups nao gerados'...",
        stage="Relatorio",
        status="running",
    )

    try:
        wait.until(EC.element_to_be_clickable((By.ID, "OF9_id-btnWrap"))).click()
        time.sleep(0.8)
        wait.until(EC.element_to_be_clickable((By.ID, "O13C_id"))).click()
        time.sleep(1.2)

        menu_item = wait.until(
            EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Backups') and contains(text(), 'gerados')]"))
        )
        menu_item.click()

        try:
            campo_filtro = wait.until(EC.presence_of_element_located((By.ID, "O205_id-triggerWrap")))
            campo_filtro.click()
            input_field = driver.find_element(By.ID, "O205_id-inputEl")
            input_field.send_keys(Keys.CONTROL + "a")
            input_field.send_keys(Keys.BACKSPACE)
            input_field.send_keys(Keys.RETURN)
            time.sleep(2)
        except Exception:
            pass

        data_str = datetime.now().strftime("%d-%m-%Y")
        caminho_xlsx = os.path.join(PASTA_RELATORIOS, f"Codigos_Clientes_{data_str}.xlsx")
        caminho_xlsx_grade = extrair_codigos_clientes_da_grade(driver, caminho_xlsx)
        if caminho_xlsx_grade:
            return caminho_xlsx_grade

        registrar_evento(
            "Extracao pela grade nao retornou clientes; tentando gerar o PDF como alternativa.",
            level="warning",
            stage="Relatorio",
            status="running",
        )

        pastas_monitoradas = [PASTA_RELATORIOS]
        if os.path.isdir(PASTA_RELATORIOS_LEGADO) and PASTA_RELATORIOS_LEGADO not in pastas_monitoradas:
            pastas_monitoradas.append(PASTA_RELATORIOS_LEGADO)
        if os.path.isdir(PASTA_RELATORIOS_LEGADO_ALT) and PASTA_RELATORIOS_LEGADO_ALT not in pastas_monitoradas:
            pastas_monitoradas.append(PASTA_RELATORIOS_LEGADO_ALT)

        arquivos_antes = {
            pasta: set(os.listdir(pasta))
            for pasta in pastas_monitoradas
            if os.path.isdir(pasta)
        }
        handles_antes = set(driver.window_handles)
        janela_original = driver.current_window_handle
        inicio_download = time.time()

        btn_visualizar = wait.until(EC.element_to_be_clickable((By.ID, "O1C2_id-btnEl")))
        ActionChains(driver).move_to_element(btn_visualizar).click().perform()
        time.sleep(1.5)

        download_acionado_no_visualizador = acionar_download_pdf_no_visualizador(driver, handles_antes)

        try:
            driver.switch_to.window(janela_original)
            driver.switch_to.default_content()
        except Exception:
            pass

        if download_acionado_no_visualizador:
            registrar_evento(
                "Download acionado pelo visualizador PDF.",
                level="info",
                stage="Relatorio",
                status="running",
                log=False,
            )
        else:
            registrar_evento(
                "Visualizador PDF nao exigiu confirmacao; aguardando download direto do navegador.",
                level="info",
                stage="Relatorio",
                status="running",
                log=False,
            )

        pasta_origem, novo_arquivo = localizar_pdf_baixado(
            pastas_monitoradas,
            arquivos_antes,
            inicio_download,
            timeout=40,
        )

        if not novo_arquivo or not pasta_origem:
            registrar_evento(
                "Nenhum PDF novo foi encontrado apos visualizar o relatorio nas pastas monitoradas.",
                level="error",
                stage="Relatorio",
                status="error",
            )
            return False

        caminho_pdf = os.path.join(PASTA_RELATORIOS, f"Relatorio_Backups_{data_str}.pdf")
        caminho_origem = os.path.join(pasta_origem, novo_arquivo)

        if os.path.exists(caminho_pdf):
            os.remove(caminho_pdf)

        if os.path.abspath(caminho_origem) != os.path.abspath(caminho_pdf):
            os.replace(caminho_origem, caminho_pdf)
        else:
            caminho_pdf = caminho_origem

        registrar_evento(
            f"PDF salvo como: {caminho_pdf}",
            level="success",
            stage="Relatorio",
            status="running",
            artifact={
                "type": "pdf",
                "label": os.path.basename(caminho_pdf),
                "path": caminho_pdf,
            },
        )

        return extrair_codigos_clientes(caminho_pdf, caminho_xlsx)
    except Exception as error:
        registrar_evento(
            f"Erro ao acessar relatorio: {error}",
            level="error",
            stage="Relatorio",
            status="error",
        )
        return False


def ler_planilha_clientes(caminho):
    registrar_evento("Lendo planilha de clientes...", stage="Preparacao da lista", status="running")
    try:
        df = pd.read_excel(caminho, dtype=str)
        if "Codigo Cliente" in df.columns:
            df = df[["Codigo Cliente"]].rename(columns={"Codigo Cliente": "Codigo do Cliente"})
        else:
            df = df.iloc[:, [0]]
            df.columns = ["Codigo do Cliente"]

        df = df.dropna(subset=["Codigo do Cliente"])
        df["Codigo do Cliente"] = df["Codigo do Cliente"].astype(str).str.replace(r"\D", "", regex=True)
        df = df[df["Codigo do Cliente"] != ""]

        registrar_evento(
            f"{len(df)} codigos unicos lidos e prontos para consulta.",
            stage="Preparacao da lista",
            status="running",
            metrics={"totalClients": len(df), "processedClients": 0},
        )
        return df
    except Exception as error:
        registrar_evento(
            f"Erro ao ler planilha: {error}",
            level="error",
            stage="Preparacao da lista",
            status="error",
        )
        return None


def abrir_botao_sair(driver):
    try:
        wait = WebDriverWait(driver, 10)
        btn = wait.until(EC.presence_of_element_located((By.ID, "O29C_id-btnEl")))
        driver.execute_script("arguments[0].click();", btn)
        registrar_evento("Botao sair clicado.", stage="Encerramento", status="running")
        time.sleep(1.5)
    except Exception:
        registrar_evento(
            "Nao foi possivel clicar no botao sair.",
            level="warning",
            stage="Encerramento",
            status="running",
        )


def abrir_cadastro_clientes(driver):
    wait = WebDriverWait(driver, 3)
    registrar_evento("Retornando a tela principal...", stage="Consulta dos clientes", status="running")

    try:
        b1 = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "O29C_id")))
        b1.click()
        time.sleep(1)
    except Exception:
        pass

    try:
        b2 = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "O1BA_id")))
        b2.click()
        time.sleep(1.5)
    except Exception:
        pass

    btn = wait.until(EC.element_to_be_clickable((By.ID, "O8C_id")))
    btn.click()
    time.sleep(2.5)

    registrar_evento("Cadastro de clientes aberto.", level="success", stage="Consulta dos clientes", status="running")


def ler_campo(driver, campo_id, timeout=15):
    try:
        el = WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, campo_id)))
        valor = el.get_attribute("value") or el.text or ""
        return valor.strip()
    except Exception:
        return ""


def ler_valor_por_rotulo(driver, rotulos, timeout=3):
    aliases = [normalizar_texto_busca(rotulo) for rotulo in rotulos if rotulo]
    if not aliases:
        return "", ""

    script = """
    const aliases = arguments[0];
    const normalize = (value) => (value || '')
      .normalize('NFD')
      .replace(/[\\u0300-\\u036f]/g, '')
      .toLowerCase()
      .replace(/\\s+/g, ' ')
      .trim();

    const isVisible = (element) => {
      if (!element) return false;
      const style = window.getComputedStyle(element);
      if (!style || style.display === 'none' || style.visibility === 'hidden') return false;
      const rect = element.getBoundingClientRect();
      return rect.width > 0 && rect.height > 0;
    };

    const getElementValue = (element) => {
      if (!element) return '';
      if (typeof element.value === 'string' && element.value.trim()) {
        return element.value.trim();
      }

      const text = (element.innerText || element.textContent || '').trim();
      return text;
    };

    const labelCandidates = Array.from(document.querySelectorAll('label, span, div, td'))
      .filter((element) => {
        if (!isVisible(element) || element.querySelector('input, textarea, select')) {
          return false;
        }

        const text = normalize(element.innerText || element.textContent || '');
        if (!text || text.length < 3 || text.length > 80) {
          return false;
        }

        return aliases.some((alias) => text.includes(alias));
      });

    const fieldCandidates = Array.from(
      document.querySelectorAll(
        'input, textarea, select, [role="textbox"], div.x-form-display-field, span.x-form-display-field'
      )
    )
      .filter((element) => {
        if (!isVisible(element)) return false;
        return Boolean(getElementValue(element));
      })
      .map((element) => {
        const rect = element.getBoundingClientRect();
        return {
          value: getElementValue(element),
          left: rect.left,
          right: rect.right,
          top: rect.top,
          bottom: rect.bottom,
        };
      });

    for (const label of labelCandidates) {
      const labelText = (label.innerText || label.textContent || '').trim();
      const labelRect = label.getBoundingClientRect();
      const labelCenterY = (labelRect.top + labelRect.bottom) / 2;
      let bestMatch = null;

      for (const field of fieldCandidates) {
        const fieldCenterY = (field.top + field.bottom) / 2;
        const horizontalGap = field.left - labelRect.right;
        const sameRow = Math.abs(fieldCenterY - labelCenterY) <= 45;
        const rightSide = field.right >= labelRect.right - 30;

        if (!sameRow || !rightSide || horizontalGap < -80) {
          continue;
        }

        const score = Math.max(horizontalGap, 0) + Math.abs(fieldCenterY - labelCenterY) * 5;
        if (!bestMatch || score < bestMatch.score) {
          bestMatch = {
            label: labelText,
            value: field.value,
            score,
          };
        }
      }

      if (bestMatch && bestMatch.value) {
        return bestMatch;
      }
    }

    return null;
    """

    fim_espera = time.time() + timeout
    while time.time() < fim_espera:
        try:
            resultado = driver.execute_script(script, aliases)
        except Exception:
            resultado = None

        if resultado and resultado.get("value"):
            return resultado["value"].strip(), resultado.get("label", "").strip()

        time.sleep(0.5)

    return "", ""


def classificar_situacao_empresa(texto):
    situacao_normalizada = normalizar_texto_busca(texto)
    if not situacao_normalizada:
        return None

    if situacao_normalizada in TERMOS_EMPRESA_ATIVA:
        return True
    if situacao_normalizada in TERMOS_EMPRESA_INATIVA:
        return False

    if any(termo in situacao_normalizada for termo in TERMOS_EMPRESA_INATIVA if len(termo) > 2):
        return False
    if any(termo in situacao_normalizada for termo in TERMOS_EMPRESA_ATIVA if len(termo) > 2):
        return True

    return None


def buscar_cliente(driver, codigo):
    wait = WebDriverWait(driver, 15)
    campo = wait.until(EC.element_to_be_clickable((By.ID, "O3F2_id-inputEl")))
    campo.click()
    campo.send_keys(Keys.CONTROL + "a")
    campo.send_keys(Keys.BACKSPACE)
    campo.send_keys(codigo)
    campo.send_keys(Keys.RETURN)
    time.sleep(2.5)


def clicar_aba_dados_adicionais(driver, wait):
    try:
        aba_inner = wait.until(EC.element_to_be_clickable((By.ID, "O300_id_tab-btnInnerEl")))
        aba_inner.click()
        time.sleep(0.8)
        aba_inner.click()

        WebDriverWait(driver, 15).until(
            EC.any_of(
                EC.presence_of_element_located((By.ID, "O358_id-inputEl")),
                EC.presence_of_element_located((By.ID, "O340_id-inputEl")),
            )
        )
        time.sleep(1.3)
        return True
    except Exception as error:
        registrar_evento(
            f"Erro ao abrir aba Dados Adicionais: {str(error)[:100]}",
            level="warning",
            stage="Consulta dos clientes",
            status="running",
        )
        return False


def clicar_aba_empresa(driver, wait):
    try:
        wait.until(EC.element_to_be_clickable((By.ID, "O20C_id_tab-btnInnerEl"))).click()
        time.sleep(1.0)
        return True
    except Exception as error:
        registrar_evento(
            f"Erro ao abrir aba principal da empresa: {str(error)[:100]}",
            level="warning",
            stage="Consulta dos clientes",
            status="running",
        )
        return False


def obter_dados_empresa(driver, wait):
    dados_empresa = {
        "nome": "Nao foi possivel ler",
        "situacao": "",
        "rotulo_situacao": "",
        "ativa": None,
    }

    if not clicar_aba_empresa(driver, wait):
        return dados_empresa

    nome_empresa = ler_campo(driver, "O2E8_id-inputEl")
    if nome_empresa:
        dados_empresa["nome"] = nome_empresa

    situacao_texto, rotulo_situacao = ler_valor_por_rotulo(driver, ROTULOS_SITUACAO_EMPRESA, timeout=2.5)
    dados_empresa["situacao"] = situacao_texto
    dados_empresa["rotulo_situacao"] = rotulo_situacao
    dados_empresa["ativa"] = classificar_situacao_empresa(situacao_texto)

    return dados_empresa


def executar_automacao(driver, df_clientes):
    wait = WebDriverWait(driver, 18)
    resultados_pendentes = []
    total_clientes = len(df_clientes)
    clientes_sem_backup_hoje = 0
    empresas_inativas_ignoradas = 0
    empresas_sem_situacao_identificada = 0

    registrar_evento(
        "Iniciando consulta cliente a cliente.",
        stage="Consulta dos clientes",
        status="running",
        metrics={"totalClients": total_clientes, "processedClients": 0, "pendingCompanies": 0},
    )

    for index, row in df_clientes.iterrows():
        codigo = row["Codigo do Cliente"]
        enviar_evento_dashboard(
            message=f"Consultando cliente {codigo} ({index + 1}/{total_clientes})",
            stage="Consulta dos clientes",
            status="running",
            metrics={"processedClients": index + 1, "totalClients": total_clientes},
            log=False,
        )

        buscar_cliente(driver, codigo)

        if not clicar_aba_dados_adicionais(driver, wait):
            continue

        texto_backup = ler_campo(driver, "O358_id-inputEl")
        if DATA_HOJE in texto_backup:
            continue

        clientes_sem_backup_hoje += 1
        representante = ler_campo(driver, "O340_id-inputEl")
        dados_empresa = obter_dados_empresa(driver, wait)
        nome_empresa = dados_empresa["nome"]
        situacao_empresa = dados_empresa["situacao"] or "Nao identificado"
        empresa_ativa = dados_empresa["ativa"]

        if empresa_ativa is False:
            empresas_inativas_ignoradas += 1
            registrar_evento(
                f"Empresa inativa ignorada no relatorio: {codigo} - {nome_empresa} ({situacao_empresa})",
                level="info",
                stage="Consulta dos clientes",
                status="running",
            )
            clicar_aba_dados_adicionais(driver, wait)
            continue

        if empresa_ativa is None:
            empresas_sem_situacao_identificada += 1
            detalhe_rotulo = (
                f" Campo identificado: {dados_empresa['rotulo_situacao']}."
                if dados_empresa["rotulo_situacao"]
                else ""
            )
            registrar_evento(
                f"Situacao da empresa nao identificada para {codigo} - {nome_empresa}; empresa nao sera levada ao relatorio.{detalhe_rotulo}",
                level="warning",
                stage="Consulta dos clientes",
                status="running",
            )
            clicar_aba_dados_adicionais(driver, wait)
            continue

        resultados_pendentes.append({
            "Codigo": codigo,
            "Empresa": nome_empresa,
            "Representante": representante,
            "Status": "Sem Backup de Hoje",
            "Situacao Empresa": situacao_empresa,
            "Ultimo Backup": texto_backup,
        })

        registrar_evento(
            f"Pendencia encontrada: {codigo} - {nome_empresa} ({situacao_empresa})",
            level="warning",
            stage="Consulta dos clientes",
            status="running",
            metrics={"pendingCompanies": len(resultados_pendentes)},
        )

        clicar_aba_dados_adicionais(driver, wait)

    if not resultados_pendentes:
        if clientes_sem_backup_hoje == 0:
            mensagem_sem_pendencias = "Todos os backups do relatorio estao em dia."
            nivel = "success"
        else:
            detalhes = []
            if empresas_inativas_ignoradas:
                detalhes.append(f"{empresas_inativas_ignoradas} inativa(s)")
            if empresas_sem_situacao_identificada:
                detalhes.append(f"{empresas_sem_situacao_identificada} com situacao nao identificada")

            sufixo = f" Filtro aplicado: {', '.join(detalhes)}." if detalhes else ""
            mensagem_sem_pendencias = (
                "Nenhuma empresa ativa ficou apta para entrar no relatorio de pendencias."
                f"{sufixo}"
            )
            nivel = "info"

        registrar_evento(
            mensagem_sem_pendencias,
            level=nivel,
            stage="Consulta dos clientes",
            status="completed",
            metrics={"pendingCompanies": 0, "processedClients": total_clientes},
        )
        return None

    df_final = pd.DataFrame(resultados_pendentes)
    df_final = df_final.sort_values(by=["Representante", "Empresa"])
    caminho_pendencias = os.path.join(DIRETORIO_SCRIPT, "PENDENCIAS.xlsx")

    try:
        df_final.to_excel(caminho_pendencias, index=False)
    except PermissionError:
        caminho_pendencias = os.path.join(DIRETORIO_SCRIPT, f"PENDENCIAS_{int(time.time())}.xlsx")
        df_final.to_excel(caminho_pendencias, index=False)

    registrar_evento(
        f"Relatorio de pendencias gerado: {caminho_pendencias}",
        level="success",
        stage="Pendencias",
        status="running",
        metrics={"pendingCompanies": len(df_final), "processedClients": total_clientes},
        artifact={
            "type": "xlsx",
            "label": os.path.basename(caminho_pendencias),
            "path": caminho_pendencias,
        },
    )
    return caminho_pendencias


def separar_por_representante(caminho_arquivo_entrada):
    registrar_evento("Separando planilhas por representante...", stage="Separacao", status="running")

    if not caminho_arquivo_entrada:
        registrar_evento(
            "Nenhum arquivo de pendencias foi informado para separacao.",
            level="warning",
            stage="Separacao",
            status="error",
        )
        return None

    if not os.path.isabs(caminho_arquivo_entrada):
        caminho_arquivo_entrada = os.path.join(DIRETORIO_SCRIPT, caminho_arquivo_entrada)

    if not os.path.exists(caminho_arquivo_entrada):
        registrar_evento(
            f"Arquivo de pendencias nao encontrado: {caminho_arquivo_entrada}",
            level="error",
            stage="Separacao",
            status="error",
        )
        return None

    try:
        df = pd.read_excel(caminho_arquivo_entrada)
        if "Representante" not in df.columns:
            registrar_evento(
                "Coluna 'Representante' nao encontrada para separacao.",
                level="error",
                stage="Separacao",
                status="error",
            )
            return None

        garantir_pasta(PASTA_PLANILHAS_SEPARADAS)
        representantes = df["Representante"].dropna().astype(str).unique()

        for rep in representantes:
            df_rep = df[df["Representante"].astype(str) == str(rep)]
            nome_arquivo_saida = f"{normalizar_nome_arquivo(rep)}.xlsx"
            caminho_saida = os.path.join(PASTA_PLANILHAS_SEPARADAS, nome_arquivo_saida)
            df_rep.to_excel(caminho_saida, index=False)
            enviar_evento_dashboard(
                message=f"Planilha separada para {rep}.",
                stage="Separacao",
                status="running",
                artifact={
                    "type": "xlsx",
                    "label": nome_arquivo_saida,
                    "path": caminho_saida,
                },
                log=True,
            )

        registrar_evento(
            f"Separacao concluida para {len(representantes)} representantes.",
            level="success",
            stage="Separacao",
            status="running",
            metrics={"representatives": len(representantes)},
        )
        return PASTA_PLANILHAS_SEPARADAS
    except Exception as error:
        registrar_evento(
            f"Erro na separacao: {error}",
            level="error",
            stage="Separacao",
            status="error",
        )
        return None


def enviar_whatsapp_api(telefone, caminho_anexo, mensagem, nome_destino):
    try:
        import requests
    except ImportError:
        registrar_evento(
            "Biblioteca 'requests' nao encontrada no Python.",
            level="error",
            stage="Envio no WhatsApp",
            status="error",
        )
        return False

    telefone_limpo = "".join(filter(str.isdigit, str(telefone)))
    if not telefone_limpo.startswith("55"):
        telefone_limpo = "55" + telefone_limpo

    registrar_evento(
        f"Enviando mensagem e anexo para {nome_destino} ({telefone_limpo})...",
        stage="Envio no WhatsApp",
        status="running",
    )

    try:
        res_msg = requests.post(
            URL_MSG,
            json={"number": telefone_limpo, "message": mensagem},
            timeout=25,
        )

        if res_msg.status_code != 200:
            registrar_evento(
                f"Erro ao enviar mensagem para {nome_destino}: {res_msg.text}",
                level="error",
                stage="Envio no WhatsApp",
                status="running",
            )
            return False

        with open(caminho_anexo, "rb") as arquivo:
            files = {
                "documento": (
                    os.path.basename(caminho_anexo),
                    arquivo,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            }
            data = {"number": telefone_limpo}
            res_media = requests.post(URL_MEDIA, data=data, files=files, timeout=25)

        if res_media.status_code == 200:
            registrar_evento(
                f"Envio concluido para {nome_destino}.",
                level="success",
                stage="Envio no WhatsApp",
                status="running",
            )
            return True

        registrar_evento(
            f"Erro ao enviar arquivo para {nome_destino}: {res_media.text}",
            level="error",
            stage="Envio no WhatsApp",
            status="running",
        )
        return False
    except Exception as error:
        registrar_evento(
            f"Erro de conexao com o servidor Baileys: {error}",
            level="error",
            stage="Envio no WhatsApp",
            status="error",
        )
        return False


def carregar_mapeamento_representantes():
    if not os.path.exists(CAMINHO_MAPEAMENTO):
        registrar_evento(
            f"Arquivo de mapeamento nao encontrado: {CAMINHO_MAPEAMENTO}",
            level="error",
            stage="Envio no WhatsApp",
            status="error",
        )
        return None

    try:
        df_mapeamento = pd.read_excel(CAMINHO_MAPEAMENTO)
        if "Representante" not in df_mapeamento.columns or "Telefone" not in df_mapeamento.columns:
            registrar_evento(
                "Colunas 'Representante' ou 'Telefone' nao encontradas no mapeamento.",
                level="error",
                stage="Envio no WhatsApp",
                status="error",
            )
            return None

        mapeamento = dict(
            zip(
                df_mapeamento["Representante"].astype(str).str.strip(),
                df_mapeamento["Telefone"].astype(str).str.strip(),
            )
        )
        registrar_evento(
            f"Mapeamento carregado com {len(mapeamento)} representantes.",
            level="success",
            stage="Envio no WhatsApp",
            status="running",
        )
        return mapeamento
    except Exception as error:
        registrar_evento(
            f"Erro ao ler mapeamento_representantes.xlsx: {error}",
            level="error",
            stage="Envio no WhatsApp",
            status="error",
        )
        return None


def enviar_planilhas_representantes_whatsapp(pasta_planilhas):
    mapeamento = carregar_mapeamento_representantes()
    if not mapeamento:
        return {
            "total": 0,
            "enviados": 0,
            "falhas": 1,
            "sem_telefone": 0,
            "erro": "Falha ao carregar o mapeamento de representantes.",
        }

    if not pasta_planilhas or not os.path.isdir(pasta_planilhas):
        registrar_evento(
            f"Pasta de planilhas nao encontrada: {pasta_planilhas}",
            level="error",
            stage="Envio no WhatsApp",
            status="error",
        )
        return {
            "total": 0,
            "enviados": 0,
            "falhas": 1,
            "sem_telefone": 0,
            "erro": "Pasta de planilhas nao encontrada.",
        }

    arquivos = sorted(
        arquivo for arquivo in os.listdir(pasta_planilhas) if arquivo.lower().endswith(".xlsx")
    )

    if not arquivos:
        registrar_evento(
            "Nenhuma planilha separada foi encontrada para envio.",
            level="warning",
            stage="Envio no WhatsApp",
            status="completed",
        )
        return {
            "total": 0,
            "enviados": 0,
            "falhas": 0,
            "sem_telefone": 0,
            "erro": "",
        }

    registrar_evento(
        f"Iniciando envio automatico para {len(arquivos)} representantes.",
        stage="Envio no WhatsApp",
        status="running",
        metrics={"representatives": len(arquivos)},
    )

    resumo_envio = {
        "total": len(arquivos),
        "enviados": 0,
        "falhas": 0,
        "sem_telefone": 0,
        "erro": "",
    }

    for arquivo in arquivos:
        caminho_anexo = os.path.join(pasta_planilhas, arquivo)
        nome_arquivo = os.path.splitext(arquivo)[0]
        telefone = None
        representante = nome_arquivo.replace("_", " ")

        for rep_nome, tel in mapeamento.items():
            if normalizar_nome_arquivo(rep_nome) == nome_arquivo:
                telefone = tel
                representante = rep_nome
                break

        if not telefone:
            registrar_evento(
                f"Nenhum telefone encontrado para o arquivo: {arquivo}",
                level="warning",
                stage="Envio no WhatsApp",
                status="running",
            )
            resumo_envio["sem_telefone"] += 1
            continue

        if enviar_whatsapp_api(telefone, caminho_anexo, MENSAGEM_WHATSAPP, representante):
            resumo_envio["enviados"] += 1
        else:
            resumo_envio["falhas"] += 1
        time.sleep(2)

    registrar_evento(
        (
            "Resumo do envio pelo WhatsApp: "
            f"{resumo_envio['enviados']} enviado(s), "
            f"{resumo_envio['falhas']} falha(s), "
            f"{resumo_envio['sem_telefone']} sem telefone."
        ),
        level="success" if resumo_envio["falhas"] == 0 and resumo_envio["sem_telefone"] == 0 else "warning",
        stage="Envio no WhatsApp",
        status="running" if resumo_envio["falhas"] or resumo_envio["sem_telefone"] else "completed",
    )

    return resumo_envio


def criar_driver():
    pasta_dest = garantir_pasta(PASTA_RELATORIOS)
    registrar_evento(
        f"Pasta de download configurada: {pasta_dest}",
        stage="Preparacao",
        status="starting",
        log=False,
    )

    chrome_options = webdriver.ChromeOptions()
    caminhos_brave = [
        r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
        r"C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe",
        os.path.expanduser(r"~\AppData\Local\BraveSoftware\Brave-Browser\Application\brave.exe"),
    ]

    for caminho in caminhos_brave:
        if os.path.exists(caminho):
            chrome_options.binary_location = caminho
            break

    prefs = {
        "download.default_directory": pasta_dest,
        "download.prompt_for_download": False,
        "safebrowsing.enabled": False,
        "safebrowsing.disable_download_protection": True,
        "profile.default_content_settings.popups": 0,
        "plugins.always_open_pdf_externally": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument("--allow-running-insecure-content")
    chrome_options.add_argument(f"--unsafely-treat-insecure-origin-as-secure={URL_SISTEMA}")

    driver = webdriver.Chrome(options=chrome_options)
    driver.get(URL_SISTEMA)
    driver.maximize_window()

    try:
        WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.ID, "details-button"))).click()
        WebDriverWait(driver, 2).until(EC.element_to_be_clickable((By.ID, "proceed-link"))).click()
    except Exception:
        pass

    return driver


def reiniciar_sessao_para_cadastro(driver, usuario, senha):
    registrar_evento("Resetando sessao para cadastro de clientes...", stage="Consulta dos clientes", status="running")
    try:
        driver.get(URL_SISTEMA)
        time.sleep(0.8)
        fazer_login(driver, usuario, senha)
    except Exception:
        pass


def main():
    global EXECUCAO_ID

    EXECUCAO_ID = datetime.now().strftime("%Y%m%d%H%M%S")

    arquivos_removidos = limpar_arquivos_execucao_anterior()
    resetar_dashboard(len(arquivos_removidos))

    if arquivos_removidos:
        registrar_evento(
            f"Arquivos antigos removidos no inicio da execucao: {len(arquivos_removidos)}.",
            level="info",
            stage="Preparacao",
            status="starting",
            metrics={"cleanedFiles": len(arquivos_removidos)},
        )
    else:
        registrar_evento(
            "Nenhum arquivo antigo precisou ser removido.",
            level="info",
            stage="Preparacao",
            status="starting",
            metrics={"cleanedFiles": 0},
        )

    print("\n=== Automacao Completa de Extracao e Checagem ===\n")
    usuario, senha, credenciais_via_painel = obter_credenciais_iniciais()

    registrar_evento("Abrindo navegador do sistema...", stage="Login", status="running")
    driver = criar_driver()

    login_sucesso = False
    while not login_sucesso:
        if fazer_login(driver, usuario, senha):
            login_sucesso = True
        else:
            if credenciais_via_painel:
                raise RuntimeError("Credenciais informadas no painel sao invalidas.")

            print("\nCredenciais incorretas. Informe novamente.\n")
            try:
                driver.refresh()
                time.sleep(2.5)
            except Exception:
                pass
            usuario = input("Usuario: ")
            senha = digitar_senha("Senha: ")

    caminho_pendencias = None
    status_final = "completed"
    mensagem_final = "Execucao concluida."

    try:
        caminho_planilha = acessar_relatorio_e_salvar_pdf(driver)
        if caminho_planilha:
            df_clientes = ler_planilha_clientes(caminho_planilha)
            if df_clientes is not None and not df_clientes.empty:
                reiniciar_sessao_para_cadastro(driver, usuario, senha)
                abrir_cadastro_clientes(driver)
                caminho_pendencias = executar_automacao(driver, df_clientes)
                abrir_botao_sair(driver)
            else:
                status_final = "error"
                mensagem_final = "Planilha de clientes nao pode ser lida."
        else:
            status_final = "error"
            mensagem_final = "Falha na extracao do relatorio."
    except Exception as error:
        status_final = "error"
        mensagem_final = f"Erro geral: {error}"
        registrar_evento(mensagem_final, level="error", stage="Execucao", status="error")

    registrar_evento("Encerrando navegador do sistema...", stage="Encerramento", status="running")
    time.sleep(3)
    driver.quit()

    if caminho_pendencias and os.path.exists(caminho_pendencias):
        pasta_planilhas = separar_por_representante(caminho_pendencias)
        if pasta_planilhas:
            resumo_envio = enviar_planilhas_representantes_whatsapp(pasta_planilhas)
            if resumo_envio["enviados"] and not resumo_envio["falhas"] and not resumo_envio["sem_telefone"]:
                mensagem_final = "Relatorios gerados e enviados para os representantes."
            elif resumo_envio["enviados"] and (resumo_envio["falhas"] or resumo_envio["sem_telefone"]):
                status_final = "error"
                mensagem_final = (
                    "Planilhas geradas com envio parcial no WhatsApp: "
                    f"{resumo_envio['enviados']} enviado(s), "
                    f"{resumo_envio['falhas']} falha(s), "
                    f"{resumo_envio['sem_telefone']} sem telefone."
                )
            elif resumo_envio["erro"]:
                status_final = "error"
                mensagem_final = resumo_envio["erro"]
            else:
                status_final = "error"
                mensagem_final = (
                    "Planilhas geradas, mas nenhum envio foi concluido no WhatsApp. "
                    f"Falhas: {resumo_envio['falhas']}. "
                    f"Sem telefone: {resumo_envio['sem_telefone']}."
                )
        else:
            status_final = "error"
            mensagem_final = "Falha ao separar planilhas por representante."
    elif status_final != "error":
        mensagem_final = "Nao houve pendencias novas; separacao e envio foram pulados."

    registrar_evento(
        mensagem_final,
        level="success" if status_final == "completed" else "error",
        stage="Concluido" if status_final == "completed" else "Falha",
        status=status_final,
    )

    print("\nPROCESSO 100% CONCLUIDO.")


if __name__ == "__main__":
    main()
