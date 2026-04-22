import json
import os
import re
import sys
import time
from datetime import datetime

from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

import baixaRel


def navegar_ate_relatorio(driver):
    wait = WebDriverWait(driver, 25)

    wait.until(EC.element_to_be_clickable((By.ID, "OF9_id-btnWrap"))).click()
    time.sleep(0.8)
    wait.until(EC.element_to_be_clickable((By.ID, "O13C_id"))).click()
    time.sleep(1.2)

    menu_item = wait.until(
        EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Backups') and contains(text(), 'gerados')]"))
    )
    menu_item.click()
    time.sleep(2.5)

    return wait


def limpar_filtro_representante(driver, wait):
    campo_filtro = wait.until(EC.presence_of_element_located((By.ID, "O205_id-triggerWrap")))
    ActionChains(driver).move_to_element(campo_filtro).click().perform()
    time.sleep(0.8)

    input_field = driver.find_element(By.ID, "O205_id-inputEl")
    input_field.send_keys(Keys.CONTROL + "a")
    input_field.send_keys(Keys.BACKSPACE)
    input_field.send_keys(Keys.RETURN)
    time.sleep(4)


def coletar_grade_visivel(driver):
    script = """
    const seletores = [
      '.x-grid-item',
      '.x-grid-row',
      '.x-grid-data-row',
      'tr',
      '[role="row"]',
      '.x-boundlist-item'
    ];
    const vistos = new Set();
    const linhas = [];

    function visivel(el) {
      const estilo = window.getComputedStyle(el);
      const rect = el.getBoundingClientRect();
      return estilo && estilo.display !== 'none' && estilo.visibility !== 'hidden' && rect.width > 0 && rect.height > 0;
    }

    for (const seletor of seletores) {
      for (const el of document.querySelectorAll(seletor)) {
        if (!visivel(el)) continue;
        const texto = (el.innerText || el.textContent || '').replace(/\\s+/g, ' ').trim();
        if (!texto || texto.length < 4 || vistos.has(texto)) continue;
        vistos.add(texto);
        linhas.push({ seletor, texto });
      }
    }

    return linhas.slice(0, 300);
    """
    return driver.execute_script(script)


def salvar_diagnostico(driver, pasta_saida):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    prefixo = os.path.join(pasta_saida, f"inspecao_backups_{timestamp}")

    os.makedirs(pasta_saida, exist_ok=True)

    html_path = prefixo + ".html"
    txt_path = prefixo + ".txt"
    json_path = prefixo + ".json"
    screenshot_path = prefixo + ".png"

    with open(html_path, "w", encoding="utf-8") as arquivo_html:
        arquivo_html.write(driver.page_source)

    texto_visivel = driver.find_element(By.TAG_NAME, "body").text
    with open(txt_path, "w", encoding="utf-8") as arquivo_txt:
        arquivo_txt.write(texto_visivel)

    linhas = coletar_grade_visivel(driver)
    dados = {
        "url": driver.current_url,
        "title": driver.title,
        "linhas_visiveis": linhas,
        "codigos_encontrados_na_pagina": sorted(set(re.findall(r"\\b\\d{4,}\\b", texto_visivel))),
        "pares_data_codigo": [
            {"data_hora": data_hora, "codigo": codigo}
            for data_hora, codigo in re.findall(r"(\\d{2}/\\d{2}/\\d{4}\\s+\\d{2}:\\d{2}:\\d{2})\\s+(\\d+)", texto_visivel)
        ],
    }

    with open(json_path, "w", encoding="utf-8") as arquivo_json:
        json.dump(dados, arquivo_json, ensure_ascii=False, indent=2)

    driver.save_screenshot(screenshot_path)

    return {
        "html": html_path,
        "txt": txt_path,
        "json": json_path,
        "png": screenshot_path,
        "dados": dados,
    }


def main():
    if len(sys.argv) != 3:
        print("Uso: python inspecao_relatorio.py <usuario> <senha>")
        sys.exit(1)

    usuario = sys.argv[1]
    senha = sys.argv[2]
    baixaRel.EXECUCAO_ID = datetime.now().strftime("%Y%m%d%H%M%S")

    driver = baixaRel.criar_driver()
    try:
        if not baixaRel.fazer_login(driver, usuario, senha):
            print("LOGIN_FALHOU")
            sys.exit(2)

        wait = navegar_ate_relatorio(driver)
        limpar_filtro_representante(driver, wait)
        resultado = salvar_diagnostico(driver, baixaRel.PASTA_RELATORIOS)

        print("INSPECAO_OK")
        print(json.dumps(resultado, ensure_ascii=False, indent=2))
    finally:
        time.sleep(2)
        driver.quit()


if __name__ == "__main__":
    main()
