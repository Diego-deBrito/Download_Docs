# -*- coding: utf-8 -*-
"""
Este script de automação (RPA) é projetado para baixar sistematicamente
todos os documentos associados a uma lista de instrumentos (contratos/convênios)
de um portal web.

O robô lê uma lista de um arquivo Excel, navega até a página de cada
instrumento, acessa a aba de contratos, percorre todas as páginas de resultados,
clica em cada botão "Detalhar" para abrir a página de documentos e, por fim,
baixa todos os arquivos disponíveis, organizando-os em pastas separadas.
"""

import logging
import os
import shutil
import time
from typing import Dict, Any

import pandas as pd
from pandas import DataFrame
from selenium import webdriver
from selenium.common.exceptions import WebDriverException, TimeoutException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.remote.webdriver import WebDriver
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# --- Configurações Globais ---
CONFIG: Dict[str, Any] = {
    "chrome_debug_port": "9222",
    "input_file": r"C:\Caminho\Para\Sua\pasta1.xlsx",
    "downloads_dir": r"C:\Caminho\Para\Sua\Pasta\Downloads",
    "output_dir": r"C:\Caminho\Para\Sua\Pasta\De\Contratos"
}

# Configuração do sistema de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("downloader_log.txt"),
        logging.StreamHandler()
    ]
)


def conectar_navegador_existente() -> WebDriver:
    """
    Conecta-se a uma instância do Chrome em execução em modo de depuração.
    """
    options = webdriver.ChromeOptions()
    options.debugger_address = f"localhost:{CONFIG['chrome_debug_port']}"
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        logging.info("Conexão com o navegador existente bem-sucedida.")
        return driver
    except WebDriverException as e:
        logging.critical(f"Não foi possível conectar ao navegador na porta {CONFIG['chrome_debug_port']}. Verifique se ele está em execução. Erro: {e}")
        exit()
    except Exception as e:
        logging.critical(f"Ocorreu um erro inesperado ao conectar ao navegador: {e}")
        exit()


def ler_planilha(file_path: str) -> DataFrame:
    """
    Lê a planilha de entrada e prepara a coluna de instrumentos.
    """
    try:
        df = pd.read_excel(file_path, engine="openpyxl")
        if "Instrumento nº" in df.columns:
            df["Instrumento nº"] = df["Instrumento nº"].astype(str).str.replace(r"\.0$", "", regex=True)
        else:
            raise ValueError("Coluna 'Instrumento nº' não encontrada na planilha.")
        logging.info("Planilha de entrada lida e preparada com sucesso.")
        return df
    except FileNotFoundError:
        logging.critical(f"Arquivo de entrada não encontrado em: {file_path}")
        exit()
    except Exception as e:
        logging.critical(f"Erro ao ler a planilha: {e}")
        exit()


def navegar_menu_principal(driver: WebDriver, instrumento: str) -> bool:
    """
    Navega pelo menu principal do sistema e pesquisa por um instrumento.
    """
    try:
        # Nota: Os seletores XPath completos são frágeis. Recomenda-se usar alternativas mais robustas.
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]"))).click()
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html[1]/body[1]/div[1]/div[3]/div[2]/div[1]/div[1]/ul[1]/li[6]/a[1]"))).click()
        
        campo_pesquisa = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/input[1]")))
        campo_pesquisa.clear()
        campo_pesquisa.send_keys(instrumento)
        
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/span[1]/input[1]"))).click()
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html[1]/body[1]/div[3]/div[15]/div[3]/div[3]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/a[1]"))).click()
        return True
    except Exception as e:
        logging.warning(f"Instrumento {instrumento} não encontrado ou falha na navegação inicial. Erro: {e}")
        return False


def acessar_aba_contratos(driver: WebDriver) -> bool:
    """
    Acessa a aba de contratos dentro da página de um instrumento.
    """
    try:
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[15]/div[1]/div/div[1]/a[6]/div/span/span"))).click()
        WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[15]/div[1]/div/div[2]/a[24]/div/span/span"))).click()
        logging.info("Aba de contratos acessada com sucesso.")
        return True
    except Exception as e:
        logging.error(f"Erro ao acessar a aba de contratos: {e}")
        return False


def mover_arquivo_baixado(arquivo_baixado: str, pasta_destino: str) -> None:
    """
    Move um arquivo da pasta de downloads para a pasta de destino do instrumento.
    """
    origem = os.path.join(CONFIG["downloads_dir"], arquivo_baixado)
    destino = os.path.join(pasta_destino, arquivo_baixado)
    try:
        shutil.move(origem, destino)
        logging.info(f"Arquivo '{arquivo_baixado}' movido para '{pasta_destino}'.")
    except Exception as e:
        logging.error(f"Falha ao mover o arquivo '{arquivo_baixado}': {e}")


def executar_acoes_detalhar(driver: WebDriver, pasta_instrumento: str) -> None:
    """
    Itera sobre os botões "Detalhar" da página atual, baixa e organiza os arquivos.
    """
    try:
        tabela_xpath = "/html/body/div[3]/div[15]/div[3]/div/div/form/div/div[2]/table"
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, tabela_xpath)))
        linhas = driver.find_elements(By.XPATH, f"{tabela_xpath}/tbody/tr")
        
        if not linhas:
            logging.info("Nenhuma linha (contrato) encontrada na tabela desta página.")
            return

        for i in range(len(linhas)):
            try:
                # Técnica para evitar "StaleElementReferenceException":
                # Re-identifica a lista de linhas a cada iteração, pois o DOM é alterado
                # após clicar em "Detalhar" e depois em "Voltar".
                linhas_atualizadas = driver.find_elements(By.XPATH, f"{tabela_xpath}/tbody/tr")
                if i >= len(linhas_atualizadas):
                    break
                
                linha = linhas_atualizadas[i]
                botao_detalhar = linha.find_element(By.XPATH, ".//a[contains(text(), 'Detalhar')]")
                botao_detalhar.click()
                time.sleep(2)

                # Lógica de Download
                try:
                    tabela_download = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[26]/td/div[1]/table")))
                    botoes_download = tabela_download.find_elements(By.XPATH, ".//a[contains(text(), 'Baixar')]")
                    
                    if not botoes_download:
                        logging.info("Nenhum botão de download encontrado na página de detalhes.")
                    else:
                        logging.info(f"Encontrados {len(botoes_download)} arquivos para download.")
                        for botao_download in botoes_download:
                            arquivos_antes = set(os.listdir(CONFIG["downloads_dir"]))
                            botao_download.click()
                            time.sleep(3)  # Aguarda o início do download.
                            arquivos_depois = set(os.listdir(CONFIG["downloads_dir"]))
                            
                            # Identifica o novo arquivo comparando o conteúdo da pasta.
                            novo_arquivo = list(arquivos_depois - arquivos_antes)
                            if novo_arquivo:
                                mover_arquivo_baixado(novo_arquivo[0], pasta_instrumento)
                            else:
                                logging.warning("Download clicado, mas nenhum novo arquivo foi detectado na pasta.")
                
                except TimeoutException:
                    logging.warning("Tabela de download não encontrada na página de detalhes.")
                finally:
                    # Clica no botão "Voltar" para retornar à lista de contratos.
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[29]/td[2]/input"))).click()
                    time.sleep(2)

            except Exception as e:
                logging.error(f"Erro ao processar uma linha da tabela de contratos: {e}")
                driver.refresh() # Tenta recarregar a página para se recuperar
                time.sleep(3)
                continue

    except Exception as e:
        logging.error(f"Erro geral ao executar ações na página de contratos: {e}")


def paginar_e_executar(driver: WebDriver, pasta_instrumento: str) -> None:
    """
    Detecta a paginação, itera por todas as páginas e executa as ações de download.
    """
    try:
        paginacao = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "/html/body/div[3]/div[15]/div[3]/div/div/form/div/div[2]/span")))
        links_paginas = paginacao.find_elements(By.TAG_NAME, "a")
        num_paginas = len(links_paginas) + 1 if links_paginas else 1
        logging.info(f"Total de {num_paginas} páginas de contratos encontradas.")

        for pagina_atual in range(1, num_paginas + 1):
            logging.info(f"Processando página {pagina_atual} de {num_paginas}...")
            if pagina_atual > 1:
                try:
                    # Encontra o link da página pelo texto e clica.
                    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, f"//a[text()='{pagina_atual}']"))).click()
                    time.sleep(2)
                except Exception as e:
                    logging.error(f"Não foi possível navegar para a página {pagina_atual}. Erro: {e}")
                    break
            
            executar_acoes_detalhar(driver, pasta_instrumento)
        
        logging.info("Processamento de todas as páginas concluído.")
    except TimeoutException:
        logging.info("Nenhuma paginação encontrada. Processando apenas a página atual.")
        executar_acoes_detalhar(driver, pasta_instrumento)
    except Exception as e:
        logging.error(f"Erro ao processar a paginação: {e}")


def main() -> None:
    """
    Função principal que orquestra todo o processo de download.
    """
    driver = conectar_navegador_existente()
    df = ler_planilha(CONFIG["input_file"])

    for instrumento in df["Instrumento nº"]:
        if pd.isna(instrumento) or not str(instrumento).strip():
            logging.warning(f"Instrumento inválido ou vazio encontrado. Pulando.")
            continue
        
        logging.info("=" * 60)
        logging.info(f"Iniciando processamento para o Instrumento: {instrumento}")
        
        pasta_instrumento = os.path.join(CONFIG["output_dir"], instrumento)
        os.makedirs(pasta_instrumento, exist_ok=True)
        
        if navegar_menu_principal(driver, instrumento):
            if acessar_aba_contratos(driver):
                paginar_e_executar(driver, pasta_instrumento)
            
            try:
                # Retorna ao menu de pesquisa para o próximo instrumento.
                WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[3]/div[2]/div[1]/a"))).click()
                time.sleep(2)
            except Exception as e:
                logging.error(f"Erro ao retornar ao menu principal. Tentando recarregar a página. Erro: {e}")
                driver.get("URL_DA_PAGINA_DE_PESQUISA") # Idealmente, usar um URL fixo para se recuperar.
                time.sleep(3)
        logging.info(f"Finalizado processamento para o Instrumento: {instrumento}")


if __name__ == "__main__":
    main()
