import time
import pandas as pd
import os
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import WebDriverException

# 🛠 Conectar ao navegador já aberto
def conectar_navegador_existente():
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"
    try:
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        print("✅ Conectado ao navegador existente!")
        return driver
    except Exception as erro:
        print(f"❌ Erro ao conectar ao navegador: {erro}")
        exit()

# Função para reconectar ao navegador em caso de erro de conexão
def reconectar_navegador(driver):
    try:
        driver.quit()  # Fechar o driver atual
    except:
        pass  # Ignorar erros ao fechar o driver

    # Tentar reconectar
    return conectar_navegador_existente()

# 📂 Ler planilha de entrada
def ler_planilha(arquivo=r"C:\Users\diego.brito\Downloads\robov1\pasta1.xlsx"):
    df = pd.read_excel(arquivo, engine="openpyxl")

    # 🛠️ Remover ".0" da coluna "Instrumento nº"
    if "Instrumento nº" in df.columns:
        df["Instrumento nº"] = df["Instrumento nº"].astype(str).str.replace(r"\.0$", "", regex=True)

    return df

# 🔄 Navegar no menu principal
def esperar_elemento(driver, xpath, timeout=10):
    return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.XPATH, xpath)))

def navegar_menu_principal(driver, instrumento):
    try:
        esperar_elemento(driver, "/html/body/div[1]/div[3]/div[1]/div[1]/div[1]/div[4]").click()
        esperar_elemento(driver, "/html[1]/body[1]/div[1]/div[3]/div[2]/div[1]/div[1]/ul[1]/li[6]/a[1]").click()
        campo_pesquisa = esperar_elemento(driver,
                                          "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/input[1]")
        campo_pesquisa.clear()
        campo_pesquisa.send_keys(instrumento)
        esperar_elemento(driver,
                         "/html[1]/body[1]/div[3]/div[15]/div[3]/div[1]/div[1]/form[1]/table[1]/tbody[1]/tr[2]/td[2]/span[1]/input[1]").click()
        time.sleep(1)
        esperar_elemento(driver,
                         "/html[1]/body[1]/div[3]/div[15]/div[3]/div[3]/table[1]/tbody[1]/tr[1]/td[1]/div[1]/a[1]").click()
        return True
    except:
        print(f"⚠️ Instrumento {instrumento} não encontrado.")
        return False

# 📝 Acessar aba contratos
def acessar_aba_contratos(driver):
    try:
        esperar_elemento(driver, "/html/body/div[3]/div[15]/div[1]/div/div[1]/a[6]/div/span/span").click()
        esperar_elemento(driver, "/html/body/div[3]/div[15]/div[1]/div/div[2]/a[24]/div/span/span").click()
        print("✅ Aba contratos acessada com sucesso!")
        return True
    except Exception as erro:
        print(f"❌ Erro ao acessar a aba contratos: {erro}")
        return False

# 🚀 Navegar para uma página específica
def navegar_para_pagina(driver, numero_pagina):
    try:
        # Se for a primeira página, não é necessário navegar
        if numero_pagina == 1:
            return

        # Identificar o elemento de paginação
        paginacao = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[3]/div/div/form/div/div[2]/span")

        # Encontrar o link da página específica
        link_pagina = paginacao.find_element(By.XPATH, f".//a[text()='{numero_pagina}']")
        link_pagina.click()
        time.sleep(2)  # Esperar a página carregar
    except Exception as e:
        print(f"⚠️ Erro ao navegar para a página {numero_pagina}: {e}")

# 📄 Executar ações nos botões "Detalhar"
def executar_acoes_detalhar(driver, pagina_atual, pasta_instrumento):
    try:
        # Reidentificar a tabela a cada iteração
        tabela = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[3]/div/div/form/div/div[2]/table")

        # Encontrar todas as linhas (<tr>) da tabela
        linhas = tabela.find_elements(By.XPATH, ".//tbody/tr")

        # Se não houver linhas, sair do loop
        if not linhas:
            print("🚨 Nenhuma linha encontrada na tabela.")
            return

        for i in range(len(linhas)):  # Percorrer as linhas pelos índices
            try:
                # Reidentificar a tabela e as linhas para evitar `stale element reference`
                tabela = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[3]/div/div/form/div/div[2]/table")
                linhas = tabela.find_elements(By.XPATH, ".//tbody/tr")

                # Verificar se o índice ainda está disponível
                if i >= len(linhas):
                    break

                linha = linhas[i]

                # Encontrar o botão "Detalhar" na linha atual
                botao_detalhar = linha.find_element(By.XPATH, ".//a[contains(text(), 'Detalhar')]")
                botao_detalhar.click()
                time.sleep(2)  # Esperar a página carregar

                try:
                    # Identificar a tabela de download
                    tabela_download = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[26]/td/div[1]/table")

                    # Encontrar todos os botões de download na tabela
                    botoes_download = tabela_download.find_elements(By.XPATH, ".//a[contains(text(), 'Baixar')]")

                    # Verificar se há botões de download
                    if not botoes_download:
                        print("⚠️ Nenhum botão de download encontrado na tabela.")
                    else:
                        print(f"✅ Encontrados {len(botoes_download)} botões de download.")

                        # Clicar em todos os botões de download
                        for botao in botoes_download:
                            # Listar os arquivos antes do download
                            arquivos_antes = set(os.listdir(caminho_downloads))

                            # Clicar no botão de download
                            botao.click()
                            time.sleep(2)  # Esperar o download

                            # Listar os arquivos após o download
                            arquivos_depois = set(os.listdir(caminho_downloads))

                            # Identificar o arquivo baixado
                            arquivos_baixados = list(arquivos_depois - arquivos_antes)
                            if arquivos_baixados:
                                for arquivo_baixado in arquivos_baixados:
                                    # Mover o arquivo baixado para a pasta do instrumento
                                    mover_arquivo_baixado(arquivo_baixado, pasta_instrumento)
                            else:
                                print("⚠️ Nenhum arquivo novo foi baixado.")
                except Exception as e:
                    print(f"⚠️ Erro ao processar a tabela de download: {e}")

                # Botão voltar
                esperar_elemento(driver,
                                 "/html/body/div[3]/div[15]/div[4]/div[1]/div/form/table/tbody/tr[29]/td[2]/input").click()
                time.sleep(2)  # Esperar a página carregar novamente

                # Voltar para a página atual após clicar em "Voltar"
                navegar_para_pagina(driver, pagina_atual)
            except WebDriverException as e:
                print(f"⚠️ Erro de conexão ao processar uma linha: {e}")
                driver = reconectar_navegador(driver)  # Reconectar ao navegador
                continue  # Continuar para a próxima linha em caso de erro
            except Exception as e:
                print(f"⚠️ Erro ao processar uma linha: {e}")
                continue  # Continuar para a próxima linha em caso de erro

        print(f"✅ Todos os botões 'Detalhar' da página {pagina_atual} foram processados com sucesso!")
    except Exception as erro:
        print(f"❌ Erro ao executar ações nos botões 'Detalhar': {erro}")

def paginar_e_executar(driver, pasta_instrumento):
    try:
        # Identificar o elemento de paginação
        paginacao = esperar_elemento(driver, "/html/body/div[3]/div[15]/div[3]/div/div/form/div/div[2]/span")

        # Extrair o texto do elemento de paginação
        texto_paginacao = paginacao.text

        # Exemplo de texto_paginacao: "Páginas 1, 2"
        # Extrair o número total de páginas
        partes = texto_paginacao.split(",")
        num_paginas = int(partes[-1].strip())  # Pega o último número após a vírgula

        print(f"📄 Total de páginas encontradas: {num_paginas}")

        # Loop para percorrer todas as páginas
        for pagina in range(1, num_paginas + 1):
            print(f"📄 Processando página {pagina} de {num_paginas}...")

            # Navegar para a página atual
            navegar_para_pagina(driver, pagina)

            # Executar as ações na página atual (ex: clicar em "Detalhar")
            executar_acoes_detalhar(driver, pagina, pasta_instrumento)

        print("✅ Todas as páginas foram processadas com sucesso!")
        return True
    except Exception as erro:
        print(f"❌ Erro ao processar a paginação: {erro}")
        return False

def criar_pasta_instrumento(instrumento):
    # Caminho base para as pastas de contratos
    caminho_base = r"C:\Users\diego.brito\Downloads\robov1\Teste Contratos"

    # Criar a pasta do instrumento
    pasta_instrumento = os.path.join(caminho_base, instrumento)
    if not os.path.exists(pasta_instrumento):
        os.makedirs(pasta_instrumento)
        print(f"✅ Pasta criada para o instrumento {instrumento} em {pasta_instrumento}")

    return pasta_instrumento

def mover_arquivo_baixado(arquivo_baixado, pasta_instrumento):
    # Caminho padrão de downloads
    caminho_downloads = r"C:\Users\diego.brito\Downloads"

    # Mover o arquivo baixado para a pasta do instrumento
    shutil.move(os.path.join(caminho_downloads, arquivo_baixado), os.path.join(pasta_instrumento, arquivo_baixado))
    print(f"✅ Arquivo {arquivo_baixado} movido para {pasta_instrumento}")

# Caminho padrão de downloads
caminho_downloads = r"C:\Users\diego.brito\Downloads"

# Exemplo de uso
if __name__ == "__main__":
    driver = conectar_navegador_existente()
    df = ler_planilha()
    for instrumento in df["Instrumento nº"]:
        # Verificar se o instrumento é um número válido
        if pd.isna(instrumento) or not str(instrumento).strip().isdigit():
            print(f"⚠️ Instrumento '{instrumento}' inválido. Pulando para o próximo.")
            continue  # Pular para o próximo instrumento

        # Criar pasta para o instrumento
        pasta_instrumento = criar_pasta_instrumento(instrumento)

        if navegar_menu_principal(driver, instrumento):
            if acessar_aba_contratos(driver):
                paginar_e_executar(driver, pasta_instrumento)

                # Clicar no botão para voltar ao menu principal
                try:
                    esperar_elemento(driver, "/html/body/div[3]/div[2]/div[1]/a").click()
                    time.sleep(2)  # Esperar a página carregar
                    print("✅ Retornou ao menu principal para o próximo instrumento.")
                except Exception as e:
                    print(f"⚠️ Erro ao clicar no botão de retorno: {e})")