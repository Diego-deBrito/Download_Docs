# RPA para Download em Massa de Documentos de Contratos

![Python](https://img.shields.io/badge/Python-3.7+-blue.svg)
![Libraries](https://img.shields.io/badge/Libraries-Selenium%20%7C%20Pandas-blue)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)

## Descrição do Projeto

Este projeto é um robô de automação (RPA) desenvolvido em Python e Selenium, cuja finalidade é automatizar o download em massa de todos os documentos associados a uma lista de contratos ou convênios de um portal web.

O script lê uma lista de "Instrumentos" de uma planilha Excel, cria uma estrutura de pastas local e, para cada instrumento, navega até a sua respectiva área no portal, percorre todas as páginas de contratos relacionados e baixa sistematicamente todos os arquivos disponíveis, organizando-os automaticamente nas pastas criadas.

## Funcionalidades Principais

-   **Download em Massa e Organização Automática:** Baixa múltiplos arquivos de diferentes páginas e os organiza em diretórios nomeados de acordo com o instrumento, eliminando a necessidade de trabalho manual.

-   **Técnicas Avançadas de Selenium:**
    -   **Prevenção de `StaleElementReferenceException`:** Re-identifica dinamicamente os elementos da página dentro de loops, uma técnica essencial para interagir com interfaces que se atualizam após cada ação (como clicar em "Voltar").
    -   **Detecção Inteligente de Downloads:** Identifica o arquivo recém-baixado comparando o conteúdo do diretório de downloads antes e depois da ação de clique, uma solução robusta para quando os nomes dos arquivos não são conhecidos previamente.

-   **Mecanismo de Paginação:** Detecta automaticamente o número de páginas em uma lista de resultados e itera sobre todas elas, garantindo que nenhum contrato ou documento seja esquecido.

-   **Conexão com Navegador Existente:** Conecta-se a uma sessão do Google Chrome em modo de depuração, permitindo que o usuário realize o login e a autenticação manualmente antes de iniciar a automação.

-   **Logging Estruturado:** Todas as ações, sucessos e erros são registrados em um arquivo de log (`downloader_log.txt`), fornecendo um histórico detalhado para auditoria e depuração.

## Pré-requisitos

-   [Python 3.7](https://www.python.org/downloads/) ou superior
-   [Google Chrome](https://www.google.com/chrome/) (navegador web)

## Instalação e Configuração

1.  **Clone o repositório:**
    ```bash
    git clone [https://github.com/seu-usuario/seu-repositorio.git](https://github.com/seu-usuario/seu-repositorio.git)
    cd seu-repositorio
    ```

2.  **Crie um ambiente virtual (recomendado):**
    ```bash
    # Windows
    python -m venv venv
    .\venv\Scripts\activate
    ```

3.  **Instale as dependências:**
    Crie um arquivo `requirements.txt` com o seguinte conteúdo:
    ```
    pandas
    openpyxl
    selenium
    webdriver-manager
    ```
    Em seguida, instale as bibliotecas:
    ```bash
    pip install -r requirements.txt
    ```

4.  **Configure os Caminhos:**
    Abra o script Python e edite o dicionário `CONFIG` no início do arquivo. **Esta é a etapa mais importante.**
    ```python
    CONFIG: Dict[str, Any] = {
        "chrome_debug_port": "9222",
        "input_file": r"C:\Caminho\Completo\Para\Sua\planilha_de_entrada.xlsx",
        "downloads_dir": r"C:\Caminho\Completo\Para\Sua\pasta_de_downloads_padrao",
        "output_dir": r"C:\Caminho\Completo\Para\onde\as_pastas_dos_contratos_serao_criadas"
    }
    ```

## Como Executar

1.  **Prepare a Planilha de Entrada:**
    Garanta que o arquivo Excel especificado em `input_file` exista e contenha uma coluna chamada `"Instrumento nº"`.

2.  **Inicie o Google Chrome em Modo de Depuração:**
    Feche todas as janelas do Chrome e inicie uma nova através do terminal com o comando abaixo.
    ```bash
    # Windows (ajuste o caminho se necessário)
    "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222
    ```

3.  **Acesse o Sistema Manualmente:**
    Na janela do Chrome que abriu, navegue até o portal, faça seu login e deixe-o pronto na página principal.

4.  **Execute o Script:**
    Abra um terminal na pasta do projeto e execute o script:
    ```bash
    python nome_do_script.py
    ```
    O robô começará a criar as pastas e a baixar os arquivos. Acompanhe o progresso pelo console ou pelo arquivo `downloader_log.txt`.

## Observações Importantes

> **Fragilidade dos Seletores (XPath):** Os seletores XPath utilizados neste script são absolutos e podem quebrar se a estrutura do site for alterada. Para uma automação mais duradoura, é fortemente recomendado substituí-los por seletores mais robustos (como IDs, classes ou XPaths relativos).

> **Ajuste de `time.sleep()`:** O script utiliza pausas fixas (`time.sleep()`). Em conexões de internet mais lentas ou sistemas mais sobrecarregados, pode ser necessário aumentar a duração dessas pausas para garantir que as páginas carreguem completamente.
