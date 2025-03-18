📌 Automacao de Navegacao e Download de Arquivos com Selenium

📖 Sobre o Projeto

Este projeto automatiza a navegação e a extração de arquivos a partir de um sistema web utilizando Selenium. O script conecta-se a um navegador Chrome já aberto, busca informações de uma planilha do Excel e interage com o site para baixar arquivos de documentos relacionados aos instrumentos listados na planilha.

🚀 Funcionalidades

Conexão com um navegador Chrome já aberto.

Leitura de dados a partir de uma planilha Excel.

Navegação automática dentro do sistema web.

Download de arquivos relacionados a cada instrumento.

Organização dos arquivos baixados em pastas específicas.

🛠 Tecnologias Utilizadas

Python (3.x)

Selenium para automação do navegador.

Pandas para manipulação de planilhas Excel.

Webdriver Manager para gestão automática do ChromeDriver.

📂 Estrutura do Projeto

/
├── main.py                   # Script principal
├── requirements.txt          # Lista de dependências
└── README.md                 # Este arquivo

🔧 Configuração e Execução

1️⃣ Instalação das Dependências

Antes de executar o script, instale os pacotes necessários utilizando o comando:

pip install -r requirements.txt

2️⃣ Inicie o Chrome com Depuração Remota

Abra o terminal e execute o seguinte comando para iniciar o Chrome com depuração remota:

chrome.exe --remote-debugging-port=9222 --user-data-dir="C:\\chrome_debug"

3️⃣ Execute o Script

Após configurar o Chrome, execute o script com:

python main.py

⚙️ Principais Funções do Código

🛠 Conectar ao navegador existente

def conectar_navegador_existente():
    options = webdriver.ChromeOptions()
    options.debugger_address = "localhost:9222"
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver

Conecta-se a uma instância do Chrome já aberta.

📂 Leitura da Planilha

def ler_planilha(arquivo):
    df = pd.read_excel(arquivo, engine="openpyxl")
    df["Instrumento nº"] = df["Instrumento nº"].astype(str).str.replace(r"\.0$", "", regex=True)
    return df

Lê os dados de uma planilha Excel e faz o tratamento de formatação na coluna "Instrumento nº".

🔄 Navegação no Sistema

def navegar_menu_principal(driver, instrumento):
    esperar_elemento(driver, "XPATH_DO_ELEMENTO").click()
    # Código para pesquisa do instrumento

Interage com o menu do site para buscar os instrumentos.

📄 Download de Arquivos

def executar_acoes_detalhar(driver, pasta_instrumento):
    # Localiza os botões "Detalhar", acessa a página e baixa arquivos

Identifica os botões de "Detalhar", acessa as páginas correspondentes e baixa arquivos.

🛑 Possíveis Erros e Soluções

Erro

Solução

selenium.common.exceptions.WebDriverException

Certifique-se de que o Chrome está aberto com depuração remota ativada.

pandas.errors.ParserError

Verifique se o arquivo Excel está no formato correto.

TimeoutException

Aumente o tempo de espera em WebDriverWait.

📜 Licença

Este projeto é de código aberto sob a licença MIT. Sinta-se à vontade para usá-lo e modificá-lo conforme necessário.

📩 Contato

Caso tenha dúvidas ou sugestões, entre em contato pelo e-mail: debrito521@gmail.com.

