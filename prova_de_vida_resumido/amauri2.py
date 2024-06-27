import os
 
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import sys
import os

# sys.path.append(r'C:\Users\ufam\OneDrive - mtegovbr\Documentos\decipyx\alexandria')
sys.path.append(r'C:\prova de vida resumido\prova_de_vida_resumido')
 
# Pega o diretório atual do script
diretorio_atual = os.path.dirname(os.path.realpath(__file__))
download_path = os.path.join(diretorio_atual, 'saida_Andrey.pdf')
 
# Configurações do Chrome
chrome_options = Options()
chrome_options.add_experimental_option('prefs', {
    "printing.print_preview_sticky_settings.appState": '{"recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}], "selectedDestinationId": "Save as PDF", "version": 2}',
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "plugins.always_open_pdf_externally": True  # Para abrir PDFs automaticamente no visualizador de PDF padrão
})
chrome_options.add_argument('--kiosk-printing')  # Ativa a impressão em modo quiosque
 
# Caminho para o WebDriver (ajuste conforme necessário)
service = Service(ChromeDriverManager().install())
 
# Inicializando o WebDriver
driver = webdriver.Chrome(options=chrome_options, service=service)
 
# Abre a página que você deseja imprimir
driver.get('https://www1.siapenet.gov.br/')
 
# Aciona a impressão para PDF
driver.execute_script('window.print();')
 
# Aguarda alguns segundos para garantir que a página foi "impressa" como PDF
import time
 
time.sleep(15)
 
# Fecha o navegador
driver.quit()
 
# Informa sobre o salvamento
print(f'Página salva como PDF em: {download_path}')
 