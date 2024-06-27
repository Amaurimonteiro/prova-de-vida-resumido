import os

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

import sys
import os
sys.path.append(r'C:\pdfs_Prova_de_Vida')


# Caminho para salvar o arquivo PDF
save_path = os.path.join(os.getcwd(), 'saida_andrey.pdf')

# Configurações do Chrome para desativar a janela de impressão e salvar a página como PDF
chrome_options = Options()
chrome_options.add_experimental_option('prefs', {
    "printing.print_preview_sticky_settings.appState": '{"recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}], "selectedDestinationId": "Save as PDF", "version": 2}',
    "savefile.default_directory": save_path
})
chrome_options.add_argument('--kiosk-printing')  # Ativa a impressão em modo quiosque

# Inicializando o WebDriver
driver = webdriver.Chrome(options=chrome_options)

# Abre a página que você deseja imprimir
driver.get('https://www1.siapenet.gov.br/')

# Aciona a impressão para PDF
driver.execute_script('window.print();')

# Aguarda alguns segundos para garantir que a página foi "impressa" como PDF
import time

time.sleep(2)

# Fecha o navegador
driver.quit()

# O arquivo output.pdf deve estar na pasta especificada
print(f'Página salva como PDF em: {save_path}')

