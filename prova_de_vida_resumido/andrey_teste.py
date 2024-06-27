# pip install PyQt5


import sys
from PyQt5.QtWidgets import QComboBox, QApplication, QMainWindow, QPushButton, QWidget, QDesktopWidget, QLabel, QFileDialog, QMessageBox, QGridLayout
from PyQt5.QtCore import Qt
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException, ElementNotInteractableException, NoSuchFrameException
from time import sleep
import sys
import os
# sys.path.append(r'C:\Users\ufam\OneDrive - mtegovbr\Documentos\decipyx\alexandria')
sys.path.append(r'C:\Users\OneDrive - mtegovbr\Documentos\decipyx\alexandria')

from socratica.web_automation import WebAutomation as web_auto
import socratica.estilos as estilo
from socratica.data_handling import DataHandling as dat_hand

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains

# # Configura o nível de logging para DEBUG
# logging.basicConfig(level=logging.DEBUG)

class App(QMainWindow):

    def __init__(self):
        super().__init__()      
        self.title = 'AutoBotESIAPE v2.5 - Decipyx '
        self.left = 10
        self.top = 10
        self.width = 400
        self.height = 200
        self.filePath = r'C:\pdfs_Prova_de_Vida'  # Para armazenar o caminho do arquivo selecionado
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        layout = QGridLayout()
        layout.setSpacing(10)  # Espaçamento entre os widgets
        self.setLayout(layout)

        # Crie um QLabel com texto clicável
        clickable_text = QLabel('<a style="text-decoration: none; color: #F3F3F3;" href="#">Sobre</a>', self)
        clickable_text.setAlignment(Qt.AlignRight)
        clickable_text.setOpenExternalLinks(False)  # Não abra um navegador externo
        clickable_text.linkActivated.connect(self.show_description)
        layout.addWidget(clickable_text)

        self.uploadButton = QPushButton('dados_beneficiarios.xlsx', self)
        self.uploadButton.clicked.connect(self.openFileNameDialog)
        layout.setContentsMargins(10, 10, 10, 10)  # Margens do layout
        layout.addWidget(self.uploadButton)

        # Adicionando um label para mostrar o caminho do arquivo selecionado
        self.filePathLabel = QLabel(self)
        file_name = os.path.basename(self.filePath)
        self.filePathLabel.setText(f"Arquivo selecionado: <b>{file_name}</b>")
        layout.addWidget(self.filePathLabel)

        self.orgao_to_xpath = {
            '20113': '//*[@id="TGROW105_0"]/td[2]/div',
            '17000': '//*[@id="TGROW105_1"]/td[2]/div',
            '40801': '//*[@id="TGROW105_5"]/td[2]/div',
            '40802': '//*[@id="TGROW105_7"]/td[2]/div',
            '40803': '//*[@id="TGROW105_10"]/td[2]/div',
            '40804': '//*[@id="TGROW105_14"]/td[2]/div',
            '40805': '//*[@id="TGROW105_18"]/td[2]/div',
            '40806': '//*[@id="TGROW105_20"]/td[2]/div',
        }

        # Adicionando uma lista suspensa (ComboBox)
        self.comboBoxLabel = QLabel('Selecione abaixo, o Órgão para rodar o programa',self)
        self.comboBox = QComboBox(self)
        list_orgao = [
                        '20113',
                        '17000',
                        '40801',
                        '40802',
                        '40803',
                        '40804',
                        '40805',
                        '40806',
                     ]
        
        for orgaos in list_orgao:
            self.comboBox.addItem(orgaos)

        layout.addWidget(self.comboBox, 5, 0)
        layout.addWidget(self.comboBoxLabel, 4, 0)

        # Conecta a mudança de valor da lista suspensa a uma função
        self.comboBox.currentIndexChanged.connect(self.on_combobox_changed)

        button = QPushButton('Iniciar', self)
        button.clicked.connect(self.on_click)
        layout.addWidget(button)

        self.itemsRemainingLabel = QLabel(self)
        layout.addWidget(self.itemsRemainingLabel)

        self.label = QLabel("Esperando...", self)
        layout.addWidget(self.label)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # Centralizando a janela
        screen = QDesktopWidget().screenGeometry()
        window = self.geometry()
        self.move(int((screen.width() - window.width()) / 2), int((screen.height() - window.height()) / 2))

        self.show()

    def show_description(self):
        desc_text = """
        Sobre esta Automação

        Este é um assistente automatizado projetado para facilitar o processo de atualização da folha de pagamento no portal siape.sigepe.gov.br. 
        Através desta ferramenta, você pode rapidamente atualizar as informações de pagamento de uma lista de servidores. 
        Para utilizar esta automação, basta fornecer uma lista de servidores em formato .xlsx.

        A ferramenta irá ler o arquivo, acessar o portal SIAPE e realizar as atualizações necessárias para cada servidor listado. 
        Isso economiza tempo e garante precisão no processo de atualização.

        Por favor, certifique-se de que o arquivo .xlsx esteja formatado corretamente e que você tenha acesso adequado ao portal antes de iniciar a automação.
        """
        msgBox = QMessageBox(self)
        msgBox.setWindowTitle("Descrição")
        msgBox.setText(desc_text)     
        msgBox.exec_()

    def on_combobox_changed(self, i):
        orgao = self.comboBox.currentText()
        print(f"O valor selecionado é {orgao}")

    def openFileNameDialog(self):
        options = QFileDialog.Options()
        filePath, _ = QFileDialog.getOpenFileName(self, "Selecione o arquivo .xlsx", "", "Text Files (*.xlsx);;All Files (*)", options=options)
        if filePath:
            self.filePath = filePath
            file_name = os.path.basename(self.filePath)
            self.filePathLabel.setText(f"Arquivo selecionado: <b>{file_name}</b>")
    
    def show_alert(self, title, message):
        alert = QMessageBox(self)
        alert.setWindowTitle(title)
        alert.setText(message)
        self.setStyleSheet(estilo.dark_stylesheet)
        alert.setIcon(QMessageBox.Warning)
        alert.exec_()

    def show_finish(self, title, message, options):
        finish = QMessageBox()
        finish.setWindowTitle(title)
        finish.setText(message)
        finish.setInformativeText(options)
        continue_button         = finish.addButton("Continuar", QMessageBox.AcceptRole)
        close_program_button    = finish.addButton("Fechar PayBot", QMessageBox.RejectRole)
        close_both_button       = finish.addButton("Fechar Ambos", QMessageBox.DestructiveRole)
        self.setStyleSheet(estilo.dark_stylesheet)
        finish.exec()

        clicked_button = finish.clickedButton()
        
        if clicked_button == continue_button:
            # Continuar com o programa aberto
            pass
        elif clicked_button == close_program_button:
            # Fechar o programa e deixar o navegador aberto
            self.close()
        elif clicked_button == close_both_button:
            # Fechar ambos o programa e o navegador
            browser.quit()  # Fechar o navegador usando Selenium
            self.close()

    def entrar_em_ficha_pensionita(self):
        self.auto = web_auto(browser)
        try:
            caminho_pesquisar = '//*[@id="ICONIMG15"]'
            self.auto.default_clica_elemento_by_xpath('WA0', caminho_pesquisar)
            sleep(2)
            self.auto.inserir_texto_enter_by_id_in_iframe('WA0', 'F_63', "FPCOPSFICF")
            sleep(2)
        except (NoSuchElementException, ElementNotInteractableException):  # ou outra exceção específica
            try:
                caminho_pesquisar = '//*[@id="ICONIMG18"]'
                self.auto.default_clica_elemento_by_xpath_com_scroll('WA0', caminho_pesquisar)
                sleep(3)
                self.auto.inserir_texto_enter_by_id_in_iframe('WA0', 'F_66', "FPCOPSFICF")
                sleep(2)
            except (NoSuchElementException, ElementNotInteractableException):           
                caminho_pesquisar = '//*[@id="ICONIMG16"]'
                self.auto.default_clica_elemento_by_xpath_com_scroll('WA1', caminho_pesquisar)
                sleep(3)
                self.auto.inserir_texto_enter_by_id_in_iframe('WA1', 'F_63', "FPCOPSFICF")
                sleep(2)
          
    def entrar_em_ficha_servidor(self):
        self.auto = web_auto(browser)
        try:
            caminho_pesquisar = '//*[@id="ICONIMG15"]'
            self.auto.default_clica_elemento_by_xpath('WA0', caminho_pesquisar)
            sleep(1)

            self.auto.inserir_texto_enter_by_id_in_iframe('WA0', 'F_63', "FPCOFICHAF")
            sleep(1)
        except (NoSuchElementException, ElementNotInteractableException):  # ou outra exceção específica
            try:
                caminho_pesquisar = '//*[@id="ICONIMG16"]'
                self.auto.default_clica_elemento_by_xpath_com_scroll('WA2', caminho_pesquisar)
                sleep(3)
                self.auto.inserir_texto_enter_by_id_in_iframe('WA2', 'F_64', "FPCOFICHAF")
                sleep(2)
            except (NoSuchElementException, ElementNotInteractableException, NoSuchFrameException):       
                caminho_pesquisar = '//*[@id="ICONIMG18"]'
                self.auto.default_clica_elemento_by_xpath_com_scroll('WA1', caminho_pesquisar)
                sleep(3)
                try:
                    self.auto.inserir_texto_enter_by_id_in_iframe('WA1', 'F_61', "FPCOFICHAF")
                    sleep(2)
                except (NoSuchElementException, ElementNotInteractableException, NoSuchFrameException):
                    self.auto.inserir_texto_enter_by_id_in_iframe('WA1', 'F_66', "FPCOFICHAF")
                    sleep(2)
    
    def clicar_imprimir_ficha_servidor(self):
        self.auto = web_auto(browser)

        browser.switch_to.default_content()
        browser.switch_to.frame('WA1')
        elemento = browser.find_element(By.ID, 'S_95')
        nome_arquivo = elemento.text
        print("Matrícula extraída:", nome_arquivo)

        caminho_imprimir = '//*[@id="B_175"]'
        self.auto.default_clica_elemento_by_xpath('WA1', caminho_imprimir)
        sleep(5)
        
        # Armazene o ID da janela original
        janela_principal = browser.current_window_handle

        # Mude para a nova janela (a última na lista)
        all_windows = browser.window_handles
        print("URL atual:", browser.current_url)
        browser.switch_to.window(all_windows[-1])
        print("Mudei para a nova janela::", browser.current_url)
        print('setp3')

        browser.maximize_window()
        sleep(5)  # Aumentar o tempo de espera

        browser.switch_to.frame("WA0")
        print('setp4')

        sleep(3)
        # Depurar
        browser.switch_to.frame("SUBPAGE4")
        print('setp5')

        wait = WebDriverWait(browser, 5)
        try:
            wait.until(EC.element_to_be_clickable((By.ID, 'open-button'))).click()
            print('setp6')
        except TimeoutException as e:
            print(f"Ainda não, parceiro: {e}")
    
        sleep(10)
        # Pega o diretório atual do script
        diretorio_download = os.path.dirname(os.path.realpath(__file__))
        nome_arquivo_destino = os.path.join(diretorio_download, nome_arquivo + ".pdf")

        # Assumindo que você sabe o nome original do arquivo baixado
        nome_arquivo_original = os.path.join(diretorio_download, "StartDynamicContent.pdf")

        # Renomeando o arquivo
        os.rename(nome_arquivo_original, nome_arquivo_destino)

        sleep(2)
        # Fechar a janela de download
        browser.close()

        # Volte para a janela principal
        browser.switch_to.window(janela_principal)

    def clicar_imprimir_ficha_pensionista(self):
        self.auto = web_auto(browser)

        browser.switch_to.default_content()
        browser.switch_to.frame('WA0')
        elemento = browser.find_element(By.ID, 'S_88')
        nome_arquivo = elemento.text
        print("Nome extraído:", nome_arquivo)

        caminho_imprimir = '//*[@id="B_175"]'
        self.auto.default_clica_elemento_by_xpath('WA0', caminho_imprimir)
        sleep(5)
        
        # Armazene o ID da janela original
        janela_principal = browser.current_window_handle

        # Mude para a nova janela (a última na lista)
        all_windows = browser.window_handles
        print("URL atual:", browser.current_url)
        browser.switch_to.window(all_windows[-1])
        print("Mudei para a nova janela::", browser.current_url)
        print('setp3')

        browser.maximize_window()
        sleep(5)  # Aumentar o tempo de espera

        browser.switch_to.frame("WA0")
        print('setp4')

        sleep(3)
        # Depurar
        browser.switch_to.frame("SUBPAGE4")
        print('setp5')

        wait = WebDriverWait(browser, 5)
        try:
            wait.until(EC.element_to_be_clickable((By.ID, 'open-button'))).click()
            print('setp6')
        except TimeoutException as e:
            print(f"Ainda não, parceiro: {e}")
    
        sleep(10)
        # Pega o diretório atual do script
        diretorio_download = os.path.dirname(os.path.realpath(__file__))
        nome_arquivo_destino = os.path.join(diretorio_download, nome_arquivo + ".pdf")

        # Assumindo que você sabe o nome original do arquivo baixado
        nome_arquivo_original = os.path.join(diretorio_download, "StartDynamicContent.pdf")

        # Renomeando o arquivo
        os.rename(nome_arquivo_original, nome_arquivo_destino)

        sleep(2)
        # Fechar a janela de download
        browser.close()

        # Volte para a janela principal
        browser.switch_to.window(janela_principal)
    
    def robo(self):
        global browser
        error_count = 0  # Variável para rastrear o número de erros
        chrome_options = webdriver.ChromeOptions()
        url = 'https://esiape.sigepe.gov.br/modsiape/servlet/StartCISPage?PAGEURL=/cisnatural/NatLogon.html&xciParameters.natsession=modsiape'

        # Pega o diretório atual do script
        current_directory = os.path.dirname(os.path.realpath(__file__))

        prefs = {
            "download.default_directory": current_directory,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True
        }

        chrome_options.add_experimental_option('prefs', prefs)

        browser = webdriver.Chrome(options=chrome_options)

        self.auto = web_auto(browser)
        self.data = dat_hand()
        self.auto.acessar_site_chrome(url, 5)

        selected_orgao = self.comboBox.currentText()
        elemento_tabela = self.orgao_to_xpath.get(selected_orgao, None)   
        print(elemento_tabela)     
        iframe_elemento = 'WA0'
        path_elemento = '//*[@id="ICONIMG15"]'
        iframe_text = 'WA0'
        element_id = 'F_63'
        caminho_avanca_comunicacao_serpro = '//*[@id="B_12"]'
        self.auto.default_clica_elemento_by_xpath('WA1', caminho_avanca_comunicacao_serpro)
        sleep(2)

        try:
            browser.window_handles
            caminho_avancar_mensagens = '//*[@id="B_377"]'
            self.auto.default_clica_elemento_by_xpath('WA2', caminho_avancar_mensagens)
            sleep(2)
            self.auto.troca_habilitacao(iframe_elemento, path_elemento, iframe_text, element_id, 'TROCAHAB', elemento_tabela)
            sleep(2)
            self.entrar_em_ficha_servidor()
            # self.entrar_em_ficha_pensionita()
        except (NoSuchElementException, ElementNotInteractableException): 
            self.auto.troca_habilitacao(iframe_elemento, path_elemento, iframe_text, element_id, 'TROCAHAB', elemento_tabela)
            sleep(2)
            # self.entrar_em_ficha_servidor()
            # self.entrar_em_ficha_pensionita()
        df_list = [self.filePath] 
        file_xlsx = self.data.arquivos_leitura(df_list)
        sleep(2)
        # Lendo os valores do arquivo txt
        colunas_manter = [
                            'Siape S/I',
                            'Siape P',
                         ]
        df_filter = self.data.filtrar_colunas(file_xlsx['Soldao_ON4_Dados_aleatórios'], colunas_manter)
        valores_float = self.data.le_e_filtra_excel(df_filter)
        valores = self.data.converte_float_list_para_int_list(valores_float)
        print(valores)
        
        for index, valor in enumerate(valores):
            remaining = len(valores) - (index + 1)
            self.itemsRemainingLabel.setText(f"Restam {remaining} itens para serem processados. Erros: {error_count}")
            QApplication.processEvents()  
            success = False  

            while not success:
                try:
                    # self.entrar_em_ficha_servidor()
                    # sleep(2)
                    print('setp7')

                    self.auto.inserir_texto_enter_by_id_in_iframe('WA2', 'F_131', valor)
                    sleep(2)
                    self.clicar_imprimir_ficha_servidor()
                    sleep(2)
                    self.entrar_em_ficha_servidor()
                    sleep(2)
                    # Verifica se existe uma mensagem de erro
                    try:
                        print('setp8')

                        # self.auto.achar_elemento_by_xpath('WA1', '//*[@id="msg-div-73"]')
                        self.auto.achar_elemento_by_xpath('WA1', '//*[@id="msg-summary-76"]')
                        
                        # Se chegou aqui, significa que um dos elementos de erro foi encontrado
                        # Entre no iframe correto e tente buscar novamente
                        self.entrar_em_ficha_pensionita()
                        # self.entrar_em_ficha_servidor()
                        self.auto.inserir_texto_enter_by_id_in_iframe('WA2', 'F_125', valor)
                        # self.auto.inserir_texto_enter_by_id_in_iframe('WA2', 'F_131', valor)
                        sleep(2)
                        self.clicar_imprimir_ficha_pensionista()
                        sleep(2)
                        self.entrar_em_ficha_servidor()
                        # self.entrar_em_ficha_pensionita()
                        sleep(2)
                        
                    except (NoSuchElementException, ElementNotInteractableException, NoSuchFrameException):
                        # Se entrar neste except, significa que os elementos de erro não foram encontrados e a inserção foi bem-sucedida
                        pass

                    success = True

                except (NoSuchElementException, ElementNotInteractableException, NoSuchFrameException):
                    error_count += 1
                    self.itemsRemainingLabel.setText(f"Restam {remaining} itens para serem processados. Erros: {error_count}")
                    QApplication.processEvents()
                    self.entrar_em_ficha_servidor()
                    # self.entrar_em_ficha_pensionita()
                    sleep(2)
                    break
 
        self.show_finish("Finalizado", "A execução finalizou.", "Deseja continuar com o Sigepe e o ExtraBot abertos?")

    def on_click(self):
        if not self.filePath:
            self.show_alert("Erro", "Por favor, selecione um arquivo .xlsx antes de prosseguir.")
        else:
            self.robo()
        return        

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyleSheet(estilo.dark_stylesheet)
    ex = App()
    sys.exit(app.exec_())