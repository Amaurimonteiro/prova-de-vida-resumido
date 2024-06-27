from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import random
from selenium.common.exceptions import NoSuchElementException, TimeoutException, NoSuchFrameException, ElementNotInteractableException, StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains
import locale
from datetime import datetime
from selenium.webdriver.support.select import Select
from num2words import num2words
import os
import glob
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium import webdriver

from webdriver_manager.chrome import ChromeDriverManager

from pprint import pprint
import time
from tempo_aleatorio import TempoAleatorio

import requests
from bs4 import BeautifulSoup 
import pandas as pd 
import re
import siapeAmauri

class WebAutomation_Amauri:
    
    def __init__(self, pegWeb):  # browser
                
        servico = Service(ChromeDriverManager().install())
        #navegador = webdriver.Chrome(service = servico)
        
        #navegador.get(pegWeb)
        
        Options = webdriver.ChromeOptions()
        print(Options)
        
        # Pega o diretório atual do script
        current_directory = os.path.dirname(os.path.realpath(__file__))
        prefs = {
            "download.default_directory": current_directory,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "plugins.always_open_pdf_externally": True
        }
        print(prefs)
        print('1')
        Options.add_experimental_option('prefs', prefs)
        
        #chrome_options.add_experimental_option('prefs', prefs)
        print('2')
        
        navegador = webdriver.Chrome(options=Options, service = servico)       
                
        #pagina = webdriver.Chrome(options=chrome_options)
        print('3')
        self.browser = navegador
        self.exception  = (NoSuchElementException, ElementNotInteractableException, NoSuchFrameException, TimeoutException)
        print('4')
        
        navegador.get(pegWeb)
        
        tempo = TempoAleatorio()
        time.sleep(tempo.tempo_aleatorio())  # Simula comportamento humano


    def trocar_pagina(self): #iframe
        
        # Encontrar o elemento no iframe
        try:
            #self.browser.get('https://loterias.caixa.gov.br/Paginas/lotofacil.aspx')
            print('Nome primeira pagina = ' + self.browser.title)
            self.browser.clica_elemento_by_xpath('//*[@id="imprime"]/a/img') # 
            
            time.sleep(5)
            #print(self.browser.window_handles)
            #self.browser.find_element(By.XPATH, '//*[@id="sidebar"]//print-preview-button-strip//div/cr-button[1]').click()
            for url in self.browser.window_handles:
                print(url)
                time.sleep(2)
                
            newUrl = self.browser.window_handles[1]
            self.browser.switch_to.window(newUrl)
            print('Nome da segunda pagina = ' + self.browser.title)
            self.browser.close()
            print('deu certo')
        except:
            # Tratamento para o caso em que o arquivo não é encontrado
            #print(f"Erro: Arquivo {arquivo} não encontrado. Detalhes: {e}")
            print('deu errado')
            return False
        else:
            # Código a ser executado se nenhum erro ocorrer durante a leitura do arquivo
            # print("Leitura do arquivo concluída com sucesso.")                         
            return busca


    def ler_elemento_enter_by_xpath(self, path): #iframe
        """
        Clica em um elemento localizado por XPath dentro de um iframe após realizar um scroll.

        Args:
            iframe (str): Nome ou ID do iframe onde o elemento está localizado.
            path (str): Caminho XPath do elemento.

        Returns:
            None
        """
        # Encontrar o elemento no iframe
        try:
            busca = self.browser.find_element(By.XPATH, path).text
        except:
            # Tratamento para o caso em que o arquivo não é encontrado
            #print(f"Erro: Arquivo {arquivo} não encontrado. Detalhes: {e}")
            return False
        else:
            # Código a ser executado se nenhum erro ocorrer durante a leitura do arquivo
            # print("Leitura do arquivo concluída com sucesso.")                         
            return busca
    
    
    def ler_elemento_enter_by_id(self, path): #iframe
        """
        Clica em um elemento localizado por XPath dentro de um iframe após realizar um scroll.

        Args:
            iframe (str): Nome ou ID do iframe onde o elemento está localizado.
            path (str): Caminho XPath do elemento.

        Returns:
            None
        """
        # Encontrar o elemento no iframe
        try:
            busca = self.browser.find_element(By.ID, path).text
        except:
            # Tratamento para o caso em que o arquivo não é encontrado
            #print(f"Erro: Arquivo {arquivo} não encontrado. Detalhes: {e}")
            return False
        else:
            # Código a ser executado se nenhum erro ocorrer durante a leitura do arquivo
            # print("Leitura do arquivo concluída com sucesso.")                         
            return busca
    
        
    def pegar_tabela(self,cabecalho, colunas, tabela):  # iframe
        """
          Insere um texto em um elemento localizado por ID dentro de um iframe e pressiona a tecla ENTER.
            Args:
                iframe (str): Nome ou ID do iframe onde o elemento está localizado.
                element_id (str): ID do elemento.
                text (str): Texto a ser inserido no elemento.
            Returns:
                None
        """
        #print(cabecalho)
        #print(colunas)
        #print(tabela)        
        
        a = 0
        colunas_num = []
        
        for i in colunas:
            colunas_num.append(a)
            a = a + 1
            #print(a)
        
        #print(colunas_num)

        #x = self.browser.ler_elemento_enter_by_xpath(tabela)
        x = self.browser.find_element(By.XPATH, tabela).text
        vezes = re.findall('\n', x + '\n')
        
        #print(x)
        #print(len(vezes))

        a = 0
        linhas = []
        for vez in vezes:
            linhas.append(a)
            a = a + 1
            #print(linhas)                  # passei aqui
                
        tab = []
        tudo = []
        for linha in linhas:
            tab = []
            tudo.append(tab)
                    
            '''print(colunas_num)
            print( 'passei aqui')
            print(colunas)
            print(linha)'''
            
            for coluna_num in colunas_num:
                busca = colunas[coluna_num]  #.replace(':' + str('0') + ':j', ':' + str(linha) + ':j')
                #print(busca)
                #print(busca.find(':da'))
                if busca.find(':da') > 0:     # in busca:
                    busca = colunas[coluna_num].replace(':' + str('0') + ':da', ':' + str(linha) + ':da')
                    #print('aqui tem da')
                else:                    
                    busca = colunas[coluna_num].replace(':' + str('0') + ':j', ':' + str(linha) + ':j')
                    #print('aqui tem j')
                    
                #print(busca)
                tab.append(self.browser.find_element(By.XPATH, busca).text) # colunas[coluna_num]
                #busca = colunas[coluna_num].replace(':' + str(linha) + ':', ':' + str(linha + 1) + ':')
                    
            #for i in colunas_num:
            #    colunas_
            # 
            # benef[i] = colunas_benef[i].replace(':' + str(linha) + ':', ':' + str(linha + 1) + ':')
        
        df_benef = pd.DataFrame(tudo, columns=cabecalho)
        df_benef = df_benef.dropna(how='all')

        return df_benef


    def clica_elemento_by_xpath(self, path):
        """
            Clica em um elemento localizado por XPath dentro de um iframe.

            Args:
                iframe (str): Nome ou ID do iframe onde o elemento está localizado.
                path (str): Caminho XPath do elemento.

            Returns:
                None
        """
        #self.browser.switch_to.default_content()
        # self.browser.switch_to.frame(iframe)
        self.browser.find_element(By.XPATH, path).click()
        #self.browser.switch_to.default_content()

    def inserir_texto_nome(self, element_name, text):  # iframe
        """
          Insere um texto em um elemento localizado por ID dentro de um iframe e pressiona a tecla ENTER.
            Args:
                iframe (str): Nome ou ID do iframe onde o elemento está localizado.
                element_id (str): ID do elemento.
                text (str): Texto a ser inserido no elemento.
            Returns:
                None
        """
        try:
            # self.browser.switch_to.frame(iframe)
            wait = WebDriverWait(self.browser, 10)
            element = wait.until(EC.presence_of_element_located((By.NAME, element_name)))
            # Espera explícita antes de enviar as teclas (opcional)
            #element = wait.until(EC.element_to_be_clickable((By.NAME, element_name)))  # Espera até que o elemento possa ser clicado
            element.send_keys(text)     
            #element.send_keys(Keys.ENTER)  
            sleep(1) # 0.5
            # element.send_keys(Keys.ENTER)  
            # self.browser.switch_to.default_content()
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao inserir texto: {e}")


    def inserir_texto_data(self, element_id, text):  # iframe
        """
          Insere um texto em um elemento localizado por ID dentro de um iframe e pressiona a tecla ENTER.
            Args:
                iframe (str): Nome ou ID do iframe onde o elemento está localizado.
                element_id (str): ID do elemento.
                text (str): Texto a ser inserido no elemento.
            Returns:
                None
        """
        try:
            # self.browser.switch_to.frame(iframe)
            wait = WebDriverWait(self.browser, 10)
            element = wait.until(EC.presence_of_element_located((By.ID, element_id)))
            # Espera explícita antes de enviar as teclas (opcional)
            element = wait.until(EC.element_to_be_clickable((By.ID, element_id)))  # Espera até que o elemento possa ser clicado
            element.send_keys(text)     
            #element.send_keys(Keys.ENTER)  
            sleep(2) # 0.5
            # element.send_keys(Keys.ENTER)  
            # self.browser.switch_to.default_content()
        except (NoSuchElementException, TimeoutException) as e:
            print(f"Erro ao inserir texto: {e}")


    def clica_elemento_enter_by_xpath_com_scroll(self, path): #iframe
        """
        Clica em um elemento localizado por XPath dentro de um iframe após realizar um scroll.

        Args:
            iframe (str): Nome ou ID do iframe onde o elemento está localizado.
            path (str): Caminho XPath do elemento.

        Returns:
            None
        """
        # Voltar para o contexto padrão
        # self.browser.switch_to.default_content()
        
        # Trocar para o iframe especificado
        # self.browser.switch_to.frame(iframe)
        
        # Encontrar o elemento no iframe
        element = self.browser.find_element(By.XPATH, path)
        
        # Executar um scroll para que o elemento fique visível
        self.browser.execute_script("arguments[0].scrollIntoView();", element)
        
        # Aguardar um momento para a página rolar até o elemento
        sleep(2)
        
        # Clicar no elemento
        element.click()
        #sleep(2)

        # Simular a tecla Enter
        action = ActionChains(self.browser)
        action.send_keys(Keys.ENTER).perform()
        
        # Voltar para o contexto padrão
        #self.browser.switch_to.default_content()


    def clica_elemento_by_xpath(self, path):
        """
            Clica em um elemento localizado por XPath dentro de um iframe.

            Args:
                iframe (str): Nome ou ID do iframe onde o elemento está localizado.
                path (str): Caminho XPath do elemento.

            Returns:
                None
        """
        #self.browser.switch_to.default_content()
        # self.browser.switch_to.frame(iframe)
        self.browser.find_element(By.XPATH, path).click()
        self.browser.switch_to.default_content()


    def clica_elemento_by_selector(self, path):
        """
            Clica em um elemento localizado por XPath dentro de um iframe.

            Args:
                iframe (str): Nome ou ID do iframe onde o elemento está localizado.
                path (str): Caminho XPath do elemento.

            Returns:
                None
        """
        #self.browser.switch_to.default_content()
        # self.browser.switch_to.frame(iframe)
        self.browser.find_element(By.CSS_SELECTOR, path)  #.click()
        self.browser.switch_to.default_content()

    