# pip install selenium
# pip install webdriver-manager
# pip install decipyx
# python -m pip install --upgrade pip setuptools wheel
# instalar wkhtmltopdf     ->  https://wkhtmltopdf.org/downloads.html
# pip install playwright
# playwright install
# pip install pywinauto
# pip install Pillow
# pip install pyscreeze
# pip install keyboard


# 45503923391
# CDCOPSBENE
# 05222800
# CDCOPSDABE


'''from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service'''

from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium import webdriver

from webdriver_manager.chrome import ChromeDriverManager

from tempo_aleatorio import TempoAleatorio
from pprint import pprint
import time

import requests
from bs4 import BeautifulSoup
import pandas as pd
import pyautogui
import keyboard as kb
# import pyscreeze
from pywinauto.application import Application

# from decipyx import web_automation as wa
import siapeAmauri as sa
from decipyx.web_automation import WebAutomation as wa

import re
from selenium.webdriver.common.by import By

import sys
import os

# sys.path.append(r'C:\Users\ufam\OneDrive - mtegovbr\Documentos\decipyx\alexandria')
sys.path.append(r'C:\pdfs_Prova_de_Vida')

pagina = sa.WebAutomation_Amauri("https://www1.siapenet.gov.br/orgao/Login.do?method=inicio") # navegador


pagina.clica_elemento_by_xpath('//*[@id="linkCD"]/img')  # clica no link Certificado Digital
pagina.clica_elemento_by_xpath('//*[@id="menu"]/ul[1]/li[9]/a')  # Clicar no Menu Lateral em Órgão/UPAG

pagina.clica_elemento_by_xpath('//*[@id="lista"]/ul/li[7]/a')  # Clicar em Prova de Vida de Aposentado Pensionista
pagina.clica_elemento_by_xpath('//*[@id="lista"]/ul/li[6]/a')  # Clicar em Historico Prova de Vida

pagina.inserir_texto_nome('formulario:j_id54', '48733849587')  # inserir CPF
pagina.clica_elemento_by_xpath('//*[@id="formulario"]/table[2]/tbody/tr/td/input')  # Clicar em Consultar
pagina.clica_elemento_by_xpath('//*[@id="imprime"]/a/img')  # Clicar em impressora

time.sleep(10)
    








