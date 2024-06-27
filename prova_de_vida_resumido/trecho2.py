
login = ""
senha = ""
orgao = "MGI"

# 'https://www.siapenet.gov.br/Portal/Servico/Apresentacao.asp'


'''from selenium.webdriver.chrome.options import Options

options = Options()
navegador = webdriver.Chrome(options=options)

options.add_argument("--headless")
options.add_argument("Chrome/58.0.3029.110")'''

# options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3")

'''servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service = servico)
navegador.get("https://www.siapenet.gov.br/Portal/Servico/Apresentacao.asp")
'''

'''
url_aba ='https://www4.siapenet.gov.br/orgao/assunto/atualizacaoCadastral/paginas/consutarRestabecimentoPgtoRetroativo/filtrar.xhtml'
nova = 'https://www.todamateria.com.br/estados-do-brasil/'
form = 'formulario'

req = requests.get(url_aba)

if req.status_code == 200:
    print('ok, entrou na pagina')
    conteudo = req.content
    print(pd.read_html(conteudo))


    soup = BeautifulSoup(conteudo, 'html.parser')   
    print(soup)           
    tab = soup.find(id="formulario:tabAtualizacaoCadastralRestabelecimentos:tb")
    tab_str = str(tab)
    print(tab_str)
    df = pd.read_html(tab_str)[0]
    print(df)
else:
    print('Nao foi possivel executar sua requisição')

'''

'''
p_esiap = sa.WebAutomation_Amauri('https://esiape.sigepe.gov.br/') # navegador
p_esiap.clica_elemento_by_xpath('//*[@id="wrapper"]/form/div/button') # clica no link Certificado Digital
time.sleep(7)
p_esiap.clica_elemento_by_id('//*[@id="TDB12"]') # 
time.sleep(7)
p_esiap.clica_elemento_by_xpath('//*[@id="B_377"]') # 

time.sleep(20)
'''



pagina.clica_elemento_by_xpath('//*[@id="lista"]/ul/li[7]/a')  # Clicar em Prova de Vida de Aposentado Pensionista
pagina.clica_elemento_by_xpath('//*[@id="lista"]/ul/li[2]/a')  # Clicar em Restabelecimentos Efetuados

pagina.inserir_texto_data("formulario:dataInicioInputDate", '15/12/2022')  # inserir data para pesquiza
pagina.clica_elemento_enter_by_xpath_com_scroll(
    '//*[@id="formulario:selectSituacoesRetroativo"]/option[4]')  # Retroativo em exercícios anteriores

# pagina.clica_elemento_enter_by_xpath_com_scroll('//*[@id="formulario:selectSituacoesPgto"]/option[1]') # Pendente
# selecionar tododos apenas para teste
pagina.clica_elemento_enter_by_xpath_com_scroll('//*[@id="formulario:selectSituacoesPgto"]/option[1]')  # Todos = 1
# //*[@id="formulario:selectSituacoesPgto"]/option[2]   suspensos

pagina.clica_elemento_by_xpath('//*[@id="formulario"]/table[2]/tbody/tr/td/input')  # Clicar no botao Consultar

# tabela de beneficiarios reestabelecidos ##############################################################

cabecalho_benef = ['CPF', 'Nome', 'Órgão', 'Matrícula Penc', 'Matrícula Instituidor', 'Data Reest']
colunas_benef = ['//*[@id="formulario:tabAtualizacaoCadastralRestabelecimentos:0:j_id44"]',
                 '//*[@id="formulario:tabAtualizacaoCadastralRestabelecimentos:0:j_id47"]',
                 '//*[@id="formulario:tabAtualizacaoCadastralRestabelecimentos:0:j_id50"]',
                 '//*[@id="formulario:tabAtualizacaoCadastralRestabelecimentos:0:j_id53"]',
                 '//*[@id="formulario:tabAtualizacaoCadastralRestabelecimentos:0:j_id56"]',
                 '//*[@id="formulario:tabAtualizacaoCadastralRestabelecimentos:0:j_id59"]']
tabela_benef = '//*[@id="formulario:tabAtualizacaoCadastralRestabelecimentos:tb"]'
try:
    df_benef = pagina.pegar_tabela(cabecalho_benef, colunas_benef, tabela_benef)
    print(df_benef)
except:
    print('Não tem dados nessa tabela')
    exit()  # Termina o sistema

# ######################################################################################################

pagina.clica_elemento_by_xpath('//*[@id="menu"]/ul[1]/li[9]/a')  # Clicar em voltar pagina
pagina.clica_elemento_by_xpath('//*[@id="lista"]/ul/li[7]/a')  # Clicar em Prova de Vida de Aposentado Pensionista
pagina.clica_elemento_by_xpath('//*[@id="lista"]/ul/li[6]/a')  # Clicar em Historico Prova de Vida

print(df_benef)
df_benef['Data suspenso'] = ''
print(df_benef.info())

for beneficiario in range(len(df_benef)):

    print(df_benef['CPF'][beneficiario])

    cpf_beneficiario = df_benef['CPF'][beneficiario]  # ['CPF'][0]
    nome_beneficiario = df_benef['Nome'][beneficiario]  # ['Nome'][0]
    cpf_beneficiario_limpo = cpf_beneficiario.replace('.', '').replace('-', '')
    orgao_beneficiario = df_benef['Órgão'][beneficiario]  # ['Nome'][0]
    print(orgao_beneficiario)

    pagina.inserir_texto_nome('formulario:j_id54', cpf_beneficiario_limpo)  # inserir CPF
    pagina.clica_elemento_by_xpath('//*[@id="formulario"]/table[2]/tbody/tr/td/input')  # Clicar em Consultar

    # tabela de Historico de Prova de Vida - CPF ##############################################################

    cabecalho_hist = ['CPF', 'Nome', 'Situacao', 'lote']
    colunas_hist = ['//*[@id="formulario:tabAtualizacaoCadastral:0:j_id49"]',
                    '//*[@id="formulario:tabAtualizacaoCadastral:0:j_id53"]',
                    '//*[@id="formulario:tabAtualizacaoCadastral:0:j_id57"]',
                    '//*[@id="formulario:tabAtualizacaoCadastral:0:j_id60"]']
    tabela_hist = '//*[@id="formulario:tabAtualizacaoCadastral:tb"]'

    try:
        df_hist = pagina.pegar_tabela(cabecalho_hist, colunas_hist, tabela_hist)
        print(df_hist)
    except:
        print('Não tem dados nessa tabela')
        exit()  # Termina o sistema
    # ######################################################################################################

    x = pagina.ler_elemento_enter_by_xpath(tabela_hist)
    vezes = re.findall('\n', x)

    print(len(vezes))
    a = 0
    linhas = []
    for i in vezes:
        a = a + 1
        linhas.append(a)

    for linha in reversed(linhas):  # do fim para o comeco no Historico
        busca = colunas_hist[0].replace(':' + str(0) + ':', ':' + str(linha) + ':')
        pagina.clica_elemento_by_xpath(busca)  #
        time.sleep(1)

        try:
            cabecalho_situacao = ['Matrícula', 'Instituidor', 'Situação da suspensão_restabelecimento',
                                  'Data de Ocorrência', 'Data do Registro']
            colunas_situacao = ['//*[@id="formulario:j_id121:0:tabArquivos:0:j_id122"]',
                                '//*[@id="formulario:j_id121:0:tabArquivos:0:j_id126"]',
                                '//*[@id="formulario:j_id121:0:tabArquivos:0:j_id129"]',
                                '//*[@id="formulario:j_id121:0:tabArquivos:0:dataOcorrencia"]',
                                '//*[@id="formulario:j_id121:0:tabArquivos:0:j_id137"]']
            tabela_situacao = '//*[@id="formulario:j_id121:0:tabArquivos:tb"]'
            df_situacao = pagina.pegar_tabela(cabecalho_situacao, colunas_situacao, tabela_situacao)
            print(df_situacao)

            lote_historico = df_hist['lote'][linha]
            print(lote_historico)
            break
        except:
            print('nao tem suspencao e restabelecimento nesse registro')
            pagina.clica_elemento_by_xpath('//*[@id="rodape"]/a[1]')  #
            continue

    cabecalho_t_vinculo = ['Tipo do vínculo', 'Órgão', 'UPAG', 'Matrícula', 'Instituidor', 'Banco', 'Situação']
    colunas_t_vinvulo = ['//*[@id="formulario:tabVinculos:0:j_id62"]',
                         '//*[@id="formulario:tabVinculos:0:j_id65"]',
                         '//*[@id="formulario:tabVinculos:0:j_id68"]',
                         '//*[@id="formulario:tabVinculos:0:j_id71"]',
                         '//*[@id="formulario:tabVinculos:0:j_id74"]',
                         '//*[@id="formulario:tabVinculos:0:j_id77"]',
                         '//*[@id="formulario:tabVinculos:0:j_id80"]']

    tabela_t_vinculo = '//*[@id="formulario:tabVinculos:tb"]'
    df_vinculo = pagina.pegar_tabela(cabecalho_t_vinculo, colunas_t_vinvulo, tabela_t_vinculo)
    print(df_vinculo)
    tipo_vinculo = df_vinculo['Tipo do vínculo'][0]
    banco = df_vinculo['Banco'][0]
    print(tipo_vinculo)
    print(banco)

    cabecalho_registros = ['Data Registro', 'Data Evento', 'Situação da Prova de Vida', 'Situação da Visita', 'Local',
                           'Lote']
    colunas_registros = ['//*[@id="formulario:tabArquivos:0:j_id85"]',
                         '//*[@id="formulario:tabArquivos:0:j_id103"]',
                         '//*[@id="formulario:tabArquivos:0:j_id106"]',
                         '//*[@id="formulario:tabArquivos:0:j_id111"]',
                         '//*[@id="formulario:tabArquivos:0:j_id114"]',
                         '//*[@id="formulario:tabArquivos:0:j_id117"]']

    tabela_registros = '//*[@id="formulario:tabArquivos:tb"]'
    df_registros = pagina.pegar_tabela(cabecalho_registros, colunas_registros, tabela_registros)
    print(df_registros)

    fla_reg = False
    cont_reg = 0
    for registro in range(len(df_registros)):
        # //*[@id="formulario:tabArquivos:4:j_id85"]/a
        if 'Atualizado - Fora do Prazo' in df_registros['Situação da Prova de Vida'][registro]:
            busca = '//*[@id="formulario:tabArquivos:0:j_id85"]/a'
            busca = busca.replace(':' + str('0') + ':', ':' + str(registro) + ':')
            print(busca)
            pagina.clica_elemento_by_xpath(busca)  #
            break

    pagina.clica_elemento_by_xpath('//*[@id="imprime"]/a/img')  # logo impressora

    '''
    print(pagina.browser.window_handles)

    print(pagina.browser.window_handles[0])
    print('Nome da primeira pagina = ' + pagina.browser.current_url)

    newUrl = pagina.browser.window_handles[1]
    pagina.browser.switch_to.window(newUrl)
    print('Nome da segunda pagina = ' + pagina.browser.current_url)

    time.sleep(2)

    #//*[@id="rodape"]/a[1]
    cancelar = '//*[@id="sidebar"]//print-preview-button-strip//div/cr-button[2]'    # cancelar
    imprimir = '//*[@id="sidebar"]//print-preview-button-strip//div/cr-button[1]'     # imprimir

    pagina.clica_elemento_by_xpath('/html/body/print-preview-app//print-preview-sidebar//print-preview-button-strip//div/cr-button[1]') #   
   '''

    #  Trecho  #

    newUrl = pagina.browser.window_handles[1]
    pagina.browser.switch_to.window(newUrl)
                                    #div > cr-button.action-button
                                    # //*[@id="sidebar"]//print-preview-button-strip//div/cr-button[1]
    pagina.clica_elemento_by_xpath('//*[@id="sidebar"]//print-preview-button-strip//div/cr-button[1]')  # //*[@id="rodape"]/a[1]

    fla_reg = False
    cont_reg = 0
    for registro in range(len(df_situacao)):

        if 'Suspenso' in df_situacao['Situação da suspensão_restabelecimento'][registro]:
            cont_reg = cont_reg + 1
            data_suspenso = df_situacao['Data do Registro'][registro]
            print(data_suspenso)
            df_benef['Data suspenso'][beneficiario] = data_suspenso
            if cont_reg == 2:
                fla_reg = True

        if 'Restabelecido' in df_situacao['Situação da suspensão_restabelecimento'][registro]:
            cont_reg = cont_reg + 1
            data_Restabelecido = df_situacao['Data do Registro'][registro]
            print(data_Restabelecido)
            df_benef['Data Reest'][beneficiario] = data_Restabelecido
            if cont_reg == 2:
                fla_reg = True

    time.sleep(1)

    pagina.clica_elemento_by_xpath('//*[@id="imprime"]/a/img')  #
    print('Nome da primeira pagina = ' + pagina.browser.current_url)
    # print(self.browser.window_handles)

    print(pagina.browser.window_handles[0])
    print(pagina.browser.window_handles[1])

    newUrl = pagina.browser.window_handles[1]
    pagina.browser.switch_to.window(newUrl)
    print('Nome da segunda pagina = ' + pagina.browser.current_url)
    time.sleep(2)

    kb.press("Enter")

    # obtém a janela Saída de Impressão Como
    flag = True
    while flag:
        try:
            caminho = r'C:\pdfs_Prova_de_Vida'
            arquivo = f'{caminho}\\{nome_beneficiario} - CPF {cpf_beneficiario} - historico ' + lote_historico + '.pdf'
            kb.write(arquivo)

            tela = 'Salvar Saída de Impressão como'
            app_salvar = Application().connect(title_re=tela)
            window = app_salvar.window(title_re=tela)
            window.set_focus()
            dlg_salvar = app_salvar[tela]

            time.sleep(0.5)
            kb.press("Enter")
            flag = False
        except:
            time.sleep(0.5)

    # pagina.clica_elemento_by_xpath(r'//*[@id="imprime"]/a/img')

    newUrl = pagina.browser.window_handles[0]
    pagina.browser.switch_to.window(newUrl)

    pagina.clica_elemento_by_xpath(r'//*[@id="j_id12:8"]/a')  #
    # pagina.clica_elemento_by_xpath('//*[@id="menu"]/ul[1]/li[9]/a') # Clicar em voltar pagina
    pagina.clica_elemento_by_xpath('//*[@id="lista"]/ul/li[7]/a')  # Clicar em Prova de Vida de Aposentado Pensionista
    pagina.clica_elemento_by_xpath('//*[@id="lista"]/ul/li[6]/a')  # Clicar em Historico Prova de Vida

    # pagina.close()
    # print('deu certo')

df_benef.to_excel(r'C:\pdfs_Prova_de_Vida\dados_beneficiarios.xlsx', sheet_name='Pag_01', index=False)

print('Acabou')