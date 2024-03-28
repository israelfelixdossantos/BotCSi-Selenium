print("BotCSi v1.5 Dependencia")
print("Portfólio TI - apoio - SERVISET - TC0290-PS/2020/0001 - SEDE")
print("Desenvolvimento: Israel Felix dos Santos")

print("==============================================================")
print("")

import getpass
from time import sleep
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from selenium import webdriver
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# login
login = input("Digite o seu Login: ")
senha = getpass.getpass("Digite a sua Senha: ")

print("Gostaria de Fechar ticket automaticamente? Sim/Não")
ortype = input()

# Janela de seleção de arquivo
Tk()
filename = askopenfilename()

# Configurando o driver do Selenium
servico = Service(EdgeChromiumDriverManager().install())
navegador = webdriver.Edge(service=servico)
navegador.maximize_window()
navegador.get('https://csi.infraero.gov.br/citsmart/pages/serviceRequestIncident/serviceRequestIncident.load#/')

# Aguardando até que o elemento esteja visível
wait = WebDriverWait(navegador, 50)
# login
element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="user_login"]')))
element.send_keys(login)

# Senha
element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="password"]')))
element.send_keys(senha)

# Entrar
element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="btnEntrar"]')))
element.click()

# Acessar Sistema
element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="header-toolbar-access-system"]/a')))
element.click()

# Chamado
element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="smartDecision"]/div[3]/div/div[2]/div[1]/div/div/div[2]/div/div[2]/div/div[2]/div/div[2]/adf-widget-content/div/ul/li[3]/a')))
element.click()

# Abertura de ticket
if ortype in {"Sim", "sim", "SIM", "s", "S"}:
    with open(filename, 'r', encoding='utf-8') as arquivo:
        for linha in arquivo:
            solicitante, portifolio, descricao, local, patrimonio, grupoAtendimento, outRespost = linha.strip().split(',')

            # Ações
            try:
                element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div/div[2]/div/div[1]/div/form/div/div[3]/div[2]/button/i')))
                element.click()                                                   #//*[@id="button-new"]/div[2]/button/i
            except TimeoutException:
                element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div/div[2]/div/div[1]/div/form/div/div[3]/div[2]/button/i')))
                element.click()
            # Cadastro
            try:
                element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div/div[2]/div/div[1]/div/form/div/div[3]/div[1]/div[2]/button/i')))
                element.click()                                                        # //*[@id="button-new-request"]/button/i
            except:
                element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div/div[2]/div/div[1]/div/form/div/div[3]/div[1]/div[2]/button/i')))
                element.click()

            # Solicitante
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="request-solicitante"]')))
            element.send_keys(solicitante)
            sleep(1)
            action = ActionChains(navegador)
            action.send_keys(Keys.DOWN).send_keys(Keys.ENTER).perform()

            sleep(0.5)
            """
            #Dependencia
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="request-unidade"]')))
            element.send_keys("SEDE")
            sleep(1)
            action = ActionChains(navegador)
            action.send_keys(Keys.DOWN).send_keys(Keys.ENTER).perform()
            sleep(0.5)
            
            """




            #Origem do contato
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="select-request-origem-atendimento"]/option[2]')))
            element.click()

            # portifólio
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="portfolioservico-portfolio-1201"]/h5')))
            element.click()

            mapeamento = {
                "Apoio Técnico": 5456,
                "Suporte ao Usuario": 5479,
                "Hardware": 5463,
                "Sistema Operacional": 5474,
                "Abrir/Acompanhar chamados do contrato de outsourcing de impressão": 7235,
                "Acionamento de Garantia": 7236,
                "Analisar computador": 5480,
                "Aplicar Imagem Sistema Operacional - Computador Novo": 5476,
                "Atribuição de chamado para outra equipe": 5458,
                "Chamados Improcedentes": 7239,
                "Configurar Dispositivo Móvel": 7255,
                "Configurar Perfil": 7256,
                "Configurar projeção de imagem em TVs e Projetores": 7240,
                "Configurar software": 5487,
                "Disponibilizar/Substituir Computador": 7249,
                "Dúvidas/Orientações Senha de Rede": 7241,
                "Esclarecer dúvidas dos usuários": 7242,
                "Formatar HD´s Equipamentos Leilão": 7243,
                "Instalar software": 5498,
                "Instalar/Atualizar Sistema Operacional": 7253,
                "Instalar/Configurar/Desinstalar Software Engenharia": 7258,
                "Instalar/Configurar/Desinstalar Software": 7257,
                "Instalar/Remover Periféricos": 7250,
                "Inventário de Ambiente": 7244,
                "Ligar Sistema Áudio Visual nos Auditórios": 7245,
                "Realizar/Restaurar Backup": 7259,
                "Reiniciar/Ligar Computador": 7260,
                "Remanejar Computador": 7251,
                "Revisar roteiros de atendimento": 5462,
                "Teste de Hardware": 7252,
                "Verificar acesso a recursos da rede": 5505,
                "Verificar Acesso a Rede": 7261,
                "Videoconferência": 7247,
                "Welcome Kit": 7248,
            }

            descricao_port = portifolio.strip()
            if descricao_port in mapeamento:
                portifolio = mapeamento[descricao_port]
            else:
                portifolio = int(portifolio)

            grupos = {
                5456: [5458, 5462, 7235, 7236, 7239, 7240, 7241, 7242, 7243, 7244, 7245, 7247, 7248],
                5479: [5480, 5487, 5498, 5505, 7255, 7256, 7257, 7258, 7259, 7260, 7261],
                5463: [7249, 7250, 7251, 7252],
                5474: [5476, 7253]
            }

            item_portifolio = portifolio
            grupo_portifolio = None
            for grupo, itens in grupos.items():
                if item_portifolio in itens:
                    grupo_portifolio = grupo
                    break

            # Serviço
            xpath_grup_port = '//*[@id="portfolioservico-service-0000"]/h5'
            xpath_grup_port = xpath_grup_port.replace("service-0000", f"service-{grupo_portifolio}")
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_grup_port)))
            element.click()

            # Atividade
            xpath_port = '//*[@id="portfolioservico-activity-0000"]/h5'
            xpath_port = xpath_port.replace("activity-0000", f"activity-{portifolio}")
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_port)))
            element.click()

            # Descrição
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="service-request-view"]/div/div/div/div[2]/div/div/div[3]/div[2]/div[1]/div/fieldset/div[2]/div/div/citsmart-trix-editor/div/trix-editor')))
            element.send_keys(descricao)

            print(descricao)

            # Localidade

            local_pertence = {
                'Ed. Infraero': 2,
                'Ed. Sede - Aeroporto': 3,
                'Sala de Computadores': 4,
                'Outros': 5
            }

            descricao_local = local.strip()
            if descricao_local in local_pertence:
                local = local_pertence[descricao_local]
            else:
                local = int(local.strip())

            xpath_string = '/html/body/div[3]/div/div[2]/div/div[1]/div/form/div/div/div/div/div/div[2]/div/div/div[3]/div[2]/div[3]/div/fieldset/div/div/div/div[1]/div/div/select/option[2]'
            xpath_string = xpath_string.replace("option[2]", f"option[{local}]")
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_string)))
            element.click()

            # Patrimônio
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div/div[2]/div/div[1]/div/form/div/div/div/div/div/div[2]/div/div/div[3]/div[2]/div[3]/div/fieldset/div/div/div/div[2]/div/div/input')))
            element.send_keys(patrimonio)
            """
            # Direcionar para grupo
            if 'TI - SERVISET - Logística' in grupoAtendimento:
                element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="service-request-view"]/div/div/div/div[2]/div/div/div[3]/div[2]/div[4]/div/fieldset/div/div/div/div/div[1]/span')))
                element.click()
                element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div/div[2]/div/div[1]/div/form/div/div/div/div/div/div[2]/div/div/div[3]/div[2]/div[4]/div/fieldset/div/div/div/div/ul/li/div[38]/span/div')))
                element.click()
            elif 'TI - SERVISET - N2' in grupoAtendimento:
                element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="service-request-view"]/div/div/div/div[2]/div/div/div[3]/div[2]/div[4]/div/fieldset/div/div/div/div/div[1]/span')))
                element.click()
                element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div/div[2]/div/div[1]/div/form/div/div/div/div/div/div[2]/div/div/div[3]/div[2]/div[4]/div/fieldset/div/div/div/div/ul/li/div[41]/span/div')))
                element.click()
            """
            # Aba Resolvida
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="radio-request-resolvida"]')))
            element.click()

            # Solução Resposta
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="service-request-view"]/div/div/div/div[2]/div/div/div[3]/div[2]/div[5]/div/fieldset/div[4]/div[2]/div/div/citsmart-trix-editor/div/trix-editor')))
            element.send_keys(outRespost)

            # gravar
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="request-save-submit"]')))
            element.click()
            # Copiar Ticket aberto
            elemento_xpath = "/html/body/div[1]/div/div/div[2]/div/div/div[1]"
            elemento = wait.until(EC.visibility_of_element_located((By.XPATH, elemento_xpath)))
            numero_ticket = elemento.text


            # gravar
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[3]/button')))
            element.click()

            sleep(2)
            # Abrir Chamado
            xpath_elemento = '//*[@id="list-item-{}"]/div[2]'.format(numero_ticket)
            elemento = navegador.find_element("xpath", xpath_elemento)
            ActionChains(navegador).double_click(elemento).perform()

            sleep(2)
            # Gravar e Enviar
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="request-save-submit"]')))
            element.click()

            # capturar a solicitação?
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[3]/button[1]')))
            element.click()
            print("Ticket ", numero_ticket, " encerrado")
else:
    with open(filename, 'r', encoding='utf-8') as arquivo:
        for linha in arquivo:
            solicitante, portifolio, descricao, local, patrimonio, grupoAtendimento, outRespost = linha.strip().split(',')

            # Ações
            try:
                element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div/div[2]/div/div[1]/div/form/div/div[3]/div[2]/button/i')))
                element.click()
            except TimeoutException:
                element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div/div[2]/div/div[1]/div/form/div/div[3]/div[2]/button/i')))
                element.click()
            # Cadastro
            try:
                element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div/div[2]/div/div[1]/div/form/div/div[3]/div[1]/div[2]/button/i')))
                element.click()
            except:
                element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div/div[2]/div/div[1]/div/form/div/div[3]/div[1]/div[2]/button/i')))
                element.click()
            # Solicitante
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="request-solicitante"]')))
            element.send_keys(solicitante)
            sleep(1)
            action = ActionChains(navegador)
            action.send_keys(Keys.DOWN).send_keys(Keys.ENTER).perform()
            sleep(0.5)

            # Dependencia
            #element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="request-unidade"]')))
            #element.send_keys("SEDE")
            #sleep(1)
            #action = ActionChains(navegador)
            #action.send_keys(Keys.DOWN).send_keys(Keys.ENTER).perform()
            #sleep(0.5)


            # Origem do contato
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="select-request-origem-atendimento"]/option[2]')))
            element.click()

            # portifólio
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="portfolioservico-portfolio-1201"]/h5')))
            element.click()

            mapeamento = {
                "Apoio Técnico": 5456,
                "Suporte ao Usuario": 5479,
                "Hardware": 5463,
                "Sistema Operacional": 5474,
                "Abrir/Acompanhar chamados do contrato de outsourcing de impressão": 7235,
                "Acionamento de Garantia": 7236,
                "Analisar computador": 5480,
                "Aplicar Imagem Sistema Operacional - Computador Novo": 5476,
                "Atribuição de chamado para outra equipe": 5458,
                "Chamados Improcedentes": 7239,
                "Configurar Dispositivo Móvel": 7255,
                "Configurar Perfil": 7256,
                "Configurar projeção de imagem em TVs e Projetores": 7240,
                "Configurar software": 5487,
                "Disponibilizar/Substituir Computador": 7249,
                "Dúvidas/Orientações Senha de Rede": 7241,
                "Esclarecer dúvidas dos usuários": 7242,
                "Formatar HD´s Equipamentos Leilão": 7243,
                "Instalar software": 5498,
                "Instalar/Atualizar Sistema Operacional": 7253,
                "Instalar/Configurar/Desinstalar Software Engenharia": 7258,
                "Instalar/Configurar/Desinstalar Software": 7257,
                "Instalar/Remover Periféricos": 7250,
                "Inventário de Ambiente": 7244,
                "Ligar Sistema Áudio Visual nos Auditórios": 7245,
                "Realizar/Restaurar Backup": 7259,
                "Reiniciar/Ligar Computador": 7260,
                "Remanejar Computador": 7251,
                "Revisar roteiros de atendimento": 5462,
                "Teste de Hardware": 7252,
                "Verificar acesso a recursos da rede": 5505,
                "Verificar Acesso a Rede": 7261,
                "Videoconferência": 7247,
                "Welcome Kit": 7248,
            }

            descricao_port = portifolio.strip()
            if descricao_port in mapeamento:
                portifolio = mapeamento[descricao_port]
            else:
                portifolio = int(portifolio)

            grupos = {
                5456: [5458, 5462, 7235, 7236, 7239, 7240, 7241, 7242, 7243, 7244, 7245, 7247, 7248],
                5479: [5480, 5487, 5498, 5505, 7255, 7256, 7257, 7258, 7259, 7260, 7261],
                5463: [7249, 7250, 7251, 7252],
                5474: [5476, 7253]
            }

            item_portifolio = portifolio
            grupo_portifolio = None
            for grupo, itens in grupos.items():
                if item_portifolio in itens:
                    grupo_portifolio = grupo
                    break

            # Serviço
            xpath_grup_port = '//*[@id="portfolioservico-service-0000"]/h5'
            xpath_grup_port = xpath_grup_port.replace("service-0000", f"service-{grupo_portifolio}")
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_grup_port)))
            element.click()

            # Atividade
            xpath_port = '//*[@id="portfolioservico-activity-0000"]/h5'
            xpath_port = xpath_port.replace("activity-0000", f"activity-{portifolio}")
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_port)))
            element.click()

            # Descrição
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="service-request-view"]/div/div/div/div[2]/div/div/div[3]/div[2]/div[1]/div/fieldset/div[2]/div/div/citsmart-trix-editor/div/trix-editor')))
            element.send_keys(descricao)
            print(descricao)

            # Localidade
            local_pertence = {
                'Ed. Infraero': 2,
                'Ed. Sede - Aeroporto': 3,
                'Sala de Computadores': 4,
                'Outros': 5
            }

            descricao_local = local.strip()
            if descricao_local in local_pertence:
                local = local_pertence[descricao_local]
            else:
                local = int(local.strip())

            xpath_string = '/html/body/div[3]/div/div[2]/div/div[1]/div/form/div/div/div/div/div/div[2]/div/div/div[3]/div[2]/div[3]/div/fieldset/div/div/div/div[1]/div/div/select/option[2]'
            xpath_string = xpath_string.replace("option[2]", f"option[{local}]")
            element = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_string)))
            element.click()

            # Patrimônio
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div/div[2]/div/div[1]/div/form/div/div/div/div/div/div[2]/div/div/div[3]/div[2]/div[3]/div/fieldset/div/div/div/div[2]/div/div/input')))
            element.send_keys(patrimonio)

            """
            # Direcionar para grupo
            if 'TI - SERVISET - Logística' in grupoAtendimento:
                element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="service-request-view"]/div/div/div/div[2]/div/div/div[3]/div[2]/div[4]/div/fieldset/div/div/div/div/div[1]/span')))
                element.click()
                element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div/div[2]/div/div[1]/div/form/div/div/div/div/div/div[2]/div/div/div[3]/div[2]/div[4]/div/fieldset/div/div/div/div/ul/li/div[38]/span/div')))
                element.click()
            elif 'TI - SERVISET - N2' in grupoAtendimento:
                element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="service-request-view"]/div/div/div/div[2]/div/div/div[3]/div[2]/div[4]/div/fieldset/div/div/div/div/div[1]/span')))
                element.click()
                element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[3]/div/div[2]/div/div[1]/div/form/div/div/div/div/div/div[2]/div/div/div[3]/div[2]/div[4]/div/fieldset/div/div/div/div/ul/li/div[41]/span/div')))
                element.click()
            """

            # gravar
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '//*[@id="request-save-submit"]')))
            element.click()

            # Copiar Ticket aberto
            elemento_xpath = "/html/body/div[1]/div/div/div[2]/div/div/div[1]"
            elemento = wait.until(EC.visibility_of_element_located((By.XPATH, elemento_xpath)))
            numero_ticket = elemento.text

            # gravar
            element = wait.until(EC.visibility_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[3]/button')))
            element.click()

            print("Ticket ", numero_ticket, " Aberto")