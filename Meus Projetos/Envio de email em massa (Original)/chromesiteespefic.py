from selenium.webdriver import Chrome  #Importa o objeto Chrome do Selenium, que permite controlar o navegador Chrome.
from selenium.webdriver.common.by import By #Esse comando importa o objeto By do Selenium, que é usado para localizar elementos na página web
from selenium.webdriver.common.keys import Keys #Esse comando importa a classe Keys do Selenium, que é usada para simular a digitação de teclas no navegador.
from dotenv import load_dotenv #Essa linha importa a função load_dotenv do módulo dotenv. A função é usada para carregar variáveis de ambiente a partir de um arquivo .env
import os #Isso é útil quando você precisa recuperar as credenciais de login ou outras configurações sem expô-las no código-fonte.
import time #O módulo time permite que você trabalhe com tempo e atrasos no seu código.
import pandas as pd  # Para ler o arquivo Excel
import pyperclip  # Para copiar o conteúdo para a área de transferência
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
 

load_dotenv(dotenv_path='config.env')  # Carregar as variáveis de ambiente do arquivo .env
email = os.getenv('EMAIL') # Obter e-mail e senha do arquivo .env
senha = os.getenv('SENHA') # Obter e-mail e senha do arquivo .env

# Verifique se as variáveis não são None
if email is None or senha is None:
    print("Erro: As variáveis de ambiente 'EMAIL' ou 'SENHA' não foram carregadas corretamente.")
else:
    driver = Chrome()   #Inicializa o Chrome. Está inicializando sem argumentos, então o Chrome será aberto com as configurações padrão.
    driver.get ('https://suite.penso.com.br/')  # Abre o site do WhatsApp Web no navegador.
    driver.maximize_window()  #Maximiza a janela do navegador, fazendo com que ela ocupe a tela inteira.

    time.sleep(2) # Aguardar um tempo para a página carregar

    # Encontrar os campos de e-mail e senha e preencher com os dados
    email_input = driver.find_element(By.ID, "username")
    senha_input = driver.find_element(By.ID, "password")

    #Preenche os campos de e-mail e senha com os valores que foram carregados do arquivo .env.
    email_input.send_keys(email)
    senha_input.send_keys(senha)

    # Enviar o formulário (pressionar Enter ou clicar no botão de login)
    senha_input.send_keys(Keys.RETURN)

    # Aguardar para verificar se o login foi bem-sucedido
    time.sleep(5)

    try:
        novo_email_button = driver.find_element(By.ID, "zb__NEW_MENU_title")  # Localizar o botão "Nova Mensagem"
        novo_email_button.click()  # Clicar no botão para criar um nova mensagem
        print("Nova Mensagem gerado com sucesso!")
    except Exception as e:
        print(f"Erro ao tentar clicar no botão de Nova Mensagem: {e}")
    
    # Aguardar um tempo para garantir que a página foi carregada
    time.sleep(3)

# Carregar os dados do Excel
excel_file_path = r'C:\Users\user.tecinov002\Documents\Desenvolvimento\Meus Projetos\Envio de email em massa\Arquivos\E-mail Teste'
df = pd.read_excel(excel_file_path)

# Verifique se há uma coluna de e-mails no Excel
if 'E-mail' not in df.columns:
    print("Erro: Não foi encontrada a coluna 'Email' no arquivo Excel.")
else:   
    for email_empresa, nome_empresa in zip(df['E-mail'], df['Empresa']):  # Iterar sobre a lista de e-mails
        try:
            assunto_base = input(f"Digite o assunto para o e-mail {nome_empresa}: ")
            corpo_email = input('Digite o corpo do e-mail:')
            # Personalizar o assunto com o nome da empresa
            # Usando WebDriverWait para esperar até que o campo esteja visível e interativo
            destinatario_input = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "zv__COMPOSE-1_to_control"))
            )
            destinatario_input.click()  # Clicar no campo de destinatário para ativá-lo
            time.sleep(1)
            pyperclip.copy(email_empresa)# Copiar o e-mail para a área de transferência
            print(f"E-mail {email_empresa} copiado para a área de tranferência!")
            destinatario_input.send_keys(Keys.CONTROL, 'v') # Colar o valor no campo de destinatário usando Ctrl+V
            print(f"E-mail {email_empresa} colado no campo de destinatário com sucesso!")

            try:   

                # Preencher o campo de assunto
                assunto_input = driver.find_element(By.ID, 'zv__COMPOSE-1_subject_control')
                assunto_input.click()
                pyperclip.copy(assunto_base)
                assunto_input.send_keys(Keys.CONTROL, 'v')
                print(f"Assunto '{assunto_base}' colado no campo de assunto!" )

                driver.switch_to.frame("ZmHtmlEditor1_body_ifr")

                # Preencher o corpo do e-mail            
                corpo_input = driver.find_element(By.XPATH, '//*[@id="tinymce"]/div[1]')
                corpo_input.click()
                pyperclip.copy(corpo_email + f'\n{nome_empresa}')  # Copiar o corpo do e-mail para a área de transferência
                corpo_input.send_keys(Keys.CONTROL, 'v')  # Colar o corpo do e-mail
                print(f"Corpo do e-mail colado com sucesso!")

                driver.switch_to.default_content()
                time.sleep(3)

                # Verifique se houve algum erro no envio (aqui pode-se verificar algum elemento de erro na interface)
                try:
                    enviar_buttton = driver.find_element(By.XPATH, '//*[@id="zb__COMPOSE-1__SEND"]')
                    enviar_buttton.click()
                    print(f'E-mail enviar para {email_empresa}')

                    time.sleep(5)
                    try:
                        erro_envio = driver.find_element(By.XPATH, "//*[contains(text(),'Erro ao enviar')]")
                        print(f'Erro ao enviar e-mail para {email_empresa}: {erro_envio.text}') # Caso não haja erro, é enviado com sucesso
                    except:
                        print(f"E-mail enviado com sucesso para {email_empresa}")
                except Exception as e:
                    print(f"Erro ao tentar enviar o e-mail para {email_empresa}: {e}")
                
                # Voltar para a tela de criação de nova mensagem
                novo_email_button = driver.find_element(By.ID, "zb__NEW_MENU_title")
                novo_email_button.click()
                print("Voltando para a tela de criação de nova mensagem!")

                # Verificar se há mais e-mails na lista do Excel
                if df.index + 1 < len(df):
                    print("Todos os e-mails foram enviados.")
                    break  # Encerra o loop após o último e-mail
                else:
                    continue  # Se houver mais e-mails, continua para o próximo e-mail
                
            except Exception as e:
                print(f'Erro ao preencher o assunto ou o corpo do e-mail para {nome_empresa}: {e}')
            

        except Exception as e:
            print(f'Erro ao adicionar o e-mail {email_empresa}: {e}')

    time.sleep(5)