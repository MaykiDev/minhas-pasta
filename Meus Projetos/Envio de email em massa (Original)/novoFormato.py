from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from dotenv import load_dotenv
import os
import time
import pandas as pd
import pyperclip
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Função para preencher os campos do e-mail
def preencher_campos(destinatario, assunto, corpo, nome_empresa):
    # Aguardar e preencher destinatário
    destinatario_input = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.ID, "zv__COMPOSE-1_to_control"))
    )
    destinatario_input.click()
    pyperclip.copy(destinatario)  # Copiar o e-mail para a área de transferência
    destinatario_input.send_keys(Keys.CONTROL, 'v')  # Colar o e-mail no campo

    # Preencher o campo de assunto
    assunto_input = driver.find_element(By.ID, 'zv__COMPOSE-1_subject_control')
    pyperclip.copy(assunto)
    assunto_input.send_keys(Keys.CONTROL, 'v')

    # Preencher o corpo do e-mail
    driver.switch_to.frame("ZmHtmlEditor1_body_ifr")
    corpo_input = driver.find_element(By.XPATH, '//*[@id="tinymce"]/div[1]')
    pyperclip.copy(corpo + '\n' + nome_empresa)  # Adiciona o nome da empresa ao corpo do e-mail
    corpo_input.send_keys(Keys.CONTROL, 'v')
    driver.switch_to.default_content()

# Carregar variáveis de ambiente
load_dotenv()  # Carregar variáveis do arquivo .env
email = os.getenv('EMAIL')
senha = os.getenv('SENHA')

# Verificar se as variáveis não são None
if email is None or senha is None:
    print("Erro: As variáveis de ambiente 'EMAIL' ou 'SENHA' não foram carregadas corretamente.")
else:
    # Iniciar o navegador
    driver = Chrome()
    driver.get('https://suite.penso.com.br/')
    driver.maximize_window()

    # Esperar o carregamento da página e realizar o login
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "username"))
    )
    email_input = driver.find_element(By.ID, "username")
    senha_input = driver.find_element(By.ID, "password")

    email_input.send_keys(email)
    senha_input.send_keys(senha)
    senha_input.send_keys(Keys.RETURN)

    # Aguardar para verificar se o login foi bem-sucedido
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "zb__NEW_MENU_title"))
    )

    try:
        novo_email_button = driver.find_element(By.ID, "zb__NEW_MENU_title")
        novo_email_button.click()
        print("Nova Mensagem gerada com sucesso!")
    except Exception as e:
        print(f"Erro ao tentar clicar no botão de Nova Mensagem: {e}")

    # Carregar os dados do Excel
    excel_file_path = r'c:\Users\user.admin006\Documents\teste.xlsx'
    df = pd.read_excel(excel_file_path)

    if 'E-mail' not in df.columns:
        print("Erro: Não foi encontrada a coluna 'Email' no arquivo Excel.")
    else:
        for email_empresa, nome_empresa in zip(df['E-mail'], df['Empresa']):
            try:
                assunto_base = input(f"Digite o assunto para o e-mail {nome_empresa}: ")
                corpo_email = input('Digite o corpo do e-mail:')

                # Preencher os campos do e-mail
                preencher_campos(email_empresa, assunto_base, corpo_email, nome_empresa)

                # Verificar e enviar o e-mail
                enviar_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="zb__COMPOSE-1__SEND"]'))
                )
                enviar_button.click()
                print(f'E-mail enviado para {email_empresa}')

                # Esperar para garantir que o e-mail seja enviado
                time.sleep(3)

                # Checar se houve erro no envio
                try:
                    erro_envio = driver.find_element(By.XPATH, "//*[contains(text(),'Erro ao enviar')]")
                    print(f'Erro ao enviar e-mail para {email_empresa}: {erro_envio.text}')
                except:
                    print(f"E-mail enviado com sucesso para {email_empresa}")

                # Voltar para a tela de criação de nova mensagem
                novo_email_button = driver.find_element(By.ID, "zb__NEW_MENU_title")
                novo_email_button.click()
                print("Voltando para a tela de criação de nova mensagem!")
            except Exception as e:
                print(f'Erro ao enviar e-mail para {email_empresa}: {e}')

    # Fechar o navegador
    driver.quit()