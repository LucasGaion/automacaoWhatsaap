# Bibliotecas
import time
import os
from datetime import date
import pandas as pd
from openpyxl import Workbook

from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Configuração das constantes para localizar elementos na página
LOGIN_USERNAME = "//*[@id='root']/main/div/form/div[1]/input"
LOGIN_PASSWORD = "//*[@id='root']/main/div/form/div[2]/input"
ADD_CONTACT_BUTTON = "//*[@id='root']/div/section/section/main/nav/div[1]/a"
CONTACT_PHONE_INPUT = "//*[@id='root']/div/section/section/main/nav/div[1]/input[2]"
SEND_CONTACT_BUTTON = "//*[@id='root']/div/section/section/main/nav/div[1]/div[2]"
TEXT_AREA = "//*[@id='root']/div/section/section/main/section/footer/textarea"
END_SESSION_BUTTON = "//*[@id='root']/div/section/section/main/section/header/div/div[3]/button/span/div"
END_SESSION_BUTTON_OPTION = "//*[@id='root']/div/section/section/main/section/header/div/div[3]/div/div[7]/div"

print(" -------------- Programa Iniciando------------------\n")

print(" -> Logando no ChatPro\n")
# Acessando o site colocando username e password
time.sleep(0.1)
username = "lucas.gaion@fastcash.com.br"
password = "f45tc45hLucas"
login_page = "https://app.chatpro.com.br/login"

print(" -> Abrindo a Planilha e Lendo\n")
# abrindo a planilha de entrada e escrevendo nas colunas
file_path = "cobranca.xlsx"
df = pd.read_excel(file_path)

client_names = []
phones = []
payment_links = []
due_dates = []

# Criação de um novo workbook para armazenar os resultados
results_wb = Workbook()
results_ws = results_wb.active
results_ws.append(['Nome do Cliente', 'Celular', 'Link de Cobrança', 'Data de Vencimento', 'Status'])

for index, row in df.iterrows():
    # Itera sobre as linhas do arquivo excel e extrai os valores das colunas correspondentes
    client_names.append(str(row['Nome_do_Cliente']))
    phones.append(str(row['Celular_do_Cliente']))
    payment_links.append(str(row['Link_de_Cobrança']))
    due_dates.append(pd.to_datetime(row['Data_de_Vencimento']).date())

driver = webdriver.Firefox()
driver.maximize_window()
driver.get(login_page)

wait = WebDriverWait(driver, 10)
wait.until(EC.presence_of_element_located((By.XPATH, LOGIN_USERNAME)))

driver.find_element(By.XPATH, LOGIN_USERNAME).send_keys(username)
driver.find_element(By.XPATH, LOGIN_PASSWORD).send_keys(password)
driver.find_element(By.XPATH, LOGIN_PASSWORD).send_keys(Keys.RETURN)

for i in range(len(client_names)):
    client_name = client_names[i]
    phone = phones[i]
    payment_link = payment_links[i]
    due_date = due_dates[i]

    # cobranca em atraso
    due_day_message = "CRÉDITO FRUBANA\n\nOlá *" + client_name + ",* \n\nEste é um lembrete para pagamento do seu crédito que vence *HOJE ou AMANHÃ.* \nA cada dia em atraso é acrescido juros e multa.\n\nPara efetuar o pagamento referente ao seu crédito, clique no link abaixo (ou copie e cole em seu navegador):\n\n*" + payment_link + "* \n\nLembrando que a data limite de vencimento é:\n*" + due_date.strftime \
        ('%d/%m/%Y') + "* \n\nSe você já tiver realizado o pagamento, por favor, desconsidere essa mensagem.\n\nQualquer dúvida, estamos a disposição!"
    first_day_message = "Olá, *" + client_name + ",* seu pagamento venceu ontem. Se você esqueceu de pagar, não tem problema, você pode pagar pelo link a seguir: *" + payment_link + "* \nÉ importante realizar o pagamento para evitar bloqueio dessa opção do Crédito Frubana e evitar juros e multas.\nSe você já tiver realizado o pagamento, por favor, desconsidere essa mensagem."
    fourth_day_message = "Olá, *" + client_name + ",* não localizamos o seu pagamento e por isso seu Crédito Frubana foi bloqueado. Para desbloquear essa opção e evitar juros diários, você ainda pode pagar pelo link a seguir: *" + payment_link + "* \nSe você já tiver realizado o pagamento, por favor, desconsidere essa mensagem."
    seventh_day_message = "Olá, *" + client_name + ",* verificamos que seu Crédito Frubana se encontra atrasado a mais de 7 dias. Mas você ainda pode pagar pelo link a seguir: *" + payment_link + "* e reativar seu Crédito Frubana e evitar juros, multas e negativação do seu nome.\nSe você já tiver realizado o pagamento, por favor, desconsidere essa mensagem."

    # cobranca em em dia
    day_message = "CRÉDITO FRUBANA\n\nOlá *" + client_name + ",* \n\nEste é um lembrete para pagamento do seu crédito que vence *HOJE.* \nA cada dia em atraso é acrescido juros e multa.\n\nPara efetuar o pagamento referente ao seu crédito, clique no link abaixo (ou copie e cole em seu navegador):\n\n*" + payment_link + "* \n\nLembrando que a data limite de vencimento é:\n*" + due_date.strftime(
        '%d/%m/%Y') + "* \n\nSe você já tiver realizado o pagamento, por favor, desconsidere essa mensagem.\n\nQualquer dúvida, estamos a disposição!"
    one_day_message = "CRÉDITO FRUBANA\n\nOlá *" + client_name + ",* \n\nEste é um lembrete para pagamento do seu crédito que vence *AMANHÃ.* \nA cada dia em atraso é acrescido juros e multa.\n\nPara efetuar o pagamento referente ao seu crédito, clique no link abaixo (ou copie e cole em seu navegador):\n\n*" + payment_link + "* \n\nSe você já tiver realizado o pagamento, por favor, desconsidere essa mensagem.\n\nQualquer dúvida, estamos a disposição!"
    three_days_message = "CRÉDITO FRUBANA\n\nOlá *" + client_name + ",* \n\nEste é um lembrete para pagamento do seu crédito que vencerá em *3 dias.* \nA cada dia em atraso é acrescido juros e multa.\n\nPara efetuar o pagamento referente ao seu crédito, clique no link abaixo (ou copie e cole em seu navegador):\n\n*" + payment_link + "* \n\nSe você já tiver realizado o pagamento, por favor, desconsidere essa mensagem.\n\nQualquer dúvida, estamos a disposição!"
    four_days_message = "CRÉDITO FRUBANA\n\nOlá *" + client_name + ",* \n\nEste é um lembrete para pagamento do seu crédito que vencerá em *4 dias.* \nA cada dia em atraso é acrescido juros e multa.\n\nPara efetuar o pagamento referente ao seu crédito, clique no link abaixo (ou copie e cole em seu navegador):\n\n*" + payment_link + "* \n\nSe você já tiver realizado o pagamento, por favor, desconsidere essa mensagem.\n\nQualquer dúvida, estamos a disposição!"
    seven_days_message = "CRÉDITO FRUBANA\n\nOlá *" + client_name + ",* \n\nEste é um lembrete para pagamento do seu crédito que vencerá em *7 dias.* \nA cada dia em atraso é acrescido juros e multa.\n\nPara efetuar o pagamento referente ao seu crédito, clique no link abaixo (ou copie e cole em seu navegador):\n\n*" + payment_link + "* \n\nSe você já tiver realizado o pagamento, por favor, desconsidere essa mensagem.\n\nQualquer dúvida, estamos a disposição!"

    current_date = date.today()

    # Verifica o número de dias de atraso do pagamento
    if (current_date - due_date).days <= 0:
        if (current_date - due_date).days == 0:
            message = day_message
        elif (current_date - due_date).days == -1:
            message = one_day_message
        elif (current_date - due_date).days == -3:
            message = three_days_message
        elif (current_date - due_date).days == -4:
            message = four_days_message
        elif (current_date - due_date).days == -7:
            message = seven_days_message
        else:
            message = due_day_message
    elif (current_date - due_date).days == 1:
        message = first_day_message
    elif (current_date - due_date).days == 4:
        message = fourth_day_message
    elif (current_date - due_date).days >= 7:
        message = seventh_day_message



    wait.until(EC.element_to_be_clickable((By.XPATH, ADD_CONTACT_BUTTON)))
    driver.find_element(By.XPATH, ADD_CONTACT_BUTTON).click()

    try:
        element = wait.until(EC.element_to_be_clickable((By.ID, 'root')))
    except TimeoutException:
        print("Loading took too much time!")
        driver.quit()

    driver.find_element(By.XPATH, CONTACT_PHONE_INPUT).send_keys(phone)
    driver.find_element(By.XPATH, SEND_CONTACT_BUTTON).click()
    time.sleep(0.1)  # Reduzido para 1 segundo

    try:
        # Verifica se o botão "OK" está presente
        ok_button_xpath = "/html/body/div[2]/div/div[3]/button[1]"
        ok_button = wait.until(EC.element_to_be_clickable((By.XPATH, ok_button_xpath)))
        if ok_button.is_displayed():
            ok_button.click()
            df.at[index, 'Status'] = 'NUMERO NÃO ENCONTRADO'   # Atualiza o status no dataframe
            results_ws.append([client_name, phone, payment_link, due_date.strftime('%d/%m/%Y'), 'NUMERO NÃO ENCONTRADO'])   # Adiciona o resultado na planilha de resultados quando nao estiver o numero
            continue
    except TimeoutException:
        pass

    wait.until(EC.element_to_be_clickable((By.XPATH, TEXT_AREA))).send_keys(message + Keys.RETURN)

    try:
        # Verifica se o botão "End Session" está presente
        end_session_button = wait.until(EC.element_to_be_clickable((By.XPATH, END_SESSION_BUTTON)))
        if end_session_button.is_displayed():
            end_session_button.click()
            time.sleep(0.0)
            driver.find_element(By.XPATH, END_SESSION_BUTTON_OPTION).click()
    except TimeoutException:
        pass

    # Atualiza o status no dataframe e adiciona o resultado na planilha de resultados
    df.at[index, 'Status'] = 'ENVIADO'
    results_ws.append([client_name, phone, payment_link, due_date.strftime('%d/%m/%Y'), 'ENVIADO'])

# Salva o dataframe atualizado e fecha o driver
df.to_excel(file_path, index=False)
driver.quit()

# Salva a planilha de resultados
print(" -> Salvo na Planilha de [RESPOSTA]\n")
results_wb.save('respostaWhats.xlsx')


print(" -------------- Programa Finalizado------------------\n")

