from selenium import webdriver as opcoesSelenium
from selenium.webdriver.common.by import By
import pyautogui as tempoEspera
import pyautogui as funcoesTeclado
import pandas as pd
import os
import smtplib
from email.message import EmailMessage

navegador = opcoesSelenium.Chrome()

navegador.get("https://www.magazineluiza.com.br/")
navegador.find_element(By.ID, "input-search").send_keys("notebook")
tempoEspera.sleep(2)
funcoesTeclado.press("enter")
tempoEspera.sleep(10)

listaMelhores = []
listaPiores = []

listaProdutos = navegador.find_elements(By.CLASS_NAME, "bDaikj")

for item in listaProdutos:
    nomeProduto = ""
    qtdAvaliacoes = ""
    urlProduto = ""

    try:
        nomeProduto = item.find_element(By.CLASS_NAME, "cQhIqz").text
    except Exception:
        pass

    try:
        qtdAvaliacoes = item.find_element(By.CLASS_NAME, "sc-fUkmAC").text
        # Extrair apenas a quantidade de avaliações dentro dos parêntes
        qtdAvaliacoes = qtdAvaliacoes.split('(')[-1].split(')')[0] 
    except Exception:
        pass

    try:
        urlProduto = item.find_element(By.TAG_NAME, "a").get_attribute("href")
    except Exception:
        pass

    if nomeProduto and qtdAvaliacoes and urlProduto:
        try:
            qtd_int = int(qtdAvaliacoes) 
            
            dadosLinha = [nomeProduto, qtd_int, urlProduto]

            if qtd_int < 100:
                listaPiores.append(dadosLinha) 
            else:
                listaMelhores.append(dadosLinha)

        except (ValueError, IndexError):
            continue

# Salvando no Excel
output_dir = "Output"
os.makedirs(output_dir, exist_ok=True)
output_file = os.path.join(output_dir, 'Notebooks.xlsx')

with pd.ExcelWriter(output_file, engine='xlsxwriter') as arquivoExcel:
    pd.DataFrame(listaMelhores, columns=['PRODUTO', 'QTD_AVAL', 'URL']).to_excel(arquivoExcel, sheet_name='Melhores', index=False)
    pd.DataFrame(listaPiores, columns=['PRODUTO', 'QTD_AVAL', 'URL']).to_excel(arquivoExcel, sheet_name='Piores', index=False)

print("Arquivo salvo em:", output_file)

def enviar_email():
    email_destino = "anamariadoceugomes@gmail.com" 
    email_remetente = "anamariadoceugomes@gmail.com" 
    senha_remetente = "jpbg tcsv uoxc twcd" 
    msg = EmailMessage()
    msg['Subject'] = "Relatório Notebooks"
    msg['From'] = email_remetente
    msg['To'] = email_destino
    msg.set_content("Olá, aqui está o seu relatório dos notebooks extraídos da Magazine Luiza.\n\nAtenciosamente,\nRobô")

    with open(output_file, 'rb') as f:
        msg.add_attachment(f.read(), maintype='application', subtype='xlsx', filename="Notebooks.xlsx")

    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(email_remetente, senha_remetente)
            smtp.send_message(msg)
            print("Email enviado com sucesso!")
    except Exception as e:
        print(f"Erro ao enviar email: {e}")

enviar_email()





