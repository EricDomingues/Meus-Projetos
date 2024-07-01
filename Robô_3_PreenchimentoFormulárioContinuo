##############################################################
## ROBÔ DE PREENCHIMENTO AUTOMÁTICO DE FORMULARIO DO CHROME ##
##############################################################

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from openpyxl import load_workbook as lw

# Inicialização do Arquivo
caminho_arquivo = "C:/Users/Dell latitude/Desktop/Estudo Python - Éric/Robô/DadosFormulario.xlsx"
planilha_aberta = lw(filename=caminho_arquivo)
sheet_sel = planilha_aberta['Dados']

# Faz o reconhecimento de todos os campos da palnilha a serem utilizados parao preenchimento continuamente
for linha in range(2, len(sheet_sel['A']) + 1):
    
    # Variaveis da Planilha
    nome = sheet_sel[f'A{linha}'].value
    email = sheet_sel[f'B{linha}'].value
    telefone = sheet_sel[f'C{linha}'].value
    sexo = sheet_sel[f'D{linha}'].value
    cargo = sheet_sel[f'E{linha}'].value

    # Abrindo Site
    driver = webdriver.Chrome()
    driver.get("https://pt.surveymonkey.com/r/WLXYDX2")
    espera = WebDriverWait(driver, 10)
    
    # Mapeamento Formulário
    nome_f = espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="166517069"]')))
    nome_f.send_keys(nome)
    
    email_f = espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="166517072"]')))
    email_f.send_keys(email)
    
    telefone_f = espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="166517070"]')))
    telefone_f.send_keys(telefone)
    
    cargo_f = espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="166517073"]')))
    cargo_f.send_keys(cargo)
    
    # Preenchimento do Sexo
    if sexo == "Masculino":
        driver.find_element(By.XPATH, '//*[@id="166517071_1215509812_label"]/span[1]').click()
    else: 
        driver.find_element(By.XPATH, '//*[@id="166517071_1215509813_label"]/span[1]').click()
    
    # Envio Formulário    
    enviar_f = driver.find_element(By.XPATH, '//*[@id="patas"]/main/article/section/form/div[2]/button').click()

driver.quit()
