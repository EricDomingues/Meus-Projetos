#################################################
## ROBÔ DE CONSULTA DO VALOR DE ALGUMAS MOEDAS ##
#################################################

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

import pyautogui as gui

import xlsxwriter as ex

while True:

    driver = webdriver.Chrome()
    driver.get("https://www.google.com.br/?hl=pt-BR")

    # Consulta do Dolar
    gui.sleep(5)
    driver.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys("Dolar hoje")
    gui.press('enter')

    gui.sleep(4)
    valor_dolar = driver.find_elements(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')[0].text
    #print("O Valor do Dolar atualmente é:", valor_dolar)
    ###################

    # Consulta do Euro
    gui.sleep(5)
    zerar_escrita = driver.find_element(By.XPATH, '//*[@id="tsf"]/div[1]/div[1]/div[2]/div/div[3]/div[1]/div').click()
    driver.find_element(By.XPATH, '//*[@id="APjFqb"]').send_keys("Euro hoje")
    gui.press('enter')

    gui.sleep(4)
    valor_euro = driver.find_elements(By.XPATH, '//*[@id="knowledge-currency__updatable-data-column"]/div[1]/div[2]/span[1]')[0].text
    #print("O Valor do Euro atualmente é:", valor_euro)
    ##################

    print("O Valor do Euro atualmente é:", valor_euro)
    print("O Valor do Dolar atualmente é:", valor_dolar)

    # Inicializando arquivo e nomeando
    caminho_arquivo = "C:/Users/Dell latitude/Desktop/Estudo Python - Éric/Robô/Robo_1.xlsx"
    workbook = ex.Workbook(caminho_arquivo)

    # Adicionando uma planilha
    planilha1 = workbook.add_worksheet()

    # Preenchendo a planilha
    planilha1.write('A1', 'Dolar')
    planilha1.write('B1', 'Euro')
    planilha1.write('A2', valor_dolar)
    planilha1.write('B2', valor_euro)

    # Salva o arquivo
    workbook.close()
    
    driver.quit()
    
    # Espera 1 minuto para rodar o código novamente
    gui.sleep(60)
