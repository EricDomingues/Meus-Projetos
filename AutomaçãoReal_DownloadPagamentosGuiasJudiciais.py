##############################################################################################################################################################################
################################# AUTOMAÇÃO DA IDENTIFICAÇÃO E DOWNLOAD DE COMPROVANTES DE PAGAMENTO DE GUIAS JUDICIAIS (CUSTOS PROCESSUAIS ##################################
##############################################################################################################################################################################



##########################################################################
################################# LIB's ##################################
##########################################################################


import os
from datetime import datetime
import time
import shutil
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


##########################################################################
################################# DEF's ##################################
##########################################################################


######################### FUNÇÃO DE ORGANIZAÇÃO DOS ARQUIVOS POR TEMPO, PARA RENOMEA-LOS (Função 1)
def obter_arquivo_mais_recente(diretorio):
    # Lista todos os arquivos no diretório
    arquivos = os.listdir(diretorio)
    
    # Verifica se o diretório está vazio
    if not arquivos:
        raise FileNotFoundError("Nenhum arquivo encontrado no diretório.")
    
    # Cria uma lista de tuplas (caminho_completo, data_modificacao)
    caminhos_completos = [os.path.join(diretorio, arquivo) for arquivo in arquivos]
    arquivos_com_data = [(caminho, os.path.getmtime(caminho)) for caminho in caminhos_completos]

    # Ordena a lista pela data de modificação (mais recente primeiro)
    arquivos_com_data.sort(key=lambda x: x[1], reverse=True)
    
    # Retorna o caminho do arquivo mais recente
    return arquivos_com_data[0][0]
#########################################################################################


################ FUNÇÃO DE IDENTIFICAÇÃO DE ";" QUANDO HOUVER MAIS DE UM CÓDIGO DE BARRAS (Função 2)
def expand_rows(planilha_1113, column_with_multiple_values):
    # Cria uma lista para armazenar as novas linhas
    rows = []
    
    # Itera sobre cada linha do DataFrame
    for _, row in planilha_1113.iterrows():
        # Se a célula contém múltiplos valores separados por ponto e vírgula
        values = row[column_with_multiple_values].split(';')
        for value in values:
            # Cria uma nova linha com os valores divididos
            new_row = row.copy()
            new_row[column_with_multiple_values] = value
            rows.append(new_row)
    
    # Cria um novo DataFrame a partir das novas linhas
    new_df = pd.DataFrame(rows)
    
    return new_df
##########################################################################################


################################ FUNÇÃO DE PADRONIZAÇÃO DA FORMATAÇÃO DAS DATAS (Função 3)
def safe_parse_date(x):
    only_date = x[:x.find(" ")]
    if only_date.find("/") > 0 and len(only_date) == 10: 
        return pd.to_datetime(only_date, format="%d/%m/%Y")
    elif only_date.find("-") > 0 and len(only_date) == 10: 
        return pd.to_datetime(only_date, format="%Y-%m-%d")
    else:
        return None
##########################################################################################


##########################################################################
######################### CONFIGURAÇÃO DO DIRETÓRIO ######################
##########################################################################


origem = "C://Users//EU//Desktop//CODIGO DE BARRAS"

caminho_download_web_comprovantes = "C://Users//EU//Desktop//CODIGO DE BARRAS//02 - COMPROVANTES"

# Caminho de Origem do Arquivo
caminho_casos = f"{origem}/00 - CASOS"
caminho_concluidos = f"{origem}/01 - ARQUIVOS CONCLUIDOS"
caminho_comprovantes = f"{origem}/02 - COMPROVANTES" 
caminho_input_banco = f"{origem}/98 - INPUT BANCO"  
caminho_controle = f"{origem}/99 - CONTROLE"


###########################################################################
############################# INPUT JURIDICO - 1113 #######################
###########################################################################


# Procura e lê o Arquivo de Controle dos Casos do 1113
itens_casos = os.listdir(caminho_casos)

# Pega o Arquivo mais recente
arquivo_caso = itens_casos[0]

# Caminho do Arquivo Sem Nome
caminho_arquivo_1113 = os.path.join(caminho_casos, arquivo_caso)

planilha_1113 = pd.read_excel(caminho_arquivo_1113, dtype=str)

# Usando a Função 2
planilha_1113 = expand_rows(planilha_1113, 'CÓDIGO DE BARRAS')

# Formatação Sem Pontuação e Espaços 
planilha_1113['CÓDIGO DE BARRAS'] = planilha_1113['CÓDIGO DE BARRAS'].astype(str).str.replace('.', '', regex=False).str.replace(' ', '', regex=False)

# Recorta a Parte Correta do Código de Barras Para a Planilha do 1113
planilha_1113['COD_BARRAS_1113'] = planilha_1113['CÓDIGO DE BARRAS'].str[10:20] + planilha_1113['CÓDIGO DE BARRAS'].str[21:31]

# Tratamento da Coluna de Data
planilha_1113["VENCIMENTO"] = planilha_1113["VENCIMENTO"].apply(safe_parse_date)
planilha_1113["VENCIMENTO"] = planilha_1113["VENCIMENTO"].dt.strftime('%d/%m/%Y')


##########################################################################
########################### INPUT BANCO DE DADOS #########################
##########################################################################


# Procura e lê o Arquivo da Consulta do Banco de Dados
itens_banco = os.listdir(caminho_input_banco)

# Pega o Arquivo mais recente
arquivo_banco = max(itens_banco)

# Caminho do Arquivo Sem Nome
caminho_arquivo_banco = os.path.join(caminho_input_banco, arquivo_banco)

planilha_consulta_banco = pd.read_excel(caminho_arquivo_banco, dtype=str)

# Recorta a Parte Correta do Código de Barras Para a Planilha do Banco e Tratamento de Dados
planilha_consulta_banco['COD_BARRAS_BANCO'] = planilha_consulta_banco['COD_BARRAS'].str[24:]

# Formatação
planilha_consulta_banco['VALOR'] = planilha_consulta_banco['VALOR'].astype(str).str.replace('.', ',', regex=False)


##########################################################################
######################## UNIFICAÇÃO DAS PLANILHAS ########################
##########################################################################


planilha_principal = pd.merge(planilha_1113, planilha_consulta_banco, left_on='COD_BARRAS_1113', right_on='COD_BARRAS_BANCO', how='left')

# Merge com chaves completas para casos onde o merge inicial não encontrou correspondência
resultado_full_match = pd.merge(planilha_1113, planilha_consulta_banco, left_on='CÓDIGO DE BARRAS', right_on='COD_BARRAS', how='left')

# Atualizar a coluna NSU onde o merge inicial falhou
resultado_full_match = planilha_principal.assign(NSU=resultado_full_match['NSU'], VALOR=resultado_full_match['VALOR'])

# Atualizar a coluna NSU onde o merge inicial falhou
planilha_principal['NSU'] = planilha_principal['NSU'].combine_first(resultado_full_match['NSU'])
planilha_principal['VALOR'] = planilha_principal['VALOR'].combine_first(resultado_full_match['VALOR'])

planilha_principal = planilha_principal.drop(columns=['CÓDIGO DE BARRAS', 'COD_BARRAS', 'COD_BARRAS_1113', 'COD_BARRAS_BANCO' ])


###########################################################################
######################### INICIALIZAÇÃO DO SISTEMA ########################
###########################################################################


# CRIA UM BOTÃO DE DOWNLOAD E ESPECIFICA UM DIRETÓRIO PADRÃO PARA OS DOWNLOADS
chrome_options = Options()
chrome_options.add_experimental_option('prefs',  {
"download.default_directory": caminho_download_web_comprovantes,
"download.prompt_for_download": False,
"download.directory_upgrade": True,
"plugins.always_open_pdf_externally": True
    }
)

driver = webdriver.Chrome(options = chrome_options)
driver.get("http://SISTEMA/PAGINA1")

espera = WebDriverWait(driver, 10)
espera.until(EC.number_of_windows_to_be(2))
driver.switch_to.window(driver.window_handles[1])

# USUARIO E SENHA CONFIGURADO COM A UTILIZAÇÃO DE VARIAVEIS DE AMBIENTE
usuario = os.getenv("USUARIO")
senha = os.getenv("SENHA")

# Usuario
espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmcaindex:NOME_USUARIO"]'))).send_keys(usuario)
# Senha
espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmcaindex:SENHA_ATUAL"]'))).send_keys(senha)

# Login
espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmcaindex:btnOk"]'))).click()

# Botão Continuar
driver.get("http://SISTEMA/PAGINA2")
espera.until(EC.element_to_be_clickable((By.NAME, "page:frmSideMenu:_id42"))).click()

# Botão OK
espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="page:frmeb_filtroini:btnok"]/input'))).click()


##########################################################################
############################ PROCESSO NO E-BANK ##########################
##########################################################################


# Aba de "Relatório"
espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmTopLayoutDoubleMenu:topmenu:t_6"]'))).click()

# Aba "Pagamentos"
espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmTopLayoutDoubleMenu:page:t_6_6"]'))).click()

# Aba "Recibo de Títulos"
espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmTopLayoutDoubleMenu:page:t_6_6_8"]'))).click()


for i in range(planilha_principal.shape[0]):

    try:

        nsu = str(planilha_principal.loc[i, "NSU"])
        pasta = str(planilha_principal.loc[i, "PASTA"])
        valor = str(planilha_principal.loc[i, "VALOR"])
        vencimento = str(planilha_principal.loc[i, "VENCIMENTO"]).replace("/",".")

        ###################################################### DOWNLOAD DOS COMPROVANTES ######################################################

        if nsu == "nan":

            planilha_principal.loc[i, "STATUS_ROBO"] = "COMPROVANTE INEXISTENTE"

        else:

            ################################################################## Preenchimento do Campo de NSU
            driver.switch_to.frame("conteudo")
            preenchimento_nsu = driver.find_element(By.XPATH, '//*[@id="page:frmeb_relrecibotitulo_l:nsu"]')
            preenchimento_nsu.clear()
            preenchimento_nsu.send_keys(nsu)
            ################################################################################################


            # Exibe o Comprovante
            espera.until(EC.presence_of_element_located((By.XPATH, '//*[@id="page:frmeb_relrecibotitulo_l:Layer5"]/input'))).click()

            # Download do Comprovante
            iframe = driver.find_elements(By.TAG_NAME,'iframe')[0]
            driver.switch_to.frame(iframe)
            driver.find_element(By.XPATH, '//*[@id=\"open-button\"]').click()

            time.sleep(5) 

            comprovante_recente = obter_arquivo_mais_recente(caminho_comprovantes)
            caminho_antigo = os.path.join(caminho_comprovantes, comprovante_recente)
            comprovante_rename = f"{pasta}_{valor}_{vencimento}_{i}.pdf"
            caminho_novo = os.path.join(caminho_comprovantes, comprovante_rename)

            os.rename(caminho_antigo, caminho_novo)


            ######################## Fechar a Aba do Download
            driver.switch_to.window(window_name="")
            driver.close()
            #################################################


            ########################## Retorna para a Aba Original
            driver.switch_to.window(driver.window_handles[0])
            driver.switch_to.frame("conteudo")
            iframe = driver.find_elements(By.TAG_NAME,'iframe')[0]
            driver.switch_to.frame(iframe)
            ######################################################

            driver.back()

            ########################################################################################################################################


            planilha_principal.loc[i, "STATUS_ROBO"] = "EFETIVADO"
        
    except:

        planilha_principal.loc[i, "STATUS_ROBO"] = "NÃO FOI EFETUADO"


timestart = datetime.now().strftime("%d.%m.%Y_%Hh.%M")

# OUTPUT 1 -------- COMPROVANTES BAIXADOS AGORA
planilha_principal.to_excel(f"{origem}/99 - CONTROLE/CONTROLE_{timestart}.xlsx", index=False)
