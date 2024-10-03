######################################################################################################
############################### AUTOMAÇÃO DE CRIAÇÃO DE EVENTOS JURIDICOS ############################
######################################################################################################


###################################################################
############################### LIB's #############################
###################################################################


from playwright.sync_api import sync_playwright
from urllib.parse import urlparse, parse_qs
import time
import os
import pandas as pd
from datetime import datetime
import shutil


###################################################################
############################### DEF's #############################
###################################################################

# MANTENDO UM FORMATO PADRÃO DE DATA
def safe_parse_date(x):
    only_date = x[:x.find(" ")]
    if only_date.find("/") > 0 and len(only_date) == 10: 
        return pd.to_datetime(only_date, format="%d/%m/%Y")
    elif only_date.find("-") > 0 and len(only_date) == 10: 
        return pd.to_datetime(only_date, format="%Y-%m-%d")
    else:
        return None


#################################################################
####################### INPUT (ARQUIVOS) ########################
#################################################################


origem = "C://Users//EU//Desktop//INTIMACAO_SIPE"

# Caminho de Origem do Arquivo
caminho_casos = f"{origem}/00 - CASOS"
caminho_mover = f"{origem}/01 - ARQUIVOS CONCLUIDOS"  

# Procura e lê o Arquivo
arquivos = os.listdir(caminho_casos)

# Pega o Arquivo mais recente
arquivo = max(arquivos)

# Caminho do Arquivo Sem Nome
caminho_arquivo = os.path.join(caminho_casos, arquivo)

planilha_principal = pd.read_excel(caminho_arquivo, dtype = str)

# Tratamento da Coluna de Data
planilha_principal["Download INT"] = planilha_principal["Download INT"].apply(safe_parse_date)
planilha_principal["Data"] = planilha_principal["Download INT"].dt.strftime('%d/%m/%Y')


##########################################################################
####################### INICIALIZAÇÃO DO SISTEMA #########################
##########################################################################


p = sync_playwright().start()
browser = p.chromium.launch(headless=False, channel='chrome')
context = browser.new_context(permissions=[])
# Abre Uma Nova Página
page = context.new_page()

# Navegue Para Uma URL
url = page.goto("http://teste.local/")

# USUARIO E SENHA COM VÁRIAVEIS DE AMBIENTE
usuario = os.getenv("USUARIO")
senha = os.getenv("SENHA")

# Usuario
page.fill('//*[@id="LOGINteste/LoginVO_*_login"]',usuario)
# Senha
page.fill('//*[@id="SENHAteste/LoginVO_*_login"]',senha)
# POP-UP
page.click('text="Entendi!"')
# Entrar
page.click('//*[@id="label_ok-FORALL_teste/LoginVO_*_login"]/tbody/tr[2]/td[2]')


##################################################################
####################### PROCESSO SISTEMA #########################
##################################################################


for i in range(planilha_principal.shape[0]):
   
    try:
    
        DATA = planilha_principal.loc[i, 'Data']
        NUMERO_PROCESSO = planilha_principal.loc[i, "Número Processo"]
        DESCRICAO = planilha_principal.loc[i, "Conteúdo"]

        # ENTENDENDO EM QUAL LINHA DA PLANILHA O ROBÔ ESTA REALIZANDO OS PREENCHIMENTOS
        print("Linha: ",i, "Desdobramento: ",NUMERO_PROCESSO)

        if pd.isna(DATA):
            
            planilha_principal.loc[i, "STATUS_ROBO"] = "DATA INVALIDA"
            
        else:

            time.sleep(2)
            
            # Aba Processos
            
            page.locator('.x-tree-node-anchor', has_text='PROCESSOS').click()
            time.sleep(3)

            # Aba Todos
            page.locator('.x-tree-node-anchor', has_text='Todos').click()


            ################################################################################################# Busca Avançada
            page.click('//*[@id="label_search-STATIC_teste/ProcessoVO_*_lista"]/tbody/tr[1]/td[2]')

            # Selecione um Campo
            page.click('text="selecione um campo"')
            page.keyboard.type("Número Atual")
            page.keyboard.press('Enter')

            # Selecione um Operador
            page.click('text="selecione um operador"')
            page.keyboard.type("é igual a")
            page.keyboard.press('Enter')

            # Pesquisa Pelo Número do Processo
            page.click('text="entre com valor"')
            time.sleep(1)
            page.keyboard.type(NUMERO_PROCESSO)
            time.sleep(5)
            
            # Buscar
            page.locator('//*[@id="buscarFILTROteste/ProcessoVO_*_lista_tree"]').get_by_role("button", name= "Buscar").click()
            ################################################################################################################


            # Abre a Pasta
            page.get_by_alt_text("Ativo").click(timeout=12000)
            page.wait_for_timeout(6000)

            # EXTRAINDO A URL ATUAL DA PÁGINA, PARA PEGAR O ID DOS CAMPOS CONTIDOS NE
            ################## Atualiza a Página Para Pegar o Link Atual
            page.reload()
            id_processo = parse_qs(urlparse(page.url).query)["VALUE"][0]
            ############################################################


            # Abre a Pagina de Publicações
            page.locator('span.x-tab-strip-text', has_text='Publicações').click()

            # Aciona o Botão Adicionar
            page.locator('.x-btn-text', has_text='Criar Evento').click()


            ################################################ Preenche o Campo de Data
            page.fill('//*[@name="DATA"]', DATA)
            #########################################################################


            ####################################################### Preenche o Campo de Evento
            page.click(f'//*[@id="ID_EVENTOteste/EventoProcessoVO_{id_processo}_inclui"]')
            page.keyboard.type("Publicação")
            page.locator('.x-layer.x-combo-list.x-resizable-pinned').locator('text="Publicação"').click()
            ##################################################################################


            page.fill(f'//*[@id="DETALHESteste/EventoProcessoVO_{id_processo}_inclui"]', DESCRICAO)


            try:
                
                ################################################################## Preenchimento do Desdobramento
                page.locator(f'//*[@id="x-form-el-ID_PROCESSO_DESDOBRAMENTOteste/EventoProcessoVO_{id_processo}_inclui"]').locator('img[src="/scripts/ext/resources/images/default/s.gif"]').click()
                page.locator('.x-combo-list-item', has_text=NUMERO_PROCESSO).click(timeout=5000)
                #################################################################################################


                # Salva o Arquivo
                page.click(f'//*[@id="label_salvar-STATIC_teste/EventoProcessoVO_{id_processo}_inclui"]/tbody/tr[2]/td[2]')

                # Fecha as Abas
                page.wait_for_timeout(3000)
                page.query_selector('.x-tab-strip-close').click()
                page.wait_for_timeout(3000)
                page.query_selector('.x-tab-strip-close').click()

                # Armazenado Como Bem Sucedido
                planilha_principal.loc[i, "STATUS_ROBO"] = "EFETIVADO"

             

            except:

                page.locator(f'//*[@id="label_cancelar-STATIC_teste/EventoProcessoVO_{id_processo}_inclui"]/tbody/tr[2]/td[2]').click()

                page.query_selector('.x-tab-strip-close').click()

                planilha_principal.loc[i, "STATUS_ROBO"] = "DESDOBRAMENTO DUPLICADO"

    except:

        page.reload()
        
        # Armazenado Como ERRO
        planilha_principal.loc[i, "STATUS_ROBO"] = "NÃO FOI EFETUADO"
        

# Extração dos Dados 
timestart = datetime.now().strftime("%d.%m.%Y_%Hh.%M")

# OUTPUT 1 -------- CASOS INCLUIDOS AGORA
planilha_principal.to_excel(f"{origem}/99 - CONTROLE/CONTROLE_{timestart}.xlsx", index=False)

# Movendo Para Outra Pasta 
lista_arquivos = os.listdir(caminho_casos)

if arquivo in lista_arquivos:
    
    # Configurando o Movimento do Arquivo
    casos = os.path.join(caminho_casos, arquivo)
    destino_MOVE = os.path.join(caminho_mover, arquivo)
    # Move os Arquivos Lidos com Sucesso    
    shutil.move(casos, destino_MOVE)
    
else:
    
    print("Arquivo Não Encontrado")
