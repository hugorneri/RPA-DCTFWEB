import os
import logging
import time
from pathlib import Path
import undetected_chromedriver as uc
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import pandas as pd

# Função para renomear arquivo
def renomear_arquivo_recente(codigo, competencia, pasta_competencia):
    try:
        arquivos = list(Path(pasta_competencia).glob("*"))
        arquivo_recente = max(arquivos, key=os.path.getctime)
        novo_nome = Path(pasta_competencia) / f"{codigo} DARFWEB {competencia}.pdf"
        arquivo_recente.rename(novo_nome)
        logging.info(f"Arquivo renomeado para: {novo_nome}")
    except Exception as e:
        logging.error(f"Erro ao renomear o arquivo: {e}")

# Configurar driver Chrome
def configurar_driver(perfil_path, pasta_competencia):
    options = uc.ChromeOptions()
    options.add_argument("--user-data-dir=" + str(perfil_path))
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-extensions")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-infobars")
    options.add_experimental_option("prefs", {
        "download.default_directory": str(pasta_competencia),
        "download.prompt_for_download": False,
        "profile.default_content_settings.popups": 0,
    })
    driver = uc.Chrome(options=options)
    driver.get('https://cav.receita.fazenda.gov.br/autenticacao/login')
    driver.maximize_window()
    driver.implicitly_wait(10)
    logging.info("Driver configurado com sucesso.")
    return driver

# Login manual
def login(driver):
    try:
        logging.info("Iniciando processo de login manual.")
        print("==== ATENÇÃO ====")
        print("O login precisa ser realizado manualmente.")
        print("Por favor, faça o login e pressione ENTER quando estiver na tela principal.")
        input("Pressione ENTER quando o login estiver concluído...")
        logging.info("Usuário confirmou que o login foi concluído.")
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="linkHome"]')))
            logging.info("Login realizado com sucesso. Página principal identificada.")
            return True
        except:
            logging.warning("Não foi possível confirmar se o login foi bem-sucedido.")
            input("Confirme que está na página principal e pressione ENTER...")
            return True
    except Exception as e:
        logging.error(f"Erro durante o processo de login: {e}")
        print("Ocorreu um erro durante o login. Verifique o arquivo de log para mais detalhes.")
        print("Tente novamente manualmente.")
        input("Pressione ENTER quando o login estiver concluído...")
        return True


# Transmissão
def transmissao(cnpjs, codigos, df, driver, competencia, pasta_competencia, data_inicial, data_final):
    for cnpj, codigo in zip(cnpjs, codigos):
        status = df.loc[df['CNPJ'] == cnpj, 'STATUS'].values
        if len(status) > 0 and ('Guia baixada' in str(status[0]) or pd.isna(status[0])):
            logging.info(f"CNPJ {cnpj} já processado ou sem status. Pulando...")
            continue
        tentativas = 3
        while tentativas > 0:
            try:
                logging.info("Iniciando navegação no sistema.")

                bt_home = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="linkHome"]'))) # Botão Home 
                logging.info("Clicando no botão Home")
                bt_home.click()

                bt_declaracoes = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//li[@id="btn214"]'))) # Botão Declarações e Demonstrativos
                logging.info("Clicando no botão Declarações e Demonstrativos")
                bt_declaracoes.click()

                bt_assinar = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="containerServicos214"]/div[2]/ul/li[1]/a'))) # Assinar e transmitir DCTF
                logging.info("Clicando no botão Assinar e transmitir DCTF")
                bt_assinar.click()

                iframe1 = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmApp"]'))) # Iframe do site
                driver.switch_to.frame(iframe1)
                iframe2 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo-pagina"]/div[2]/iframe'))) # Iframe do Captcha   
                driver.switch_to.frame(iframe2)

                time.sleep(3)
                bt_sou_humano = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="checkbox"]'))) # Botao sou Humano
                logging.info("Clicando no botão Sou Humano (captcha)")
                bt_sou_humano.click()

                driver.switch_to.default_content()
                driver.switch_to.frame(iframe1)

                bt_prosseguir = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_cphConteudo_btnProsseguir"]'))) # Botao Prosseguir
                logging.info("Clicando no botão Prosseguir")
                time.sleep(2)
                bt_prosseguir.click()
                driver.switch_to.default_content()

                iframe = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmApp"]'))) # Iframe do site
                driver.switch_to.frame(iframe)
                bt_sou_procurador = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_cphConteudo_chkListarOutorgantes"]'))) # Botão Sou Procurador
                logging.info("Clicando no botão Sou Procurador")
                bt_sou_procurador.click()
                driver.switch_to.default_content()

                logging.info(f'Iniciando a transmissão da empresa: {cnpj}')
                iframe = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmApp"]'))) # Iframe do site
                driver.switch_to.frame(iframe)

                data_inicio = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="txtDataInicio"]'))) # Campo de data inicial
                data_inicio.clear() # Limpa o campo de data inicial
                data_inicio.send_keys(data_inicial) # Escreve a data inicial

                data_fim = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="txtDataFinal"]'))) # Campo de data final
                data_fim.clear() # Limpa o campo de data final
                data_fim.send_keys(data_final) # Escreve a data final

                bt_ortogante = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_UpdatePanelListaOutorgantes"]/div/div[2]/div/div/div/button')))
                logging.info("Clicando no botão Ortogante")
                bt_ortogante.click() # Campo ortogante onde tem entrada do cnpj do cliente

                bt_nenhum = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_UpdatePanelListaOutorgantes"]/div/div[2]/div/div/div/div/div[2]/div/button[2]')))
                logging.info("Clicando no botão Nenhum")
                bt_nenhum.click() # remover clientes selecionados antes de inserir o novo cnpj para pesquisar

                campo_cnpj = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_UpdatePanelListaOutorgantes"]/div/div[2]/div/div/div/div/div[1]/input')))
                campo_cnpj.send_keys(cnpj) # Insere o CNPJ na barra de busca

                selecionar_cnpj = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_UpdatePanelListaOutorgantes"]/div/div[2]/div/div/div/div/ul')))
                selecionar_cnpj.click() # Seleciona o CNPJ na barra de busca

                bt_pesquisar = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_btnFiltar"]')))
                bt_pesquisar.click() # Clica para buscar o Cliente

                try:
                    bt_visualizar = WebDriverWait(driver, 15).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs_ctl02_lbkVisualizarDctf"]')))
                    logging.info("Clicando no botão Visualizar")
                    bt_visualizar.click() # Clica no botão visualizar
                    
                except Exception as e:
                    logging.info("Nenhuma declaração encontrada.")
                    df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Nenhuma declaração encontrada' # Atualiza o status do cliente na planilha
                    df.to_excel('database.xlsx', index=False)
                    driver.switch_to.default_content()
                    break

                bt_emitir_darf = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="LinkEmitirDARFIntegral"]')))
                logging.info("Clicando no botão Emitir DARF")
                bt_emitir_darf.click() # Clica no botão emitir DARF

                time.sleep(5)
                renomear_arquivo_recente(codigo, competencia, pasta_competencia)

                logging.info(f"Download successful for {cnpj}")
                df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Guia baixada'
                df.to_excel('database.xlsx', index=False)

                bt_ok = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//button[text()='OK']")))
                logging.info("Clicando no botão OK")
                bt_ok.click()
                driver.switch_to.default_content()

                df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Guia baixada'
                df.to_excel('database.xlsx', index=False)
                break
                
            except Exception as e:
                logging.error(f"Erro no processamento do cliente {cnpj}: {e}")
                df.loc[df['CNPJ'] == cnpj, 'STATUS'] = 'Erro no download'
                df.to_excel('database.xlsx', index=False)
                tentativas -= 1
                if tentativas > 0:
                    logging.info(f"Tentando novamente ({tentativas} tentativas restantes)")
                    time.sleep(3)
                    transmissao(cnpjs, codigos, df, driver, competencia, pasta_competencia, data_inicial, data_final)
                else:
                    logging.error(f"Falha após {3} tentativas para o cliente {cnpj}") 