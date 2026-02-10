"""
Módulo de automação para transmissão DCTF no e-CAC.
"""
import logging
import time
from pathlib import Path
from typing import Callable, Optional

import undetected_chromedriver as uc
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

from src.utils import get_chrome_version, renomear_arquivo_recente
from src.planilha import atualizar_status


def configurar_driver(pasta_competencia):
    """
    Configura e retorna uma instância do ChromeDriver com o diretório de download definido.
    O Chrome usa um perfil temporário (sem cache persistente).

    Args:
        pasta_competencia (str or Path): Pasta onde os arquivos serão baixados.

    Returns:
        driver (uc.Chrome): Instância do Chrome configurada.

    Raises:
        Exception: Com mensagem completa do erro e dicas de solução em caso de falha.
    """
    dicas = (
        "Dicas: (1) Feche todas as janelas do Chrome antes de iniciar. "
        "(2) Atualize o Google Chrome (Menu ⋮ → Ajuda → Sobre o Google Chrome). "
        "(3) Se o erro for de versão (Chrome 144 vs driver 145), apague a pasta de cache do driver e tente de novo: "
        "%APPDATA%\\undetected_chromedriver "
        "(4) Reinstale o pacote: pip install -U undetected-chromedriver"
    )

    try:
        options = uc.ChromeOptions()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_argument("--disable-extensions")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-infobars")
        options.add_experimental_option("prefs", {
            "download.default_directory": str(pasta_competencia),
            "download.prompt_for_download": False,
            "profile.default_content_settings.popups": 0,
        })

        # Usar a versão do Chrome instalada para baixar o driver compatível (evita erro 145 vs 144)
        version_main = get_chrome_version()
        if version_main is not None:
            logging.info(f"Chrome detectado: versão principal {version_main}. Usando driver compatível.")
            driver = uc.Chrome(options=options, version_main=version_main)
        else:
            driver = uc.Chrome(options=options)
        driver.get('https://cav.receita.fazenda.gov.br/autenticacao/login')
        driver.maximize_window()
        driver.implicitly_wait(10)
        logging.info("Driver configurado com sucesso.")
        return driver

    except Exception as e:
        msg_original = str(e).strip()
        logging.error(f"Falha ao configurar driver: {msg_original}")
        raise Exception(
            f"Não foi possível iniciar o navegador.\n\n"
            f"Erro original: {msg_original}\n\n"
            f"{dicas}"
        ) from e


def login(driver, callback: Optional[Callable[[str], None]] = None):
    """
    Realiza o processo de login manual no e-CAC, aguardando confirmação do usuário.
    
    Args:
        driver (uc.Chrome): Instância do Chrome já aberta na página de login.
        callback: Função opcional para feedback de status (para GUI).
        
    Returns:
        bool: True se o login foi confirmado pelo usuário.
    """
    def notify(msg):
        if callback:
            callback(msg)
        print(msg)
    
    try:
        logging.info("Iniciando processo de login manual.")
        notify("==== ATENÇÃO ====")
        notify("O login precisa ser realizado manualmente.")
        notify("Por favor, faça o login no navegador.")
        
        # No modo GUI, não usa input() - a GUI controla o fluxo
        if callback is None:
            input("Pressione ENTER quando o login estiver concluído...")
        
        logging.info("Verificando se login foi concluído.")
        
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//*[@id="linkHome"]'))
            )
            logging.info("Login realizado com sucesso. Página principal identificada.")
            return True
        except TimeoutException:
            logging.warning("Não foi possível confirmar se o login foi bem-sucedido.")
            return True
            
    except Exception as e:
        logging.error(f"Erro durante o processo de login: {e}")
        return True


def transmissao(
    cnpjs, 
    codigos, 
    df, 
    driver, 
    competencia, 
    pasta_competencia, 
    data_inicial, 
    data_final,
    timeout_elemento: int = 30,
    tentativas_por_cnpj: int = 3,
    callback: Optional[Callable[[str, int, int], None]] = None,
    should_stop: Optional[Callable[[], bool]] = None,
    planilha_path: Optional[str] = None
):
    """
    Realiza o processo de transmissão e download dos DARFs para cada cliente da lista.
    
    Args:
        cnpjs (list): Lista de CNPJs.
        codigos (list): Lista de códigos dos clientes.
        df (pd.DataFrame): DataFrame da planilha de clientes.
        driver (uc.Chrome): Instância do Chrome já logada.
        competencia (str): Competência (ex: '06 2025').
        pasta_competencia (str or Path): Pasta de download dos arquivos.
        data_inicial (str): Data inicial do filtro.
        data_final (str): Data final do filtro.
        timeout_elemento (int): Tempo máximo de espera por elementos (segundos).
        tentativas_por_cnpj (int): Número de tentativas por CNPJ.
        callback: Função para reportar progresso (mensagem, atual, total).
        should_stop: Função que retorna True se deve parar a execução.
        planilha_path: Caminho para salvar a planilha (opcional).
    """
    total = len(cnpjs)
    planilha_save_path = planilha_path or 'database.xlsx'
    
    for idx, (cnpj, codigo) in enumerate(zip(cnpjs, codigos)):
        # Verificar se deve parar
        if should_stop and should_stop():
            logging.info("Execução interrompida pelo usuário.")
            break
        
        # Reportar progresso
        if callback:
            callback(f"Processando {cnpj}...", idx + 1, total)
        
        # Verificar status - CORRIGIDO: só pula se já foi baixada com sucesso
        cnpj_str = str(cnpj).strip()
        status = df.loc[df['CNPJ'] == cnpj_str, 'STATUS'].values
        
        if len(status) > 0 and 'Guia baixada' in str(status[0]):
            logging.info(f"CNPJ {cnpj} já processado com sucesso. Pulando...")
            continue
        
        tentativas = tentativas_por_cnpj
        sucesso = False
        
        while tentativas > 0 and not sucesso:
            # Verificar se deve parar
            if should_stop and should_stop():
                logging.info("Execução interrompida pelo usuário.")
                break
            
            try:
                # Garantir que estamos no contexto principal antes de começar
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass
                
                logging.info(f"Iniciando navegação no sistema para CNPJ {cnpj}.")

                bt_home = WebDriverWait(driver, timeout_elemento).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="linkHome"]'))
                )
                logging.info("Clicando no botão Home")
                bt_home.click()
                time.sleep(2)  # Aguardar página principal carregar completamente

                bt_declaracoes = WebDriverWait(driver, timeout_elemento).until(
                    EC.element_to_be_clickable((By.XPATH, '//li[@id="btn214"]'))
                )
                logging.info("Clicando no botão Declarações e Demonstrativos")
                time.sleep(1)
                bt_declaracoes.click()
                time.sleep(2)  # Aguardar submenu expandir

                bt_assinar = WebDriverWait(driver, timeout_elemento).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="containerServicos214"]/div[2]/ul/li[1]/a'))
                )
                logging.info("Clicando no botão Assinar e transmitir DCTF")
                time.sleep(1)  # Aguardar link ficar visível
                bt_assinar.click()
                time.sleep(2)  # Aguardar página carregar antes de buscar iframe

                iframe = WebDriverWait(driver, timeout_elemento).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="frmApp"]'))
                )
                driver.switch_to.frame(iframe)
                
                bt_sou_procurador = WebDriverWait(driver, timeout_elemento).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_cphConteudo_chkListarOutorgantes"]'))
                )
                logging.info("Clicando no botão Sou Procurador")
                bt_sou_procurador.click()
                driver.switch_to.default_content()

                logging.info(f'Iniciando a transmissão da empresa: {cnpj}')
                iframe = WebDriverWait(driver, timeout_elemento).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="frmApp"]'))
                )
                driver.switch_to.frame(iframe)

                data_inicio = WebDriverWait(driver, timeout_elemento).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="txtDataInicio"]'))
                )
                data_inicio.clear()
                data_inicio.send_keys(data_inicial)

                data_fim = WebDriverWait(driver, timeout_elemento).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="txtDataFinal"]'))
                )
                data_fim.clear()
                data_fim.send_keys(data_final)

                bt_ortogante = WebDriverWait(driver, timeout_elemento).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_UpdatePanelListaOutorgantes"]/div/div[2]/div/div/div/button'))
                )
                logging.info("Clicando no botão Outorgante")
                bt_ortogante.click()

                bt_nenhum = WebDriverWait(driver, timeout_elemento).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_UpdatePanelListaOutorgantes"]/div/div[2]/div/div/div/div/div[2]/div/button[2]'))
                )
                logging.info("Clicando no botão Nenhum")
                bt_nenhum.click()

                campo_cnpj = WebDriverWait(driver, timeout_elemento).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_UpdatePanelListaOutorgantes"]/div/div[2]/div/div/div/div/div[1]/input'))
                )
                campo_cnpj.send_keys(cnpj)

                selecionar_cnpj = WebDriverWait(driver, timeout_elemento).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_UpdatePanelListaOutorgantes"]/div/div[2]/div/div/div/div/ul'))
                )
                selecionar_cnpj.click()

                bt_pesquisar = WebDriverWait(driver, timeout_elemento).until(
                    EC.presence_of_element_located((By.XPATH, '//*[@id="ctl00_cphConteudo_btnFiltar"]'))
                )
                bt_pesquisar.click()

                try:
                    bt_visualizar = WebDriverWait(driver, 15).until(
                        EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_cphConteudo_tabelaListagemDctf_GridViewDctfs_ctl02_lbkVisualizarDctf"]'))
                    )
                    logging.info("Clicando no botão Visualizar")
                    bt_visualizar.click()
                except (TimeoutException, NoSuchElementException):
                    logging.info(f"Nenhuma declaração encontrada para CNPJ {cnpj}.")
                    atualizar_status(df, cnpj, 'Nenhuma declaração encontrada')
                    driver.switch_to.default_content()
                    break

                bt_emitir_darf = WebDriverWait(driver, timeout_elemento).until(
                    EC.element_to_be_clickable((By.XPATH, '//*[@id="LinkEmitirDARFIntegral"]'))
                )
                logging.info("Clicando no botão Emitir DARF")
                bt_emitir_darf.click()

                time.sleep(5)
                renomear_arquivo_recente(codigo, competencia, pasta_competencia)

                logging.info(f"Download concluído para {cnpj}")
                atualizar_status(df, cnpj, 'Guia baixada')

                bt_ok = WebDriverWait(driver, timeout_elemento).until(
                    EC.presence_of_element_located((By.XPATH, "//button[text()='OK']"))
                )
                logging.info("Clicando no botão OK")
                bt_ok.click()
                driver.switch_to.default_content()

                sucesso = True
                
            except (TimeoutException, NoSuchElementException) as e:
                logging.error(f"Erro de elemento Selenium no processamento do cliente {cnpj}: {e}")
                
                # Garantir retorno ao contexto principal
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass
                
                atualizar_status(df, cnpj, 'Erro no download')
                tentativas -= 1
                
                if tentativas > 0:
                    logging.info(f"Tentando novamente ({tentativas} tentativas restantes)")
                    time.sleep(3)
                else:
                    logging.error(f"Falha após {tentativas_por_cnpj} tentativas para o cliente {cnpj}")
                    
            except Exception as e:
                logging.error(f"Erro inesperado no processamento do cliente {cnpj}: {e}")
                
                # Garantir retorno ao contexto principal
                try:
                    driver.switch_to.default_content()
                except Exception:
                    pass
                
                atualizar_status(df, cnpj, 'Erro inesperado')
                tentativas = 0
        
        # Salva a planilha ao final do processamento de cada cliente
        try:
            df.to_excel(planilha_save_path, index=False)
        except Exception as e:
            logging.error(f"Erro ao salvar planilha: {e}")
    
    # Reportar conclusão
    if callback:
        callback("Processamento concluído!", total, total)
