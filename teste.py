import undetected_chromedriver as uc
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

URL = 'https://cav.receita.fazenda.gov.br/autenticacao/login'

def debug_input(msg):
    input(f"[DEBUG] {msg} Pressione ENTER para continuar...")

if __name__ == '__main__':
    options = uc.ChromeOptions()
    driver = uc.Chrome(options=options)
    driver.get(URL)
    driver.maximize_window()
    print('Navegador aberto. Faça login manualmente.')
    debug_input('Faça login manualmente e vá para a tela principal do e-CAC.')

    # Clique no botão Home
    bt_home = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="linkHome"]')))
    print('Botão Home localizado.')
    bt_home.click()
    debug_input('Cliquei no botão Home.')

    # Clique no botão Declarações e Demonstrativos
    bt_declaracoes = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//li[@id="btn214"]')))
    print('Botão Declarações e Demonstrativos localizado.')
    bt_declaracoes.click()
    debug_input('Cliquei no botão Declarações e Demonstrativos.')

    # Clique em Assinar e transmitir DCTF
    bt_assinar = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="containerServicos214"]/div[2]/ul/li[1]/a')))
    print('Botão Assinar e transmitir DCTF localizado.')
    bt_assinar.click()
    debug_input('Cliquei em Assinar e transmitir DCTF.')

    # Troca para o iframe principal
    iframe1 = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmApp"]')))
    driver.switch_to.frame(iframe1)
    print('Entrou no iframe principal.')
    debug_input('Dentro do iframe principal.')

    # Troca para o iframe do captcha (se existir)
    try:
        iframe2 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, '//*[@id="conteudo-pagina"]/div[2]/iframe')))
        driver.switch_to.frame(iframe2)
        print('Entrou no iframe do captcha.')
        debug_input('Dentro do iframe do captcha.')
        bt_sou_humano = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="checkbox"]')))
        print('Botão Sou Humano localizado.')
        bt_sou_humano.click()
        debug_input('Cliquei no botão Sou Humano (captcha).')
        driver.switch_to.default_content()
        driver.switch_to.frame(iframe1)
    except Exception as e:
        print('Captcha não encontrado ou não necessário.')
        driver.switch_to.default_content()
        driver.switch_to.frame(iframe1)

    # Botão Prosseguir
    bt_prosseguir = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_cphConteudo_btnProsseguir"]')))
    print('Botão Prosseguir localizado.')
    bt_prosseguir.click()
    debug_input('Cliquei no botão Prosseguir.')
    driver.switch_to.default_content()

    # Volta para o iframe principal
    iframe = WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, '//*[@id="frmApp"]')))
    driver.switch_to.frame(iframe)
    print('Entrou novamente no iframe principal.')
    debug_input('Dentro do iframe principal novamente.')

    # Botão Sou Procurador
    bt_sou_procurador = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="ctl00_cphConteudo_chkListarOutorgantes"]')))
    print('Botão Sou Procurador localizado.')
    bt_sou_procurador.click()
    debug_input('Cliquei no botão Sou Procurador.')

    print('Fim do fluxo de debug!')
    driver.quit() 