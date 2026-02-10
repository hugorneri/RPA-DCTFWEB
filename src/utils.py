import os
import re
import shutil
import time
import logging
from pathlib import Path

try:
    import winreg
except ImportError:
    winreg = None

import pyautogui

# Função de reconhecimento de imagem na tela
def reconhecimento(imagens_referencia, tempo_limite, confidence=1.0):
    tempo_inicio = time.time()
    while time.time() - tempo_inicio < tempo_limite:
        for imagem_referencia in imagens_referencia:
            posicao = pyautogui.locateCenterOnScreen(imagem_referencia, confidence=confidence)
            if posicao is not None:
                logging.info(f"Imagem encontrada: {imagem_referencia}")
                pyautogui.moveTo(posicao.x, posicao.y, duration=0.5)
                return True
            logging.info(f"Imagem não encontrada: {imagem_referencia}")
        tempo_espera = 1.0 + (time.time() % 0.5)
        time.sleep(tempo_espera)
    logging.info("Nenhuma imagem encontrada dentro do tempo limite.")
    return False

# Função de clique em imagem na tela
def clique(imagens_referencia, tempo_limite, confidence=1.0):
    tempo_inicio = time.time()
    while time.time() - tempo_inicio < tempo_limite:
        for imagem_referencia in imagens_referencia:
            posicao = pyautogui.locateCenterOnScreen(imagem_referencia, confidence=confidence)
            if posicao is not None:
                pyautogui.moveTo(posicao.x, posicao.y, duration=0.6)
                time.sleep(0.1 + (time.time() % 0.3))
                pyautogui.click()
                logging.info(f"Clique na imagem: {imagem_referencia}")
                time.sleep(0.2 + (time.time() % 0.3))
                return True
            logging.info(f"Imagem não reconhecida: {imagem_referencia}")
        tempo_espera = 1.0 + (time.time() % 0.8)
        time.sleep(tempo_espera)
    return False

# Função de clique em ocorrência específica de imagem
def clique2(imagens_referencia, tempo_limite, confidence=1.0, ocorrencia=1):
    tempo_inicio = time.time()
    while time.time() - tempo_inicio < tempo_limite:
        for imagem_referencia in imagens_referencia:
            posicoes = list(pyautogui.locateAllOnScreen(imagem_referencia, confidence=confidence))
            logging.info(f"Número de ocorrências da imagem '{imagem_referencia}' encontradas na tela: {len(posicoes)}")
            if len(posicoes) >= ocorrencia:
                posicao_target = posicoes[ocorrencia-1]
                centro_x = posicao_target.left + posicao_target.width / 2
                centro_y = posicao_target.top + posicao_target.height / 2
                pyautogui.moveTo(centro_x, centro_y, duration=0.4 + (time.time() % 0.3))
                time.sleep(0.1 + (time.time() % 0.2))
                pyautogui.click()
                logging.info(f"Clique na imagem: {imagem_referencia} (ocorrência {ocorrencia})")
                time.sleep(0.2 + (time.time() % 0.3))
                return True
        logging.info(f"Imagem não reconhecida: {imagem_referencia}")
        tempo_espera = 1.0 + (time.time() % 0.5)
        time.sleep(tempo_espera)
    return False

# Função para simular lentidão humana
def lentidao():
    time.sleep(1.5)


def get_chrome_version():
    """
    Obtém o número da versão principal do Google Chrome instalado (ex: 144).
    Usado para alinhar o ChromeDriver à versão do Chrome no Windows.

    Returns:
        int ou None: Versão principal (major) do Chrome, ou None se não detectada.
    """
    if winreg is None:
        return None
    keys = [
        (winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\BLBeacon"),
        (winreg.HKEY_LOCAL_MACHINE, r"Software\Google\Chrome\BLBeacon"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Wow6432Node\Google\Chrome\BLBeacon"),
    ]
    for hkey, subkey in keys:
        try:
            key = winreg.OpenKey(hkey, subkey)
            version, _ = winreg.QueryValueEx(key, "version")
            winreg.CloseKey(key)
            if version:
                match = re.match(r"^(\d+)", str(version))
                if match:
                    return int(match.group(1))
        except (OSError, TypeError):
            continue
    return None

def limpar_pasta(pasta):
    """Remove todos os arquivos e subpastas de uma pasta."""
    pasta = str(pasta)
    if not os.path.exists(pasta):
        return
    for item in os.listdir(pasta):
        item_path = os.path.join(pasta, item)
        try:
            if os.path.isfile(item_path) or os.path.islink(item_path):
                os.unlink(item_path)
            elif os.path.isdir(item_path):
                shutil.rmtree(item_path)
        except Exception as e:
            print(f'Erro ao remover {item_path}: {e}')

def renomear_arquivo_recente(codigo, competencia, pasta_competencia):
    """
    Renomeia o arquivo mais recente da pasta para o padrão '<codigo> DARFWEB <competencia>.pdf'.
    
    Args:
        codigo (str): Código do cliente.
        competencia (str): Competência (ex: '06 2025').
        pasta_competencia (str or Path): Caminho da pasta onde está o arquivo.
        
    Returns:
        bool: True se o arquivo foi renomeado com sucesso, False caso contrário.
    """
    try:
        pasta = Path(pasta_competencia)
        
        # Filtrar apenas arquivos (não diretórios) e excluir arquivos já renomeados
        arquivos = [
            f for f in pasta.glob("*") 
            if f.is_file() and "DARFWEB" not in f.name
        ]
        
        if not arquivos:
            logging.warning(f"Nenhum arquivo novo encontrado na pasta {pasta_competencia}")
            return False
        
        arquivo_recente = max(arquivos, key=os.path.getctime)
        novo_nome = pasta / f"{codigo} DARFWEB {competencia}.pdf"
        
        # Verificar se o arquivo de destino já existe
        if novo_nome.exists():
            logging.warning(f"Arquivo {novo_nome} já existe. Removendo...")
            novo_nome.unlink()
        
        arquivo_recente.rename(novo_nome)
        logging.info(f"Arquivo renomeado para: {novo_nome}")
        return True
        
    except ValueError as e:
        logging.error(f"Erro ao encontrar arquivo mais recente: {e}")
        return False
    except Exception as e:
        logging.error(f"Erro ao renomear o arquivo: {e}")
        return False 