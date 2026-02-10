"""
Pacote src - Módulos da automação DCTF.

Módulos:
    - config: Configurações centralizadas
    - gui: Interface gráfica
    - automacao: Lógica de automação Selenium
    - planilha: Manipulação de planilhas Excel
    - utils: Funções utilitárias
"""

from src.config import Config, get_config, save_config
from src.automacao import configurar_driver, login, transmissao
from src.planilha import ler_planilha, atualizar_status
from src.utils import limpar_pasta, renomear_arquivo_recente

__all__ = [
    'Config',
    'get_config', 
    'save_config',
    'configurar_driver',
    'login',
    'transmissao',
    'ler_planilha',
    'atualizar_status',
    'limpar_pasta',
    'renomear_arquivo_recente',
]
