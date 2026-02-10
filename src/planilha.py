import pandas as pd
from pathlib import Path
import logging


def ler_planilha(planilha_path):
    """
    Lê a planilha de clientes e retorna listas de CNPJs, códigos e o DataFrame.
    
    Args:
        planilha_path (str or Path): Caminho da planilha Excel.
        
    Returns:
        tuple: (lista de CNPJs, lista de códigos, DataFrame)
    """
    df = pd.read_excel(planilha_path)
    
    # Garantir que CNPJ seja string para comparações consistentes
    df['CNPJ'] = df['CNPJ'].astype(str).str.strip()
    
    # Preservar STATUS existente - só cria coluna se não existir
    if 'STATUS' not in df.columns:
        df['STATUS'] = ''
    else:
        # Preencher valores NaN com string vazia
        df['STATUS'] = df['STATUS'].fillna('')
    
    cnpjs = df['CNPJ'].tolist()
    codigos = df['COD'].astype(str).tolist()
    
    logging.info(f"Planilha carregada: {len(cnpjs)} CNPJs encontrados")
    return cnpjs, codigos, df


def atualizar_status(df, cnpj, status):
    """
    Atualiza o status de um cliente no DataFrame.
    
    Args:
        df (pd.DataFrame): DataFrame da planilha.
        cnpj (str): CNPJ do cliente.
        status (str): Novo status a ser atribuído.
    """
    # Garantir que cnpj seja string para comparação consistente
    cnpj_str = str(cnpj).strip()
    mask = df['CNPJ'] == cnpj_str
    
    if mask.any():
        df.loc[mask, 'STATUS'] = status
        logging.info(f"Status atualizado para CNPJ {cnpj_str}: {status}")
    else:
        logging.warning(f"CNPJ {cnpj_str} não encontrado na planilha") 