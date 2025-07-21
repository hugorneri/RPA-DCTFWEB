import logging
import os
from pathlib import Path
from src.automacao import configurar_driver, login, transmissao
from src.planilha import ler_planilha
from src.utils import limpar_pasta

data_inicial = '01062025'
data_final = '30062025'
competencia = '06 2025'

PASTA_COMPETENCIA = Path(__file__).parent / "Competencias executadas"
PASTA_DOWNLOAD = PASTA_COMPETENCIA / f"{competencia}"
IMAGEM_DIR = Path(__file__).parent / "img"
PLANILHA = Path(__file__).parent / "database.xlsx"
CACHE = Path(__file__).parent / 'perfil-path'

logging.basicConfig(filename='AUTOMACAO-DCTF.log', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def main():
    tentativas_gerais = 3
    pasta_competencia = os.path.join(PASTA_COMPETENCIA, competencia)
    if not os.path.exists(pasta_competencia):
        os.mkdir(pasta_competencia)
    if not CACHE.exists():
        CACHE.mkdir()
    # Limpar a pasta de perfil antes de iniciar
    limpar_pasta(CACHE)
    while tentativas_gerais > 0:
        try:
            logging.info("Iniciando automação DCTF")
            print("="*50)
            print("AUTOMAÇÃO DCTF - INICIANDO")
            print("="*50)
            driver = configurar_driver(CACHE, pasta_competencia)
            cnpjs, codigos, df = ler_planilha(PLANILHA)
            print("Aguardando login manual...")
            login(driver)
            input("Login concluído? Pressione ENTER para continuar com a navegação e processamento...")
            transmissao(cnpjs, codigos, df, driver, competencia, pasta_competencia, data_inicial, data_final)
            df.to_excel('database.xlsx', index=False)
            print("="*50)
            print("AUTOMAÇÃO CONCLUÍDA COM SUCESSO!")
            print("="*50)
            logging.info("Automação concluída com sucesso")
            try:
                driver.quit()
            except:
                pass
            break
        except Exception as e:
            tentativas_gerais -= 1
            logging.error(f"Erro geral na execução: {e}")
            print(f"Ocorreu um erro: {e}")
            if tentativas_gerais > 0:
                print(f"Tentando novamente. Restam {tentativas_gerais} tentativas.")
                try:
                    driver.quit()
                except:
                    pass
                tempo_espera = 5
                print(f"Aguardando {tempo_espera} segundos antes da próxima tentativa...")
                import time
                time.sleep(tempo_espera)
            else:
                print("Número máximo de tentativas excedido. Encerrando programa.")
                logging.error("Número máximo de tentativas excedido. Programa finalizado com erro.")
    try:
        df.to_excel('database.xlsx', index=False)
        print("Planilha salva com status final dos processamentos.")
    except:
        print("Não foi possível salvar a planilha final.")

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        print("\nPrograma interrompido pelo usuário.")
        logging.info("Programa interrompido pelo usuário (KeyboardInterrupt)")
    except Exception as e:
        print(f"\nErro não tratado: {e}")
        logging.critical(f"Erro crítico não tratado: {e}")
    finally:
        print("\nFinalizando...")
        logging.info("Programa finalizado") 
        # Limpar a pasta de perfil após a execução
        from src.utils import limpar_pasta
        limpar_pasta(CACHE) 