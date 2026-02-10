"""
Ponto de entrada principal para a automação DCTF.

Uso:
    python main.py          # Abre a interface gráfica (padrão)
    python main.py --cli    # Executa no modo linha de comando
    python main.py --help   # Mostra ajuda
"""
import sys
import logging
import time
from pathlib import Path

# Adicionar o diretório do projeto ao path
PROJECT_ROOT = Path(__file__).parent
sys.path.insert(0, str(PROJECT_ROOT))

from src.config import Config, get_config
from src.automacao import configurar_driver, login, transmissao
from src.planilha import ler_planilha


def setup_logging(config: Config):
    """Configura o sistema de logging."""
    logging.basicConfig(
        filename=str(config.log_file),
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )


def run_cli(config: Config = None):
    """
    Executa a automação no modo CLI (linha de comando).
    
    Args:
        config: Configuração a ser usada. Se None, carrega do arquivo.
    """
    # Carregar configuração
    if config is None:
        config = get_config()
    
    setup_logging(config)
    
    tentativas_gerais = config.tentativas_gerais
    pasta_competencia = config.pasta_download
    
    # Criar diretórios necessários
    if not pasta_competencia.exists():
        pasta_competencia.mkdir(parents=True)
    
    driver = None
    df = None
    
    while tentativas_gerais > 0:
        try:
            logging.info("Iniciando automação DCTF")
            print("=" * 50)
            print("AUTOMAÇÃO DCTF - MODO CLI")
            print("=" * 50)
            print(f"Competência: {config.competencia}")
            print(f"Período: {config.data_inicial} a {config.data_final}")
            print("=" * 50)
            
            driver = configurar_driver(pasta_competencia)
            cnpjs, codigos, df = ler_planilha(config.planilha)
            
            print("Aguardando login manual...")
            login(driver)
            input("Login concluído? Pressione ENTER para continuar com a navegação e processamento...")
            
            transmissao(
                cnpjs=cnpjs,
                codigos=codigos,
                df=df,
                driver=driver,
                competencia=config.competencia,
                pasta_competencia=pasta_competencia,
                data_inicial=config.data_inicial,
                data_final=config.data_final,
                timeout_elemento=config.timeout_elemento,
                tentativas_por_cnpj=config.tentativas_por_cnpj,
                planilha_path=str(config.planilha)
            )
            
            # Salvar planilha final
            df.to_excel(config.planilha, index=False)
            
            print("=" * 50)
            print("AUTOMAÇÃO CONCLUÍDA COM SUCESSO!")
            print("=" * 50)
            logging.info("Automação concluída com sucesso")
            
            try:
                driver.quit()
            except Exception:
                pass
            
            break
            
        except Exception as e:
            tentativas_gerais -= 1
            logging.error(f"Erro geral na execução: {e}")
            print(f"Ocorreu um erro: {e}")
            
            if tentativas_gerais > 0:
                print(f"Tentando novamente. Restam {tentativas_gerais} tentativas.")
                try:
                    if driver:
                        driver.quit()
                except Exception:
                    pass
                
                tempo_espera = 5
                print(f"Aguardando {tempo_espera} segundos antes da próxima tentativa...")
                time.sleep(tempo_espera)
            else:
                print("Número máximo de tentativas excedido. Encerrando programa.")
                logging.error("Número máximo de tentativas excedido. Programa finalizado com erro.")
    
    # Salvar planilha com status final
    if df is not None:
        try:
            df.to_excel(config.planilha, index=False)
            print("Planilha salva com status final dos processamentos.")
        except Exception as e:
            print(f"Não foi possível salvar a planilha final: {e}")


def run_gui():
    """Executa a automação com interface gráfica."""
    from src.gui import run_gui as start_gui
    start_gui()


def show_help():
    """Mostra a ajuda do programa."""
    print("""
Automação DCTF - Sistema de transmissão automática de declarações

Uso:
    python main.py              Abre a interface gráfica (padrão)
    python main.py --gui        Abre a interface gráfica
    python main.py --cli        Executa no modo linha de comando
    python main.py --help       Mostra esta ajuda

Configurações:
    As configurações são salvas em config.json na raiz do projeto.
    Você pode editar manualmente ou usar a interface gráfica.

Arquivos:
    - database.xlsx         Planilha com CNPJs e códigos
    - config.json           Arquivo de configurações
    - AUTOMACAO-DCTF.log    Log de execução
    - Competencias executadas/  Pasta com os DARFs baixados
""")


def main():
    """Função principal - ponto de entrada do programa."""
    # Processar argumentos de linha de comando
    args = sys.argv[1:]
    
    if '--help' in args or '-h' in args:
        show_help()
        return
    
    if '--cli' in args:
        # Modo CLI
        config = get_config()
        try:
            run_cli(config)
        except KeyboardInterrupt:
            print("\nPrograma interrompido pelo usuário.")
            logging.info("Programa interrompido pelo usuário (KeyboardInterrupt)")
        except Exception as e:
            print(f"\nErro não tratado: {e}")
            logging.critical(f"Erro crítico não tratado: {e}")
        finally:
            print("\nFinalizando...")
            logging.info("Programa finalizado")
    else:
        # Modo GUI (padrão)
        run_gui()


if __name__ == '__main__':
    main()
