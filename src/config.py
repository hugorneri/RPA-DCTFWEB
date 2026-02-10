"""
Módulo de configuração centralizada para a automação DCTF.
Permite persistência das configurações em arquivo JSON.
"""
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


def get_project_root() -> Path:
    """Retorna o diretório raiz do projeto (pai da pasta src)."""
    return Path(__file__).parent.parent


# Caminho padrão para o arquivo de configuração (na raiz do projeto)
CONFIG_FILE = get_project_root() / "config.json"


@dataclass
class Config:
    """Configurações da automação DCTF."""
    
    # Datas e competência
    data_inicial: str = '01062025'
    data_final: str = '30062025'
    competencia: str = '06 2025'
    
    # Timeouts e tentativas
    timeout_elemento: int = 30
    tentativas_por_cnpj: int = 3
    tentativas_gerais: int = 3
    
    # Caminho da planilha (pode ser personalizado)
    planilha_path: str = ''
    
    # Paths (não serializados no JSON, calculados dinamicamente)
    _pasta_base: Optional[Path] = field(default=None, repr=False)
    
    def __post_init__(self):
        """Inicializa paths após criação do objeto."""
        if self._pasta_base is None:
            self._pasta_base = get_project_root()
    
    @property
    def pasta_base(self) -> Path:
        """Retorna o diretório base do projeto."""
        return self._pasta_base or get_project_root()
    
    @property
    def pasta_competencia(self) -> Path:
        """Retorna o caminho da pasta de competências executadas."""
        return self.pasta_base / "Competencias executadas"
    
    @property
    def pasta_download(self) -> Path:
        """Retorna o caminho da pasta de download para a competência atual."""
        return self.pasta_competencia / self.competencia
    
    @property
    def imagem_dir(self) -> Path:
        """Retorna o diretório de imagens."""
        return self.pasta_base / "img"
    
    @property
    def planilha(self) -> Path:
        """Retorna o caminho da planilha de dados."""
        if self.planilha_path and Path(self.planilha_path).exists():
            return Path(self.planilha_path)
        return self.pasta_base / "database.xlsx"
    
    @property
    def cache(self) -> Path:
        """Retorna o caminho do cache do perfil do Chrome."""
        return self.pasta_base / "perfil-path"
    
    @property
    def log_file(self) -> Path:
        """Retorna o caminho do arquivo de log."""
        return self.pasta_base / "AUTOMACAO-DCTF.log"
    
    def to_dict(self) -> dict:
        """Converte configurações serializáveis para dicionário."""
        return {
            'data_inicial': self.data_inicial,
            'data_final': self.data_final,
            'competencia': self.competencia,
            'timeout_elemento': self.timeout_elemento,
            'tentativas_por_cnpj': self.tentativas_por_cnpj,
            'tentativas_gerais': self.tentativas_gerais,
            'planilha_path': self.planilha_path,
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'Config':
        """Cria uma instância de Config a partir de um dicionário."""
        return cls(
            data_inicial=data.get('data_inicial', '01062025'),
            data_final=data.get('data_final', '30062025'),
            competencia=data.get('competencia', '06 2025'),
            timeout_elemento=data.get('timeout_elemento', 30),
            tentativas_por_cnpj=data.get('tentativas_por_cnpj', 3),
            tentativas_gerais=data.get('tentativas_gerais', 3),
            planilha_path=data.get('planilha_path', ''),
        )
    
    def save(self, filepath: Optional[Path] = None) -> None:
        """
        Salva as configurações em um arquivo JSON.
        
        Args:
            filepath: Caminho do arquivo. Se None, usa o padrão.
        """
        filepath = filepath or CONFIG_FILE
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(self.to_dict(), f, indent=4, ensure_ascii=False)
    
    @classmethod
    def load(cls, filepath: Optional[Path] = None) -> 'Config':
        """
        Carrega as configurações de um arquivo JSON.
        
        Args:
            filepath: Caminho do arquivo. Se None, usa o padrão.
            
        Returns:
            Instância de Config com as configurações carregadas.
        """
        filepath = filepath or CONFIG_FILE
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return cls.from_dict(data)
        except FileNotFoundError:
            # Retorna configuração padrão se arquivo não existe
            return cls()
        except json.JSONDecodeError:
            # Retorna configuração padrão se arquivo está corrompido
            return cls()


def get_config() -> Config:
    """
    Obtém a configuração atual, carregando do arquivo se existir.
    
    Returns:
        Instância de Config.
    """
    return Config.load()


def save_config(config: Config) -> None:
    """
    Salva a configuração no arquivo padrão.
    
    Args:
        config: Instância de Config a ser salva.
    """
    config.save()
