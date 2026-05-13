import os
import json

# Pasta base do projeto (onde está este arquivo)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "config_relatorios.json")

# Pasta padrão (se o usuário não configurar outra)
DEFAULT_OUTPUT_DIR = os.environ.get("RL_METAIS_OUTPUT_DIR") or os.path.join(BASE_DIR, "output")


def _carregar_config() -> dict:
    """Lê o arquivo de configuração (se existir) e devolve como dict."""
    if not os.path.exists(CONFIG_FILE):
        return {}
    try:
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        # Se der erro ao ler (arquivo corrompido, etc.), ignora e volta vazio
        return {}


def get_output_dir() -> str:
    """
    Retorna a pasta padrão onde os relatórios serão salvos.
    Se não houver configuração, usa a pasta 'output' na raiz do projeto.
    Garante que a pasta exista.
    """
    cfg = _carregar_config()
    path = cfg.get("output_dir") or DEFAULT_OUTPUT_DIR

    # Se for caminho relativo, converte para absoluto em relação à BASE_DIR
    if not os.path.isabs(path):
        path = os.path.abspath(os.path.join(BASE_DIR, path))

    os.makedirs(path, exist_ok=True)
    return path


def set_output_dir(path: str) -> None:
    """
    Define a pasta padrão para salvar os relatórios e grava no arquivo de config.
    """
    if not path:
        return

    # Normaliza o caminho
    path = os.path.abspath(path)

    cfg = _carregar_config()
    cfg["output_dir"] = path

    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)
