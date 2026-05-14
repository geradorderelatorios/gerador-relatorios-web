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

    # Se estiver em Linux (Render) e o caminho tiver formato de Windows, ignora
    if os.name != "nt" and (":" in path or "\\" in path):
        path = DEFAULT_OUTPUT_DIR

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

# ================================
# Função usada pela versão WEB (Flask / Render)
# ================================

def gerar_relatorio_principal(caminho_final: str) -> str:
    """
    Gera um relatório END 'demo' para a versão web e copia
    o arquivo gerado para `caminho_final`, que é o caminho
    esperado pelo Flask (pasta temp_reports no Render).

    Retorna o caminho final do arquivo DOCX.
    """
    import os
    import shutil
    from datetime import date
    from reports_combo import generate_end_combo_report

    # ---------- DADOS "DEMO" ----------
    dados_comuns = {
        "NUMRELATORIO": "WEB-0001",
        "EMPRESA": "RL Metais - Versão Web",
        "ENDEREÇO": "Rua Exemplo, 123",
        "BAIRRO": "Centro",
        "CIDADE": "Americana",
        "ESTADO": "SP",
        "CEP": "00000-000",
        "CONTATO": "Responsável Técnico",
        "DDD": "19",
        "FONE": "0000-0000",
        "EMAIL": "contato@rlmetais.com.br",
        "PECA_INSP": "Peça de Demonstração",
        "NUM_DESENHO": "OP-123456",
        "QUANTIDADE": "1",
        "LOCAL_INSP": "Laboratório RL Metais",
        "DATA_INSP": date.today(),
        # Esses campos são usados nos relatórios específicos
        "LAUDO": "A",
        "LAUDO_EXTENSO": "APROVADO",
    }

    # Não vamos usar fotos na versão web demo
    dados_lp = {
        "FAB_PENETRANTE": "",
        "VAL_PENETRANTE": "",
        "LOTE_PENETRANTE": "",
        "FAB_REVELADOR": "",
        "VAL_REVELADOR": "",
        "LOTE_REVELADOR": "",
        "TEMPERATURA": "",
        "COND_SUPERFICIAL": "",
    }

    dados_pm = {
        "FAB_PARTICULA": "",
        "VAL_PARTICULA": "",
        "LOTE_PARTICULA": "",
        "TEMPERATURA": "",
        "COND_SUPERFICIAL": "",
    }

    dados_us = {
        "MATERIAL": "Aço carbono",
        "COND_SUPERFICIAL": "Usinada",
        "REGIAO_INSP": "Região crítica",
        "ESPESSURA": "10,0 mm",
    }

    # Gera o combo END (CAPA + LP + PM + US) na pasta padrão de saída
    docx_path, _ = generate_end_combo_report(
        dados_comuns=dados_comuns,
        incluir_lp=True,
        incluir_pm=True,
        incluir_us=True,
        dados_lp=dados_lp,
        dados_pm=dados_pm,
        dados_us=dados_us,
        foto_capa=None,  # sem foto na versão web demo
    )

    # Garante que a pasta de destino (TEMP_DIR no app.py) exista
    pasta_destino = os.path.dirname(caminho_final)
    os.makedirs(pasta_destino, exist_ok=True)

    # Copia/renomeia o arquivo gerado para o caminho_final
    shutil.copy2(docx_path, caminho_final)

    return caminho_final

from datetime import date
import os
import shutil
from reports_combo import generate_end_combo_report

def gerar_relatorio_principal(caminho_final: str) -> str:
    """
    Função usada pela API Flask no Render.
    Gera um relatório END simples e copia para `caminho_final`
    """

    # Dados mínimos só pra passar nos templates
    hoje = date.today()
    numrel = hoje.strftime("%Y%m%d")

    dados = {
        "NUMRELATORIO": numrel,
        "DATA_INSP": hoje,
        "PECA_INSP": "Item Teste",
        "NUM_DESENHO": "",
        "QUANTIDADE": "",
        "LOCAL_INSP": "",
        "EMPRESA": "Teste Render",
        "ENDEREÇO": "",
        "BAIRRO": "",
        "CIDADE": "",
        "ESTADO": "",
        "CEP": "",
        "CONTATO": "",
        "DDD": "",
        "FONE": "",
        "EMAIL": "",
    }

    # só com LP para evitar erro
    incluir_lp = True
    incluir_pm = False
    incluir_us = False

    docx_path, _ = generate_end_combo_report(
        dados_comuns=dados,
        incluir_lp=incluir_lp,
        incluir_pm=incluir_pm,
        incluir_us=incluir_us,
        dados_lp=None,
        dados_pm=None,
        dados_us=None,
        foto_capa=None,
    )

    os.makedirs(os.path.dirname(caminho_final), exist_ok=True)
    shutil.copy2(docx_path, caminho_final)

    return caminho_final
