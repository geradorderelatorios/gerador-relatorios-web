import os
from datetime import date
from typing import Optional

from docx import Document
from docxcompose.composer import Composer

from config_relatorios import get_output_dir
from reports_capa import generate_capa_report
from reports_lp import generate_lp_report
from reports_pm import generate_pm_report
from reports_ultrassom import generate_ultrassom_report


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")


def _format_tipo_ensaio_label(us: bool, lp: bool, pm: bool) -> str:
    """
    Monta a string que alimenta {{TIPO_ENSAIO}} na capa,
    conforme os ensaios selecionados.
    """
    labels = []
    if us:
        labels.append("Ultrassom")
    if lp:
        labels.append("Líquido Penetrante")
    if pm:
        labels.append("Partículas Magnéticas")

    if not labels:
        return ""

    if len(labels) == 1:
        return labels[0]
    if len(labels) == 2:
        return f"{labels[0]} e {labels[1]}"
    # 3 ou mais
    return ", ".join(labels[:-1]) + " e " + labels[-1]


def _try_generate_pdf(docx_path: str) -> Optional[str]:
    """
    Tenta gerar PDF a partir do DOCX.

    1) Primeiro tenta com docx2pdf (mais simples).
    2) Se falhar, tenta fallback com automação do Word via win32com.
    3) Se nada der certo, retorna None.
    """
    pdf_path = os.path.splitext(docx_path)[0] + ".pdf"

    # --- Tentativa 1: docx2pdf ---
    try:
        from docx2pdf import convert  # type: ignore
        convert(docx_path, pdf_path)
        if os.path.exists(pdf_path):
            return pdf_path
    except Exception as e:
        print(f"[PDF] Falha com docx2pdf: {e}")

    # --- Tentativa 2: automação do Word via win32com ---
    try:
        import win32com.client  # type: ignore
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(docx_path)
        # 17 = formato PDF
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close(False)
        word.Quit()
        if os.path.exists(pdf_path):
            return pdf_path
    except Exception as e:
        print(f"[PDF] Falha com win32com/Word: {e}")

    # Se chegou aqui, não conseguiu gerar
    return None


def generate_end_combo_report(
    dados_comuns: dict,
    incluir_lp: bool,
    incluir_pm: bool,
    incluir_us: bool,
    dados_lp: Optional[dict] = None,
    dados_pm: Optional[dict] = None,
    dados_us: Optional[dict] = None,
    foto_capa: Optional[str] = None,
    template_paths: Optional[dict[str, str]] = None,
) -> tuple[str, Optional[str]]:
    """
    Gera um único DOCX + (opcionalmente) um PDF contendo:

        CAPA_TEMPLATE
        + LP_TEMPLATE (se incluir_lp)
        + PM_TEMPLATE (se incluir_pm)
        + US_TEMPLATE (se incluir_us)

    na ordem LP, PM, US.

    Parâmetros:
        dados_comuns: campos comuns (NUMRELATORIO, EMPRESA, DATA_INSP, etc.)
        incluir_lp / pm / us: flags informando quais relatórios gerar
        dados_lp / pm / us: dicionários específicos de cada ensaio
                             (se None, assume só os dados comuns)
        foto_capa: caminho da foto de capa (FOTO_1)

    Retorna:
        (caminho_docx_gerado, caminho_pdf_gerado_ou_None)
    """
    if not (incluir_lp or incluir_pm or incluir_us):
        raise ValueError("É necessário selecionar ao menos um tipo de ensaio (LP, PM ou US).")

    template_paths = template_paths or {}
    output_dir = get_output_dir()
    os.makedirs(output_dir, exist_ok=True)

    # pasta temporária para TODOS os docs intermediários (incluindo a CAPA)
    tmp_dir = os.path.join(output_dir, "_tmp_merge")
    os.makedirs(tmp_dir, exist_ok=True)

    # --------- Monta string do tipo de ensaio para a CAPA ---------
    tipo_ensaio_str = _format_tipo_ensaio_label(us=incluir_us, lp=incluir_lp, pm=incluir_pm)

    # NUMRELATORIO é obrigatório para nomear os arquivos
    num_rel = (dados_comuns.get("NUMRELATORIO") or "").strip() or "0000"

    # --------- Dados da capa ---------
    from utils_docx import date_por_extenso

    dados_capa = dict(dados_comuns)  # copia
    dados_capa["TIPO_ENSAIO"] = tipo_ensaio_str
    if foto_capa:
        dados_capa["FOTO_1"] = foto_capa

    data_insp = dados_comuns.get("DATA_INSP")
    if isinstance(data_insp, date):
        dados_capa["DATA_INSP_EXTENSO"] = date_por_extenso(data_insp)

    capa_template = template_paths.get("capa") or os.path.join(TEMPLATES_DIR, "CAPA_TEMPLATE.docx")
    # gera a capa também na pasta temporária
    capa_path = generate_capa_report(dados_capa, capa_template, tmp_dir)

    # --------- Gera os relatórios individuais em pasta temporária ---------
    generated_paths: list[str] = [capa_path]  # capa sempre primeiro

    # LP
    if incluir_lp:
        lp_template = template_paths.get("lp") or os.path.join(TEMPLATES_DIR, "LP_TEMPLATE.docx")
        dados_lp_full = dict(dados_comuns)
        if dados_lp:
            dados_lp_full.update(dados_lp)
        lp_path = generate_lp_report(dados_lp_full, lp_template, tmp_dir)
        generated_paths.append(lp_path)

    # PM
    if incluir_pm:
        pm_template = template_paths.get("pm") or os.path.join(TEMPLATES_DIR, "PM_TEMPLATE.docx")
        dados_pm_full = dict(dados_comuns)
        if dados_pm:
            dados_pm_full.update(dados_pm)
        pm_path = generate_pm_report(dados_pm_full, pm_template, tmp_dir)
        generated_paths.append(pm_path)

    # US
    if incluir_us:
        us_template = template_paths.get("us") or os.path.join(TEMPLATES_DIR, "US_TEMPLATE.docx")
        dados_us_full = dict(dados_comuns)
        if dados_us:
            dados_us_full.update(dados_us)
        us_path = generate_ultrassom_report(
            dados_us_full,
            us_template,
            tmp_dir,
            incluir_capa=False,
        )
        generated_paths.append(us_path)

    # --------- MERGE SEGURO COM DOXCCOMPOSE ---------
    # Usa o primeiro documento como base (a capa)
    base_doc = Document(generated_paths[0])
    composer = Composer(base_doc)

    # Anexa os demais na ordem
    for sub_path in generated_paths[1:]:
        sub_doc = Document(sub_path)
        composer.append(sub_doc)

    final_docx_name = f"Relatório {num_rel}-END.docx"
    final_docx_path = os.path.join(output_dir, final_docx_name)
    composer.save(final_docx_path)

    # O download web usa DOCX; evitar conversao para PDF reduz o tempo de geracao.
    final_pdf_path = None

    # --------- Limpa pasta temporária ---------
    import shutil
    try:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    except Exception:
        pass

    return final_docx_path, final_pdf_path
