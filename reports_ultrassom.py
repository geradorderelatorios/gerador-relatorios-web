import os
from datetime import date

from docx import Document
from docxcompose.composer import Composer

from utils_docx import (
    replace_placeholders,
    inserir_foto_por_placeholder,
    format_date_br,
    date_por_extenso,
)

CAPA_TEMPLATE_NAME = "CAPA_TEMPLATE.docx"


def _montar_mapping(dados: dict) -> dict:
    """
    Monta o dicionário de placeholders -> valores em string,
    incluindo tratamento de datas e conclusão automática.
    """
    dados_limpos = dict(dados)

    # Trata data de inspeção
    data_insp = dados_limpos.get("DATA_INSP")
    if isinstance(data_insp, date):
        dados_limpos["DATA_INSP"] = format_date_br(data_insp)
        dados_limpos["DATA_INSP_EXTENSO"] = date_por_extenso(data_insp)
    elif data_insp:
        dados_limpos["DATA_INSP"] = str(data_insp)
        dados_limpos["DATA_INSP_EXTENSO"] = ""
    else:
        dados_limpos["DATA_INSP"] = ""
        dados_limpos["DATA_INSP_EXTENSO"] = ""

    # Conclusão automática
    laudo = dados_limpos.get("LAUDO")
    laudo_ext = dados_limpos.get("LAUDO_EXTENSO") or ""

    if laudo == "A":
        dados_limpos["TIPO_DESC"] = "Isento"
        dados_limpos["LOC_DESC"] = "Isento"
        dados_limpos["DIM_DESC"] = "Isento"
    else:
        dados_limpos.setdefault("TIPO_DESC", dados_limpos.get("US_TIPO_DESC", ""))
        dados_limpos.setdefault("LOC_DESC", dados_limpos.get("US_LOC_DESC", ""))
        dados_limpos.setdefault("DIM_DESC", dados_limpos.get("US_DIM_DESC", ""))
    dados_limpos["US_TIPO_DESC"] = dados_limpos["TIPO_DESC"]
    dados_limpos["US_LOC_DESC"] = dados_limpos["LOC_DESC"]
    dados_limpos["US_DIM_DESC"] = dados_limpos["DIM_DESC"]
    dados_limpos["US_LAUDO"] = dados_limpos.get("LAUDO", "")
    dados_limpos["US_LAUDO_EXTENSO"] = dados_limpos.get("LAUDO_EXTENSO", "")

    if laudo == "A":
        conclusao = (
            "CONCLUSÃO: Não Foram evidenciadas indicações relevantes, "
            f"estando o ensaio {laudo_ext} segundo o critério de aceitação da norma acima"
        )
    else:
        conclusao = (
            "CONCLUSÃO: Foram evidenciadas indicações relevantes, "
            f"estando o ensaio {laudo_ext} segundo o critério de aceitação da norma acima"
        )

    dados_limpos["CONCLUSAO"] = conclusao

    mapping: dict[str, str] = {}
    for k, v in dados_limpos.items():
        # Fotos não entram como texto
        if k.startswith("FOTO_"):
            continue
        placeholder = "{{" + k + "}}"
        mapping[placeholder] = "" if v is None else str(v)

    mapping["{{CONCLUSAO}}"] = conclusao
    return mapping


def generate_ultrassom_report(
    dados: dict,
    template_path: str,
    output_dir: str,
    incluir_capa: bool = True,
) -> str:
    """
    Gera o relatório de Ultrassom (US) com:
    - CAPA (CAPA_TEMPLATE.docx)
    - LAUDO (US_TEMPLATE.docx passado em template_path)
    em um único arquivo .docx.
    Se a capa não existir, gera apenas o laudo.
    """
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template não encontrado: {template_path}")

    os.makedirs(output_dir, exist_ok=True)

    numrel = str(dados.get("NUMRELATORIO", "")).strip()
    if not numrel:
        raise ValueError("O campo NUMRELATORIO é obrigatório.")

    mapping = _montar_mapping(dados)

    # Diretório dos templates (onde está o US_TEMPLATE.docx)
    base_dir = os.path.dirname(template_path)
    capa_path = os.path.join(base_dir, CAPA_TEMPLATE_NAME)

    doc_capa = None
    if incluir_capa and os.path.exists(capa_path):
        # CAPA usa placeholders próprios, inclusive {{FOTO_1}}
        doc_capa = Document(capa_path)
        replace_placeholders(doc_capa, mapping)

        foto1 = dados.get("FOTO_1")
        if foto1:
            inserir_foto_por_placeholder(doc_capa, "{{FOTO_1}}", foto1, height_cm=10.0)

    # LAUDO US
    doc_laudo = Document(template_path)
    replace_placeholders(doc_laudo, mapping)

    foto2 = dados.get("FOTO_2")
    if foto2:
        inserir_foto_por_placeholder(doc_laudo, "{{FOTO_2}}", foto2, height_cm=4.0)

    foto3 = dados.get("FOTO_3")
    if foto3:
        inserir_foto_por_placeholder(doc_laudo, "{{FOTO_3}}", foto3, height_cm=4.0)

    file_name = f"Relatório {numrel}-US.docx"
    output_path = os.path.join(output_dir, file_name)

    # Junta capa + laudo, se a capa existir
    if doc_capa is not None:
        composer = Composer(doc_capa)
        composer.append(doc_laudo)
        composer.save(output_path)
    else:
        doc_laudo.save(output_path)

    return output_path
