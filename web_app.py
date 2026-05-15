import json
import os
import re
import urllib.error
import urllib.request
import zipfile
from datetime import date, datetime

from flask import Flask, flash, g, jsonify, redirect, render_template, request, send_file, session as browser_session, url_for
from jinja2 import ChoiceLoader, DictLoader, FileSystemLoader
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

from config_relatorios import get_output_dir
from database import get_session, init_db
from models import Cliente, EntradaRelatorio, Insumo, Organization, TemplateEmpresa, TipoRelatorio, User
from reports_lp import generate_lp_report
from reports_pm import generate_pm_report
from reports_combo import generate_end_combo_report
from reports_ultrassom import generate_ultrassom_report


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
UPLOAD_DIR = os.environ.get("RL_METAIS_UPLOAD_DIR") or os.path.join(BASE_DIR, "web_uploads")
ORG_TEMPLATES_DIR = os.environ.get("RL_METAIS_TEMPLATE_DIR") or os.path.join(BASE_DIR, "organization_templates")

REPORT_TYPES = {
    "lp": {
        "name": "Líquido Penetrante - LP",
        "db_name": "Líquido Penetrante - LP",
        "template": "LP_TEMPLATE.docx",
        "generator": generate_lp_report,
        "suffix": "LP",
        "fields": [
            ("PECA_INSP", "Peça inspecionada", "text"),
            ("NUM_DESENHO", "Número do desenho / OP", "text"),
            ("QUANTIDADE", "Quantidade", "text"),
            ("LOCAL_INSP", "Local da inspeção", "text"),
            ("TEMPERATURA", "Temperatura", "text"),
            ("COND_SUPERFICIAL", "Condição da superfície", "text"),
        ],
        "photos": ["FOTO_1", "FOTO_2", "FOTO_3"],
    },
    "pm": {
        "name": "Partículas Magnéticas - PM",
        "db_name": "Partículas Magnéticas - PM",
        "template": "PM_TEMPLATE.docx",
        "generator": generate_pm_report,
        "suffix": "PM",
        "fields": [
            ("PECA_INSP", "Peça inspecionada", "text"),
            ("NUM_DESENHO", "Número do desenho / OP", "text"),
            ("QUANTIDADE", "Quantidade", "text"),
            ("LOCAL_INSP", "Local da inspeção", "text"),
            ("FAB_PARTICULA", "Partícula - fabricação", "text"),
            ("VAL_PARTICULA", "Partícula - validade", "text"),
            ("LOTE_PARTICULA", "Partícula - lote", "text"),
            ("TEMPERATURA", "Temperatura", "text"),
            ("COND_SUPERFICIAL", "Condição da superfície", "text"),
        ],
        "photos": ["FOTO_1", "FOTO_2"],
    },
    "us": {
        "name": "Ultrassom - US",
        "db_name": "Ultrassom - US",
        "template": "US_TEMPLATE.docx",
        "generator": generate_ultrassom_report,
        "suffix": "US",
        "fields": [
            ("PECA_INSP", "Peça ensaiada", "text"),
            ("NUM_DESENHO", "Número da OP", "text"),
            ("QUANTIDADE", "Quantidade", "text"),
            ("LOCAL_INSP", "Local do ensaio", "text"),
            ("MATERIAL", "Material", "text"),
            ("COND_SUPERFICIAL", "Condição da superfície", "text"),
            ("REGIAO_INSP", "Região inspecionada", "text"),
            ("ESPESSURA", "Espessura", "text"),
        ],
        "photos": ["FOTO_1", "FOTO_2", "FOTO_3"],
    },
}


app = Flask(__name__)
app.secret_key = os.environ.get("RL_METAIS_SECRET", "dev-secret-change-me")

FALLBACK_TEMPLATES = {
    "base.html": """
<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>RL Metais Relatórios</title>
  <style>
    :root{--bg:#f3f5f8;--surface:#fff;--text:#111;--muted:#66717f;--line:#d8dde5;--primary:#0057b8;--primary-dark:#003f87;font-family:Arial,Helvetica,sans-serif}
    *{box-sizing:border-box}body{margin:0;background:radial-gradient(circle at top left,rgba(0,87,184,.1),transparent 34rem),linear-gradient(180deg,#f8fafc 0%,var(--bg) 42%);color:var(--text)}
    a{color:var(--primary);text-decoration:none}.topbar{align-items:center;background:rgba(255,255,255,.94);border-bottom:1px solid var(--line);display:flex;justify-content:space-between;min-height:72px;padding:0 32px;position:sticky;top:0;z-index:10}
    .brand{font-weight:900;color:#111}.topbar nav{align-items:center;display:flex;gap:22px}.topbar nav a,.topbar nav span{color:#111;font-size:14px;font-weight:700}.nav-button{background:transparent;border:0;color:#111;font-weight:700;padding:0}
    .page{margin:0 auto;max-width:1180px;padding:44px 24px 64px}.panel,.tile{background:var(--surface);border:1px solid var(--line);border-radius:14px;padding:24px;box-shadow:0 16px 40px rgba(15,23,42,.05)}
    .auth{margin:40px auto 0;max-width:460px}.form{display:grid;gap:16px}.grid{display:grid;gap:16px}.grid.two{grid-template-columns:repeat(2,minmax(0,1fr))}.grid.three{grid-template-columns:repeat(3,minmax(0,1fr))}
    .split{display:grid;gap:22px;grid-template-columns:minmax(0,1.15fr) minmax(320px,.85fr)}label{color:#334155;display:grid;font-size:14px;font-weight:700;gap:6px}
    input,select{border:1px solid var(--line);border-radius:8px;font:inherit;min-height:42px;padding:10px 12px;width:100%}input:focus,select:focus{border-color:var(--primary);outline:3px solid rgba(0,87,184,.12)}
    .button,button{background:#fff;border:1px solid var(--line);border-radius:8px;color:var(--text);cursor:pointer;display:inline-flex;font:inherit;font-weight:800;justify-content:center;min-height:44px;padding:12px 16px}
    .button.primary,button.primary{background:var(--primary);border-color:var(--primary);color:white}.button.primary:hover,button.primary:hover{background:var(--primary-dark)}.button.danger,button.danger{background:#b42318;border-color:#b42318;color:#fff}
    .eyebrow{color:var(--primary);font-size:13px;font-weight:800;text-transform:uppercase;margin:0 0 8px}h1{margin:0 0 14px}.actions,.home-actions{display:flex;gap:12px;flex-wrap:wrap}
    .home-hero{align-items:stretch;background:linear-gradient(135deg,rgba(255,255,255,.96),rgba(244,248,253,.96));border:1px solid var(--line);border-radius:18px;box-shadow:0 24px 70px rgba(15,23,42,.08);display:grid;gap:32px;grid-template-columns:minmax(0,1.2fr) minmax(320px,.8fr);padding:44px}
    .home-copy{display:grid;gap:16px}.home-copy h1{font-size:46px;line-height:1.05}.home-copy p,.muted{color:var(--muted);line-height:1.6}.hero-panel{background:#0f172a;border-radius:14px;color:white;display:grid;gap:18px;padding:28px}.flow-line{display:grid;gap:10px;grid-template-columns:1fr repeat(3,auto)}.flow-line strong,.flow-line span{background:rgba(255,255,255,.1);border-radius:8px;padding:12px;text-align:center}.flow-line strong{background:var(--primary)}
    .quick-grid{display:grid;gap:16px;grid-template-columns:repeat(3,minmax(0,1fr));margin-top:24px}.quick-card{background:#fff;border:1px solid var(--line);border-radius:14px;color:#111;display:grid;gap:10px;min-height:145px;padding:24px}
    .section-title{align-items:center;display:flex;justify-content:space-between;gap:16px}.list,.client-list{display:grid;gap:8px}.client-list{max-height:calc(100vh - 210px);overflow-y:auto;padding-right:6px}.client-item{border:1px solid transparent;border-radius:8px;color:#111;display:grid;gap:4px;padding:11px 12px}.client-item:hover,.client-item.active{background:#f1f6ff;border-color:#c9dcf8}.client-item span,.client-item small,.empty{color:var(--muted);font-size:13px}
    .check-option{align-items:center;background:#f7f9fc;border:1px solid var(--line);border-radius:8px;display:flex;gap:10px;min-height:56px;padding:12px}.check-option input{min-height:auto;width:auto}.subsection{border-top:1px solid var(--line);display:grid;gap:14px;padding-top:18px}.form-note{background:#f1f5fb;border-left:4px solid var(--primary);color:var(--muted);margin:0;padding:12px 14px}
    .alert{border-radius:8px;padding:12px 14px;margin-bottom:16px}.alert.error{background:#fff0f0;color:#9f1d1d}
    @media(max-width:780px){.grid.two,.grid.three,.split,.home-hero,.quick-grid{grid-template-columns:1fr}.home-hero{padding:28px}.home-copy h1{font-size:34px}.topbar{align-items:flex-start;flex-direction:column;padding:18px 24px}.topbar nav{flex-wrap:wrap}}
  </style>
</head>
<body>
  <header class="topbar">
    <a class="brand" href="{{ url_for('index') }}">RL Metais</a>
    <nav>
      {% if g.user_id %}
        <span>{{ g.organization_name }}</span>
        <a href="{{ url_for('clientes') }}">Clientes</a>
        <a href="{{ url_for('insumos') }}">Insumos</a>
        <a href="{{ url_for('emitir_relatorio') }}">Gerar Relatórios</a>
        <form method="post" action="{{ url_for('logout') }}"><button class="nav-button" type="submit">Sair</button></form>
      {% endif %}
    </nav>
  </header>
  <main class="page">
    {% with messages = get_flashed_messages(with_categories=true) %}
      {% for category, message in messages if category != 'success' %}<div class="alert {{ category }}">{{ message }}</div>{% endfor %}
    {% endwith %}
    {% block content %}{% endblock %}
  </main>
</body>
</html>
""",
    "setup.html": """{% extends "base.html" %}{% block content %}<section class="auth panel"><form class="form" method="post"><div><p class="eyebrow">Primeiro acesso</p><h1>Criar empresa e usuário administrador</h1><p class="muted">Configure a conta principal para acessar o sistema de relatórios.</p></div><label>Empresa<input name="organization_name" required autofocus></label><label>Seu nome<input name="name" required></label><label>E-mail<input name="email" type="email" required></label><label>Senha<input name="password" type="password" minlength="6" required></label><button class="primary" type="submit">Criar acesso</button></form></section>{% endblock %}""",
    "login.html": """{% extends "base.html" %}{% block content %}<section class="auth panel"><form class="form" method="post"><div><p class="eyebrow">Acesso</p><h1>Entrar</h1></div><label>E-mail<input name="email" type="email" required autofocus></label><label>Senha<input name="password" type="password" required></label><button class="primary" type="submit">Entrar</button></form></section>{% endblock %}""",
    "index.html": """{% extends "base.html" %}{% block content %}<section class="home-hero"><div class="home-copy"><p class="eyebrow">Sistema de relatórios técnicos</p><h1>Emissão profissional de laudos RL Metais</h1><p>Organize clientes, insumos e relatórios de ensaios em um fluxo único, com capa padronizada e documentos prontos para entrega.</p><div class="home-actions"><a class="button primary" href="{{ url_for('emitir_relatorio') }}">Gerar Relatórios</a><a class="button" href="{{ url_for('clientes') }}">Gerenciar Clientes</a></div></div><aside class="hero-panel"><span class="eyebrow">Fluxo de emissão</span><div class="flow-line"><strong>Capa</strong><span>LP</span><span>PM</span><span>US</span></div><p>Gere laudos separados ou combinados em um único arquivo, sempre com a capa oficial.</p></aside></section><section class="quick-grid"><a class="quick-card" href="{{ url_for('clientes') }}"><span>01</span><strong>Clientes</strong><small>Cadastre, consulte por CNPJ, altere ou remova empresas.</small></a><a class="quick-card" href="{{ url_for('insumos') }}"><span>02</span><strong>Insumos</strong><small>Controle lotes, fabricação e validade dos materiais usados.</small></a><a class="quick-card" href="{{ url_for('emitir_relatorio') }}"><span>03</span><strong>Gerar Relatórios</strong><small>Selecione os ensaios e emita documentos completos.</small></a></section>{% endblock %}""",
    "clientes.html": """{% extends "base.html" %}{% block content %}<section class="split"><form class="panel form" method="post" action="{{ url_for('salvar_cliente') }}"><div class="section-title"><h1>{{ "Alterar cliente" if selected else "Novo cliente" }}</h1>{% if selected %}<a href="{{ url_for('clientes') }}">Novo cadastro</a>{% endif %}</div><input type="hidden" name="cliente_id" value="{{ selected.id if selected else '' }}"><div class="grid two"><label>CNPJ<input name="cnpj" id="cnpj" value="{{ selected.cnpj if selected else '' }}"></label><label>&nbsp;<button class="button" type="button" id="buscar-cnpj">Buscar CNPJ</button></label></div><label>Razão social *<input name="razao_social" id="razao_social" required value="{{ selected.razao_social if selected else '' }}"></label><label>Contato<input name="contato" id="contato" value="{{ selected.contato if selected else '' }}"></label><label>Inscrição estadual<input name="ie" id="ie" value="{{ selected.ie if selected else '' }}"></label><div class="grid two"><label>Rua<input name="rua" id="rua" value="{{ selected.rua if selected else '' }}"></label><label>Número<input name="numero" id="numero" value="{{ selected.numero if selected else '' }}"></label></div><div class="grid three"><label>Bairro<input name="bairro" id="bairro" value="{{ selected.bairro if selected else '' }}"></label><label>Cidade<input name="cidade" id="cidade" value="{{ selected.cidade if selected else '' }}"></label><label>UF<input name="uf" id="uf" maxlength="2" value="{{ selected.uf if selected else '' }}"></label></div><div class="grid three"><label>CEP<input name="cep" id="cep" value="{{ selected.cep if selected else '' }}"></label><label>DDD<input name="ddd" id="ddd" value="{{ selected.ddd if selected else '' }}"></label><label>Telefone<input name="telefone" id="telefone" value="{{ selected.telefone if selected else '' }}"></label></div><label>E-mail<input name="email" id="email" type="email" value="{{ selected.email if selected else '' }}"></label><div class="actions"><button class="primary" type="submit" name="action" value="save">Salvar</button>{% if selected %}<button class="danger" type="submit" name="action" value="delete" onclick="return confirm('Deseja excluir este cliente?')">Excluir</button>{% endif %}</div><p class="form-note" id="cnpj-status" hidden></p></form><section class="panel"><div class="section-title"><h2>Clientes cadastrados</h2><span>{{ clientes|length }}</span></div>{% if clientes %}<div class="client-list">{% for cliente in clientes %}<a class="client-item {% if selected and selected.id == cliente.id %}active{% endif %}" href="{{ url_for('editar_cliente', cliente_id=cliente.id) }}"><strong>{{ cliente.razao_social }}</strong><span>{{ cliente.cnpj or 'CNPJ não informado' }}</span><small>{{ cliente.cidade or 'Cidade não informada' }}{% if cliente.uf %} / {{ cliente.uf }}{% endif %}</small></a>{% endfor %}</div>{% else %}<p class="empty">Nenhum cliente cadastrado.</p>{% endif %}</section></section><script>const b=document.getElementById('buscar-cnpj'),s=document.getElementById('cnpj-status');function v(i,x){const f=document.getElementById(i);if(f&&x)f.value=x}b?.addEventListener('click',async()=>{const c=document.getElementById('cnpj').value.replace(/\\D/g,'');s.hidden=false;s.textContent='Consultando CNPJ...';try{const r=await fetch(`/api/cnpj/${c}`),d=await r.json();if(!r.ok){s.textContent=d.error||'Não foi possível consultar o CNPJ.';return}['razao_social','cnpj','rua','numero','bairro','cidade','uf','cep','ddd','telefone','email'].forEach(k=>v(k,d[k]));s.textContent='Dados preenchidos pelo CNPJ.'}catch(e){s.textContent='Serviço de CNPJ indisponível no momento.'}});</script>{% endblock %}""",
    "insumos.html": """{% extends "base.html" %}{% block content %}<section class="split"><form class="panel form" method="post"><h1>Cadastro de insumos</h1><label>Tipo *<select name="tipo" required><option value="penetrante">Líquido penetrante</option><option value="revelador">Revelador</option><option value="particula">Partícula magnética</option></select></label><label>Fabricante<input name="fabricante"></label><div class="grid two"><label>Fabricação<input name="data_fabricacao" type="month"></label><label>Validade<input name="data_validade" type="month"></label></div><label>Lote<input name="lote"></label><button class="primary" type="submit">Salvar insumo</button></form><section class="panel"><h2>Insumos cadastrados</h2>{% if insumos %}<div class="list">{% for insumo in insumos %}<div class="client-item"><strong>{{ insumo.nome }}</strong><span>Lote {{ insumo.lote or '-' }} | Validade {{ insumo.data_validade or '-' }}</span></div>{% endfor %}</div>{% else %}<p class="empty">Nenhum insumo cadastrado.</p>{% endif %}</section></section>{% endblock %}""",
    "emitir_relatorio.html": """{% extends "base.html" %}{% block content %}<form class="panel form" method="post" enctype="multipart/form-data"><div class="section-title"><h1>Gerar Relatórios</h1><a href="{{ url_for('insumos') }}">Cadastrar insumos</a></div><div class="grid three"><label class="check-option"><input name="incluir_lp" type="checkbox" value="1" {% if default_report == 'lp' %}checked{% endif %}>Gerar Relatório de Líquidos Penetrantes</label><label class="check-option"><input name="incluir_pm" type="checkbox" value="1" {% if default_report == 'pm' %}checked{% endif %}>Gerar Relatório de Partículas Magnéticas</label><label class="check-option"><input name="incluir_us" type="checkbox" value="1" {% if default_report == 'us' %}checked{% endif %}>Gerar Relatório de Ultrassom</label></div><div class="grid two"><label>Formato de saída<select name="output_mode"><option value="unico">Gerar todos em um único arquivo</option><option value="separados">Gerar laudos separados</option></select></label><label>Cliente *<select name="cliente_id" required><option value="">Selecione...</option>{% for cliente in clientes %}<option value="{{ cliente.id }}">{{ cliente.razao_social }}</option>{% endfor %}</select></label></div><div class="grid three"><label>Número do relatório *<input name="NUMRELATORIO" required></label><label>Data da inspeção<input name="DATA_INSP" type="date" value="{{ today }}"></label><label>Laudo<select name="LAUDO"><option value="A">Aprovado</option><option value="R">Reprovado</option></select></label></div><div class="grid two"><label>Peça inspecionada<input name="PECA_INSP"></label><label>Número do desenho / OP<input name="NUM_DESENHO"></label><label>Quantidade<input name="QUANTIDADE"></label><label>Local da inspeção<input name="LOCAL_INSP"></label><label>Temperatura<input name="TEMPERATURA"></label><label>Condição da superfície<input name="COND_SUPERFICIAL"></label><label>Material<input name="MATERIAL"></label><label>Região inspecionada<input name="REGIAO_INSP"></label><label>Espessura<input name="ESPESSURA"></label></div><section class="subsection"><h2>Insumos</h2><div class="grid three"><label>Líquido penetrante<select name="penetrante_id"><option value="">Selecione...</option>{% for insumo in penetrantes %}<option value="{{ insumo.id }}">{{ insumo.nome }}</option>{% endfor %}</select></label><label>Revelador<select name="revelador_id"><option value="">Selecione...</option>{% for insumo in reveladores %}<option value="{{ insumo.id }}">{{ insumo.nome }}</option>{% endfor %}</select></label><label>Partícula magnética<select name="particula_id"><option value="">Selecione...</option>{% for insumo in particulas %}<option value="{{ insumo.id }}">{{ insumo.nome }}</option>{% endfor %}</select></label></div></section><section class="subsection"><h2>Fotos</h2><div class="grid three"><label>Foto 1 / capa<input name="FOTO_1" type="file" accept="image/*"></label><label>Foto 2<input name="FOTO_2" type="file" accept="image/*"></label><label>Foto 3<input name="FOTO_3" type="file" accept="image/*"></label></div></section><button class="primary" type="submit">Gerar relatório</button></form>{% endblock %}""",
    "novo_relatorio.html": """{% extends "emitir_relatorio.html" %}""",
}

app.jinja_loader = ChoiceLoader([
    FileSystemLoader(os.path.join(BASE_DIR, "templates")),
    DictLoader(FALLBACK_TEMPLATES),
    FileSystemLoader(BASE_DIR),
])


def _ensure_dirs() -> None:
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(ORG_TEMPLATES_DIR, exist_ok=True)
    os.makedirs(get_output_dir(), exist_ok=True)


def _current_user(db_session):
    user_id = browser_session.get("user_id")
    if not user_id:
        return None
    return db_session.get(User, int(user_id))


def _require_login():
    open_routes = {"index", "login", "signup", "setup", "forgot_password", "oauth_google", "oauth_facebook", "oauth_apple", "static"}
    g.user_id = None
    g.organization_id = None
    g.user_name = ""
    g.organization_name = ""
    g.user_role = ""

    db_session = get_session()
    try:
        user = _current_user(db_session)
        if user:
            g.user_id = user.id
            g.organization_id = user.organization_id
            g.user_name = user.nome
            g.organization_name = user.organization.nome if user.organization else ""
            g.user_role = user.role
    finally:
        db_session.close()

    if request.endpoint in open_routes:
        return None

    if not g.user_id:
        return redirect(url_for("login"))
    return None


def _current_org_id() -> int:
    return int(g.organization_id)


def _ensure_report_type(session, config: dict) -> TipoRelatorio:
    tipo = session.query(TipoRelatorio).filter_by(nome=config["db_name"]).first()
    if tipo:
        return tipo

    tipo = TipoRelatorio(
        nome=config["db_name"],
        descricao=config["name"],
        schema_json="{}",
        template_path=os.path.join("templates", config["template"]),
    )
    session.add(tipo)
    session.commit()
    return tipo


def _cliente_mapping(cliente: Cliente) -> dict:
    endereco = cliente.rua or ""
    if cliente.numero:
        endereco = f"{endereco}, {cliente.numero}" if endereco else cliente.numero

    return {
        "EMPRESA": cliente.razao_social or "",
        "ENDERECO": endereco,
        "ENDEREÇO": endereco,
        "BAIRRO": cliente.bairro or "",
        "CIDADE": cliente.cidade or "",
        "ESTADO": cliente.uf or "",
        "CEP": cliente.cep or "",
        "CONTATO": cliente.contato or "",
        "DDD": cliente.ddd or "",
        "FONE": cliente.telefone or "",
        "EMAIL": cliente.email or "",
    }


def _parse_date(value: str) -> date:
    if not value:
        return date.today()
    return datetime.strptime(value, "%Y-%m-%d").date()


def _save_upload(report_num: str, field_name: str):
    upload = request.files.get(field_name)
    if not upload or not upload.filename:
        return None

    folder = os.path.join(UPLOAD_DIR, secure_filename(report_num or "sem-numero"))
    os.makedirs(folder, exist_ok=True)
    filename = f"{field_name}_{secure_filename(upload.filename)}"
    path = os.path.join(folder, filename)
    upload.save(path)
    return path


def _json_ready(data: dict) -> dict:
    converted = {}
    for key, value in data.items():
        converted[key] = value.isoformat() if isinstance(value, (date, datetime)) else value
    return converted


def _laudo_extenso(laudo: str) -> str:
    return "aprovado" if laudo == "A" else "reprovado"


def _month_year(value: str) -> str:
    if not value:
        return ""
    try:
        parsed = datetime.strptime(value, "%Y-%m")
        return parsed.strftime("%m/%Y")
    except ValueError:
        return value


def _apply_insumos(
    db_session,
    dados: dict,
    org_id: int,
    penetrante_id: str = "",
    revelador_id: str = "",
    particula_id: str = "",
) -> None:
    penetrante = (
        db_session.query(Insumo)
        .filter_by(id=int(penetrante_id), organization_id=org_id, tipo="penetrante")
        .first()
        if penetrante_id
        else None
    )
    revelador = (
        db_session.query(Insumo)
        .filter_by(id=int(revelador_id), organization_id=org_id, tipo="revelador")
        .first()
        if revelador_id
        else None
    )
    particula = (
        db_session.query(Insumo)
        .filter_by(id=int(particula_id), organization_id=org_id, tipo="particula")
        .first()
        if particula_id
        else None
    )

    if penetrante:
        dados["FAB_PENETRANTE"] = penetrante.data_fabricacao or ""
        dados["VAL_PENETRANTE"] = penetrante.data_validade or ""
        dados["LOTE_PENETRANTE"] = penetrante.lote or ""

    if revelador:
        dados["FAB_REVELADOR"] = revelador.data_fabricacao or ""
        dados["VAL_REVELADOR"] = revelador.data_validade or ""
        dados["LOTE_REVELADOR"] = revelador.lote or ""

    if particula:
        dados["FAB_PARTICULA"] = particula.data_fabricacao or ""
        dados["VAL_PARTICULA"] = particula.data_validade or ""
        dados["LOTE_PARTICULA"] = particula.lote or ""


def _selected_reports() -> list[str]:
    reports = []
    if request.form.get("incluir_lp"):
        reports.append("lp")
    if request.form.get("incluir_pm"):
        reports.append("pm")
    if request.form.get("incluir_us"):
        reports.append("us")
    return reports


def _make_zip(paths: list[str], report_num: str) -> str:
    zip_path = os.path.join(get_output_dir(), f"Relatorios {report_num}.zip")
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in paths:
            archive.write(path, arcname=os.path.basename(path))
    return zip_path


def _active_template_paths(db_session, org_id: int) -> dict[str, str]:
    templates = (
        db_session.query(TemplateEmpresa)
        .filter_by(organization_id=org_id, ativo=1)
        .order_by(TemplateEmpresa.criado_em.desc())
        .all()
    )
    selected = {}
    for template in templates:
        if template.tipo not in selected and os.path.exists(template.file_path):
            selected[template.tipo] = template.file_path
    return selected


def _only_digits(value: str) -> str:
    return re.sub(r"\D", "", value or "")


def _fill_cliente(cliente: Cliente, form) -> None:
    cliente.razao_social = form.get("razao_social", "").strip()
    cliente.contato = form.get("contato") or None
    cliente.cnpj = form.get("cnpj") or None
    cliente.ie = form.get("ie") or None
    cliente.rua = form.get("rua") or None
    cliente.numero = form.get("numero") or None
    cliente.bairro = form.get("bairro") or None
    cliente.cidade = form.get("cidade") or None
    cliente.uf = form.get("uf") or None
    cliente.cep = form.get("cep") or None
    cliente.ddd = form.get("ddd") or None
    cliente.telefone = form.get("telefone") or None
    cliente.email = form.get("email") or None


@app.before_request
def bootstrap():
    init_db()
    _ensure_dirs()
    login_redirect = _require_login()
    if login_redirect:
        return login_redirect


@app.route("/setup", methods=["GET", "POST"])
def setup():
    db_session = get_session()
    try:
        if db_session.query(User).first():
            return redirect(url_for("login"))

        if request.method == "POST":
            org_name = request.form.get("organization_name", "").strip()
            name = request.form.get("name", "").strip()
            email = request.form.get("email", "").strip().lower()
            password = request.form.get("password", "")

            if not org_name or not name or not email or len(password) < 6:
                flash("Preencha empresa, nome, email e uma senha com pelo menos 6 caracteres.", "error")
                return redirect(url_for("setup"))

            org = Organization(nome=org_name)
            db_session.add(org)
            db_session.flush()

            user = User(
                organization_id=org.id,
                nome=name,
                email=email,
                password_hash=generate_password_hash(password),
                role="master",
            )
            db_session.add(user)

            for cliente in db_session.query(Cliente).filter(Cliente.organization_id.is_(None)).all():
                cliente.organization_id = org.id
            for entrada in db_session.query(EntradaRelatorio).filter(EntradaRelatorio.organization_id.is_(None)).all():
                entrada.organization_id = org.id

            db_session.commit()
            browser_session["user_id"] = user.id
            flash("Conta criada. Agora a base web já está isolada por empresa.", "success")
            return redirect(url_for("dashboard"))

        return render_template("setup.html")
    finally:
        db_session.close()


@app.route("/cadastro", methods=["GET", "POST"])
def signup():
    db_session = get_session()
    try:
        if request.method == "POST":
            org_name = request.form.get("organization_name", "").strip()
            name = request.form.get("name", "").strip()
            email = request.form.get("email", "").strip().lower()
            password = request.form.get("password", "")

            if not org_name or not name or not email or len(password) < 6:
                flash("Preencha empresa, nome, e-mail e uma senha com pelo menos 6 caracteres.", "error")
                return redirect(url_for("signup"))

            existing = db_session.query(User).filter_by(email=email).first()
            if existing:
                flash("Já existe uma conta com este e-mail.", "error")
                return redirect(url_for("signup"))

            org = Organization(nome=org_name)
            db_session.add(org)
            db_session.flush()

            user = User(
                organization_id=org.id,
                nome=name,
                email=email,
                password_hash=generate_password_hash(password),
                role="admin",
            )
            db_session.add(user)
            db_session.commit()
            browser_session["user_id"] = user.id
            return redirect(url_for("dashboard"))

        return render_template("signup.html")
    finally:
        db_session.close()


@app.route("/login", methods=["GET", "POST"])
def login():
    db_session = get_session()
    try:
        if not db_session.query(User).first():
            return redirect(url_for("setup"))

        if request.method == "POST":
            email = request.form.get("email", "").strip().lower()
            password = request.form.get("password", "")
            user = db_session.query(User).filter_by(email=email).first()
            if not user or not check_password_hash(user.password_hash, password):
                flash("E-mail ou senha inválidos.", "error")
                return redirect(url_for("login"))

            browser_session["user_id"] = user.id
            flash("Login realizado.", "success")
            return redirect(url_for("dashboard"))

        return render_template("login.html")
    finally:
        db_session.close()


@app.route("/esqueci-minha-senha")
def forgot_password():
    flash("Recuperação de senha será habilitada na próxima etapa.", "error")
    return redirect(url_for("login"))


@app.route("/auth/google")
def oauth_google():
    flash("Login com Google ainda precisa ser configurado no provedor OAuth.", "error")
    return redirect(url_for("login"))


@app.route("/auth/facebook")
def oauth_facebook():
    flash("Login com Facebook ainda precisa ser configurado no provedor OAuth.", "error")
    return redirect(url_for("login"))


@app.route("/auth/apple")
def oauth_apple():
    flash("Login com Apple ainda precisa ser configurado no provedor OAuth.", "error")
    return redirect(url_for("login"))


@app.route("/logout", methods=["POST"])
def logout():
    browser_session.clear()
    flash("Você saiu da aplicação.", "success")
    return redirect(url_for("login"))


@app.route("/")
def index():
    return render_template("landing.html")


@app.route("/app")
def dashboard():
    session = get_session()
    try:
        org_id = _current_org_id()
        clientes_count = session.query(Cliente).filter_by(organization_id=org_id).count()
        return render_template(
            "index.html",
            clientes_count=clientes_count,
            report_types=REPORT_TYPES,
        )
    finally:
        session.close()


@app.route("/templates", methods=["GET", "POST"])
def templates_empresa():
    db_session = get_session()
    try:
        org_id = _current_org_id()

        if request.method == "POST":
            tipo = request.form.get("tipo", "").strip()
            nome = request.form.get("nome", "").strip()
            upload = request.files.get("template_file")

            if tipo not in {"capa", "lp", "pm", "us"} or not nome or not upload or not upload.filename:
                flash("Informe o tipo, o nome e selecione um arquivo DOCX.", "error")
                return redirect(url_for("templates_empresa"))

            filename = secure_filename(upload.filename)
            if not filename.lower().endswith(".docx"):
                flash("O template deve estar no formato DOCX.", "error")
                return redirect(url_for("templates_empresa"))

            folder = os.path.join(ORG_TEMPLATES_DIR, str(org_id), tipo)
            os.makedirs(folder, exist_ok=True)
            path = os.path.join(folder, f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{filename}")
            upload.save(path)

            for old_template in db_session.query(TemplateEmpresa).filter_by(organization_id=org_id, tipo=tipo, ativo=1).all():
                old_template.ativo = 0

            db_session.add(TemplateEmpresa(
                organization_id=org_id,
                tipo=tipo,
                nome=nome,
                file_path=path,
                ativo=1,
            ))
            db_session.commit()
            flash("Template personalizado salvo para esta empresa.", "success")
            return redirect(url_for("templates_empresa"))

        templates = (
            db_session.query(TemplateEmpresa)
            .filter_by(organization_id=org_id)
            .order_by(TemplateEmpresa.criado_em.desc())
            .all()
        )
        return render_template("templates_empresa.html", templates=templates)
    finally:
        db_session.close()


@app.route("/clientes", methods=["GET", "POST"])
def clientes():
    session = get_session()
    try:
        if request.method == "POST":
            razao_social = request.form.get("razao_social", "").strip()
            if not razao_social:
                flash("Razão social é obrigatória.", "error")
                return redirect(url_for("clientes"))

            cliente = Cliente(
                razao_social=razao_social,
                contato=request.form.get("contato") or None,
                cnpj=request.form.get("cnpj") or None,
                ie=request.form.get("ie") or None,
                rua=request.form.get("rua") or None,
                numero=request.form.get("numero") or None,
                bairro=request.form.get("bairro") or None,
                cidade=request.form.get("cidade") or None,
                uf=request.form.get("uf") or None,
                cep=request.form.get("cep") or None,
                ddd=request.form.get("ddd") or None,
                telefone=request.form.get("telefone") or None,
                email=request.form.get("email") or None,
                organization_id=_current_org_id(),
            )
            session.add(cliente)
            session.commit()
            flash("Cliente cadastrado.", "success")
            return redirect(url_for("clientes"))

        items = (
            session.query(Cliente)
            .filter_by(organization_id=_current_org_id())
            .order_by(Cliente.razao_social.asc())
            .all()
        )
        return render_template("clientes.html", clientes=items)
    finally:
        session.close()


@app.route("/clientes/editar/<int:cliente_id>")
def editar_cliente(cliente_id):
    session = get_session()
    try:
        org_id = _current_org_id()
        items = (
            session.query(Cliente)
            .filter_by(organization_id=org_id)
            .order_by(Cliente.razao_social.asc())
            .all()
        )
        selected = session.query(Cliente).filter_by(id=cliente_id, organization_id=org_id).first()
        if not selected:
            flash("Cliente não encontrado.", "error")
            return redirect(url_for("clientes"))
        return render_template("clientes.html", clientes=items, selected=selected)
    finally:
        session.close()


@app.route("/clientes/salvar", methods=["POST"])
def salvar_cliente():
    session = get_session()
    try:
        org_id = _current_org_id()
        action = request.form.get("action", "save")
        cliente_id = request.form.get("cliente_id")

        if action == "delete":
            if not cliente_id:
                flash("Selecione um cliente para excluir.", "error")
                return redirect(url_for("clientes"))
            cliente = session.query(Cliente).filter_by(id=int(cliente_id), organization_id=org_id).first()
            if not cliente:
                flash("Cliente não encontrado.", "error")
                return redirect(url_for("clientes"))
            session.delete(cliente)
            session.commit()
            return redirect(url_for("clientes"))

        if not request.form.get("razao_social", "").strip():
            flash("Razão social é obrigatória.", "error")
            return redirect(url_for("clientes"))

        if cliente_id:
            cliente = session.query(Cliente).filter_by(id=int(cliente_id), organization_id=org_id).first()
            if not cliente:
                flash("Cliente não encontrado.", "error")
                return redirect(url_for("clientes"))
        else:
            cliente = Cliente(organization_id=org_id)
            session.add(cliente)

        _fill_cliente(cliente, request.form)
        session.commit()
        return redirect(url_for("clientes"))
    finally:
        session.close()


@app.route("/api/cnpj/<cnpj>")
def buscar_cnpj(cnpj):
    cnpj_digits = _only_digits(cnpj)
    if len(cnpj_digits) != 14:
        return jsonify({"error": "CNPJ inválido."}), 400

    url = f"https://brasilapi.com.br/api/cnpj/v1/{cnpj_digits}"
    try:
        request_obj = urllib.request.Request(url, headers={"User-Agent": "RLMetais/1.0"})
        with urllib.request.urlopen(request_obj, timeout=12) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            return jsonify({"error": "CNPJ não encontrado."}), 404
        return jsonify({"error": "Não foi possível consultar o CNPJ."}), 502
    except Exception:
        return jsonify({"error": "Serviço de CNPJ indisponível no momento."}), 502

    ddd_telefone = _only_digits(payload.get("ddd_telefone_1", ""))
    ddd = ddd_telefone[:2] if len(ddd_telefone) >= 10 else ""
    telefone = ddd_telefone[2:] if ddd else ddd_telefone

    return jsonify({
        "razao_social": payload.get("razao_social") or payload.get("nome_fantasia") or "",
        "cnpj": cnpj_digits,
        "rua": payload.get("logradouro") or "",
        "numero": payload.get("numero") or "",
        "bairro": payload.get("bairro") or "",
        "cidade": payload.get("municipio") or "",
        "uf": payload.get("uf") or "",
        "cep": payload.get("cep") or "",
        "ddd": ddd,
        "telefone": telefone,
        "email": payload.get("email") or "",
    })


@app.route("/insumos", methods=["GET", "POST"])
def insumos():
    db_session = get_session()
    try:
        org_id = _current_org_id()
        if request.method == "POST":
            tipo = request.form.get("tipo", "").strip()
            if tipo not in {"penetrante", "revelador", "particula"}:
                flash("Informe o tipo do insumo.", "error")
                return redirect(url_for("insumos"))

            lote = request.form.get("lote") or None
            nomes = {
                "penetrante": "Líquido penetrante",
                "revelador": "Revelador",
                "particula": "Partícula magnética",
            }
            nome = nomes[tipo]
            if lote:
                nome = f"{nome} - Lote {lote}"

            insumo = Insumo(
                organization_id=org_id,
                tipo=tipo,
                nome=nome,
                fabricante=request.form.get("fabricante") or None,
                data_fabricacao=_month_year(request.form.get("data_fabricacao", "")) or None,
                data_validade=_month_year(request.form.get("data_validade", "")) or None,
                lote=lote,
            )
            db_session.add(insumo)
            db_session.commit()
            flash("Insumo cadastrado.", "success")
            return redirect(url_for("insumos"))

        items = (
            db_session.query(Insumo)
            .filter_by(organization_id=org_id)
            .order_by(Insumo.tipo.asc(), Insumo.nome.asc())
            .all()
        )
        return render_template("insumos.html", insumos=items)
    finally:
        db_session.close()


@app.route("/relatorios/novo/<report_key>", methods=["GET", "POST"])
def novo_relatorio(report_key):
    config = REPORT_TYPES.get(report_key)
    if not config:
        flash("Tipo de relatório inválido.", "error")
        return redirect(url_for("index"))

    session = get_session()
    try:
        org_id = _current_org_id()
        clientes = (
            session.query(Cliente)
            .filter_by(organization_id=org_id)
            .order_by(Cliente.razao_social.asc())
            .all()
        )
        if request.method == "POST":
            numrel = request.form.get("NUMRELATORIO", "").strip()
            cliente_id = request.form.get("cliente_id")

            if not numrel or not cliente_id:
                flash("Informe o número do relatório e o cliente.", "error")
                return redirect(url_for("novo_relatorio", report_key=report_key))

            cliente = (
                session.query(Cliente)
                .filter_by(id=int(cliente_id), organization_id=org_id)
                .first()
            )
            if not cliente:
                flash("Cliente não encontrado.", "error")
                return redirect(url_for("novo_relatorio", report_key=report_key))

            dados = {
                "NUMRELATORIO": numrel,
                "DATA_INSP": _parse_date(request.form.get("DATA_INSP", "")),
                "LAUDO": request.form.get("LAUDO") or "A",
            }
            dados["LAUDO_EXTENSO"] = _laudo_extenso(dados["LAUDO"])
            for field_name, _label, _field_type in config["fields"]:
                dados[field_name] = request.form.get(field_name, "").strip()

            if report_key == "lp":
                if not request.form.get("penetrante_id") or not request.form.get("revelador_id"):
                    flash("Selecione o líquido penetrante e o revelador.", "error")
                    return redirect(url_for("novo_relatorio", report_key=report_key))
                _apply_lp_insumos(
                    session,
                    dados,
                    request.form.get("penetrante_id", ""),
                    request.form.get("revelador_id", ""),
                    org_id,
                )

            for photo_name in config["photos"]:
                dados[photo_name] = _save_upload(numrel, photo_name)

            dados.update(_cliente_mapping(cliente))

            tipo = _ensure_report_type(session, config)
            template_path = os.path.join(TEMPLATES_DIR, config["template"])
            output_dir = get_output_dir()
            caminho_arquivo = config["generator"](dados, template_path, output_dir)

            entrada = EntradaRelatorio(
                cliente_id=cliente.id,
                organization_id=org_id,
                tipo_relatorio_id=tipo.id,
                relatorio_num=numrel,
                titulo_personalizado=f"Relatório {numrel}-{config['suffix']}",
                dados_json=json.dumps(_json_ready(dados), ensure_ascii=False, indent=2),
                criado_em=datetime.now(),
                caminho_arquivo_gerado=caminho_arquivo,
            )
            session.add(entrada)
            session.commit()
            flash("Relatório gerado com sucesso.", "success")
            return redirect(url_for("download_relatorio", entrada_id=entrada.id))

        return render_template(
            "novo_relatorio.html",
            clientes=clientes,
            config=config,
            report_key=report_key,
            penetrantes=session.query(Insumo).filter_by(organization_id=org_id, tipo="penetrante").order_by(Insumo.nome.asc()).all(),
            reveladores=session.query(Insumo).filter_by(organization_id=org_id, tipo="revelador").order_by(Insumo.nome.asc()).all(),
            today=date.today().isoformat(),
        )
    except Exception as exc:
        session.rollback()
        flash(f"Falha ao gerar relatório: {exc}", "error")
        return redirect(url_for("novo_relatorio", report_key=report_key))
    finally:
        session.close()


@app.route("/relatorios/emitir", methods=["GET", "POST"])
def emitir_relatorio():
    db_session = get_session()
    try:
        org_id = _current_org_id()
        clientes = (
            db_session.query(Cliente)
            .filter_by(organization_id=org_id)
            .order_by(Cliente.razao_social.asc())
            .all()
        )

        if request.method == "POST":
            numrel = request.form.get("NUMRELATORIO", "").strip()
            cliente_id = request.form.get("cliente_id")
            selected = _selected_reports()

            if not numrel or not cliente_id:
                flash("Informe o número do relatório e o cliente.", "error")
                return redirect(url_for("emitir_relatorio"))
            if not selected:
                flash("Selecione ao menos um relatório para gerar.", "error")
                return redirect(url_for("emitir_relatorio"))

            cliente = (
                db_session.query(Cliente)
                .filter_by(id=int(cliente_id), organization_id=org_id)
                .first()
            )
            if not cliente:
                flash("Cliente não encontrado.", "error")
                return redirect(url_for("emitir_relatorio"))

            if "lp" in selected and (
                not request.form.get("penetrante_id") or not request.form.get("revelador_id")
            ):
                flash("Selecione o líquido penetrante e o revelador.", "error")
                return redirect(url_for("emitir_relatorio"))
            if "pm" in selected and not request.form.get("particula_id"):
                flash("Selecione a partícula magnética.", "error")
                return redirect(url_for("emitir_relatorio"))

            dados = {
                "NUMRELATORIO": numrel,
                "DATA_INSP": _parse_date(request.form.get("DATA_INSP", "")),
                "LAUDO": request.form.get("LAUDO") or "A",
            }
            dados["LAUDO_EXTENSO"] = _laudo_extenso(dados["LAUDO"])

            field_names = []
            for report_key in selected:
                field_names.extend(field_name for field_name, _label, _field_type in REPORT_TYPES[report_key]["fields"])
            for field_name in set(field_names):
                dados[field_name] = request.form.get(field_name, "").strip()

            _apply_insumos(
                db_session,
                dados,
                org_id,
                penetrante_id=request.form.get("penetrante_id", ""),
                revelador_id=request.form.get("revelador_id", ""),
                particula_id=request.form.get("particula_id", ""),
            )

            for photo_name in ["FOTO_1", "FOTO_2", "FOTO_3"]:
                dados[photo_name] = _save_upload(numrel, photo_name)
            dados.update(_cliente_mapping(cliente))
            template_paths = _active_template_paths(db_session, org_id)

            generated_paths = []
            if request.form.get("output_mode") == "separados":
                for report_key in selected:
                    docx_path, _pdf_path = generate_end_combo_report(
                        dados,
                        incluir_lp=report_key == "lp",
                        incluir_pm=report_key == "pm",
                        incluir_us=report_key == "us",
                        dados_lp=dados,
                        dados_pm=dados,
                        dados_us=dados,
                        foto_capa=dados.get("FOTO_1"),
                        template_paths=template_paths,
                    )
                    final_path = os.path.splitext(docx_path)[0].replace("-END", f"-{REPORT_TYPES[report_key]['suffix']}") + ".docx"
                    if docx_path != final_path:
                        os.replace(docx_path, final_path)
                    generated_paths.append(final_path)
                caminho_arquivo = generated_paths[0] if len(generated_paths) == 1 else _make_zip(generated_paths, numrel)
            else:
                caminho_arquivo, _pdf_path = generate_end_combo_report(
                    dados,
                    incluir_lp="lp" in selected,
                    incluir_pm="pm" in selected,
                    incluir_us="us" in selected,
                    dados_lp=dados,
                    dados_pm=dados,
                    dados_us=dados,
                    foto_capa=dados.get("FOTO_1"),
                    template_paths=template_paths,
                )

            tipo = _ensure_report_type(
                db_session,
                {
                    "db_name": "Relatório composto",
                    "name": "Relatório composto",
                    "template": "CAPA_TEMPLATE.docx",
                },
            )
            entrada = EntradaRelatorio(
                cliente_id=cliente.id,
                organization_id=org_id,
                tipo_relatorio_id=tipo.id,
                relatorio_num=numrel,
                titulo_personalizado=f"Relatório {numrel}",
                dados_json=json.dumps(_json_ready(dados), ensure_ascii=False, indent=2),
                criado_em=datetime.now(),
                caminho_arquivo_gerado=caminho_arquivo,
            )
            db_session.add(entrada)
            db_session.commit()
            return redirect(url_for("download_relatorio", entrada_id=entrada.id))

        return render_template(
            "emitir_relatorio.html",
            clientes=clientes,
            report_types=REPORT_TYPES,
            default_report=request.args.get("tipo", ""),
            penetrantes=db_session.query(Insumo).filter_by(organization_id=org_id, tipo="penetrante").order_by(Insumo.nome.asc()).all(),
            reveladores=db_session.query(Insumo).filter_by(organization_id=org_id, tipo="revelador").order_by(Insumo.nome.asc()).all(),
            particulas=db_session.query(Insumo).filter_by(organization_id=org_id, tipo="particula").order_by(Insumo.nome.asc()).all(),
            today=date.today().isoformat(),
        )
    except Exception as exc:
        db_session.rollback()
        flash(f"Falha ao gerar relatório: {exc}", "error")
        return redirect(url_for("emitir_relatorio"))
    finally:
        db_session.close()


@app.route("/relatorios/<int:entrada_id>/download")
def download_relatorio(entrada_id: int):
    session = get_session()
    try:
        entrada = (
            session.query(EntradaRelatorio)
            .filter_by(id=entrada_id, organization_id=_current_org_id())
            .first()
        )
        if not entrada or not entrada.caminho_arquivo_gerado:
            flash("Arquivo não encontrado.", "error")
            return redirect(url_for("index"))
        return send_file(entrada.caminho_arquivo_gerado, as_attachment=True)
    finally:
        session.close()


if __name__ == "__main__":
    init_db()
    _ensure_dirs()
    app.run(host="127.0.0.1", port=5000, debug=True)
