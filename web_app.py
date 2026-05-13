import json
import os
import re
import urllib.error
import urllib.request
import zipfile
from datetime import date, datetime

from flask import Flask, flash, g, jsonify, redirect, render_template, request, send_file, session as browser_session, url_for
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

from config_relatorios import get_output_dir
from database import get_session, init_db
from models import Cliente, EntradaRelatorio, Insumo, Organization, TipoRelatorio, User
from reports_lp import generate_lp_report
from reports_pm import generate_pm_report
from reports_combo import generate_end_combo_report
from reports_ultrassom import generate_ultrassom_report


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
UPLOAD_DIR = os.environ.get("RL_METAIS_UPLOAD_DIR") or os.path.join(BASE_DIR, "web_uploads")

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


def _ensure_dirs() -> None:
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(get_output_dir(), exist_ok=True)


def _current_user(db_session):
    user_id = browser_session.get("user_id")
    if not user_id:
        return None
    return db_session.get(User, int(user_id))


def _require_login():
    open_routes = {"login", "setup", "static"}
    if request.endpoint in open_routes:
        return None

    db_session = get_session()
    try:
        user = _current_user(db_session)
        if not user:
            return redirect(url_for("login"))
        g.user_id = user.id
        g.organization_id = user.organization_id
        g.user_name = user.nome
        g.organization_name = user.organization.nome if user.organization else ""
    finally:
        db_session.close()
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
                role="admin",
            )
            db_session.add(user)

            for cliente in db_session.query(Cliente).filter(Cliente.organization_id.is_(None)).all():
                cliente.organization_id = org.id
            for entrada in db_session.query(EntradaRelatorio).filter(EntradaRelatorio.organization_id.is_(None)).all():
                entrada.organization_id = org.id

            db_session.commit()
            browser_session["user_id"] = user.id
            flash("Conta criada. Agora a base web já está isolada por empresa.", "success")
            return redirect(url_for("index"))

        return render_template("setup.html")
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
            return redirect(url_for("index"))

        return render_template("login.html")
    finally:
        db_session.close()


@app.route("/logout", methods=["POST"])
def logout():
    browser_session.clear()
    flash("Você saiu da aplicação.", "success")
    return redirect(url_for("login"))


@app.route("/")
def index():
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
