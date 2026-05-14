from flask import (
    Flask,
    send_file,
    request,
    render_template,
    redirect,
    url_for,
    session,
)
import os
from datetime import datetime

from database import get_session
from models import Cliente, TipoRelatorio, EntradaRelatorio

from reports_ultrassom import generate_ultrassom_report
from config_relatorios import get_output_dir
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = "super_secret_key"  # TODO trocar depois

# =====================================================
# LOGIN
# =====================================================

VALID_EMAIL = "admin@admin.com"
VALID_PASSWORD = "123456"


def login_required(view):
    def wrapped(*args, **kwargs):
        if "email" not in session:
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    wrapped.__name__ = view.__name__
    return wrapped


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        email = request.form.get("email")
        password = request.form.get("password")

        if email == VALID_EMAIL and password == VALID_PASSWORD:
            session["email"] = email
            return redirect(url_for("dashboard"))
        else:
            error = "Usuário ou senha inválidos"
    return render_template("login.html", error=error)

@app.route("/clientes", methods=["GET", "POST"])
@login_required
def clientes_form():
    """
    Tela de Cadastro / Edição de Clientes (versão web).
    - GET: lista clientes e, se tiver cliente_id na querystring, carrega esse cliente no formulário
    - POST: trata botões Novo / Salvar / Excluir
    """
    session_db = get_session()

    try:
        # ------------------------------
        # 1) Tratamento das ações (POST)
        # ------------------------------
        if request.method == "POST":
            acao = request.form.get("acao")
            cliente_id = request.form.get("cliente_id")

            # Campos do formulário
            dados = {
                "razao_social": request.form.get("razao_social") or "",
                "contato": request.form.get("contato") or "",
                "cnpj": request.form.get("cnpj") or "",
                "ie": request.form.get("ie") or "",
                "rua": request.form.get("rua") or "",
                "numero": request.form.get("numero") or "",
                "bairro": request.form.get("bairro") or "",
                "cidade": request.form.get("cidade") or "",
                "uf": request.form.get("uf") or "",
                "cep": request.form.get("cep") or "",
                "ddd": request.form.get("ddd") or "",
                "telefone": request.form.get("telefone") or "",
                "email": request.form.get("email") or "",
            }

            # 👉 NOVO: limpa o formulário (não mexe em banco)
            if acao == "novo":
                return redirect(url_for("clientes_form"))

            # 👉 SALVAR: cria ou atualiza
            if acao == "salvar":
                if cliente_id:  # atualizar existente
                    cli = session_db.get(Cliente, int(cliente_id))
                    if cli:
                        for campo, valor in dados.items():
                            setattr(cli, campo, valor)
                else:  # novo cliente
                    cli = Cliente(**dados)
                    session_db.add(cli)

                session_db.commit()
                # volta para a tela com o cliente recém-salvo selecionado
                return redirect(url_for("clientes_form", cliente_id=cli.id))

            # 👉 EXCLUIR: apaga o cliente atual
            if acao == "excluir" and cliente_id:
                cli = session_db.get(Cliente, int(cliente_id))
                if cli:
                    session_db.delete(cli)
                    session_db.commit()
                return redirect(url_for("clientes_form"))

        # --------------------------------------
        # 2) GET normal: carrega lista + cliente
        # --------------------------------------
        clientes = session_db.query(Cliente).order_by(Cliente.razao_social).all()

        cliente_atual = None
        cliente_id_get = request.args.get("cliente_id")
        if cliente_id_get:
            cliente_atual = session_db.get(Cliente, int(cliente_id_get))

        return render_template(
            "clientes.html",
            clientes=clientes,
            cliente_atual=cliente_atual,
        )

    finally:
        session_db.close()

@app.route("/relatorio/download/<nome_arquivo>")
@login_required
def download_relatorio(nome_arquivo):
    output_dir = get_output_dir()
    file_path = os.path.join(output_dir, nome_arquivo)
    return send_file(file_path, as_attachment=True)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# =====================================================
# DASHBOARD
# =====================================================

@app.route("/")
@login_required
def dashboard():
    return render_template("dashboard.html")

@app.route("/relatorio/novo")
@login_required
def relatorio_novo():
    # só redireciona para a tela de escolha de tipo
    return redirect(url_for("relatorio_tipo"))


@app.route("/relatorio/tipo", methods=["GET", "POST"])
@login_required
def relatorio_tipo():
    if request.method == "POST":
        tipo = request.form.get("tipo")
        if tipo == "US":
            return redirect(url_for("relatorio_us"))
        elif tipo == "LP":
            return redirect(url_for("relatorio_lp"))
        elif tipo == "PM":
            return redirect(url_for("relatorio_pm"))
    return render_template("relatorio_tipo.html")

@app.route("/relatorio/us", methods=["GET", "POST"])
@login_required
def relatorio_us():
    session_db = get_session()
    try:
        clientes = session_db.query(Cliente).order_by(Cliente.razao_social.asc()).all()

        if request.method == "POST":
            cliente_id = request.form.get("cliente_id")
            numrel = request.form.get("numrel")
            peca = request.form.get("peca")
            num_desenho = request.form.get("num_desenho")
            quantidade = request.form.get("quantidade")
            local_insp = request.form.get("local_insp")
            data_insp_str = request.form.get("data_insp")
            material = request.form.get("material")
            cond_superficial = request.form.get("cond_superficial")
            regiao_insp = request.form.get("regiao_insp")
            espessura = request.form.get("espessura")

            # validações simples
            if not cliente_id or not numrel:
                error = "Cliente e Número do Relatório são obrigatórios."
                return render_template(
                    "relatorio_us.html",
                    clientes=clientes,
                    error=error,
                )

            # converte data (YYYY-MM-DD vindo do input type=date)
            data_insp = datetime.strptime(data_insp_str, "%Y-%m-%d").date()

            # busca cliente
            cliente = session_db.get(Cliente, int(cliente_id))

            # monta endereço
            endereco_txt = ""
            if cliente.rua:
                endereco_txt = cliente.rua
                if cliente.numero:
                    endereco_txt += f", {cliente.numero}"

            # salva fotos enviadas em uma pasta (ex: "output/fotos")
            upload_dir = os.path.join(get_output_dir(), "fotos")
            os.makedirs(upload_dir, exist_ok=True)

            fotos = {}
            for campo in ["foto1", "foto2", "foto3"]:
                file = request.files.get(campo)
                if file and file.filename:
                    filename = secure_filename(file.filename)
                    caminho = os.path.join(upload_dir, filename)
                    file.save(caminho)
                    fotos[campo] = caminho
                else:
                    fotos[campo] = None

            dados = {
                "NUMRELATORIO": numrel,
                "PECA_INSP": peca,
                "NUM_DESENHO": num_desenho,
                "QUANTIDADE": quantidade,
                "LOCAL_INSP": local_insp,
                "DATA_INSP": data_insp,
                "MATERIAL": material,
                "COND_SUPERFICIAL": cond_superficial,
                "REGIAO_INSP": regiao_insp,
                "ESPESSURA": espessura,
                "FOTO_1": fotos["foto1"],
                "FOTO_2": fotos["foto2"],
                "FOTO_3": fotos["foto3"],
            }

            # completa com dados do cliente (EMPRESA, ENDEREÇO etc)
            dados.update({
                "EMPRESA": cliente.razao_social or "",
                "ENDEREÇO": endereco_txt,
                "BAIRRO": cliente.bairro or "",
                "CIDADE": cliente.cidade or "",
                "ESTADO": cliente.uf or "",
                "CEP": cliente.cep or "",
                "CONTATO": cliente.contato or "",
                "DDD": cliente.ddd or "",
                "FONE": cliente.telefone or "",
                "EMAIL": cliente.email or "",
            })

            # pega TipoRelatorio (Ultrassom - US) já cadastrado
            tipo = session_db.query(TipoRelatorio).filter_by(nome="Ultrassom - US").first()
            if not tipo:
                # aqui podemos só dar erro, ou criar automaticamente como no desktop
                return "Tipo de relatório 'Ultrassom - US' não encontrado.", 500

            base_dir = os.path.dirname(os.path.abspath(__file__))
            template_path = os.path.join(base_dir, "templates", "US_TEMPLATE.docx")

            output_dir = get_output_dir()
            caminho_docx = generate_ultrassom_report(dados, template_path, output_dir)

            # grava no banco
            entrada = EntradaRelatorio(
                cliente_id=cliente.id,
                tipo_relatorio_id=tipo.id,
                relatorio_num=numrel,
                titulo_personalizado=f"Relatório {numrel}-US",
                dados_json="{}",  # depois podemos salvar o json bonitinho
                criado_em=datetime.now(),
                caminho_arquivo_gerado=caminho_docx,
            )
            session_db.add(entrada)
            session_db.commit()

            # manda para página de confirmação / download
            return render_template(
                "relatorio_confirmar.html",
                caminho_docx=os.path.basename(caminho_docx),
            )

        # GET → exibe formulário
        return render_template(
            "relatorio_us.html",
            clientes=clientes,
            error=None,
        )
    finally:
        session_db.close()

