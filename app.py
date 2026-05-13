from flask import Flask, send_file, request, render_template
import os
import io
from datetime import datetime # ⬅️ IMPORTAÇÃO NECESSÁRIA

# Importa a lógica do seu gerador (config_relatorios)
import config_relatorios 

# Cria uma pasta temporária segura que o Render pode usar para salvar os relatórios
TEMP_DIR = os.path.join(os.getcwd(), 'temp_reports')
os.makedirs(TEMP_DIR, exist_ok=True)

app = Flask(__name__)

@app.route("/")
def home():
    # Rota principal para confirmar que está online
    return "<p>Gerador de Relatórios RL Metais está ONLINE! <a href='/gerar'>Clique aqui para Gerar</a></p>"

@app.route("/gerar", methods=["GET"])
def gerar_relatorio():
    
    # 1. DEFINIÇÃO DE CAMINHOS DINÂMICOS (NOVO CÓDIGO AQUI)
    # =======================================================
    data_atual = datetime.now().strftime("%Y%m%d_%H%M%S")
    nome_arquivo_final = f"Relatorio_RL_Metais_{data_atual}.docx" # ⬅️ Variável definida
    caminho_final = os.path.join(TEMP_DIR, nome_arquivo_final)       # ⬅️ Variável definida
    # =======================================================
    
    # 2. ACESSO À LÓGICA DO SEU GERADOR (BLOCO DE EXECUÇÃO)
    try:
        # ATENÇÃO CRUCIAL:
        # -------------------------------------------------------------
        # Você deve substituir o bloco ABAIXO pela sua FUNÇÃO REAL 
        # de geração de relatórios, garantindo que o arquivo seja salvo
        # no caminho: 'caminho_final'.
        # -------------------------------------------------------------

        # Código REAL de Geração:
        config_relatorios.gerar_relatorio_principal(caminho_final)
        
    except Exception as e:
        return f"Erro na Geração do Relatório: {e}", 500
    
    # 3. DOWNLOAD DO RELATÓRIO
    
    # Usa a função send_file do Flask para forçar o download no navegador.
    try:
        return send_file(caminho_final, 
                         mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                         as_attachment=True, 
                         download_name=nome_arquivo_final)
    except Exception as e:
        return f"Erro ao enviar arquivo: {e}", 500