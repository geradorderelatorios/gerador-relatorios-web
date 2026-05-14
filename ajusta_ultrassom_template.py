import os
import sqlite3

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "relatorios_end.db")

conn = sqlite3.connect(DB_PATH)
cur = conn.cursor()

cur.execute("""
    UPDATE tipos_relatorio
    SET template_path = 'templates/US_TEMPLATE.docx'
    WHERE nome = 'Ultrassom - US'
""")

print("Linhas afetadas:", cur.rowcount)

conn.commit()
conn.close()
print("Template do Ultrassom atualizado com sucesso.")
