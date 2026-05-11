import sqlite3
import os

# Caminho do seu banco de dados
caminho_db = os.path.join("dados", "banco_sistema.db")

print("🧹 Iniciando a faxina do Banco de Dados...")

try:
    conn = sqlite3.connect(caminho_db)
    c = conn.cursor()
    
    # 1. Apaga todo o histórico da Timeline
    c.execute("DELETE FROM log_atividades")
    print("✅ Timeline de Atividades zerada.")
    
    # 2. Apaga todo o histórico das Tabelas (Último E-mail)
    c.execute("DELETE FROM historico")
    print("✅ Histórico de Cobranças da Tabela zerado.")
    
    conn.commit()
    conn.close()
    print("🚀 Faxina concluída! O sistema está pronto para a estreia oficial.")

except Exception as e:
    print(f"❌ Erro ao zerar o banco: {e}")