import os
import json
import sqlite3

PATH_DB = os.path.join("dados", "banco_sistema.db")
PATH_HISTORICO_JSON = r"S:\Gestão Financeira\Controle cobranças\historico_envios.json"
PATH_HISTORICO_BACKUP = r"S:\Gestão Financeira\Controle cobranças\historico_envios.json.backup"

def configurar_banco():
    print("🔧 Iniciando configuração do Banco de Dados...")
    
    # 1. Cria a pasta dados se não existir
    if not os.path.exists("dados"):
        os.makedirs("dados")
        print("📁 Pasta 'dados' criada.")

    # 2. Conecta/Cria o banco
    conn = sqlite3.connect(PATH_DB)
    c = conn.cursor()

    # 3. Cria as Tabelas (Apagamos a de usuários para resetar as senhas para 123)
    print("📊 Construindo tabelas...")
    c.execute('DROP TABLE IF EXISTS usuarios')
    
    c.execute('''CREATE TABLE IF NOT EXISTS usuarios 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                  usuario TEXT UNIQUE, 
                  senha TEXT,
                  funcao TEXT,
                  telefone TEXT)''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS historico 
                 (nota_fiscal TEXT PRIMARY KEY, 
                  data_envio TEXT, 
                  caminho_pdf TEXT,
                  usuario TEXT)''') # <-- A coluna nova tem que nascer com o banco!
    
    # 4. Insere os Usuários Iniciais com senha 123
    print("👤 Cadastrando usuários padrão com senha '123'...")
    usuarios_padrao = [
        ('ruan', '123', 'Financeiro', '555193371657'),
        ('jose', '123', 'Operacional', '5551988888888'),
        ('renato', '123', 'Licitação', '5551977777777'),
        ('otniel', '123', 'Licitação', '5551966666666')
    ]
    for user, senha, func, tel in usuarios_padrao:
        c.execute("INSERT OR IGNORE INTO usuarios (usuario, senha, funcao, telefone) VALUES (?, ?, ?, ?)", 
                  (user, senha, func, tel))
    conn.commit()

    # 5. Migração Inteligente (Procura o JSON ou o BACKUP)
    arquivo_alvo = None
    if os.path.exists(PATH_HISTORICO_JSON):
        arquivo_alvo = PATH_HISTORICO_JSON
    elif os.path.exists(PATH_HISTORICO_BACKUP):
        arquivo_alvo = PATH_HISTORICO_BACKUP

    if arquivo_alvo:
        print(f"📦 Migrando histórico encontrado em: {arquivo_alvo}")
        try:
            with open(arquivo_alvo, 'r') as f:
                dados_antigos = json.load(f)
            
            for nota, info in dados_antigos.items():
                data = info.get("data", "")
                caminho = info.get("caminho", "")
                c.execute("INSERT OR IGNORE INTO historico (nota_fiscal, data_envio, caminho_pdf) VALUES (?, ?, ?)", 
                          (nota, data, caminho))
            conn.commit()
            
            if arquivo_alvo == PATH_HISTORICO_JSON:
                os.replace(PATH_HISTORICO_JSON, PATH_HISTORICO_BACKUP)
                print("✅ Migração do JSON concluída e arquivo renomeado para .backup!")
            else:
                print("✅ Migração do arquivo .backup concluída com sucesso para o banco!")
                
        except Exception as e:
            print(f"❌ Erro na migração: {e}")
    else:
        print("ℹ️ Nenhum arquivo JSON ou Backup antigo encontrado.")

    conn.close()
    print("\n🚀 Tudo pronto! O banco de dados foi configurado com sucesso.")

if __name__ == "__main__":
    configurar_banco()