import os
import json
import sqlite3
import pandas as pd
import eel
from datetime import datetime, timedelta
import win32com.client as win32
import pythoncom
import codecs
import base64
import locale
import re
import tkinter as tk
from dotenv import load_dotenv
from supabase import create_client, Client
from tkinter import filedialog
import warnings
import shutil     # <--- NOVO IMPORT PARA COPIAR ARQUIVO
import tempfile   # <--- NOVO IMPORT PARA ACHAR A PASTA TEMP DO WINDOWS
import zipfile

load_dotenv()

SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

try:
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise ValueError("O Python não conseguiu ler o arquivo .env! Verifique se ele não se chama .env.txt")
        
    supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
    print("✅ Conectado ao Supabase com sucesso!")
except Exception as e:
    supabase = None
    print(f"❌ Erro ao conectar ao Supabase: {e}")

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

try: locale.setlocale(locale.LC_ALL, 'pt_BR.utf8')
except:
    try: locale.setlocale(locale.LC_ALL, 'pt_BR')
    except: pass

# ==============================================================================
# 📂 CAMINHOS DO SISTEMA (PORTABILIDADE MÁXIMA)
# ==============================================================================
PATH_DADOS = "dados"
PATH_ASSINATURAS = os.path.join(PATH_DADOS, "assinaturas")

# O Python cria a pasta "dados" e a "assinaturas" sozinho se elas não existirem!
if not os.path.exists(PATH_DADOS): os.makedirs(PATH_DADOS)
if not os.path.exists(PATH_ASSINATURAS): os.makedirs(PATH_ASSINATURAS)

PATH_DB = os.path.join(PATH_DADOS, "banco_sistema.db")

# Seus caminhos originais (mantenha os que você usa na rede)
PATH_2025 = r"S:\Gestão Financeira\Faturamento contratos\Relatório de Serviços  2025\Relatório de serviços Grupo Nascimento 2025.xlsx"
PATH_2026 = r"S:\Gestão Financeira\Faturamento contratos\Relatório de Serviços  2026\Relatório de serviços Grupo Nascimento 2026.xlsx"
PATH_EVIDENCIAS_RAIZ = r"S:\Gestão Financeira\Controle cobranças\Cobranças"

# ==============================================================================
# 🚀 FUNÇÃO UTILITÁRIA PARA LEITURA RÁPIDA DE EXCEL NA REDE
# ==============================================================================
def ler_excel_ninja(caminho_rede, **kwargs):
    """Copia o Excel da rede para o PC local antes de ler, deixando 10x mais rápido"""
    if not os.path.exists(caminho_rede):
        return pd.DataFrame()
    
    nome_arquivo = os.path.basename(caminho_rede)
    pasta_temp = tempfile.gettempdir()
    caminho_local = os.path.join(pasta_temp, f"temp_{nome_arquivo}")
    
    try:
        shutil.copy2(caminho_rede, caminho_local)
        # O **kwargs repassa o usecols e o engine certinho pro Pandas!
        df = pd.read_excel(caminho_local, **kwargs)
        return df
    except Exception as e:
        print(f"Erro ao ler o Excel rápido: {e}")
        return pd.DataFrame()

# ==============================================================================
# 🏢 CLASSE PRINCIPAL DO GESTOR
# ==============================================================================
class GestorCobrancaEel:
    def __init__(self):
        # Garante que a pasta existe antes de conectar
        if not os.path.exists("dados"): os.makedirs("dados")
            
        self.conn = sqlite3.connect(PATH_DB, check_same_thread=False)
        
        try:
            # --- TABELA DE VÍNCULOS OPERACIONAIS ---
            self.conn.execute('''CREATE TABLE IF NOT EXISTS vinculos_operacao 
                                (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                                contrato_sistema TEXT UNIQUE, 
                                analista_nome TEXT,
                                analista_email TEXT,
                                supervisor_nome TEXT,
                                supervisor_tel TEXT,
                                gerente_nome TEXT,
                                gerente_tel TEXT)''')
            self.conn.commit()
            print("✅ Tabela vinculos_operacao verificada/criada com sucesso!")
        except Exception as e:
            # 👇 ISSO AQUI É O SEU MELHOR AMIGO AGORA!
            print(f"❌ ERRO CRÍTICO AO CRIAR TABELA: {e}")

        try:
            self.conn.execute("ALTER TABLE log_atividades ADD COLUMN caminho_anexo TEXT")
            self.conn.commit()
        except:
            pass
        
        # 🚀 MIGRATION: Tenta adicionar a coluna 'anexo' se ela não existir nas antigas
        try:
            self.conn.execute("ALTER TABLE log_atividades ADD COLUMN anexo TEXT DEFAULT ''")
            self.conn.commit()
            print("✅ Coluna 'anexo' adicionada na tabela log_atividades com sucesso!")
        except:
            pass # Se der erro, é porque a coluna já existe, então segue o baile!

        # --- TABELA HISTÓRICO ---
        try:
            self.conn.execute("ALTER TABLE historico ADD COLUMN usuario TEXT")
            self.conn.commit()
        except:
            pass 

        # --- TABELA DE TEMPLATES DE E-MAIL E WHATSAPP ---
        try:
            # 1. Cria a tabela com todas as 6 colunas (Para bancos novos, do zero)
            self.conn.execute('''CREATE TABLE IF NOT EXISTS templates_email 
                                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                                  nome_identificador TEXT UNIQUE, 
                                  assunto TEXT, 
                                  corpo TEXT,
                                  anexo TEXT,
                                  responsavel TEXT)''')
            
            # 2. Tenta injetar as colunas novas no banco velho separadamente
            # Fazemos um try/except para cada coluna, assim se uma já existir, não trava a outra
            try:
                self.conn.execute("ALTER TABLE templates_email ADD COLUMN anexo TEXT")
            except: pass
            
            try:
                self.conn.execute("ALTER TABLE templates_email ADD COLUMN responsavel TEXT")
            except: pass
            
            self.conn.commit()
            
            # 3. Planta os textos iniciais caso a tabela esteja zerada
            self.semear_templates_iniciais() 
        except Exception as e: 
            print(f"Erro ao inicializar banco de dados: {e}")

        # --- VARIÁVEIS DO SISTEMA ---
        self.usuario_logado = "Sistema" 
        self.todas_as_notas = []
        self.historico = self.carregar_historico()
        self.ultima_filtragem = [] 
        self.assinatura_base64 = ""
        
    def semear_templates_iniciais(self):
        """Insere os textos padrões da nova regra da CEO completos"""
        c = self.conn.cursor()
        c.execute("SELECT COUNT(*) FROM templates_email")
        if c.fetchone()[0] == 0:
            templates = [
                ("FAIXA AMARELA (16 A 30 DIAS)", "Cobrança Administrativa - Contrato {cliente}", 
                 "Prezados(as),<br><br>Cumprimentando-os cordialmente, vimos, por meio deste, reiterar formalmente a pendência administrativa e financeira relacionada ao Contrato <b>{cliente}</b>, firmado entre este órgão e a empresa <b>{ass_emp}</b>.<br><br>Conforme histórico de tratativas já estabelecido, foram realizadas diversas tentativas de regularização ao longo das últimas semanas, por meio de comunicações formais por e-mail, contatos institucionais com o setor financeiro, alinhamentos com a fiscalização contratual e solicitações expressas de providência.<br><br>Não obstante tais esforços, até o presente momento permanece pendente de regularização o pagamento referente {t_nota}{t_num}, {t_emi}{t_comp}{t_venc} conforme informações abaixo:<br><br>{lista_html}<br><b>Valor Total Atualizado: {valor_total}</b><br><br>Importa destacar que, desde o início da contratação, a empresa vem mantendo postura absolutamente colaborativa, diligente e comprometida com a execução integral dos serviços, assegurando padrão técnico elevado e regularidade operacional.<br><br>Ressalta-se que, mesmo diante do cenário de inadimplência, os serviços vêm sendo executados de forma contínua e ininterrupta. Nesse sentido, registra-se de forma expressa que:<br><ul><li>Os colaboradores permanecem regularmente em atividade;</li><li>A empresa mantém integral cumprimento de suas obrigações trabalhistas e fiscais;</li><li>Não houve interrupção ou descontinuidade na execução dos serviços.</li></ul><br>Todavia, a permanência da pendência ora relatada passa a produzir impacto relevante sobre a equação econômico-financeira do contrato, afetando diretamente o fluxo financeiro necessário à manutenção dos serviços continuados.<br><br>Diante desse contexto, solicitamos, de forma expressa e objetiva a indicação das providências adotadas até o momento e a informação precisa da data prevista para regularização.<br><br>Ressaltamos que esta empresa permanece, como sempre, aberta ao diálogo institucional.<br>Permanecemos à disposição para realização de reunião institucional imediata, caso entendam pertinente."),
                
                ("FAIXA LARANJA (31 A 60 DIAS)", "Notificação à Autoridade Superior I - {cliente}", 
                 "Prezados(as),<br><br>Encaminhamos a presente comunicação, em caráter formal e institucional, para fins de ciência da autoridade superior acerca da pendência financeira relacionada ao Contrato <b>{cliente}</b>, firmado entre este órgão e a empresa <b>{ass_emp}</b>.<br><br>Conforme histórico já registrado, foram realizadas diversas tratativas administrativas junto à fiscalização contratual e ao setor financeiro. Não obstante tais esforços, permanece pendente a regularização do pagamento devido referente {t_nota}{t_num}, conforme detalhado abaixo:<br><br>{lista_html}<br><b>Valor Total Atualizado: {valor_total}</b><br><br>Cumpre destacar que a empresa vem mantendo integralmente a execução dos serviços contratados, preservando a regularidade operacional e a manutenção dos postos de trabalho.<br><br>Todavia, a permanência da pendência financeira, após sucessivas tentativas de solução em nível técnico-administrativo, ultrapassa o âmbito ordinário de gestão contratual, passando a demandar conhecimento e eventual atuação da instância superior, em razão dos reflexos diretos na sustentabilidade da execução contratual.<br><br>Nesse contexto, a presente comunicação tem por finalidade dar ciência à autoridade superior, possibilitando o apoio institucional necessário à adoção de providências administrativas adequadas, com vistas à regularização da pendência e à preservação do equilíbrio contratual.<br><br>A empresa reafirma sua postura colaborativa, sua boa-fé na condução da relação contratual e sua plena disposição para construção de solução administrativa célere e eficiente."),
                
                ("FAIXA VERMELHA (61 A 90 DIAS)", "NOTIFICAÇÃO ADMINISTRATIVA FORMAL FORTE - {cliente}", 
                 "Prezados(as),<br><br>Cumprimentando-os cordialmente, vimos, por meio da presente <b>NOTIFICAÇÃO ADMINISTRATIVA FORMAL</b>, reiterar, em caráter expresso, técnico e fundamentado, a pendência financeira relacionada ao Contrato <b>{cliente}</b>.<br><br>Até o presente momento, permanece <b>sem a devida liquidação</b> o pagamento referente {t_nota}{t_num}, {t_emi}{t_comp}{t_venc}, configurando atraso administrativo relevante e injustificado de <b>{dias_max} dias</b>.<br><br>{lista_html}<br><b>Valor Total Atualizado: {valor_total}</b><br><br>Importa destacar que a empresa, desde o início da contratação, vem adotando postura absolutamente colaborativa. Durante todo o período contratual, os serviços foram mantidos de forma contínua e ininterrupta.<br><br>Não obstante o histórico consistente de tratativas, a pendência permanece sem qualquer solução efetiva até o momento, impactando diretamente a equação econômico-financeira do contrato. Cumpre registrar, de forma expressa, que a contratada <b>não deu causa ao atraso verificado</b>.<br><br><b>DA POSSÍVEL ATUAÇÃO DO TRIBUNAL DE CONTAS</b><br>Diante disso, a persistência da inadimplência poderá ensejar, caso não regularizada em prazo razoável, a comunicação formal dos fatos ao Tribunal de Contas competente, para fins de apuração e adoção das medidas cabíveis.<br><br>Tal providência poderá ser adotada sem prejuízo da utilização concomitante ou posterior das medidas judiciais cabíveis.<br><br>Dessa forma, solicitamos formalmente a imediata regularização dos pagamentos em atraso ou a indicação precisa da data prevista para quitação do débito. Fica desde já formalmente registrada a reserva de adoção de todas as medidas administrativas e jurídicas cabíveis."),
                
                ("FAIXA VERMELHA - AUTORIDADE SUP. II (61 A 90 DIAS)", "Notificação à Autoridade Superior II - {cliente}", 
                 "Prezados(as),<br><br>Encaminhamos a presente comunicação, em caráter formal e institucional, para fins de ciência da autoridade superior, acerca da pendência financeira/administrativa relacionada ao Contrato <b>{cliente}</b>.<br><br>Não obstante o cenário de plena adimplência contratual por parte da empresa, permanece sem solução a pendência financeira referente {t_nota}{t_num}, com emissão em {t_com_venc}, totalizando <b>{dias_max} dias de atraso</b>, detalhada abaixo:<br><br>{lista_html}<br><b>Valor Total Atualizado: {valor_total}</b><br><br>A empresa vem atuando com absoluta boa-fé, responsabilidade institucional e espírito colaborativo. Contudo, a persistência da inadimplência, mesmo diante da existência de dotação orçamentária, pode indicar situação de irregularidade na execução orçamentária e financeira.<br><br>O prolongamento da situação ultrapassa as tratativas administrativas ordinárias, produzindo impacto relevante sobre o equilíbrio econômico-financeiro do contrato. A manutenção do quadro poderá ensejar a necessidade de comunicação formal dos fatos ao Tribunal de Contas competente.<br><br>Diante do exposto, solicita-se, respeitosamente:<br><ul><li>A ciência da autoridade superior quanto à situação relatada;</li><li>O apoio institucional para viabilização da regularização da pendência;</li><li>A informação acerca da previsão de regularização da obrigação pendente.</li></ul><br>Seguimos à disposição para quaisquer esclarecimentos."),
                
                ("FAIXA ROXA - MINUTA JURÍDICO (91 A 120 DIAS)", "MINUTA INTERNA PARA ENCAMINHAMENTO JURÍDICO | {cliente}", 
                 "<b>MINUTA INTERNA PARA ENCAMINHAMENTO JURÍDICO</b><br><br><b>Empresa:</b> {ass_emp}<br><b>Contrato:</b> {cliente}<br><b>Valor em aberto:</b> {valor_total}<br><b>Atraso Máximo:</b> {dias_max} dias<br><br>{lista_html}<br><b>1. HISTÓRICO RESUMIDO</b><br>A empresa realizou, ao longo do período de inadimplência, sucessivas e devidamente documentadas tratativas administrativas. Dentre as medidas adotadas, destacam-se: envio de comunicações formais por e-mail, reuniões administrativas e escalonamento institucional.<br><br><b>2. EXECUÇÃO CONTRATUAL</b><br>Durante todo o período, a empresa manteve integralmente a execução contratual. Registra-se que a inadimplência não decorre de qualquer falha da contratada, mas exclusivamente de pendência administrativa do contratante.<br><br><b>3. ENCAMINHAMENTO PARA ANÁLISE</b><br>Encaminha-se o presente dossiê ao Comitê Gestor e ao setor jurídico para avaliação estratégica quanto às medidas cabíveis, incluindo, mas não se limitando a:<br><ul><li>Intensificação da cobrança administrativa formal;</li><li>Adoção de medida judicial cabível;</li><li>Requerimento de recomposição do equilíbrio econômico-financeiro.</li></ul>"),
                
                ("FAIXA ROXA - PARECER PREVENTIVO (91 A 120 DIAS)", "PARECER INTERNO PREVENTIVO | {cliente}", 
                 "<b>PARECER INTERNO PREVENTIVO PARA RESERVA DE MEDIDAS ADMINISTRATIVAS E JUDICIAIS</b><br><br>Considerando o histórico consistente de sucessivas tratativas administrativas realizadas pela empresa vinculadas ao Contrato <b>{cliente}</b>;<br><br>Considerando que a contratada não deu causa à pendência financeira verificada ({valor_total} com {dias_max} dias de atraso), tendo cumprido integralmente suas obrigações contratuais, trabalhistas e fiscais;<br><br>A persistência da inadimplência passa a gerar risco concreto e relevante ao equilíbrio econômico-financeiro do contrato. Fica tecnicamente registrada, em caráter preventivo e resguardatório, a possibilidade de adoção futura das seguintes medidas administrativas e judiciais:<br><ul><li>Intensificação da notificação administrativa formal;</li><li>Formulação de pedido de recomposição do equilíbrio econômico-financeiro;</li><li>Propositura de ação de obrigação de fazer ou cobrança judicial.</li></ul><br>O presente parecer deverá integrar o dossiê interno do contrato, compondo o conjunto probatório consolidado."),
                 
                ("FAIXA PRETA (+120 DIAS)", "ENCAMINHAMENTO OBRIGATÓRIO AO JURÍDICO - AÇÃO DE COBRANÇA | {cliente}", 
                 "<b>ENCAMINHAMENTO DE DOSSIÊ PARA ADOÇÃO DE MEDIDAS JUDICIAIS</b><br><br><b>Empresa:</b> {ass_emp}<br><b>Contrato:</b> {cliente}<br><b>Valor em aberto:</b> {valor_total}<br><b>Atraso Máximo:</b> {dias_max} dias<br><br>{lista_html}<br><b>1. CONFIGURAÇÃO DA INADIMPLÊNCIA GRAVE</b><br>Configurada a inadimplência grave, reiterada e prolongada, com evidente comprometimento do equilíbrio econômico-financeiro do contrato, encaminhamos a presente demanda ao advogado para adoção das medidas judiciais cabíveis.<br><br><b>2. INSTRUÇÃO DO DOSSIÊ</b><br>O presente encaminhamento encontra-se instruído com dossiê completo e devidamente organizado, formado ao longo de todas as etapas anteriores do fluxo, contemplando: comunicações formais (ANEXO I), notificações administrativas (ANEXO III), comunicações à autoridade superior (ANEXOS II e IV), além do contrato administrativo, notas fiscais emitidas, registros de tratativas e memória detalhada do impacto financeiro.<br><br><b>3. ROBUSTEZ PROBATÓRIA E AÇÃO JUDICIAL</b><br>Nesta fase, o procedimento já se encontra plenamente estruturado sob os aspectos probatório, técnico e jurídico, com demonstração inequívoca da boa-fé da contratada e da responsabilidade da Administração pela mora. Tal robustez documental viabiliza a adoção segura e estratégica das medidas judiciais pertinentes, inclusive a propositura de ação de cobrança, obrigação de fazer e/ou pedido de tutela de urgência, com o objetivo de viabilizar o faturamento e resguardar, de forma efetiva, o equilíbrio econômico-financeiro do contrato."), # <---- ESSA VÍRGULA AQUI FALTAVA!

                 ("WHATSAPP - MENSAGEM PADRÃO", "Mensagem Direta pelo WhatsApp", 
                 "Oi, tudo bem?\n\nAcabei de enviar um e-mail com as notas pendentes de pagamento do *{cliente}*. Precisamos da sua ajuda para cobrar esses pagamentos, pois o atraso pode gerar sérios problemas financeiros para a empresa, afetando até o pagamento dos colaboradores.\n\nPor favor, dê uma olhada no e-mail e nos ajude a agilizar isso.\n\nAgradeço desde já pela colaboração!")
            ]

            c.executemany("INSERT INTO templates_email (nome_identificador, assunto, corpo) VALUES (?, ?, ?)", templates)
            self.conn.commit()       

    def carregar_historico(self):
        c = self.conn.cursor()
        c.execute("SELECT nota_fiscal, data_envio, caminho_pdf, usuario FROM historico")
        resultados = c.fetchall()
        
        hist_dict = {}
        for linha in resultados:
            nota_banco_limpa = str(linha[0]).replace('.0', '').strip()
            user = linha[3] if linha[3] else "Desconhecido"
            hist_dict[nota_banco_limpa] = {"data": linha[1], "caminho": linha[2], "usuario": user}
        return hist_dict

    def salvar_historico_envio(self, numero_nota, caminho_arquivo_salvo):
        hoje_str = datetime.now().strftime("%d/%m/%Y")
        nota_str = str(numero_nota)
        try:
            c = self.conn.cursor()
            c.execute("INSERT OR REPLACE INTO historico (nota_fiscal, data_envio, caminho_pdf, usuario) VALUES (?, ?, ?, ?)", 
                      (nota_str, hoje_str, caminho_arquivo_salvo, self.usuario_logado))
            self.conn.commit()
        except Exception as e:
            print(f"Erro ao salvar no banco: {e}")

        self.historico[nota_str] = {"data": hoje_str, "caminho": caminho_arquivo_salvo, "usuario": self.usuario_logado}
        
    def autenticar_no_banco(self, usuario, senha):
        c = self.conn.cursor()
        # Buscamos o id e a função (setor) do usuário
        c.execute("SELECT id, funcao FROM usuarios WHERE usuario = ? AND senha = ?", (usuario, senha))
        resultado = c.fetchone()
        if resultado:
            self.usuario_logado = str(usuario).capitalize()
            self.funcao_logada = str(resultado[1]).upper() if resultado[1] else "" # <--- SALVA O SETOR
            return True
        return False

    def obter_assinatura_base64(self, nome_usuario):
        import base64
        import os
        
        # 1. Limpa o nome (ex: "Ruan" vira "ruan.png")
        usuario_limpo = str(nome_usuario).strip().lower()
        caminho_img = os.path.join(PATH_ASSINATURAS, f"{usuario_limpo}.png")
        
        print(f"\n🕵️‍♂️ ASSINATURA: O usuário logado é '{usuario_limpo}'")
        print(f"👉 Procurando arquivo em: {caminho_img}")
        
        # 2. Se não achar a do usuário, tenta a 'padrao.png'
        if not os.path.exists(caminho_img):
            print(f"⚠️ Arquivo '{usuario_limpo}.png' não achado! Tentando o estepe 'padrao.png'...")
            caminho_img = os.path.join(PATH_ASSINATURAS, "padrao.png")
            
        # 3. Se não houver nenhuma, retorna vazio
        if not os.path.exists(caminho_img):
            print(f"❌ Falha total: Nenhuma imagem encontrada em {caminho_img}\n")
            return ""

        try:
            with open(caminho_img, "rb") as img_file:
                print("✅ Imagem de assinatura carregada com sucesso!\n")
                return "data:image/png;base64," + base64.b64encode(img_file.read()).decode('utf-8')
        except Exception as e:
            print(f"❌ Erro ao converter assinatura: {e}\n")
            return ""

    def extrair_float(self, valor):
        if pd.isna(valor) or valor == '-': return 0.0
        try:
            if isinstance(valor, str): return float(valor.replace('R$', '').replace('.', '').replace(',', '.').strip())
            return float(valor)
        except: return 0.0

    def formatar_moeda(self, valor_float):
        texto = f"{valor_float:,.2f}".replace(',', 'v').replace('.', ',').replace('v', '.')
        return f"R$ {texto}"

    def definir_parecer(self, dias):
        # dias_atraso = dias desde a emissão - 30 (prazo da empresa)
        dias_atraso = dias - 30

        # --- FASE 1: NOTAS QUE NÃO VENCERAM (No Prazo / Aviso Amigável) ---
        # Se o atraso é 0 ou negativo, a nota está saudável.
        # Retornamos o nome EXATO que você cadastrou no Banco de Dados.
        if dias_atraso <= 0:
            return "AVISO AMIGÁVEL (PRÉ-VENCIMENTO)"

        # --- FASE 2: FLUXO DE COBRANÇA (A PARTIR DO 1º DIA DE ATRASO REAL) ---
        # Daqui para baixo, o código só chega se a nota estiver VENCIDA (> 0)
        if dias_atraso <= 15: return "FAIXA VERDE (ATÉ 15 DIAS)"
        if dias_atraso <= 30: return "FAIXA AMARELA (16 A 30 DIAS)"
        if dias_atraso <= 60: return "FAIXA LARANJA (31 A 60 DIAS)"
        if dias_atraso <= 90: return "FAIXA VERMELHA (61 A 90 DIAS)"
        if dias_atraso <= 120: return "FAIXA ROXA (91 A 120 DIAS)"
        
        return "FAIXA PRETA (+120 DIAS)"

    def encontrar_ou_criar_pasta_cliente(self, nome_cliente_excel, empresa_excel):
        try:
            ano_atual = datetime.now().strftime("%Y")
            empresa_limpa = str(empresa_excel).strip().upper()
            nome_seguro_contrato = nome_cliente_excel.replace("/", ".").replace("\\", ".")
            nome_seguro_contrato = re.sub(r'[<>:"|?*]', '', nome_seguro_contrato).strip()
            caminho_base = os.path.join(PATH_EVIDENCIAS_RAIZ, ano_atual, empresa_limpa)
            if not os.path.exists(caminho_base): os.makedirs(caminho_base)
            
            nome_limpo_excel = re.sub(r'[^a-zA-Z0-9]', '', nome_cliente_excel).upper()
            pastas_existentes = [f for f in os.listdir(caminho_base) if os.path.isdir(os.path.join(caminho_base, f))]
            melhor_pasta = None
            for pasta in pastas_existentes:
                nome_limpo_pasta = re.sub(r'[^a-zA-Z0-9]', '', pasta).upper()
                if nome_limpo_excel in nome_limpo_pasta or nome_limpo_pasta in nome_limpo_excel:
                    melhor_pasta = pasta
                    break
            
            if melhor_pasta: return os.path.join(caminho_base, melhor_pasta)
            else:
                nova_pasta = os.path.join(caminho_base, nome_seguro_contrato)
                if not os.path.exists(nova_pasta): os.makedirs(nova_pasta)
                return nova_pasta
        except: return PATH_EVIDENCIAS_RAIZ

    def gerar_evidencia_pdf(self, assunto, corpo, remetente, cliente_nome, nome_arquivo, empresa_nome):
        word = None
        doc = None
        try:
            pasta_destino = self.encontrar_ou_criar_pasta_cliente(cliente_nome, empresa_nome)
            nome_limpo = "".join([c for c in nome_arquivo if c.isalpha() or c.isdigit() or c in ' ._-'])
            
            # 🚨 ABSOLUTE PATH: Essencial para o Word não dar erro silencioso
            caminho_final = os.path.abspath(os.path.join(pasta_destino, f"{nome_limpo}.pdf"))
            temp_html = os.path.abspath(os.path.join(PATH_DADOS, "temp_email_evidencia.html"))

            if os.path.exists(caminho_final):
                try: os.rename(caminho_final, caminho_final)
                except: return "ERRO_ABERTO"

            data_extenso = datetime.now().strftime("%A, %d de %B de %Y %H:%M")
            
            # 🚀 ASSINATURA DINÂMICA: Puxa de quem está logado no sistema
            assinatura_img = self.obter_assinatura_base64(self.usuario_logado)
            tag_assinatura = f'<div class="assinatura"><img src="{assinatura_img}" alt="Assinatura" width="200"></div>' if assinatura_img else ""

            html_content = f"""
            <html><head><meta charset="utf-8">
            <style>
                body {{ font-family: 'Calibri', Arial, sans-serif; font-size: 11pt; color: #000; margin: 20px; }}
                .remetente-topo {{ font-weight: bold; font-size: 12pt; border-bottom: 2px solid #000; padding-bottom: 5px; margin-bottom: 15px; }}
                .header-block {{ margin-bottom: 20px; }} 
                .row {{ margin-bottom: 3px; }} 
                .label {{ font-weight: bold; color: #000; display: inline-block; width: 80px; }}
                .content {{ margin-top: 30px; line-height: 1.6; }} 
                .assinatura {{ margin-top: 40px; }}
            </style></head>
            <body>
                <div class="remetente-topo">{remetente}</div>
                <div class="header-block">
                    <div class="row"><span class="label">De:</span> {remetente}</div>
                    <div class="row"><span class="label">Enviado em:</span> {data_extenso}</div>
                    <div class="row"><span class="label">Para:</span> {cliente_nome}</div>
                    <div class="row"><span class="label">Assunto:</span> {assunto}</div>
                </div>
                <div class="content">{corpo}</div>
                {tag_assinatura}
            </body></html>"""
            
            with open(temp_html, "w", encoding="utf-8") as f: 
                f.write(html_content)

            pythoncom.CoInitialize()
            word = win32.Dispatch('Word.Application')
            word.Visible = False
            
            # O Word agora abre o caminho completo, sem erro!
            doc = word.Documents.Open(temp_html)
            doc.PageSetup.LeftMargin, doc.PageSetup.RightMargin = 50, 50
            doc.PageSetup.TopMargin, doc.PageSetup.BottomMargin = 50, 50
            
            doc.SaveAs(caminho_final, FileFormat=17) 
            doc.Close(False)
            
            if os.path.exists(temp_html): 
                os.remove(temp_html)
                
            return caminho_final

        except Exception as e: 
            print(f"❌ Erro ao gerar evidência: {e}")
            return str(e)
        finally:
            if doc: 
                try: doc.Close(False)
                except: pass
            if word: 
                try: word.Quit()
                except: pass

gestor = GestorCobrancaEel()

@eel.expose
def autenticar_usuario(usuario, senha):
    # Agora consulta de verdade no Banco de Dados!
    sucesso = gestor.autenticar_no_banco(usuario.lower(), senha)
    if sucesso:
        return {"status": "sucesso"}
    return {"status": "erro", "mensagem": "Usuário ou senha inválidos."}

@eel.expose
def registrar_log_atividade(cliente, nota_fiscal, acao):
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    try:
        c = gestor.conn.cursor()
        c.execute("INSERT INTO log_atividades (cliente, nota_fiscal, data_hora, usuario, acao) VALUES (?, ?, ?, ?, ?)",
                  (cliente, nota_fiscal, agora, gestor.usuario_logado, acao))
        gestor.conn.commit()
    except Exception as e: print("Erro no log:", e)

@eel.expose
def buscar_historico_cliente(cliente):
    try:
        c = gestor.conn.cursor()
        
        # 🚀 O SEGREDO ESTÁ AQUI: Adicionamos a coluna 'anexo' no SELECT!
        c.execute("""
            SELECT id, nota_fiscal, data_hora, usuario, acao, anexo 
            FROM log_atividades 
            WHERE cliente = ? 
            ORDER BY id DESC
        """, (cliente,))
        
        linhas = c.fetchall()
        
        logs = []
        for r in linhas:
            logs.append({
                "id": r[0],
                "nota": r[1],
                "data": r[2],
                "usuario": r[3],
                "acao": r[4],
                "anexo": r[5] if r[5] else ""  # 👈 Puxa o caminho do anexo pro Javascript
            })
            
        return logs
    except Exception as e:
        print(f"❌ Erro ao buscar histórico do cliente: {e}")
        return []

@eel.expose
def obter_contas_outlook():
    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('outlook.application')
        contas = []
        for account in outlook.Session.Accounts:
            try: contas.append(account.SmtpAddress if account.SmtpAddress else account.DisplayName)
            except: contas.append(account.DisplayName)
        return contas if contas else ["Conta Padrão"]
    except: return ["Conta Padrão"]

@eel.expose
def obter_contatos_operacao():
    # Busca apenas os usuários que têm telefone cadastrado
    c = gestor.conn.cursor()
    c.execute("SELECT usuario, funcao, telefone FROM usuarios WHERE telefone IS NOT NULL AND telefone != ''")
    resultados = c.fetchall()
    
    contatos = []
    for r in resultados:
        contatos.append({
            "nome": str(r[0]).capitalize(), 
            "funcao": str(r[1]), 
            "telefone": str(r[2])
        })
    return contatos

@eel.expose
def carregar_dados_reais(forcar=False):
    gestor.historico = gestor.carregar_historico() # Puxa do DB fresquinho (Demora 0.01 segundos!)
    
    # 🛡️ O ESCUDO DO CACHE: Se já temos as notas na memória, pulamos o Excel pesado!
    if len(gestor.todas_as_notas) > 0 and not forcar:
        # Apenas atualizamos a etiqueta de último envio de cada nota caso o outro usuário tenha mandado e-mail
        for nota in gestor.todas_as_notas:
            dados_hist = gestor.historico.get(str(nota['nota']), "-")
            if isinstance(dados_hist, dict):
                nota['ultimo_envio'] = dados_hist.get("data", "-")
                nota['caminho_evidencia'] = dados_hist.get("caminho", "Não disponível")
                nota['usuario_envio'] = dados_hist.get("usuario", "Desconhecido")
        return {"status": "sucesso", "msg": "Carregado do cache rápido"}

    # Se a lista estiver vazia (primeira vez que abre o sistema no dia), ele lê o Excel!
    gestor.todas_as_notas.clear()
    arquivos = [PATH_2025, PATH_2026]
    hoje = datetime.now()
    lista_final = []

    for caminho in arquivos:
        if os.path.exists(caminho):
            try:
                # 🚀 USANDO O CARREGAMENTO NINJA LOCAL AQUI!
                df = ler_excel_ninja(caminho, usecols=[1, 3, 4, 5, 8, 11, 14, 15], engine='openpyxl')
                
                if df.empty: continue
                    
                df.columns = ['empresa', 'data', 'nota', 'competencia', 'cliente', 'situacao', 'valor', 'pagamento']
                df['linha_excel'] = df.index + 2
                df['arquivo_origem'] = "2025" if "2025" in caminho else "2026"
                df = df[df['situacao'].str.upper() == 'NORMAL']
                df = df[df['pagamento'].isna() | (df['pagamento'].astype(str).str.strip() == '-')]
                lista_final.append(df)
            except: pass

    if not lista_final: return {"status": "sucesso", "dados": [], "total_atraso": "R$ 0,00", "total_geral": "R$ 0,00"}

    df_total = pd.concat(lista_final)
    for _, row in df_total.iterrows():
        try:
            dt_venc = row['data']
            if not isinstance(dt_venc, datetime): dt_venc = pd.to_datetime(dt_venc)
            atraso = (hoje - dt_venc).days
            
            if atraso >= 20:
                parecer = gestor.definir_parecer(atraso)
                valor_num = gestor.extrair_float(row['valor'])
                
                try: nota_str = str(int(float(row['nota']))).strip()
                except: nota_str = str(row['nota']).strip()
                
                comp_val = str(row['competencia']).strip()
                if comp_val.lower() == 'nan' or pd.isna(row['competencia']): comp_val = ""
                
                dados_hist = gestor.historico.get(nota_str, "-")
                if isinstance(dados_hist, dict):
                    ultimo_envio = dados_hist.get("data", "-")
                    caminho_evidencia = dados_hist.get("caminho", "Não disponível")
                    usuario_envio = dados_hist.get("usuario", "Desconhecido") 
                else:
                    ultimo_envio = dados_hist 
                    caminho_evidencia = "Salvo em versão anterior"
                    usuario_envio = "Desconhecido" 

                gestor.todas_as_notas.append({
                    "emissao": dt_venc.strftime("%d/%m/%Y"),
                    "nota": nota_str,
                    "cliente": str(row['cliente']).strip(),
                    "empresa": str(row['empresa']).strip().upper(),
                    "valor_str": gestor.formatar_moeda(valor_num),
                    "valor_num": valor_num,
                    "dias": atraso,
                    "parecer": parecer,
                    "ultimo_envio": ultimo_envio,
                    "usuario_envio": usuario_envio, 
                    "linha_excel": row['linha_excel'],
                    "arquivo": row['arquivo_origem'],
                    "caminho_evidencia": caminho_evidencia,
                    "competencia": comp_val
                })
        except: continue

    return {"status": "sucesso"}

@eel.expose
def filtrar_dados(termo, empresa_filtro):
    termo = termo.upper()
    total_30_dias = 0.0
    total_geral = 0.0
    notas_filtradas = []

    # 🕵️‍♂️ IDENTIFICAÇÃO DO PERFIL LOGADO
    setor = getattr(gestor, 'funcao_logada', '').upper()
    usuario = getattr(gestor, 'usuario_logado', '').upper()
    
    # Define quem é quem no jogo
    is_juridico = "JURIDICO" in setor or "JURÍDICO" in setor or "NATALIA" in usuario
    is_licitacao = "LICITAÇÃO" in setor or "LICITACAO" in setor or "RENATO" in usuario or "OTNIEL" in usuario

    for nota in gestor.todas_as_notas:
        dias_emissao = nota.get('dias', 0)
        dias_atraso = dias_emissao - 30 # Atraso real após os 30 dias de prazo
        
        # 🚨 REGRA NATÁLIA (JURÍDICO): Só vê do LARANJA em diante (> 30 dias de atraso real)
        if is_juridico and dias_atraso <= 30:
            continue
            
        # 🚨 REGRA RENATO/OTNIEL (LICITAÇÃO): Só vê do VERMELHA em diante (> 60 dias de atraso real)
        if is_licitacao and dias_atraso <= 60:
            continue

        # Filtro de Empresa
        if empresa_filtro != "TODAS" and empresa_filtro not in nota['empresa']: 
            continue
            
        # Filtro de Busca (Cliente, Nota ou Faixa)
        if termo in str(nota.get('cliente', '')).upper() or termo in str(nota.get('nota', '')).upper() or termo in str(nota.get('parecer', '')).upper():
            notas_filtradas.append(nota)

    # --- PROCESSAMENTO DOS DADOS FILTRADOS ---
    notas_filtradas.sort(key=lambda x: (x['cliente'], -x['dias']))
    
    subtotais = {}
    for nota in notas_filtradas:
        cli = nota['cliente']
        subtotais[cli] = subtotais.get(cli, 0.0) + nota['valor_num']
        total_geral += nota['valor_num']
        if nota['dias'] >= 30: 
            total_30_dias += nota['valor_num']

    dados_finais = []
    cliente_atual = None
    for nota in notas_filtradas:
        cli = nota['cliente']
        if cli != cliente_atual:
            str_subtotal = gestor.formatar_moeda(subtotais[cli])
            val_subtotal_num = subtotais[cli]
            cliente_atual = cli
        else:
            str_subtotal = ""
            val_subtotal_num = None
        
        n = nota.copy()
        n['total_contrato_str'] = str_subtotal
        n['total_contrato_num'] = val_subtotal_num
        dados_finais.append(n)
    
    gestor.ultima_filtragem = dados_finais
    return {
        "dados": dados_finais,
        "total_atraso": gestor.formatar_moeda(total_30_dias),
        "total_geral": gestor.formatar_moeda(total_geral)
    }

@eel.expose
def exportar_relatorio():
    if not gestor.ultima_filtragem: return {"status": "erro", "msg": "Não há dados para exportar."}
    
    root = tk.Tk(); root.withdraw(); root.attributes('-topmost', True)
    caminho = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], title="Salvar", initialfile=f"Relatorio_{datetime.now().strftime('%d-%m-%Y')}.xlsx")
    if not caminho: return {"status": "cancelado"}

    dados_exportar = []
    total_g = 0.0
    for d in gestor.ultima_filtragem:
        total_g += d['valor_num']
        dados_exportar.append({
            "Empresa": d["empresa"], 
            "Cliente": d["cliente"], 
            "Nota Fiscal": d["nota"],
            "Competência": d["competencia"], 
            "Emissão": d["emissao"], 
            "Valor": d["valor_num"],
            "Total Contrato": d["total_contrato_num"], 
            "Dias Atraso": d["dias"],
            "Status": d["parecer"], 
            "Ult. Envio": f"{d['ultimo_envio']} (por {d.get('usuario_envio', 'Desconhecido')})", # <-- MUDANÇA AQUI
            "Local Evidência": d["caminho_evidencia"]
        })
    dados_exportar.append({"Empresa": "", "Cliente": "", "Nota Fiscal": "", "Competência": "", "Emissão": "TOTAL GERAL:", "Valor": total_g, "Total Contrato": None, "Dias Atraso": "", "Status": "", "Ult. Envio": "", "Local Evidência": ""})

    df_export = pd.DataFrame(dados_exportar)
    try:
        with pd.ExcelWriter(caminho, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False, sheet_name='Cobranças')
            ws = writer.sheets['Cobranças']
            for row in ws.iter_rows(min_row=2, max_col=7, max_row=ws.max_row):
                if isinstance(row[5].value, (int, float)): row[5].number_format = '#,##0.00'
                if isinstance(row[6].value, (int, float)): row[6].number_format = '#,##0.00'
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try: max_length = max(max_length, len(str(cell.value)))
                    except: pass
                ws.column_dimensions[column].width = min(max_length + 2, 45)
        os.startfile(caminho)
        return {"status": "sucesso"}
    except Exception as e: return {"status": "erro", "msg": str(e)}

@eel.expose
def enviar_email_backend(notas_selecionadas, conta_selecionada, escolha_90_dias):
    if not notas_selecionadas: return {"status": "erro", "msg": "Nenhuma nota."}
    
    cliente_ref = notas_selecionadas[0]['cliente']
    empresa_ref = notas_selecionadas[0]['empresa']
    valor_total = sum(n['valor_num'] for n in notas_selecionadas)
    
    # =================================================================
    # 🕵️‍♂️ CÉREBRO DE CÓPIA (CC) - BUSCA NO BANCO LOCAL (vinculos_operacao)
    # =================================================================
    lista_cc = []
    
    # 1. REGRA DE OURO: O Daison está em todas!
    lista_cc.append("gerenciaoperacional@haggltda.com.br") 

    # 2. Dicionário de E-mails dos Supervisores (Para traduzir Nome -> E-mail)
    emails_supervisores = {
        "AFRANIO": "adm5@haggltda.com.br",
        "CASSIO": "cassioduarte.operacional@haggltda.com.br",
        "ISMAEL": "adm9@haggltda.com.br",
        "GUSTAVO B": "gustavobarcelos.operacional@haggltda.com.br",
        "GUSTAVO": "gustavobarcelos.operacional@haggltda.com.br"
    }

    try:
        # Pega a primeira palavra do cliente para buscar no banco (Ex: BENTO)
        termo_busca = str(cliente_ref).split('-')[0].strip()
        
        c = gestor.conn.cursor()
        c.execute("""SELECT analista_email, supervisor_nome 
                     FROM vinculos_operacao 
                     WHERE contrato_sistema LIKE ?""", (f"%{termo_busca}%",))
        resultado = c.fetchone()
        
        if resultado:
            email_analista = resultado[0]
            nome_supervisor = str(resultado[1]).upper()
            
            # Adiciona o Analista se o e-mail for válido
            if email_analista and "@" in str(email_analista):
                lista_cc.append(email_analista)
                print(f"🎯 Analista {email_analista} adicionado ao CC.")

            # Busca o e-mail do supervisor no nosso dicionário
            for chave_sup in emails_supervisores.keys():
                if chave_sup in nome_supervisor:
                    lista_cc.append(emails_supervisores[chave_sup])
                    print(f"👮 Supervisor {nome_supervisor} adicionado ao CC.")
                    break
        else:
            print(f"⚠️ Vínculo não encontrado para {termo_busca}. Seguindo apenas com CC padrão.")

    except Exception as e:
        print(f"❌ Erro ao buscar vínculos no DB Local: {e}")

    # Remove duplicados e junta com ponto e vírgula para o Outlook
    cc_final = "; ".join(list(set(lista_cc)))
    # =================================================================

    tipo_cobranca_final = ""
    maior_sev = -1 
    
    niveis = {
        "AVISO AMIGÁVEL (PRÉ-VENCIMENTO)": 0,
        "FAIXA VERDE (ATÉ 15 DIAS)": 1, 
        "FAIXA AMARELA (16 A 30 DIAS)": 2, 
        "FAIXA LARANJA (31 A 60 DIAS)": 3, 
        "FAIXA VERMELHA (61 A 90 DIAS)": 4, 
        "FAIXA ROXA (91 A 120 DIAS)": 5,
        "FAIXA PRETA (+120 DIAS)": 6
    }
    
    for n in notas_selecionadas:
        if n['cliente'] != cliente_ref: return {"status": "erro", "msg": "Selecione notas do MESMO CLIENTE."}
        parecer_recebido = str(n['parecer']).strip().upper()
        parecer_oficial = parecer_recebido 
        
        for chave_oficial in niveis.keys():
            if parecer_recebido in chave_oficial:
                parecer_oficial = chave_oficial
                break
        
        niv = niveis.get(parecer_oficial, -1) 
        if niv > maior_sev:
            maior_sev = niv
            tipo_cobranca_final = parecer_oficial

    qtd_notas = len(notas_selecionadas)
    lista_html = "<ul>"
    for n in notas_selecionadas:
        comp = f" - Ref: {n['competencia']}" if n['competencia'] else ""
        lista_html += f"<li>Nota <b>{n['nota']}</b> (Emissão: {n['emissao']}{comp}) - {n['valor_str']}</li>"
    lista_html += "</ul>"

    dias_max = max(n['dias'] for n in notas_selecionadas)
    data_lim = (datetime.now() + timedelta(days=5)).strftime("%d/%m/%Y")
    rz_map = {"SN": "SN Serviços de Limpeza e Zeladoria Predial Ltda", "HAGG": "Nascimento Serviços de Limpeza Ltda", "NH": "NH Prestação de Serviços Ltda", "CANAÃ": "INSTITUTO DE ENSINO CANAA", "CANAA": "INSTITUTO DE ENSINO CANAA"}
    ass_emp = rz_map.get(empresa_ref, f"{empresa_ref} - GRUPO NASCIMENTO")

    # ... (código anterior da sua função, incluindo a formatação das Variáveis Mágicas) ...
    t_nota = "à Nota Fiscal" if qtd_notas == 1 else "às Notas Fiscais"
    t_da_nota = "da Nota Fiscal" if qtd_notas == 1 else "das Notas Fiscais"
    t_num = f" nº {notas_selecionadas[0]['nota']}" if qtd_notas == 1 else " detalhadas abaixo"
    comp_ref = notas_selecionadas[0]['competencia']
    t_comp = f" relativas aos serviços prestados no período de {comp_ref}," if qtd_notas == 1 and comp_ref else ""
    t_emi = f"emitida pela empresa <b>{ass_emp}</b>," if qtd_notas == 1 else f"emitidas pela empresa <b>{ass_emp}</b>,"
    t_venc = f" cujo vencimento está programado para o dia {notas_selecionadas[0]['emissao']}," if qtd_notas == 1 else ""
    t_com_venc = f" com emissão em {notas_selecionadas[0]['emissao']}," if qtd_notas == 1 else ""

    # =================================================================
    # ⚙️ MÁGICA DOS TEMPLATES DINÂMICOS NO BANCO DE DADOS (SELETOR)
    # AQUI ENTRA A MODIFICAÇÃO QUE VOCÊ PERGUNTOU! 👇
    # =================================================================
    if "FAIXA LARANJA" in tipo_cobranca_final:
        if str(escolha_90_dias) == "1":
            tipo_cobranca_final = "FAIXA LARANJA - ANEXO I (ADMINISTRATIVO)"
        else:
            tipo_cobranca_final = "FAIXA LARANJA - ANEXO II (JURÍDICO)"
            
    elif "FAIXA VERMELHA" in tipo_cobranca_final:
        if str(escolha_90_dias) == "1":
            tipo_cobranca_final = "FAIXA VERMELHA - ANEXO III (ADMINISTRATIVO)"
        else:
            tipo_cobranca_final = "FAIXA VERMELHA - ANEXO IV (JURÍDICO)"

    elif "FAIXA ROXA" in tipo_cobranca_final:
        if str(escolha_90_dias) == "1":
            tipo_cobranca_final = "FAIXA ROXA - MINUTA JURÍDICO (91 A 120 DIAS)"
        else:
            tipo_cobranca_final = "FAIXA ROXA - PARECER PREVENTIVO (91 A 120 DIAS)"

    # =================================================================
    # AGORA O SISTEMA VAI PRO BANCO BUSCAR O NOME CORRETO QUE FOI ESCOLHIDO ACIMA!
    # =================================================================
    c = gestor.conn.cursor()
    c.execute("SELECT assunto, corpo FROM templates_email WHERE nome_identificador = ?", (tipo_cobranca_final,))
    res_temp = c.fetchone()

    if not res_temp:
        return {"status": "erro", "msg": f"O modelo de e-mail '{tipo_cobranca_final}' não foi encontrado no Banco de Dados."}

    substituicoes = {
        "{cliente}": cliente_ref, "{lista_html}": lista_html, "{valor_total}": gestor.formatar_moeda(valor_total),
        "{dias_max}": str(dias_max), "{data_lim}": data_lim, "{ass_emp}": ass_emp, "{t_nota}": t_nota,
        "{t_da_nota}": t_da_nota, "{t_num}": t_num, "{t_comp}": t_comp, "{t_emi}": t_emi, "{t_venc}": t_venc, "{t_com_venc}": t_com_venc
    }

    ass = res_temp[0]
    corpo = res_temp[1]
    for tag, valor_real in substituicoes.items():
        ass = ass.replace(tag, valor_real)
        corpo = corpo.replace(tag, valor_real)
    
    # ... (o resto da função com a geração do PDF e Outlook continua igualzinho ao que você mandou!) ...

    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        c_enc = None
        
        # Identifica a conta que será usada
        if conta_selecionada and conta_selecionada != "Conta Padrão":
            for acc in outlook.Session.Accounts:
                if acc.DisplayName == conta_selecionada or acc.SmtpAddress == conta_selecionada:
                    c_enc = acc
                    break
        
        if c_enc:
            try: mail.SendUsingAccount = c_enc
            except: pass
            try: mail._oleobj_.Invoke(*(64209, 0, 8, 0, c_enc))
            except: pass

        mail.Subject = ass
        
        if cc_final:
            mail.CC = cc_final
            print(f"📧 Outlook configurado com CC: {cc_final}")

        mail.Display()
        ass_out = mail.HTMLBody 
        mail.HTMLBody = f"<div style='font-family:Calibri; font-size:11pt;'>{corpo}</div>{ass_out}" 
        
        # ======================================================================
        # 🕵️‍♂️ CAPTURA O E-MAIL REAL PARA A EVIDÊNCIA (FIM DO "CONTA PADRÃO")
        # ======================================================================
        try:
            if c_enc:
                email_real_remetente = c_enc.SmtpAddress
            else:
                email_real_remetente = outlook.Session.Accounts.Item(1).SmtpAddress
        except:
            email_real_remetente = conta_selecionada 
            
        print(f"📄 Remetente real identificado: {email_real_remetente}")

        c_ev = corpo + "<br><br>Atenciosamente,"
        nms = ", ".join([n['nota'] for n in notas_selecionadas])
        cli_l = cliente_ref.replace("/", "-").replace("\\", "-")
        
        # 🚀 CORREÇÃO DO BUG DE SOBREPOSIÇÃO: Adicionando Hora/Minuto, Tipo e Usuário
        dt_env = datetime.now().strftime("%d.%m.%Y %Hh%M") # Ex: 27.04.2026 10h18
        usuario_up = str(gestor.usuario_logado).upper()
        
        # Cria um nome riquíssimo que nunca vai se repetir!
        nm_arq = f"{cli_l} - {dt_env} - NFs {nms} - {tipo_cobranca_final} ({usuario_up})"[:200]
        
        # Limpa qualquer caractere que o Windows proíba em nomes de arquivos
        nm_arq = nm_arq.replace(":", "h").replace("/", "-").replace("\\", "-")

        # Chama a geração do PDF
        cam = gestor.gerar_evidencia_pdf(ass, c_ev, email_real_remetente, cliente_ref, nm_arq, empresa_ref)
        
        if cam and cam != "ERRO_ABERTO":
            for n in notas_selecionadas: 
                gestor.salvar_historico_envio(n['nota'], cam)
                try:
                    gestor.conn.execute("INSERT INTO log_atividades (cliente, nota_fiscal, data_hora, usuario, acao) VALUES (?, ?, ?, ?, ?)",
                                        (cliente_ref, n['nota'], datetime.now().strftime("%d/%m/%Y %H:%M:%S"), gestor.usuario_logado, f"E-mail: {tipo_cobranca_final}"))
                except: pass
            gestor.conn.commit()
            gestor.historico = gestor.carregar_historico()
            carregar_dados_reais()  
            
        return {"status": "sucesso"}
        
    except Exception as e: 
        return {"status": "erro", "msg": str(e)}

@eel.expose
def substituir_evidencia_python(nota_fiscal, cliente, empresa):
    import tkinter as tk
    from tkinter import filedialog
    import shutil
    import os
    
    # 1. Abre a janelinha do Windows para o Ruan escolher o novo arquivo
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    arquivo_novo = filedialog.askopenfilename(
        title=f"Selecione a nova evidência para a Nota {nota_fiscal}", 
        filetypes=[("Arquivos PDF", "*.pdf"), ("Imagens", "*.png;*.jpg;*.jpeg"), ("Todos", "*.*")]
    )
    
    if not arquivo_novo: return {"status": "cancelado"} # Se ele fechar a janela

    try:
        # 2. Descobre qual é a pasta oficial desse cliente lá no disco S:
        pasta_destino = gestor.encontrar_ou_criar_pasta_cliente(cliente, empresa)
        
        # 3. Cria um nome padronizado para a nova evidência
        extensao = os.path.splitext(arquivo_novo)[1]
        data_hoje = datetime.now().strftime("%d.%m.%Y")
        nome_arquivo = f"{cliente.replace('/', '-')} - {data_hoje} - NF {nota_fiscal} - EVIDENCIA MANUAL{extensao}"
        caminho_final = os.path.join(pasta_destino, nome_arquivo)
        
        # 4. Copia o arquivo do PC do Ruan para o Disco S:
        shutil.copy2(arquivo_novo, caminho_final)
        
        # 5. Atualiza o Banco de Dados (Substituindo o caminho velho pelo novo)
        gestor.salvar_historico_envio(nota_fiscal, caminho_final)
        
        # 6. Atualiza a memória para a tela não precisar recarregar o excel
        gestor.historico = gestor.carregar_historico()
        carregar_dados_reais(forcar=False)
        
        return {"status": "sucesso", "novo_caminho": caminho_final}
    except Exception as e:
        return {"status": "erro", "msg": str(e)}
    
# ==============================================================================
# ⚙️ GERADOR DE DOSSIÊ
# ==============================================================================
@eel.expose
def gerar_dossie_zip_python(cliente):
    import tkinter as tk
    from tkinter import filedialog
    import os
    import zipfile
    
    try:
        # 1. Busca no Banco de Dados todos os caminhos de arquivos desse cliente
        # (Nós pegamos tudo que está no histórico onde a nota pertence a esse cliente)
        c = gestor.conn.cursor()
        c.execute("SELECT nota_fiscal, caminho_pdf FROM historico WHERE caminho_pdf IS NOT NULL")
        resultados = c.fetchall()
        
        # Filtra as notas que estão na memória atual e pertencem a este cliente
        notas_do_cliente = [str(n['nota']) for n in gestor.todas_as_notas if n['cliente'] == cliente]
        
        arquivos_para_zipar = []
        for linha in resultados:
            nota_banco = str(linha[0])
            caminho = linha[1]
            if nota_banco in notas_do_cliente and os.path.exists(caminho):
                arquivos_para_zipar.append(caminho)
                
        if not arquivos_para_zipar:
            return {"status": "erro", "msg": "Nenhuma evidência física encontrada para este cliente."}

        # 2. Pergunta onde o usuário quer salvar o Dossiê (ZIP)
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        nome_sugerido = f"DOSSIE_COBRANCA_{cliente.replace('/', '_')}.zip"
        caminho_zip = filedialog.asksaveasfilename(
            title="Salvar Dossiê Jurídico", 
            initialfile=nome_sugerido,
            defaultextension=".zip", 
            filetypes=[("Arquivo ZIP", "*.zip")]
        )
        
        if not caminho_zip: return {"status": "cancelado"}

        # 3. Cria o arquivo ZIP e empacota tudo lá dentro!
        with zipfile.ZipFile(caminho_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for arq in arquivos_para_zipar:
                nome_base = os.path.basename(arq)
                zipf.write(arq, arcname=nome_base) # arcname evita que o zip crie pastas S:/Gestao... lá dentro
                
        # Abre a pasta onde o ZIP foi salvo para o usuário já ver
        os.startfile(os.path.dirname(caminho_zip))
        
        return {"status": "sucesso"}
    except Exception as e:
        return {"status": "erro", "msg": str(e)}

@eel.expose
def abrir_arquivo_evidencia(caminho_arquivo):
    import os
    try:
        # Se o arquivo existir, o Python manda o Windows abrir (vai abrir no Leitor de PDF padrão)
        if os.path.exists(caminho_arquivo):
            os.startfile(caminho_arquivo)
            return {"status": "sucesso"}
        else:
            return {"status": "erro", "msg": "O arquivo não foi encontrado na rede. Ele pode ter sido movido."}
    except Exception as e:
        return {"status": "erro", "msg": str(e)}

# ==============================================================================
# ⚙️ MÓDULO DE ADMINISTRAÇÃO DE USUÁRIOS
# ==============================================================================

@eel.expose
def obter_perfil_usuario():
    nome = gestor.usuario_logado
    setor = gestor.funcao_logada
    
    # LIBERAÇÃO: Se for Ruan OU se o setor for FINANCEIRO ou CONTROLADORIA
    permitidos = ["FINANCEIRO", "CONTROLADORIA"]
    is_admin = True if (nome.lower() == 'ruan' or setor in permitidos) else False
    
    return {"nome": nome, "is_admin": is_admin}

@eel.expose
def alterar_senha_python(nova_senha):
    """Altera a senha do usuário que está logado no momento"""
    try:
        c = gestor.conn.cursor()
        user_atual = gestor.usuario_logado.lower()
        c.execute("UPDATE usuarios SET senha = ? WHERE usuario = ?", (nova_senha, user_atual))
        gestor.conn.commit()
        return {"status": "sucesso"}
    except Exception as e:
        return {"status": "erro", "msg": str(e)}
    
@eel.expose
def listar_usuarios():
    """Puxa todo mundo do banco para montar a tabela do Painel"""
    c = gestor.conn.cursor()
    c.execute("SELECT id, usuario, funcao, telefone FROM usuarios ORDER BY usuario ASC")
    resultados = c.fetchall()
    lista = []
    for r in resultados:
        lista.append({
            "id": r[0], 
            "usuario": str(r[1]).capitalize(), 
            "funcao": r[2] if r[2] else "", 
            "telefone": r[3] if r[3] else ""
        })
    return lista

@eel.expose
def salvar_usuario(id_user, usuario, senha, funcao, telefone):
    """Cria um usuário novo ou edita um existente"""
    try:
        c = gestor.conn.cursor()
        user_limpo = usuario.lower().strip()
        
        if id_user: 
            # É uma EDIÇÃO
            if senha.strip() != "": # Se digitou senha nova, atualiza tudo
                c.execute("UPDATE usuarios SET usuario=?, senha=?, funcao=?, telefone=? WHERE id=?", 
                          (user_limpo, senha, funcao, telefone, id_user))
            else: # Se deixou a senha em branco, altera os dados mas mantém a senha antiga
                c.execute("UPDATE usuarios SET usuario=?, funcao=?, telefone=? WHERE id=?", 
                          (user_limpo, funcao, telefone, id_user))
        else: 
            # É um NOVO USUÁRIO
            if not senha: return {"status": "erro", "msg": "A senha é obrigatória para novos usuários!"}
            c.execute("INSERT INTO usuarios (usuario, senha, funcao, telefone) VALUES (?, ?, ?, ?)", 
                      (user_limpo, senha, funcao, telefone))
        
        gestor.conn.commit()
        return {"status": "sucesso"}
    except Exception as e:
        return {"status": "erro", "msg": "Erro: Esse nome de usuário já existe ou o banco falhou."}

@eel.expose
def excluir_usuario(id_user):
    """Manda o usuário embora do banco"""
    try:
        c = gestor.conn.cursor()
        c.execute("DELETE FROM usuarios WHERE id=?", (id_user,))
        gestor.conn.commit()
        return {"status": "sucesso"}
    except Exception as e:
        return {"status": "erro", "msg": str(e)}
    
# ==============================================================================
# ⚙️ EDIÇÃO DE PARECERES
# ==============================================================================

@eel.expose
def listar_templates_email():
        try:
            c = gestor.conn.cursor()
            c.execute("SELECT id, nome_identificador, assunto, corpo, anexo, responsavel FROM templates_email")
            linhas = c.fetchall()

            lista = []
            for r in linhas:
                lista.append({
                    "id": r[0], "nome": r[1], "assunto": r[2], 
                    "corpo": r[3], "anexo": r[4] if r[4] else "", 
                    "responsavel": r[5] if r[5] else ""
                })

            # 🧠 MÁGICA DA ORDENAÇÃO (Dando peso para as palavras-chave)
            def peso_ordem(t):
                n = str(t['nome']).upper()
                if "AMIGÁVEL" in n: return 1
                if "VERDE" in n: return 2
                if "AMARELA" in n: return 3
                
                # Agrupa a Laranja
                if "LARANJA" in n: 
                    return 5 if "II" in n else 4
                
                # Agrupa a Vermelha
                if "VERMELHA" in n: 
                    return 7 if "IV" in n or "SUP" in n else 6
                
                # Agrupa a Roxa
                if "ROXA" in n: 
                    return 9 if "PARECER" in n else 8
                    
                if "PRETA" in n: return 10
                if "WHATSAPP" in n: return 11
                
                return 99 # Desconhecidos vão pro final da fila

            # Ordena a lista usando a função de peso antes de mandar pro Javascript
            lista_ordenada = sorted(lista, key=peso_ordem)
            return lista_ordenada
            
        except Exception as e:
            print(f"Erro ao listar templates: {e}")
            return []

@eel.expose
def salvar_template_email(id_temp, nome, assunto, corpo, anexo="", responsavel=""):
        try:
            # Pega a "mão" (cursor) a partir do seu "caderno" oficial (conn)
            cursor = gestor.conn.cursor()
            
            # Se o ID vier vazio, significa que é um NOVO MODELO (INSERT)
            if not id_temp or id_temp == "":
                cursor.execute("""
                    INSERT INTO templates_email (nome_identificador, assunto, corpo, anexo, responsavel) 
                    VALUES (?, ?, ?, ?, ?)
                """, (nome, assunto, corpo, anexo, responsavel))
            
            # Se vier com ID, é a EDIÇÃO de um modelo que já existe (UPDATE)
            else:
                cursor.execute("""
                    UPDATE templates_email 
                    SET nome_identificador = ?, assunto = ?, corpo = ?, anexo = ?, responsavel = ? 
                    WHERE id = ?
                """, (nome, assunto, corpo, anexo, responsavel, id_temp))
            
            # Salva de verdade no banco de dados!
            gestor.conn.commit()
            
            return {"status": "sucesso"}
        except Exception as e:
            print(f"Erro ao salvar template: {e}")
            return {"status": "erro", "msg": str(e)}
        

@eel.expose
def obter_texto_whatsapp():
    try:
        c = gestor.conn.cursor()
        c.execute("SELECT corpo FROM templates_email WHERE nome_identificador = 'WHATSAPP - MENSAGEM PADRÃO'")
        res = c.fetchone()
        return res[0] if res else "Texto do WhatsApp não encontrado."
    except Exception as e:
        return str(e)
    
@eel.expose
def gerar_preview_whats_python(notas_selecionadas):
    try:
        if not notas_selecionadas: return {"status": "erro", "msg": "Nenhuma nota"}
        
        # 👇 AQUI ESTAVA O ERRO: Tabela e colunas corrigidas! 👇
        c = gestor.conn.cursor()
        c.execute("SELECT corpo FROM templates_email WHERE nome_identificador = 'WHATSAPP - MENSAGEM PADRÃO'")
        res = c.fetchone()
        
        texto_base = res[0] if res else "Olá! Temos notas pendentes do contrato {cliente}."

        # Pega os dados da primeira nota para as variáveis gerais
        cliente_ref = notas_selecionadas[0]['cliente']
        valor_total = sum(n['valor_num'] for n in notas_selecionadas)
        
        # Faz as substituições mágicas
        texto_pronto = texto_base.replace('{cliente}', f"*{cliente_ref}*")
        texto_pronto = texto_pronto.replace('{valor_total}', f"R$ {valor_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        
        # Monta a listinha de notas
        lista_notas = ""
        for n in notas_selecionadas:
            lista_notas += f"• NF {n['nota']} - Venc: {n['emissao']} - R$ {n['valor_str']}\n"
            
        texto_pronto = texto_pronto.replace('{lista_html}', lista_notas)
        
        # Limpa o HTML que pode ter vindo do editor para virar texto de WhatsApp
        texto_pronto = texto_pronto.replace('<br>', '\n').replace('<b>', '*').replace('</b>', '*')

        return {"status": "sucesso", "texto": texto_pronto}
    except Exception as e:
        print(f"Erro no Python: {e}") # Isso vai fofocar no terminal se der mais algum B.O.
        return {"status": "erro", "msg": str(e)}
    
@eel.expose
def abrir_pasta_cliente_python(cliente, empresa):
    import subprocess
    import os
    try:
        pasta = gestor.encontrar_ou_criar_pasta_cliente(cliente, empresa)
        
        if os.path.exists(pasta):
            # 🚀 MÁGICA: Chama o Explorer diretamente. 
            # Isso costuma forçar o Windows a dar foco na nova janela.
            subprocess.Popen(f'explorer "{os.path.normpath(pasta)}"')
            return {"status": "sucesso"}
        else:
            return {"status": "erro", "msg": "A pasta ainda não existe."}
    except Exception as e:
        return {"status": "erro", "msg": str(e)}
    
@eel.expose
def verificar_cobranca_duplicada_python(notas_selecionadas, faixa):
    try:
        c = gestor.conn.cursor()
        
        # Pega só o nome principal da faixa para garantir a busca (Ex: "FAIXA ROXA")
        termo_busca = str(faixa).split('(')[0].strip() 
        
        for n in notas_selecionadas:
            nota_num = str(n['nota'])
            
            # 🕵️‍♂️ MÁGICA: Procura na Timeline se essa nota já teve um e-mail dessa faixa
            query = "SELECT data_hora, usuario FROM log_atividades WHERE nota_fiscal = ? AND acao LIKE ? ORDER BY id DESC"
            c.execute(query, (nota_num, f"%E-mail: %{termo_busca}%"))
            resultado = c.fetchone()
            
            if resultado:
                data_hora = resultado[0]
                usuario = resultado[1]
                return {
                    "duplicado": True, 
                    "msg": f"A Nota Fiscal {nota_num} já teve um e-mail de '{termo_busca}' gerado em {data_hora} por {usuario}."
                }
                
        return {"duplicado": False}
    except Exception as e:
        print(f"Erro ao verificar duplicidade: {e}")
        return {"duplicado": False}
    
@eel.expose
def importar_planilha_relacao_python():
    import pandas as pd
    import os
    
    caminho = r"C:\Users\João Victor Peretti\Documents\FINANCEIRO\COBRANÇA v2\RELAÇÃO OPERACIONAL.xlsx"
    
    if not os.path.exists(caminho):
        return {"status": "erro", "msg": f"Arquivo não encontrado: {caminho}"}

    # 🕵️‍♂️ TRADUTOR 1: CONTRATOS (DE -> PARA) - O DICIONÁRIO DEFINITIVO
    dicionario_de_para = {
        "BENTO HIG/AUX ADM": "BENTO GONÇALVES",
        "PREF POA SAUDE/SAMU": "SAMU",
        "PREF CAXIAS": "CAXIAS DO SUL",
        "SÃO CAMILO": "HOSPITAL SÃO CAMILO",
        "CAM RG LIMPEZA": "CAMARA RIO GRANDE LIMP E COPA 001.2023",
        "CAM RG PORTARIA": "CAMARA DE RIO GRANDE - PORTARIA 002.2023",
        "CANAÃ": "ESCOLA DE ENSINO CANAA EIRELI - ME",
        "DEMAE": "DMAE 8950",
        "EMBRAPA CANOINHAS": "CANOINHA EMBRAPA 47.2024",
        "EMBRAPA PELOTAS": "EMBRAPA 93.2021",
        "FUNARBE": "FUNARBE PELOTAS - 58164/2025",
        "FURG HU": "HU FURG RIO GRANDE 006.2023",
        "FURG JARDINAGEM": "FURG JARDINAGEM 049.2022",
        "FURG PORTARIA": "FURG PORTARIA 55.2023",
        "HCPA": "HCPA MENSAGEIROS", 
        "HUSM": "LAVANDERIA HUSM 020.2021",
        "IPAM": "IPAM CAXIAS 012.2022",
        "IPASEM": "IPASEM - NH 13.2022",
        "PENHA": "PENHA LIMPEZA 039.2025",
        "PREF POA CULTURA": "SEC. DA CULTURA - POA 88123.2024",
        "PREF POA SAUDE DVS": "PREF POA SMS RECEPÇÃO - 98672/2025",
        "PREF POA SAUDE HPS": "PREF POA SMS RECEPÇÃO - 98672/2025",
        "PREF POA SAUDE MATERNO": "PREF POA SMS RECEPÇÃO - 98672/2025",
        "SALTO DO JACUI": "PM SALTO DO JACUI 722.2021",
        "SEMAE": "SEMAE 3038.2020",
        "TJ COORDENADORA": "TJ RS 023.2025",
        "TJ SUPERVISORES": "TJ RS 023.2025",
        "TRIUNFO MOT/OP MAQ": "TRIUNFO", 
        "TRIUNFO VIGIAS": "PM TRIUNFO VIGIAS 33.2024",
        "UFFS CERRO LARGO": "UFFS CERRO LARGO 041/2021",
        "UFFS CHAPECO": "UFFS CHAPECO 041/2021",
        "UFFS ERECHIM": "UFFS ERECHIM 041/2021",
        "UFFS LARENJEIRAS": "UFFS LARANJEIRAS 041/2021",
        "UFFS PASSO FUNDO": "UFFS PASSO FUNDO 041/2021",
        "UFFS REALEZA": "UFFS REALEZA 041/2021",
        "UFRGS CARREGADORES CENTRO": "UFRGS CARREGADORES 095.2024",
        "UFRGS CARREGADORES SAUDE": "UFRGS CARREGADORES 095.2024",
        "UFRGS CARREGADORES VALE": "UFRGS CARREGADORES 095.2024",
        "UFRGS HCVET": "UFRGS PORTO ALEGRE 020.2022",
        "UFRGS JARDINAGEM CENTRO": "UFRGS JARDINAGEM 062/2025",
        "UFRGS JARDINAGEM ESEFID": "UFRGS JARDINAGEM 062/2025",
        "UFRGS JARDINAGEM LITORAL": "UFRGS JARDINAGEM 062/2025",
        "UFRGS JARDINAGEM SAUDE": "UFRGS JARDINAGEM 062/2025",
        "UFRGS JARDINAGEM VALE AGRONOMIA": "UFRGS JARDINAGEM 062/2025",
        "UFRGS JARDINAGEM VALE": "UFRGS JARDINAGEM 062/2025",
        "UFRGS LIMP GERAL AGRONOMIA": "UFRGS LIMPEZA GERAL 047.2022",
        "UFRGS LIMP GERAL CENTRO 1": "UFRGS LIMPEZA GERAL 047.2022",
        "UFRGS LIMP GERAL CENTRO 2": "UFRGS LIMPEZA GERAL 047.2022",
        "UFRGS LIMP GERAL ESEFID": "UFRGS LIMPEZA GERAL 047.2022",
        "UFRGS LIMP GERAL FERISTA": "UFRGS LIMPEZA GERAL 047.2022",
        "UFRGS LIMP GERAL SAUDE": "UFRGS LIMPEZA GERAL 047.2022",
        "UFRGS LIMP GERAL VALE CENTRO 2": "UFRGS LIMPEZA GERAL 047.2022",
        "UFRGS LIMP GERAL VALE IPH": "UFRGS LIMPEZA GERAL 047.2022",
        "UFRGS LIMP GERAL VALE PREFEITURA": "UFRGS LIMPEZA GERAL 047.2022",
        "UFRGS LIMP GERAL VALE SETOR 4": "UFRGS LIMPEZA GERAL 047.2022",
        "UFRGS LIMP GERAL VETERINARIA": "UFRGS LIMPEZA GERAL 047.2022",
        "UFRGS LITORAL LIMP GERAL": "UFRGS LIMPEZA GERAL 047.2022",
        "UFRGS MOTORISTAS DITRAN": "UFRGS MOTORISTAS 034.2022",
        "UFRGS MOTORISTAS FROTA": "UFRGS MOTORISTAS 034.2022",
        "UFRGS MOTORISTAS SUINFRA": "UFRGS MOTORISTAS 034.2022",
        "UFRGS ODONTO": "UFRGS - AUXILIAR DE SAUDE BUCAL 033.2021",
        "UFRGS SERV SAUDE BUCAL": "UFRGS - AUXILIAR DE SAUDE BUCAL 033.2021",
        "VERANOPOLIS": "MUNICIPIO DE VERANOPOLIS 001.2021"
    }

    # 📧 TRADUTOR 2: E-MAILS DOS ANALISTAS (A prova de falhas!)
    # Preencha ou corrija os e-mails oficiais aqui:
    dicionario_emails = {
        "ALESSANDRA": "operacional2@haggltda.com.br",
        "CAREN": "analistapoa1@haggltda.com.br",
        "CARLA ANDREA RIBAS DA SILVA": "assistenteoperacionalpoa1@haggltda.com.br",
        "MICHELE ROSA CELLAS": "gestao.contratos@haggltda.com.br",
        "LAINE NIENOV KREMER": "admalocaxias2@haggltda.com.br",
        "DAIANA KLEIN DE OLIVEIRA": "admalocaxias1@haggltda.com.br"
        # Pode adicionar mais se precisar, sempre em MAIÚSCULO na esquerda!
    }

    try:
        df = pd.read_excel(caminho)
        c = gestor.conn.cursor()
        c.execute("DELETE FROM vinculos_operacao") # Limpa a bagunça antiga

        inseridos = 0
        for index, row in df.iterrows():
            try:
                supervisor_bruto = str(row.iloc[0]).strip()
                analista = str(row.iloc[1]).strip().upper() # Analista em Maiúsculo para o dicionário achar
                contrato_bruto = str(row.iloc[2]).strip().upper()
                
                if not contrato_bruto or contrato_bruto == 'NAN': continue

                # 1. Traduz o Contrato
                contrato_final = dicionario_de_para.get(contrato_bruto, contrato_bruto)
                
                # 2. Puxa o E-mail REAL do Analista pelo nosso dicionário VIP (ignora o excel)
                email_correto = ""
                for chave_nome in dicionario_emails.keys():
                    if chave_nome in analista or analista in chave_nome:
                        email_correto = dicionario_emails[chave_nome]
                        break

                # 3. Arruma o Supervisor (Nome + Telefone)
                sup_nome = supervisor_bruto
                sup_tel = ""
                if '\n' in supervisor_bruto:
                    partes = supervisor_bruto.split('\n')
                    sup_nome = partes[0].strip()
                    sup_tel = partes[1].strip() if len(partes) > 1 else ""

                # 4. Salva no Banco de Dados lindamente
                c.execute("""INSERT INTO vinculos_operacao 
                             (contrato_sistema, analista_nome, analista_email, supervisor_nome, supervisor_tel) 
                             VALUES (?, ?, ?, ?, ?)""",
                          (contrato_final, analista, email_correto, sup_nome, sup_tel))
                inseridos += 1
            except Exception as ex:
                continue
            
        gestor.conn.commit()
        return {"status": "sucesso", "msg": f"✅ {inseridos} vínculos importados! E-mails corrigidos com sucesso!"}
    except Exception as e:
        return {"status": "erro", "msg": str(e)}
    
@eel.expose
def debug_planilha_python():
    import pandas as pd
    
    caminho = r"C:\Users\João Victor Peretti\Documents\FINANCEIRO\COBRANÇA v2\RELAÇÃO OPERACIONAL.xlsx"
    
    print("\n" + "="*50)
    print("🕵️‍♂️ INICIANDO RAIO-X DO SISTEMA")
    print("="*50)
    
    # 1. RAIO-X DA PLANILHA EXCEL
    try:
        df = pd.read_excel(caminho)
        print("\n📊 1. CABEÇALHOS E COLUNAS ENCONTRADAS NO EXCEL:")
        print(df.columns.tolist())
        
        print("\n📄 2. TESTE DE LEITURA DA PRIMEIRA LINHA:")
        # Vamos imprimir índice por índice para ver o que o Python está enxergando
        for i, col in enumerate(df.columns):
            valor = df.iloc[0, i]
            print(f"   [Índice {i}] Coluna '{col}': {valor}")
            
    except Exception as e:
        print(f"❌ Erro ao ler planilha: {e}")

    # 2. RAIO-X DO BANCO DE DADOS
    try:
        c = gestor.conn.cursor()
        c.execute("SELECT contrato_sistema, analista_nome, analista_email FROM vinculos_operacao LIMIT 5")
        resultados = c.fetchall()
        
        print("\n🗄️ 3. O QUE FOI SALVO NO BANCO DE DADOS (5 primeiros):")
        for r in resultados:
            print(f"   Contrato: {r[0]} | Analista: {r[1]} | E-mail: {r[2]}")
            
    except Exception as e:
        print(f"❌ Erro ao ler o banco: {e}")
        
    print("\n" + "="*50)
    return "Raio-X Finalizado! Olhe o terminal preto."

@eel.expose
def obter_contatos_whats_contrato(cliente):
    try:
        c = gestor.conn.cursor()

        # 1. Pega o telefone do Daison direto da tabela de usuários
        c.execute("SELECT telefone FROM usuarios WHERE usuario LIKE 'daison'")
        res_daison = c.fetchone()
        tel_daison = res_daison[0] if res_daison else ""

        # 2. Descobre quem é o supervisor desse contrato no nosso Banco Local
        termo_busca = str(cliente).split('-')[0].strip()
        c.execute("SELECT supervisor_nome FROM vinculos_operacao WHERE contrato_sistema LIKE ?", (f"%{termo_busca}%",))
        res_vinculo = c.fetchone()

        sup_nome = "Supervisor"
        tel_sup = ""

        if res_vinculo and res_vinculo[0]:
            nome_bruto = str(res_vinculo[0]).upper()

            # Traduz pro nome exato que está na tela de Gestão de Usuários
            nome_pesquisa = ""
            if "CASSIO" in nome_bruto: nome_pesquisa = "cassio"
            elif "AFRANIO" in nome_bruto: nome_pesquisa = "afranio"
            elif "ISMAEL" in nome_bruto: nome_pesquisa = "ismael"
            elif "GUSTAVO" in nome_bruto: nome_pesquisa = "gustavo b"

            if nome_pesquisa:
                # Pega o telefone do supervisor no banco de dados!
                c.execute("SELECT usuario, telefone FROM usuarios WHERE usuario LIKE ?", (nome_pesquisa,))
                res_sup = c.fetchone()
                if res_sup:
                    sup_nome = str(res_sup[0]).capitalize()
                    tel_sup = res_sup[1]

        return {
            "status": "sucesso",
            "daison_tel": tel_daison,
            "sup_nome": sup_nome,
            "sup_tel": tel_sup
        }
    except Exception as e:
        return {"status": "erro", "msg": str(e)}
    
@eel.expose
def debug_supervisores_python():
    try:
        c = gestor.conn.cursor()
        
        print("\n" + "="*50)
        print("🕵️‍♂️ INICIANDO INVESTIGAÇÃO DE SUPERVISORES")
        print("="*50)
        
        # 1. Puxa os nomes que vieram da planilha
        c.execute("SELECT DISTINCT supervisor_nome FROM vinculos_operacao")
        supervisores_planilha = c.fetchall()
        
        print("\n📋 1. O que veio da Planilha (vinculos_operacao):")
        for s in supervisores_planilha:
            print(f"   -> '{s[0]}'")
            
        # 2. Puxa os usuários que existem no sistema
        c.execute("SELECT usuario, telefone, funcao FROM usuarios")
        usuarios_sistema = c.fetchall()
        
        print("\n👥 2. O que está cadastrado no Sistema (tabela usuarios):")
        for u in usuarios_sistema:
            print(f"   -> Nome: '{u[0]}' | Tel: '{u[1]}' | Função: '{u[2]}'")
            
        print("\n" + "="*50)
        return "Raio-X de Supervisores Finalizado! Olhe o terminal."
    except Exception as e:
        return f"Erro no debug: {str(e)}"
    
@eel.expose
def debug_divergencia_contratos_python():
    import pandas as pd
    import os

    caminho = r"C:\Users\João Victor Peretti\Documents\FINANCEIRO\COBRANÇA v2\RELAÇÃO OPERACIONAL.xlsx"

    if not os.path.exists(caminho):
        return "Arquivo não encontrado."

    # 🕵️‍♂️ O SEU DICIONÁRIO ATUAL (Gabarito)
    dicionario_de_para = {
        "BENTO HIG/AUX ADM": "BENTO GONÇALVES",
        "PREF POA SAUDE/SAMU": "SAMU",
        "PREF CAXIAS": "CAXIAS DO SUL",
        "TRIUNFO MOT/OP MAQ": "TRIUNFO - MOTORISTAS",
        "TRIUNFO VIGIAS": "TRIUNFO - VIGIAS",
        "SÃO CAMILO": "HOSPITAL SÃO CAMILO"
    }

    print("\n" + "="*70)
    print("🕵️‍♂️ INICIANDO RADAR DE DIVERGÊNCIA DE CONTRATOS")
    print("="*70)

    try:
        # 1. Pega os contratos oficiais que estão nas notas carregadas na tela!
        if not hasattr(gestor, 'todas_as_notas') or not gestor.todas_as_notas:
            print("⚠️ AVISO: Carregue as notas no sistema primeiro para eu ter o Gabarito Oficial!")
            return "Sem notas carregadas no sistema."

        # Extrai os nomes limpos e únicos do sistema
        oficiais_completos = list(set([str(n['cliente']).upper() for n in gestor.todas_as_notas]))
        
        # 2. Lê a planilha do operacional
        df = pd.read_excel(caminho)
        # Pega a Coluna C (índice 2) inteira e tira os duplicados
        contratos_planilha = df.iloc[:, 2].dropna().astype(str).str.strip().str.upper().unique()

        ok_list = []
        erro_list = []

        # 3. O Combate!
        for c_planilha in contratos_planilha:
            if c_planilha == 'NAN': continue

            # Tenta traduzir usando nosso dicionário
            c_traduzido = dicionario_de_para.get(c_planilha, c_planilha)

            # Verifica se o nome traduzido está "dentro" do nome oficial
            match = False
            nome_oficial_encontrado = ""
            for oficial in oficiais_completos:
                if c_traduzido in oficial or oficial in c_traduzido:
                    match = True
                    nome_oficial_encontrado = oficial
                    break

            if match:
                ok_list.append(f"   ✅ '{c_planilha}' -> 🤝 Deu Match com: {nome_oficial_encontrado}")
            else:
                erro_list.append(f"   ❌ '{c_planilha}' -> ⚠️ NÃO ACHOU NINGUÉM OFICIAL")

        # 4. Imprime os Resultados
        print("\n🟢 CONTRATOS RECONHECIDOS (Já estão perfeitos):")
        for ok in ok_list: print(ok)

        print("\n🔴 CONTRATOS DIVERGENTES (Você precisa adicionar no Dicionário do Python):")
        for erro in sorted(erro_list): print(erro)

        print("\n💡 DICA DE OURO: Pegue os nomes da lista 🔴 e adicione no 'dicionario_de_para'")
        print(" apontando para o nome oficial que aparece no seu sistema!")
        print("="*70)
        
        return "Radar Finalizado! Olhe o terminal para ver as divergências."

    except Exception as e:
        print(f"Erro no radar: {e}")
        return str(e)

@eel.expose
def remover_anexo_timeline_python(id_log):
        try:
            c = gestor.conn.cursor()
            
            # 1. Busca o caminho do anexo para tentar apagar o arquivo físico
            c.execute("SELECT anexo FROM log_atividades WHERE id = ?", (id_log,))
            res = c.fetchone()
            
            if res and res[0]:
                caminho_arquivo = str(res[0])
                if os.path.exists(caminho_arquivo):
                    try:
                        os.remove(caminho_arquivo)
                    except Exception as e:
                        print(f"⚠️ Aviso: Não foi possível excluir o arquivo físico: {e}")
            
            # 2. Limpa a coluna no banco de dados
            c.execute("UPDATE log_atividades SET anexo = '' WHERE id = ?", (id_log,))
            gestor.conn.commit()
            
            return {"status": "sucesso"}
            
        except Exception as e:
            return {"status": "erro", "msg": str(e)}  

@eel.expose
def anexar_doc_timeline(id_log, cliente_nome, nota_fiscal):
    try:
        # 1. Abre a janelinha do Windows para escolher o arquivo
        root = tk.Tk()
        root.withdraw()
        root.attributes('-topmost', True)
        arquivo_origem = filedialog.askopenfilename(
            title="Selecione o arquivo para anexar",
            filetypes=[("Todos os Arquivos", "*.*"), ("PDFs", "*.pdf"), ("Imagens", "*.png *.jpg *.jpeg")]
        )
        root.destroy()

        if not arquivo_origem:
            return {"status": "cancelado"} # Usuário fechou a janela sem escolher nada

        # 2. Cria uma pasta segura para os anexos da Timeline dentro da raiz
        ano_atual = str(datetime.now().year)
        pasta_destino = os.path.join(PATH_EVIDENCIAS_RAIZ, ano_atual, cliente_nome, "Dossie_Anexos")
        
        if not os.path.exists(pasta_destino):
            os.makedirs(pasta_destino)

        # 3. Copia o arquivo e coloca um RG nele pra não ter nome duplicado
        nome_original = os.path.basename(arquivo_origem)
        nome_limpo = "".join([c for c in nome_original if c.isalnum() or c in ' ._-'])
        data_hora = datetime.now().strftime("%d%m%Y_%H%M")
        
        novo_nome = f"Nota_{nota_fiscal}_{data_hora}_{nome_limpo}"
        caminho_final = os.path.abspath(os.path.join(pasta_destino, novo_nome))
        
        shutil.copy2(arquivo_origem, caminho_final)

        # 4. Grava o caminho absoluto no nosso banco de dados
        c = gestor.conn.cursor()
        c.execute("UPDATE log_atividades SET anexo = ? WHERE id = ?", (caminho_final, id_log))
        gestor.conn.commit()

        return {"status": "sucesso"}

    except Exception as e:
        print(f"❌ Erro ao anexar documento na timeline: {e}")
        return {"status": "erro", "msg": str(e)}  

eel.init('web')
eel.start('login.html', size=(1400, 850), port=0)
