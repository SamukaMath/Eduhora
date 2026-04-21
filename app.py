import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io
import re
import psycopg2
import json
import hashlib

# Dependências de Exportação
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from reportlab.lib.pagesizes import landscape, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

import smtplib
from email.mime.text import MIMEText
import random
import string

# ================= FUNÇÃO DE ENVIO DE E-MAIL =================
def enviar_email_recuperacao(destinatario, nova_senha):
    try:
        remetente = st.secrets["EMAIL_USER"]
        senha_app = st.secrets["EMAIL_PASS"]
        
        assunto = "EduHora  - Recuperação de Senha"
        corpo = f"""Olá!
        
Sua senha foi redefinida com sucesso.
Sua nova senha temporária é: {nova_senha}

Recomendamos que você faça login e atualize sua senha assim que possível.

Equipe EduHora """

        msg = MIMEText(corpo)
        msg['Subject'] = assunto
        msg['From'] = remetente
        msg['To'] = destinatario
        
        # Conecta ao servidor do Gmail
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
            server.login(remetente, senha_app)
            server.send_message(msg)
        return True
    except Exception as e:
        st.error(f"Erro ao enviar o e-mail: {e}")
        return False

# ================= CONFIGURAÇÃO GERAL =================
st.set_page_config(page_title="EduHora - Plataforma", page_icon="🏫", layout="wide")

# ================= CONEXÃO COM POSTGRESQL (NUVEM) =================
# O Streamlit vai buscar essa URL nos "Secrets" que configuraremos depois
DB_URL = st.secrets["DB_URL"]

def run_query(query, params=(), is_select=False):
    # Converte a sintaxe do SQLite (?) para PostgreSQL (%s)
    query = query.replace('?', '%s')
    
    with psycopg2.connect(DB_URL) as conn:
        with conn.cursor() as c:
            c.execute(query, params)
            if is_select:
                return c.fetchall()
            conn.commit()
            return None

def init_db():
    # Em PostgreSQL, AUTOINCREMENT é SERIAL
    run_query('''CREATE TABLE IF NOT EXISTS usuarios (
                    id SERIAL PRIMARY KEY, 
                    email TEXT UNIQUE, 
                    password TEXT,
                    nome TEXT,
                    sobrenome TEXT
                )''')
    run_query('CREATE TABLE IF NOT EXISTS projetos (id SERIAL PRIMARY KEY, user_id INTEGER, nome TEXT)')
    run_query('CREATE TABLE IF NOT EXISTS professores (id SERIAL PRIMARY KEY, projeto_id INTEGER, nome TEXT, manha TEXT, tarde TEXT)')
    run_query('CREATE TABLE IF NOT EXISTS disciplinas (id SERIAL PRIMARY KEY, projeto_id INTEGER, nome TEXT)')
    run_query('CREATE TABLE IF NOT EXISTS turmas (id SERIAL PRIMARY KEY, projeto_id INTEGER, nome TEXT, turno TEXT)')
    run_query('CREATE TABLE IF NOT EXISTS requerimentos (id SERIAL PRIMARY KEY, projeto_id INTEGER, turma TEXT, disciplina TEXT, professor TEXT, aulas INTEGER)')

# Inicializa o banco de dados na nuvem (cria as tabelas se não existirem)
try:
    init_db()
except Exception as e:
    st.error(f"Erro ao conectar com o banco de dados: {e}")

# ================= SEGURANÇA =================
def hash_password(password):
    return hashlib.sha256(password.encode()).hexdigest()

def is_valid_email(email):
    padrao = r'^[\w\.-]+@[\w\.-]+\.\w+$'
    return re.match(padrao, email) is not None


# ================= CONTROLE DE SESSÃO =================
if 'logged_in' not in st.session_state: st.session_state.logged_in = False
if 'user_id' not in st.session_state: st.session_state.user_id = None
if 'user_nome' not in st.session_state: st.session_state.user_nome = None
if 'user_email' not in st.session_state: st.session_state.user_email = None
if 'projeto_id' not in st.session_state: st.session_state.projeto_id = None
if 'projeto_nome' not in st.session_state: st.session_state.projeto_nome = None
if 'erro_login' not in st.session_state: st.session_state.erro_login = None

def tentar_login():
    email = st.session_state.input_email.strip().lower()
    senha = st.session_state.input_senha
    if email and senha:
        # Busca no banco de dados
        user_data = run_query('SELECT id, password, nome FROM usuarios WHERE email=%s', (email,), True)
        # Verifica se achou o usuário e se a senha bate
        if user_data and user_data[0][1] == hash_password(senha):
            st.session_state.logged_in = True
            st.session_state.user_id = user_data[0][0]
            st.session_state.user_nome = user_data[0][2]
            st.session_state.user_email = email
            st.session_state.erro_login = None
        else:
            st.session_state.erro_login = "E-mail ou senha incorretos."

def logout():
    for key in list(st.session_state.keys()): del st.session_state[key]

def fechar_projeto():
    st.session_state.projeto_id = None
    st.session_state.projeto_nome = None
    reset_project_data()

def reset_project_data():
    st.session_state.professores = {}
    st.session_state.disciplinas = []
    st.session_state.turmas = {}
    st.session_state.requerimentos = []

def load_project_data():
    reset_project_data()
    pid = st.session_state.projeto_id
    profs = run_query('SELECT nome, manha, tarde FROM professores WHERE projeto_id=?', (pid,), True)
    for row in profs: st.session_state.professores[row[0]] = {"manha": json.loads(row[1]), "tarde": json.loads(row[2])}
    discs = run_query('SELECT nome FROM disciplinas WHERE projeto_id=?', (pid,), True)
    st.session_state.disciplinas = [row[0] for row in discs]
    turmas = run_query('SELECT nome, turno FROM turmas WHERE projeto_id=?', (pid,), True)
    for row in turmas: st.session_state.turmas[row[0]] = row[1]
    reqs = run_query('SELECT turma, disciplina, professor, aulas FROM requerimentos WHERE projeto_id=?', (pid,), True)
    st.session_state.requerimentos = [{'turma': r[0], 'disciplina': r[1], 'professor': r[2], 'aulas': r[3]} for r in reqs]



# ================= TELA 1: LOGIN E REGISTRO =================
if not st.session_state.logged_in:
    st.title("🏫 EduHora - Acesso à Plataforma")
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("Já tenho uma conta")
        
        # Exibe o erro caso o login tenha falhado
        if st.session_state.erro_login:
            st.error(st.session_state.erro_login)
            st.session_state.erro_login = None # Limpa o erro para não ficar travado
            
        with st.form("login_form"):
            st.text_input("E-mail", key="input_email")
            st.text_input("Senha", type="password", key="input_senha")
            
            # O botão agora aciona o callback ANTES de recarregar a tela
            st.form_submit_button("Entrar", type="primary", use_container_width=True, on_click=tentar_login)
            
        # (Deixe o seu bloco do "Esqueci minha senha" aqui embaixo do jeito que já estava)
        
        # -------- NOVO BLOCO: RECUPERAÇÃO DE SENHA --------
        with st.expander("Esqueci minha senha"):
            st.markdown("Digite seu e-mail cadastrado para receber uma nova senha temporária.")
            rec_email = st.text_input("E-mail para recuperação:", key="rec_email").strip().lower()
            
            if st.button("Enviar nova senha", use_container_width=True):
                if not is_valid_email(rec_email):
                    st.error("Insira um e-mail válido.")
                else:
                    # Verifica se o e-mail existe no banco
                    user_exists = run_query('SELECT id FROM usuarios WHERE email=%s', (rec_email,), True)
                    
                    if user_exists:
                        # Gera uma senha aleatória de 8 caracteres
                        nova_senha = "".join(random.choices(string.ascii_letters + string.digits, k=8))
                        senha_hasheada = hash_password(nova_senha)
                        
                        # Atualiza no banco de dados
                        run_query('UPDATE usuarios SET password=%s WHERE email=%s', (senha_hasheada, rec_email))
                        
                        # Tenta enviar o e-mail
                        with st.spinner("Enviando e-mail..."):
                            if enviar_email_recuperacao(rec_email, nova_senha):
                                st.success("Uma nova senha foi enviada para o seu e-mail! (Verifique também a caixa de Spam).")
                    else:
                        # Por segurança, não dizemos se o e-mail existe ou não de forma explícita para evitar rastreio de dados
                        st.success("Se o e-mail estiver cadastrado, uma nova senha será enviada em instantes.")
        # --------------------------------------------------

    with col2:
        st.subheader("Criar Nova Conta")
        with st.form("register_form"):
            c_nome, c_sobre = st.columns(2)
            reg_nome = c_nome.text_input("Nome").strip()
            reg_sobrenome = c_sobre.text_input("Sobrenome").strip()
            reg_email = st.text_input("E-mail válido").strip().lower()
            reg_pwd = st.text_input("Nova Senha", type="password")
            reg_pwd2 = st.text_input("Confirmar Senha", type="password")
            
            if st.form_submit_button("Cadastrar", use_container_width=True):
                if not reg_nome or not reg_sobrenome or not reg_email or not reg_pwd:
                    st.error("Preencha todos os campos obrigatórios.")
                elif not is_valid_email(reg_email):
                    st.error("Por favor, insira um endereço de e-mail válido.")
                elif reg_pwd != reg_pwd2:
                    st.error("As senhas não coincidem.")
                else:
                    try:
                        run_query('INSERT INTO usuarios (email, password, nome, sobrenome) VALUES (?, ?, ?, ?)', 
                                  (reg_email, hash_password(reg_pwd), reg_nome, reg_sobrenome))
                        st.success(f"Conta criada com sucesso, {reg_nome}! Faça login ao lado.")
                    except psycopg2.IntegrityError:
                        st.error("Este e-mail já está cadastrado.")
                    except Exception as e:
                        st.error(f"Erro no banco: {e}")
    st.stop()
# ================= BARRA LATERAL E PERFIL (USUÁRIO LOGADO) =================
if st.session_state.logged_in:
    st.sidebar.markdown(f"👤 **{st.session_state.user_nome}**\n📧 {st.session_state.user_email}")
    
    # Bloco Expansível para Edição de Perfil
    with st.sidebar.expander("⚙️ Minha Conta / Perfil"):
        # Busca os dados fresquinhos do banco
        dados_usuario = run_query('SELECT nome, sobrenome FROM usuarios WHERE id=%s', (st.session_state.user_id,), True)
        atual_nome, atual_sobre = dados_usuario[0] if dados_usuario else (st.session_state.user_nome, "")
        
        with st.form("form_atualizar_perfil"):
            st.markdown("📝 **Dados Pessoais**")
            upd_nome = st.text_input("Nome", value=atual_nome).strip()
            upd_sobre = st.text_input("Sobrenome", value=atual_sobre).strip()
            
            st.markdown("🔒 **Segurança**")
            upd_senha = st.text_input("Nova Senha (deixe em branco para não alterar)", type="password")
            upd_senha2 = st.text_input("Confirmar Nova Senha", type="password")
            
            if st.form_submit_button("Salvar Alterações", type="primary", use_container_width=True):
                if upd_senha and upd_senha != upd_senha2:
                    st.error("As senhas não coincidem.")
                else:
                    try:
                        if upd_senha:
                            run_query('UPDATE usuarios SET nome=%s, sobrenome=%s, password=%s WHERE id=%s', 
                                      (upd_nome, upd_sobre, hash_password(upd_senha), st.session_state.user_id))
                        else:
                            run_query('UPDATE usuarios SET nome=%s, sobrenome=%s WHERE id=%s', 
                                      (upd_nome, upd_sobre, st.session_state.user_id))
                        
                        # Atualiza a sessão para o nome mudar na interface na mesma hora
                        st.session_state.user_nome = upd_nome
                        st.toast("Perfil atualizado com sucesso!", icon="✅")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Erro ao atualizar: {e}")

    st.sidebar.markdown("---")
    
    # Se estiver dentro de um projeto (TELA 3), mostra o botão de voltar e o nome da escola
    if st.session_state.projeto_id:
        st.sidebar.title(f"🏫 {st.session_state.projeto_nome}")
        st.sidebar.button("⬅️ Trocar de Escola", on_click=fechar_projeto, use_container_width=True)
    
    # Botão de sair global
    st.sidebar.button("🚪 Sair (Logout)", on_click=logout, use_container_width=True)
    
# ================= TELA 2: SELEÇÃO DE PROJETOS =================
if not st.session_state.projeto_id:

    st.title(f"👋 Olá, {st.session_state.user_nome}!")
    st.subheader("Seus Projetos / Escolas")
    
    meus_projetos = run_query('SELECT id, nome FROM projetos WHERE user_id=?', (st.session_state.user_id,), True)
    
    if meus_projetos:
        cols = st.columns(3)
        for i, (p_id, p_nome) in enumerate(meus_projetos):
            with cols[i % 3]:
                st.info(f"**{p_nome}**")
                c1, c2 = st.columns([2, 1])
                with c1:
                    if st.button("Abrir", key=f"abrir_{p_id}", use_container_width=True):
                        st.session_state.projeto_id = p_id
                        st.session_state.projeto_nome = p_nome
                        load_project_data()
                        st.rerun()
                with c2:
                    if st.button("🗑️", key=f"del_{p_id}", help="Excluir Escola"):
                        run_query('DELETE FROM projetos WHERE id=?', (p_id,))
                        run_query('DELETE FROM professores WHERE projeto_id=?', (p_id,))
                        run_query('DELETE FROM disciplinas WHERE projeto_id=?', (p_id,))
                        run_query('DELETE FROM turmas WHERE projeto_id=?', (p_id,))
                        run_query('DELETE FROM requerimentos WHERE projeto_id=?', (p_id,))
                        st.success("Projeto excluído com sucesso!")
                        st.rerun()
    else:
        st.info("Você ainda não tem projetos criados. Crie o primeiro abaixo!")

    st.markdown("---")
    st.subheader("Criar Novo Projeto")
    with st.form("novo_projeto"):
        nome_escola = st.text_input("Nome da Escola / Instituição:")
        if st.form_submit_button("➕ Criar Projeto", type="primary"):
            if nome_escola.strip():
                run_query('INSERT INTO projetos (user_id, nome) VALUES (?, ?)', (st.session_state.user_id, nome_escola.strip()))
                st.rerun()
    st.stop()

# ================= TELA 3: O APLICATIVO PRINCIPAL =================


st.title(f"Projeto Ativo: {st.session_state.projeto_nome}")
dias_semana = ["Segunda", "Terça", "Quarta", "Quinta", "Sexta"]

aba_prof, aba_disc, aba_turmas, aba_distrib, aba_gerar = st.tabs([
    "👩‍🏫 Professores", "📚 Disciplinas", "🎓 Turmas", "⚙️ Grade", "🚀 Gerar Horários"
])

pid = st.session_state.projeto_id

with aba_prof:
    st.subheader("Cadastro de Professores")
    with st.form("form_prof", clear_on_submit=True):
        nome_prof = st.text_input("Nome do Professor:").strip().upper()
        col_m, col_t = st.columns(2)
        m_vars, t_vars = [], []
        with col_m:
            for d in dias_semana: m_vars.append(st.checkbox(f"Evitar Manhã - {d}", key=f"m_{d}"))
        with col_t:
            for d in dias_semana: t_vars.append(st.checkbox(f"Evitar Tarde - {d}", key=f"t_{d}"))
                
        if st.form_submit_button("➕ Salvar Professor", type="primary") and nome_prof:
            off_m = [i for i, v in enumerate(m_vars) if v]
            off_t = [i for i, v in enumerate(t_vars) if v]
            run_query('DELETE FROM professores WHERE projeto_id=? AND nome=?', (pid, nome_prof))
            run_query('INSERT INTO professores (projeto_id, nome, manha, tarde) VALUES (?, ?, ?, ?)', 
                      (pid, nome_prof, json.dumps(off_m), json.dumps(off_t)))
            st.session_state.professores[nome_prof] = {"manha": off_m, "tarde": off_t}
            st.rerun()

    if st.session_state.professores:
        df_profs = pd.DataFrame([
            {"Nome": p, "Evitar Manhã": ", ".join([dias_semana[i] for i in d["manha"]]) or "-", 
             "Evitar Tarde": ", ".join([dias_semana[i] for i in d["tarde"]]) or "-"}
            for p, d in st.session_state.professores.items()
        ])
        st.dataframe(df_profs, use_container_width=True)
        excluir_prof = st.selectbox("Excluir professor:", [""] + list(st.session_state.professores.keys()))
        if st.button("🗑️ Excluir Professor") and excluir_prof:
            run_query('DELETE FROM professores WHERE projeto_id=? AND nome=?', (pid, excluir_prof))
            del st.session_state.professores[excluir_prof]
            st.rerun()

with aba_disc:
    st.subheader("Cadastro de Disciplinas")
    col1, col2 = st.columns([3, 1])
    nova_disc = col1.text_input("Disciplina:").strip()
    if col2.button("➕ Adicionar", use_container_width=True) and nova_disc not in st.session_state.disciplinas:
        run_query('INSERT INTO disciplinas (projeto_id, nome) VALUES (?, ?)', (pid, nova_disc))
        st.session_state.disciplinas.append(nova_disc)
        st.rerun()

    if st.session_state.disciplinas:
        for i, d in enumerate(st.session_state.disciplinas):
            c1, c2 = st.columns([4, 1])
            c1.write(f"📚 {d}")
            if c2.button("🗑️", key=f"del_disc_{i}"):
                run_query('DELETE FROM disciplinas WHERE projeto_id=? AND nome=?', (pid, d))
                st.session_state.disciplinas.remove(d)
                st.rerun()

with aba_turmas:
    st.subheader("Cadastro de Turmas")
    with st.form("form_turmas", clear_on_submit=True):
        col1, col2, col3 = st.columns([2, 2, 1])
        nome_t = col1.text_input("Turma:").strip().upper()
        turno_t = col2.selectbox("Turno:", ["Manhã", "Tarde"])
        if col3.form_submit_button("➕ Adicionar") and nome_t:
            run_query('DELETE FROM turmas WHERE projeto_id=? AND nome=?', (pid, nome_t))
            run_query('INSERT INTO turmas (projeto_id, nome, turno) VALUES (?, ?, ?)', (pid, nome_t, turno_t))
            st.session_state.turmas[nome_t] = turno_t
            st.rerun()

    if st.session_state.turmas:
        df_turmas = pd.DataFrame([{"Turma": t, "Turno": trn} for t, trn in st.session_state.turmas.items()])
        st.dataframe(df_turmas, use_container_width=True)
        excluir_turma = st.selectbox("Excluir turma:", [""] + list(st.session_state.turmas.keys()))
        if st.button("🗑️ Excluir Turma") and excluir_turma:
            run_query('DELETE FROM turmas WHERE projeto_id=? AND nome=?', (pid, excluir_turma))
            del st.session_state.turmas[excluir_turma]
            st.rerun()

with aba_distrib:
    st.subheader("Distribuição de Aulas")
    if not (st.session_state.turmas and st.session_state.disciplinas and st.session_state.professores):
        st.warning("Cadastre Turmas, Disciplinas e Professores primeiro.")
    else:
        with st.form("form_req"):
            c1, c2, c3, c4, c5 = st.columns([2, 2, 2, 1, 1])
            t = c1.selectbox("Turma", sorted(st.session_state.turmas.keys()))
            d = c2.selectbox("Disciplina", sorted(st.session_state.disciplinas))
            p = c3.selectbox("Professor", sorted(st.session_state.professores.keys()))
            a = c4.number_input("Aulas", min_value=1, value=2)
            if c5.form_submit_button("🔗 Vincular"):
                run_query('INSERT INTO requerimentos (projeto_id, turma, disciplina, professor, aulas) VALUES (?, ?, ?, ?, ?)', (pid, t, d, p, a))
                st.session_state.requerimentos.append({'turma': t, 'disciplina': d, 'professor': p, 'aulas': a})
                st.rerun()

        if st.session_state.requerimentos:
            st.dataframe(pd.DataFrame(st.session_state.requerimentos), use_container_width=True)
            if st.button("🗑️ Limpar Toda a Grade desta Escola"):
                run_query('DELETE FROM requerimentos WHERE projeto_id=?', (pid,))
                st.session_state.requerimentos = []
                st.rerun()

with aba_gerar:
    col1, col2, col3 = st.columns(3)
    spin_manha = col1.number_input("Aulas/dia (Manhã):", min_value=1, max_value=10, value=5)
    spin_tarde = col2.number_input("Aulas/dia (Tarde):", min_value=1, max_value=10, value=6)
    max_aulas_disc = col3.number_input("Máx. mesma disciplina/dia:", min_value=1, max_value=5, value=2)
    
    if st.button("⚙️ INICIAR MOTOR MATEMÁTICO", type="primary", use_container_width=True):
        if not st.session_state.requerimentos:
            st.error("Adicione vínculos de aulas primeiro.")
        else:
            with st.spinner("Calculando possibilidades..."):
                modelo = cp_model.CpModel()
                dias = len(dias_semana)
                total_aulas_dia = spin_manha + spin_tarde
                alocacoes = {}
                
                for r_idx, req in enumerate(st.session_state.requerimentos):
                    turno_turma = st.session_state.turmas.get(req['turma'], "Manhã")
                    for d in range(dias):
                        for a in range(total_aulas_dia):
                            if turno_turma == "Manhã" and a >= spin_manha: continue
                            if turno_turma == "Tarde" and a < spin_manha: continue
                            alocacoes[(r_idx, d, a)] = modelo.NewBoolVar(f"R{r_idx}_D{d}_A{a}")
                
                for r_idx, req in enumerate(st.session_state.requerimentos):
                    ap = [alocacoes[(r_idx, d, a)] for d in range(dias) for a in range(total_aulas_dia) if (r_idx, d, a) in alocacoes]
                    modelo.Add(sum(ap) == int(req['aulas']))
                    
                for d in range(dias):
                    for a in range(total_aulas_dia):
                        for turma in st.session_state.turmas.keys():
                            at = [alocacoes[(r_idx, d, a)] for r_idx, req in enumerate(st.session_state.requerimentos) if req['turma'] == turma and (r_idx, d, a) in alocacoes]
                            if at: modelo.AddAtMostOne(at)
                        for prof in st.session_state.professores.keys():
                            ap = [alocacoes[(r_idx, d, a)] for r_idx, req in enumerate(st.session_state.requerimentos) if req['professor'] == prof and (r_idx, d, a) in alocacoes]
                            if ap: modelo.AddAtMostOne(ap)

                for d in range(dias):
                    for turma in st.session_state.turmas.keys():
                        disciplinas_turma = set([req['disciplina'] for req in st.session_state.requerimentos if req['turma'] == turma])
                        for disc in disciplinas_turma:
                            ad = [alocacoes[(r_idx, d, a)] for r_idx, req in enumerate(st.session_state.requerimentos) if req['turma'] == turma and req['disciplina'] == disc for a in range(total_aulas_dia) if (r_idx, d, a) in alocacoes]
                            if ad: modelo.Add(sum(ad) <= max_aulas_disc)

                penalidades = []
                for r_idx, req in enumerate(st.session_state.requerimentos):
                    dados_prof = st.session_state.professores.get(req['professor'], {})
                    off_m, off_t = dados_prof.get("manha", []), dados_prof.get("tarde", [])
                    for d in range(dias):
                        for a in range(total_aulas_dia):
                            if (r_idx, d, a) in alocacoes:
                                if a < spin_manha and d in off_m: penalidades.append(alocacoes[(r_idx, d, a)])
                                elif a >= spin_manha and d in off_t: penalidades.append(alocacoes[(r_idx, d, a)])
                modelo.Minimize(sum(penalidades))
                
                solver = cp_model.CpSolver()
                solver.parameters.max_time_in_seconds = 30.0
                status = solver.Solve(modelo)

                if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                    def ordenar_turmas(lista):
                        return sorted(lista, key=lambda x: [int(c) if c.isdigit() else c.lower() for c in re.split(r'(\d+)', x)])
                    
                    turmas_m = ordenar_turmas([t for t, trn in st.session_state.turmas.items() if trn == "Manhã"])
                    turmas_t = ordenar_turmas([t for t, trn in st.session_state.turmas.items() if trn == "Tarde"])
                    
                    dados_m, dados_t = [], []
                    for d in range(dias):
                        for a in range(spin_manha):
                            linha = {"Dia": dias_semana[d], "Horário": f"{a+1}ª Aula"}
                            for turma in turmas_m:
                                linha[turma] = "---"
                                for r_idx, req in enumerate(st.session_state.requerimentos):
                                    if req['turma'] == turma and (r_idx, d, a) in alocacoes and solver.Value(alocacoes[(r_idx, d, a)]) == 1:
                                        linha[turma] = f"{req['disciplina']} ({req['professor']})"
                                        break
                            dados_m.append(linha)
                        dados_m.append({"Dia": "---", "Horário": "---", **{t: "---" for t in turmas_m}})
                        
                        for a in range(spin_tarde):
                            aula_g = a + spin_manha
                            linha = {"Dia": dias_semana[d], "Horário": f"{a+1}ª Aula"}
                            for turma in turmas_t:
                                linha[turma] = "---"
                                for r_idx, req in enumerate(st.session_state.requerimentos):
                                    if req['turma'] == turma and (r_idx, d, aula_g) in alocacoes and solver.Value(alocacoes[(r_idx, d, aula_g)]) == 1:
                                        linha[turma] = f"{req['disciplina']} ({req['professor']})"
                                        break
                            dados_t.append(linha)
                        dados_t.append({"Dia": "---", "Horário": "---", **{t: "---" for t in turmas_t}})
                    
                    st.session_state.df_manha = pd.DataFrame(dados_m)
                    st.session_state.df_tarde = pd.DataFrame(dados_t)
                    st.success("🎉 Horários gerados com sucesso!")
                else:
                    st.error("Não foi possível gerar os horários. Reveja as regras e cargas horárias.")

    if 'df_manha' in st.session_state:
        st.markdown("### ☀️ Grade - Turno da Manhã")
        st.data_editor(st.session_state.df_manha, use_container_width=True, hide_index=True)
        st.markdown("### 🌆 Grade - Turno da Tarde")
        st.data_editor(st.session_state.df_tarde, use_container_width=True, hide_index=True)
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            st.session_state.df_manha.to_excel(writer, index=False, sheet_name='Manhã')
            st.session_state.df_tarde.to_excel(writer, index=False, sheet_name='Tarde')
        st.download_button("📥 Baixar Excel", data=output.getvalue(), file_name=f"Horarios_{st.session_state.projeto_nome}.xlsx")