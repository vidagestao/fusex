import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from num2words import num2words
import os
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import cm
from io import BytesIO
import re
import pdfplumber
import fitz  # PyMuPDF
import easyocr 
import numpy as np
import gc
import bcrypt
import time

# --- CONFIGURAÇÃO INICIAL ---
st.set_page_config(page_title="Corpore - Acesso Seguro", layout="wide", page_icon="🏥")

# --- CONEXÃO GOOGLE SHEETS ---
conn = st.connection("gsheets", type=GSheetsConnection)

# ==========================================
# 🔐 MÓDULO DE SEGURANÇA E USUÁRIOS
# ==========================================

def carregar_usuarios():
    """Carrega a tabela de usuários. Se não existir, cria estrutura."""
    try:
        # ttl=0 garante que pegamos os dados mais frescos (sem cache) no login
        df = conn.read(worksheet="usuarios", ttl=0)
        return df
    except:
        return pd.DataFrame(columns=["username", "name", "password_hash", "created_at"])

def salvar_novo_usuario(username, name, password):
    """Criptografa a senha e salva o novo usuário no Sheets."""
    df_users = carregar_usuarios()
    
    # Verifica se usuário já existe
    if not df_users.empty and username in df_users['username'].values:
        return False, "Usuário já existe!"
    
    # Criptografia (Hash)
    # O encode converte string para bytes, o salt garante unicidade
    hashed = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
    
    novo_usuario = pd.DataFrame([{
        "username": username,
        "name": name,
        "password_hash": hashed,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }])
    
    # Concatena e Salva
    df_final = pd.concat([df_users, novo_usuario], ignore_index=True)
    try:
        conn.update(worksheet="usuarios", data=df_final)
        return True, "Usuário criado com sucesso!"
    except Exception as e:
        return False, f"Erro ao salvar: {e}"

def autenticar_usuario(username, password):
    """Verifica se usuário e senha batem com o banco."""
    df_users = carregar_usuarios()
    
    if df_users.empty:
        return False, "Nenhum usuário cadastrado."
    
    # Busca o usuário
    user_row = df_users[df_users['username'] == username]
    
    if user_row.empty:
        return False, "Usuário não encontrado."
    
    # Pega o hash salvo
    stored_hash = user_row.iloc[0]['password_hash']
    
    # Verifica a senha
    # checkpw compara a senha digitada (em bytes) com o hash salvo (em bytes)
    if bcrypt.checkpw(password.encode('utf-8'), stored_hash.encode('utf-8')):
        return True, user_row.iloc[0]['name']
    else:
        return False, "Senha incorreta."

def tela_login():
    """Gerencia a interface de Login e Cadastro."""
    st.markdown("## 🏥 Corpore Centro de Saúde")
    
    tab1, tab2 = st.tabs(["🔓 Entrar", "📝 Criar Conta"])
    
    # --- ABA DE LOGIN ---
    with tab1:
        with st.form("login_form"):
            user = st.text_input("Usuário")
            senha = st.text_input("Senha", type="password")
            submitted = st.form_submit_button("Acessar Sistema", type="primary")
            
            if submitted:
                if not user or not senha:
                    st.warning("Preencha todos os campos.")
                else:
                    sucesso, msg = autenticar_usuario(user, senha)
                    if sucesso:
                        st.session_state['logado'] = True
                        st.session_state['usuario_nome'] = msg
                        st.success("Login realizado! Redirecionando...")
                        time.sleep(1)
                        st.rerun()
                    else:
                        st.error(msg)
    
    # --- ABA DE CADASTRO ---
    with tab2:
        st.caption("Primeiro acesso? Crie seu login aqui. A senha será criptografada.")
        with st.form("register_form"):
            new_user = st.text_input("Escolha um Usuário (Login)")
            new_name = st.text_input("Seu Nome Completo")
            new_pass = st.text_input("Escolha uma Senha", type="password")
            new_pass_conf = st.text_input("Confirme a Senha", type="password")
            
            reg_submitted = st.form_submit_button("Cadastrar Usuário")
            
            if reg_submitted:
                if new_pass != new_pass_conf:
                    st.error("As senhas não coincidem!")
                elif len(new_pass) < 4:
                    st.error("A senha deve ter pelo menos 4 caracteres.")
                elif not new_user:
                    st.error("O campo usuário é obrigatório.")
                else:
                    sucesso, msg = salvar_novo_usuario(new_user, new_name, new_pass)
                    if sucesso:
                        st.success(msg)
                        st.info("Agora vá para a aba 'Entrar' e faça login.")
                    else:
                        st.error(msg)

def logout():
    st.session_state['logado'] = False
    st.session_state['usuario_nome'] = ""
    st.rerun()

# ==========================================
# 🏥 MÓDULO PRINCIPAL (SEU SISTEMA)
# ==========================================

def sistema_principal():
    # --- SIDEBAR ---
    with st.sidebar:
        st.write(f"Olá, **{st.session_state['usuario_nome']}**! 👋")
        if st.button("Sair / Logout"):
            logout()
        st.divider()

    st.title("🏥 Gestão de Faturas e Guias")

    # --- FUNÇÕES DE DADOS (Mantidas do seu código anterior) ---
    def carregar_dados_sheets():
        try:
            return conn.read(worksheet="guias", ttl=5)
        except:
            return pd.DataFrame(columns=["fatura_ref", "mes_competencia", "ano_competencia", "tipo_usuario", "servicos_fatura", "paciente_nome", "nr_guia", "data_atend", "cod_proced", "valor", "data_lancamento"])

    def salvar_no_sheets(df_novo, meta_dados):
        # Tenta ler, se falhar cria vazio
        try: df_existente = conn.read(worksheet="guias", ttl=5)
        except: df_existente = pd.DataFrame()
        
        data_hoje = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        lista_novos = []
        for _, row in df_novo.iterrows():
            try: val = row['VALOR_CALC']
            except: val = 0.0
            
            lista_novos.append({
                "fatura_ref": meta_dados['fatura'], "mes_competencia": meta_dados['mes'],
                "ano_competencia": meta_dados['ano'], "tipo_usuario": meta_dados['usuario'],
                "servicos_fatura": meta_dados['servico'], "paciente_nome": row['NOME DO PACIENTE'],
                "nr_guia": row['NR DA GUIA'], "data_atend": row['DATA ATEND.'],
                "cod_proced": row['CÓDIGO PROCED.'], "valor": val,
                "data_lancamento": data_hoje
            })
        
        df_append = pd.DataFrame(lista_novos)
        df_final = pd.concat([df_existente, df_append], ignore_index=True) if not df_existente.empty else df_append
        
        try: conn.update(worksheet="guias", data=df_final)
        except: conn.update(worksheet=0, data=df_final)

    def atualizar_fatura_sheets(fatura_ref, df_editado, meta_dados):
        df_completo = carregar_dados_sheets()
        df_limpo = df_completo[df_completo['fatura_ref'] != fatura_ref]
        
        data_hoje = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        lista_novos = []
        for _, row in df_editado.iterrows():
            try:
                raw_val = str(row['VALOR (R$)'])
                val = float(raw_val.replace('R$', '').replace('.', '').replace(',', '.'))
            except: val = 0.0
            
            lista_novos.append({
                "fatura_ref": fatura_ref, "mes_competencia": meta_dados['mes'],
                "ano_competencia": meta_dados['ano'], "tipo_usuario": meta_dados['usuario'],
                "servicos_fatura": meta_dados['servico'], "paciente_nome": row['NOME DO PACIENTE'],
                "nr_guia": row['NR DA GUIA'], "data_atend": row['DATA ATEND.'],
                "cod_proced": row['CÓDIGO PROCED.'], "valor": val, "data_lancamento": data_hoje
            })
            
        df_append = pd.DataFrame(lista_novos)
        df_final = pd.concat([df_limpo, df_append], ignore_index=True)
        conn.update(worksheet="guias", data=df_final)

    # --- OCR E EXTRAÇÃO ---
    @st.cache_resource
    def load_ocr_reader():
        return easyocr.Reader(['pt'], gpu=False, quantize=True)

    def extrair_texto_hibrido(arquivo_bytes):
        texto_final = ""; usou_ocr = False
        with pdfplumber.open(arquivo_bytes) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t: texto_final += t + "\n"
                
        if len(texto_final.strip()) < 20:
            try:
                reader = load_ocr_reader() 
                arquivo_bytes.seek(0)
                doc = fitz.open(stream=arquivo_bytes.read(), filetype="pdf")
                texto_ocr = ""
                for pagina in doc:
                    pix = pagina.get_pixmap(dpi=150) 
                    img_data = pix.tobytes("png")
                    resultado = reader.readtext(img_data, detail=0, paragraph=True)
                    texto_ocr += "\n".join(resultado) + "\n"
                    del pix, img_data; gc.collect()
                texto_final = texto_ocr; usou_ocr = True
            except Exception as e: return f"ERRO_OCR: {str(e)}", False
        return texto_final, usou_ocr

    def extrair_dados_pdf(arquivo):
        dados = {"NOME DO PACIENTE": "", "NR DA GUIA": "", "DATA ATEND.": "", "PREC-CP/SIAPE": "", "CÓDIGO PROCED.": "", "VALOR (R$)": 0.0}
        try:
            arquivo.seek(0)
            text, _ = extrair_texto_hibrido(arquivo)
            match_guia = re.search(r'(?:Nr|Numero)[:\.]?\s*(\d+)', text, flags=re.IGNORECASE)
            if match_guia: dados["NR DA GUIA"] = match_guia.group(1)
            match_data = re.search(r'Data:\s*(\d{2}/\d{2}/\d{4})', text, flags=re.IGNORECASE)
            if match_data: dados["DATA ATEND."] = match_data.group(1)
            match_titular = re.search(r'Titular:\s*\(.*?\)\s*\n?(.+)', text, flags=re.IGNORECASE)
            match_dependente = re.search(r'Dependente:\s*\(.*?\)\s*\n?(.+)', text, flags=re.IGNORECASE)
            if match_dependente: dados["NOME DO PACIENTE"] = match_dependente.group(1).strip()
            elif match_titular: dados["NOME DO PACIENTE"] = match_titular.group(1).strip()
            if "UG Origem" in dados["NOME DO PACIENTE"]: dados["NOME DO PACIENTE"] = dados["NOME DO PACIENTE"].split("UG Origem")[0].strip()
            match_idt = re.search(r'Idt:\s*([\d-]+)', text, flags=re.IGNORECASE)
            if match_idt: dados["PREC-CP/SIAPE"] = match_idt.group(1)
            else:
                match_prec = re.search(r'Prec CP:\s*(\d+)', text, flags=re.IGNORECASE)
                if match_prec: dados["PREC-CP/SIAPE"] = match_prec.group(1)
            codigos = re.findall(r'(?<!\d)(\d{8})(?!\d)', text)
            if codigos:
                codigos_validos = [c for c in codigos if not c.startswith("202")]
                dados["CÓDIGO PROCED."] = ", ".join(sorted(set(codigos_validos)))
            match_total = re.search(r'Total\s*:?\s*([\d\.,]+)', text, flags=re.IGNORECASE)
            if match_total:
                 try: dados["VALOR (R$)"] = float(match_total.group(1).replace('.', '').replace(',', '.'))
                 except: pass
        except: pass
        return dados

    # --- TEMPLATES WORD E PDF (Mantidos mas resumidos visualmente) ---
    def gerar_doc_word(doc, df_dados, tags, tipo_usuario):
        for p in doc.paragraphs:
            for key, val in tags.items(): 
                if key in p.text: p.text = p.text.replace(key, str(val))
            if "REFERENTE A USUÁRIO" in p.text:
                opcoes = ["FUSEX", "PASS (S.CIVIL)", "FATOR DE CUSTO", "Ex-Combatente"]
                texto_base = "REFERENTE A USUÁRIO:   "
                for op in opcoes: 
                    marcador = "( X )" if tipo_usuario == op else "( )"
                    texto_base += f"{op} {marcador}    "
                p.text = texto_base
        if doc.tables:
            tabela = doc.tables[0]
            for _, row in df_dados.iterrows():
                cells = tabela.add_row().cells
                vals = [row["NOME DO PACIENTE"], row["NR DA GUIA"], row["DATA ATEND."], row.get("PREC-CP/SIAPE", ""), row["CÓDIGO PROCED."], row["VALOR (R$)"]]
                for i, v in enumerate(vals): 
                    if i==5: 
                        try: cells[i].text = f"{float(str(v).replace('R$','').replace(',','.')):,.2f}".replace('.',',')
                        except: cells[i].text = str(v)
                    else: cells[i].text = str(v)
                    if cells[i].paragraphs: cells[i].paragraphs[0].runs[0].font.size = Pt(9)
            row_total = tabela.add_row().cells
            row_total[4].text = "TOTAL"; row_total[5].text = tags["{{TOTAL}}"]
        return doc

    def criar_template_padrao():
        doc = Document()
        style = doc.styles['Normal']; style.font.name = 'Arial'; style.font.size = Pt(10)
        p = doc.add_paragraph('Corpore Centro de Saúde Ltda - CNPJ 15.259.434/0001-88')
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.runs[0].bold = True
        doc.add_paragraph('')
        p_fat = doc.add_paragraph(); p_fat.add_run('FATURA Nº: ').bold = True; p_fat.add_run('{{NUM_FATURA}} – {{SERVICO}} – {{MES_ANO}}')
        doc.add_paragraph('REFERENTE A USUÁRIO:   FUSEX( ) ...') 
        table = doc.add_table(rows=1, cols=6); table.style = 'Table Grid'
        hdr = ["NOME DO PACIENTE", "NR DA GUIA", "DATA ATEND.", "PREC-CP/SIAPE", "CÓD PROCED.", "VALOR R$"]
        for i, h in enumerate(hdr): table.rows[0].cells[i].text = h
        doc.add_paragraph(''); p_ext = doc.add_paragraph(); p_ext.add_run('VALOR POR EXTENSO: ').bold = True; p_ext.add_run('{{EXTENSO}} ({{TOTAL}})')
        return doc

    def gerar_pdf_protocolo(faturas_selecionadas, qtd_guias, total_faturas):
        buffer = BytesIO(); c = canvas.Canvas(buffer, pagesize=A4)
        data_envio = datetime.now().strftime("%d/%m/%Y às %H:%M")
        endereco_fusex = ["Aos Cuidados FUSEX", "Hospital Geral de Juiz de Fora - HGeJF", "Endereço: R. Gen. Deschamps Cavalcante, s/n - Fábrica", "Juiz de Fora - MG, 36080-220"]
        def desenhar_via(y_inicial):
            c.setFont("Helvetica-Bold", 14); c.drawString(2*cm, y_inicial, "CORPORE CENTRO DE SAÚDE LTDA")
            c.setFont("Helvetica", 10); c.drawString(2*cm, y_inicial - 0.5*cm, "PROTOCOLO DE REMESSA DE FATURAS FUSEX")
            c.setFont("Helvetica", 9); y_dest = y_inicial
            for linha in endereco_fusex: c.drawRightString(19*cm, y_dest, linha); y_dest -= 0.4*cm
            y_box = y_inicial - 2.8*cm; c.rect(2*cm, y_box - 2.5*cm, 17*cm, 2.5*cm)
            c.setFont("Helvetica-Bold", 11); c.drawString(2.5*cm, y_box - 0.8*cm, f"QTD FATURAS: {len(faturas_selecionadas)}"); c.drawString(10*cm, y_box - 0.8*cm, f"TOTAL DE GUIAS: {qtd_guias}")
            lista_faturas_str = [str(f) for f in faturas_selecionadas]; texto_faturas = ", ".join(lista_faturas_str)
            if len(texto_faturas) > 90: texto_faturas = texto_faturas[:90] + "..."
            c.setFont("Helvetica", 10); c.drawString(2.5*cm, y_box - 1.5*cm, f"Ref. Faturas: {texto_faturas}")
            valor_fmt = f"{total_faturas:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'); c.drawString(2.5*cm, y_box - 2.2*cm, f"Valor Total Declarado: R$ {valor_fmt}")
            y_ass = y_box - 4.5*cm; c.line(2*cm, y_ass, 9*cm, y_ass); c.setFont("Helvetica", 8); c.drawString(2*cm, y_ass - 0.4*cm, "Despachado por (Corpore)")
            c.line(11*cm, y_ass, 19*cm, y_ass); c.drawString(11*cm, y_ass - 0.4*cm, "Transportado por (Motoboy)")
            y_ass2 = y_ass - 2.5*cm; c.line(2*cm, y_ass2, 19*cm, y_ass2); c.setFont("Helvetica-Bold", 9); c.drawString(2*cm, y_ass2 - 0.5*cm, "Recebido por (Carimbo/Assinatura HGeJF)")
            c.setFont("Helvetica-Oblique", 7); c.drawRightString(19*cm, y_ass2 - 1.2*cm, f"Gerado via Sistema Corpore em: {data_envio}")
        desenhar_via(27*cm); c.setDash(4, 4); c.line(1*cm, 14.85*cm, 20*cm, 14.85*cm); c.setFont("Helvetica", 6); c.drawCentredString(10.5*cm, 14.95*cm, "- - - Corte Aqui - - -"); c.setDash([]); desenhar_via(13*cm)
        c.save(); buffer.seek(0); return buffer

    # --- INTERFACE ABAS ---
    tab1, tab2, tab3, tab4 = st.tabs(["📝 Nova Fatura", "✏ Editar (Nuvem)", "📈 Relatórios", "📦 Protocolo"])
    meses = {"Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12}

    with tab1:
        st.header("📝 Nova Fatura")
        if 'df_input' not in st.session_state: st.session_state['df_input'] = pd.DataFrame(columns=["NOME DO PACIENTE", "NR DA GUIA", "DATA ATEND.", "PREC-CP/SIAPE", "CÓDIGO PROCED.", "VALOR (R$)"])
        c1, c2, c3 = st.columns(3)
        mes_nome = c1.selectbox("Mês", list(meses.keys()), index=datetime.now().month - 1); seq = c1.number_input("Sequencial", 1, 100, 1); fatura_ref = f"{meses[mes_nome]}.{seq}"; c1.info(f"Fatura: **{fatura_ref}**")
        ano = c2.number_input("Ano", 2024, 2030, 2025); servico = c2.multiselect("Serviço", ["Fisioterapia", "Fonoaudiologia", "Consulta"], default=["Fisioterapia"]); servico_txt = ", ".join(servico); usuario = c3.radio("Convênio", ["FUSEX", "PASS", "S.CIVIL"])
        uploaded = st.file_uploader("Arraste os PDFs", type="pdf", accept_multiple_files=True)
        if uploaded and st.button("Processar PDFs"):
            lista = []; bar = st.progress(0)
            for i, f in enumerate(uploaded):
                d = extrair_dados_pdf(f)
                if d["NR DA GUIA"]: lista.append(d)
                bar.progress((i+1)/len(uploaded))
            if lista:
                novo = pd.DataFrame(lista)
                st.session_state['df_input'] = pd.concat([st.session_state['df_input'], novo], ignore_index=True)
                st.success(f"{len(lista)} guias lidas!")
        df_editor = st.data_editor(st.session_state['df_input'], num_rows="dynamic")
        try: vals = df_editor['VALOR (R$)'].apply(lambda x: float(str(x).replace('R$','').replace('.','').replace(',','.')) if isinstance(x, str) else x); total = vals.sum()
        except: total = 0.0
        st.metric("Total", f"R$ {total:,.2f}")
        if st.button("💾 Salvar na Nuvem"):
            df_editor['VALOR_CALC'] = vals; meta = {'fatura': fatura_ref, 'mes': mes_nome, 'ano': ano, 'usuario': usuario, 'servico': servico_txt}
            salvar_no_sheets(df_editor, meta)
            doc = criar_template_padrao(); extenso = num2words(total, lang='pt_BR', to='currency').upper()
            tags = {"{{NUM_FATURA}}": fatura_ref, "{{MES_ANO}}": f"{mes_nome}/{ano}", "{{SERVICO}}": servico_txt, "{{TOTAL}}": f"R$ {total:,.2f}", "{{EXTENSO}}": extenso}
            doc = gerar_doc_word(doc, df_editor, tags, usuario)
            buf = BytesIO(); doc.save(buf); buf.seek(0)
            st.download_button("📥 Download DOCX", buf, f"Fatura_{fatura_ref}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            st.success("Salvo com sucesso!")

    with tab2:
        st.header("✏ Editar Faturas")
        df_nuvem = carregar_dados_sheets()
        if not df_nuvem.empty:
            faturas = df_nuvem['fatura_ref'].unique(); sel_fat = st.selectbox("Editar fatura:", faturas)
            if sel_fat:
                df_filtrado = df_nuvem[df_nuvem['fatura_ref'] == sel_fat].copy()
                meta_orig = {'mes': df_filtrado.iloc[0]['mes_competencia'], 'ano': df_filtrado.iloc[0]['ano_competencia'], 'usuario': df_filtrado.iloc[0]['tipo_usuario'], 'servico': df_filtrado.iloc[0]['servicos_fatura']}
                st.info(f"Editando: {meta_orig['servico']} | {meta_orig['usuario']}")
                cols_show = ["paciente_nome", "nr_guia", "data_atend", "cod_proced", "valor"]
                df_edit = df_filtrado[cols_show].rename(columns={"paciente_nome": "NOME DO PACIENTE", "nr_guia": "NR DA GUIA", "data_atend": "DATA ATEND.", "cod_proced": "CÓDIGO PROCED.", "valor": "VALOR (R$)"})
                df_final_edit = st.data_editor(df_edit, num_rows="dynamic")
                if st.button("🔄 Atualizar Fatura"):
                    atualizar_fatura_sheets(sel_fat, df_final_edit, meta_orig); st.success("Atualizado!"); st.rerun()

    with tab3:
        st.header("📊 Dashboard")
        df = carregar_dados_sheets()
        if not df.empty:
            df['valor'] = pd.to_numeric(df['valor'], errors='coerce'); total_geral = df['valor'].sum()
            c1, c2 = st.columns(2); c1.metric("Faturamento Total", f"R$ {total_geral:,.2f}"); c1.metric("Guias", len(df))
            st.bar_chart(df.groupby("mes_competencia")["valor"].sum()); st.dataframe(df)

    with tab4:
        st.header("📦 Protocolo")
        df = carregar_dados_sheets()
        if not df.empty:
            sel = st.multiselect("Selecione Faturas:", df['fatura_ref'].unique())
            if sel:
                sub = df[df['fatura_ref'].isin(sel)]; tot = sub['valor'].sum(); qtd = sub['nr_guia'].nunique()
                st.info(f"Total: R$ {tot:,.2f} ({qtd} guias)")
                if st.button("🖨 Baixar Protocolo"):
                    pdf = gerar_pdf_protocolo(sel, qtd, tot); st.download_button("📥 PDF", pdf, "Protocolo.pdf", "application/pdf")

# ==========================================
# 🏁 PONTO DE PARTIDA (MAIN)
# ==========================================

if __name__ == "__main__":
    if 'logado' not in st.session_state:
        st.session_state['logado'] = False

    if st.session_state['logado']:
        sistema_principal()
    else:
        tela_login()
