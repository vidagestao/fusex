import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from num2words import num2words
import os
import sqlite3
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
from PIL import Image
import gc # Garbage Collector para limpar memória

# --- CONFIGURAÇÃO INICIAL ---
st.set_page_config(page_title="Corpore - Gestão Estratégica", layout="wide")
st.title("🏥 Sistema de Gestão de Faturas e Guias")

# --- CARREGAMENTO DO MODELO OCR (CACHEADO) ---
@st.cache_resource
def load_ocr_reader():
    # Carrega modelo leve para evitar travar PC com pouca memória
    # Nota: No Streamlit Cloud Free, a RAM é limitada. Otimizações são cruciais.
    return easyocr.Reader(['pt', 'en'], gpu=False, quantize=True)

# --- BANCO DE DADOS ---
def init_db():
    conn = sqlite3.connect('faturas.db')
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS guias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fatura_ref TEXT,
            mes_competencia TEXT,
            ano_competencia INTEGER,
            tipo_usuario TEXT,
            servicos_fatura TEXT,
            paciente_nome TEXT,
            nr_guia TEXT,
            data_atend TEXT,
            cod_proced TEXT,
            valor REAL,
            data_lancamento DATE
        )
    ''')
    conn.commit()
    conn.close()

def salvar_no_banco(df, meta_dados):
    conn = sqlite3.connect('faturas.db')
    c = conn.cursor()
    data_hoje = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for _, row in df.iterrows():
        try: val = row['VALOR_CALC']
        except: val = 0.0
        
        c.execute('''
            INSERT INTO guias (
                fatura_ref, mes_competencia, ano_competencia, tipo_usuario, 
                servicos_fatura, paciente_nome, nr_guia, data_atend, 
                cod_proced, valor, data_lancamento
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            meta_dados['fatura'], meta_dados['mes'], meta_dados['ano'],
            meta_dados['usuario'], meta_dados['servico'], row['NOME DO PACIENTE'],
            row['NR DA GUIA'], row['DATA ATEND.'], row['CÓDIGO PROCED.'],
            val, data_hoje
        ))
    conn.commit()
    conn.close()

def atualizar_fatura_existente(fatura_ref, df_novo, meta_dados_originais):
    conn = sqlite3.connect('faturas.db')
    c = conn.cursor()
    c.execute("DELETE FROM guias WHERE fatura_ref = ?", (fatura_ref,))
    data_hoje = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    for _, row in df_novo.iterrows():
        try:
            raw_val = str(row['VALOR (R$)'])
            val = float(raw_val.replace('R$', '').replace('.', '').replace(',', '.'))
        except: val = 0.0
        
        c.execute('''
            INSERT INTO guias (
                fatura_ref, mes_competencia, ano_competencia, tipo_usuario, 
                servicos_fatura, paciente_nome, nr_guia, data_atend, 
                cod_proced, valor, data_lancamento
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            fatura_ref, meta_dados_originais['mes'], meta_dados_originais['ano'],
            meta_dados_originais['usuario'], meta_dados_originais['servico'], 
            row['NOME DO PACIENTE'], row['NR DA GUIA'], row['DATA ATEND.'], 
            row['CÓDIGO PROCED.'], val, data_hoje
        ))
    conn.commit()
    conn.close()

init_db()

# --- INTEGRAÇÃO COM PDF (EASYOCR OTIMIZADO) ---
def extrair_texto_hibrido(arquivo_bytes):
    texto_final = ""
    usou_ocr = False
    
    # 1. Tenta leitura direta (rápida)
    with pdfplumber.open(arquivo_bytes) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t: texto_final += t + "\n"
            
    # 2. Se o texto for muito curto ou vazio, ativa o EasyOCR
    if len(texto_final.strip()) < 20:
        try:
            reader = load_ocr_reader() 
            arquivo_bytes.seek(0)
            doc = fitz.open(stream=arquivo_bytes.read(), filetype="pdf")
            texto_ocr = ""
            
            for pagina in doc:
                # OTIMIZAÇÃO DE MEMÓRIA:
                # DPI reduzido para 150. Isso evita o erro de "not enough memory".
                try:
                    pix = pagina.get_pixmap(dpi=150) 
                    img_data = pix.tobytes("png")
                    resultado = reader.readtext(img_data, detail=0, paragraph=True)
                    texto_ocr += "\n".join(resultado) + "\n"
                except Exception:
                    # Se ainda der erro de memória, tenta com qualidade mínima
                    pix = pagina.get_pixmap(dpi=72)
                    img_data = pix.tobytes("png")
                    resultado = reader.readtext(img_data, detail=0, paragraph=True)
                    texto_ocr += "\n".join(resultado) + "\n"
                
                # Força limpeza da memória
                del pix
                del img_data
                gc.collect()
                
            texto_final = texto_ocr
            usou_ocr = True
        except Exception as e:
            return f"ERRO_OCR: {str(e)}", False

    return texto_final, usou_ocr

def extrair_dados_pdf(arquivo):
    """
    Extrai dados aplicando Regex ajustado para o layout da GUIA.
    """
    dados = {
        "NOME DO PACIENTE": "", "NR DA GUIA": "", "DATA ATEND.": "",
        "PREC-CP/SIAPE": "", "CÓDIGO PROCED.": "", "VALOR (R$)": 0.0,
        "_DEBUG_TEXTO": "", "_USOU_OCR": False
    }
    
    try:
        arquivo.seek(0)
        text, usou_ocr = extrair_texto_hibrido(arquivo)
        dados["_USOU_OCR"] = usou_ocr
        
        if str(text).startswith("ERRO_OCR"):
            dados["_DEBUG_TEXTO"] = text
            return dados
            
        dados["_DEBUG_TEXTO"] = text[:300] if text else "Texto vazio."

        # --- APLICAÇÃO DOS REGEX ---
        
        # 1. NR DA GUIA
        match_guia = re.search(r'(?:GUIA.*?Nr|Nr)[:\.]?\s*(\d+)', text, flags=re.IGNORECASE)
        if match_guia: dados["NR DA GUIA"] = match_guia.group(1)

        # 2. DATA (Manual conforme solicitado, deixamos em branco para preenchimento)
        dados["DATA ATEND."] = "" 

        # 3. PACIENTE (Prioridade Dependente > Titular)
        match_dependente = re.search(r'Dependente:.*?\((?:PACIENTE|PARTICIPANTE)\)\s*(.+)', text, flags=re.IGNORECASE)
        
        if match_dependente:
            raw_name = match_dependente.group(1).split('\n')[0].strip()
            dados["NOME DO PACIENTE"] = raw_name
            
            try:
                idx_nome = text.find(raw_name)
                texto_pos_nome = text[idx_nome:]
                match_prec = re.search(r'Prec CP:[\s\n]*(\d+)', texto_pos_nome, flags=re.IGNORECASE)
                if match_prec: dados["PREC-CP/SIAPE"] = match_prec.group(1)
            except: pass
            
        else:
            match_titular = re.search(r'Titular:\s*\n?(.+)', text, flags=re.IGNORECASE)
            if match_titular:
                raw_name = match_titular.group(1).split('\n')[0].strip()
                if len(raw_name) > 3 and "UG" not in raw_name:
                    dados["NOME DO PACIENTE"] = raw_name
                
                texto_pre_dep = text.split("Dependente")[0] if "Dependente" in text else text
                match_prec_tit = re.search(r'Prec CP:[\s\n]*(\d+)', texto_pre_dep, flags=re.IGNORECASE)
                if match_prec_tit: dados["PREC-CP/SIAPE"] = match_prec_tit.group(1)

        # 4. CÓDIGO DO PROCEDIMENTO
        codigos = re.findall(r'(\d{8})', text)
        if codigos:
            seen = set()
            codigos_unicos = [x for x in codigos if not (x in seen or seen.add(x))]
            dados["CÓDIGO PROCED."] = ", ".join(codigos_unicos)

        # 5. VALOR
        match_total = re.search(r'Total:.*?R?\$?\s*([\d\.,]+)', text, flags=re.IGNORECASE)
        if match_total:
             valor_str = match_total.group(1).replace('.', '').replace(',', '.')
             try: dados["VALOR (R$)"] = float(valor_str)
             except: pass
        else:
            match_devido = re.search(r'Valor Devido:.*?R?\$?\s*([\d\.,]+)', text, flags=re.IGNORECASE)
            if match_devido:
                valor_str = match_devido.group(1).replace('.', '').replace(',', '.')
                try: dados["VALOR (R$)"] = float(valor_str)
                except: pass

    except Exception as e:
        dados["_DEBUG_TEXTO"] = f"Erro geral: {str(e)}"

    return dados

# --- FUNÇÕES DOCX/PDF ---
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
            colunas_dados = [
                row["NOME DO PACIENTE"], 
                row["NR DA GUIA"], 
                row["DATA ATEND."], 
                row.get("PREC-CP/SIAPE", ""), 
                row["CÓDIGO PROCED."], 
                row["VALOR (R$)"]
            ]
            
            for i, val in enumerate(colunas_dados): 
                if i == 5: 
                    try:
                        val_float = float(val) if not isinstance(val, str) else float(str(val).replace('R$','').replace('.','').replace(',','.'))
                        cells[i].text = f"{val_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
                    except: cells[i].text = str(val)
                else: 
                    cells[i].text = str(val)
                
                if cells[i].paragraphs and cells[i].paragraphs[0].runs:
                    cells[i].paragraphs[0].runs[0].font.size = Pt(9)
                    
        row_total = tabela.add_row().cells
        row_total[4].text = "TOTAL"
        row_total[5].text = tags["{{TOTAL}}"]
        
    section = doc.sections[0]
    footer = section.footer
    p_footer = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    agora = datetime.now().strftime("%d/%m/%Y às %H:%M")
    p_footer.text = f"Gestão Corpore - Documento gerado em: {agora}"
    p_footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    if p_footer.runs: p_footer.runs[0].font.size = Pt(8)
    
    return doc

def criar_template_padrao():
    doc = Document()
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(1); section.bottom_margin = Cm(1)
        section.left_margin = Cm(1); section.right_margin = Cm(1)
        
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(10)
    
    p1 = doc.add_paragraph('Corpore Centro de Saúde Ltda')
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if p1.runs: p1.runs[0].bold = True
    
    p2 = doc.add_paragraph('CNPJ 15.259.434/0001-88')
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    doc.add_paragraph('') 
    
    p_fatura = doc.add_paragraph()
    p_fatura.add_run('FATURA Nº: ').bold = True
    p_fatura.add_run('{{NUM_FATURA}} – {{SERVICO}} – {{MES_ANO}}')
    
    doc.add_paragraph('REFERENTE A USUÁRIO:   FUSEX( )      PASS (S.CIVIL)( )    FATOR DE CUSTO( )    Ex- Combatente ( )')
    doc.add_paragraph('')
    
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cols = ["NOME DO PACIENTE", "NR DA GUIA", "DATA ATEND.", "PREC-CP/SIAPE", "CÓD PROCED.", "VALOR R$"]
    for i, col_name in enumerate(hdr_cols):
        hdr_cells[i].text = col_name
        hdr_cells[i].paragraphs[0].runs[0].bold = True
        hdr_cells[i].paragraphs[0].runs[0].font.size = Pt(9)
        
    doc.add_paragraph('')
    
    p_extenso = doc.add_paragraph()
    p_extenso.add_run('VALOR DA FATURA POR EXTENSO:  ').bold = True
    p_extenso.add_run('{{EXTENSO}} ({{TOTAL}})')
    
    return doc

def gerar_pdf_protocolo(faturas_selecionadas, qtd_guias, total_faturas):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    data_envio = datetime.now().strftime("%d/%m/%Y às %H:%M")
    endereco_fusex = ["Aos Cuidados FUSEX", "Hospital Geral de Juiz de Fora - HGeJF", "Endereço: R. Gen. Deschamps Cavalcante, s/n - Fábrica", "Juiz de Fora - MG, 36080-220"]
    
    def desenhar_via(y_inicial):
        c.setFont("Helvetica-Bold", 14)
        c.drawString(2*cm, y_inicial, "CORPORE CENTRO DE SAÚDE LTDA")
        c.setFont("Helvetica", 10)
        c.drawString(2*cm, y_inicial - 0.5*cm, "PROTOCOLO DE REMESSA DE FATURAS FUSEX")
        
        y_dest = y_inicial
        for linha in endereco_fusex: 
            c.drawRightString(19*cm, y_dest, linha)
            y_dest -= 0.5*cm
            
        y_box = y_inicial - 2.5*cm
        c.rect(2*cm, y_box - 2.5*cm, 17*cm, 2.5*cm)
        c.setFont("Helvetica-Bold", 11)
        c.drawString(2.5*cm, y_box - 0.8*cm, f"QUANTIDADE DE FATURAS: {len(faturas_selecionadas)}")
        c.drawString(10*cm, y_box - 0.8*cm, f"TOTAL DE GUIAS ÚNICAS: {qtd_guias}")
        c.setFont("Helvetica", 10)
        c.drawString(2.5*cm, y_box - 1.5*cm, f"Faturas: {', '.join(faturas_selecionadas)[:80]}...")
        c.drawString(2.5*cm, y_box - 2.2*cm, f"Valor Total Declarado: R$ {total_faturas:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        
        y_ass = y_box - 4.5*cm
        c.line(2*cm, y_ass, 9*cm, y_ass)
        c.drawString(2*cm, y_ass - 0.5*cm, "Despachado por (Corpore)")
        c.line(11*cm, y_ass, 19*cm, y_ass)
        c.drawString(11*cm, y_ass - 0.5*cm, "Transportado por (Motoboy)")
        
        y_ass2 = y_ass - 2.5*cm
        c.line(2*cm, y_ass2, 19*cm, y_ass2)
        c.drawString(2*cm, y_ass2 - 0.5*cm, "Recebido por (Carimbo/Assinatura HGeJF)")
        
        c.setFont("Helvetica-Oblique", 8)
        c.drawRightString(19*cm, y_ass2 - 1.5*cm, f"Gestão Corpore - Pronto para envio em: {data_envio}")
        
    desenhar_via(27*cm)
    c.setDash(4, 4)
    c.line(1*cm, 14.85*cm, 20*cm, 14.85*cm)
    c.setDash([])
    desenhar_via(13*cm)
    
    c.save()
    buffer.seek(0)
    return buffer

# --- INTERFACE ---
tab1, tab2, tab3, tab4 = st.tabs(["📝 Nova Fatura (OCR)", "✏ Gerenciar e Editar", "📈 Relatórios", "📦 Protocolo"])
meses_dict = {"Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12}

# ================= ABA 1: NOVA FATURA COM IMPORTAÇÃO =================
with tab1:
    st.header("📝 Nova Fatura & Importação (com EasyOCR)")
    st.info("Atenção: A primeira leitura pode demorar alguns segundos enquanto o sistema carrega a Inteligência Artificial. Após o cache, o processo será rápido.")
    
    if 'df_input_data' not in st.session_state:
        st.session_state['df_input_data'] = pd.DataFrame(columns=["NOME DO PACIENTE", "NR DA GUIA", "DATA ATEND.", "PREC-CP/SIAPE", "CÓDIGO PROCED.", "VALOR (R$)"])

    col1, col2, col3 = st.columns(3)
    
    with col1:
        mes_nome = st.selectbox("Mês de Referência", list(meses_dict.keys()), index=datetime.now().month - 1)
        mes_numero = meses_dict[mes_nome]
        sequencial = st.number_input("Sequencial", min_value=1, value=1)
        fatura_num = f"{mes_numero}.{sequencial}"
        st.info(f"Fatura: *{fatura_num}*")
  
    with col2:
        ano_competencia = st.number_input("Ano", min_value=2024, value=datetime.now().year)
        opcoes_servicos = ["Fisioterapia", "Fonoaudiologia", "Psicologia", "Terapia Ocupacional", "Consulta médica", "Terapias Especiais TEA/TGD", "Nutrição"]
        servicos_selecionados = st.multiselect("Serviços", options=opcoes_servicos, default=["Fisioterapia"])
        servico_texto = ", ".join(servicos_selecionados)
        
    with col3:
        tipo_usuario = st.radio("Convênio", ["FUSEX", "PASS (S.CIVIL)", "FATOR DE CUSTO", "Ex-Combatente"])
    
    st.divider()

    st.markdown("### 📤 Importação Automática")
    uploaded_files = st.file_uploader("Arraste as Guias (PDF) aqui", type=["pdf"], accept_multiple_files=True)
    
    if uploaded_files:
        if st.button(f"Processar {len(uploaded_files)} Arquivos"):
            lista_dados = []
            debug_infos = []
            
            bar = st.progress(0)
            status_text = st.empty() 
            
            for i, pdf_file in enumerate(uploaded_files):
                status_text.text(f"Lendo arquivo {i+1}/{len(uploaded_files)}: {pdf_file.name} (usando IA)...")
                try:
                    dados_extraidos = extrair_dados_pdf(pdf_file)
            
                    texto_debug = dados_extraidos.pop("_DEBUG_TEXTO", "Sem texto")
                    usou_ocr = dados_extraidos.pop("_USOU_OCR", False)
                    
                    status_ocr = " [EASYOCR]" if usou_ocr else ""
                  
                    debug_infos.append(f"ARQUIVO: {pdf_file.name}{status_ocr} | LOG: {texto_debug[:100]}...")

                    if dados_extraidos["NR DA GUIA"] or dados_extraidos["NOME DO PACIENTE"]:
                        lista_dados.append(dados_extraidos)
                    else:
                        st.warning(f"Aviso: Não foi possível ler dados de {pdf_file.name}.")
                except Exception as e:
                    st.error(f"Erro no arquivo {pdf_file.name}: {e}")
                bar.progress((i + 1) / len(uploaded_files))
            
            status_text.text("Processamento concluído!")
    
            if lista_dados:
                novo_df = pd.DataFrame(lista_dados)
                if not novo_df.empty:
                    st.session_state['df_input_data'] = pd.concat(
                        [st.session_state['df_input_data'], novo_df], 
                        ignore_index=True
                    ).drop_duplicates(subset=["NR DA GUIA", "NOME DO PACIENTE"], keep='first').reset_index(drop=True)
                    st.success(f"{len(lista_dados)} guias importadas!")
            else:
                st.error("Nenhum dado encontrado.")
                with st.expander("🕵 Debug"):
                    for d in debug_infos: st.text(d)

    st.divider()
    
    column_config = {
        "NOME DO PACIENTE": st.column_config.TextColumn("Nome do Paciente", width="medium", required=True),
        "NR DA GUIA": st.column_config.TextColumn("Nr. Guia", width="small", required=True),
        "DATA ATEND.": st.column_config.TextColumn("Data", width="small"),
        "PREC-CP/SIAPE": st.column_config.TextColumn("Prec-CP", width="small"),
        "CÓDIGO PROCED.": st.column_config.TextColumn("Cód. Proc.", width="small"),
        "VALOR (R$)": st.column_config.NumberColumn("Valor (R$)", format="R$ %.2f", min_value=0.0)
    }
    
    df_dados = st.data_editor(st.session_state['df_input_data'], column_config=column_config, num_rows="dynamic", key="editor_principal")
    
    def limpar_valor(v):
        """Converte valor para float de forma segura."""
        try:
            if isinstance(v, str): 
                return float(v.replace('R$', '').replace('.', '').replace(',', '.'))
            return float(v)
        except: 
            return 0.0

    total_fatura = 0.0
    valor_extenso = "ZERO REAIS"
    
    if not df_dados.empty:
        df_dados['VALOR_CALC'] = df_dados['VALOR (R$)'].apply(limpar_valor)
        total_fatura = df_dados['VALOR_CALC'].sum()
        try: 
            valor_extenso = num2words(total_fatura, lang='pt_BR', to='currency').upper()
        except: 
            valor_extenso = "---"
            
        st.markdown(f"### Total: {f'R$ {total_fatura:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')}")

    if st.button("💾 Salvar Fatura e Gerar Arquivo"):
        if df_dados.empty: 
            st.warning("Tabela vazia. Não é possível salvar.")
        else:
            try:
                meta_dados = {'fatura': fatura_num, 'mes': mes_nome, 'ano': ano_competencia, 'usuario': tipo_usuario, 'servico': servico_texto}
                salvar_no_banco(df_dados, meta_dados)
                
                # Usa o template padrão criado via código, não precisa subir arquivo .docx
                doc = criar_template_padrao()
                    
                tags = {
                    "{{NUM_FATURA}}": fatura_num, 
                    "{{MES_ANO}}": f"{mes_nome[:3]}/{ano_competencia}", 
                    "{{SERVICO}}": servico_texto, 
                    "{{EXTENSO}}": valor_extenso, 
                    "{{TOTAL}}": f'R$ {total_fatura:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
                }
                
                doc = gerar_doc_word(doc, df_dados, tags, tipo_usuario)
                
                # Salva em memória para download imediato
                buffer = BytesIO()
                doc.save(buffer)
                buffer.seek(0)
                
                safe_usuario = tipo_usuario.replace(" ", "_").replace("(", "").replace(")", "")
                nome_arquivo = f"Fatura_{fatura_num}_{safe_usuario}.docx"
                
                st.success("Salvo com sucesso!")
                st.download_button("📥 Baixar Word", buffer, file_name=nome_arquivo)
                    
                st.session_state['df_input_data'] = pd.DataFrame(columns=["NOME DO PACIENTE", "NR DA GUIA", "DATA ATEND.", "PREC-CP/SIAPE", "CÓDIGO PROCED.", "VALOR (R$)"])
                st.rerun()
            except Exception as e: 
                st.error(f"Erro ao salvar/gerar documento: {e}")

# ================= ABA 2: GERENCIAR E EDITAR =================
with tab2:
    st.header("✏ Editar Faturas Existentes")
    conn = sqlite3.connect('faturas.db')
    try:
        faturas_disponiveis = pd.read_sql_query("SELECT DISTINCT fatura_ref FROM guias ORDER BY fatura_ref DESC", conn)
    except Exception as e:
        # st.error(f"Erro ao carregar faturas: {e}") # Silenciar erro se for primeira vez
        faturas_disponiveis = pd.DataFrame()
        
    conn.close()
    
    if not faturas_disponiveis.empty:
        fatura_selecionada = st.selectbox("Selecione a Fatura para Editar:", faturas_disponiveis['fatura_ref'])
        if fatura_selecionada:
            conn = sqlite3.connect('faturas.db')
            df_edit = pd.read_sql_query(f"SELECT * FROM guias WHERE fatura_ref = '{fatura_selecionada}'", conn)
            
            meta_mes = df_edit['mes_competencia'].iloc[0]
            meta_ano = df_edit['ano_competencia'].iloc[0]
            meta_usuario = df_edit['tipo_usuario'].iloc[0]
            meta_servico = df_edit['servicos_fatura'].iloc[0]
            conn.close()
            
            st.markdown(f"*Detalhes:* {meta_servico} | {meta_usuario} | {meta_mes}/{meta_ano}")
            
            df_para_editor = df_edit[['paciente_nome', 'nr_guia', 'data_atend', 'cod_proced', 'valor']].copy()
            df_para_editor.columns = ["NOME DO PACIENTE", "NR DA GUIA", "DATA ATEND.", "CÓDIGO PROCED.", "VALOR (R$)"]
            df_para_editor['PREC-CP/SIAPE'] = "" 
            
            df_alterado = st.data_editor(df_para_editor, num_rows="dynamic", key="editor_edicao")
            
            col_btn1, col_btn2 = st.columns(2)
            
            if col_btn1.button("💾 Salvar Alterações no Banco"):
                try:
                    meta_dados_originais = {'mes': meta_mes, 'ano': meta_ano, 'usuario': meta_usuario, 'servico': meta_servico}
                  
                    atualizar_fatura_existente(fatura_selecionada, df_alterado, meta_dados_originais)
                    st.success("Atualizado!")
                    st.rerun()
                except Exception as e: st.error(f"Erro: {e}")
                
            if col_btn2.button("📄 Regenerar Documento Word"):
                try:
                    def safe_float(x):
                        try: 
                            if isinstance(x, str): 
                                x = x.replace('R$', '').replace('.', '').replace(',', '.')
                            return float(x)
                        except: return 0.0
                        
                    df_alterado['VALOR_CALC'] = df_alterado['VALOR (R$)'].apply(safe_float)
                    total_novo = df_alterado['VALOR_CALC'].sum()
                    extenso_novo = num2words(total_novo, lang='pt_BR', to='currency').upper()
     
                    doc = criar_template_padrao()
                        
                    tags = {
                        "{{NUM_FATURA}}": fatura_selecionada, 
                        "{{MES_ANO}}": f"{meta_mes[:3]}/{meta_ano}", 
                        "{{SERVICO}}": meta_servico, 
                        "{{EXTENSO}}": extenso_novo, 
                        "{{TOTAL}}": f'R$ {total_novo:,.2f}'.replace(',', 'X').replace('.', ',').replace('X', '.')
                    }
                    
                    doc = gerar_doc_word(doc, df_alterado, tags, meta_usuario)
                    
                    buffer = BytesIO()
                    doc.save(buffer)
                    buffer.seek(0)
                    
                    nome_arq = f"Fatura_{fatura_selecionada}_REVISADA.docx"
                    st.download_button("📥 Baixar", buffer, file_name=nome_arq)
                        
                except Exception as e: st.error(f"Erro: {e}")
    else: 
        st.warning("Nenhuma fatura encontrada no banco de dados local.")

# ================= ABA 3: ESTATÍSTICAS =================
with tab3:
    st.header("📈 Relatório Estatístico")
    conn = sqlite3.connect('faturas.db')
    try:
        c_ano, c_mes = st.columns(2)
        ano_stat = c_ano.number_input("Ano Estatística", 2024, 2030, datetime.now().year)
        lista_meses = ["Todos"] + list(meses_dict.keys())
        mes_stat = c_mes.selectbox("Mês Estatística", lista_meses)
        
        query = "SELECT * FROM guias WHERE ano_competencia = ?"
        params = [ano_stat]
        if mes_stat != "Todos": 
            query += " AND mes_competencia = ?"
            params.append(mes_stat)
            
        df_stat = pd.read_sql_query(query, conn, params=params)
        
        if not df_stat.empty:
            total_periodo = df_stat['valor'].sum()
            kpi1, kpi2, kpi3 = st.columns(3)
            
            kpi1.metric("Faturamento", f"R$ {total_periodo:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            kpi2.metric("Faturas", df_stat['fatura_ref'].nunique())
            kpi3.metric("Guias", df_stat['nr_guia'].nunique())
        
            df_report = df_stat.groupby('fatura_ref').agg(
                Valor_Total=('valor', 'sum'), 
                Qtd_Guias=('nr_guia', 'nunique'), 
                Tipo=('tipo_usuario', 'first')
            ).reset_index()
            
            df_report['Valor_Total'] = df_report['Valor_Total'].apply(
                lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            )
            
            st.dataframe(df_report, width=None)
        else: 
            st.warning("Sem dados no período selecionado.")
    except:
        st.info("O banco de dados ainda está vazio.")
    finally: 
        conn.close()

# ================= ABA 4: PROTOCOLO =================
with tab4:
    st.header("📦 Protocolo de Remessa FUSEX")
    st.markdown("Selecione abaixo as faturas físicas que estão sendo enviadas hoje.")
    conn = sqlite3.connect('faturas.db')
    try:
        faturas_disp = pd.read_sql_query("SELECT DISTINCT fatura_ref FROM guias ORDER BY fatura_ref DESC", conn)
        selecao = st.multiselect("Selecione as Faturas:", options=faturas_disp['fatura_ref'].tolist())
        
        if selecao:
            ph = ','.join(['?'] * len(selecao))
            df_remessa = pd.read_sql_query(f"SELECT * FROM guias WHERE fatura_ref IN ({ph})", conn, params=selecao)
            total_remessa = df_remessa['valor'].sum()
            
            st.info(f"Resumo: {len(selecao)} faturas | R$ {total_remessa:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            
            if st.button("🖨 Gerar Protocolo PDF"):
                pdf_bytes = gerar_pdf_protocolo(selecao, df_remessa['nr_guia'].nunique(), total_remessa)
                st.download_button("📥 Baixar Protocolo", pdf_bytes, f"Protocolo_{datetime.now().strftime('%d-%m-%Y')}.pdf", "application/pdf")
                
    except Exception as e:
        st.error(f"Erro ao gerar protocolo: {e}")
    finally: 
        conn.close()
