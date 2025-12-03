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

# --- CONFIGURAÇÃO INICIAL ---
st.set_page_config(page_title="Corpore - Gestão na Nuvem", layout="wide", page_icon="🏥")
st.title("🏥 Sistema de Gestão de Faturas (Google Sheets)")

# --- CONEXÃO GOOGLE SHEETS ---
conn = st.connection("gsheets", type=GSheetsConnection)

def carregar_dados_sheets():
    """Carrega os dados da aba 'guias' ou cria um DF vazio se não existir."""
    try:
        # ttl=5 garante atualização rápida (cache de 5 segundos)
        df = conn.read(worksheet="guias", ttl=5)
        return df
    except Exception:
        # Se a aba não existir ou estiver vazia, retorna estrutura padrão
        return pd.DataFrame(columns=[
            "fatura_ref", "mes_competencia", "ano_competencia", "tipo_usuario", 
            "servicos_fatura", "paciente_nome", "nr_guia", "data_atend", 
            "cod_proced", "valor", "data_lancamento"
        ])

def salvar_no_sheets(df_novo, meta_dados):
    """Salva novos dados no Google Sheets, adicionando ao que já existe."""
    df_existente = carregar_dados_sheets()
    
    data_hoje = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    lista_novos = []
    for _, row in df_novo.iterrows():
        try: val = row['VALOR_CALC']
        except: val = 0.0
        
        lista_novos.append({
            "fatura_ref": meta_dados['fatura'],
            "mes_competencia": meta_dados['mes'],
            "ano_competencia": meta_dados['ano'],
            "tipo_usuario": meta_dados['usuario'],
            "servicos_fatura": meta_dados['servico'],
            "paciente_nome": row['NOME DO PACIENTE'],
            "nr_guia": row['NR DA GUIA'],
            "data_atend": row['DATA ATEND.'],
            "cod_proced": row['CÓDIGO PROCED.'],
            "valor": val,
            "data_lancamento": data_hoje
        })
    
    df_append = pd.DataFrame(lista_novos)
    
    # Concatena antigo com novo
    df_final = pd.concat([df_existente, df_append], ignore_index=True)
    
    # Atualiza a planilha
    conn.update(worksheet="guias", data=df_final)

def atualizar_fatura_sheets(fatura_ref, df_editado, meta_dados):
    """Atualiza uma fatura existente (apaga a antiga e escreve a nova versão)."""
    df_completo = carregar_dados_sheets()
    
    # Remove as linhas da fatura antiga
    df_limpo = df_completo[df_completo['fatura_ref'] != fatura_ref]
    
    # Prepara os novos dados
    data_hoje = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lista_novos = []
    for _, row in df_editado.iterrows():
        try:
            raw_val = str(row['VALOR (R$)'])
            val = float(raw_val.replace('R$', '').replace('.', '').replace(',', '.'))
        except: val = 0.0
        
        lista_novos.append({
            "fatura_ref": fatura_ref, # Mantém a referência original
            "mes_competencia": meta_dados['mes'],
            "ano_competencia": meta_dados['ano'],
            "tipo_usuario": meta_dados['usuario'],
            "servicos_fatura": meta_dados['servico'],
            "paciente_nome": row['NOME DO PACIENTE'],
            "nr_guia": row['NR DA GUIA'],
            "data_atend": row['DATA ATEND.'],
            "cod_proced": row['CÓDIGO PROCED.'],
            "valor": val,
            "data_lancamento": data_hoje
        })
        
    df_append = pd.DataFrame(lista_novos)
    df_final = pd.concat([df_limpo, df_append], ignore_index=True)
    
    conn.update(worksheet="guias", data=df_final)

# --- OCR E EXTRAÇÃO (REFINADO PARA SEU PDF) ---
@st.cache_resource
def load_ocr_reader():
    return easyocr.Reader(['pt'], gpu=False, quantize=True)

def extrair_texto_hibrido(arquivo_bytes):
    texto_final = ""
    usou_ocr = False
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
                del pix, img_data
                gc.collect()
            texto_final = texto_ocr
            usou_ocr = True
        except Exception as e: return f"ERRO_OCR: {str(e)}", False
    return texto_final, usou_ocr

def extrair_dados_pdf(arquivo):
    """
    Regex ajustado especificamente para o layout 'GUIA 12 2025'.
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
        dados["_DEBUG_TEXTO"] = text[:500]

        # 1. NR DA GUIA (Baseado em: "GUIA DE ENCAMINHAMENTO Nr: 56756 / FUSEX")
        match_guia = re.search(r'(?:Nr|Numero)[:\.]?\s*(\d+)', text, flags=re.IGNORECASE)
        if match_guia: dados["NR DA GUIA"] = match_guia.group(1)

        # 2. DATA (Baseado em: "Data: 05/11/2025")
        match_data = re.search(r'Data:\s*(\d{2}/\d{2}/\d{4})', text, flags=re.IGNORECASE)
        if match_data: dados["DATA ATEND."] = match_data.group(1)

        # 3. PACIENTE (Baseado em: "Titular: (PACIENTE)\nNOME...")
        # Procura por Titular ou Dependente e pega a linha seguinte ou o resto da linha
        match_titular = re.search(r'Titular:\s*\(.*?\)\s*\n?(.+)', text, flags=re.IGNORECASE)
        match_dependente = re.search(r'Dependente:\s*\(.*?\)\s*\n?(.+)', text, flags=re.IGNORECASE)
        
        if match_dependente:
            dados["NOME DO PACIENTE"] = match_dependente.group(1).strip()
        elif match_titular:
            dados["NOME DO PACIENTE"] = match_titular.group(1).strip()
        
        # Limpeza do nome (remove UG Origem se vier junto)
        if "UG Origem" in dados["NOME DO PACIENTE"]:
             dados["NOME DO PACIENTE"] = dados["NOME DO PACIENTE"].split("UG Origem")[0].strip()

        # 4. PREC-CP / IDT (Baseado em "Idt: 043516954-5" ou Prec CP)
        match_idt = re.search(r'Idt:\s*([\d-]+)', text, flags=re.IGNORECASE)
        if match_idt: dados["PREC-CP/SIAPE"] = match_idt.group(1)
        else:
            match_prec = re.search(r'Prec CP:\s*(\d+)', text, flags=re.IGNORECASE)
            if match_prec: dados["PREC-CP/SIAPE"] = match_prec.group(1)

        # 5. CÓDIGO PROCEDIMENTO (Busca padrão de 8 dígitos na tabela)
        codigos = re.findall(r'(?<!\d)(\d{8})(?!\d)', text)
        if codigos:
            # Filtra códigos que parecem datas (começam com 2024/2025)
            codigos_validos = [c for c in codigos if not c.startswith("202")]
            dados["CÓDIGO PROCED."] = ", ".join(sorted(set(codigos_validos)))

        # 6. VALOR (Prioridade: "Total :" da tabela ou "Valor Devido")
        # No seu PDF: "Total : 126,00"
        match_total = re.search(r'Total\s*:?\s*([\d\.,]+)', text, flags=re.IGNORECASE)
        if match_total:
             try: dados["VALOR (R$)"] = float(match_total.group(1).replace('.', '').replace(',', '.'))
             except: pass
        
    except Exception as e:
        dados["_DEBUG_TEXTO"] = f"Erro leitura: {str(e)}"

    return dados

# --- WORD E PDF (MANTIDOS) ---
def gerar_doc_word(doc, df_dados, tags, tipo_usuario):
    # (Mantendo sua lógica original de Word)
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
            vals = [row["NOME DO PACIENTE"], row["NR DA GUIA"], row["DATA ATEND."], 
                    row.get("PREC-CP/SIAPE", ""), row["CÓDIGO PROCED."], row["VALOR (R$)"]]
            for i, v in enumerate(vals): 
                if i==5: # Formata Valor
                    try: cells[i].text = f"{float(str(v).replace('R$','').replace(',','.')):,.2f}".replace('.',',')
                    except: cells[i].text = str(v)
                else: cells[i].text = str(v)
                if cells[i].paragraphs: cells[i].paragraphs[0].runs[0].font.size = Pt(9)
        
        # Linha Total
        row_total = tabela.add_row().cells
        row_total[4].text = "TOTAL"
        row_total[5].text = tags["{{TOTAL}}"]
        
    return doc

def criar_template_padrao():
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'; style.font.size = Pt(10)
    
    p = doc.add_paragraph('Corpore Centro de Saúde Ltda - CNPJ 15.259.434/0001-88')
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER; p.runs[0].bold = True
    doc.add_paragraph('')
    
    p_fat = doc.add_paragraph()
    p_fat.add_run('FATURA Nº: ').bold = True
    p_fat.add_run('{{NUM_FATURA}} – {{SERVICO}} – {{MES_ANO}}')
    
    doc.add_paragraph('REFERENTE A USUÁRIO:   FUSEX( ) ...') # Simplificado para brevidade
    
    table = doc.add_table(rows=1, cols=6)
    table.style = 'Table Grid'
    hdr = ["NOME DO PACIENTE", "NR DA GUIA", "DATA ATEND.", "PREC-CP/SIAPE", "CÓD PROCED.", "VALOR R$"]
    for i, h in enumerate(hdr): table.rows[0].cells[i].text = h
    
    doc.add_paragraph('')
    p_ext = doc.add_paragraph()
    p_ext.add_run('VALOR POR EXTENSO: ').bold = True
    p_ext.add_run('{{EXTENSO}} ({{TOTAL}})')
    return doc

def gerar_pdf_protocolo(faturas, qtd_guias, total):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    c.drawString(3*cm, 27*cm, "PROTOCOLO DE REMESSA - CORPORE")
    
    # --- CORREÇÃO: Garante que tudo vire texto antes de juntar ---
    # Se 'faturas' tiver números (ex: 2.1), o .join falha. O str(f) resolve.
    lista_faturas_str = ", ".join([str(f) for f in faturas])
    
    c.drawString(3*cm, 26*cm, f"Faturas: {lista_faturas_str}")
    c.drawString(3*cm, 25.5*cm, f"Total Guias: {qtd_guias} | Valor Total: R$ {total:,.2f}")
    c.save()
    buffer.seek(0)
    return buffer
    
# --- INTERFACE PRINCIPAL ---
tab1, tab2, tab3, tab4 = st.tabs(["📝 Nova Fatura", "✏ Editar (Nuvem)", "📈 Relatórios", "📦 Protocolo"])
meses = {"Janeiro": 1, "Fevereiro": 2, "Março": 3, "Abril": 4, "Maio": 5, "Junho": 6, "Julho": 7, "Agosto": 8, "Setembro": 9, "Outubro": 10, "Novembro": 11, "Dezembro": 12}

# === ABA 1: IMPORTAÇÃO E SALVAMENTO ===
with tab1:
    st.header("📝 Nova Fatura (Salva no Google Sheets)")
    
    if 'df_input' not in st.session_state:
        st.session_state['df_input'] = pd.DataFrame(columns=["NOME DO PACIENTE", "NR DA GUIA", "DATA ATEND.", "PREC-CP/SIAPE", "CÓDIGO PROCED.", "VALOR (R$)"])

    c1, c2, c3 = st.columns(3)
    mes_nome = c1.selectbox("Mês", list(meses.keys()), index=datetime.now().month - 1)
    seq = c1.number_input("Sequencial", 1, 100, 1)
    fatura_ref = f"{meses[mes_nome]}.{seq}"
    c1.info(f"Fatura: **{fatura_ref}**")
    
    ano = c2.number_input("Ano", 2024, 2030, 2025)
    servico = c2.multiselect("Serviço", ["Fisioterapia", "Fonoaudiologia", "Consulta"], default=["Fisioterapia"])
    servico_txt = ", ".join(servico)
    usuario = c3.radio("Convênio", ["FUSEX", "PASS", "S.CIVIL"])

    # Upload e Processamento
    uploaded = st.file_uploader("Arraste os PDFs das Guias", type="pdf", accept_multiple_files=True)
    if uploaded and st.button("Processar PDFs"):
        lista = []
        bar = st.progress(0)
        for i, f in enumerate(uploaded):
            d = extrair_dados_pdf(f)
            if d["NR DA GUIA"]: lista.append(d)
            bar.progress((i+1)/len(uploaded))
        
        if lista:
            novo = pd.DataFrame(lista).drop(columns=["_DEBUG_TEXTO", "_USOU_OCR"])
            st.session_state['df_input'] = pd.concat([st.session_state['df_input'], novo], ignore_index=True)
            st.success(f"{len(lista)} guias lidas com sucesso!")
        else: st.warning("Nenhum dado encontrado nos PDFs.")

    # Editor
    df_editor = st.data_editor(st.session_state['df_input'], num_rows="dynamic")
    
    # Cálculos Finais
    try: 
        vals = df_editor['VALOR (R$)'].apply(lambda x: float(str(x).replace('R$','').replace('.','').replace(',','.')) if isinstance(x, str) else x)
        total = vals.sum()
    except: total = 0.0
    
    st.metric("Total da Fatura", f"R$ {total:,.2f}")
    
    if st.button("💾 Salvar na Nuvem e Baixar Word"):
        df_editor['VALOR_CALC'] = vals
        meta = {'fatura': fatura_ref, 'mes': mes_nome, 'ano': ano, 'usuario': usuario, 'servico': servico_txt}
        
        # 1. Salva no Google Sheets
        salvar_no_sheets(df_editor, meta)
        
        # 2. Gera Word
        doc = criar_template_padrao()
        extenso = num2words(total, lang='pt_BR', to='currency').upper()
        tags = {
            "{{NUM_FATURA}}": fatura_ref, "{{MES_ANO}}": f"{mes_nome}/{ano}",
            "{{SERVICO}}": servico_txt, "{{TOTAL}}": f"R$ {total:,.2f}", "{{EXTENSO}}": extenso
        }
        doc = gerar_doc_word(doc, df_editor, tags, usuario)
        
        buf = BytesIO()
        doc.save(buf)
        buf.seek(0)
        st.download_button("📥 Download DOCX", buf, f"Fatura_{fatura_ref}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        st.success("Dados salvos no Google Sheets!")

# === ABA 2: EDITAR (LENDO DO SHEETS) ===
with tab2:
    st.header("✏ Editar Faturas da Nuvem")
    df_nuvem = carregar_dados_sheets()
    
    if not df_nuvem.empty:
        faturas = df_nuvem['fatura_ref'].unique()
        sel_fat = st.selectbox("Escolha a fatura para editar:", faturas)
        
        if sel_fat:
            df_filtrado = df_nuvem[df_nuvem['fatura_ref'] == sel_fat].copy()
            
            # Recupera metadados da primeira linha
            meta_orig = {
                'mes': df_filtrado.iloc[0]['mes_competencia'],
                'ano': df_filtrado.iloc[0]['ano_competencia'],
                'usuario': df_filtrado.iloc[0]['tipo_usuario'],
                'servico': df_filtrado.iloc[0]['servicos_fatura']
            }
            
            st.info(f"Editando: {meta_orig['servico']} | {meta_orig['usuario']}")
            
            # Prepara para edição
            cols_show = ["paciente_nome", "nr_guia", "data_atend", "cod_proced", "valor"]
            df_edit = df_filtrado[cols_show].rename(columns={
                "paciente_nome": "NOME DO PACIENTE", "nr_guia": "NR DA GUIA", 
                "data_atend": "DATA ATEND.", "cod_proced": "CÓDIGO PROCED.", "valor": "VALOR (R$)"
            })
            
            df_final_edit = st.data_editor(df_edit, num_rows="dynamic")
            
            if st.button("🔄 Atualizar Fatura na Nuvem"):
                atualizar_fatura_sheets(sel_fat, df_final_edit, meta_orig)
                st.success("Planilha atualizada com sucesso!")
                st.rerun()

# === ABA 3: RELATÓRIOS ===
with tab3:
    st.header("📊 Dashboard Financeiro")
    df = carregar_dados_sheets()
    if not df.empty:
        df['valor'] = pd.to_numeric(df['valor'], errors='coerce')
        total_geral = df['valor'].sum()
        c1, c2 = st.columns(2)
        c1.metric("Faturamento Total", f"R$ {total_geral:,.2f}")
        c1.metric("Total de Guias", len(df))
        
        # Gráfico simples
        st.bar_chart(df.groupby("mes_competencia")["valor"].sum())
        st.dataframe(df)

# === ABA 4: PROTOCOLO ===
with tab4:
    st.header("📦 Gerar Protocolo")
    df = carregar_dados_sheets()
    if not df.empty:
        sel = st.multiselect("Selecione Faturas para Envio:", df['fatura_ref'].unique())
        if sel:
            sub = df[df['fatura_ref'].isin(sel)]
            tot = sub['valor'].sum()
            qtd = sub['nr_guia'].nunique()
            st.info(f"Total: R$ {tot:,.2f} ({qtd} guias)")
            
            if st.button("🖨 Baixar PDF Protocolo"):
                pdf = gerar_pdf_protocolo(sel, qtd, tot)
                st.download_button("📥 Download PDF", pdf, "Protocolo.pdf", "application/pdf")

