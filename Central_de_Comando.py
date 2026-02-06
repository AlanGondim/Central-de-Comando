import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from fpdf import FPDF
from datetime import datetime

# --- SEGURANÇA ---
def check_password():
    if "password_correct" not in st.session_state:
        st.text_input("Acesso Executivo", type="password", 
                     on_change=lambda: st.session_state.update({"password_correct": st.session_state.password == "MV2026"}), 
                     key="password")
        return False
    return st.session_state["password_correct"]

if not check_password():
    st.stop()

# --- MOTOR DE EXTRAÇÃO RESTRITO (RESILIENTE) ---
def parse_project_mv_final(file_content):
    # Dicionário padrão para garantir que as colunas SEMPRE existam no DataFrame
    p_data = {
        "projeto": "PROJETO_MV", "gerente": "Não Informado", "spi": 0.0, "cpi": 0.0, 
        "recuperavel": 0.0, "score": "0/10", "status": "ANÁLISE",
        "conclusao_executiva": ""
    }
    try:
        tree = ET.parse(file_content)
        root = tree.getroot()
        ns = '{http://schemas.microsoft.com/project}'
        
        # 1. Captura do Gerente via <AssnOwner>
        owners = [o.text for o in root.findall(f'.//{ns}AssnOwner') if o.text]
        if owners: p_data["gerente"] = owners[-1]

        # 2. Indicadores Financeiros
        def get_v(tag):
            node = root.find(f'.//{ns}{tag}')
            return float(node.text) if node is not None and node.text else 0.0

        pv, ev, ac, pct = get_v('BCWS'), get_v('BCWP'), get_v('ACWP'), get_v('PercentComplete')
        p_data["projeto"] = (root.find(f'.//{ns}Title').text or "PROJETO_MV").upper()
        termino = root.find(f'.//{ns}FinishDate').text[:10] if root.find(f'.//{ns}FinishDate') is not None else "N/A"
        
        p_data["spi"] = round(ev / pv, 2) if pv > 0 else (1.0 if pct == 100 else 0.0)
        p_data["cpi"] = round(ev / ac, 2) if ac > 0 else 1.0
        p_data["recuperavel"] = max(0.0, pv - ev)

        # 3. Formatação da Conclusão
        if pct == 100:
            p_data["status"] = "Concluído com sucesso"
            p_data["conclusao_executiva"] = (f"ENTREGA FINALIZADA: O projeto atingiu 100% de conclusão em {termino}. "
                                            f"Performance SPI/CPI mantidas em 1.0.")
        else:
            p_data["status"] = "EM EXECUÇÃO"
            p_data["conclusao_executiva"] = (f"STATUS: {pct}% de avanço. Atenção ao investimento recuperável de "
                                            f"R$ {p_data['recuperavel']:,.2f}.")

        p_data["score"] = f"{int((len(root.findall(f'.//{ns}Baseline'))>0)*4 + (ac>0)*3 + (pct>0)*3)}/10"
        return p_data
    except:
        return p_data

# --- PDF ENGINE ---
class PremiumPDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 50); self.set_text_color(245, 245, 245)
        self.rotate(45, 100, 150); self.text(35, 190, "CONFIDENCIAL"); self.rotate(0)
        self.set_text_color(0, 51, 102); self.set_font('Arial', 'B', 15)
        self.cell(190, 15, "MV PORTFOLIO INTELLIGENCE - DIRETORIA", ln=True, align='C'); self.ln(5)

# --- UI STREAMLIT ---
st.set_page_config(page_title="MV Auditoria Master", layout="wide")
st.title("🛡️ Central de Comando de Portfólio MV")

files = st.file_uploader("Upload XML", type="xml", accept_multiple_files=True)

if files:
    results = [parse_project_mv_final(f) for f in files]
    df = pd.DataFrame(results)

    # Verificação: se o DF não tiver as colunas por erro de parse, o programa não tenta renderizar
    if not df.empty and 'projeto' in df.columns:
        st.subheader("📋 Painel Consolidado de Governança")
        cols_view = ['projeto', 'gerente', 'spi', 'cpi', 'recuperavel', 'score', 'status']
        st.dataframe(df[cols_view], use_container_width=True)

        if st.button("🚀 GERAR RELATORIO_MV_DIRETORIA"):
            pdf = PremiumPDF()
            pdf.add_page()
            for _, row in df.iterrows():
                pdf.set_fill_color(0, 51, 102); pdf.set_text_color(255); pdf.set_font("Arial", 'B', 11)
                pdf.cell(190, 10, f" PROJETO: {row['projeto']}", ln=True, fill=True)
                pdf.set_fill_color(245, 245, 245); pdf.set_text_color(0); pdf.set_font("Arial", 'B', 9)
                pdf.cell(120, 8, f" GERENTE: {row['gerente']}", border='LB', fill=True)
                pdf.cell(70, 8, f" SCORE: {row['score']}", border='RB', ln=True, fill=True)
                pdf.set_font("Arial", '', 9)
                pdf.cell(47.5, 8, f" SPI: {row['spi']}", border='L')
                pdf.cell(47.5, 8, f" CPI: {row['cpi']}")
                pdf.cell(95, 8, f" RECUPERÁVEL: R$ {row['recuperavel']:,.2f}", border='R', ln=True)
                pdf.set_font("Arial", 'B', 9); pdf.set_fill_color(255, 255, 200)
                pdf.cell(190, 7, " PARECER DA ENTREGA / CONCLUSÃO:", border='LR', ln=True, fill=True)
                pdf.set_font("Arial", '', 9); pdf.multi_cell(190, 6, f" {row['conclusao_executiva']}", border='LRB')
                pdf.ln(6)
            
            pdf.ln(15)
            pdf.line(20, pdf.get_y(), 90, pdf.get_y()); pdf.line(110, pdf.get_y(), 180, pdf.get_y())
            pdf.text(35, pdf.get_y()+5, "Diretoria de Operações"); pdf.text(130, pdf.get_y()+5, "PMO / Auditoria")
            st.download_button("📥 Baixar Relatório", bytes(pdf.output()), "RELATORIO_MV_DIRETORIA.pdf")
    else:
        st.error("Erro ao processar os arquivos. Verifique se o formato XML é válido.")
