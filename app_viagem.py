import streamlit as st
import pandas as pd
import gspread
import re
import os
import base64
import json
import unicodedata
from datetime import datetime, date, timedelta, time
from fpdf import FPDF

# --- UTILS ---
def remove_accents(input_str):
    if not isinstance(input_str, str): return input_str
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return u"".join([c for c in nfkd_form if not unicodedata.combining(c)])

class PDF(FPDF):
    def header(self):
        if os.path.exists('assets/brasao.jpg'):
            try:
                self.image('assets/brasao.jpg', 10, 8, 25)
            except:
                pass
        self.set_font('Arial', 'B', 14)
        self.cell(28) 
        self.cell(0, 8, 'ADIANTAMENTO VIAGEM - SERVICO/TREINAMENTO', 0, 1, 'C')
        self.set_font('Arial', '', 11)
        self.cell(28)
        self.cell(0, 6, 'PREFEITURA MUNICIPAL', 0, 1, 'C')
        self.ln(10)

def cl(t): return str(t).encode('latin-1', 'replace').decode('latin-1')

def criar_pdf_b64(dados, str_pass, tabela_itens, colunas_dias):
    pdf = PDF()
    
    # ---- PAGE 1 ----
    pdf.add_page()
    pdf.ln(12)
    pdf.set_font('Arial', '', 10)
    
    ano_atual = date.today().year
    pdf.cell(0, 6, cl(f"   --- / {ano_atual}" if not dados.get('Numero_Adiantamento') else f"   Nº {dados['Numero_Adiantamento']} / {ano_atual}"), 0, 1)
    
    import textwrap
    def write_wrapped(txt, width=95):
        for line in textwrap.wrap(txt, width=width, break_long_words=True):
            pdf.cell(0, 6, cl(line), 0, 1)

    def separator():
        pdf.ln(2)
        pdf.line(10, pdf.get_y(), 200, pdf.get_y())
        pdf.ln(2)

    pdf.cell(0, 6, cl(f"Solicitante: {dados['Solicitante']}"), 0, 1)
    pdf.cell(0, 6, cl(f"Cargo: {dados.get('Cargo_Solicitante', '-')}"), 0, 1)
    pdf.cell(0, 6, cl(f"Departamento: {dados.get('Setor_Solicitante', '-')}"), 0, 1)
    pdf.cell(0, 6, cl(f"Setor: {dados.get('Setor_Solicitante', '-')}"), 0, 1)
    
    separator()
    
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 6, f"VALOR: R$ {dados['Valor_Final']:.2f}", 0, 1)
    pdf.set_font('Arial', '', 10)
    
    separator()
    
    import textwrap
    def write_wrapped(txt, width=95):
        for line in textwrap.wrap(txt, width=width, break_long_words=True):
            pdf.cell(0, 6, cl(line), 0, 1)

    pdf.cell(0, 6, cl("FUNDAMENTO LEGAL: LEI No 4.297/2018 E PORTARIA 303, DE 15 DE MARCO DE 2018."), 0, 1)
    write_wrapped(f"Destino: {dados['Destino']}")
    write_wrapped(f"Finalidade: {dados['Finalidade']}")
    write_wrapped(f"Data/Periodo: {dados['Data_Saida']} a {dados['Data_Retorno']}")
    
    separator()
    
    v_parts = dados['Veiculo'].split(" - ", 1)
    if len(v_parts) == 2:
        placa, nome_veic = v_parts[0], v_parts[1]
    else:
        placa, nome_veic = "-", dados['Veiculo']
        
    pdf.cell(0, 6, cl(f"Veiculo: {placa}"), 0, 1)
    pdf.cell(0, 6, cl(f"PLACA: {nome_veic}"), 0, 1)
    write_wrapped(f"Horario de Saida: {dados['Hora_Saida']}    Horario de Chegada: {dados['Hora_Retorno']}")
    
    separator()
    
    write_wrapped('Quantidade de Pessoas, Nomes e Cargos: ' + str_pass)
    
    separator()
    
    meses_pt = {1: 'janeiro', 2: 'fevereiro', 3: 'marco', 4: 'abril', 5: 'maio', 6: 'junho', 7: 'julho', 8: 'agosto', 9: 'setembro', 10: 'outubro', 11: 'novembro', 12: 'dezembro'}
    hoje = date.today()
    data_extenso = f"Vargem Grande do Sul, {hoje.day} de {meses_pt[hoje.month]} de {hoje.year}"
    pdf.cell(0, 6, data_extenso, 0, 1)
    
    pdf.ln(15) 
    
    # Linha 1 de assinaturas
    pdf.cell(90, 6, "________________________________________", 0, 0, 'C')
    pdf.cell(90, 6, "________________________________________", 0, 1, 'C')
    pdf.cell(90, 5, "Atendente", 0, 0, 'C')
    pdf.cell(90, 5, "Solicitante", 0, 1, 'C')
    
    pdf.ln(12)
    # Linha 2 de assinaturas
    pdf.cell(190, 6, "________________________________________", 0, 1, 'C')
    pdf.cell(190, 5, cl("Departamento de Financas"), 0, 1, 'C')
    
    pdf.ln(12)
    # Área de Conclusão (Ocupando a largura total)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(190, 6, cl("Conclusao do Controle Interno:"), 0, 1, 'L')
    pdf.set_font('Arial', '', 10)
    
    # Linhas pautadas para preenchimento
    linha_cheia = "_" * 92
    pdf.cell(190, 8, linha_cheia, 0, 1, 'L')
    pdf.cell(190, 8, linha_cheia, 0, 1, 'L')
    pdf.cell(190, 8, linha_cheia, 0, 1, 'L')


    # ---- PAGE 2 ----
    num_dias = max(1, len(colunas_dias))
    orient = 'P' if num_dias <= 6 else 'L'
    pdf.add_page(orientation=orient)
    
    printable_width = 190.0 if orient == 'P' else 277.0
    
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 6, "PREFEITURA MUNICIPAL DE VARGEM GRANDE DO SUL", 0, 1, 'C')
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 6, "CNPJ: 46.248.837/0001-55", 0, 1, 'C')
    if dados.get('Numero_Adiantamento'):
        pdf.cell(0, 6, f"Adiantamento Nº {dados['Numero_Adiantamento']} / {date.today().year}", 0, 1, 'C')
    pdf.ln(4)
    
    pdf.cell(0, 6, cl(f"Responsavel: {dados['Solicitante']}"), 0, 1)
    pdf.cell(0, 6, cl(f"Destino: {dados['Destino']}"), 0, 1)
    pdf.cell(0, 6, f"Data da Viagem: {dados['Data_Saida']} a {dados['Data_Retorno']}", 0, 1)
    
    pdf.ln(3)
    pdf.line(10, pdf.get_y(), 10 + printable_width, pdf.get_y())
    pdf.ln(3)
    
    pdf.cell(0, 6, cl("Segue, os valores referentes a alimentacao:"), 0, 1)
    
    font_s = 8
    if num_dias > 12: font_s = 7
    if num_dias > 20: font_s = 6
    
    pdf.set_fill_color(220, 220, 220)
    pdf.set_font('Arial', 'B', font_s)
    
    w_item = 40 if orient == 'P' else 50
    w_total = 26
    w_pessoa = 26
    w_rest = printable_width - (w_item + w_total + w_pessoa)
    w_dia = max(5.0, w_rest / num_dias)
    
    pdf.cell(w_item, 7, 'Item', 1, 0, 'C', 1)
    for d in colunas_dias:
        pdf.cell(w_dia, 7, cl(d), 1, 0, 'C', 1)
    pdf.cell(w_total, 7, 'Total (R$)', 1, 0, 'C', 1)
    pdf.cell(w_pessoa, 7, 'Por P./dia', 1, 1, 'C', 1)
    
    pdf.set_font('Arial', '', font_s)
    for dt in tabela_itens:
        item_str = dt['Item'].replace('ç','c').replace('õ','o').replace('é','e').replace('ã','a').replace('Í','I').replace('Á','A').replace('Ô','O')
        pdf.cell(w_item, 7, cl(item_str), 1, 0, 'L')
        
        for d in colunas_dias:
            pdf.cell(w_dia, 7, cl(dt[d]), 1, 0, 'C')
            
        pdf.cell(w_total, 7, cl(dt['Total (R$)']), 1, 0, 'C')
        pdf.cell(w_pessoa, 7, cl(dt['Por Pessoa/dia']), 1, 1, 'C')
            
    pdf.ln(6)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 8, f"TOTAL: R$ {dados['Valor_Final']:.2f}", 0, 1)
    pdf.cell(0, 4, "", 0, 1)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 6, f"Numero de Pessoas: {dados['Qtd_Pessoas']}", 0, 1)
    
    pdf.ln(15) # Espaço aumentado
    pdf.cell(190, 6, "Ciente em ____/____/____", 0, 1, 'C')
    pdf.ln(15) # Espaço aumentado
    pdf.cell(190, 6, "______________________________________", 0, 1, 'C')
    pdf.cell(190, 6, cl(dados['Solicitante']), 0, 1, 'C')
    
    res = pdf.output(dest='S')
    if isinstance(res, str): res = res.encode('latin1')
    return base64.b64encode(res).decode('ascii')


def renderizar_recibo_visual(dados_pedido, str_passageiros_cargos, tabela_itens, colunas_dias, is_nova=False):
    th_dias = "".join([f'<th style="border: 1px solid #ddd; padding: 10px; text-align: center; color: #333;">{d}</th>' for d in colunas_dias])
    html_table = f"""
<table style="width: 100%; border-collapse: collapse; margin-top: 10px; margin-bottom: 20px; font-size: 13px;">
    <thead>
        <tr style="background-color: #f8f9fa;">
            <th style="border: 1px solid #ddd; padding: 10px; text-align: left; color: #333;">Item de Custeio</th>
            {th_dias}
            <th style="border: 1px solid #ddd; padding: 10px; text-align: center; color: #333; background-color: #fff1e6;">Total (R$)</th>
            <th style="border: 1px solid #ddd; padding: 10px; text-align: center; color: #333;">Por Pessoa/dia</th>
        </tr>
    </thead>
    <tbody>
"""
    for item in tabela_itens:
        td_dias = "".join([f'<td style="border: 1px solid #ddd; padding: 8px; text-align: center; color: #555;">{item[d]}</td>' for d in colunas_dias])
        html_table += f"""
        <tr>
            <td style="border: 1px solid #ddd; padding: 8px; text-align: left; font-weight: bold;">{item['Item']}</td>
            {td_dias}
            <td style="border: 1px solid #ddd; padding: 8px; text-align: center; background-color: #fff1e6;"><strong>{item['Total (R$)']}</strong></td>
            <td style="border: 1px solid #ddd; padding: 8px; text-align: center; color: #666;">{item['Por Pessoa/dia']}</td>
        </tr>
"""
    html_table += "    </tbody>\n</table>"

    num_ad = f"Nº {dados_pedido.get('Numero_Adiantamento', '')}" if dados_pedido.get('Numero_Adiantamento') else ""

    html_final = f"""
<div style="border: 1px solid #ccc; padding: 30px; border-radius: 8px; background-color: #ffffff; color: #111; font-family: sans-serif; box-shadow: 0 4px 6px rgba(0,0,0,0.05);">
    <h3 style="text-align: center; color: #333; margin-bottom: 2px;">ADIANTAMENTO VIAGEM - SERVIÇO/TREINAMENTO</h3>
    <p style="text-align: center; font-size: 13px; margin-top: 0; color: #666;">PREFEITURA MUNICIPAL<br>{num_ad}</p>
    <hr style="margin: 15px 0; border: 0; border-top: 1px solid #eee;">
    <table style="width: 100%; border: none; font-size: 15px;">
        <tr>
            <td style="width: 50%; vertical-align: top; padding-right: 15px;">
                <p style="margin: 5px 0;"><strong>Solicitante:</strong> {dados_pedido['Solicitante']}</p>
                <p style="margin: 5px 0;"><strong>Cargo:</strong> {dados_pedido.get('Cargo_Solicitante', '-')}</p>
                <p style="margin: 5px 0;"><strong>Departamento/Setor:</strong> {dados_pedido.get('Setor_Solicitante', '-')}</p>
            </td>
            <td style="width: 50%; vertical-align: top;">
                <p style="margin: 5px 0;"><strong>Destino:</strong> {dados_pedido['Destino']}</p>
                <p style="margin: 5px 0;"><strong>Finalidade:</strong> {dados_pedido['Finalidade']}</p>
                <p style="margin: 5px 0;"><strong>Veículo/Placa:</strong> {dados_pedido['Veiculo']}</p>
            </td>
        </tr>
    </table>
    <br>
    <div style="background-color: #f8f9fa; padding: 15px; border-radius: 6px; border: 1px solid #eee;">
        <p style="margin: 4px 0; font-size: 15px;"><strong>Data / Período:</strong> {dados_pedido['Data_Saida']} a {dados_pedido['Data_Retorno']}</p>
        <p style="margin: 4px 0; font-size: 15px;"><strong>Horário de Saída:</strong> {dados_pedido['Hora_Saida']} &nbsp;&nbsp;|&nbsp;&nbsp; <strong>Horário de Chegada:</strong> {dados_pedido['Hora_Retorno']}</p>
        <br>
        <p style="margin: 4px 0; font-size: 15px;"><strong>Quantidade de Pessoas, Nomes e Cargos:</strong></p>
        <p style="margin: 4px 0; font-size: 14px; color: #444;"><i>{str_passageiros_cargos}</i></p>
    </div>
    <hr style="margin: 25px 0; border: 0; border-top: 1px solid #eee;">
    <h4 style="color: #444; font-size: 16px;">Tabela Detalhada de Custeio:</h4>
{html_table}
    <div style="background-color: #f1f8ff; padding: 20px; border-radius: 6px; border: 1px solid #cce5ff; margin-top: 20px;">
        <p style="margin: 4px 0; font-size: 15px;"><strong>Número de Pessoas Assistidas:</strong> {dados_pedido['Qtd_Pessoas']}</p>
        <h2 style="margin-top: 20px; color: #0056b3; text-align: right; font-size: 26px;">TOTAL A REPASSAR: R$ {dados_pedido['Valor_Final']:.2f}</h2>
    </div>
    
    <p style="text-align: center; color: #555; margin-top: 60px;">Ciente em ____/____/____</p>
    <br><br><br><br>
    <p style="text-align: center; color: #555;">____________________________________________________________<br>
    <strong style="color: #111;">{dados_pedido['Solicitante']}</strong></p>
</div>
"""
    html_final = html_final.replace('\n', '')
    st.markdown(html_final, unsafe_allow_html=True)
    
    try:
        b64_pdf = criar_pdf_b64(dados_pedido, str_passageiros_cargos, tabela_itens, colunas_dias)
        href = f'<a href="data:application/pdf;base64,{b64_pdf}" download="Adiantamento_Ordem.pdf" target="_blank"><button style="width:100%; padding:15px; font-size:18px; font-weight:bold; cursor:pointer; background-color:#2e86c1; color:white; border:none; border-radius:5px;">📥 BAIXAR RECIBO PDF OFICIAL</button></a>'
        
        if is_nova:
            if not st.session_state.get('pedido_salvo', False):
                st.warning("⚠️ Atenção: Este cálculo ainda é provisório e **NÃO** foi salvo no banco de dados.")
                if st.button("💾 CONFIRMAR RESULTADO E SALVAR NO BANCO", type="primary", use_container_width=True):
                    with st.spinner("Gravando no Banco de Dados..."):
                        sucesso = save_pedido(dados_pedido, tabela_itens, colunas_dias)
                    if sucesso:
                        st.session_state['pedido_salvo'] = True
                        st.rerun()
            else:
                st.success("✅ Pedido gravado com sucesso! Já pode emitir o PDF oficial.")
                st.markdown(href, unsafe_allow_html=True)
        else:
            st.info("Sua Ficha Oficial reconstituída foi gerada abaixo:")
            st.markdown(href, unsafe_allow_html=True)
            
    except Exception as e:
        st.error(f"Erro ao computar PDF nativo: {e}")

# --- GSPREAD ---
st.set_page_config(page_title="Adiantamento de Viagem", page_icon="🚗", layout="wide")

SHEET_ID = '17LKTIN_wEMCqbN7L04cT2fqTsCABCJDsgLY7wKP88xo'
CRED_FILE = 'credenciais.json'

@st.cache_resource
def init_gspread():
    try:
        try:
            has_secrets = "gcp_service_account" in st.secrets
        except FileNotFoundError:
            has_secrets = False
            
        if has_secrets:
            cred_dict = dict(st.secrets["gcp_service_account"])
            cred_dict["private_key"] = cred_dict["private_key"].replace('\\n', '\n')
            gc = gspread.service_account_from_dict(cred_dict)
        else:
            gc = gspread.service_account(filename=CRED_FILE)
        return gc
    except Exception as e:
        st.error(f"Erro ao autenticar com o Google Sheets: {e}.")
        return None

@st.cache_data(ttl=300)
def pull_sheet_data():
    gc = init_gspread()
    if not gc: return None, None, None, None
    sh = gc.open_by_key(SHEET_ID)
    
    ws_mun = sh.worksheet("Mun")
    df_mun = pd.DataFrame(ws_mun.get_all_values())
    
    df_cidades = df_mun.iloc[1:, 1:6].copy()
    df_cidades.columns = ['Cod', 'Municipio', 'UF', 'Populacao', 'Porte']
    df_cidades = df_cidades[df_cidades['Municipio'].str.strip() != '']
    df_cidades['Municipio_UF'] = df_cidades['Municipio'] + " - " + df_cidades['UF']
    # Aplica remoção de acentos para busca facilitada
    df_cidades['Municipio_UF'] = df_cidades['Municipio_UF'].apply(remove_accents).str.upper()
    
    df_mot = df_mun.iloc[1:, 14:18].copy()
    df_mot.columns = ['Matricula', 'Nome_Motorista', 'Cargo', 'Setor']
    df_mot = df_mot[df_mot['Nome_Motorista'].str.strip() != '']
    
    df_veic = df_mun.iloc[1:, 25:29].copy()
    df_veic.columns = ['Placa', 'Veiculo', 'Ano', 'Combustivel']
    df_veic = df_veic[df_veic['Placa'].str.strip() != '']
    df_veic['Label'] = df_veic['Placa'] + " - " + df_veic['Veiculo']
    
    valores_porte = {"GP": 150.0, "MP": 100.0, "PP": 80.0}
    for r in range(min(50, df_mun.shape[0])):
        for c in range(df_mun.shape[1]-1):
            val = str(df_mun.iloc[r, c]).strip().upper()
            if val in ["GP", "MP", "PP"]:
                try:
                    viz = str(df_mun.iloc[r, c+1])
                    num_str = re.sub(r'[^\d\,]', '', viz).replace(',', '.')
                    if num_str: valores_porte[val] = float(num_str)
                except: pass
                    
    return df_cidades, df_mot, df_veic, valores_porte

@st.cache_data(ttl=60)
def pull_banco_dados():
    gc = init_gspread()
    if not gc: return []
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws = sh.worksheet("Banco de Dados")
        return ws.get_all_values()
    except:
        return []

def save_pedido(dados, tabela_itens, colunas_dias):
    gc = init_gspread()
    if not gc: return False
    sh = gc.open_by_key(SHEET_ID)
    try:
        ws_pedido = sh.worksheet("Banco de Dados")
    except:
        ws_pedido = sh.add_worksheet(title="Banco de Dados", rows="1000", cols="30")
        
    if not ws_pedido.row_values(1):
        ws_pedido.append_row([
            "Carimbo", "Status", "Solicitante", "Tripulação", 
            "Destino", "Finalidade", "Veículo/Placa", "Período (Início e Fim)", 
            "Total Pessoas", "Qtd Cafés (Equipe)", "Total R$ Cafés", 
            "Qtd Almoços (Equipe)", "Total R$ Almoços", "Qtd Jantas (Equipe)", "Total R$ Jantas", 
            "Qtd Pernoites (Equipe)", "Total R$ Pernoites", 
            "Despesa Extra (Ped/Comb)", "Despesa Inesperada", "TOTAL GERAL ADIANTAMENTO"
        ])
        
    json_str = json.dumps({
        'dados': dados,
        'str_pass': dados['Nomes_Passageiros'],
        'tabela': tabela_itens,
        'colunas_dias': colunas_dias
    })
    
    row = [
        datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        dados['Status'],
        dados['Solicitante'],
        dados['Nomes_Passageiros'],
        dados['Destino'],
        dados['Finalidade'],
        dados['Veiculo'],
        f"{dados['Data_Saida']} a {dados['Data_Retorno']}",
        dados['Qtd_Pessoas'],
        dados['Qtd_Cafes'], dados['V_Cafes'],
        dados['Qtd_Almocos'], dados['V_Almocos'],
        dados['Qtd_Jantas'], dados['V_Jantas'],
        dados['Qtd_Pernoites'], dados['V_Pernoites'],
        dados['Extras'],
        dados['Inesperadas'],
        dados['Valor_Final'],
        dados.get('Numero_Adiantamento', ''),
        json_str
    ]
    ws_pedido.append_row(row)
    return True

# --- UI ---
st.title("🚗 Sistema de Adiantamento de Viagem")

df_cid, df_mot, df_veic, val_porte = pull_sheet_data()
if df_cid is None:
    st.error("⚠️ Não foi possível carregar os dados. Verifique a conexão.")
    st.stop()

# SIDEBAR: Cadastros
st.sidebar.header("⚙️ Cadastros Rápidos")

with st.sidebar.expander("➕ Cadastrar Funcionário", expanded=False):
    with st.form("form_nova_pessoa", clear_on_submit=True):
        new_nome = st.text_input("Nome Completo (*)").strip()
        new_cargo = st.text_input("Cargo").strip()
        new_setor = st.text_input("Setor").strip()
        submit_pessoa = st.form_submit_button("Salvar")
        if submit_pessoa and new_nome:
            with st.spinner("Salvando..."):
                try:
                    gc = init_gspread()
                    sh = gc.open_by_key(SHEET_ID)
                    ws_mun = sh.worksheet("Mun")
                    nomes_atuais = ws_mun.col_values(16)
                    next_row = len(nomes_atuais) + 1
                    ws_mun.update(f"O{next_row}:R{next_row}", [["NOVO REGISTRO", new_nome.upper(), new_cargo.upper(), new_setor.upper()]])
                    st.cache_data.clear()
                    st.success(f"{new_nome.upper()} cadastrado!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro: {e}")

with st.sidebar.expander("🚙 Cadastrar Veículo", expanded=False):
    with st.form("form_novo_veic", clear_on_submit=True):
        new_placa = st.text_input("Placa (*)").strip()
        new_veiculo = st.text_input("Veículo/Modelo").strip()
        submit_veic = st.form_submit_button("Salvar")
        if submit_veic and new_placa:
            with st.spinner("Salvando..."):
                try:
                    gc = init_gspread()
                    sh = gc.open_by_key(SHEET_ID)
                    ws_mun = sh.worksheet("Mun")
                    placas_atuais = ws_mun.col_values(26)
                    next_row = len(placas_atuais) + 1
                    ws_mun.update(f"Z{next_row}:AC{next_row}", [[new_placa.upper(), new_veiculo.upper(), "", ""]])
                    st.cache_data.clear()
                    st.success(f"Veículo cadastrado!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Erro: {e}")

# TABS
tab1, tab2 = st.tabs(["NOVO ADIANTAMENTO", "🖨️ REIMPRIMIR ANTIGOS"])

with tab1:
    with st.form("form_viagem"):
        st.subheader("📍 Detalhes da Rota")
        col1, col2, col3 = st.columns([2, 2, 1])
        destino = col1.selectbox("Município de Destino", df_cid['Municipio_UF'].tolist())
        finalidade = col2.text_input("Finalidade da Viagem")
        numero_adiantamento = col3.text_input("Nº Adiantamento (Opcional)")
        
        st.subheader("🕒 Cronograma")
        c1, c2, c3, c4 = st.columns(4)
        d_saida = c1.date_input("Data de Saída", format="DD/MM/YYYY")
        h_saida = c2.time_input("Horário de Saída", value=None)
        d_retorno = c3.date_input("Data de Retorno", format="DD/MM/YYYY")
        h_retorno = c4.time_input("Horário de Retorno", value=None)
        
        st.subheader("👥 Equipe e Logística")
        lista_pessoas = sorted(df_mot['Nome_Motorista'].tolist())
        c5, c7 = st.columns(2)
        solicitante = c5.selectbox("Solicitante/Motorista principal", lista_pessoas)
        veiculo = c7.selectbox("Veículo", df_veic['Label'].tolist())
        
        passageiros = st.multiselect("Nomes dos demais passageiros (Opcional)", lista_pessoas)
        qtd_pessoas = 1 + len(passageiros)
        st.caption(f"**Total de Pessoas a Viajar:** {qtd_pessoas}")
        
        st.subheader("💰 Despesas Adicionais Estimadas")
        cx1, cx2, cx3 = st.columns(3)
        combustivel = cx1.number_input("Adiant. Combustível (R$)", min_value=0.0, step=50.0)
        pedagio = cx2.number_input("Pedágio / Táxi / Ônibus (R$)", min_value=0.0, step=10.0)
        hospedagem = cx3.checkbox("🏨 Cobrir Hospedagem Extra além da Diária?")
        
        inesperada = st.checkbox("🚩 Despesa Inesperada (Soma Automática de 3x Almoço do respectivo porte)")
        
        st.divider()
        submetido = st.form_submit_button("🧮 Calcular Diárias", type="primary", use_container_width=True)

    if submetido:
        row_mun = df_cid[df_cid['Municipio_UF'] == destino].iloc[0]
        porte = str(row_mun['Porte']).strip().upper()
        if porte not in val_porte: porte = "PP"
            
        tabela_fixa = {
            "GP": {"almoco": 55.0, "jantar": 55.0, "cafe": 20.0, "pernoite": 190.0},
            "MP": {"almoco": 40.0, "jantar": 40.0, "cafe": 18.0, "pernoite": 140.0},
            "PP": {"almoco": 31.0, "jantar": 31.0, "cafe": 15.0, "pernoite": 100.0}
        }
        if porte not in tabela_fixa: porte = "PP"
            
        dias = (d_retorno - d_saida).days
        if dias < 0:
            st.error("A Data de Retorno não pode ser anterior à Saída!")
            st.stop()
            
        if h_saida is None or h_retorno is None:
            st.warning("⚠️ Preencha os Horários de Saída e Retorno para calcular as frações de refeição!")
            st.stop()
            
        time_limit_saida = time(7, 0)
        time_limit_retorno = time(19, 0)
        
        v_cafe = tabela_fixa[porte]['cafe']
        v_almoco = tabela_fixa[porte]['almoco']
        v_janta = tabela_fixa[porte]['jantar']
        v_pernoite = tabela_fixa[porte]['pernoite']
        v_almoco_base = v_almoco

        cafe_row = {"Item": "CAFÉ"}
        almoco_row = {"Item": "ALMOÇO"}
        janta_row = {"Item": "JANTAR"}
        pernoite_row = {"Item": "HOSPEDAGEM"}
        
        qtd_cafes = 0; qtd_almocos = 0; qtd_jantas = 0; qtd_pernoites = 0
        colunas_dias = []
        
        for i in range(dias + 1):
            dia_atual = d_saida + timedelta(days=i)
            str_dia = dia_atual.strftime('%d/%m')
            colunas_dias.append(str_dia)
            
            is_primeiro = (i == 0)
            is_ultimo = (i == dias)
            
            cafe = True; almoco = True; jantar = True; pernoite = False if is_ultimo else True
            if is_primeiro and h_saida >= time_limit_saida: cafe = False
            if is_ultimo and h_retorno <= time_limit_retorno: jantar = False
                
            if cafe: qtd_cafes += 1
            if almoco: qtd_almocos += 1
            if jantar: qtd_jantas += 1
            if (pernoite and hospedagem): qtd_pernoites += 1
            
            cafe_row[str_dia] = f"R$ {(v_cafe * qtd_pessoas):.2f}" if cafe else "-"
            almoco_row[str_dia] = f"R$ {(v_almoco * qtd_pessoas):.2f}" if almoco else "-"
            janta_row[str_dia] = f"R$ {(v_janta * qtd_pessoas):.2f}" if jantar else "-"
            pernoite_row[str_dia] = f"R$ {(v_pernoite * qtd_pessoas):.2f}" if (pernoite and hospedagem) else "-"

        v_total_cafes = (qtd_cafes * v_cafe) * qtd_pessoas
        v_total_almocos = (qtd_almocos * v_almoco) * qtd_pessoas
        v_total_jantas = (qtd_jantas * v_janta) * qtd_pessoas
        v_total_pernoites = (qtd_pernoites * v_pernoite) * qtd_pessoas
        
        total_diarias_equipe = v_total_cafes + v_total_almocos + v_total_jantas + v_total_pernoites
        valor_inesperado = (v_almoco_base * 3.0) if inesperada else 0.0
        valor_final = total_diarias_equipe + combustivel + pedagio + valor_inesperado
        total_diarias_pessoa = (total_diarias_equipe / qtd_pessoas) if qtd_pessoas > 0 else 0.0

        cafe_row["Total (R$)"] = f"R$ {v_total_cafes:.2f}"
        almoco_row["Total (R$)"] = f"R$ {v_total_almocos:.2f}"
        janta_row["Total (R$)"] = f"R$ {v_total_jantas:.2f}"
        pernoite_row["Total (R$)"] = f"R$ {v_total_pernoites:.2f}"
        
        cafe_row["Por Pessoa/dia"] = f"R$ {v_cafe:.2f}" if qtd_cafes > 0 else "-"
        almoco_row["Por Pessoa/dia"] = f"R$ {v_almoco:.2f}" if qtd_almocos > 0 else "-"
        janta_row["Por Pessoa/dia"] = f"R$ {v_janta:.2f}" if qtd_jantas > 0 else "-"
        pernoite_row["Por Pessoa/dia"] = f"R$ {v_pernoite:.2f}" if qtd_pernoites > 0 else "-"
        
        combust_row = {"Item": "COMBUSTÍVEL", "Total (R$)": f"R$ {combustivel:.2f}", "Por Pessoa/dia": "-"}
        pedagio_row = {"Item": "PEDÁGIO / TÁXI / ÔNIBUS", "Total (R$)": f"R$ {pedagio:.2f}", "Por Pessoa/dia": "-"}
        inesp_row = {"Item": "DESPESA INESPERADA", "Total (R$)": f"R$ {valor_inesperado:.2f}", "Por Pessoa/dia": "-"}
        capital_row = {"Item": "CAPITAL", "Total (R$)": "-", "Por Pessoa/dia": "-"}
        
        for c_dia in colunas_dias:
            combust_row[c_dia] = "-"
            pedagio_row[c_dia] = "-"
            inesp_row[c_dia] = "-"
            capital_row[c_dia] = "-"
            
        todas_linhas = [cafe_row, almoco_row, janta_row, pernoite_row, combust_row, pedagio_row, inesp_row, capital_row]
        tabela_itens = [r for r in todas_linhas if r.get("Total (R$)") not in ["R$ 0.00", "-"]]
        
        cargo_map = dict(zip(df_mot['Nome_Motorista'], df_mot['Cargo']))
        setor_map = dict(zip(df_mot['Nome_Motorista'], df_mot['Setor']))
        
        lista_detalhada_pessoas = []
        solic_cargo = cargo_map.get(solicitante, "PASSAGEIRO")
        solic_setor = setor_map.get(solicitante, "-")
        lista_detalhada_pessoas.append(f"{solicitante} ({solic_cargo})")
        
        for p in passageiros:
            lista_detalhada_pessoas.append(f"{p} ({cargo_map.get(p, 'PASSAGEIRO')})")
            
        str_passageiros_cargos = " , ".join(lista_detalhada_pessoas)
        
        dados_pedido = {
            'Numero_Adiantamento': numero_adiantamento,
            'Status': "APROVADO",
            'Solicitante': solicitante,
            'Cargo_Solicitante': solic_cargo,
            'Setor_Solicitante': solic_setor,
            'Nomes_Passageiros': str_passageiros_cargos,
            'Destino': destino,
            'Finalidade': finalidade,
            'Veiculo': veiculo,
            'Data_Saida': d_saida.strftime("%d/%m/%Y"),
            'Hora_Saida': h_saida.strftime("%H:%M"),
            'Data_Retorno': d_retorno.strftime("%d/%m/%Y"),
            'Hora_Retorno': h_retorno.strftime("%H:%M"),
            'Qtd_Pessoas': qtd_pessoas,
            'Qtd_Cafes': qtd_cafes * qtd_pessoas, 'V_Cafes': v_total_cafes,
            'Qtd_Almocos': qtd_almocos * qtd_pessoas, 'V_Almocos': v_total_almocos,
            'Qtd_Jantas': qtd_jantas * qtd_pessoas, 'V_Jantas': v_total_jantas,
            'Qtd_Pernoites': qtd_pernoites * qtd_pessoas, 'V_Pernoites': v_total_pernoites,
            'Valor_Diaria': total_diarias_pessoa,
            'Extras': combustivel + pedagio,
            'Inesperadas': valor_inesperado,
            'Valor_Final': valor_final
        }
        
        st.session_state['pedido_atual'] = {
            'dados_pedido': dados_pedido,
            'str_passageiros_cargos': str_passageiros_cargos,
            'tabela_itens': tabela_itens,
            'colunas_dias': colunas_dias
        }
        st.session_state['pedido_salvo'] = False
        
    if 'pedido_atual' in st.session_state:
        st.markdown("---")
        p = st.session_state['pedido_atual']
        renderizar_recibo_visual(p['dados_pedido'], p['str_passageiros_cargos'], p['tabela_itens'], p['colunas_dias'], is_nova=True)

with tab2:
    st.subheader("🖨️ Histórico do Banco de Dados")
    dados_db = pull_banco_dados()
    if len(dados_db) > 1:
        rows = dados_db[1:]
        opcoes = []
        for r in rows:
            carimbo = r[0] if len(r) > 0 else "?"
            num = r[20] if len(r) > 20 and r[20] else "S/N"
            solic = r[2] if len(r) > 2 else "?"
            dest = r[4] if len(r) > 4 else "?"
            opcoes.append(f"{carimbo} | Nº {num} | {solic} | {dest}")
            
        opcoes.reverse()
        rows.reverse()
        
        escolha = st.selectbox("Selecione um pedido para reimprimir:", [""] + opcoes)
        if escolha:
            idx = opcoes.index(escolha)
            r = rows[idx]
            
            json_str = None
            for val in reversed(r):
                if isinstance(val, str) and val.startswith("{") and "tabela" in val:
                    json_str = val
                    break
                    
            if json_str:
                try:
                    obj = json.loads(json_str)
                    st.markdown("---")
                    renderizar_recibo_visual(obj['dados'], obj['str_pass'], obj['tabela'], obj['colunas_dias'])
                except Exception as e:
                    st.error("Erro ao ler dados detalhados deste pedido.")
            else:
                st.warning("⚠️ Os pedidos mais antigos não possuem os dados detalhados para reimpressão automática da ficha (esse recurso foi implantado agora).")
    else:
        st.info("Nenhum pedido encontrado no Banco de Dados.")
