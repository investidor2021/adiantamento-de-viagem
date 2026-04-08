import streamlit as st
import pandas as pd
import gspread
import base64
import os
import unicodedata
from datetime import datetime, date
from fpdf import FPDF

# --- SETUP E UTILS ---
st.set_page_config(page_title="Formulário de Solicitação de Viagem", page_icon="📝", layout="centered")

SHEET_ID = '17LKTIN_wEMCqbN7L04cT2fqTsCABCJDsgLY7wKP88xo'
CRED_FILE = 'credenciais.json'

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
        self.cell(0, 8, 'SOLICITACAO DE ADIANTAMENTO DE VIAGEM', 0, 1, 'C')
        self.set_font('Arial', '', 11)
        self.cell(28)
        self.cell(0, 6, 'PREFEITURA MUNICIPAL', 0, 1, 'C')
        self.ln(10)

def cl(t): return str(t).encode('latin-1', 'replace').decode('latin-1')

# --- AUTENTICACAO GSPREAD ---
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
def pull_sheet_data_basico():
    gc = init_gspread()
    if not gc: return None, None, None
    sh = gc.open_by_key(SHEET_ID)
    ws_mun = sh.worksheet("Mun")
    df_mun = pd.DataFrame(ws_mun.get_all_values())
    
    df_cidades = df_mun.iloc[1:, 1:6].copy()
    df_cidades.columns = ['Cod', 'Municipio', 'UF', 'Populacao', 'Porte']
    df_cidades = df_cidades[df_cidades['Municipio'].str.strip() != '']
    df_cidades['Municipio_UF'] = df_cidades['Municipio'] + " - " + df_cidades['UF']
    df_cidades['Municipio_UF'] = df_cidades['Municipio_UF'].apply(remove_accents).str.upper()
    
    df_mot = df_mun.iloc[1:, 14:18].copy()
    df_mot.columns = ['Matricula', 'Nome_Motorista', 'Cargo', 'Setor']
    df_mot = df_mot[df_mot['Nome_Motorista'].str.strip() != '']
    
    df_veic = df_mun.iloc[1:, 25:29].copy()
    df_veic.columns = ['Placa', 'Veiculo', 'Ano', 'Combustivel']
    df_veic = df_veic[df_veic['Placa'].str.strip() != '']
    df_veic['Label'] = df_veic['Placa'] + " - " + df_veic['Veiculo']
    
    return df_cidades, df_mot, df_veic

def save_formulario(dados):
    gc = init_gspread()
    if not gc: return False
    try:
        sh = gc.open_by_key(SHEET_ID)
        ws = sh.worksheet("Formulario")
        
        row = [
            datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
            dados['Solicitante'],
            dados['Cargo_Solicitante'],
            dados['Setor_Solicitante'],
            dados['Destino'],
            dados['Finalidade'],
            dados['Veiculo'],
            dados['Data_Saida'],
            dados['Hora_Saida'],
            dados['Data_Retorno'],
            dados['Hora_Retorno'],
            dados['Nomes_Passageiros'],
            "PENDENTE"
        ]
        ws.append_row(row, value_input_option='USER_ENTERED')
        return True
    except Exception as e:
        st.error(f"Erro ao salvar na aba Formulario: {e}")
        return False

def criar_pdf_solicitacao_b64(dados):
    pdf = PDF()
    pdf.add_page()
    pdf.ln(12)
    pdf.set_font('Arial', '', 10)
    
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
    pdf.cell(0, 6, cl(f"Setor/Departamento: {dados.get('Setor_Solicitante', '-')}"), 0, 1)
    
    separator()
    write_wrapped(f"Destino: {dados['Destino']}")
    write_wrapped(f"Finalidade da Viagem: {dados['Finalidade']}")
    write_wrapped(f"Periodo: Saida em {dados['Data_Saida']} (as {dados['Hora_Saida']}) e Retorno em {dados['Data_Retorno']} (as {dados['Hora_Retorno']})")
    
    separator()
    v_parts = dados['Veiculo'].split(" - ", 1)
    if len(v_parts) == 2:
        placa, nome_veic = v_parts[0], v_parts[1]
    else:
        placa, nome_veic = "-", dados['Veiculo']
        
    pdf.cell(0, 6, cl(f"Veiculo: {placa} | Modelo: {nome_veic}"), 0, 1)
    if dados['Nomes_Passageiros']:
        write_wrapped(f"Passageiros Acompanhantes: {dados['Nomes_Passageiros']}")
    else:
        pdf.cell(0, 6, cl("Passageiros Acompanhantes: Nenhum"), 0, 1)
    
    separator()
    
    pdf.ln(10)
    pdf.set_font('Arial', 'I', 10)
    pdf.cell(0, 6, cl("Declaro ter ciencia de que esta ficha e apenas uma SOLICITACAO PREVIA,"), 0, 1, 'C')
    pdf.cell(0, 6, cl("estando sujeita a aprovacao da chefia e calculo final pelo departamento de financas."), 0, 1, 'C')
    
    pdf.ln(30)
    pdf.set_font('Arial', '', 10)
    pdf.cell(190, 6, "________________________________________", 0, 1, 'C')
    pdf.cell(190, 5, cl("Assinatura do Solicitante"), 0, 1, 'C')
    
    pdf.ln(20)
    pdf.cell(190, 6, "________________________________________", 0, 1, 'C')
    pdf.cell(190, 5, cl("Assinatura da Chefia Imediata (Aprovacao)"), 0, 1, 'C')
    
    res = pdf.output(dest='S')
    if isinstance(res, str): res = res.encode('latin1')
    return base64.b64encode(res).decode('ascii')


# --- INTERFACE PRINCIPAL ---

st.title("📄 Formulário Novo Pedido de Viagem")
st.markdown("Preencha cuidadosamente os dados abaixo para solicitar sua viagem. **Esta etapa não calcula valores financeiros**, serve apenas para a chefia autorizar sua liberação.")

df_cidades, df_mot, df_veic = pull_sheet_data_basico()

if df_cidades is None:
    st.error("Não foi possível carregar os dados. Verifique a conexão com o Google Sheets.")
    st.stop()

with st.form("form_solicitacao"):
    st.subheader("👤 Seu Cadastro")
    
    lista_pessoas = sorted(df_mot['Nome_Motorista'].tolist())
    solicitante = st.selectbox("Você (Solicitante/Motorista)", lista_pessoas)
    
    cargo_map = dict(zip(df_mot['Nome_Motorista'], df_mot['Cargo']))
    setor_map = dict(zip(df_mot['Nome_Motorista'], df_mot['Setor']))
    
    st.subheader("📍 Detalhes da Rota")
    col1, col2 = st.columns(2)
    destino = col1.selectbox("Município de Destino", df_cidades['Municipio_UF'].tolist())
    finalidade = col2.text_input("Qual a Finalidade da Viagem?", placeholder="Ex: Reunião do CONDERG")
    
    st.subheader("🕒 Cronograma")
    c1, c2, c3, c4 = st.columns(4)
    d_saida = c1.date_input("Data de Saída", format="DD/MM/YYYY")
    h_saida = c2.time_input("Horário de Saída", value=None)
    d_retorno = c3.date_input("Data de Retorno", format="DD/MM/YYYY")
    h_retorno = c4.time_input("Horário de Retorno", value=None)
    
    st.subheader("👥 Equipe e Logística")
    c5, c6 = st.columns(2)
    veiculo = c5.selectbox("Veículo que será utilizado", df_veic['Label'].tolist())
    passageiros = c6.multiselect("Nomes dos passageiros (se houver)", lista_pessoas)
    
    st.divider()
    submetido = st.form_submit_button("Gerar Solicitação e Imprimir", type="primary", use_container_width=True)

if submetido:
    if not finalidade:
        st.warning("⚠️ Por favor, preencha a Finalidade da Viagem.")
    elif h_saida is None or h_retorno is None:
        st.warning("⚠️ Por favor, preencha corretamente os horários de Saída e Retorno.")
    else:
        # Preparar dados
        str_pass = ", ".join(passageiros) if passageiros else ""
        solic_cargo = cargo_map.get(solicitante, "PASSAGEIRO")
        solic_setor = setor_map.get(solicitante, "")
        
        dados = {
            'Solicitante': solicitante,
            'Cargo_Solicitante': solic_cargo,
            'Setor_Solicitante': solic_setor,
            'Destino': destino,
            'Finalidade': finalidade,
            'Veiculo': veiculo,
            'Data_Saida': d_saida.strftime("%d/%m/%Y"),
            'Hora_Saida': h_saida.strftime("%H:%M"),
            'Data_Retorno': d_retorno.strftime("%d/%m/%Y"),
            'Hora_Retorno': h_retorno.strftime("%H:%M"),
            'Nomes_Passageiros': str_pass
        }
        
        with st.spinner("Salvando na Aba Formulario e Gerando Impressão..."):
            sucesso = save_formulario(dados)
            
        if sucesso:
            st.success("✅ Solicitação gravada com sucesso na aba 'Formulario'!")
            try:
                b64_pdf = criar_pdf_solicitacao_b64(dados)
                href = f'<a href="data:application/pdf;base64,{b64_pdf}" download="Solicitacao_Viagem_{solicitante}.pdf" target="_blank"><button style="width:100%; padding:15px; font-size:18px; font-weight:bold; cursor:pointer; background-color:#2e86c1; color:white; border:none; border-radius:5px;">📥 BAIXAR PEDIDO PRELIMINAR (PDF)</button></a>'
                st.markdown(href, unsafe_allow_html=True)
                st.info("⬆️ Baixe o PDF acima, assine e entregue à sua chefia/departamento de finanças.")
            except Exception as e:
                st.error(f"Erro ao gerar o PDF de solicitação: {e}")
