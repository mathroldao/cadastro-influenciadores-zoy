import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from docx import Document
from io import BytesIO
from datetime import datetime
import uuid
import os

APP_NAME = "Cadastro de Influenciadores | Zoy"
SHEET_NAME = "[ZOY] Cadastro Influenciadores | Base"
WORKSHEET_NAME = "Cadastro"

COLUMNS = [
    "ID","DATA_CADASTRO","STATUS","RAZAO_SOCIAL","CNPJ","REPRESENTANTE_LEGAL",
    "EMAIL","TELEFONE","INSTAGRAM","TIKTOK","CEP","ENDERECO","NUMERO","COMPLEMENTO",
    "BAIRRO","CIDADE","ESTADO","BANCO","AGENCIA","CONTA","TIPO_CONTA","PIX",
    "CLIENTE","CAMPANHA","VALOR_CACHE","ESCOPO","DIREITO_IMAGEM","EXCLUSIVIDADE",
    "REPOST","IMPULSIONAMENTO","CONTRATO_GERADO","CARTA_GERADA","LINK_CONTRATO",
    "LINK_CARTA","DATA_FORMALIZACAO","D4SIGN_STATUS","D4SIGN_ENVELOPE_ID","OBSERVACOES"
]

st.set_page_config(page_title=APP_NAME, page_icon="💜", layout="wide")

st.markdown("""
<style>
.block-container{max-width:1100px;padding-top:30px}
h1,h2,h3{color:#2B2B36}
.stButton>button{background:#7C3AED;color:white;border:none;border-radius:10px;height:46px;font-weight:700}
.stButton>button:hover{background:#6D28D9;color:white}
.card{border:1px solid #e5e7eb;border-radius:16px;padding:24px;background:white;margin-bottom:18px}
</style>
""", unsafe_allow_html=True)

def conectar():
    scopes = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    client = gspread.authorize(creds)
    sheet = client.open(SHEET_NAME)
    try:
        ws = sheet.worksheet(WORKSHEET_NAME)
    except:
        ws = sheet.add_worksheet(title=WORKSHEET_NAME, rows=1000, cols=50)
    if ws.row_values(1) != COLUMNS:
        ws.clear()
        ws.append_row(COLUMNS)
    return ws

def listar():
    return conectar().get_all_records()

def salvar(dados):
    ws = conectar()
    ws.append_row([dados.get(c, "") for c in COLUMNS])

def replace_docx(template_path, dados):
    doc = Document(template_path)
    mapa = {f"{{{{{k}}}}}": str(v) for k, v in dados.items()}

    def replace_paragraph(p):
        texto = p.text
        for k, v in mapa.items():
            texto = texto.replace(k, v)
        if texto != p.text:
            p.clear()
            p.add_run(texto)

    for p in doc.paragraphs:
        replace_paragraph(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_paragraph(p)

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

def formata_valor(v):
    try:
        return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return str(v)

menu = st.sidebar.radio("Menu", ["Cadastro Público", "Jurídico"])

if menu == "Cadastro Público":
    st.title("💜 Cadastro de Influenciadores | Zoy")
    st.info("Este cadastro será utilizado para formalização contratual entre o influenciador e a Agência Zoy. Tempo médio: **3 minutos**.")

    with st.form("form_cadastro"):
        st.subheader("🏢 Dados da Empresa")
        razao = st.text_input("Razão Social *")
        cnpj = st.text_input("CNPJ *")
        representante = st.text_input("Representante Legal *")

        st.subheader("📱 Contato")
        email = st.text_input("E-mail *")
        telefone = st.text_input("Telefone / WhatsApp *")
        instagram = st.text_input("Instagram *")
        tiktok = st.text_input("TikTok")

        st.subheader("📍 Endereço")
        cep = st.text_input("CEP")
        endereco = st.text_input("Endereço")
        numero = st.text_input("Número")
        complemento = st.text_input("Complemento")
        bairro = st.text_input("Bairro")
        cidade = st.text_input("Cidade")
        estado = st.text_input("Estado")

        st.subheader("🏦 Dados Bancários")
        banco = st.text_input("Banco")
        agencia = st.text_input("Agência")
        conta = st.text_input("Conta")
        tipo_conta = st.selectbox("Tipo de Conta", ["Corrente", "Poupança", "Pagamento"])
        pix = st.text_input("Chave Pix")

        st.subheader("📄 Dados da Campanha")
        cliente = st.text_input("Cliente *")
        campanha = st.text_input("Campanha *")
        valor = st.number_input("Valor do cachê negociado", min_value=0.0, step=100.0)
        escopo = st.text_area("Escopo contratado *")
        direito = st.text_input("Direito de imagem", value="N/A")
        exclusividade = st.text_input("Exclusividade", value="N/A")
        repost = st.text_input("Repost", value="N/A")
        impulsionamento = st.text_input("Impulsionamento", value="N/A")

        enviar = st.form_submit_button("Enviar Cadastro")

    if enviar:
        dados = {
            "ID": str(uuid.uuid4()),
            "DATA_CADASTRO": datetime.now().strftime("%d/%m/%Y %H:%M"),
            "STATUS": "Pendente",
            "RAZAO_SOCIAL": razao,
            "CNPJ": cnpj,
            "REPRESENTANTE_LEGAL": representante,
            "EMAIL": email,
            "TELEFONE": telefone,
            "INSTAGRAM": instagram,
            "TIKTOK": tiktok,
            "CEP": cep,
            "ENDERECO": endereco,
            "NUMERO": numero,
            "COMPLEMENTO": complemento,
            "BAIRRO": bairro,
            "CIDADE": cidade,
            "ESTADO": estado,
            "BANCO": banco,
            "AGENCIA": agencia,
            "CONTA": conta,
            "TIPO_CONTA": tipo_conta,
            "PIX": pix,
            "CLIENTE": cliente,
            "CAMPANHA": campanha,
            "VALOR_CACHE": formata_valor(valor),
            "ESCOPO": escopo,
            "DIREITO_IMAGEM": direito,
            "EXCLUSIVIDADE": exclusividade,
            "REPOST": repost,
            "IMPULSIONAMENTO": impulsionamento,
            "CONTRATO_GERADO": "Não",
            "CARTA_GERADA": "Não",
            "LINK_CONTRATO": "",
            "LINK_CARTA": "",
            "DATA_FORMALIZACAO": "",
            "D4SIGN_STATUS": "",
            "D4SIGN_ENVELOPE_ID": "",
            "OBSERVACOES": ""
        }
        salvar(dados)
        st.success("Cadastro enviado com sucesso! O time jurídico da Zoy recebeu suas informações.")
        st.balloons()

if menu == "Jurídico":
    st.title("⚖️ Jurídico | Cadastros Recebidos")

    registros = listar()
    if not registros:
        st.info("Nenhum cadastro encontrado.")
        st.stop()

    df = pd.DataFrame(registros)

    busca = st.text_input("Buscar por razão social, Instagram, cliente ou campanha")
    if busca:
        b = busca.lower()
        df = df[df.astype(str).apply(lambda x: x.str.lower().str.contains(b)).any(axis=1)]

    for i, row in df.iterrows():
        with st.expander(f"{row['RAZAO_SOCIAL']} | {row['INSTAGRAM']} | {row['CLIENTE']} - {row['CAMPANHA']}"):
            c1, c2 = st.columns(2)

            with c1:
                st.write("### Empresa")
                st.write(f"**Razão Social:** {row['RAZAO_SOCIAL']}")
                st.write(f"**CNPJ:** {row['CNPJ']}")
                st.write(f"**Representante:** {row['REPRESENTANTE_LEGAL']}")
                st.write(f"**E-mail:** {row['EMAIL']}")
                st.write(f"**Telefone:** {row['TELEFONE']}")
                st.write(f"**Instagram:** {row['INSTAGRAM']}")
                st.write(f"**TikTok:** {row['TIKTOK']}")

            with c2:
                st.write("### Campanha")
                st.write(f"**Cliente:** {row['CLIENTE']}")
                st.write(f"**Campanha:** {row['CAMPANHA']}")
                st.write(f"**Valor:** R$ {row['VALOR_CACHE']}")
                st.write(f"**Escopo:** {row['ESCOPO']}")
                st.write(f"**Direito de imagem:** {row['DIREITO_IMAGEM']}")
                st.write(f"**Exclusividade:** {row['EXCLUSIVIDADE']}")
                st.write(f"**Repost:** {row['REPOST']}")
                st.write(f"**Impulsionamento:** {row['IMPULSIONAMENTO']}")

            dados_doc = row.to_dict()
            dados_doc["DATA_EXTENSO"] = datetime.now().strftime("%d/%m/%Y")

            col_a, col_b = st.columns(2)

            with col_a:
                if os.path.exists("templates/contrato.docx"):
                    contrato = replace_docx("templates/contrato.docx", dados_doc)
                    st.download_button(
                        "📄 Baixar Contrato",
                        contrato,
                        file_name=f"Contrato - {row['RAZAO_SOCIAL']}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.warning("Falta subir templates/contrato.docx")

            with col_b:
                if os.path.exists("templates/carta.docx"):
                    carta = replace_docx("templates/carta.docx", dados_doc)
                    st.download_button(
                        "📄 Baixar Carta",
                        carta,
                        file_name=f"Carta - {row['RAZAO_SOCIAL']}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                else:
                    st.warning("Falta subir templates/carta.docx")
