import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from docx import Document
from io import BytesIO
from datetime import datetime
from urllib.parse import urlencode
import uuid
import os
import re

APP_NAME = "Cadastro de Influenciadores | Zoy"
SHEET_NAME = "[ZOY] Cadastro Influenciadores | Base"
WORKSHEET_NAME = "Cadastro"
ADMIN_USER = "juridico"
ADMIN_PASSWORD = "zoy2026"
PUBLIC_BASE_URL = "https://cadastro-zoy.streamlit.app/"

COLUMNS = [
    "ID","DATA_CADASTRO","STATUS","RAZAO_SOCIAL","CNPJ","REPRESENTANTE_LEGAL",
    "EMAIL","TELEFONE","INSTAGRAM","TIKTOK","CEP","ENDERECO","NUMERO","COMPLEMENTO",
    "BAIRRO","CIDADE","ESTADO","BANCO","AGENCIA","CONTA","TIPO_CONTA","PIX",
    "CLIENTE","CAMPANHA","VALOR_CACHE","ESCOPO","DIREITO_IMAGEM","EXCLUSIVIDADE",
    "REPOST","IMPULSIONAMENTO","CONTRATO_GERADO","CARTA_GERADA","LINK_CONTRATO",
    "LINK_CARTA","DATA_FORMALIZACAO","D4SIGN_STATUS","D4SIGN_ENVELOPE_ID","OBSERVACOES"
]

st.set_page_config(page_title=APP_NAME, page_icon="📄", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

.block-container {
    max-width: 1180px;
    padding-top: 42px;
    padding-bottom: 72px;
}

h1 {
    font-weight: 800 !important;
    letter-spacing: -1.5px;
    color: #272735;
}

h2, h3 {
    color: #272735;
    font-weight: 800 !important;
}

div[data-testid="stSidebar"] {
    background: #F7F2FF;
}

.stButton > button {
    background: linear-gradient(135deg, #7C3AED 0%, #A855F7 100%);
    color: white;
    border: none;
    border-radius: 14px;
    min-height: 48px;
    padding: 0 24px;
    font-weight: 800;
    box-shadow: 0 12px 26px rgba(124,58,237,.20);
}

.stButton > button:hover {
    background: linear-gradient(135deg, #6D28D9 0%, #9333EA 100%);
    color: white;
    border: none;
}

div[data-testid="stDownloadButton"] > button {
    background: white;
    color: #6D28D9;
    border: 1px solid #DDD6FE;
    border-radius: 14px;
    min-height: 48px;
    font-weight: 800;
}

div[data-testid="stDownloadButton"] > button:hover {
    background: #F5F3FF;
    color: #5B21B6;
    border: 1px solid #C4B5FD;
}

.card {
    border: 1px solid #E7E3F3;
    background: #FFFFFF;
    border-radius: 22px;
    padding: 24px;
    box-shadow: 0 12px 30px rgba(39,39,53,.04);
    margin-bottom: 18px;
}

.metric-card {
    border: 1px solid #E7E3F3;
    background: #FFFFFF;
    border-radius: 22px;
    padding: 22px;
    min-height: 116px;
    box-shadow: 0 12px 30px rgba(39,39,53,.04);
}

.metric-label {
    color: #737387;
    text-transform: uppercase;
    font-size: 12px;
    font-weight: 800;
    letter-spacing: .6px;
}

.metric-value {
    color: #272735;
    font-size: 34px;
    font-weight: 800;
    margin-top: 10px;
}

.small-muted {
    color: #737387;
    font-size: 14px;
}

.zoy-pill {
    display: inline-block;
    padding: 7px 12px;
    border-radius: 999px;
    background: #F5F3FF;
    color: #6D28D9;
    font-size: 12px;
    font-weight: 800;
}

.login-box {
    max-width: 430px;
    margin: 40px auto;
    border: 1px solid #E7E3F3;
    border-radius: 24px;
    padding: 32px;
    background: white;
    box-shadow: 0 20px 50px rgba(39,39,53,.08);
}

hr {
    margin-top: 28px;
    margin-bottom: 28px;
}

input, textarea, select {
    border-radius: 12px !important;
}
</style>
""", unsafe_allow_html=True)


# ======================
# Google Sheets
# ======================

@st.cache_resource(ttl=600)
def conectar():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )
    client = gspread.authorize(creds)
    sheet = client.open(SHEET_NAME)

    try:
        ws = sheet.worksheet(WORKSHEET_NAME)
    except Exception:
        ws = sheet.add_worksheet(title=WORKSHEET_NAME, rows=1000, cols=50)

    if ws.row_values(1) != COLUMNS:
        ws.clear()
        ws.append_row(COLUMNS)

    return ws


def listar():
    return conectar().get_all_records()


def salvar(dados):
    conectar().append_row([dados.get(c, "") for c in COLUMNS])


# ======================
# Helpers
# ======================

def get_query_value(key, default=""):
    try:
        value = st.query_params.get(key, default)
        if isinstance(value, list):
            return value[0] if value else default
        return value if value is not None else default
    except Exception:
        return default


def is_admin_url():
    return get_query_value("admin", "") == "zoy"


def formata_valor(v):
    try:
        return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)


def safe_text(value):
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value)


def achar_template(nome):
    paths = [
        os.path.join("templates", nome),
        nome
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return None


def normalizar_dados_doc(row):
    dados = {}
    for col in COLUMNS:
        dados[col] = safe_text(row.get(col, ""))

    dados["VALOR"] = dados.get("VALOR_CACHE", "")
    dados["DATA_EXTENSO"] = datetime.now().strftime("%d/%m/%Y")

    # Compatibilidade com variações comuns de placeholder
    dados["RAZÃO_SOCIAL"] = dados.get("RAZAO_SOCIAL", "")
    dados["DIREITO_DE_IMAGEM"] = dados.get("DIREITO_IMAGEM", "")
    dados["PRAZO_PAGAMENTO"] = dados.get("PRAZO_PAGAMENTO", "30")

    return dados


def replace_in_paragraph(paragraph, replacements):
    """
    Substitui placeholders mesmo quando o Word quebra o texto em vários runs.
    Mantém a formatação do primeiro run do parágrafo/célula.
    """
    if not paragraph.runs:
        return

    original_text = "".join(run.text for run in paragraph.runs)
    new_text = original_text

    for key, value in replacements.items():
        new_text = new_text.replace(key, value)

    if new_text != original_text:
        for run in paragraph.runs:
            run.text = ""
        paragraph.runs[0].text = new_text


def replace_in_table(table, replacements):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                replace_in_paragraph(paragraph, replacements)
            for nested_table in cell.tables:
                replace_in_table(nested_table, replacements)


def replace_docx(template_path, dados):
    doc = Document(template_path)
    replacements = {f"{{{{{k}}}}}": safe_text(v) for k, v in dados.items()}

    for paragraph in doc.paragraphs:
        replace_in_paragraph(paragraph, replacements)

    for table in doc.tables:
        replace_in_table(table, replacements)

    for section in doc.sections:
        for paragraph in section.header.paragraphs:
            replace_in_paragraph(paragraph, replacements)
        for table in section.header.tables:
            replace_in_table(table, replacements)

        for paragraph in section.footer.paragraphs:
            replace_in_paragraph(paragraph, replacements)
        for table in section.footer.tables:
            replace_in_table(table, replacements)

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio


def top_header(title, subtitle=None):
    st.markdown(f"""
    <div style="margin-bottom: 28px;">
        <div class="zoy-pill">Zoy Jurídico</div>
        <h1 style="margin-top: 12px; margin-bottom: 6px;">{title}</h1>
        {f'<p class="small-muted">{subtitle}</p>' if subtitle else ''}
    </div>
    """, unsafe_allow_html=True)


# ======================
# Público
# ======================

def cadastro_publico():
    st.title("Cadastro de Influenciadores | Zoy")
    st.info("Este cadastro será utilizado para formalização contratual entre o influenciador e a Agência Zoy. Tempo médio: **3 minutos**.")

    cliente_pre = get_query_value("cliente", "")
    campanha_pre = get_query_value("campanha", "")
    valor_pre = get_query_value("valor", "0")
    escopo_pre = get_query_value("escopo", "")
    direito_pre = get_query_value("direito_imagem", "N/A")
    exclusividade_pre = get_query_value("exclusividade", "N/A")
    repost_pre = get_query_value("repost", "N/A")
    impulsionamento_pre = get_query_value("impulsionamento", "N/A")

    try:
        valor_pre_float = float(str(valor_pre).replace(".", "").replace(",", "."))
    except Exception:
        valor_pre_float = 0.0

    with st.form("form_cadastro"):
        st.markdown('<div class="card">', unsafe_allow_html=True)

        st.subheader("Dados da Empresa")
        razao = st.text_input("Razão Social *")
        cnpj = st.text_input("CNPJ *")
        representante = st.text_input("Representante Legal *")

        st.divider()

        st.subheader("Contato")
        email = st.text_input("E-mail *")
        telefone = st.text_input("Telefone / WhatsApp *")
        instagram = st.text_input("Instagram *")
        tiktok = st.text_input("TikTok")

        st.divider()

        st.subheader("Endereço")
        c1, c2 = st.columns([1, 2])
        with c1:
            cep = st.text_input("CEP")
        with c2:
            endereco = st.text_input("Endereço")

        c3, c4, c5 = st.columns([1, 1, 1])
        with c3:
            numero = st.text_input("Número")
        with c4:
            complemento = st.text_input("Complemento")
        with c5:
            bairro = st.text_input("Bairro")

        c6, c7 = st.columns([2, 1])
        with c6:
            cidade = st.text_input("Cidade")
        with c7:
            estado = st.text_input("Estado")

        st.divider()

        st.subheader("Dados Bancários")
        c8, c9 = st.columns(2)
        with c8:
            banco = st.text_input("Banco")
            agencia = st.text_input("Agência")
            pix = st.text_input("Chave Pix")
        with c9:
            conta = st.text_input("Conta")
            tipo_conta = st.selectbox("Tipo de Conta", ["Corrente", "Poupança", "Pagamento"])

        st.divider()

        st.subheader("Dados da Campanha")
        c10, c11 = st.columns(2)
        with c10:
            cliente = st.text_input("Cliente *", value=cliente_pre)
            valor = st.number_input("Valor do cachê negociado", min_value=0.0, step=100.0, value=valor_pre_float)
        with c11:
            campanha = st.text_input("Campanha *", value=campanha_pre)

        escopo = st.text_area("Escopo contratado *", value=escopo_pre, height=130)

        c12, c13 = st.columns(2)
        with c12:
            direito = st.text_input("Direito de imagem", value=direito_pre)
            repost = st.text_input("Repost", value=repost_pre)
        with c13:
            exclusividade = st.text_input("Exclusividade", value=exclusividade_pre)
            impulsionamento = st.text_input("Impulsionamento", value=impulsionamento_pre)

        enviar = st.form_submit_button("Enviar Cadastro")
        st.markdown('</div>', unsafe_allow_html=True)

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


# ======================
# Login / Jurídico
# ======================

def login():
    top_header("Cadastro de Influenciadores", "Área interna da Zoy para gestão de cadastros e documentos.")

    st.markdown('<div class="login-box">', unsafe_allow_html=True)
    st.subheader("Login do Jurídico")
    user = st.text_input("Usuário")
    password = st.text_input("Senha", type="password")

    if st.button("Entrar", use_container_width=True):
        if user == ADMIN_USER and password == ADMIN_PASSWORD:
            st.session_state["logado"] = True
            st.rerun()
        else:
            st.error("Usuário ou senha inválidos.")

    st.markdown('</div>', unsafe_allow_html=True)


def tela_link():
    st.subheader("Criar link para influenciador")
    st.caption("Preencha os dados da campanha e envie o link gerado para o influenciador.")

    with st.form("form_link"):
        c1, c2 = st.columns(2)

        with c1:
            cliente = st.text_input("Cliente")
            campanha = st.text_input("Campanha")
            valor = st.text_input("Valor do cachê")

        with c2:
            direito = st.text_input("Direito de imagem", value="N/A")
            exclusividade = st.text_input("Exclusividade", value="N/A")
            repost = st.text_input("Repost", value="N/A")
            impulsionamento = st.text_input("Impulsionamento", value="N/A")

        escopo = st.text_area("Escopo contratado", height=130)

        gerar = st.form_submit_button("Gerar link")

    if gerar:
        query = urlencode({
            "cliente": cliente,
            "campanha": campanha,
            "valor": valor,
            "escopo": escopo,
            "direito_imagem": direito,
            "exclusividade": exclusividade,
            "repost": repost,
            "impulsionamento": impulsionamento
        })
        link = f"{PUBLIC_BASE_URL}?{query}"

        st.success("Link criado com sucesso.")
        st.code(link, language="text")


def tela_cadastros():
    registros = listar()
    if not registros:
        st.info("Nenhum cadastro encontrado.")
        return

    df = pd.DataFrame(registros)

    total = len(df)
    pendentes = len(df[df["STATUS"].astype(str).str.lower() == "pendente"]) if "STATUS" in df.columns else 0
    contratos = len(df[df["CONTRATO_GERADO"].astype(str).str.lower() == "sim"]) if "CONTRATO_GERADO" in df.columns else 0
    cartas = len(df[df["CARTA_GERADA"].astype(str).str.lower() == "sim"]) if "CARTA_GERADA" in df.columns else 0

    m1, m2, m3, m4 = st.columns(4)
    for col, label, value in [
        (m1, "Total de cadastros", total),
        (m2, "Pendentes", pendentes),
        (m3, "Contratos gerados", contratos),
        (m4, "Cartas geradas", cartas),
    ]:
        with col:
            st.markdown(f"""
            <div class="metric-card">
                <div class="metric-label">{label}</div>
                <div class="metric-value">{value}</div>
            </div>
            """, unsafe_allow_html=True)

    st.write("")
    c1, c2 = st.columns([3, 1])

    with c1:
        busca = st.text_input("Buscar por razão social, Instagram, cliente ou campanha")

    with c2:
        status = st.selectbox("Status", ["Todos", "Pendente", "Formalizado", "Finalizado"])

    if busca:
        b = busca.lower()
        df = df[df.astype(str).apply(lambda x: x.str.lower().str.contains(b)).any(axis=1)]

    if status != "Todos" and "STATUS" in df.columns:
        df = df[df["STATUS"].astype(str) == status]

    st.write("")

    for idx, row in df.iterrows():
        titulo = f"{row.get('RAZAO_SOCIAL','')} | {row.get('INSTAGRAM','')} | {row.get('CLIENTE','')} - {row.get('CAMPANHA','')}"
        with st.expander(titulo):
            c1, c2 = st.columns(2)

            with c1:
                st.markdown("### Empresa")
                st.write(f"**Razão Social:** {row.get('RAZAO_SOCIAL','')}")
                st.write(f"**CNPJ:** {row.get('CNPJ','')}")
                st.write(f"**Representante:** {row.get('REPRESENTANTE_LEGAL','')}")
                st.write(f"**E-mail:** {row.get('EMAIL','')}")
                st.write(f"**Telefone:** {row.get('TELEFONE','')}")
                st.write(f"**Instagram:** {row.get('INSTAGRAM','')}")
                st.write(f"**TikTok:** {row.get('TIKTOK','')}")

                st.markdown("### Endereço")
                st.write(f"{row.get('ENDERECO','')}, {row.get('NUMERO','')} {row.get('COMPLEMENTO','')}")
                st.write(f"{row.get('BAIRRO','')} - {row.get('CIDADE','')}/{row.get('ESTADO','')}")
                st.write(f"CEP: {row.get('CEP','')}")

            with c2:
                st.markdown("### Campanha")
                st.write(f"**Cliente:** {row.get('CLIENTE','')}")
                st.write(f"**Campanha:** {row.get('CAMPANHA','')}")
                st.write(f"**Valor:** R$ {row.get('VALOR_CACHE','')}")
                st.write(f"**Escopo:** {row.get('ESCOPO','')}")
                st.write(f"**Direito de imagem:** {row.get('DIREITO_IMAGEM','')}")
                st.write(f"**Exclusividade:** {row.get('EXCLUSIVIDADE','')}")
                st.write(f"**Repost:** {row.get('REPOST','')}")
                st.write(f"**Impulsionamento:** {row.get('IMPULSIONAMENTO','')}")

                st.markdown("### Bancário")
                st.write(f"**Banco:** {row.get('BANCO','')}")
                st.write(f"**Agência:** {row.get('AGENCIA','')}")
                st.write(f"**Conta:** {row.get('CONTA','')}")
                st.write(f"**Tipo:** {row.get('TIPO_CONTA','')}")
                st.write(f"**Pix:** {row.get('PIX','')}")

            dados_doc = normalizar_dados_doc(row.to_dict())

            contrato_path = achar_template("contrato.docx")
            carta_path = achar_template("carta.docx")

            col_a, col_b = st.columns(2)

            with col_a:
                if contrato_path:
                    contrato = replace_docx(contrato_path, dados_doc)
                    st.download_button(
                        "Baixar Contrato",
                        contrato,
                        file_name=f"Contrato - {row.get('RAZAO_SOCIAL','Influenciador')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"contrato_{row.get('ID', idx)}",
                        use_container_width=True
                    )
                else:
                    st.warning("Arquivo contrato.docx não encontrado.")

            with col_b:
                if carta_path:
                    carta = replace_docx(carta_path, dados_doc)
                    st.download_button(
                        "Baixar Carta",
                        carta,
                        file_name=f"Carta - {row.get('RAZAO_SOCIAL','Influenciador')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"carta_{row.get('ID', idx)}",
                        use_container_width=True
                    )
                else:
                    st.warning("Arquivo carta.docx não encontrado.")


def juridico():
    if not st.session_state.get("logado", False):
        login()
        return

    top_header("Cadastro de Influenciadores", "Área interna da Zoy para gestão de cadastros e documentos.")

    col1, col2, col3 = st.columns([1.1, 1.4, .7])

    with col1:
        if st.button("Cadastros", use_container_width=True):
            st.session_state["aba_admin"] = "Cadastros"

    with col2:
        if st.button("Link para Influenciador", use_container_width=True):
            st.session_state["aba_admin"] = "Link"

    with col3:
        if st.button("Sair", use_container_width=True):
            st.session_state["logado"] = False
            st.rerun()

    if "aba_admin" not in st.session_state:
        st.session_state["aba_admin"] = "Cadastros"

    st.divider()

    if st.session_state["aba_admin"] == "Link":
        tela_link()
    else:
        st.subheader("Cadastros")
        tela_cadastros()


if is_admin_url():
    juridico()
else:
    cadastro_publico()
