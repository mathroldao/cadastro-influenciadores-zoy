
import os
import uuid
from io import BytesIO
from datetime import datetime
from urllib.parse import urlencode

import gspread
import pandas as pd
import streamlit as st
from docx import Document
from google.oauth2.service_account import Credentials


APP_NAME = "Cadastro de Influenciadores | Zoy"
SHEET_NAME = "[ZOY] Cadastro Influenciadores | Base"
WORKSHEET_NAME = "Cadastro"
BASE_URL = "https://cadastro-zoy.streamlit.app/"
ADMIN_PASSWORD = "zoy"

COLUMNS = [
    "ID", "DATA_CADASTRO", "STATUS",
    "RAZAO_SOCIAL", "CNPJ", "REPRESENTANTE_LEGAL",
    "EMAIL", "TELEFONE", "INSTAGRAM", "TIKTOK",
    "CEP", "ENDERECO", "NUMERO", "COMPLEMENTO", "BAIRRO", "CIDADE", "ESTADO",
    "BANCO", "AGENCIA", "CONTA", "TIPO_CONTA", "PIX",
    "CLIENTE", "CAMPANHA", "VALOR_CACHE", "ESCOPO",
    "DIREITO_IMAGEM", "EXCLUSIVIDADE", "REPOST", "IMPULSIONAMENTO",
    "CONTRATO_GERADO", "CARTA_GERADA",
    "LINK_CONTRATO", "LINK_CARTA",
    "DATA_FORMALIZACAO",
    "D4SIGN_STATUS", "D4SIGN_ENVELOPE_ID",
    "OBSERVACOES"
]


st.set_page_config(
    page_title=APP_NAME,
    page_icon="💜",
    layout="wide",
    initial_sidebar_state="collapsed"
)


st.markdown("""
<style>
[data-testid="stSidebar"] {
    display: none;
}

.block-container {
    max-width: 1080px;
    padding-top: 42px;
    padding-bottom: 80px;
}

h1, h2, h3 {
    color: #2B2B36;
}

.zoy-header {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 18px;
    margin-bottom: 18px;
}

.zoy-title {
    font-size: 46px;
    line-height: 1.05;
    font-weight: 800;
    color: #2B2B36;
    margin: 0;
}

.zoy-subtitle {
    font-size: 17px;
    color: #4B5563;
    margin-top: 8px;
}

.zoy-pill {
    background: #F3E8FF;
    color: #7C3AED;
    padding: 8px 14px;
    border-radius: 999px;
    font-weight: 700;
    font-size: 14px;
    white-space: nowrap;
}

.zoy-card {
    border: 1px solid #E5E7EB;
    border-radius: 18px;
    padding: 26px;
    background: #FFFFFF;
    margin: 18px 0 24px 0;
    box-shadow: 0 8px 24px rgba(17, 24, 39, 0.04);
}

.zoy-info {
    background: #F3E8FF;
    border: 1px solid #E9D5FF;
    color: #4C1D95;
    border-radius: 16px;
    padding: 18px 22px;
    font-size: 16px;
    margin: 14px 0 24px 0;
}

.zoy-success {
    background: #ECFDF5;
    border: 1px solid #A7F3D0;
    color: #065F46;
    border-radius: 16px;
    padding: 22px;
    font-size: 18px;
    font-weight: 700;
    margin-top: 22px;
}

.stButton>button,
.stDownloadButton>button,
button[kind="primary"] {
    background: #7C3AED !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    height: 48px !important;
    font-weight: 800 !important;
}

.stButton>button:hover,
.stDownloadButton>button:hover {
    background: #6D28D9 !important;
    color: white !important;
}

input, textarea, [data-baseweb="select"] {
    border-radius: 12px !important;
}

hr {
    margin: 26px 0;
}

.small-muted {
    color: #6B7280;
    font-size: 14px;
}
</style>
""", unsafe_allow_html=True)


def is_admin():
    params = st.query_params
    return params.get("admin", "") == ADMIN_PASSWORD


def conectar():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
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


def formata_valor(v):
    try:
        return f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)


def achar_template(nome):
    for caminho in [f"templates/{nome}", nome]:
        if os.path.exists(caminho):
            return caminho
    return None


def replace_docx(template_path, dados):
    doc = Document(template_path)
    mapa = {f"{{{{{k}}}}}": str(v) for k, v in dados.items()}

    def replace_in_paragraph(paragraph):
        texto = paragraph.text
        for chave, valor in mapa.items():
            texto = texto.replace(chave, valor)

        if texto != paragraph.text:
            paragraph.clear()
            paragraph.add_run(texto)

    for p in doc.paragraphs:
        replace_in_paragraph(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p)

    arquivo = BytesIO()
    doc.save(arquivo)
    arquivo.seek(0)
    return arquivo


def header_publico():
    st.markdown("""
    <div class="zoy-header">
        <div>
            <div class="zoy-title">💜 Cadastro de Influenciadores | Zoy</div>
            <div class="zoy-subtitle">Preencha seus dados para formalização da campanha.</div>
        </div>
        <div class="zoy-pill">Tempo médio: 3 min</div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="zoy-info">
        Este cadastro será utilizado apenas para coleta de dados cadastrais e formalização interna pela equipe jurídica da Zoy.
        O influenciador não terá acesso aos documentos nesta etapa.
    </div>
    """, unsafe_allow_html=True)


def cadastro_publico():
    header_publico()

    params = st.query_params

    cliente_url = params.get("cliente", "")
    campanha_url = params.get("campanha", "")
    escopo_url = params.get("escopo", "")
    valor_url = params.get("valor", "0")
    direito_url = params.get("direito_imagem", "N/A")
    exclusividade_url = params.get("exclusividade", "N/A")
    repost_url = params.get("repost", "N/A")
    impulsionamento_url = params.get("impulsionamento", "N/A")

    try:
        valor_url = float(valor_url)
    except Exception:
        valor_url = 0.0

    with st.form("form_cadastro"):
        st.markdown('<div class="zoy-card">', unsafe_allow_html=True)

        st.subheader("🏢 Dados da Empresa")
        razao = st.text_input("Razão Social *")
        cnpj = st.text_input("CNPJ *")
        representante = st.text_input("Representante Legal *")

        st.divider()

        st.subheader("📱 Contato")
        email = st.text_input("E-mail *")
        telefone = st.text_input("Telefone / WhatsApp *")
        instagram = st.text_input("Instagram *")
        tiktok = st.text_input("TikTok")

        st.divider()

        st.subheader("📍 Endereço")
        col1, col2 = st.columns([1, 2])
        with col1:
            cep = st.text_input("CEP")
        with col2:
            endereco = st.text_input("Endereço")

        col3, col4 = st.columns([1, 2])
        with col3:
            numero = st.text_input("Número")
        with col4:
            complemento = st.text_input("Complemento")

        col5, col6, col7 = st.columns([2, 2, 1])
        with col5:
            bairro = st.text_input("Bairro")
        with col6:
            cidade = st.text_input("Cidade")
        with col7:
            estado = st.text_input("Estado")

        st.divider()

        st.subheader("🏦 Dados Bancários")
        col8, col9 = st.columns(2)
        with col8:
            banco = st.text_input("Banco")
            agencia = st.text_input("Agência")
            tipo_conta = st.selectbox("Tipo de Conta", ["Corrente", "Poupança", "Pagamento"])
        with col9:
            conta = st.text_input("Conta")
            pix = st.text_input("Chave Pix")

        st.divider()

        st.subheader("📄 Dados da Campanha")
        cliente = st.text_input("Cliente *", value=cliente_url)
        campanha = st.text_input("Campanha *", value=campanha_url)
        valor = st.number_input(
            "Valor do cachê negociado",
            min_value=0.0,
            step=100.0,
            value=valor_url
        )
        escopo = st.text_area("Escopo contratado *", value=escopo_url, height=130)

        col10, col11 = st.columns(2)
        with col10:
            direito = st.text_input("Direito de imagem", value=direito_url)
            repost = st.text_input("Repost", value=repost_url)
        with col11:
            exclusividade = st.text_input("Exclusividade", value=exclusividade_url)
            impulsionamento = st.text_input("Impulsionamento", value=impulsionamento_url)

        confirmar = st.checkbox("Confirmo que as informações preenchidas são verdadeiras.")
        enviar = st.form_submit_button("Enviar Cadastro")

        st.markdown("</div>", unsafe_allow_html=True)

    if enviar:
        obrigatorios = [razao, cnpj, representante, email, telefone, instagram, cliente, campanha, escopo]

        if not all(obrigatorios):
            st.error("Preencha todos os campos obrigatórios antes de enviar.")
            st.stop()

        if not confirmar:
            st.error("Confirme a veracidade das informações antes de enviar.")
            st.stop()

        protocolo = f"ZOY-{datetime.now().strftime('%Y%m%d')}-{str(uuid.uuid4())[:6].upper()}"

        dados = {
            "ID": protocolo,
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

        st.markdown(f"""
        <div class="zoy-success">
            Cadastro enviado com sucesso.<br>
            Protocolo: {protocolo}<br><br>
            O time jurídico da Zoy recebeu suas informações.
        </div>
        """, unsafe_allow_html=True)

        st.balloons()


def criar_link():
    st.subheader("🔗 Criar link para influenciador")

    with st.form("form_link"):
        col1, col2 = st.columns(2)

        with col1:
            link_cliente = st.text_input("Cliente")
            link_campanha = st.text_input("Campanha")
            link_valor = st.text_input("Valor do cachê")

        with col2:
            link_direito = st.text_input("Direito de imagem", value="N/A")
            link_exclusividade = st.text_input("Exclusividade", value="N/A")
            link_repost = st.text_input("Repost", value="N/A")
            link_impulsionamento = st.text_input("Impulsionamento", value="N/A")

        link_escopo = st.text_area("Escopo contratado")
        gerar = st.form_submit_button("Gerar Link")

    if gerar:
        query = urlencode({
            "cliente": link_cliente,
            "campanha": link_campanha,
            "valor": link_valor,
            "escopo": link_escopo,
            "direito_imagem": link_direito,
            "exclusividade": link_exclusividade,
            "repost": link_repost,
            "impulsionamento": link_impulsionamento
        })

        link_final = f"{BASE_URL}?{query}"

        st.success("Link criado com sucesso.")
        st.code(link_final)


def juridico():
    st.title("⚖️ Jurídico | Cadastro de Influenciadores")
    st.caption("Área interna da Zoy. Esta página não aparece para o influenciador.")

    tab1, tab2 = st.tabs(["🔗 Criar Link", "📋 Cadastros"])

    with tab1:
        criar_link()

    with tab2:
        registros = listar()

        if not registros:
            st.info("Nenhum cadastro encontrado.")
            return

        df = pd.DataFrame(registros)

        busca = st.text_input("Buscar por razão social, Instagram, cliente ou campanha")

        if busca:
            b = busca.lower()
            df = df[df.astype(str).apply(lambda x: x.str.lower().str.contains(b)).any(axis=1)]

        st.write(f"**{len(df)} cadastro(s) encontrado(s)**")

        for _, row in df.iterrows():
            titulo = f"{row.get('RAZAO_SOCIAL', '')} | {row.get('INSTAGRAM', '')} | {row.get('CLIENTE', '')} - {row.get('CAMPANHA', '')}"

            with st.expander(titulo):
                c1, c2 = st.columns(2)

                with c1:
                    st.write("### Empresa")
                    st.write(f"**Razão Social:** {row.get('RAZAO_SOCIAL', '')}")
                    st.write(f"**CNPJ:** {row.get('CNPJ', '')}")
                    st.write(f"**Representante:** {row.get('REPRESENTANTE_LEGAL', '')}")
                    st.write(f"**E-mail:** {row.get('EMAIL', '')}")
                    st.write(f"**Telefone:** {row.get('TELEFONE', '')}")
                    st.write(f"**Instagram:** {row.get('INSTAGRAM', '')}")
                    st.write(f"**TikTok:** {row.get('TIKTOK', '')}")

                    st.write("### Endereço")
                    st.write(f"{row.get('ENDERECO', '')}, {row.get('NUMERO', '')} {row.get('COMPLEMENTO', '')}")
                    st.write(f"{row.get('BAIRRO', '')} - {row.get('CIDADE', '')}/{row.get('ESTADO', '')}")
                    st.write(f"CEP: {row.get('CEP', '')}")

                    st.write("### Bancário")
                    st.write(f"**Banco:** {row.get('BANCO', '')}")
                    st.write(f"**Agência:** {row.get('AGENCIA', '')}")
                    st.write(f"**Conta:** {row.get('CONTA', '')}")
                    st.write(f"**Tipo:** {row.get('TIPO_CONTA', '')}")
                    st.write(f"**Pix:** {row.get('PIX', '')}")

                with c2:
                    st.write("### Campanha")
                    st.write(f"**Cliente:** {row.get('CLIENTE', '')}")
                    st.write(f"**Campanha:** {row.get('CAMPANHA', '')}")
                    st.write(f"**Valor:** R$ {row.get('VALOR_CACHE', '')}")
                    st.write(f"**Escopo:** {row.get('ESCOPO', '')}")

                    st.write("### Condições")
                    st.write(f"**Direito de imagem:** {row.get('DIREITO_IMAGEM', '')}")
                    st.write(f"**Exclusividade:** {row.get('EXCLUSIVIDADE', '')}")
                    st.write(f"**Repost:** {row.get('REPOST', '')}")
                    st.write(f"**Impulsionamento:** {row.get('IMPULSIONAMENTO', '')}")

                dados_doc = row.to_dict()
                dados_doc["DATA_EXTENSO"] = datetime.now().strftime("%d/%m/%Y")

                contrato_path = achar_template("contrato.docx")
                carta_path = achar_template("carta.docx")

                col_a, col_b = st.columns(2)

                with col_a:
                    if contrato_path:
                        contrato = replace_docx(contrato_path, dados_doc)
                        st.download_button(
                            "📄 Baixar Contrato",
                            contrato,
                            file_name=f"Contrato - {row.get('RAZAO_SOCIAL', '')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    else:
                        st.warning("Arquivo contrato.docx não encontrado.")

                with col_b:
                    if carta_path:
                        carta = replace_docx(carta_path, dados_doc)
                        st.download_button(
                            "📄 Baixar Carta",
                            carta,
                            file_name=f"Carta - {row.get('RAZAO_SOCIAL', '')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )
                    else:
                        st.warning("Arquivo carta.docx não encontrado.")


if is_admin():
    juridico()
else:
    cadastro_publico()
