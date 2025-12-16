import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata
import re

# ============================
# CONFIG DO APP
# ============================

st.set_page_config(
    page_title="Conversor W4 -> Omie",
    layout="wide"
)

# ============================
# CSS – NATAL 🎄
# ============================

st.markdown("""
<style>
body {
    background-image: url('https://images.unsplash.com/photo-1513670800287-29d3b6b4a3d8');
    background-size: cover;
    background-repeat: no-repeat;
    background-attachment: fixed;
}
.block-container {
    backdrop-filter: blur(6px);
    background: rgba(255, 255, 255, 0.85);
    padding: 2rem;
    border-radius: 12px;
}
h1 {
    text-align: center;
    color: #003366 !important; /* Azul Omie */
    font-weight: 900 !important;
}
.stButton>button {
    background-color: #00C8FF; /* Azul claro Omie */
    color: white;
    border-radius: 10px;
    padding: 0.6rem 1.2rem;
    border: none;
    font-weight: bold;
}
.stButton>button:hover {
    background-color: #0099cc;
}
</style>
""", unsafe_allow_html=True)

# ============================
# FUNÇÕES AUXILIARES
# ============================

def normalize_text(texto):
    texto = str(texto).lower().strip()
    texto = ''.join(c for c in unicodedata.normalize('NFKD', texto) if not unicodedata.combining(c))
    texto = re.sub(r'[^a-z0-9]+', ' ', texto)
    return re.sub(r'\s+', ' ', texto).strip()

def preparar_categorias(df_cat):
    col = "Descrição da categoria financeira"
    df = df_cat.copy()

    def tirar_codigo(txt):
        txt = str(txt).strip()
        parts = txt.split(" ", 1)
        if len(parts) == 2 and any(ch.isdigit() for ch in parts[0]):
            return parts[1].strip()
        return txt

    df["nome_base"] = df[col].apply(tirar_codigo).apply(normalize_text)
    return df

def formatar_data_coluna(serie):
    datas = pd.to_datetime(serie, errors="coerce")
    return datas.dt.strftime("%d/%m/%Y")

def converter_valor_omie(valor_str):
    """
    Converte para string brasileira, mas SEMPRE positivo,
    pois o arquivo é de Contas a Pagar.
    """
    if pd.isna(valor_str):
        return ""
    
    # Se já for número
    if isinstance(valor_str, (int, float)):
        val = abs(valor_str)
        return f"{val:.2f}".replace(".", ",")
    
    # Se for string
    base = str(valor_str).strip()
    # Remove qualquer sinal existente para garantir absoluto
    base = base.replace("R$", "").replace(" ", "")
    
    # Tenta tratar formato brasileiro 1.000,00 vs americano 1000.00
    try:
        if "," in base and "." in base: # Ex: 1.000,00
            base = base.replace(".", "").replace(",", ".")
        elif "," in base: # Ex: 1000,00
            base = base.replace(",", ".")
        
        val_float = float(base)
        val_abs = abs(val_float)
        return f"{val_abs:.2f}".replace(".", ",")
    except:
        return base # Retorna original se falhar

# ============================
# FUNÇÃO PRINCIPAL
# ============================

def converter_w4_para_omie(df_w4, df_categorias_prep):

    if "Detalhe Conta / Objeto" not in df_w4.columns:
        raise ValueError("Coluna 'Detalhe Conta / Objeto' não existe no W4.")

    col_cat = "Detalhe Conta / Objeto"

    # Filtra transferências internas
    df = df_w4.loc[
        ~df_w4[col_cat].astype(str).str.contains("Transferência Entre Disponíveis", case=False, na=False)
    ].copy()

    # --- Cruzamento de Categorias ---
    col_desc_cat = "Descrição da categoria financeira"
    df["nome_base_w4"] = df[col_cat].astype(str).apply(normalize_text)

    df = df.merge(
        df_categorias_prep[["nome_base", col_desc_cat]],
        left_on="nome_base_w4",
        right_on="nome_base",
        how="left"
    )

    df["Categoria_final"] = df[col_desc_cat].where(df[col_desc_cat].notna(), df[col_cat])

    # --- Tratamento de Empréstimos e Processos ---
    fluxo = df.get("Fluxo", pd.Series("", index=df.index)).astype(str).str.lower()
    fluxo_vazio = fluxo.str.strip().isin(["", "none", "nan"])
    cond_fluxo_receita = fluxo.str.contains("receita", na=False)
    cond_fluxo_despesa = fluxo.str.contains("despesa", na=False)

    proc_original = df.get("Processo", pd.Series("", index=df.index)).astype(str)
    proc = proc_original.str.lower()
    proc = proc.apply(lambda t: unicodedata.normalize("NFKD", t).encode("ascii", "ignore").decode("ascii"))

    pessoa = df.get("Pessoa", pd.Series("", index=df.index)).astype(str)

    cond_emprestimo = proc.str.contains("emprestimo", na=False)
    cond_pag_emp = proc.str.contains("pagamento", na=False) & cond_emprestimo
    cond_rec_emp = proc.str.contains("recebimento", na=False) & cond_emprestimo

    df.loc[cond_pag_emp, "Categoria_final"] = proc_original[cond_pag_emp] + " " + pessoa[cond_pag_emp]
    df.loc[cond_rec_emp, "Categoria_final"] = proc_original[cond_rec_emp] + " " + pessoa[cond_rec_emp]
    df.loc[cond_emprestimo & ~cond_pag_emp & ~cond_rec_emp, "Categoria_final"] = proc_original[cond_emprestimo]

    # --- Definição: É Despesa (Saída)? ---
    detalhe_lower = df[col_cat].astype(str).str.lower()

    cond_palavra_despesa = (
        fluxo_vazio & 
        ~(cond_rec_emp) & 
        (
            detalhe_lower.str.contains("custo", na=False) | 
            detalhe_lower.str.contains("despesa", na=False)
        )
    )
    cond_imobilizado = fluxo.str.contains("imobilizado", na=False)

    df["is_despesa"] = (
        cond_fluxo_despesa | 
        cond_pag_emp | 
        cond_palavra_despesa | 
        cond_imobilizado
    )
    df.loc[cond_fluxo_receita | cond_rec_emp, "is_despesa"] = False

    # Regra de fallback
    cond_sem_def = df["is_despesa"].isna() | (
        (df["is_despesa"] == False) & 
        (~cond_fluxo_receita) & 
        (~cond_rec_emp) & 
        (~cond_imobilizado) & 
        (~cond_palavra_despesa)
    )
    cond_pag_proc = proc.str.contains("pagamento", na=False)
    cond_rec_proc = proc.str.contains("recebimento", na=False)

    df.loc[cond_sem_def & cond_pag_proc, "is_despesa"] = True
    df.loc[cond_sem_def & cond_rec_proc, "is_despesa"] = False

    # === FILTRO DE SAÍDA ===
    # O usuário pediu para jogar todas as SAÍDAS para o layout do Omie (Contas a Pagar)
    df_saidas = df[df["is_despesa"] == True].copy()

    if df_saidas.empty:
        return pd.DataFrame() # Retorna vazio se não tiver saídas

    # --- Formatação de Valores e Datas ---
    df_saidas["Valor_Formatado"] = df_saidas["Valor total"].apply(converter_valor_omie)
    data_tes = formatar_data_coluna(df_saidas["Data da Tesouraria"])

    # --- MONTAGEM DO LAYOUT OMIE ---
    colunas_omie = [
        "Código de Integração", "Fornecedor * (Razão Social, Nome Fantasia, CNPJ ou CPF)", 
        "Categoria *", "Conta Corrente *", "Valor da Conta *", "Vendedor", "Projeto", 
        "Data de Emissão", "Data de Registro *", "Data de Vencimento *", "Data de Previsão", 
        "Data do Pagamento", "Valor do Pagamento", "Juros", "Multa", "Desconto", 
        "Data de Conciliação", "Observações", "Tipo de Documento", "Número do Documento", 
        "Parcela", "Total de Parcelas", "Número do Pedido", "Nota Fiscal", "Chave da NF-e", 
        "Forma de Pagamento", "Código de Barras do Boleto", "% de Juros ao Mês do Boleto", 
        "% de Multa por Atraso do Boleto", "Banco da Transferência", "Agência da Transferência", 
        "Conta Corrente da Transferência", "CNPJ ou CPF do Titular", "Nome do Titular da Conta", 
        "Finalidade da Transferência", "Chave Pix", "Valor PIS", "Reter PIS", "Valor COFINS", 
        "Reter COFINS", "Valor CSLL", "Reter CSLL", "Valor IR", "Reter IR", "Valor ISS", 
        "Reter ISS", "Valor INSS", "Reter INSS", "Departamento (100%)", "Número da NF (serviço tomado)", 
        "Série", "Código do Serviço (LC116)", "Valor total da NF", "CST do PIS", 
        "Base de Cálculo - PIS", "Alíquota do PIS (%)", "Valor do PIS", "CST do COFINS", 
        "Base de cálculo - COFINS", "Alíquota  do COFINS (%)", "Valor do COFINS"
    ]

    out = pd.DataFrame(columns=colunas_omie)

    # Preenchimento das colunas mapeadas
    # Usando .get() para evitar erro se a coluna original não existir, mas o user disse que existe "Id Item tesouraria"
    out["Código de Integração"] = df_saidas.get("Id Item tesouraria", "")
    out["Fornecedor * (Razão Social, Nome Fantasia, CNPJ ou CPF)"] = df_saidas.get("Pessoa", "")
    out["Categoria *"] = df_saidas["Categoria_final"]
    out["Valor da Conta *"] = df_saidas["Valor_Formatado"]
    out["Valor do Pagamento"] = df_saidas["Valor_Formatado"] # Assume que foi pago integralmente
    
    # Datas
    out["Data de Emissão"] = data_tes
    out["Data de Registro *"] = data_tes
    out["Data de Vencimento *"] = data_tes
    out["Data do Pagamento"] = data_tes # Se já está na tesouraria W4, consideramos pago
    
    # Outros
    out["Observações"] = df_saidas.get("Descrição", "")
    
    # Coluna Conta Corrente é obrigatória no Omie, mas não temos o De/Para. 
    # Deixaremos vazia ou fixa se você quiser depois.
    out["Conta Corrente *"] = "" 

    return out

# ============================
# CARREGAR ARQUIVO W4
# ============================

def carregar_arquivo_w4(arq):
    if arq.name.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(arq)
    else:
        return pd.read_csv(arq, sep=";", encoding="latin1")

# ============================
# CARREGAR CATEGORIAS
# ============================
try:
    df_cat_raw = pd.read_excel("categorias_contabeis.xlsx")
    df_cat_prep = preparar_categorias(df_cat_raw)
except FileNotFoundError:
    # Cria um DF vazio dummy para o app não quebrar se não tiver o arquivo local no teste
    st.warning("Arquivo 'categorias_contabeis.xlsx' não encontrado. O De/Para de categorias não funcionará.")
    df_cat_prep = pd.DataFrame(columns=["Descrição da categoria financeira", "nome_base"])

# ============================
# INTERFACE
# ============================

st.title("Conversor W4 -> Omie (Contas a Pagar)")
st.markdown("### Envie o arquivo W4")

arq_w4 = st.file_uploader("Selecione o arquivo W4", type=["csv", "xlsx", "xls"])

if arq_w4:
    if st.button("Gerar Planilha Omie"):
        try:
            df_w4 = carregar_arquivo_w4(arq_w4)
            df_final = converter_w4_para_omie(df_w4, df_cat_prep)

            if df_final.empty:
                st.warning("Nenhuma saída/despesa identificada no arquivo enviado.")
            else:
                st.success(f"Conversão realizada! {len(df_final)} lançamentos de saída identificados.")

                buffer = BytesIO()
                df_final.to_excel(buffer, index=False, engine="openpyxl")
                buffer.seek(0)

                st.download_button(
                    label="Baixar Importação Omie",
                    data=buffer,
                    file_name="importacao_omie_pagar.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Erro no processamento: {e}")
            st.exception(e) # Mostra o erro detalhado para debug
else:
    st.info("Aguardando upload...")
