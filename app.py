import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import json

# ─── CONFIG ───────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Controle Financeiro",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# ─── ESTILO VISUAL ────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&display=swap');

* { font-family: 'DM Sans', sans-serif; }

.stApp {
    background: linear-gradient(145deg, #0f172a, #1e293b);
    color: #e2e8f0;
}

/* Cards de resumo */
.summary-card {
    background: #1e293b;
    border-radius: 14px;
    padding: 20px 24px;
    border: 1px solid rgba(255,255,255,0.07);
}
.summary-label {
    color: #94a3b8;
    font-size: 13px;
    font-weight: 500;
    margin-bottom: 6px;
}
.summary-value {
    font-size: 26px;
    font-weight: 700;
    margin: 0;
}
.value-white  { color: #f1f5f9; }
.value-amber  { color: #fbbf24; }
.value-green  { color: #34d399; }
.value-red    { color: #f87171; }

/* Formulário */
.form-section {
    background: rgba(255,255,255,0.04);
    border: 1px solid rgba(255,255,255,0.09);
    border-radius: 14px;
    padding: 28px;
    margin-bottom: 24px;
}

/* Histórico item */
.tx-item {
    background: #1e293b;
    border-radius: 10px;
    padding: 14px 18px;
    margin-bottom: 8px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    border: 1px solid rgba(255,255,255,0.06);
}
.tx-desc { font-weight: 600; font-size: 15px; }
.tx-meta { color: #64748b; font-size: 12px; margin-top: 3px; }
.tx-amount { font-weight: 700; font-size: 16px; }
.tx-entrada { color: #34d399; }
.tx-saida   { color: #f87171; }
.tx-neutro  { color: #94a3b8; }

/* Inputs e selects */
div[data-baseweb="select"] > div,
div[data-baseweb="input"] > div {
    background-color: #1e293b !important;
    border-color: #334155 !important;
    color: #f1f5f9 !important;
}

/* Botão principal */
.stButton > button {
    background: #6366f1 !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    font-size: 15px !important;
    padding: 10px 20px !important;
    width: 100%;
}
.stButton > button:hover {
    background: #4f46e5 !important;
}

/* Título principal */
h1 { color: #f1f5f9 !important; font-size: 32px !important; font-weight: 700 !important; }
h2 { color: #e2e8f0 !important; font-size: 22px !important; font-weight: 600 !important; }
h3 { color: #cbd5e1 !important; }

/* Esconde elementos do Streamlit */
#MainMenu, footer, header { visibility: hidden; }
.block-container { padding-top: 2rem !important; }

/* Filtros */
.filter-box {
    background: rgba(255,255,255,0.04);
    border: 1px solid rgba(255,255,255,0.09);
    border-radius: 12px;
    padding: 20px;
    margin-bottom: 16px;
}

/* Tag de cartão */
.tag {
    display: inline-block;
    background: rgba(99,102,241,0.2);
    color: #a5b4fc;
    font-size: 11px;
    font-weight: 600;
    padding: 2px 8px;
    border-radius: 20px;
    margin-left: 6px;
}
.tag-person {
    background: rgba(251,191,36,0.15);
    color: #fbbf24;
}
</style>
""", unsafe_allow_html=True)

# ─── GOOGLE SHEETS ────────────────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

@st.cache_resource
def get_sheet():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    client = gspread.authorize(creds)
    sheet = client.open("financas-app").worksheet("transacoes")
    return sheet

def load_data(sheet):
    rows = sheet.get_all_records()
    if not rows:
        return pd.DataFrame(columns=[
            "type","payment_method","amount","installments",
            "installment_value","description","created_at",
            "card","is_for_someone","bought_by"
        ])
    df = pd.DataFrame(rows)
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
    df["installment_value"] = pd.to_numeric(df["installment_value"], errors="coerce").fillna(0)
    df["installments"] = pd.to_numeric(df["installments"], errors="coerce").fillna(1)
    df["created_at"] = pd.to_datetime(df["created_at"], errors="coerce")
    df["is_for_someone"] = df["is_for_someone"].astype(str).str.upper() == "TRUE"
    return df

def save_transaction(sheet, data: dict):
    row = [
        data["type"],
        data["payment_method"],
        data["amount"],
        data["installments"],
        data["installment_value"],
        data["description"],
        data["created_at"],
        data.get("card", ""),
        str(data.get("is_for_someone", False)).upper(),
        data.get("bought_by", "")
    ]
    sheet.append_row(row)

def delete_transaction(sheet, row_index: int):
    sheet.delete_rows(row_index + 2)  # +2: header + 0-index

# ─── LABELS ───────────────────────────────────────────────────────────────────
METHOD_LABELS = {
    "pix_conta":        "Pix na conta",
    "pix":              "Pix",
    "dinheiro_vivo":    "Dinheiro vivo",
    "saque_dinheiro":   "Saque dinheiro",
    "credito_parcelado":"Crédito parcelado",
}
CARD_LABELS = {
    "inter":        "Inter",
    "mercado_pago": "Mercado Pago",
    "nubank":       "Nubank",
    "nu_pj":        "Nu PJ",
    "picpay":       "PicPay",
    "amazon":       "Amazon",
    "mei_pj":       "Mei PJ",
}

# ─── CÁLCULOS ─────────────────────────────────────────────────────────────────
def calc_summary(df):
    bank = cash = income = expense = 0.0
    for _, r in df.iterrows():
        amt = r["installment_value"] if r["payment_method"] == "credito_parcelado" else r["amount"]
        if r["type"] == "entrada":
            income += amt
            if r["payment_method"] == "pix_conta":
                bank += amt
            elif r["payment_method"] == "dinheiro_vivo":
                cash += amt
        else:
            if r["payment_method"] == "saque_dinheiro":
                bank -= amt
                cash += amt
            else:
                expense += amt
                if r["payment_method"] == "dinheiro_vivo":
                    cash -= amt
                else:
                    bank -= amt
    return bank, cash, income, expense

def fmt(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def top5_expenses(df):
    cutoff_str = (datetime.now() - timedelta(days=30)).isoformat()
    dfc = df.copy()
    dfc["created_at_str"] = dfc["created_at"].astype(str)
    mask = (
        (dfc["type"] == "saida") &
        (dfc["payment_method"] != "saque_dinheiro") &
        (dfc["created_at_str"] >= cutoff_str)
    )
    recent = dfc[mask].copy()
    recent["val"] = recent.apply(
        lambda r: r["installment_value"] if r["payment_method"] == "credito_parcelado" else r["amount"],
        axis=1
    )
    top = recent.groupby("description")["val"].sum().nlargest(5).reset_index()
    return top

# ─── MAIN ─────────────────────────────────────────────────────────────────────
def main():
    try:
        sheet = get_sheet()
    except Exception as e:
        st.error(f"Erro ao conectar com o Google Sheets: {e}")
        st.stop()

    df = load_data(sheet)

    # ── TÍTULO
    st.markdown("<h1>💰 Controle Financeiro</h1>", unsafe_allow_html=True)
    st.markdown("<div style='margin-bottom:24px'></div>", unsafe_allow_html=True)

    # ── CARDS DE RESUMO
    bank, cash, income, expense = calc_summary(df)
    top5 = top5_expenses(df)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""
        <div class="summary-card">
            <div class="summary-label">Saldo em banco</div>
            <p class="summary-value value-white">{fmt(bank)}</p>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="summary-card">
            <div class="summary-label">Dinheiro vivo</div>
            <p class="summary-value value-amber">{fmt(cash)}</p>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class="summary-card">
            <div class="summary-label">Entradas</div>
            <p class="summary-value value-green">{fmt(income)}</p>
        </div>""", unsafe_allow_html=True)
    with c4:
        col_val, col_chart = st.columns([1, 1])
        with col_val:
            st.markdown(f"""
            <div class="summary-card" style="height:100%">
                <div class="summary-label">Saídas</div>
                <p class="summary-value value-red">{fmt(expense)}</p>
            </div>""", unsafe_allow_html=True)
        with col_chart:
            if not top5.empty:
                fig = go.Figure(go.Bar(
                    x=top5["description"],
                    y=top5["val"],
                    marker_color="#ef4444",
                    hovertemplate="<b>%{x}</b><br>%{y:,.2f}<extra></extra>"
                ))
                fig.update_layout(
                    paper_bgcolor="rgba(0,0,0,0)",
                    plot_bgcolor="rgba(0,0,0,0)",
                    margin=dict(l=0, r=0, t=0, b=0),
                    height=80,
                    showlegend=False,
                    xaxis=dict(visible=False),
                    yaxis=dict(visible=False),
                )
                st.plotly_chart(fig, use_container_width=True, config={"displayModeBar": False})

    st.markdown("<div style='margin:28px 0 8px'></div>", unsafe_allow_html=True)

    # ── FORMULÁRIO
    st.markdown("<h2>Nova Transação</h2>", unsafe_allow_html=True)
    with st.container():
        with st.form("tx_form", clear_on_submit=True):
            col1, col2 = st.columns(2)
            with col1:
                tx_type = st.selectbox("Tipo", ["entrada", "saida"],
                    format_func=lambda x: "Entrada" if x == "entrada" else "Saída")
            with col2:
                if tx_type == "entrada":
                    method_opts = ["pix_conta", "dinheiro_vivo"]
                else:
                    method_opts = ["pix", "dinheiro_vivo", "saque_dinheiro", "credito_parcelado"]
                tx_method = st.selectbox("Método", method_opts,
                    format_func=lambda x: METHOD_LABELS.get(x, x))

            col3, col4 = st.columns(2)
            with col3:
                default_desc = "Saque dinheiro" if tx_method == "saque_dinheiro" else ""
                tx_desc = st.text_input("Descrição", value=default_desc, placeholder="Ex: Salário")
            with col4:
                tx_amount = st.number_input("Valor (R$)", min_value=0.01, step=0.01, format="%.2f")

            # Campos extras para crédito parcelado
            tx_installments = 1
            tx_card = None
            tx_is_for_someone = False
            tx_bought_by = ""

            if tx_method == "credito_parcelado":
                col5, col6 = st.columns(2)
                with col5:
                    tx_installments = st.number_input("Parcelas", min_value=2, max_value=48, value=2, step=1)
                with col6:
                    installment_val = tx_amount / tx_installments if tx_installments > 0 else 0
                    st.markdown(f"""
                    <div style="margin-top:28px;background:#1e293b;border:1px solid #334155;
                    border-radius:8px;padding:9px 14px;color:#f1f5f9;font-size:15px;">
                        {fmt(installment_val)} / parcela
                    </div>""", unsafe_allow_html=True)

                tx_card = st.selectbox("Cartão", list(CARD_LABELS.keys()),
                    format_func=lambda x: CARD_LABELS.get(x, x))

                tx_is_for_someone = st.checkbox("🛍️ Compra de alguém")
                if tx_is_for_someone:
                    tx_bought_by = st.text_input("Quem comprou?", placeholder="Ex: João")

            submitted = st.form_submit_button("➕ Adicionar transação")

            if submitted:
                if not tx_desc.strip():
                    st.error("Preencha a descrição.")
                else:
                    installment_value = (tx_amount / tx_installments
                                         if tx_method == "credito_parcelado" else tx_amount)
                    save_transaction(sheet, {
                        "type": tx_type,
                        "payment_method": tx_method,
                        "amount": tx_amount,
                        "installments": tx_installments,
                        "installment_value": installment_value,
                        "description": tx_desc.strip(),
                        "created_at": datetime.now().isoformat(),
                        "card": tx_card or "",
                        "is_for_someone": tx_is_for_someone,
                        "bought_by": tx_bought_by.strip() if tx_is_for_someone else "",
                    })
                    st.success("Transação adicionada! ✅")
                    st.cache_resource.clear()
                    st.rerun()

    st.markdown("<div style='margin:8px 0'></div>", unsafe_allow_html=True)

    # ── HISTÓRICO
    st.markdown("<h2>Histórico</h2>", unsafe_allow_html=True)

    # Filtros
    with st.expander("🔍 Filtros", expanded=False):
        fc1, fc2, fc3, fc4 = st.columns(4)
        with fc1:
            f_type = st.selectbox("Tipo", ["Todos", "Entradas", "Saídas"])
        with fc2:
            # Cartões usados nos últimos 30 dias
            cutoff30_str = (datetime.now() - timedelta(days=30)).isoformat()
            used_cards = df[
                (df["payment_method"] == "credito_parcelado") &
                (df["created_at"].astype(str) >= cutoff30_str) &
                (df["card"].astype(str).str.strip() != "")
            ]["card"].unique().tolist()
            card_opts = ["Todos"] + [CARD_LABELS.get(c, c) for c in used_cards]
            f_card = st.selectbox("Cartão (últimos 30 dias)", card_opts)
        with fc3:
            buyers = df[df["is_for_someone"] & (df["bought_by"].astype(str).str.strip() != "")]["bought_by"].unique().tolist()
            buyer_opts = ["Todos"] + sorted(buyers)
            f_buyer = st.selectbox("Compra de alguém", buyer_opts)
        with fc4:
            st.markdown("<div style='margin-top:28px'></div>", unsafe_allow_html=True)
            reset = st.button("Limpar filtros")

    if reset:
        f_type = "Todos"
        f_card = "Todos"
        f_buyer = "Todos"

    # Aplicar filtros
    filtered = df.copy()
    if f_type == "Entradas":
        filtered = filtered[filtered["type"] == "entrada"]
    elif f_type == "Saídas":
        filtered = filtered[filtered["type"] == "saida"]
    if f_card != "Todos":
        card_key = next((k for k, v in CARD_LABELS.items() if v == f_card), None)
        if card_key:
            filtered = filtered[filtered["card"] == card_key]
    if f_buyer != "Todos":
        filtered = filtered[filtered["bought_by"] == f_buyer]

    # Ordenar mais recente primeiro
    filtered = filtered.sort_values("created_at", ascending=False)

    if filtered.empty:
        st.markdown("""
        <div style="text-align:center;color:#475569;padding:40px 0;font-size:15px;">
            Nenhuma transação encontrada.
        </div>""", unsafe_allow_html=True)
    else:
        for idx, (df_idx, row) in enumerate(filtered.iterrows()):
            amt = row["installment_value"] if row["payment_method"] == "credito_parcelado" else row["amount"]
            is_saque = row["payment_method"] == "saque_dinheiro"

            if row["type"] == "entrada":
                prefix = "+"
                amount_class = "tx-entrada"
            elif is_saque:
                prefix = ""
                amount_class = "tx-neutro"
            else:
                prefix = "−"
                amount_class = "tx-saida"

            # Meta info
            method_label = METHOD_LABELS.get(row["payment_method"], row["payment_method"])
            if row["payment_method"] == "credito_parcelado":
                method_label = f"Crédito {int(row['installments'])}x"
            dt = row["created_at"]
            dt_str = dt.strftime("%d/%m/%Y às %H:%M") if pd.notnull(dt) else ""

            card_tag = ""
            if row["payment_method"] == "credito_parcelado" and row.get("card"):
                card_tag = f'<span class="tag">{CARD_LABELS.get(row["card"], row["card"])}</span>'

            person_tag = ""
            if row.get("is_for_someone") and row.get("bought_by"):
                person_tag = f'<span class="tag tag-person">🛍️ {row["bought_by"]}</span>'

            col_item, col_del = st.columns([20, 1])
            with col_item:
                st.markdown(f"""
                <div class="tx-item">
                    <div>
                        <div class="tx-desc" style="color:{'#f87171' if amount_class=='tx-saida' else '#f1f5f9'}">
                            {row['description']}{card_tag}{person_tag}
                        </div>
                        <div class="tx-meta">{method_label} • {dt_str}</div>
                    </div>
                    <div class="tx-amount {amount_class}">{prefix}{fmt(amt)}</div>
                </div>""", unsafe_allow_html=True)
            with col_del:
                st.markdown("<div style='margin-top:14px'></div>", unsafe_allow_html=True)
                if st.button("🗑️", key=f"del_{df_idx}_{idx}", help="Excluir"):
                    delete_transaction(sheet, df_idx)
                    st.cache_resource.clear()
                    st.rerun()

if __name__ == "__main__":
    main()
