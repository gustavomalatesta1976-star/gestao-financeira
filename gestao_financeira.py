import streamlit as st
import pandas as pd
import plotly.express as px
import os
import re
from datetime import date
import io

st.set_page_config(
    page_title="Livro despesa residencial",
    page_icon="📒",
    layout="wide"
)

ARQUIVO_LIVRO = r"C:\Users\Windows\Desktop\livro_caixa_2026.xlsx"

CATEGORIAS = [
    "Moradia / Apartamento",
    "Contas Fixas (Luz, Água, Internet, Celular)",
    "Cartão de Crédito / Parcelas",
    "Veículo (IPVA, Seguro, Combustível)",
    "Pets (Charlotte)",
    "Alimentação / Mercado",
    "Lazer / Viagens",
    "Salário / Receitas",
    "Outros"
]

def item(mes, dt, tipo, cat, desc, val, status):
    return {
        "Mês/Ano": mes,
        "Data": dt,
        "Tipo": tipo,
        "Categoria": cat,
        "Descrição": desc,
        "Valor": float(val),
        "Status": status
    }

def limpar_vazios(df):
    if df.empty:
        return df
    df = df.dropna(subset=["Descrição", "Valor"]).copy()
    df = df[df["Descrição"].astype(str).str.strip() != ""]
    df = df[df["Descrição"].astype(str).str.lower() != "none"]
    return df.reset_index(drop=True)

def carregar_livro():
    if os.path.exists(ARQUIVO_LIVRO):
        df_lido = pd.read_excel(ARQUIVO_LIVRO)
        return limpar_vazios(df_lido)
    
    M = "Moradia / Apartamento"
    F = "Contas Fixas (Luz, Água, Internet, Celular)"
    C = "Cartão de Crédito / Parcelas"
    V = "Veículo (IPVA, Seguro, Combustível)"
    P = "Pets (Charlotte)"
    L = "Lazer / Viagens"
    R = "Salário / Receitas"
    O = "Outros"

    dados = [
        # --- AGOSTO 2026 ---
        item("2026-08", "2026-08-04", "Receita", R, "Aporte Gustavo Agosto", 5556.00, "Recebido"),
        item("2026-08", "2026-08-04", "Receita", R, "Receita Jéssica Agosto", 1268.00, "Recebido"),
        item("2026-08", "2026-08-07", "Despesa", M, "Ap - Caixa", 1290.00, "Pago"),
        item("2026-08", "2026-08-07", "Despesa", M, "Ap (B/G) (12/71)", 1500.00, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", M, "Ap - Condomínio", 100.00, "Pago"),
        item("2026-08", "2026-08-07", "Despesa", M, "Ap - IPTU", 94.00, "Pago"),
        item("2026-08", "2026-08-20", "Despesa", F, "Ap - Internet", 125.00, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", F, "Ap - Água + Condomínio", 65.50, "Pago"),
        item("2026-08", "2026-08-07", "Despesa", F, "Ap - Luz", 266.65, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", F, "Gustavo - Celular", 45.00, "Pago"),
        item("2026-08", "2026-08-20", "Despesa", C, "Gustavo - Bruno (7/10)", 250.00, "Pago"),
        item("2026-08", "2026-08-07", "Despesa", V, "Gustavo - Seguro", 332.00, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", C, "Gustavo - Matheus", 16.36, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", C, "Gustavo - Cartão Tânia óculos (6/8)", 92.50, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", C, "Gustavo - Cartão Tânia roupa Itui (2/2)", 74.80, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", C, "Gustavo - Cartão Hélio (3/10)", 300.00, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", C, "Gustavo - Cartão Thamiris mat. ap (3/4)", 96.45, "Pago"),
        item("2026-08", "2026-08-20", "Despesa", V, "Gustavo - Cartão Mãe carro (1/12)", 250.00, "Pago"),
        item("2026-08", "2026-08-11", "Despesa", O, "Gustavo - Mãe", 300.00, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", C, "Gustavo - Tênis avulso", 139.00, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", C, "Gustavo - Tênis (1/4)", 160.00, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", F, "Jéssica - Celular", 110.00, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", F, "Jéssica - INSS", 88.00, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", C, "Jéssica - Cartão Thamiris forno (9/10)", 138.00, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", C, "Jéssica - Cartão Tânia laser (9/12)", 167.00, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", C, "Jéssica - Cartão Tânia mesa (5/10)", 187.00, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", P, "Jéssica - Cartão Tânia vacina (2/10)", 295.00, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", C, "Jéssica - Gabi (3/3)", 137.00, "Pago"),
        item("2026-08", "2026-08-12", "Despesa", O, "Jéssica - Leninha", 146.00, "Pago"),

        # --- SETEMBRO 2026 ---
        item("2026-09", "2026-09-26", "Receita", R, "Recebimento Jéssica Setembro", 1200.00, "Recebido"),
        item("2026-09", "2026-09-08", "Despesa", M, "Ap - Caixa", 1290.00, "Pago"),
        item("2026-09", "2026-09-08", "Despesa", M, "Ap (B/G) (13/71)", 1500.00, "Pago"),
        item("2026-09", "2026-09-21", "Despesa", M, "Ap - Condomínio", 100.00, "Pago"),
        item("2026-09", "2026-09-08", "Despesa", M, "Ap - IPTU", 94.00, "Pago"),
        item("2026-09", "2026-09-08", "Despesa", F, "Ap - Internet", 138.00, "Pago"),
        item("2026-09", "2026-09-08", "Despesa", F, "Ap - Água + Condomínio", 59.59, "Pago"),
        item("2026-09", "2026-09-08", "Despesa", F, "Ap - Luz", 205.53, "Pago"),
        item("2026-09", "2026-09-09", "Despesa", F, "Gustavo - Celular", 45.00, "Pago"),
        item("2026-09", "2026-09-14", "Despesa", C, "Gustavo - Bruno (8/10)", 250.00, "Pago"),
        item("2026-09", "2026-09-11", "Despesa", V, "Gustavo - Seguro", 332.00, "Pago"),
        item("2026-09", "2026-09-11", "Despesa", C, "Gustavo - Matheus", 16.36, "Pago"),
        item("2026-09", "2026-09-10", "Despesa", C, "Gustavo - Cartão Tânia óculos (7/8)", 92.50, "Pago"),
        item("2026-09", "2026-09-10", "Despesa", C, "Gustavo - Cartão Hélio (4/10)", 300.00, "Pago"),
        item("2026-09", "2026-09-10", "Despesa", C, "Gustavo - Cartão Thamiris mat. ap (4/4)", 96.45, "Pago"),
        item("2026-09", "2026-09-15", "Despesa", V, "Gustavo - Cartão Mãe carro (2/12)", 250.00, "Pago"),
        item("2026-09", "2026-09-15", "Despesa", V, "Gustavo - Cartão Mãe carro (1/4)", 116.00, "Pago"),
        item("2026-09", "2026-09-12", "Despesa", C, "Gustavo - Tênis (2/4)", 160.00, "Pago"),
        item("2026-09", "2026-09-21", "Despesa", F, "Jéssica - Celular (acumulado p/ Out)", 110.00, "Pendente"),
        item("2026-09", "2026-09-21", "Despesa", F, "Jéssica - INSS (acumulado p/ Out)", 88.00, "Pendente"),
        item("2026-09", "2026-09-10", "Despesa", C, "Jéssica - Cartão Thamiris forno (10/10)", 138.00, "Pago"),
        item("2026-09", "2026-09-10", "Despesa", C, "Jéssica - Cartão Thamiris (1/3)", 100.00, "Pago"),
        item("2026-09", "2026-09-10", "Despesa", C, "Jéssica - Cartão Tânia laser (10/12)", 167.00, "Pago"),
        item("2026-09", "2026-09-10", "Despesa", C, "Jéssica - Cartão Tânia mesa (6/10)", 187.00, "Pago"),
        item("2026-09", "2026-09-10", "Despesa", P, "Jéssica - Cartão Tânia vacina (3/10)", 295.00, "Pago"),
        item("2026-09", "2026-09-10", "Despesa", P, "Jéssica - Banho Charlotte", 100.00, "Pago"),
        item("2026-09", "2026-09-21", "Despesa", O, "Jéssica - Leninha (shampoo/cond.)", 80.00, "Pago"),
        item("2026-09", "2026-09-15", "Despesa", C, "Jéssica - Tênis (1/4)", 67.47, "Pago"),

        # --- OUTUBRO 2026 ---
        item("2026-10", "2026-10-14", "Receita", R, "Recebimento Jéssica (Leninha)", 1207.50, "Recebido"),
        item("2026-10", "2026-10-08", "Despesa", M, "Ap - Caixa", 1321.82, "Pago"),
        item("2026-10", "2026-10-08", "Despesa", M, "Ap (B/G) (14/71)", 1500.00, "Pago"),
        item("2026-10", "2026-10-08", "Despesa", M, "Ap - Condomínio", 100.00, "Pago"),
        item("2026-10", "2026-10-09", "Despesa", M, "Ap - IPTU", 94.00, "Pago"),
        item("2026-10", "2026-10-08", "Despesa", F, "Ap - Internet", 138.99, "Pago"),
        item("2026-10", "2026-10-08", "Despesa", F, "Ap - Luz", 194.91, "Pago"),
        item("2026-10", "2026-10-15", "Despesa", F, "Ap - Água + Condomínio", 46.64, "Pago"),
        item("2026-10", "2026-10-09", "Despesa", F, "Gustavo - Celular", 45.00, "Pago"),
        item("2026-10", "2026-10-14", "Despesa", C, "Gustavo - Bruno (9/10)", 250.00, "Pago"),
        item("2026-10", "2026-10-15", "Despesa", V, "Gustavo - Seguro", 332.00, "Pago"),
        item("2026-10", "2026-10-14", "Despesa", C, "Gustavo - Matheus", 16.36, "Pago"),
        item("2026-10", "2026-10-14", "Despesa", C, "Gustavo - Cartão Tânia óculos (8/8)", 92.50, "Pago"),
        item("2026-10", "2026-10-14", "Despesa", C, "Gustavo - Cartão Hélio (5/10)", 300.00, "Pago"),
        item("2026-10", "2026-10-14", "Despesa", V, "Gustavo - Cartão Mãe carro (3/12)", 250.00, "Pago"),
        item("2026-10", "2026-10-14", "Despesa", V, "Gustavo - Cartão Mãe carro (2/4)", 116.00, "Pago"),
        item("2026-10", "2026-10-14", "Despesa", C, "Gustavo - Tênis (3/4)", 160.00, "Pago"),
        item("2026-10", "2026-10-29", "Despesa", V, "Gustavo - IPVA (2/3)", 599.22, "Pago"),
        item("2026-10", "2026-10-14", "Despesa", F, "Jéssica - Celular (Set + Out)", 220.00, "Pago"),
        item("2026-10", "2026-10-14", "Despesa", F, "Jéssica - INSS (Set + Out)", 176.00, "Pago"),
        item("2026-10", "2026-10-14", "Despesa", C, "Jéssica - Cartão Thamiris (2/3)", 100.00, "Pago"),
        item("2026-10", "2026-10-14", "Despesa", C, "Jéssica - Cartão Tânia laser (11/12)", 167.00, "Pago"),
        item("2026-10", "2026-10-14", "Despesa", C, "Jéssica - Cartão Tânia mesa (7/10)", 187.00, "Pago"),
        item("2026-10", "2026-10-14", "Despesa", P, "Jéssica - Cartão Tânia vacina (4/10)", 295.00, "Pago"),
        item("2026-10", "2026-10-14", "Despesa", P, "Jéssica - Banho Charlotte", 100.00, "Pago"),
        item("2026-10", "2026-10-14", "Despesa", C, "Jéssica - Tênis (2/4)", 67.47, "Pago"),

        # --- NOVEMBRO 2026 ---
        item("2026-11", "2026-11-10", "Receita", R, "Depósito Gustavo p/ Contas", 3000.00, "Recebido"),
        item("2026-11", "2026-11-10", "Receita", R, "Recebimento Jéssica", 1000.00, "Recebido"),
        item("2026-11", "2026-11-11", "Receita", R, "Repasse Jéssica (Cartões Tânia + Hélio)", 949.00, "Recebido"),
        item("2026-11", "2026-11-10", "Despesa", M, "Ap - Caixa", 1321.82, "Pago"),
        item("2026-11", "2026-11-10", "Despesa", M, "Ap (B/G) (15/71)", 1500.00, "Pago"),
        item("2026-11", "2026-11-10", "Despesa", M, "Ap - Condomínio e Luz", 100.00, "Pago"),
        item("2026-11", "2026-11-13", "Despesa", M, "Ap - IPTU", 94.00, "Pago"),
        item("2026-11", "2026-11-13", "Despesa", F, "Ap - Internet", 138.99, "Pago"),
        item("2026-11", "2026-11-13", "Despesa", F, "Ap - Água + Condomínio", 54.37, "Pago"),
        item("2026-11", "2026-11-13", "Despesa", F, "Ap - Luz", 178.10, "Pago"),
        item("2026-11", "2026-11-13", "Despesa", F, "Gustavo - Celular", 45.00, "Pago"),
        item("2026-11", "2026-11-13", "Despesa", V, "Gustavo - Seguro", 332.00, "Pago"),
        item("2026-11", "2026-11-11", "Despesa", C, "Gustavo - Cartão Hélio (6/10)", 300.00, "Pago"),
        item("2026-11", "2026-11-13", "Despesa", C, "Gustavo - Bruno (10/10)", 250.00, "Pendente"),
        item("2026-11", "2026-11-13", "Despesa", C, "Gustavo - Matheus", 16.36, "Pendente"),
        item("2026-11", "2026-11-13", "Despesa", V, "Gustavo - Cartão Mãe carro (4/12)", 250.00, "Pendente"),
        item("2026-11", "2026-11-13", "Despesa", V, "Gustavo - Cartão Mãe carro (3/4)", 116.00, "Pendente"),
        item("2026-11", "2026-11-13", "Despesa", V, "Gustavo - Cartão mãe carro (1/3)", 107.00, "Pendente"),
        item("2026-11", "2026-11-13", "Despesa", C, "Gustavo - Cartão Mãe aniversário", 214.00, "Pendente"),
        item("2026-11", "2026-11-13", "Despesa", C, "Gustavo - Tênis (4/4)", 160.00, "Pendente"),
        item("2026-11", "2026-11-27", "Despesa", V, "Gustavo - IPVA (3/3)", 599.22, "Pendente"),
        item("2026-11", "2026-11-13", "Despesa", L, "Gustavo - Praia (1/5)", 185.00, "Pendente"),
        item("2026-11", "2026-11-11", "Despesa", C, "Jéssica - Cartão Tânia laser (12/12)", 167.00, "Pago"),
        item("2026-11", "2026-11-11", "Despesa", C, "Jéssica - Cartão Tânia mesa (8/10)", 187.00, "Pago"),
        item("2026-11", "2026-11-11", "Despesa", P, "Jéssica - Cartão Tânia vacina (5/10)", 295.00, "Pago"),
        item("2026-11", "2026-11-13", "Despesa", P, "Jéssica - Banho Charlotte", 100.00, "Pago"),
        item("2026-11", "2026-11-13", "Despesa", F, "Jéssica - Celular", 75.00, "Pendente"),
        item("2026-11", "2026-11-13", "Despesa", F, "Jéssica - INSS", 88.00, "Pendente"),
        item("2026-11", "2026-11-13", "Despesa", C, "Jéssica - Cartão Thamiris (3/3)", 100.00, "Pendente"),
        item("2026-11", "2026-11-13", "Despesa", P, "Jéssica - Tapete higiênico Charlotte", 90.00, "Pendente"),
        item("2026-11", "2026-11-13", "Despesa", P, "Jéssica - Extras Charlotte", 60.00, "Pendente"),
        item("2026-11", "2026-11-13", "Despesa", C, "Jéssica - Tênis (3/4)", 67.47, "Pendente"),
    ]
    df_init = pd.DataFrame(dados)
    df_init.to_excel(ARQUIVO_LIVRO, index=False)
    return df_init

def salvar_livro(df):
    df_limpo = limpar_vazios(df)
    df_limpo.to_excel(ARQUIVO_LIVRO, index=False)

def brl(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

df = carregar_livro()

st.title("📒 Livro despesa residencial")

# ==========================================
# 1. LANÇAMENTO RÁPIDO & EXCLUSÃO DE ITEM
# ==========================================
col_add, col_del = st.columns([2, 1])

with col_add:
    with st.expander("➕ Novo Lançamento (Receita ou Despesa)", expanded=False):
        with st.form("form_caixa", clear_on_submit=True):
            c1, c2, c3 = st.columns([1.2, 1.2, 1.8])
            with c1:
                tipo = st.selectbox("Tipo", ["Despesa", "Receita"])
            with c2:
                data_mov = st.date_input("Data", value=date.today(), format="DD/MM/YYYY")
            with c3:
                categoria = st.selectbox("Categoria", CATEGORIAS)

            c4, c5, c6 = st.columns([2.5, 1.2, 1.2])
            with c4:
                descricao = st.text_input("Descrição", placeholder="Ex: Caixa Ap, Luz, Salário...")
            with c5:
                valor = st.number_input("Valor (R$)", min_value=0.01, step=10.0, format="%.2f")
            with c6:
                status = st.selectbox("Situação", ["Pago", "Pendente", "Recebido"])

            enviar = st.form_submit_button("💾 Registrar Lançamento", use_container_width=True, type="primary")

            if enviar:
                if not descricao.strip():
                    st.error("Digite a descrição do lançamento.")
                else:
                    nova_linha = pd.DataFrame([
                        item(data_mov.strftime("%Y-%m"), data_mov.strftime("%Y-%m-%d"), tipo, categoria, descricao.strip(), valor, status)
                    ])
                    df = pd.concat([df, nova_linha], ignore_index=True)
                    salvar_livro(df)
                    st.success("Lançamento registrado!")
                    st.rerun()

with col_del:
    with st.expander("🗑️ Apagar Lançamento", expanded=True):
        if not df.empty:
            opcoes_apagar = {
                idx: f"[{row['Data']}] {row['Tipo']} - {row['Descrição']} ({brl(row['Valor'])})"
                for idx, row in df.sort_values(by="Data", ascending=False).iterrows()
            }
            idx_escolhido = st.selectbox(
                "Selecione o item para remover:",
                options=list(opcoes_apagar.keys()),
                format_func=lambda x: opcoes_apagar[x]
            )
            if st.button("🗑️ Apagar Item Selecionado", use_container_width=True, type="secondary"):
                df = df.drop(index=idx_escolhido).reset_index(drop=True)
                salvar_livro(df)
                st.success("Lançamento apagado!")
                st.rerun()
        else:
            st.info("Nenhum lançamento cadastrado.")

# Filtro de Mês
if df.empty:
    st.info("Nenhum lançamento no momento. Adicione um item acima.")
    st.stop()

meses = sorted(df["Mês/Ano"].unique().tolist(), reverse=True)
col_mes, _ = st.columns([1, 3])
with col_mes:
    mes_atual = st.selectbox("📅 Mês de Referência:", meses, index=0)

df_mes = df[df["Mês/Ano"] == mes_atual].sort_values(by="Data").copy()

# ==========================================
# 2. DASHBOARD & PONTOS CRÍTICOS
# ==========================================
entradas = df_mes[df_mes["Tipo"] == "Receita"]["Valor"].sum()
saidas = df_mes[df_mes["Tipo"] == "Despesa"]["Valor"].sum()
saldo = entradas - saidas
pendentes = df_mes[(df_mes["Tipo"] == "Despesa") & (df_mes["Status"] == "Pendente")]["Valor"].sum()

k1, k2, k3, k4 = st.columns(4)
k1.metric("🟢 Entradas (Receitas)", brl(entradas))
k2.metric("🔴 Saídas (Despesas)", brl(saidas))
k3.metric("⚖️ Saldo do Caixa", brl(saldo))
k4.metric("⏳ Contas Pendentes", brl(pendentes))

st.markdown("---")

# Radar de Pontos Críticos
st.subheader("🚨 Pontos Críticos Sinalizados")
df_desp = df_mes[df_mes["Tipo"] == "Despesa"]

col_crit1, col_crit2 = st.columns(2)
with col_crit1:
    if entradas == 0:
        st.error("🚨 **Sem Receitas Lançadas:** Registre suas entradas do mês para calcular o saldo real.")
    elif saidas > entradas:
        st.error(f"🚨 **Caixa Negativo (Déficit de {brl(saidas - entradas)}):** As despesas superaram as receitas lançadas neste mês.")
    elif (saidas / entradas) >= 0.85:
        st.warning(f"⚠️ **Orçamento no Limite ({(saidas/entradas)*100:.0f}% comprometido):** Sobra apenas {brl(saldo)} livre no caixa.")
    else:
        st.success(f"✅ **Caixa Saudável:** Sobram {brl(saldo)} ({100 - (saidas/entradas)*100:.0f}% da receita livre).")

    if pendentes > 0:
        qtd_p = len(df_desp[df_desp["Status"] == "Pendente"])
        maior_p = df_desp[df_desp["Status"] == "Pendente"].sort_values(by="Valor", ascending=False).iloc[0]
        st.warning(f"⚠️ **Contas Pendentes:** **{qtd_p} conta(s)** em aberto somando **{brl(pendentes)}** (Maior: *{maior_p['Descrição']}* — {brl(maior_p['Valor'])}).")

with col_crit2:
    if not df_desp.empty and saidas > 0:
        por_cat = df_desp.groupby("Categoria")["Valor"].sum().sort_values(ascending=False)
        maior_cat = por_cat.index[0]
        maior_val = por_cat.iloc[0]
        pct_cat = (maior_val / saidas) * 100
        if pct_cat >= 35:
            st.warning(f"⚠️ **Maior Peso no Caixa:** `{maior_cat}` representa **{pct_cat:.1f}%** ({brl(maior_val)}) dos seus gastos.")
        else:
            st.info(f"ℹ️ **Maior Gasto:** `{maior_cat}` com **{brl(maior_val)}** ({pct_cat:.1f}%).")

        alivio = 0.0
        for _, r in df_desp.iterrows():
            m = re.search(r"\((\d+)\s*/\s*(\d+)\)", str(r["Descrição"]))
            if m and m.group(1) == m.group(2):
                alivio += r["Valor"]
        if alivio > 0:
            st.success(f"✅ **Alívio no Próximo Mês:** **{brl(alivio)}** em parcelas chegam ao fim neste mês.")

# ==========================================
# GRÁFICO ÚNICO EM FORMATO DE PIZZA
# ==========================================
st.subheader(f"🍕 Distribuição das Despesas por Categoria — {mes_atual}")
if not df_desp.empty:
    df_pizza = df_desp.groupby("Categoria", as_index=False)["Valor"].sum().sort_values(by="Valor", ascending=False)
    fig = px.pie(
        df_pizza,
        names="Categoria",
        values="Valor",
        hole=0.35
    )
    fig.update_traces(
        textposition="inside",
        textinfo="percent+label",
        hovertemplate="<b>%{label}</b><br>Valor: R$ %{value:,.2f}<br>Percentual: %{percent}<extra></extra>"
    )
    fig.update_layout(height=460, margin=dict(t=20, b=20, l=20, r=20))
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("Sem despesas lançadas para exibir no gráfico.")

st.markdown("---")

# ==========================================
# 3. LIVRO CAIXA (EXTRATO EDITÁVEL)
# ==========================================
st.subheader(f"📖 Extrato do Livro Caixa — {mes_atual}")

df_editado = st.data_editor(
    df_mes,
    num_rows="fixed",
    use_container_width=True,
    hide_index=True,
    column_config={
        "Valor": st.column_config.NumberColumn("Valor (R$)", format="R$ %.2f"),
        "Tipo": st.column_config.SelectboxColumn("Tipo", options=["Despesa", "Receita"]),
        "Status": st.column_config.SelectboxColumn("Situação", options=["Pago", "Pendente", "Recebido"]),
        "Categoria": st.column_config.SelectboxColumn("Categoria", options=CATEGORIAS)
    }
)

b1, b2 = st.columns(2)
with b1:
    if st.button("💾 Salvar Alterações no Livro Caixa", use_container_width=True, type="primary"):
        df_outros = df[df["Mês/Ano"] != mes_atual]
        df_final = pd.concat([df_outros, df_editado], ignore_index=True)
        salvar_livro(df_final)
        st.success("Livro Caixa atualizado!")
        st.rerun()

with b2:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Livro_Caixa")
    st.download_button(
        label="📥 Exportar Livro Caixa para Excel",
        data=buffer.getvalue(),
        file_name=f"livro_caixa_{mes_atual}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
