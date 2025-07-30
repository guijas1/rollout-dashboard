import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from datetime import datetime, timedelta

# =========================
# CONFIGURAÇÕES / ENTRADAS
# =========================
CAMINHO_PLANILHA = "Agendamentos_Rollout_2025_SO.xlsx"
ABA = "Planilha SO"
META_TOTAL = 346
DATA_INICIO_ROLLOUT = datetime(2025, 7, 7)
DATA_FIM_SUPORTE = datetime(2025, 10, 14)
HOJE = datetime.today()

# =========================
# CARREGAMENTO E PREPARO
# =========================
df = pd.read_excel(CAMINHO_PLANILHA, sheet_name=ABA)
df['Concluido'] = df['Concluido'].astype(str).str.strip().str.upper()
df['Data'] = pd.to_datetime(df['Data'], errors='coerce')
df = df.dropna(subset=['Data'])

# =========================
# CÁLCULOS
# =========================
total_concluido = df[df['Concluido'] == 'SIM'].shape[0]
agendados_atuais = df[df['Concluido'] == 'NÃO'].shape[0]
pendentes_atual = META_TOTAL - (total_concluido + agendados_atuais)
pendentes_atual = max(pendentes_atual, 0)

dias_decorridos = max((HOJE - DATA_INICIO_ROLLOUT).days, 1)
dias_uteis_restantes = np.busday_count(HOJE.date(), DATA_FIM_SUPORTE.date())

ritmo_atual = total_concluido / dias_decorridos
meta_diaria_uteis = pendentes_atual / dias_uteis_restantes if dias_uteis_restantes > 0 else 0
expectativa_ate_fim = total_concluido + ritmo_atual * dias_uteis_restantes

if ritmo_atual > 0:
    dias_uteis_necessarios_no_ritmo_atual = int(np.ceil(pendentes_atual / ritmo_atual))
    previsao_conclusao_no_ritmo_atual = np.busday_offset(HOJE.date(), dias_uteis_necessarios_no_ritmo_atual)
else:
    previsao_conclusao_no_ritmo_atual = None

ultimos_periodos = [7, 14, 21]
ritmos_recentes = {}
for d in ultimos_periodos:
    limite = HOJE - timedelta(days=d)
    concl_ult = df[(df['Concluido'] == 'SIM') & (df['Data'] >= limite)].shape[0]
    ritmos_recentes[d] = concl_ult / d

def simular_data_conclusao_por_produtividade(produtividade_dia_util: float):
    if produtividade_dia_util <= 0:
        return None
    dias = int(np.ceil(pendentes_atual / produtividade_dia_util))
    return np.busday_offset(HOJE.date(), dias)

def dias_uteis_extra_para_compensar(incremento_acima_do_atual: float):
    capacidade_prevista = total_concluido + ritmo_atual * dias_uteis_restantes
    gap = pendentes_atual - (capacidade_prevista - total_concluido)
    if gap <= 0:
        return 0
    if incremento_acima_do_atual <= 0:
        return None
    return int(np.ceil(gap / incremento_acima_do_atual))

# =========================
# DASHBOARD STREAMLIT
# =========================
st.set_page_config(page_title="Rollout Windows 11", layout="wide")
st.title("📊 Dashboard Executivo - Rollout Windows 11 (RJ)")
st.subheader("📆 Dados atualizados em tempo real com base na planilha")

col1, col2, col3 = st.columns(3)
col1.metric("✅ Concluídos", total_concluido)
col2.metric("📅 Agendados", agendados_atuais)
col3.metric("📦 Pendentes (ainda fora da planilha)", pendentes_atual)

st.caption("📌 Pendentes = 346 - (Concluídos + Agendados)")
st.caption("📅 Agendados = status 'NÃO' | ✅ Concluídos = status 'SIM'")

col4, col5, col6 = st.columns(3)
col4.metric("📆 Dias corridos desde 07/07/2025", dias_decorridos)
col5.metric("📆 Dias úteis restantes até 14/10", dias_uteis_restantes)
col6.metric("📈 Ritmo médio atual", f"{ritmo_atual:.2f}/dia corrido")

col7, col8 = st.columns(2)
col7.metric("⚖️ Meta diária ideal (dias úteis)", f"{meta_diaria_uteis:.2f}/dia útil",
            f"{(ritmo_atual - meta_diaria_uteis):+.2f} vs atual")
col8.metric("🔵 Expectativa no ritmo atual", int(expectativa_ate_fim))

st.markdown("---")

st.subheader("🔮 Previsão de conclusão")
if previsao_conclusao_no_ritmo_atual:
    st.info(f"📅 Conclusão prevista no ritmo atual: **{previsao_conclusao_no_ritmo_atual}**")
else:
    st.warning("Ritmo atual é zero — não foi possível prever a data de conclusão.")

st.markdown("### 🕒 Ritmo dos últimos dias")
cols = st.columns(len(ultimos_periodos))
for i, d in enumerate(ultimos_periodos):
    cols[i].metric(f"Últimos {d} dias", f"{ritmos_recentes[d]:.2f}/dia")

st.markdown("### 📈 Progresso geral")
categorias = ['Concluído', 'Agendado', 'Pendente']
valores = [total_concluido, agendados_atuais, pendentes_atual]
cores = ['green', 'orange', 'red']

fig1, ax1 = plt.subplots()
bars = ax1.bar(categorias, valores, color=cores)
for b in bars:
    ax1.text(b.get_x() + b.get_width()/2, b.get_height() + 3, f"{int(b.get_height())}", ha='center')
ax1.set_ylim(0, max(valores) + 50)
ax1.set_ylabel("Pessoas")
ax1.set_title("Distribuição Atual")
ax1.grid(axis='y', linestyle='--', alpha=0.5)
st.pyplot(fig1)

st.markdown("### 📌 Comparativo de ritmo diário")
fig2, ax2 = plt.subplots(figsize=(6, 2.6))
labels = ['Ritmo atual', 'Meta diária']
valores = [ritmo_atual, meta_diaria_uteis]
ax2.barh(labels, valores)
ax2.set_xlim(0, max(valores) * 1.5)
for i, v in enumerate(valores):
    ax2.text(v + 0.05, i, f"{v:.2f}/dia", va='center')
ax2.set_xlabel("Upgrades por dia")
ax2.grid(axis='x', linestyle='--', alpha=0.3)
st.pyplot(fig2)

st.markdown("### ⚖️ Simulador de esforço extra")
incremento = st.number_input("Quantos upgrades extras/dia útil?", min_value=0.0, value=2.0, step=0.5)
dias_extra = dias_uteis_extra_para_compensar(incremento)
if dias_extra == 0:
    st.success("🟢 Nenhum esforço adicional necessário. Ritmo atual é suficiente.")
elif dias_extra is None:
    st.warning("Informe um valor válido > 0 para calcular.")
else:
    data_fim_boost = np.busday_offset(HOJE.date(), dias_extra)
    st.info(f"📅 Com +{incremento:.2f}/dia útil, bastariam **{dias_extra} dias úteis** para compensar o atraso até **{data_fim_boost}**.")

st.markdown("### 🧮 Simulador de cenários")
cenarios = [4, 5, 6]
cols_sim = st.columns(len(cenarios))
for i, v in enumerate(cenarios):
    data_prev = simular_data_conclusao_por_produtividade(v)
    cols_sim[i].metric(f"{v} upgrades/dia útil", str(data_prev) if data_prev else "—")

custom = st.slider("Experimente outro valor:", 1, 20, 5)
custom_data = simular_data_conclusao_por_produtividade(custom)
st.write(f"📦 Com **{custom} upgrades/dia útil**, término em **{custom_data}**")
