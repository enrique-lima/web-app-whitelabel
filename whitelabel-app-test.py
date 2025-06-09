import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime
from statsmodels.tsa.statespace.sarimax import SARIMAX
from pytrends.request import TrendReq
import requests
from bs4 import BeautifulSoup

st.set_page_config(page_title="Previsão de Vendas & Posicionamento de Marca", layout="wide")
st.title("📈 Previsão de Vendas com Google Trends & Reposição")

# --- Cache Data Loading ---
@st.cache_data(show_spinner=False)
def load_data(file_vendas, file_estoque):
    df_v = pd.read_excel(file_vendas)
    df_e = pd.read_excel(file_estoque)
    for df in (df_v, df_e):
        df.columns = (
            df.columns.str.strip().str.lower()
              .str.replace(" ", "_")
              .str.normalize('NFKD')
              .str.encode('ascii', errors='ignore').str.decode('utf-8')
        )
    df_v['date'] = pd.to_datetime(
        dict(year=df_v['ano_venda'], month=df_v['mes_venda'], day=1)
    )
    return df_v, df_e

# --- Cache Google Trends ---
@st.cache_data(show_spinner=False)
def fetch_trends(terms):
    pytrends = TrendReq(hl='pt-BR', tz=-180)
    frames = []
    for i in range(0, len(terms), 5):
        batch = terms[i:i+5]
        pytrends.build_payload(batch, timeframe='today 12-m')
        df = pytrends.interest_over_time().drop(columns=['isPartial'], errors='ignore')
        frames.append(df)
    trends = pd.concat(frames, axis=1)
    idx = trends.mean(axis=1)
    idx.name = 'trend_index'
    return idx

# --- Cache Forecast Computation ---
@st.cache_data(show_spinner=False)
def compute_forecast(sales_ts, trend_idx, periods=6):
    df_m = pd.concat([sales_ts, trend_idx], axis=1).dropna()
    model = SARIMAX(df_m[sales_ts.name], exog=df_m[['trend_index']], order=(1,1,1))
    res = model.fit(disp=False)
    last_trend = trend_idx.iloc[-3:].mean()
    future_exog = pd.DataFrame(
        {'trend_index': [last_trend]*periods},
        index=pd.date_range(
            start=sales_ts.index[-1] + pd.offsets.MonthBegin(),
            periods=periods, freq='MS'
        )
    )
    forecast = res.get_forecast(steps=periods, exog=future_exog).predicted_mean
    return forecast

# --- Sidebar Setup ---
st.sidebar.title("Configurações")

# --- Main Tabs ---
tab1, tab2 = st.tabs(["Forecast", "SERP Google"])

# Forecast Tab
with tab1:
    uploaded = st.file_uploader(
        "Upload: vendas históricas e estoque atual", type=["xlsx","xls"], accept_multiple_files=True
    )
    if uploaded and len(uploaded) == 2:
        df_vendas, df_estoque = load_data(uploaded[0], uploaded[1])
        key = 'codigo_produto'
        generic = ["acessorios","alpargata","anabela","mocassim","bolsa","bota","cinto",
                   "loafer","rasteira","sandalia","sapatilha","scarpin","tenis","meia",
                   "meia pata","salto","salto fino","salto normal","sapato tratorado",
                   "mule","oxford","papete","peep flat","slide","sandália spike",
                   "salto spike","papete spike"]
        competitors = ["Alexander Birman","Luiza Barcelos","Schutz","Arezzo","Carmen Steffens"]
        terms = generic + competitors

        prod = st.selectbox("Código do Produto", df_vendas[key].unique())
        df_p = df_vendas[df_vendas[key] == prod]
        sales_ts = df_p.groupby('date')['saldo_empresa'].sum().asfreq('MS').fillna(0)

        if st.sidebar.button("Atualizar Forecast com Trends"):
            trend_idx = fetch_trends(terms)
            forecast = compute_forecast(sales_ts, trend_idx)
            # Exibição
            st.subheader(f"Previsão de Vendas (c/ Trends) para {prod}")
            st.line_chart(pd.concat([sales_ts, forecast]))
            estoque = df_estoque.loc[df_estoque[key] == prod, 'saldo_empresa'].sum()
            need = max(0, forecast.sum() - estoque)
            st.write(f"Estoque atual: {estoque:.0f} unidades")
            st.write(f"Sugestão de compra (6m): {need:.0f} unidades")

            # Preparar arquivo de retorno
            df_out = pd.DataFrame({
                'date': forecast.index,
                'forecast': forecast.values
            })
            # Gravar Excel em memória
            buffer = BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_out.to_excel(writer, sheet_name='Forecast', index=False)
                pd.DataFrame({
                    'Sugestao_compra_6m': [need]
                }).to_excel(writer, sheet_name='Resumo', index=False)
                writer.save()
                buffer.seek(0)
            # Botão de download
            st.download_button(
                label='📥 Baixar Forecast e Sugestão',
                data=buffer,
                file_name=f'retorno_{prod}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

# SERP Tab
with tab2:
    st.header("🔍 SERP Google - Tatiana Loureiro")
    term = st.text_input("Digite o termo para pesquisar (ex: sandália spike)")
    max_pages = st.number_input("Número de páginas a vasculhar", min_value=1, max_value=10, value=3)
    if st.button("Buscar posição"):  
        ua = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36"
            )
        }
        found = False
        for page in range(max_pages):
            start = page * 10
            url = f"https://www.google.com/search?q={term.replace(' ','+')}" + f"&start={start}"
            res = requests.get(url, headers=ua)
            soup = BeautifulSoup(res.text, 'html.parser')
            results = soup.select('div.g')
            for idx, g in enumerate(results, start=1):
                if "tatiana loureiro" in g.get_text().lower():
                    position = start + idx
                    st.write(f"Encontrado na posição {position} (página {page+1})")
                    found = True
                    break
            if found:
                break
        if not found:
            st.write("Nenhum resultado da Tatiana Loureiro encontrado nas páginas buscadas.")
