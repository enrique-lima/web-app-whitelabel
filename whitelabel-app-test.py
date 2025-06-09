import streamlit as st
import pandas as pd
import unicodedata
import time
from io import BytesIO
from datetime import datetime
from dateutil.relativedelta import relativedelta
from statsmodels.tsa.api import ExponentialSmoothing
from pytrends.request import TrendReq
from bs4 import BeautifulSoup
import requests

# Verifica dependências
try:
    import openpyxl
except ImportError:
    st.error("Biblioteca 'openpyxl' não encontrada. Adicione 'openpyxl' no requirements.txt.")
    st.stop()

# Normaliza colunas
def normalizar_colunas(df):
    df.columns = [
        unicodedata.normalize("NFKD", c).encode("ASCII", "ignore").decode("utf-8").strip()
        .lower().replace(" ", "_") for c in df.columns
    ]
    return df

# Carrega vendas e estoque
@st.cache_data(show_spinner=False)
def carregar_dados(uploaded_file):
    df_venda = pd.read_excel(uploaded_file, sheet_name="VENDA", engine="openpyxl")
    df_estoque = pd.read_excel(uploaded_file, sheet_name="ESTOQUE", engine="openpyxl")
    df_venda = normalizar_colunas(df_venda)
    df_estoque = normalizar_colunas(df_estoque)
    meses_map = {"janeiro":1,"fevereiro":2,"marco":3,"abril":4,"maio":5,"junho":6,
                 "julho":7,"agosto":8,"setembro":9,"outubro":10,"novembro":11,"dezembro":12}
    df_venda['mes_num'] = df_venda['mes_venda'].str.lower().map(meses_map)
    df_venda = df_venda.dropna(subset=['ano_venda','mes_num'])
    df_venda['ano_mes'] = pd.to_datetime(
        df_venda['ano_venda'].astype(int).astype(str) + '-' + df_venda['mes_num'].astype(int).astype(str) + '-01'
    )
    return df_venda, df_estoque

# Obtém uplift de trends
@st.cache_data(show_spinner=False, max_entries=32, ttl=21600)
def get_trend_uplift(linhas_otb):
    pytrends = TrendReq(hl='pt-BR', tz=-180)
    genericos = [...]  # manter sua lista
    concorrentes = [...]  # manter sua lista
    registros=[]; tendencias={}
    for linha in linhas_otb:
        termos=[linha.lower()]+genericos+concorrentes
        try:
            pytrends.build_payload(termos, timeframe='today 3-m', geo='BR')
            df_tr = pytrends.interest_over_time()
            if not df_tr.empty:
                base=df_tr[linha.lower()].mean()
                gen=df_tr[genericos].mean(axis=1).mean()
                conc=df_tr[concorrentes].mean(axis=1).mean()
                uplift=((base+gen+conc)/3-50)/100
            else:
                base=gen=conc=uplift=0
        except:
            base=gen=conc=uplift=0
        tendencias[linha]=round(uplift,3)
        registros.append({'linha_otb':linha,'score_linha':round(base,2),'score_generico':round(gen,2),
                          'score_concorrente':round(conc,2),'uplift_aplicado':round(uplift,3)})
        time.sleep(1)
    return tendencias, pd.DataFrame(registros)

# Forecast de série
@st.cache_data(show_spinner=False, max_entries=128)
def forecast_serie(serie, passos, saz):
    if serie.count()>=24 and saz:
        modelo=ExponentialSmoothing(serie, trend='add', seasonal='add', seasonal_periods=12)
        prev=modelo.fit().forecast(passos)
    elif serie.count()>=6:
        modelo=ExponentialSmoothing(serie, trend='add', seasonal=None)
        prev=modelo.fit().forecast(passos)
    else:
        prev=pd.Series([serie.mean()]*passos,
            index=pd.date_range(serie.index[-1]+relativedelta(months=1), periods=passos, freq='MS'))
    return prev.clip(lower=0)

# Interface
st.image('https://raw.githubusercontent.com/enrique-lima/compra-moda-app/main/LOGO_TL.png', width=300)
st.title('Previsão de Vendas e Reposição de Estoque')

uploaded_file = st.file_uploader('📂 Faça upload do arquivo Excel', type=['xlsx'], key='tpl')
if uploaded_file:
    progresso = st.progress(0)
    status = st.empty()

    status.text('1/4 - Lendo e normalizando dados...')
    df_venda, df_estoque = carregar_dados(uploaded_file)
    progresso.progress(25)

    status.text('2/4 - Obtendo tendências Google Trends...')
    trend_uplift, df_trends = get_trend_uplift(df_venda['linha_otb'].dropna().unique().tolist())
    progresso.progress(50)

    status.text('3/4 - Calculando forecast por grupo...')
    peso = st.sidebar.slider('Peso Google Trends (%)', 0, 100, 100, 5) / 100
    saz = st.sidebar.checkbox('Considerar sazonalidade', True)
    periodos = 6
    records = []

    for (l, c, f), grp in df_venda.groupby(['linha_otb','cor_produto','filial']):
        serie = grp.set_index('ano_mes')['qtd_vendida'].resample('MS').sum().fillna(0)
        prev = forecast_serie(serie, periodos, saz)
        ajuste = trend_uplift.get(l, 0) * peso
        prev_adj = prev * (1 + ajuste)
        prev_adj = prev_adj.clip(lower=0)

        # Estoquepor grupo
        estoque_atual = int(
            df_estoque[
                (df_estoque['linha_otb']==l)&
                (df_estoque['cor_produto']==c)&
                (df_estoque['filial']==f)
            ]['saldo_empresa'].sum()
        )

        for date, val in prev_adj.items():
            cobertura = estoque_atual / val if val>0 else None
            compra = max(int(val - estoque_atual), 0)
            records.append({
                'linha_otb': l, 'cor_produto': c, 'filial': f,
                'mes': date.strftime('%Y-%m'), 'forecast': int(val),
                'estoque_atual': estoque_atual, 'cobertura_meses': round(cobertura,2) if cobertura else None,
                'compra_sugerida': compra
            })

    df_monthly = pd.DataFrame(records)
    progresso.progress(75)
    status.text('4/4 - Pronto!')
    st.success('Forecast gerado com sucesso!')

    # Download
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_monthly.to_excel(writer, sheet_name='Forecast_Mensal', index=False)
        df_trends.to_excel(writer, sheet_name='Tendencias', index=False)
    buffer.seek(0)
    progresso.progress(100)
    status.text('100% concluído')
    st.download_button(
        '⬇️ Baixar Forecast Mensal e Tendências', buffer.getvalue(),
        'output_forecast.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )
