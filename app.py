import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import yfinance as yf


st.set_page_config(page_title="오늘의 종목 대시보드", layout="wide")


KR_STOCKS = {
    "삼성전자": "005930.KS",
    "SK하이닉스": "000660.KS",
    "LG에너지솔루션": "373220.KS",
    "현대차": "005380.KS",
    "기아": "000270.KS",
    "네이버": "035420.KS",
    "카카오": "035720.KS",
    "셀트리온": "068270.KS",
    "POSCO홀딩스": "005490.KS",
    "KB금융": "105560.KS",
}

US_STOCKS = {
    "Apple": "AAPL",
    "Microsoft": "MSFT",
    "NVIDIA": "NVDA",
    "Amazon": "AMZN",
    "Alphabet": "GOOGL",
}

PERIOD_OPTIONS = {
    "7일": "7d",
    "30일": "1mo",
    "90일": "3mo",
}


st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@400;500;700&display=swap');

    html, body, [class*="css"], .stApp {
        font-family: 'Noto Sans KR', sans-serif;
    }

    .stock-card {
        background: linear-gradient(145deg, #ffffff 0%, #f6f8fb 100%);
        border: 1px solid #e5e7eb;
        border-radius: 18px;
        padding: 18px 20px;
        box-shadow: 0 10px 30px rgba(15, 23, 42, 0.06);
        min-height: 120px;
    }

    .stock-name {
        font-size: 1rem;
        font-weight: 700;
        color: #111827;
        margin-bottom: 10px;
    }

    .stock-price {
        font-size: 1.5rem;
        font-weight: 700;
        color: #111827;
        margin-bottom: 8px;
    }

    .stock-change {
        font-size: 1rem;
        font-weight: 700;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(ttl=900, show_spinner=False)
def load_market_data(tickers, period):
    return yf.download(
        tickers=tickers,
        period=period,
        interval="1d",
        auto_adjust=False,
        progress=False,
        group_by="ticker",
        threads=True,
    )


def get_stock_frame(raw_data, name_map):
    records = []

    for name, ticker in name_map.items():
        try:
            ticker_frame = raw_data[ticker].dropna(how="all").copy()
        except Exception:
            continue

        if ticker_frame.empty:
            continue

        ticker_frame["Date"] = ticker_frame.index
        ticker_frame["Stock"] = name
        ticker_frame["Ticker"] = ticker
        records.append(ticker_frame.reset_index(drop=True))

    if not records:
        return pd.DataFrame()

    combined = pd.concat(records, ignore_index=True)
    combined = combined.sort_values(["Stock", "Date"]).reset_index(drop=True)
    combined["ChangePct"] = combined.groupby("Stock")["Close"].pct_change() * 100
    return combined


def get_card_metrics(raw_data, name_map):
    metrics = []
    for name, ticker in name_map.items():
        try:
            ticker_frame = raw_data[ticker].dropna(how="all")
        except Exception:
            ticker_frame = pd.DataFrame()

        if ticker_frame.empty:
            metrics.append(
                {
                    "name": name,
                    "price": None,
                    "change_pct": None,
                }
            )
            continue

        current_price = float(ticker_frame["Close"].iloc[-1])

        if len(ticker_frame) < 2:
            change_pct = None
        else:
            prev_price = float(ticker_frame["Close"].iloc[-2])
            change_pct = ((current_price - prev_price) / prev_price) * 100 if prev_price else 0.0

        metrics.append(
            {
                "name": name,
                "price": current_price,
                "change_pct": change_pct,
            }
        )
    return metrics


def render_cards(metrics):
    st.subheader("한국 종목 한눈에 보기")
    columns = st.columns(5)

    for idx, metric in enumerate(metrics):
        if metric["change_pct"] is None:
            color = "#6b7280"
            change_text = "데이터 없음"
        else:
            color = "#FF4444" if metric["change_pct"] >= 0 else "#0066CC"
            sign = "+" if metric["change_pct"] >= 0 else ""
            change_text = f"{sign}{metric['change_pct']:.2f}%"

        if metric["price"] is None:
            price_text = "가격 정보 없음"
        else:
            price_text = f"{metric['price']:,.0f}원"

        with columns[idx % 5]:
            st.markdown(
                f"""
                <div class="stock-card">
                    <div class="stock-name">{metric["name"]}</div>
                    <div class="stock-price">{price_text}</div>
                    <div class="stock-change" style="color:{color};">
                        {change_text}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )


def make_line_chart(frame, title):
    fig = px.line(
        frame,
        x="Date",
        y="Close",
        color="Stock",
        title=title,
        markers=True,
    )
    fig.update_layout(
        font=dict(family="Noto Sans KR, sans-serif"),
        hovermode="x unified",
        legend_title_text="종목",
        xaxis_title="날짜",
        yaxis_title="종가",
        template="plotly_white",
        margin=dict(l=20, r=20, t=60, b=20),
    )
    return fig


def recent_days(frame, days=30):
    return (
        frame.sort_values(["Stock", "Date"])
        .groupby("Stock", group_keys=False)
        .tail(days)
        .reset_index(drop=True)
    )


def make_volatility_chart(frame, title):
    volatility = (
        frame.groupby("Stock", as_index=False)["ChangePct"]
        .std()
        .rename(columns={"ChangePct": "Volatility"})
        .dropna()
    )
    fig = go.Figure(
        data=[
            go.Bar(
                x=volatility["Stock"],
                y=volatility["Volatility"],
                marker_color="#1f77b4",
                text=volatility["Volatility"].round(2),
                textposition="outside",
            )
        ]
    )
    fig.update_layout(
        title=title,
        font=dict(family="Noto Sans KR, sans-serif"),
        xaxis_title="종목",
        yaxis_title="일간 변동성(표준편차, %)",
        template="plotly_white",
        margin=dict(l=20, r=20, t=60, b=20),
    )
    return fig


st.title("오늘의 종목 대시보드")

selected_label = st.sidebar.radio("조회 기간", options=list(PERIOD_OPTIONS.keys()), index=1)
selected_period = PERIOD_OPTIONS[selected_label]

kr_raw = load_market_data(list(KR_STOCKS.values()), selected_period)
us_raw = load_market_data(list(US_STOCKS.values()), selected_period)

kr_metrics = get_card_metrics(kr_raw, KR_STOCKS)
kr_frame = get_stock_frame(kr_raw, KR_STOCKS)
us_frame = get_stock_frame(us_raw, US_STOCKS)

if not any(metric["price"] is not None for metric in kr_metrics):
    st.error("한국 종목 데이터를 불러오지 못했습니다. 잠시 후 다시 시도해 주세요.")
    st.stop()

render_cards(kr_metrics)

st.subheader("종목별 종가 추이")
tab_kr, tab_us = st.tabs(["한국", "해외"])

with tab_kr:
    if kr_frame.empty:
        st.info("한국 종목 차트 데이터가 없습니다.")
    else:
        st.plotly_chart(
            make_line_chart(recent_days(kr_frame, 30), "한국 종목 최근 30일 종가 추이"),
            use_container_width=True,
        )

with tab_us:
    if us_frame.empty:
        st.info("해외 종목 차트 데이터가 없습니다.")
    else:
        st.plotly_chart(
            make_line_chart(recent_days(us_frame, 30), "해외 종목 최근 30일 종가 추이"),
            use_container_width=True,
        )

st.subheader("종목별 일간 변동성")
vol_tab_kr, vol_tab_us = st.tabs(["한국", "해외"])

with vol_tab_kr:
    if kr_frame.empty:
        st.info("한국 변동성 데이터를 표시할 수 없습니다.")
    else:
        st.plotly_chart(
            make_volatility_chart(kr_frame, f"한국 종목 일간 변동성 ({selected_label} 기준)"),
            use_container_width=True,
        )

with vol_tab_us:
    if us_frame.empty:
        st.info("해외 변동성 데이터를 표시할 수 없습니다.")
    else:
        st.plotly_chart(
            make_volatility_chart(us_frame, f"해외 종목 일간 변동성 ({selected_label} 기준)"),
            use_container_width=True,
        )
