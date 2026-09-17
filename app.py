
import streamlit as st
import pandas as pd
import numpy as np
import os
import json

try:
    import yfinance as yf
    HAS_YFINANCE = True
except ImportError:
    HAS_YFINANCE = False

# Import standalone strategy modules
import v40_auto_fetcher
import v40_geometric_engine_v5

st.set_page_config(page_title="V40 Trading Strategy Radar", page_icon="📈", layout="wide")

# Persistent Notes Helper Functions
NOTES_FILE = "notes.json"

def load_notes():
    if os.path.exists(NOTES_FILE):
        try:
            with open(NOTES_FILE, "r") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def save_notes(notes):
    with open(NOTES_FILE, "w") as f:
        json.dump(notes, f, indent=4)

notes = load_notes()

def get_ticker_col(df):
    for col in ['Ticker', 'Ticker / Company', 'ticker', 'COMPANY']:
        if col in df.columns:
            return col
    return df.columns if len(df.columns) > 0 else None

# Master Watchlist Definitions
V40_CORE = [
    "ASIANPAINT", "TITAN", "PIDILITIND", "HDFCBANK", "ICICIBANK", "TCS", "INFY", 
    "HINDUNILVR", "NESTLEIND", "BAJFINANCE", "BAJAJFINSV", "HCLTECH", "RELIANCE", 
    "KOTAKBANK", "DABUR", "BRITANNIA", "HAVELLS", "BERGEPAINT", "BAJAJ-AUTO", 
    "MARUTI", "EICHERMOT", "POLYCAB", "SUPREMEIND", "ASTRAL", "TRENT", "PAGEIND", 
    "COALINDIA", "BATAINDIA", "TATAPOWER", "APLLTD"
]

V40_NEXT = [
    "RELAXO", "INDIGOPNTS", "BECTORFOOD", "5PAISA", "ALEMBICLTD", "AKZOINDIA", 
    "ABBOTTINDIA", "FLUOROCHEM", "CERA", "ERIS", "FINEORG", "DEEPAKNTR", 
    "VINATIORGA", "KEI", "VIPIND", "SONACOMS", "METROPOLIS", "LALPATHLAB", 
    "CLEAN", "LAURUSLABS", "RAJRATAN", "MAPMYINDIA", "CRAFTSMAN", "CAMPUS", "HAPPSTMNDS"
]

def analyze_v20_sma(ticker_list):
    """
    Scans tickers for 3-SMA Moving Average Alignment (20, 50, 200) and V20 Continuous Green Bunch Swings.
    """
    results = []
    for ticker in ticker_list:
        ticker = ticker.strip().upper()
        res = {
            'Ticker / Company': ticker,
            'CMP (₹)': 0.0,
            '20 SMA (₹)': 0.0,
            '50 SMA (₹)': 0.0,
            '200 SMA (₹)': 0.0,
            'Dist to 200 SMA (%)': 0.0,
            '3-SMA Alignment State': 'NO DATA',
            'SMA Trigger Signal': 'HOLD / NO TRIGGER',
            'V20 Bunch Found': 'No',
            'Bunch Growth %': 0.0,
            'Bottom Support (₹)': 0.0,
            'V20 Swing Signal': 'NO V20 BUNCH'
        }
        
        if not HAS_YFINANCE:
            results.append(res)
            continue
            
        try:
            yf_ticker = f"{ticker}.NS"
            df = yf.Ticker(yf_ticker).history(period="1y")
            
            if df.empty or len(df) < 50:
                results.append(res)
                continue
                
            close_arr = df['Close'].values
            open_arr = df['Open'].values
            high_arr = df['High'].values
            low_arr = df['Low'].values
            
            cmp_val = round(float(close_arr[-1]), 2)
            res['CMP (₹)'] = cmp_val
            
            sma200 = round(float(pd.Series(close_arr).rolling(200).mean().iloc[-1]), 2) if len(close_arr) >= 200 else cmp_val
            sma50 = round(float(pd.Series(close_arr).rolling(50).mean().iloc[-1]), 2) if len(close_arr) >= 50 else cmp_val
            sma20 = round(float(pd.Series(close_arr).rolling(20).mean().iloc[-1]), 2) if len(close_arr) >= 20 else cmp_val
            
            res['20 SMA (₹)'] = sma20
            res['50 SMA (₹)'] = sma50
            res['200 SMA (₹)'] = sma200
            
            dist_200 = round(((cmp_val - sma200) / sma200) * 100, 1)
            res['Dist to 200 SMA (%)'] = dist_200
            
            # SMA Alignment Logic
            if sma200 > sma50 > sma20 > cmp_val:
                res['3-SMA Alignment State'] = 'PESSIMISTIC (200 > 50 > 20 > CMP)'
                res['SMA Trigger Signal'] = 'BUY 1ST TRANCHE (Counter-Intuitive)'
            elif cmp_val > sma20 > sma50 > sma200:
                res['3-SMA Alignment State'] = 'OPTIMISTIC (CMP > 20 > 50 > 200)'
                res['SMA Trigger Signal'] = 'PROFIT TAKING / OPTIMISTIC'
            elif dist_200 <= -10.0:
                res['3-SMA Alignment State'] = 'DEEP DISCOUNT (-10% below 200 SMA)'
                res['SMA Trigger Signal'] = 'AVERAGE DOWN (2nd Tranche)'
            else:
                res['3-SMA Alignment State'] = 'NEUTRAL / TRANSITIONAL'
                res['SMA Trigger Signal'] = 'HOLD / NO TRIGGER'
                
            # V20 Green Bunch Analysis (Last 60 Days)
            sub_df = df.tail(60).copy()
            sub_closes = sub_df['Close'].values
            sub_opens = sub_df['Open'].values
            sub_lows = sub_df['Low'].values
            sub_highs = sub_df['High'].values
            
            max_bunch_gain = 0.0
            best_bunch_bot = 0.0
            best_bunch_top = 0.0
            
            cur_low = sub_lows if len(sub_lows) > 0 else 0.0
            cur_high = sub_highs if len(sub_highs) > 0 else 0.0
            
            for i in range(len(sub_closes)):
                if sub_closes[i] >= sub_opens[i]:
                    cur_high = max(cur_high, sub_highs[i])
                    if cur_low > 0:
                        gain_pct = ((cur_high - cur_low) / cur_low) * 100
                        if gain_pct > max_bunch_gain:
                            max_bunch_gain = gain_pct
                            best_bunch_bot = cur_low
                            best_bunch_top = cur_high
                else:
                    if i + 1 < len(sub_lows):
                        cur_low = sub_lows[i+1]
                        cur_high = sub_highs[i+1]
                        
            if max_bunch_gain >= 18.0:
                res['V20 Bunch Found'] = 'Yes (≥20% Bunch)'
                res['Bunch Growth %'] = round(max_bunch_gain, 1)
                res['Bottom Support (₹)'] = round(best_bunch_bot, 2)
                
                dist_to_bot = ((cmp_val - best_bunch_bot) / best_bunch_bot) * 100
                if dist_to_bot <= 3.0:
                    res['V20 Swing Signal'] = 'BUY AT BOTTOM SUPPORT'
                elif dist_to_bot <= 8.0:
                    res['V20 Swing Signal'] = 'NEAR BOTTOM SUPPORT'
                else:
                    res['V20 Swing Signal'] = 'IN SWING RANGE'
            else:
                res['V20 Bunch Found'] = 'No'
                res['V20 Swing Signal'] = 'NO V20 BUNCH'
                
        except Exception:
            pass
            
        results.append(res)
        
    return pd.DataFrame(results)

# Dashboard UI
st.title("📈 V40 Trading Strategy Radar & Backtest Journal")
st.write("Run Fundamental Moat Screening, V20 & SMA Moving Average Swings, or Geometric Breakouts.")

# Sidebar Controls
st.sidebar.header("⚙️ Watchlist & Selection")

watchlist_choice = st.sidebar.selectbox(
    "Choose Watchlist Source:",
    ["🛡️ V40 Core Moats", "🚀 V40 Next Watchlist", "🎯 Custom Tickers"]
)

if watchlist_choice == "🛡️ V40 Core Moats":
    base_tickers = V40_CORE
elif watchlist_choice == "🚀 V40 Next Watchlist":
    base_tickers = V40_NEXT
else:
    base_tickers = []

if watchlist_choice != "🎯 Custom Tickers":
    st.sidebar.markdown("**Filter Tickers:**")
    selected_subset = st.sidebar.multiselect(
        "Select specific company/companies (leave blank to analyze ALL in list):",
        options=base_tickers,
        default=[]
    )
    tickers = selected_subset if selected_subset else base_tickers
else:
    ticker_input = st.sidebar.text_area(
        "Enter Custom Tickers (comma separated):",
        value="CEATLTD, COALINDIA, TATAPOWER, APLLTD",
        height=120
    )
    tickers = [t.strip().upper() for t in ticker_input.split(",") if t.strip()]

st.sidebar.markdown(f"**Selected Companies Count:** `{len(tickers)}`")

strategy = st.sidebar.radio(
    "Select Strategy Engine:",
    [
        "1. 🛡️ Fundamental Moat Screener (V40 / V40 Next)",
        "2. 📉 Trend & Moving Average Radar (V20 & SMA)",
        "3. 📐 Short-Term Geometric Pattern Radar"
    ]
)

# Execution Logic
if st.sidebar.button("🚀 Run Analysis", type="primary"):
    if not tickers:
        st.warning("Please select or enter at least one valid ticker symbol.")
    else:
        st.subheader(f"Scanning {len(tickers)} Tickers: {', '.join(tickers)}")
        
        # Engine 1: Screener Fundamentals
        if "1. 🛡️ Fundamental" in strategy:
            st.info("Fetching Screener.in live balance sheets & moat metrics...")
            out_file = "v40_automated_analysis.xlsx"
            v40_auto_fetcher.create_v40_analysis_workbook(tickers, out_file)
            
            try:
                df = pd.read_excel(out_file, skiprows=3)
            except Exception:
                df = pd.read_excel(out_file)
                
            t_col = get_ticker_col(df)
            if t_col:
                df['Backtest Sticky Notes'] = df[t_col].apply(lambda x: notes.get(str(x).upper(), ''))
                
            st.dataframe(df, use_container_width=True)
            df.to_excel(out_file, index=False)
            
            with open(out_file, "rb") as f:
                st.download_button(
                    label="📥 Download Screener Fundamental Excel Report (.xlsx)",
                    data=f,
                    file_name=out_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        # Engine 2: V20 & SMA Trend Radar
        elif "2. 📉 Trend" in strategy:
            st.info("Scanning 20/50/200 SMA alignment and V20 green bunch swings...")
            df = analyze_v20_sma(tickers)
            
            t_col = get_ticker_col(df)
            if t_col:
                df['Backtest Sticky Notes'] = df[t_col].apply(lambda x: notes.get(str(x).upper(), ''))
                
            st.dataframe(df, use_container_width=True)
            out_file = "v40_v20_sma_analysis.xlsx"
            df.to_excel(out_file, index=False)
            
            with open(out_file, "rb") as f:
                st.download_button(
                    label="📥 Download V20 & SMA Trend Excel Report (.xlsx)",
                    data=f,
                    file_name=out_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        # Engine 3: Geometric Pattern Radar
        elif "3. 📐 Short-Term" in strategy:
            st.info("Scanning 45-day price candles for geometric breakouts...")
            out_file = "v40_geometric_analysis-v5.xlsx"
            v40_geometric_engine_v5.create_geometric_workbook(tickers, out_file)
            
            try:
                df = pd.read_excel(out_file, skiprows=3)
            except Exception:
                df = pd.read_excel(out_file)
                
            t_col = get_ticker_col(df)
            if t_col:
                df['Backtest Sticky Notes'] = df[t_col].apply(lambda x: notes.get(str(x).upper(), ''))
                
            st.dataframe(df, use_container_width=True)
            df.to_excel(out_file, index=False)
            
            with open(out_file, "rb") as f:
                st.download_button(
                    label="📥 Download Geometric Pattern Radar Excel (.xlsx)",
                    data=f,
                    file_name=out_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# Sticky Notes & Backtest Journal Section
st.markdown("---")
st.header("📌 Sticky Notes & Backtest Journal")
st.write("Write notes or backtest findings for any stock. They persist across sessions and attach to your Excel exports.")

if tickers:
    selected_ticker = st.selectbox("Select Ticker to View/Edit Note:", tickers)
    current_note = notes.get(selected_ticker, "")
    new_note = st.text_area(f"📝 Post-it Note for {selected_ticker}:", value=current_note, height=120)
    
    if st.button("💾 Save Sticky Note"):
        notes[selected_ticker] = new_note
        save_notes(notes)
        st.success(f"Saved note for {selected_ticker}!")
        st.rerun()

    if current_note:
        st.info(f"**Saved Note for {selected_ticker}:**\n\n{current_note}")
