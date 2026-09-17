
import streamlit as st
import pandas as pd
import os
import json

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

st.title("📈 V40 Trading Strategy Radar & Backtest Journal")
st.write("Run Screener Moat Fundamentals or Short-Term Geometric Pattern Scans with Watchlists & Sticky Notes.")

# Sidebar controls
st.sidebar.header("⚙️ Watchlist & Selection")

watchlist_choice = st.sidebar.selectbox(
    "Choose Watchlist Source:",
    ["🛡️ V40 Core Moats", "🚀 V40 Next Watchlist", "🎯 Custom Tickers"]
)

# Ticker Selection Logic
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
    ["1. Screener Moat Fundamentals", "2. Short-Term Geometric Pattern Radar"]
)

# Main Scan Button
if st.sidebar.button("🚀 Run Analysis", type="primary"):
    if not tickers:
        st.warning("Please select or enter at least one valid ticker symbol.")
    else:
        st.subheader(f"Scanning {len(tickers)} Tickers: {', '.join(tickers)}")
        
        if "1. Screener" in strategy:
            st.info("Fetching Screener.in live fundamentals...")
            out_file = "v40_automated_analysis.xlsx"
            v40_auto_fetcher.create_v40_analysis_workbook(tickers, out_file)
            
            df = pd.read_excel(out_file, skiprows=3)
            # Safely resolve Ticker column name
            ticker_col = 'Ticker' if 'Ticker' in df.columns else ('Ticker / Company' if 'Ticker / Company' in df.columns else df.columns)
            df['Backtest Sticky Notes'] = df[ticker_col].apply(lambda x: notes.get(str(x).upper(), ''))
            
            st.dataframe(df, use_container_width=True)
            df.to_excel(out_file, index=False)
            
            with open(out_file, "rb") as f:
                st.download_button(
                    label="📥 Download Screener Excel Report (.xlsx)",
                    data=f,
                    file_name=out_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        elif "2. Short-Term" in strategy:
            st.info("Scanning 45-day price candles for geometric breakouts...")
            out_file = "v40_geometric_analysis-v5.xlsx"
            v40_geometric_engine_v5.create_geometric_workbook(tickers, out_file)
            
            df = pd.read_excel(out_file, skiprows=3)
            # Safely resolve Ticker column name
            ticker_col = 'Ticker' if 'Ticker' in df.columns else ('Ticker / Company' if 'Ticker / Company' in df.columns else df.columns)
            df['Backtest Sticky Notes'] = df[ticker_col].apply(lambda x: notes.get(str(x).upper(), ''))
            
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
st.write("Write notes or backtest findings for any stock.")

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

