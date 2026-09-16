
import streamlit as st
import pandas as pd
import os

# Import standalone strategy modules
import v40_auto_fetcher
import v40_geometric_engine_v5

st.set_page_config(page_title="V40 Trading Strategy Radar", page_icon="📈", layout="wide")

st.title("📈 V40 Trading Strategy Radar")
st.write("Run Screener Moat Fundamentals or Short-Term Geometric Pattern Scans on Demand.")

# Sidebar controls
st.sidebar.header("⚙️ Watchlist & Controls")

ticker_input = st.sidebar.text_area(
    "Enter Stock Tickers (comma separated):",
    value="CEATLTD, COALINDIA, TATAPOWER, APLLTD",
    height=120
)

tickers = [t.strip().upper() for t in ticker_input.split(",") if t.strip()]

strategy = st.sidebar.radio(
    "Select Strategy Engine:",
    ["1. Screener Moat Fundamentals", "2. Short-Term Geometric Pattern Radar"]
)

if st.sidebar.button("🚀 Run Analysis", type="primary"):
    if not tickers:
        st.warning("Please enter at least one valid ticker symbol.")
    else:
        st.subheader(f"Scanning {len(tickers)} Tickers: {', '.join(tickers)}")
        
        if "1. Screener" in strategy:
            st.info("Fetching Screener.in live fundamentals...")
            out_file = "v40_automated_analysis.xlsx"
            v40_auto_fetcher.create_v40_analysis_workbook(tickers, out_file)
            
            # Display results
            df = pd.read_excel(out_file, skiprows=3)
            st.dataframe(df, use_container_width=True)
            
            # Download button
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
            st.dataframe(df, use_container_width=True)
            
            with open(out_file, "rb") as f:
                st.download_button(
                    label="📥 Download Geometric Pattern Radar Excel (.xlsx)",
                    data=f,
                    file_name=out_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )



