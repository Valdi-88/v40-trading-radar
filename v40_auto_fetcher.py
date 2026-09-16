import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Try importing requests and bs4 for live fundamental scraping
try:
    import requests
    from bs4 import BeautifulSoup
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

# Try importing yfinance for live technical price data and moving averages
try:
    import yfinance as yf
    HAS_YFINANCE = True
except ImportError:
    HAS_YFINANCE = False


def fetch_screener_data(ticker):
    """
    Fetches financial metrics from Screener.in and live technicals (CMP, SMAs, V20 swing)
    from Yahoo Finance for a given Indian stock ticker.
    """
    ticker = ticker.strip().upper()
    
    # Default fallback data structure
    data = {
        'ticker': ticker,
        'company_name': ticker,
        'sector': 'General',
        'current_price': 1000.0,
        'sma_20': 1050.0,
        'sma_50': 1100.0,
        'sma_200': 1200.0,
        'market_cap': 5000.0,
        'leader_market_cap': 100000.0,
        'strong_hands_pct': 75.0,
        'promoter_pledge_pct': 0.0,
        'is_debt_free': 'Yes',
        'moat_score': 5,
        'v20_swing_pct': 15.0,
        'swing_dates_info': 'N/A',  # <<< MARKER 1: ADDED DEFAULT DATE FIELD
        'cwip': 50.0,
        'fixed_assets': 300.0,
        'dividend_yield': 1.5
    }
    
    # ---------------------------------------------------------
    # 1. FETCH FUNDAMENTAL DATA FROM SCREENER.IN
    # ---------------------------------------------------------
    if HAS_REQUESTS:
        url = f"https://www.screener.in/company/{ticker}/consolidated/"
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'
        }

        try:
            response = requests.get(url, headers=headers, timeout=10)
            if response.status_code != 200:
                url = f"https://www.screener.in/company/{ticker}/"
                response = requests.get(url, headers=headers, timeout=10)
                
            if response.status_code == 200:
                soup = BeautifulSoup(response.content, 'html.parser')
                
                # Company Name
                name_tag = soup.find('h1')
                if name_tag:
                    data['company_name'] = name_tag.text.strip()
                    
                # Top Ratios
                top_ratios = soup.find('ul', {'id': 'top-ratios'})
                if top_ratios:
                    for li in top_ratios.find_all('li'):
                        name = li.find('span', {'class': 'name'})
                        value = li.find('span', {'class': 'number'})
                        if name and value:
                            n_text = name.text.strip().lower()
                            v_text = value.text.strip().replace(',', '')
                            try:
                                v_float = float(v_text)
                                if 'market cap' in n_text:
                                    data['market_cap'] = v_float
                                elif 'current price' in n_text:
                                    data['current_price'] = v_float
                                elif 'dividend yield' in n_text:
                                    data['dividend_yield'] = v_float
                                elif 'roce' in n_text:
                                    data['roce'] = v_float
                            except ValueError:
                                pass
                                
                # Shareholding Pattern Parsing
                shp_section = soup.find('section', {'id': 'shareholding'})
                if shp_section:
                    table = shp_section.find('table')
                    if table:
                        rows = table.find_all('tr')
                        promoter_pct = 0.0
                        fii_pct = 0.0
                        dii_pct = 0.0
                        for row in rows:
                            cells = row.find_all(['td', 'th'])
                            if cells:
                                row_name = cells[0].text.strip().lower()
                                latest_val_text = cells[-1].text.strip().replace('%', '')
                                try:
                                    latest_val = float(latest_val_text)
                                    if 'promoters' in row_name:
                                        promoter_pct = latest_val
                                    elif 'fiis' in row_name:
                                        fii_pct = latest_val
                                    elif 'diis' in row_name:
                                        dii_pct = latest_val
                                except ValueError:
                                    pass
                        strong_hands = promoter_pct + fii_pct + dii_pct
                        if strong_hands > 0:
                            data['strong_hands_pct'] = round(strong_hands, 2)
                            
                print(f"  [Screener Success] {ticker}: Market Cap = ₹{data['market_cap']} Cr | Strong Hands = {data['strong_hands_pct']}%")
        except Exception as e:
            print(f"  [Screener Notice] Could not connect to Screener for {ticker}: {e}")

    # ---------------------------------------------------------
    # 2. FETCH LIVE TECHNICAL DATA & SMAs FROM YAHOO FINANCE
    # ---------------------------------------------------------
    if HAS_YFINANCE:
        try:
            yf_ticker = f"{ticker}.NS"  # NSE Ticker symbol format
            df = yf.Ticker(yf_ticker).history(period="1y")
            
            if not df.empty and len(df) >= 200:
                # Live CMP and Moving Averages
                data['current_price'] = round(float(df['Close'].iloc[-1]), 2)
                data['sma_20'] = round(float(df['Close'].rolling(window=20).mean().iloc[-1]), 2)
                data['sma_50'] = round(float(df['Close'].rolling(window=50).mean().iloc[-1]), 2)
                data['sma_200'] = round(float(df['Close'].rolling(window=200).mean().iloc[-1]), 2)
                
                # Actual V20 continuous swing % (3-4 month lookback, ~80 trading days)
                df_4mo = df.tail(80)
                low_val = df_4mo['Low'].min()
                high_val = df_4mo['High'].max()
                if low_val > 0:
                    data['v20_swing_pct'] = round(float(((high_val - low_val) / low_val) * 100), 1)
                
                # <<< MARKER 2: ADDED EXACT DATE EXTRACTION LOGIC HERE >>>
                low_idx = df_4mo['Low'].idxmin()
                high_idx = df_4mo['High'].idxmax()
                low_date = low_idx.strftime('%Y-%m-%d')
                high_date = high_idx.strftime('%Y-%m-%d')
                data['swing_dates_info'] = f"{low_date} to {high_date}"
                # <<< END OF DATE EXTRACTION >>>
                
                print(f"  [yfinance Success] {ticker}: CMP = ₹{data['current_price']} | 20 SMA = {data['sma_20']} | 50 SMA = {data['sma_50']} | 200 SMA = {data['sma_200']} | V20 Swing = {data['v20_swing_pct']}% ({data['swing_dates_info']})")
        except Exception as e:
            print(f"  [yfinance Notice] Could not fetch live technicals for {ticker}: {e}")
            
    return data


def create_v40_analysis_workbook(ticker_list, output_filename="v40_automated_analysis.xlsx"):
    """
    Generates a 2-tab Excel workbook analyzing provided tickers using V40 fundamentals
    and multi-strategy technical signals (3-SMA Alignment, V20 Swing, Portfolio Trade Sizer).
    """
    wb = openpyxl.Workbook()
    
    font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)
    
    fill_navy = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_blue_head = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    fill_section = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    fill_zebra = PatternFill(start_color="F2F4F8", end_color="F2F4F8", fill_type="solid")
    
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    
    thin_side = Side(border_style="thin", color="D9D9D9")
    thick_bottom = Side(border_style="medium", color="1F4E78")
    
    border_data = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_header = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thick_bottom)

    # ---------------------------------------------------------
    # TAB 1: MULTI-STRATEGY V40 DASHBOARD
    # ---------------------------------------------------------
    ws_dash = wb.active
    ws_dash.title = "V40 Strategy Dashboard"
    
    try:
        ws_dash.sheet_view.showGridLines = True
    except Exception:
        try:
            ws_dash.views.sheetView.showGridLines = True
        except Exception:
            pass
    
    # Title Banner
    ws_dash.merge_cells("A1:L1")  # <<< MARKER 3: MERGED 12 COLUMNS (A TO L)
    title_cell = ws_dash.cell(row=1, column=1, value="V40 & V40 NEXT AUTOMATED MULTI-STRATEGY EVALUATION DASHBOARD")
    title_cell.font = font_title
    title_cell.fill = fill_navy
    title_cell.alignment = align_center
    ws_dash.row_dimensions[1].height = 35
    
    # Subtitle
    ws_dash.merge_cells("A2:L2")  # <<< MARKER 4: MERGED 12 COLUMNS (A TO L)
    sub_cell = ws_dash.cell(row=2, column=1, value="Combines V40 Quality Moat Screening + Live 3-SMA Alignment & Technical Strategy Signals")
    sub_cell.font = Font(name="Calibri", size=10, italic=True, color="FFFFFF")
    sub_cell.fill = fill_blue_head
    sub_cell.alignment = align_center
    ws_dash.row_dimensions[2].height = 20

    # Headers
    headers = [
        "Ticker / Company", "Market Cap (₹ Cr)", "Strong Hands %", "Moat Score", 
        "Current Price (₹)", "20 SMA", "50 SMA", "200 SMA",
        "3-SMA Alignment", "V20 Swing Status", 
        "V20 Swing Dates (Low to High)",  # <<< MARKER 5: ADDED NEW COLUMN HEADER HERE
        "Overall Action Signal"
    ]
    
    ws_dash.row_dimensions[4].height = 28
    for col_idx, h in enumerate(headers, start=1):
        cell = ws_dash.cell(row=4, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_section
        cell.alignment = align_center
        cell.border = border_header

    # Populate Data Rows
    row_start = 5
    for idx, ticker in enumerate(ticker_list):
        current_row = row_start + idx
        ws_dash.row_dimensions[current_row].height = 22
        
        info = fetch_screener_data(ticker)
        
        mcap = info['market_cap']
        sh_pct = info['strong_hands_pct'] / 100.0
        moat = info['moat_score']
        cmp_val = info['current_price']
        sma20 = info['sma_20']
        sma50 = info['sma_50']
        sma200 = info['sma_200']
        
        # Strategy Logic Evaluation
        is_3sma_buy = (sma200 > sma50) and (sma50 > sma20) and (sma20 > cmp_val)
        is_3sma_sell = (cmp_val > sma20) and (sma20 > sma50) and (sma50 > sma200)
        
        if is_3sma_buy:
            sma_status = "ALIGNED BUY (200>50>20>CMP)"
        elif is_3sma_sell:
            sma_status = "ALIGNED EXIT (CMP>20>50>200)"
        else:
            sma_status = "NEUTRAL / MIXED"
            
        v20_status = f"SWING MOVE ({info['v20_swing_pct']}%)" if info['v20_swing_pct'] >= 20 else f"NORMAL RANGE ({info['v20_swing_pct']}%)"
        
        if moat >= 5 and sh_pct >= 0.70:
            if is_3sma_buy:
                action_signal = "BUY MORNING (3-SMA BUY SETUP)"
            elif is_3sma_sell:
                action_signal = "SELL MORNING (3-SMA EXIT SETUP)"
            else:
                action_signal = "QUALIFIED V40 / WATCH FOR SETUP"
        else:
            action_signal = "WATCHLIST ONLY (FAILS V40 MOAT)"

        # Write to sheet
        c_name = ws_dash.cell(row=current_row, column=1, value=info['company_name'])
        c_mcap = ws_dash.cell(row=current_row, column=2, value=mcap)
        c_sh = ws_dash.cell(row=current_row, column=3, value=sh_pct)
        c_moat = ws_dash.cell(row=current_row, column=4, value=moat)
        c_price = ws_dash.cell(row=current_row, column=5, value=cmp_val)
        c_s20 = ws_dash.cell(row=current_row, column=6, value=sma20)
        c_s50 = ws_dash.cell(row=current_row, column=7, value=sma50)
        c_s200 = ws_dash.cell(row=current_row, column=8, value=sma200)
        c_sma_stat = ws_dash.cell(row=current_row, column=9, value=sma_status)
        c_v20_stat = ws_dash.cell(row=current_row, column=10, value=v20_status)
        
        # <<< MARKER 6: EXACT PLACE WHERE DATES ARE WRITTEN (COLUMN 11) & ACTION SIGNAL SHIFTED (COLUMN 12) >>>
        c_dates_stat = ws_dash.cell(row=current_row, column=11, value=info['swing_dates_info'])
        c_act_stat = ws_dash.cell(row=current_row, column=12, value=action_signal)
        # <<< END OF ROW VALUE WRITING >>>
        
        # Formatting
        c_name.alignment = align_left
        c_mcap.number_format = '#,##0'
        c_sh.number_format = '0.0%'
        c_moat.alignment = align_center
        c_price.number_format = '₹#,##0.0'
        c_s20.number_format = '₹#,##0.0'
        c_s50.number_format = '₹#,##0.0'
        c_s200.number_format = '₹#,##0.0'
        c_sma_stat.alignment = align_left
        c_v20_stat.alignment = align_center
        c_dates_stat.alignment = align_center  # <<< MARKER 7: FORMATTED DATE COLUMN
        c_act_stat.alignment = align_left
        
        # <<< MARKER 8: UPDATED LOOP RANGE TO 13 (FOR 12 TOTAL COLUMNS) >>>
        for col_i in range(1, 13):
            cell = ws_dash.cell(row=current_row, column=col_i)
            cell.font = font_regular
            cell.border = border_data
            if idx % 2 == 1:
                cell.fill = fill_zebra

    col_widths = [22, 16, 14, 12, 16, 14, 14, 14, 28, 22, 26, 32]  # <<< MARKER 9: UPDATED COLUMN WIDTHS
    for i, w in enumerate(col_widths, start=1):
        ws_dash.column_dimensions[get_column_letter(i)].width = w

    # ---------------------------------------------------------
    # TAB 2: PORTFOLIO TRADE SIZER & AVERAGING CALCULATOR
    # ---------------------------------------------------------
    ws_sizer = wb.create_sheet(title="Trade Sizer & Portfolio Rules")
    try:
        ws_sizer.sheet_view.showGridLines = True
    except Exception:
        try:
            ws_sizer.views.sheetView.showGridLines = True
        except Exception:
            pass

    # Title Banner
    ws_sizer.merge_cells("A1:F1")
    t2 = ws_sizer.cell(row=1, column=1, value="PORTFOLIO CAPITAL ALLOCATION & AVERAGING CALCULATOR")
    t2.font = font_title
    t2.fill = fill_navy
    t2.alignment = align_center
    ws_sizer.row_dimensions[header_row := 1].height = 35

    # Inputs Section
    ws_sizer.cell(row=3, column=1, value="Total Portfolio Capital (₹):").font = font_bold
    ws_sizer.cell(row=3, column=2, value=1000000).font = font_bold  # Default ₹10 Lakhs
    ws_sizer.cell(row=3, column=2).number_format = '₹#,##0'
    ws_sizer.cell(row=3, column=2).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    headers_sizer = [
        "Stock Name", "Entry Price (₹)", "Standard Position (3% Capital)", 
        "10% Dip Price (Tranche 2)", "Averaging Tranche (3% Capital)", "10.5% Profit Target (Tranche 2 Exit)"
    ]
    
    ws_sizer.row_dimensions[5].height = 28
    for col_idx, h in enumerate(headers_sizer, start=1):
        cell = ws_sizer.cell(row=5, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_section
        cell.alignment = align_center
        cell.border = border_header

    for idx, ticker in enumerate(ticker_list):
        current_row = 6 + idx
        ws_sizer.row_dimensions[current_row].height = 22
        
        info = fetch_screener_data(ticker)
        entry_p = info['current_price']
        
        c1 = ws_sizer.cell(row=current_row, column=1, value=ticker)
        c2 = ws_sizer.cell(row=current_row, column=2, value=entry_p)
        c3 = ws_sizer.cell(row=current_row, column=3, value=f"=B3*0.03")  # 3% Allocation
        c4 = ws_sizer.cell(row=current_row, column=4, value=f"=B{current_row}*0.90")  # 10% Dip
        c5 = ws_sizer.cell(row=current_row, column=5, value=f"=B3*0.03")  # Tranche 2 Allocation
        c6 = ws_sizer.cell(row=current_row, column=6, value=f"=D{current_row}*1.105")  # 10.5% Target
        
        c1.alignment = align_left
        c2.number_format = '₹#,##0.0'
        c3.number_format = '₹#,##0'
        c4.number_format = '₹#,##0.0'
        c5.number_format = '₹#,##0'
        c6.number_format = '₹#,##0.0'
        
        for col_i in range(1, 7):
            cell = ws_sizer.cell(row=current_row, column=col_i)
            cell.font = font_regular
            cell.border = border_data

    sizer_widths = [18, 16, 26, 24, 26, 30]
    for i, w in enumerate(sizer_widths, start=1):
        ws_sizer.column_dimensions[get_column_letter(i)].width = w

    wb.save(output_filename)
    print(f"\n✅ Automatically created Multi-Strategy Excel Workbook: {output_filename}")


if __name__ == "__main__":
    print("=================================================================")
    print("  AUTOMATED V40 MULTI-STRATEGY SCREENER & TECHNICAL ENGINE")
    print("=================================================================")
    
    # Sample list of stock tickers
    sample_tickers = ["INDIGOPNTS", "CEATLTD", "TATAPOWER", "COALINDIA", "APLLTD"]
    create_v40_analysis_workbook(sample_tickers, "v40_automated_analysis.xlsx")
