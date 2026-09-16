import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import numpy as np
    import pandas as pd
    from scipy.signal import find_peaks
    import yfinance as yf
    HAS_LIBS = True
except ImportError:
    HAS_LIBS = False


def safe_1d(df, col_name):
    """
    Extracts a column from df as a guaranteed 1D float numpy array.
    Handles DataFrame columns, Series, multi-index columns, and 2D arrays.
    """
    if col_name not in df.columns:
        matches = [c for c in df.columns if (isinstance(c, tuple) and c == col_name) or c == col_name]
        if matches:
            sub = df[matches]
        else:
            return np.array([], dtype=float)
    else:
        sub = df[col_name]
    
    if isinstance(sub, pd.DataFrame):
        sub = sub.iloc[:, 0]
    
    return np.asarray(sub, dtype=float).ravel()


def analyze_geometric_patterns(ticker, lookback_days=45, max_breakout_pct=3.0):
    """
    Scans a stock's daily price candles for FRESH, active geometric structures 
    (W-Pattern, Reverse Head & Shoulders, Cup with Handle) with strict 2nd Green Candle 
    and Failed Breakout verification.
    """
    ticker = ticker.strip().upper()
    
    result = {
        'ticker': ticker,
        'pattern_type': 'None',
        'cmp': 1000.0,
        'neckline_price': 0.0,
        'dist_to_breakout_pct': 0.0,
        'breakout_status': 'No Pattern',
        'projected_target': 0.0,
        'anchor_dates': 'N/A',
        'action_signal': 'NO GEOMETRIC SETUP'
    }

    if not HAS_LIBS:
        return result

    try:
        yf_ticker = f"{ticker}.NS"
        df = yf.Ticker(yf_ticker).history(period="6mo")

        if df.empty or len(df) < lookback_days:
            return result

        # Extract 1D price arrays safely
        close_arr = safe_1d(df, 'Close')
        open_arr = safe_1d(df, 'Open')
        if len(close_arr) < 2 or len(open_arr) < 2:
            return result

        cmp_val = round(float(close_arr[-1]), 2)
        result['cmp'] = cmp_val

        # Tight recent lookback window (default 45 trading days ~ 2 months)
        df_recent = df.tail(lookback_days).copy()
        prices = safe_1d(df_recent, 'Close')
        opens = safe_1d(df_recent, 'Open')
        lows = safe_1d(df_recent, 'Low')
        highs = safe_1d(df_recent, 'High')
        dates = [d.strftime('%Y-%m-%d') for d in df_recent.index]

        if len(prices) < 15 or len(lows) < 15 or len(highs) < 15:
            return result

        # Candle color checks for latest 2 candles
        is_green_today = prices[-1] > opens[-1]
        is_green_prev = prices[-2] > opens[-2]

        # Find local peaks and troughs
        std_val = float(np.std(prices))
        prom = std_val * 0.25 if std_val > 0 else 1.0
        
        peaks, _ = find_peaks(prices, distance=5, prominence=prom)
        troughs, _ = find_peaks(-prices, distance=5, prominence=prom)

        peaks = [int(p) for p in peaks]
        troughs = [int(t) for t in troughs]

        pattern_found = False

        # -------------------------------------------------------------
        # 1. CHECK FOR FRESH W-PATTERN (DOUBLE BOTTOM)
        # -------------------------------------------------------------
        if len(troughs) >= 2:
            for i in range(len(troughs) - 2, -1, -1):
                t1, t2 = troughs[i], troughs[i+1]
                
                # Recency check: Right bottom (L2) must be within last 22 trading days
                if (len(prices) - t2) > 22:
                    continue

                p1_val, p2_val = float(lows[t1]), float(lows[t2])
                min_bottom = min(p1_val, p2_val)

                if min_bottom > 0:
                    diff_pct = abs(p1_val - p2_val) / min_bottom
                    if diff_pct <= 0.04:
                        between_peaks = [p for p in peaks if t1 < p < t2]
                        if between_peaks:
                            nk_idx = max(between_peaks, key=lambda p: float(highs[p]))
                            neckline = round(float(highs[nk_idx]), 2)
                            
                            height = neckline - min_bottom
                            target = round(neckline + height, 2)
                            dist_pct = round(((cmp_val - neckline) / neckline) * 100, 1)

                            d_t1 = dates[t1]
                            d_nk = dates[nk_idx]
                            d_t2 = dates[t2]

                            result['pattern_type'] = 'Fresh W-Pattern (Double Bottom)'
                            result['neckline_price'] = neckline
                            result['dist_to_breakout_pct'] = dist_pct
                            result['projected_target'] = target
                            result['anchor_dates'] = f"{d_t1} (L1) / {d_nk} (Nk) / {d_t2} (L2)"

                            # --- STRICT BREAKOUT CONFIRMATION LOGIC ---
                            prev_close = float(prices[-2])
                            
                            # Case A: Price is currently above Neckline
                            if cmp_val >= neckline:
                                if dist_pct > max_breakout_pct:
                                    continue  # Already done and dusted
                                
                                if is_green_today and is_green_prev:
                                    result['breakout_status'] = 'CONFIRMED BREAKOUT (2 Green Candles)'
                                    result['action_signal'] = 'BUY MORNING (2nd Green Candle)'
                                elif is_green_today:
                                    result['breakout_status'] = 'INITIAL BREAKOUT (1st Green Candle)'
                                    result['action_signal'] = 'WATCHLIST (Wait for 2nd Green Candle)'
                                else:
                                    result['breakout_status'] = 'PULLBACK AT NECKLINE (Red Candle)'
                                    result['action_signal'] = 'WATCHLIST (Pullback)'
                            
                            # Case B: Price fell below Neckline
                            else:
                                if prev_close >= neckline:
                                    result['breakout_status'] = 'FAILED BREAKOUT (Slipped Below Neckline)'
                                    result['action_signal'] = 'AVOID (Breakout Failed)'
                                elif dist_pct >= -4.0:
                                    result['breakout_status'] = 'APPROACHING BREAKOUT'
                                    result['action_signal'] = f'WATCHLIST (Alert at ₹{neckline})'
                                else:
                                    result['breakout_status'] = 'FORMING RIGHT LEG'
                                    result['action_signal'] = 'WATCHLIST ONLY'

                            pattern_found = True
                            break

        # -------------------------------------------------------------
        # 2. CHECK FOR FRESH REVERSE HEAD & SHOULDERS
        # -------------------------------------------------------------
        if not pattern_found and len(troughs) >= 3:
            for i in range(len(troughs) - 3, -1, -1):
                s1_idx, h_idx, s2_idx = troughs[i], troughs[i+1], troughs[i+2]
                
                if (len(prices) - s2_idx) > 22:
                    continue

                s1_p, h_p, s2_p = float(lows[s1_idx]), float(lows[h_idx]), float(lows[s2_idx])

                if h_p < s1_p and h_p < s2_p:
                    min_shoulder = min(s1_p, s2_p)
                    if min_shoulder > 0 and abs(s1_p - s2_p) / min_shoulder <= 0.05:
                        p1_between = [p for p in peaks if s1_idx < p < h_idx]
                        p2_between = [p for p in peaks if h_idx < p < s2_idx]

                        if p1_between and p2_between:
                            nk1_idx = max(p1_between, key=lambda p: float(highs[p]))
                            nk2_idx = max(p2_between, key=lambda p: float(highs[p]))
                            nk1 = float(highs[nk1_idx])
                            nk2 = float(highs[nk2_idx])
                            neckline = round(float((nk1 + nk2) / 2.0), 2)

                            height = neckline - h_p
                            target = round(neckline + height, 2)
                            dist_pct = round(((cmp_val - neckline) / neckline) * 100, 1)

                            prev_close = float(prices[-2])

                            result['pattern_type'] = 'Fresh Reverse H&S'
                            result['neckline_price'] = neckline
                            result['dist_to_breakout_pct'] = dist_pct
                            result['projected_target'] = target
                            result['anchor_dates'] = f"{dates[s1_idx]} (LS) / {dates[h_idx]} (H) / {dates[s2_idx]} (RS)"

                            if cmp_val >= neckline:
                                if dist_pct > max_breakout_pct:
                                    continue
                                
                                if is_green_today and is_green_prev:
                                    result['breakout_status'] = 'CONFIRMED BREAKOUT (2 Green Candles)'
                                    result['action_signal'] = 'BUY MORNING (2nd Green Candle)'
                                else:
                                    result['breakout_status'] = 'PULLBACK AT NECKLINE (Red Candle)'
                                    result['action_signal'] = 'WATCHLIST (Pullback)'
                            else:
                                if prev_close >= neckline:
                                    result['breakout_status'] = 'FAILED BREAKOUT (Slipped Below Neckline)'
                                    result['action_signal'] = 'AVOID (Breakout Failed)'
                                else:
                                    result['breakout_status'] = 'APPROACHING BREAKOUT'
                                    result['action_signal'] = f'WATCHLIST (Alert at ₹{neckline})'

                            pattern_found = True
                            break

        if not pattern_found:
            result['action_signal'] = 'NO ACTIVE GEOMETRIC SETUP'

        print(f"  [Short-Term Radar] {ticker}: {result['pattern_type']} | CMP: ₹{result['cmp']} | Status: {result['breakout_status']}")

    except Exception as e:
        print(f"  [Short-Term Radar Notice] Error processing {ticker}: {e}")

    return result


def create_geometric_workbook(ticker_list, output_filename="v40_geometric_analysis-v5.xlsx"):
    """
    Generates a formatted Excel workbook containing short-term fresh pattern scan results 
    with custom color fills for Breakout Status and Action Signals.
    """
    wb = openpyxl.Workbook()
    
    font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_regular = Font(name="Calibri", size=11)
    
    # Custom Color Fills & Fonts
    fill_navy = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_blue_head = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    fill_section = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    fill_zebra = PatternFill(start_color="F2F4F8", end_color="F2F4F8", fill_type="solid")
    
    # Status & Signal Highlights
    fill_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") # Soft Green
    font_green = Font(name="Calibri", size=11, bold=True, color="006100")

    fill_yellow = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid") # Soft Yellow
    font_yellow = Font(name="Calibri", size=11, bold=True, color="9C6500")

    fill_red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") # Soft Red
    font_red = Font(name="Calibri", size=11, bold=True, color="9C0006")

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center")
    
    thin_side = Side(border_style="thin", color="D9D9D9")
    thick_bottom = Side(border_style="medium", color="1F4E78")
    
    border_data = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_header = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thick_bottom)

    ws = wb.active
    ws.title = "Short-Term Pattern Radar"
    
    try:
        ws.sheet_view.showGridLines = True
    except Exception:
        pass

    # Title Banner
    ws.merge_cells("A1:I1")
    t_cell = ws.cell(row=1, column=1, value="V40 SHORT-TERM GEOMETRIC PATTERN RADAR (COLOR-CODED SIGNALS)")
    t_cell.font = font_title
    t_cell.fill = fill_navy
    t_cell.alignment = align_center
    ws.row_dimensions[1].height = 35

    # Subtitle
    ws.merge_cells("A2:I2")
    sub_cell = ws.cell(row=2, column=1, value="Green = Confirmed Breakout / Buy | Yellow = Watchlist / Pullback | Red = Failed Breakout / Avoid")
    sub_cell.font = Font(name="Calibri", size=10, italic=True, color="FFFFFF")
    sub_cell.fill = fill_blue_head
    sub_cell.alignment = align_center
    ws.row_dimensions[2].height = 20

    headers = [
        "Ticker / Company", "Pattern Type", "CMP (₹)", "Neckline / Rim (₹)", 
        "Distance to Breakout", "Breakout Status", "Projected Target (₹)", 
        "Key Anchor Dates", "Action Signal"
    ]

    ws.row_dimensions[4].height = 28
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_section
        cell.alignment = align_center
        cell.border = border_header

    for idx, ticker in enumerate(ticker_list):
        current_row = 5 + idx
        ws.row_dimensions[current_row].height = 22

        res = analyze_geometric_patterns(ticker)

        c1 = ws.cell(row=current_row, column=1, value=res['ticker'])
        c2 = ws.cell(row=current_row, column=2, value=res['pattern_type'])
        c3 = ws.cell(row=current_row, column=3, value=res['cmp'])
        c4 = ws.cell(row=current_row, column=4, value=res['neckline_price'])
        c5 = ws.cell(row=current_row, column=5, value=res['dist_to_breakout_pct'] / 100.0 if res['pattern_type'] != 'None' else 0.0)
        c6 = ws.cell(row=current_row, column=6, value=res['breakout_status'])
        c7 = ws.cell(row=current_row, column=7, value=res['projected_target'])
        c8 = ws.cell(row=current_row, column=8, value=res['anchor_dates'])
        c9 = ws.cell(row=current_row, column=9, value=res['action_signal'])

        c1.alignment = align_left
        c2.alignment = align_left
        c3.number_format = '₹#,##0.0'
        c4.number_format = '₹#,##0.0'
        c5.number_format = '0.0%'
        c6.alignment = align_center
        c7.number_format = '₹#,##0.0'
        c8.alignment = align_center
        c9.alignment = align_left

        # Default styling
        for col_i in range(1, 10):
            cell = ws.cell(row=current_row, column=col_i)
            cell.font = font_regular
            cell.border = border_data
            if idx % 2 == 1:
                cell.fill = fill_zebra

        # --- COLOR CODING LOGIC FOR BREAKOUT STATUS (COL 6) ---
        status_str = str(res['breakout_status']).upper()
        if 'CONFIRMED BREAKOUT' in status_str or 'FRESH BREAKOUT' in status_str:
            c6.fill = fill_green
            c6.font = font_green
        elif 'FAILED BREAKOUT' in status_str or 'SLIPPED' in status_str:
            c6.fill = fill_red
            c6.font = font_red
        elif 'PULLBACK' in status_str or 'APPROACHING' in status_str or 'FORMING' in status_str or 'INITIAL' in status_str:
            c6.fill = fill_yellow
            c6.font = font_yellow

        # --- COLOR CODING LOGIC FOR ACTION SIGNAL (COL 9) ---
        signal_str = str(res['action_signal']).upper()
        if 'BUY MORNING' in signal_str or 'BUY' in signal_str:
            c9.fill = fill_green
            c9.font = font_green
        elif 'AVOID' in signal_str or 'FAILED' in signal_str:
            c9.fill = fill_red
            c9.font = font_red
        elif 'WATCHLIST' in signal_str or 'WAIT' in signal_str:
            c9.fill = fill_yellow
            c9.font = font_yellow

    col_widths = [18, 28, 14, 18, 20, 32, 18, 38, 32]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(output_filename)
    print(f"\n✅ Automatically created Short-Term Geometric Pattern Workbook: {output_filename}")


if __name__ == "__main__":
    print("=================================================================")
    print("  AUTOMATED V40 SHORT-TERM GEOMETRIC PATTERN RADAR")
    print("=================================================================")

    sample_tickers = ["CEATLTD", "COALINDIA", "TATAPOWER", "APLLTD"]
    create_geometric_workbook(sample_tickers, "v40_geometric_analysis-v5.xlsx")
