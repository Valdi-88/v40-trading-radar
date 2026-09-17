
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
    PARALLEL PATTERN EVALUATION (OPTION 2: SEPARATE ROWS PER PATTERN):
    Scans a stock's daily price candles independently for ALL active geometric structures.
    Returns a LIST of pattern result dictionaries (one dict per detected pattern).
    If no pattern is found, returns a list containing one default 'No Pattern' dictionary.
    """
    ticker = ticker.strip().upper()
    
    default_no_pattern = {
        'ticker': ticker,
        'pattern_type': 'None',
        'cmp': 0.0,
        'neckline_price': 0.0,
        'dist_to_breakout_pct': 0.0,
        'breakout_status': 'No Pattern',
        'projected_target': 0.0,
        'anchor_dates': 'N/A',
        'action_signal': 'NO ACTIVE GEOMETRIC SETUP'
    }

    if not HAS_LIBS:
        return [default_no_pattern]

    try:
        yf_ticker = f"{ticker}.NS"
        df = yf.Ticker(yf_ticker).history(period="6mo")

        if df.empty or len(df) < lookback_days:
            return [default_no_pattern]

        close_arr = safe_1d(df, 'Close')
        open_arr = safe_1d(df, 'Open')
        if len(close_arr) < 2 or len(open_arr) < 2:
            return [default_no_pattern]

        cmp_val = round(float(close_arr[-1]), 2)
        default_no_pattern['cmp'] = cmp_val

        df_recent = df.tail(lookback_days).copy()
        prices = safe_1d(df_recent, 'Close')
        opens = safe_1d(df_recent, 'Open')
        lows = safe_1d(df_recent, 'Low')
        highs = safe_1d(df_recent, 'High')
        dates = [d.strftime('%Y-%m-%d') for d in df_recent.index]

        if len(prices) < 15 or len(lows) < 15 or len(highs) < 15:
            return [default_no_pattern]

        is_green_today = prices[-1] > opens[-1]
        is_green_prev = prices[-2] > opens[-2]
        prev_close = float(prices[-2])

        std_val = float(np.std(prices))
        prom = std_val * 0.25 if std_val > 0 else 1.0
        
        peaks, _ = find_peaks(prices, distance=5, prominence=prom)
        troughs, _ = find_peaks(-prices, distance=5, prominence=prom)

        peaks = [int(p) for p in peaks]
        troughs = [int(t) for t in troughs]

        detected_patterns = []

        # 1. INDEPENDENT CHECK: FRESH W-PATTERN (DOUBLE BOTTOM)
        if len(troughs) >= 2:
            for i in range(len(troughs) - 2, -1, -1):
                t1, t2 = troughs[i], troughs[i+1]
                if (len(prices) - t2) > 22:
                    continue

                p1_val, p2_val = float(lows[t1]), float(lows[t2])
                min_bottom = min(p1_val, p2_val)

                if min_bottom > 0 and (abs(p1_val - p2_val) / min_bottom) <= 0.05:
                    between_peaks = [p for p in peaks if t1 < p < t2]
                    if between_peaks:
                        nk_idx = max(between_peaks, key=lambda p: float(highs[p]))
                        neckline = round(float(highs[nk_idx]), 2)
                        height = neckline - min_bottom
                        target = round(neckline + height, 2)
                        dist_pct = round(((cmp_val - neckline) / neckline) * 100, 1)

                        if cmp_val >= neckline:
                            if dist_pct > max_breakout_pct:
                                continue
                            status = 'CONFIRMED BREAKOUT (2 Green)' if (is_green_today and is_green_prev) else ('INITIAL BREAKOUT (1st Green)' if is_green_today else 'PULLBACK AT NECKLINE')
                            signal = 'BUY MORNING (2nd Green)' if (is_green_today and is_green_prev) else ('WATCHLIST (Wait 2nd Green)' if is_green_today else 'WATCHLIST (Pullback)')
                        else:
                            if prev_close >= neckline:
                                status = 'FAILED BREAKOUT (Slipped Below)'
                                signal = 'AVOID (Breakout Failed)'
                            elif dist_pct >= -6.0:
                                status = 'APPROACHING BREAKOUT'
                                signal = f'WATCHLIST (Alert at ₹{neckline})'
                            else:
                                status = 'FORMING RIGHT LEG'
                                signal = 'WATCHLIST ONLY'

                        detected_patterns.append({
                            'ticker': ticker,
                            'pattern_type': 'Fresh W-Pattern',
                            'cmp': cmp_val,
                            'neckline_price': neckline,
                            'dist_to_breakout_pct': dist_pct,
                            'breakout_status': status,
                            'projected_target': target,
                            'anchor_dates': f"{dates[t1]} (L1) / {dates[nk_idx]} (Nk) / {dates[t2]} (L2)",
                            'action_signal': signal
                        })
                        break

        # 2. INDEPENDENT CHECK: FRESH REVERSE HEAD & SHOULDERS
        if len(troughs) >= 3:
            for i in range(len(troughs) - 3, -1, -1):
                s1_idx, h_idx, s2_idx = troughs[i], troughs[i+1], troughs[i+2]
                if (len(prices) - s2_idx) > 22:
                    continue

                s1_p, h_p, s2_p = float(lows[s1_idx]), float(lows[h_idx]), float(lows[s2_idx])

                if h_p < s1_p and h_p < s2_p and abs(s1_p - s2_p) / min(s1_p, s2_p) <= 0.06:
                    p1_between = [p for p in peaks if s1_idx < p < h_idx]
                    p2_between = [p for p in peaks if h_idx < p < s2_idx]

                    if p1_between and p2_between:
                        nk1_idx = max(p1_between, key=lambda p: float(highs[p]))
                        nk2_idx = max(p2_between, key=lambda p: float(highs[p]))
                        neckline = round(float((highs[nk1_idx] + highs[nk2_idx]) / 2.0), 2)
                        height = neckline - h_p
                        target = round(neckline + height, 2)
                        dist_pct = round(((cmp_val - neckline) / neckline) * 100, 1)

                        if cmp_val >= neckline:
                            if dist_pct > max_breakout_pct:
                                continue
                            status = 'CONFIRMED BREAKOUT (2 Green)' if (is_green_today and is_green_prev) else 'PULLBACK AT NECKLINE'
                            signal = 'BUY MORNING (2nd Green)' if (is_green_today and is_green_prev) else 'WATCHLIST (Pullback)'
                        else:
                            if prev_close >= neckline:
                                status = 'FAILED BREAKOUT (Slipped Below)'
                                signal = 'AVOID (Breakout Failed)'
                            else:
                                status = 'APPROACHING BREAKOUT'
                                signal = f'WATCHLIST (Alert at ₹{neckline})'

                        detected_patterns.append({
                            'ticker': ticker,
                            'pattern_type': 'Fresh Reverse H&S',
                            'cmp': cmp_val,
                            'neckline_price': neckline,
                            'dist_to_breakout_pct': dist_pct,
                            'breakout_status': status,
                            'projected_target': target,
                            'anchor_dates': f"{dates[s1_idx]} (LS) / {dates[h_idx]} (H) / {dates[s2_idx]} (RS)",
                            'action_signal': signal
                        })
                        break

        # 3. INDEPENDENT CHECK: FRESH CUP WITH HANDLE
        if len(troughs) >= 2 and len(peaks) >= 2:
            for i in range(len(troughs) - 2, -1, -1):
                t_cup = troughs[i]
                if len(prices) - t_cup > 35:
                    continue
                rim_peaks = [p for p in peaks if p < t_cup]
                handle_peaks = [p for p in peaks if p > t_cup]
                if rim_peaks and handle_peaks:
                    p_rim = max(rim_peaks, key=lambda p: float(highs[p]))
                    p_handle = max(handle_peaks, key=lambda p: float(highs[p]))
                    rim_val = float(highs[p_rim])
                    handle_val = float(highs[p_handle])
                    if abs(rim_val - handle_val) / rim_val <= 0.06:
                        neckline = round(float((rim_val + handle_val) / 2.0), 2)
                        cup_depth = neckline - float(lows[t_cup])
                        target = round(neckline + cup_depth, 2)
                        dist_pct = round(((cmp_val - neckline) / neckline) * 100, 1)

                        if cmp_val >= neckline:
                            if dist_pct <= max_breakout_pct:
                                status = 'CONFIRMED BREAKOUT' if (is_green_today and is_green_prev) else 'INITIAL BREAKOUT'
                                signal = 'BUY MORNING (2nd Green)' if (is_green_today and is_green_prev) else 'WATCHLIST'
                                detected_patterns.append({
                                    'ticker': ticker,
                                    'pattern_type': 'Fresh Cup with Handle',
                                    'cmp': cmp_val,
                                    'neckline_price': neckline,
                                    'dist_to_breakout_pct': dist_pct,
                                    'breakout_status': status,
                                    'projected_target': target,
                                    'anchor_dates': f"{dates[p_rim]} (Rim1) / {dates[t_cup]} (Cup) / {dates[p_handle]} (Handle)",
                                    'action_signal': signal
                                })
                                break
                        elif dist_pct >= -6.0:
                            detected_patterns.append({
                                'ticker': ticker,
                                'pattern_type': 'Fresh Cup with Handle',
                                'cmp': cmp_val,
                                'neckline_price': neckline,
                                'dist_to_breakout_pct': dist_pct,
                                'breakout_status': 'FORMING HANDLE / NEAR RIM',
                                'projected_target': target,
                                'anchor_dates': f"{dates[p_rim]} (Rim1) / {dates[t_cup]} (Cup) / {dates[p_handle]} (Handle)",
                                'action_signal': f'WATCHLIST (Alert at ₹{neckline})'
                            })
                            break

        if detected_patterns:
            return detected_patterns
        else:
            return [default_no_pattern]

    except Exception as e:
        print(f"Notice: Error processing {ticker}: {e}")
        return [default_no_pattern]

def create_geometric_workbook(ticker_list, output_filename="v40_geometric_analysis-v5.xlsx"):
    """
    Generates a formatted Excel workbook containing short-term parallel pattern scan results 
    with separate rows per pattern and company group color fills.
    """
    wb = openpyxl.Workbook()
    
    font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_regular = Font(name="Calibri", size=11)
    
    fill_navy = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_blue_head = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    fill_section = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    
    # Soft Pastel Fills per Ticker Group (distinct colors for consecutive tickers)
    group_fills = [
        PatternFill(start_color="EBF1F5", end_color="EBF1F5", fill_type="solid"), # Light Slate Blue
        PatternFill(start_color="F9F2EC", end_color="F9F2EC", fill_type="solid"), # Soft Cream
        PatternFill(start_color="EBF5EC", end_color="EBF5EC", fill_type="solid"), # Soft Mint
        PatternFill(start_color="F5EBF5", end_color="F5EBF5", fill_type="solid"), # Soft Lavender
        PatternFill(start_color="F5F5EB", end_color="F5F5EB", fill_type="solid")  # Soft Gold
    ]

    fill_green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    font_green = Font(name="Calibri", size=11, bold=True, color="006100")

    fill_yellow = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    font_yellow = Font(name="Calibri", size=11, bold=True, color="9C6500")

    fill_red = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    font_red = Font(name="Calibri", size=11, bold=True, color="9C0006")

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    
    thin_side = Side(border_style="thin", color="D9D9D9")
    thick_bottom = Side(border_style="medium", color="1F4E78")
    
    border_data = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    border_header = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thick_bottom)

    ws = wb.active
    ws.title = "Parallel Pattern Radar"
    
    try:
        ws.views.sheetView.showGridLines = True
    except Exception:
        pass

    ws.merge_cells("A1:I1")
    t_cell = ws.cell(row=1, column=1, value="V40 PARALLEL GEOMETRIC PATTERN RADAR (GROUP-COLOR CODED ROWS)")
    t_cell.font = font_title
    t_cell.fill = fill_navy
    t_cell.alignment = align_center
    ws.row_dimensions.height = 35

    ws.merge_cells("A2:I2")
    sub_cell = ws.cell(row=2, column=1, value="Option 2: Individual Rows per Detected Pattern | Ticker Groups Color-Coded Together")
    sub_cell.font = Font(name="Calibri", size=10, italic=True, color="FFFFFF")
    sub_cell.fill = fill_blue_head
    sub_cell.alignment = align_center
    ws.row_dimensions.height = 20

    headers = [
        "Ticker / Company", "Pattern Type", "CMP (₹)", "Neckline / Rim (₹)", 
        "Distance to Breakout", "Breakout Status", "Projected Target (₹)", 
        "Key Anchor Dates", "Action Signal"
    ]

    ws.row_dimensions.height = 28
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_section
        cell.alignment = align_center
        cell.border = border_header

    current_row = 5
    for ticker_idx, ticker in enumerate(ticker_list):
        pattern_list = analyze_geometric_patterns(ticker)
        company_fill = group_fills[ticker_idx % len(group_fills)]

        for res in pattern_list:
            ws.row_dimensions[current_row].height = 22

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

            # Apply Company Group Color
            for col_i in range(1, 10):
                cell = ws.cell(row=current_row, column=col_i)
                cell.font = font_regular
                cell.border = border_data
                cell.fill = company_fill

            # Highlight Breakout Status
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

            # Highlight Action Signal
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

            current_row += 1

    col_widths = [18, 26, 14, 18, 20, 32, 18, 38, 32]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(output_filename)

