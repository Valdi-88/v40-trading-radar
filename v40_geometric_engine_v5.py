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

def analyze_geometric_patterns(ticker, lookback_days=120, max_breakout_pct=4.0):
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
        df = yf.Ticker(yf_ticker).history(period="1y")

        if df.empty or len(df) < 60:
            return [default_no_pattern]

        close_arr = safe_1d(df, 'Close')
        open_arr = safe_1d(df, 'Open')
        high_arr = safe_1d(df, 'High')
        low_arr = safe_1d(df, 'Low')

        if len(close_arr) < 2 or len(open_arr) < 2:
            return [default_no_pattern]

        cmp_val = round(float(close_arr[-1]), 2)
        default_no_pattern['cmp'] = cmp_val

        lifetime_high = float(np.max(high_arr))
        
        df_recent = df.tail(lookback_days).copy()
        prices = safe_1d(df_recent, 'Close')
        opens = safe_1d(df_recent, 'Open')
        lows = safe_1d(df_recent, 'Low')
        highs = safe_1d(df_recent, 'High')
        dates = [d.strftime('%Y-%m-%d') for d in df_recent.index]

        n_recent = len(prices)
        if n_recent < 20:
            return [default_no_pattern]

        is_green_today = prices[-1] > opens[-1]

        std_val = float(np.std(prices))
        prom = std_val * 0.30 if std_val > 0 else 1.0
        
        peaks, _ = find_peaks(prices, distance=6, prominence=prom)
        troughs, _ = find_peaks(-prices, distance=6, prominence=prom)

        peaks = [int(p) for p in peaks]
        troughs = [int(t) for t in troughs]

        detected_patterns = []

        # -------------------------------------------------------------
        # 1. W-PATTERN (NON-ADJACENT PAIRING, 2-4 MO DURATION, LEG BALANCE)
        # -------------------------------------------------------------
        if len(troughs) >= 2:
            for i in range(len(troughs) - 1, -1, -1):
                found_w = False
                for j in range(i + 1, min(i + 3, len(troughs))):
                    t1, t2 = troughs[i], troughs[j]
                    
                    # Recency check on L2: must be within last 30 trading days
                    if (n_recent - t2) > 30:
                        continue

                    # Timeframe Gate: 15 to 80 trading days (~2 to 4 months)
                    if not (15 <= (t2 - t1) <= 80):
                        continue

                    p1_val, p2_val = float(lows[t1]), float(lows[t2])
                    min_bottom = min(p1_val, p2_val)

                    # Strict floor alignment: L1 and L2 within 1.5% tolerance
                    if min_bottom > 0 and (abs(p1_val - p2_val) / min_bottom) <= 0.015:
                        between_peaks = [p for p in peaks if t1 < p < t2]
                        if between_peaks:
                            nk_idx = max(between_peaks, key=lambda p: float(highs[p]))
                            
                            # Leg Balance Gate: mid-point peak must be at least 5 trading days away from both L1 and L2
                            if (nk_idx - t1) < 5 or (t2 - nk_idx) < 5:
                                continue

                            neckline = round(float(highs[nk_idx]), 2)
                            
                            if neckline >= (0.92 * lifetime_high):
                                continue

                            pattern_depth_pct = ((neckline - min_bottom) / neckline) * 100
                            if not (4.5 <= pattern_depth_pct <= 15.0):
                                continue

                            # Identify the actual peak before t1 where the initial fall started
                            prior_peaks = [p for p in peaks if p < t1]
                            if prior_peaks:
                                p_start = max(prior_peaks, key=lambda p: float(highs[p]))
                            else:
                                p_start = int(np.argmax(highs[:t1])) if t1 > 0 else 0

                            fall_start_price = round(float(highs[p_start]), 2)
                            fall_start_date = dates[p_start]

                            # SOURCE RULE: Top of fall verification (must have fallen >=4% into L1)
                            if fall_start_price > 0 and ((fall_start_price - p1_val) / fall_start_price) < 0.04:
                                continue

                            # Preferred target from source: highest closing price from where original fall began
                            full_t1_idx = len(close_arr) - n_recent + t1
                            full_prior_prices = close_arr[:full_t1_idx]
                            if len(full_prior_prices) > 0:
                                prior_highest_close = round(float(np.max(full_prior_prices)), 2)
                                target = max(prior_highest_close, round(neckline + (neckline - min_bottom), 2))
                            else:
                                target = round(neckline + (neckline - min_bottom), 2)

                            dist_pct = round(((cmp_val - neckline) / neckline) * 100, 1)

                            # Floor invalidation gate
                            if cmp_val < min_bottom:
                                continue  

                            # Search for breakout candle starting from t2
                            breakout_idx = None
                            for idx_c in range(t2, n_recent):
                                if prices[idx_c] >= neckline and prices[idx_c] > opens[idx_c]:
                                    breakout_idx = idx_c
                                    break

                            # Fresh breakout recency gate (must be within last 2 days)
                            if breakout_idx is not None and breakout_idx < (n_recent - 2):
                                continue  

                            t1_date = dates[t1]
                            nk_date = dates[nk_idx]
                            t2_date = dates[t2]
                            end_date = dates[-1]

                            date_span_str = f"Fall Start: {fall_start_date} (₹{fall_start_price}) ➔ L1: {t1_date} ➔ Peak: {nk_date} ➔ L2: {t2_date} | Span: {fall_start_date} ➔ {end_date}"

                            if breakout_idx == (n_recent - 1):
                                if dist_pct > max_breakout_pct:
                                    continue
                                status = 'INITIAL BREAKOUT (1st Green Candle Today)'
                                signal = 'WATCHLIST (Wait for 2nd Green Candle)'
                            elif breakout_idx == (n_recent - 2):
                                breakout_high = highs[breakout_idx]
                                if cmp_val > breakout_high and is_green_today:
                                    status = 'CONFIRMED BREAKOUT (Closed Above Breakout High Today)'
                                    signal = 'BUY MORNING (2-Step Confirmed)'
                                else:
                                    continue
                            else:
                                if dist_pct >= -5.0:
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
                                'anchor_dates': date_span_str,
                                'action_signal': signal
                            })
                            found_w = True
                            break
                if found_w:
                    break

        # -------------------------------------------------------------
        # 2. REVERSE HEAD & SHOULDERS (STANDARD & COMPLEX)
        # -------------------------------------------------------------
        if len(troughs) >= 3:
            for num_tr in range(min(5, len(troughs)), 2, -1):
                found_rhs = False
                for i in range(len(troughs) - num_tr, -1, -1):
                    sub_troughs = troughs[i:i+num_tr]
                    if (n_recent - sub_troughs[-1]) > 30:
                        continue

                    h_local_idx = min(range(len(sub_troughs)), key=lambda k: float(lows[sub_troughs[k]]))
                    if h_local_idx == 0 or h_local_idx == len(sub_troughs) - 1:
                        continue

                    h_idx = sub_troughs[h_local_idx]
                    h_p = float(lows[h_idx])

                    left_shoulders = sub_troughs[:h_local_idx]
                    right_shoulders = sub_troughs[h_local_idx+1:]

                    all_shoulders_p = [float(lows[s]) for s in left_shoulders + right_shoulders]
                    min_shoulder_p = min(all_shoulders_p)

                    if h_p < (min_shoulder_p * 0.96):
                        if (max(all_shoulders_p) - min(all_shoulders_p)) / min_shoulder_p <= 0.055:
                            shoulder_peaks = []
                            for idx_s in range(len(sub_troughs) - 1):
                                t_a, t_b = sub_troughs[idx_s], sub_troughs[idx_s+1]
                                between_p = [p for p in peaks if t_a < p < t_b]
                                if between_p:
                                    peak_idx = max(between_p, key=lambda p: float(highs[p]))
                                    shoulder_peaks.append(peak_idx)

                            if len(shoulder_peaks) >= 2:
                                peak_vals = [float(highs[p]) for p in shoulder_peaks]
                                min_nk, max_nk = min(peak_vals), max(peak_vals)

                                if (max_nk - min_nk) / min_nk <= 0.020:
                                    neckline = round(float(np.mean(peak_vals)), 2)

                                    if neckline >= (0.92 * lifetime_high):
                                        continue

                                    height = neckline - h_p
                                    target = round(neckline + height, 2)
                                    dist_pct = round(((cmp_val - neckline) / neckline) * 100, 1)

                                    if cmp_val < h_p:
                                        continue

                                    breakout_idx = None
                                    for idx_c in range(sub_troughs[-1], n_recent):
                                        if prices[idx_c] >= neckline and prices[idx_c] > opens[idx_c]:
                                            breakout_idx = idx_c
                                            break

                                    if breakout_idx is not None and breakout_idx < (n_recent - 2):
                                        continue  

                                    is_complex = len(left_shoulders) > 1 or len(right_shoulders) > 1
                                    pattern_label = 'Fresh Complex Reverse H&S (Multi-Shoulder)' if is_complex else 'Fresh Reverse H&S'

                                    ls_first_idx = left_shoulders
                                    ls_start = dates[max(0, ls_first_idx - 7)]
                                    head_date = dates[h_idx]
                                    nk1_date = dates[shoulder_peaks]
                                    nk2_date = dates[shoulder_peaks[-1]]
                                    rs_end = dates[-1]

                                    if is_complex:
                                        date_span_str = f"LS ({len(left_shoulders)}): {ls_start}➔{nk1_date} | Head: {nk1_date}➔{head_date} | RS ({len(right_shoulders)}): {head_date}➔{rs_end}"
                                    else:
                                        date_span_str = f"LS: {ls_start}➔{nk1_date} | Head: {nk1_date}➔{nk2_date} | RS: {nk2_date}➔{rs_end}"

                                    if breakout_idx == (n_recent - 1):
                                        if dist_pct > max_breakout_pct:
                                            continue
                                        status = 'INITIAL BREAKOUT (1st Green Candle Today)'
                                        signal = 'WATCHLIST (Wait for 2nd Green Candle)'
                                    elif breakout_idx == (n_recent - 2):
                                        breakout_high = highs[breakout_idx]
                                        if cmp_val > breakout_high and is_green_today:
                                            status = 'CONFIRMED BREAKOUT (Closed Above Breakout High Today)'
                                            signal = 'BUY MORNING (2-Step Confirmed)'
                                        else:
                                            continue
                                    else:
                                        if dist_pct >= -5.0:
                                            status = 'APPROACHING BREAKOUT'
                                            signal = f'WATCHLIST (Alert at ₹{neckline})'
                                        else:
                                            status = 'FORMING RIGHT SHOULDER'
                                            signal = 'WATCHLIST ONLY'

                                    detected_patterns.append({
                                        'ticker': ticker,
                                        'pattern_type': pattern_label,
                                        'cmp': cmp_val,
                                        'neckline_price': neckline,
                                        'dist_to_breakout_pct': dist_pct,
                                        'breakout_status': status,
                                        'projected_target': target,
                                        'anchor_dates': date_span_str,
                                        'action_signal': signal
                                    })
                                    found_rhs = True
                                    break
                if found_rhs:
                    break

        # -------------------------------------------------------------
        # 3. CUP WITH HANDLE (SOURCE GROUNDED: TOP OF FALL & HANDLE BOTTOM)
        # -------------------------------------------------------------
        if len(troughs) >= 2 and len(peaks) >= 2:
            for i in range(len(troughs) - 2, -1, -1):
                t_cup = troughs[i]
                if n_recent - t_cup > 60:
                    continue
                rim_peaks = [p for p in peaks if p < t_cup]
                handle_peaks = [p for p in peaks if p > t_cup]
                if rim_peaks and handle_peaks:
                    p_rim = max(rim_peaks, key=lambda p: float(highs[p]))
                    p_handle = max(handle_peaks, key=lambda p: float(highs[p]))
                    
                    # SOURCE RULE 1: Top of Fall Verification
                    # Ensure p_rim was preceded by an upward rally (not starting midway through a fall)
                    prior_troughs_before_rim = [t for t in troughs if t < p_rim]
                    if prior_troughs_before_rim:
                        t_before_rim = max(prior_troughs_before_rim)
                        pre_rally_pct = (highs[p_rim] - lows[t_before_rim]) / lows[t_before_rim]
                        if pre_rally_pct < 0.04:  # Must have rallied >=4% into p_rim
                            continue  # Rejects Cup starting from "Mid of Fall"
                    
                    handle_troughs = [t for t in troughs if t > t_cup]
                    is_double_handle = len(handle_troughs) >= 2

                    rim_val = float(highs[p_rim])
                    handle_val = float(highs[p_handle])

                    if abs(rim_val - handle_val) / min(rim_val, handle_val) <= 0.020:
                        neckline = round(float((rim_val + handle_val) / 2.0), 2)
                        
                        if neckline >= (0.92 * lifetime_high):
                            continue

                        cup_depth = neckline - float(lows[t_cup])
                        cup_depth_pct = (cup_depth / neckline) * 100

                        if 6.0 <= cup_depth_pct <= 30.0:
                            # SOURCE RULE 2: Handle Bottom Breakdown Invalidation
                            if handle_troughs:
                                min_handle_bottom = min([float(lows[ht]) for ht in handle_troughs])
                                if cmp_val < min_handle_bottom:
                                    continue  # Pattern invalid if price closes below handle bottom

                            target = round(neckline + cup_depth, 2)
                            dist_pct = round(((cmp_val - neckline) / neckline) * 100, 1)

                            if cmp_val < float(lows[t_cup]):
                                continue

                            breakout_idx = None
                            for idx_c in range(t_cup, n_recent):
                                if prices[idx_c] >= neckline and prices[idx_c] > opens[idx_c]:
                                    breakout_idx = idx_c
                                    break

                            if breakout_idx is not None and breakout_idx < (n_recent - 2):
                                continue  # Past breakout -> Skip

                            pattern_label = 'Fresh Complex Cup with Handle (Double Handle)' if is_double_handle else 'Fresh Cup with Handle'
                            cup_start = dates[p_rim]
                            cup_end = dates[p_handle]
                            
                            if is_double_handle:
                                h1_end = dates[handle_troughs]
                                h2_end = dates[-1]
                                date_span_str = f"Cup: {cup_start}➔{cup_end} | Handle 1: {cup_end}➔{h1_end} | Handle 2: {h1_end}➔{h2_end}"
                            else:
                                handle_end = dates[-1]
                                date_span_str = f"Cup: {cup_start} ➔ {cup_end} | Handle: {cup_end} ➔ {handle_end}"

                            if breakout_idx == (n_recent - 1):
                                if dist_pct > max_breakout_pct:
                                    continue
                                status = 'INITIAL BREAKOUT (1st Green Candle Today)'
                                signal = 'WATCHLIST (Wait for 2nd Green Candle)'
                            elif breakout_idx == (n_recent - 2):
                                breakout_high = highs[breakout_idx]
                                if cmp_val > breakout_high and is_green_today:
                                    status = 'CONFIRMED BREAKOUT (Closed Above Breakout High Today)'
                                    signal = 'BUY MORNING (2-Step Confirmed)'
                                else:
                                    continue
                            else:
                                status = 'FORMING HANDLE / NEAR RESISTANCE'
                                signal = f'WATCHLIST (Alert at ₹{neckline})'

                            detected_patterns.append({
                                'ticker': ticker,
                                'pattern_type': pattern_label,
                                'cmp': cmp_val,
                                'neckline_price': neckline,
                                'dist_to_breakout_pct': dist_pct,
                                'breakout_status': status,
                                'projected_target': target,
                                'anchor_dates': date_span_str,
                                'action_signal': signal
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
    wb = openpyxl.Workbook()
    
    font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_regular = Font(name="Calibri", size=11)
    
    fill_navy = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_blue_head = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
    fill_section = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    
    group_fills = [
        PatternFill(start_color="EBF1F5", end_color="EBF1F5", fill_type="solid"),
        PatternFill(start_color="F9F2EC", end_color="F9F2EC", fill_type="solid"),
        PatternFill(start_color="EBF5EC", end_color="EBF5EC", fill_type="solid"),
        PatternFill(start_color="F5EBF5", end_color="F5EBF5", fill_type="solid"),
        PatternFill(start_color="F5F5EB", end_color="F5F5EB", fill_type="solid")
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
    t_cell = ws.cell(row=1, column=1, value="V40 PARALLEL GEOMETRIC PATTERN RADAR (STANDARD & COMPLEX PATTERNS)")
    t_cell.font = font_title
    t_cell.fill = fill_navy
    t_cell.alignment = align_center
    ws.row_dimensions.height = 35

    ws.merge_cells("A2:I2")
    sub_cell = ws.cell(row=2, column=1, value="Detects Complex Multi-Shoulders (Rev H&S) & Double Handles (Cup & Handle) with Exact Date Spans")
    sub_cell.font = Font(name="Calibri", size=10, italic=True, color="FFFFFF")
    sub_cell.fill = fill_blue_head
    sub_cell.alignment = align_center
    ws.row_dimensions.height = 20

    headers = [
        "Ticker / Company", "Pattern Type", "CMP (₹)", "Neckline / Resistance (₹)", 
        "Distance to Breakout", "Breakout Status", "Projected Target (₹)", 
        "Pattern Date Spans (Start ➔ End)", "Action Signal"
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
            c8.alignment = align_left
            c9.alignment = align_left

            for col_i in range(1, 10):
                cell = ws.cell(row=current_row, column=col_i)
                cell.font = font_regular
                cell.border = border_data
                cell.fill = company_fill

            status_str = str(res['breakout_status']).upper()
            if 'CONFIRMED BREAKOUT' in status_str or 'CLOSED ABOVE BREAKOUT HIGH' in status_str:
                c6.fill = fill_green
                c6.font = font_green
            elif 'FAILED BREAKOUT' in status_str or 'SLIPPED' in status_str:
                c6.fill = fill_red
                c6.font = font_red
            elif 'PULLBACK' in status_str or 'APPROACHING' in status_str or 'FORMING' in status_str or 'INITIAL' in status_str:
                c6.fill = fill_yellow
                c6.font = font_yellow

            signal_str = str(res['action_signal']).upper()
            if 'BUY MORNING' in signal_str or 'CONFIRMED' in signal_str:
                c9.fill = fill_green
                c9.font = font_green
            elif 'AVOID' in signal_str or 'FAILED' in signal_str:
                c9.fill = fill_red
                c9.font = font_red
            elif 'WATCHLIST' in signal_str or 'WAIT' in signal_str:
                c9.fill = fill_yellow
                c9.font = font_yellow

            current_row += 1

    col_widths = [4-9]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(output_filename)
