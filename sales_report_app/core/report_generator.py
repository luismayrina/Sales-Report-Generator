import os
from datetime import datetime
from collections import defaultdict
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

from .workbook_utils import copy_style
from .sku_normalizer import normalize_sku, extract_valid_skus_from_template, match_sku

# ── SHEET 1: DAILY SALES SUMMARY ─────────────────────────────────────────────
def write_daily_sales(ws_out, ws_sample, all_orders):
    header_fill = PatternFill(start_color="F0F0F0", end_color="F0F0F0", fill_type="solid")
    alt_fill = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")

    # Copy header row
    for j, cell in enumerate(ws_sample[1], 1):
        dst = ws_out.cell(row=1, column=j, value=cell.value)
        copy_style(cell, dst)
        dst.fill = header_fill
        dst.font = Font(bold=True)

    # Column widths
    for col_letter, dim in ws_sample.column_dimensions.items():
        ws_out.column_dimensions[col_letter].width = dim.width

    DATE_FMT  = "d-mmm"
    DATE2_FMT = "mm-dd-yy"
    MONEY_FMT = "#,##0.00"

    # Sort by date, then source
    source_order = {"Shopify": 0, "Shopee": 1, "Lazada": 2}
    sorted_orders = sorted(
        all_orders,
        key=lambda o: (o.get("date") or datetime(2000,1,1), source_order.get(o["source"], 9))
    )

    for ri, o in enumerate(sorted_orders, start=2):
        vals = [
            o["source"],           # A  au
            o.get("date"),         # B  Date
            o["customer"],         # C  Customer Name
            o["order_no"],         # D  Order No.
            o.get("gross"),        # E  Gross Amount
            o.get("shipping"),     # F  Plus Shipping fee
            o.get("payment_fee"),  # G  Less payment fee
            None,                  # H  Less shipping fee
            o.get("discount"),     # I  Less Discount
            None,None,None,None,None,None,None,None,  # J-Q
            o.get("paymongo_fee"), # R  Less Paymongo fee
            None,                  # S  EWT
            o.get("net"),          # T  Net Amount
            o.get("deposit_date"), # U  Payout/Date Deposit
            o.get("bank"),         # V  BANK
            None,                  # W  SI
            o.get("deposit_date"), # X  Date Deposit
            o.get("ref"),          # Y  Reference No.
            None,None,None,None,   # Z-AC
        ]
        for ci, v in enumerate(vals, 1):
            cell = ws_out.cell(row=ri, column=ci, value=v)
            if ri % 2 == 0:
                cell.fill = alt_fill
            if ci == 2 and v:
                cell.number_format = DATE_FMT
            elif ci == 21 and v:
                cell.number_format = DATE_FMT
            elif ci == 24 and v:
                cell.number_format = DATE2_FMT
            elif ci in (5, 6, 7, 9, 18, 20) and v is not None:
                cell.number_format = MONEY_FMT

    ws_out.freeze_panes = "B2"


# ── SHEET 2: REAL TIME INVENTORY ─────────────────────────────────────────────
def write_inventory(ws_out, ws_sample):
    for i, row in enumerate(ws_sample.iter_rows(), 1):
        for j, cell in enumerate(row, 1):
            dst = ws_out.cell(row=i, column=j, value=cell.value)
            copy_style(cell, dst)
    
    ws_out.cell(row=3, column=1, value="[Pending Data Source for Finished Goods + Raw Materials]")


# ── SHEET 3: OFFTAKE REPORT SUMMARY ──────────────────────────────────────────
MONTH_NAMES = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]
SECTION_WIDTH = 10  # Each month block is 10 columns wide
# Block layout (1-indexed offsets from block start):
#  +0 : % cont (monthly)
#  +1 : Actual  ← this col holds the datetime header
#  +2 : Target
#  +3 : PY
#  +4 : vs PY
#  +5 : (gap)
#  +6 : % cont (YTD)
#  +7 : YTD actual
#  +8 : YTD PY
#  +9 : YTD vs PY
# First block starts at col 2 (B).

def _build_section(month_num, section_start):
    """Return a dict mapping semantic names -> 1-indexed column numbers."""
    mn = month_num
    s = section_start
    return {
        "month_num":    mn,
        "section_start": s,
        "pct_cont_col": s,
        "actual_col":   s + 1,
        "target_col":   s + 2,
        "py_col":       s + 3,
        "vs_py_col":    s + 4,
        "ytd_pct_col":  s + 6,
        "ytd_col":      s + 7,
        "ytd_py_col":   s + 8,
        "ytd_vs_py_col": s + 9,
    }


def _copy_block_styles(ws, src_start, dst_start, max_row, log):
    """
    Copy styles from every cell in a 10-wide source block to a destination block.
    Used when creating new month columns not present in the template.
    """
    for row in range(1, max_row + 1):
        for offset in range(SECTION_WIDTH):
            src_cell = ws.cell(row=row, column=src_start + offset)
            dst_cell = ws.cell(row=row, column=dst_start + offset)
            copy_style(src_cell, dst_cell)
            # Clear any stale values in the new block
            dst_cell.value = None


def write_offtake_summary(ws_out, ws_sample, all_orders, py_data, latest_sales_month, current_year=2026, logger=None):
    def log(msg):
        if logger:
            logger(msg)

    SKIP_ROWS = {"TOTAL WITH INDUSTRIAL", "INDUSTRIAL", "RETAIL", "ECOMMERCE", "TOTAL"}

    # ── Step 1: Copy the entire template sheet ──────────────────────────────
    for i, row in enumerate(ws_sample.iter_rows(), 1):
        for j, cell in enumerate(row, 1):
            dst = ws_out.cell(row=i, column=j, value=cell.value)
            copy_style(cell, dst)

    # ── Step 2: Detect existing month sections from template header ─────────
    # A section is identified by a datetime value in row 1 (the "actual" col).
    template_month_sections = {}   # month_num -> section dict
    hdr_len = ws_out.max_column
    hdr_row = 1
    for ci in range(1, hdr_len + 2):
        val = ws_out.cell(1, ci).value
        if isinstance(val, datetime):
            month_num = val.month
            section_start = ci - 1   # % cont col is one before the datetime col
            template_month_sections[month_num] = _build_section(month_num, section_start)
            log(f"Template has month {val.strftime('%B')} (block start col {section_start})")

    if not template_month_sections:
        log("WARNING: No month columns found in template. Aborting Offtake Summary generation.")
        return

    # ── Step 5: Detect channel rows (moved up to use in month detection) ────
    channel_row_map = {}
    total_row = None
    for i in range(1, 35):
        v = ws_out.cell(i, 1).value
        if v is None:
            continue
        name = str(v).strip()
        if name:
            channel_row_map[name] = i
            log(f"Channel row: '{name}' -> row {i}")
        nu = name.upper()
        if nu == "TOTAL WITH INDUSTRIAL":
            total_row = i

    # ── Step 3: Determine all months needed from sales data ─────────────────
    # Determine the contiguous range Jan→latest month.
    earliest_template_month = min(template_month_sections.keys())
    log(f"Latest sales month detected: {MONTH_NAMES[latest_sales_month - 1]} (month {latest_sales_month})")

    # Months we need to have blocks for (from earliest in template to latest in data)
    needed_months = list(range(earliest_template_month, latest_sales_month + 1))
    log(f"Month blocks needed: {[MONTH_NAMES[m-1] for m in needed_months]}")

    # ── Step 4: Create missing month blocks ─────────────────────────────────
    last_template_month = max(template_month_sections.keys())
    last_section = template_month_sections[last_template_month]
    last_block_start = last_section["section_start"]
    next_block_start = last_block_start + SECTION_WIDTH

    all_sections = dict(template_month_sections)

    max_data_row = max(30, ws_out.max_row)
    for month_num in needed_months:
        if month_num in all_sections:
            continue
        log(f"Creating new month block for {MONTH_NAMES[month_num-1]} at col {next_block_start}")
        _copy_block_styles(ws_out, last_block_start, next_block_start, max_data_row, log)

        # Write row-1 header cells for the new block
        month_abbr = MONTH_NAMES[month_num - 1]
        insert_col = next_block_start + 1

        sec = _build_section(month_num, next_block_start)
        # For new blocks, insert the month label exactly as formatted in template.
        month_dt = datetime(current_year, month_num, 1)
        ws_out.cell(row=hdr_row, column=insert_col).value = month_dt
        ws_out.cell(row=hdr_row, column=insert_col).number_format = "MMM-yy"
        ws_out.cell(row=1, column=sec["target_col"],      value="Target")
        ws_out.cell(row=1, column=sec["vs_py_col"],       value="vs PY")
        ws_out.cell(row=1, column=sec["ytd_pct_col"],     value="% cont")
        ws_out.cell(row=1, column=sec["ytd_col"],         value=f"{month_abbr} YTD")
        ws_out.cell(row=1, column=sec["ytd_py_col"],      value="PY")
        ws_out.cell(row=1, column=sec["ytd_vs_py_col"],   value="vs PY")

        # Apply % and money number formats to data rows
        for ri in range(2, max_data_row + 1):
            ws_out.cell(ri, sec["pct_cont_col"]).number_format  = "0%"
            ws_out.cell(ri, sec["actual_col"]).number_format    = "#,##0.00"
            ws_out.cell(ri, sec["target_col"]).number_format    = "#,##0.00"
            ws_out.cell(ri, sec["ytd_pct_col"]).number_format   = "0%"
            ws_out.cell(ri, sec["ytd_col"]).number_format       = "#,##0.00"

        all_sections[month_num] = sec
        last_block_start = next_block_start
        next_block_start += SECTION_WIDTH

    # ── Step 5.5: Clear ALL app-calculated cells ──────────────────────────
    APP_CALC_OFFSETS = [0, 1, 4, 6, 7, 9]
    for ch_name, ridx in channel_row_map.items():
        if ch_name.upper() in SKIP_ROWS:
            continue
        for sec in all_sections.values():
            for offset in APP_CALC_OFFSETS:
                col = sec["section_start"] + offset
                ws_out.cell(row=ridx, column=col).value = None

    # ── Step 6: Calculate monthly totals per channel ────────────────────────
    channel_monthly = defaultdict(lambda: defaultdict(float))
    for o in all_orders:
        if o["status"] == "cancelled": continue
        src = o.get("source")
        if src not in channel_row_map: continue
        dt = o.get("date")
        if not dt: continue
        channel_monthly[src][dt.month] += o.get("net") or 0

    # ── Step 7: Write actuals, YTD, vs PY for all channels ──────────────────
    sorted_months = sorted(all_sections.keys())

    for ch, ridx in channel_row_map.items():
        if ch.upper() in SKIP_ROWS: continue
        ytd_sum = 0.0
        for month_num in sorted_months:
            sec = all_sections[month_num]
            monthly_val = channel_monthly[ch].get(month_num, 0.0)
            ytd_sum += monthly_val

            cell_actual = ws_out.cell(row=ridx, column=sec["actual_col"], value=round(monthly_val, 2) if monthly_val else 0)
            cell_actual.number_format = "#,##0.00"
            cell_ytd = ws_out.cell(row=ridx, column=sec["ytd_col"], value=round(ytd_sum, 2) if ytd_sum else 0)
            cell_ytd.number_format = "#,##0.00"

            py_val = 0.0
            if py_data and ch in py_data:
                py_val = py_data[ch].get(month_num, 0.0)
            else:
                py_val_sample = ws_sample.cell(ridx, sec["py_col"]).value if month_num in template_month_sections else None
                if isinstance(py_val_sample, (int, float)): py_val = py_val_sample

            cell_py = ws_out.cell(row=ridx, column=sec["py_col"], value=round(py_val, 2) if py_val else 0)
            cell_py.number_format = "#,##0.00"

            if py_val:
                pct_val = (monthly_val - py_val) / py_val
                cell_vs_py = ws_out.cell(row=ridx, column=sec["vs_py_col"], value=pct_val)
                cell_vs_py.number_format = "0%"
                if pct_val > 0:
                    cell_vs_py.fill = PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
                elif pct_val < 0:
                    cell_vs_py.fill = PatternFill(start_color="FCE8E6", end_color="FCE8E6", fill_type="solid")
            else: ws_out.cell(row=ridx, column=sec["vs_py_col"]).value = None

            ytd_py_sum = sum(py_data[ch].get(m, 0.0) for m in range(1, month_num + 1)) if py_data and ch in py_data else 0.0
            cell_ytd_py = ws_out.cell(row=ridx, column=sec["ytd_py_col"], value=round(ytd_py_sum, 2) if ytd_py_sum else 0)
            cell_ytd_py.number_format = "#,##0.00"

            if ytd_py_sum:
                ytd_pct_val = (ytd_sum - ytd_py_sum) / ytd_py_sum
                cell_ytd_vs_py = ws_out.cell(row=ridx, column=sec["ytd_vs_py_col"], value=ytd_pct_val)
                cell_ytd_vs_py.number_format = "0%"
                if ytd_pct_val > 0:
                    cell_ytd_vs_py.fill = PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
                elif ytd_pct_val < 0:
                    cell_ytd_vs_py.fill = PatternFill(start_color="FCE8E6", end_color="FCE8E6", fill_type="solid")
            else: ws_out.cell(row=ridx, column=sec["ytd_vs_py_col"]).value = None

    # ── Step 8: Recalculate TOTAL row ────────────────────────────────────────
    total_rows = [ridx for name, ridx in channel_row_map.items() if "TOTAL" in name.upper()]
    for t_ridx in total_rows:
        for month_num in sorted_months:
            sec = all_sections[month_num]
            total_actual = sum(ws_out.cell(ridx, sec["actual_col"]).value or 0 for name, ridx in channel_row_map.items() if name.upper() not in SKIP_ROWS)
            total_ytd = sum(ws_out.cell(ridx, sec["ytd_col"]).value or 0 for name, ridx in channel_row_map.items() if name.upper() not in SKIP_ROWS)
            cell_t_act = ws_out.cell(row=t_ridx, column=sec["actual_col"], value=round(total_actual, 2))
            cell_t_act.number_format = "#,##0.00"
            cell_t_ytd = ws_out.cell(row=t_ridx, column=sec["ytd_col"], value=round(total_ytd, 2))
            cell_t_ytd.number_format = "#,##0.00"
            
            # Recalculate totals for vs PY using existing PY sums
            total_py = sum(ws_out.cell(ridx, sec["py_col"]).value or 0 for name, ridx in channel_row_map.items() if name.upper() not in SKIP_ROWS)
            total_ytd_py = sum(ws_out.cell(ridx, sec["ytd_py_col"]).value or 0 for name, ridx in channel_row_map.items() if name.upper() not in SKIP_ROWS)
            
            if total_py:
                t_pct_val = (total_actual - total_py) / total_py
                cell_vs_py = ws_out.cell(row=t_ridx, column=sec["vs_py_col"], value=t_pct_val)
                cell_vs_py.number_format = "0%"
                if t_pct_val > 0:
                    cell_vs_py.fill = PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
                elif t_pct_val < 0:
                    cell_vs_py.fill = PatternFill(start_color="FCE8E6", end_color="FCE8E6", fill_type="solid")
            else:
                ws_out.cell(row=t_ridx, column=sec["vs_py_col"]).value = None
                
            if total_ytd_py:
                t_ytd_pct_val = (total_ytd - total_ytd_py) / total_ytd_py
                cell_ytd_vs_py = ws_out.cell(row=t_ridx, column=sec["ytd_vs_py_col"], value=t_ytd_pct_val)
                cell_ytd_vs_py.number_format = "0%"
                if t_ytd_pct_val > 0:
                    cell_ytd_vs_py.fill = PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
                elif t_ytd_pct_val < 0:
                    cell_ytd_vs_py.fill = PatternFill(start_color="FCE8E6", end_color="FCE8E6", fill_type="solid")
            else:
                ws_out.cell(row=t_ridx, column=sec["ytd_vs_py_col"]).value = None

    # ── Step 9: Recalculate % cont for all channels ──────────────────────────
    # Find main total row for division
    main_total_row = total_row
    if not main_total_row and total_rows:
        main_total_row = total_rows[0]

    for month_num in sorted_months:
        sec = all_sections[month_num]
        total_actual = ws_out.cell(main_total_row, sec["actual_col"]).value if main_total_row else None
        total_ytd    = ws_out.cell(main_total_row, sec["ytd_col"]).value    if main_total_row else None

        for ch_name, ridx in channel_row_map.items():
            if ch_name.upper() in SKIP_ROWS:
                continue
            row_actual = ws_out.cell(ridx, sec["actual_col"]).value
            row_ytd    = ws_out.cell(ridx, sec["ytd_col"]).value

            if isinstance(row_actual, (int, float)) and isinstance(total_actual, (int, float)) and total_actual:
                pct = round(row_actual / total_actual, 4)
                ws_out.cell(row=ridx, column=sec["pct_cont_col"], value=pct)
                ws_out.cell(row=ridx, column=sec["pct_cont_col"]).number_format = "0%"
            else:
                ws_out.cell(row=ridx, column=sec["pct_cont_col"]).value = None

            if isinstance(row_ytd, (int, float)) and isinstance(total_ytd, (int, float)) and total_ytd:
                ytd_pct = round(row_ytd / total_ytd, 4)
                ws_out.cell(row=ridx, column=sec["ytd_pct_col"], value=ytd_pct)
                ws_out.cell(row=ridx, column=sec["ytd_pct_col"]).number_format = "0%"
            else:
                ws_out.cell(row=ridx, column=sec["ytd_pct_col"]).value = None

    # ── Step 9.5: Add Zebra Striping ────────────────────────────────────────────
    alt_fill = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")
    data_rows = sorted([ridx for name, ridx in channel_row_map.items() if name.upper() not in SKIP_ROWS])
    for i, ridx in enumerate(data_rows):
        if i % 2 == 1:
            for sec in all_sections.values():
                for col in range(sec["section_start"], sec["section_start"] + SECTION_WIDTH):
                    if col not in (sec["vs_py_col"], sec["ytd_vs_py_col"]):
                        ws_out.cell(row=ridx, column=col).fill = alt_fill
            # Also stripe the channel name column
            ws_out.cell(row=ridx, column=1).fill = alt_fill

    # ── Step 10: Auto-set column widths to prevent ##### display ─────────────
    # Currency columns (actual, target, YTD) → min 16; % cont → min 10
    for sec in all_sections.values():
        currency_cols = [sec["actual_col"], sec["target_col"], sec["ytd_col"], sec["py_col"], sec["ytd_py_col"]]
        pct_cols = [sec["pct_cont_col"], sec["ytd_pct_col"]]
        for col in currency_cols:
            letter = ws_out.cell(1, col).column_letter
            cur = ws_out.column_dimensions[letter].width or 0
            ws_out.column_dimensions[letter].width = max(cur, 16)
        for col in pct_cols:
            letter = ws_out.cell(1, col).column_letter
            cur = ws_out.column_dimensions[letter].width or 0
            ws_out.column_dimensions[letter].width = max(cur, 10)

    # Freeze column A and Row 1
    ws_out.freeze_panes = "B2"
    
    # ── Step 11: Cleanup leftover template errors ────────────────────────────
    for r in range(1, ws_out.max_row + 1):
        for c in range(1, ws_out.max_column + 1):
            val = ws_out.cell(r, c).value
            if isinstance(val, str) and val in ("#DIV/0!", "#N/A", "#VALUE!", "#REF!"):
                ws_out.cell(r, c).value = None

# ── SHEET 4: OFFTAKE PER SKU ─────────────────────────────────────────────────
def write_offtake_per_sku(ws_out, ws_sample, all_orders, valid_skus):
    for i, row in enumerate(ws_sample.iter_rows(values_only=True), 1):
        for j, val in enumerate(row, 1):
            ws_out.cell(row=i, column=j, value=val)

    months = ["JAN","FEB","MAR","APR","MAY","JUN","JUL","AUG","SEP","OCT","NOV","DEC"]
    month_num = {m: i+1 for i, m in enumerate(months)}
    hdr_row = None
    for i, row in enumerate(ws_sample.iter_rows(values_only=True), 1):
        if row[0] == "VARIANT":
            hdr_row = i
            break
    if not hdr_row: return

    hdr = [c.value for c in ws_sample[hdr_row]]
    month_col = {}
    for ci, val in enumerate(hdr, 1):
        if val in months: month_col[month_num[val]] = ci

    sku_row = {}
    for i, row in enumerate(ws_sample.iter_rows(min_row=hdr_row+1, values_only=True), hdr_row+1):
        if row[0]: sku_row[str(row[0]).strip()] = i

    sku_qty = defaultdict(lambda: defaultdict(int))
    for o in all_orders:
        if o["status"] == "cancelled": continue
        dt = o.get("date")
        if not dt: continue
        month = dt.month
        
        for (item_name, qty, price) in o.get("items", []):
            normalized = normalize_sku(item_name)
            matched_sku = match_sku(normalized, valid_skus) or normalized
            sku_qty[matched_sku][month] += qty

    last_row_idx = max(sku_row.values()) if sku_row else hdr_row + 1

    for sku, month_data in sku_qty.items():
        row_idx = sku_row.get(sku)
        if not row_idx:
            last_row_idx += 1
            row_idx = last_row_idx
            sku_row[sku] = row_idx
            ws_out.cell(row=row_idx, column=1, value=sku)
            if last_row_idx > hdr_row + 1:
                for col in range(1, ws_out.max_column + 1):
                    src_cell = ws_out.cell(row=last_row_idx - 1, column=col)
                    dst_cell = ws_out.cell(row=row_idx, column=col)
                    copy_style(src_cell, dst_cell)
                    
        for month, qty in month_data.items():
            col_idx = month_col.get(month)
            if col_idx:
                existing = ws_out.cell(row=row_idx, column=col_idx).value
                ws_out.cell(row=row_idx, column=col_idx, value=(existing or 0) + qty)

    ws_out.freeze_panes = "B2"
    alt_fill = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")
    for row_idx in range(hdr_row + 1, last_row_idx + 1):
        if row_idx % 2 == 0:
            for col in range(1, ws_out.max_column + 1):
                ws_out.cell(row=row_idx, column=col).fill = alt_fill


# ── MAIN GENERATOR ───────────────────────────────────────────────────────────
def generate_full_report(all_orders, template_path, output_path, py_data=None, current_year=2026, logger=print):
    def log(msg):
        if logger: logger(msg)

    log("Loading template workbook...")
    wb_sample = load_workbook(template_path)
    
    valid_skus = extract_valid_skus_from_template(wb_sample["Offtake Report Summary per sku"])

    wb_out = openpyxl.Workbook()

    log("Generating Daily Sales Summary...")
    ws1 = wb_out.active
    ws1.title = "Daily Sales Summary"
    write_daily_sales(ws1, wb_sample["Daily Sales Summary "], all_orders)

    log("Generating Real Time Inventory Report...")
    ws2 = wb_out.create_sheet("Real Time Inventory Report")
    write_inventory(ws2, wb_sample["Real Time Inventory Report"])

    latest_sales_month = max([o["date"].month for o in all_orders if o.get("date")], default=1)
    
    log("Generating Offtake Report Summary...")
    ws3 = wb_out.create_sheet("Offtake Report Summary")
    
    if py_data is None:
        base_dir = os.path.dirname(template_path)
        offtake_path = os.path.join(base_dir, "docs", f"{current_year} OFFTAKE REPORT.xlsx")
        if os.path.exists(offtake_path):
            try:
                from .physical_parser import load_py_data
                py_data = load_py_data(offtake_path, py_year=(current_year - 1))
                log("Loaded historical PY data automatically.")
            except Exception as e:
                log(f"Warning: Could not auto-load PY data: {e}")
                py_data = {}
        else:
            py_data = {}
            
    write_offtake_summary(ws3, wb_sample["Offtake Report Summary"], all_orders, py_data or {}, latest_sales_month, current_year, logger=log)

    log("Generating Offtake Report Summary per sku...")
    ws4 = wb_out.create_sheet("Offtake Report Summary per sku")
    write_offtake_per_sku(ws4, wb_sample["Offtake Report Summary per sku"], all_orders, valid_skus)

    log(f"Saving output to {output_path}...")
    try:
        wb_out.save(output_path)
        if logger: logger("Save successful!")
        return True, output_path
    except PermissionError:
        # File is likely open
        timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        dir_name = os.path.dirname(output_path)
        base_name = os.path.basename(output_path).replace(".xlsx", "")
        new_path = os.path.join(dir_name, f"{base_name}_{timestamp}.xlsx")
        
        if logger: logger(f"Permission denied. File might be open. Saving as {new_path}")
        wb_out.save(new_path)
        return True, new_path
    except Exception as e:
        if logger: logger(f"Error saving file: {str(e)}")
        return False, str(e)
