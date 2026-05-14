import os
from datetime import datetime
from collections import defaultdict
import openpyxl
from openpyxl import load_workbook

from .workbook_utils import copy_style
from .sku_normalizer import normalize_sku, extract_valid_skus_from_template, match_sku

# ── SHEET 1: DAILY SALES SUMMARY ─────────────────────────────────────────────
def write_daily_sales(ws_out, ws_sample, all_orders):
    # Copy header row
    for j, cell in enumerate(ws_sample[1], 1):
        dst = ws_out.cell(row=1, column=j, value=cell.value)
        copy_style(cell, dst)

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


def write_offtake_summary(ws_out, ws_sample, all_orders, logger=None):
    def log(msg):
        if logger:
            logger(msg)

    ONLINE_CHANNELS = {"Shopify", "Shopee", "Lazada"}
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

    # ── Step 3: Determine all months needed from sales data ─────────────────
    sales_months = set()
    for o in all_orders:
        if o["status"] == "cancelled":
            continue
        src = o.get("source")
        if src not in ONLINE_CHANNELS:
            continue
        dt = o.get("date")
        if dt:
            sales_months.add(dt.month)

    if not sales_months:
        log("No valid online sales found. Copying template only.")
        return

    # Determine the contiguous range Jan→latest month.
    # We only create blocks from month 1 up to the latest month with data,
    # skipping any months earlier than the first template month.
    earliest_template_month = min(template_month_sections.keys())
    latest_sales_month = max(sales_months)
    log(f"Latest sales month detected: {MONTH_NAMES[latest_sales_month - 1]} (month {latest_sales_month})")

    # Months we need to have blocks for (from earliest in template to latest in data)
    needed_months = list(range(earliest_template_month, latest_sales_month + 1))
    log(f"Month blocks needed: {[MONTH_NAMES[m-1] for m in needed_months]}")

    # ── Step 4: Create missing month blocks ─────────────────────────────────
    # Find the last template block to use as a style source for new blocks.
    last_template_month = max(template_month_sections.keys())
    last_section = template_month_sections[last_template_month]
    last_block_start = last_section["section_start"]
    next_block_start = last_block_start + SECTION_WIDTH  # where the next new block goes

    all_sections = dict(template_month_sections)   # month_num -> section dict

    max_data_row = max(30, ws_out.max_row)  # limit style copying to data area
    for month_num in needed_months:
        if month_num in all_sections:
            continue  # already exists in template
        log(f"Creating new month block for {MONTH_NAMES[month_num-1]} at col {next_block_start}")
        _copy_block_styles(ws_out, last_block_start, next_block_start, max_data_row, log)

        # Write row-1 header cells for the new block
        month_dt = datetime(2026, month_num, 1)
        month_abbr = MONTH_NAMES[month_num - 1]

        sec = _build_section(month_num, next_block_start)
        ws_out.cell(row=1, column=sec["pct_cont_col"],   value="% cont")
        ws_out.cell(row=1, column=sec["actual_col"],      value=month_dt)
        ws_out.cell(row=1, column=sec["actual_col"]).number_format = "mmm-yy"
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

    # ── Step 5: Detect channel rows ─────────────────────────────────────────
    channel_row_map = {}
    total_row = None
    ecommerce_row = None
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
        if nu == "ECOMMERCE":
            ecommerce_row = i

    # ── Step 6: Calculate monthly totals per online channel ─────────────────
    channel_monthly = defaultdict(lambda: defaultdict(float))
    for o in all_orders:
        if o["status"] == "cancelled":
            continue
        src = o.get("source")
        if src not in ONLINE_CHANNELS:
            continue
        dt = o.get("date")
        if not dt:
            continue
        channel_monthly[src][dt.month] += o.get("net") or 0

    for ch, mdata in channel_monthly.items():
        for m, total in sorted(mdata.items()):
            log(f"  {ch} {MONTH_NAMES[m-1]}: {total:.2f}")

    # ── Step 7: Write actuals, YTD, vs PY for online channels ───────────────
    sorted_months = sorted(all_sections.keys())

    for ch in ONLINE_CHANNELS:
        ridx = channel_row_map.get(ch)
        if not ridx:
            log(f"WARNING: No row for '{ch}'. Skipping.")
            continue
        ytd_sum = 0.0
        for month_num in sorted_months:
            sec = all_sections[month_num]
            monthly_val = channel_monthly[ch].get(month_num, 0.0)
            ytd_sum += monthly_val

            # Always overwrite actual & YTD so old template data is cleared
            ws_out.cell(row=ridx, column=sec["actual_col"],
                        value=round(monthly_val, 2) if monthly_val else 0)
            ws_out.cell(row=ridx, column=sec["ytd_col"],
                        value=round(ytd_sum, 2) if ytd_sum else 0)
            log(f"  {ch} {MONTH_NAMES[month_num-1]} actual={monthly_val:.2f}  YTD={ytd_sum:.2f}")

            # vs PY — use template PY value if it exists
            py_val = ws_sample.cell(ridx, sec["py_col"]).value if month_num in template_month_sections else None
            if isinstance(py_val, (int, float)) and py_val:
                ws_out.cell(row=ridx, column=sec["vs_py_col"],
                            value=round(monthly_val - py_val, 2))

            ytd_py_val = ws_sample.cell(ridx, sec["ytd_py_col"]).value if month_num in template_month_sections else None
            if isinstance(ytd_py_val, (int, float)) and ytd_py_val:
                ws_out.cell(row=ridx, column=sec["ytd_vs_py_col"],
                            value=round(ytd_sum - ytd_py_val, 2))

    # ── Step 8: Recalculate TOTAL row ────────────────────────────────────────
    if total_row:
        for month_num in sorted_months:
            sec = all_sections[month_num]
            total_actual = 0.0
            total_ytd = 0.0
            for ch_name, ridx in channel_row_map.items():
                if ch_name.upper() in SKIP_ROWS:
                    continue
                v_a = ws_out.cell(ridx, sec["actual_col"]).value
                v_y = ws_out.cell(ridx, sec["ytd_col"]).value
                if isinstance(v_a, (int, float)):
                    total_actual += v_a
                if isinstance(v_y, (int, float)):
                    total_ytd += v_y
            ws_out.cell(row=total_row, column=sec["actual_col"],
                        value=round(total_actual, 2) if total_actual else 0)
            ws_out.cell(row=total_row, column=sec["ytd_col"],
                        value=round(total_ytd, 2) if total_ytd else 0)
            log(f"  TOTAL {MONTH_NAMES[month_num-1]} actual={total_actual:.2f} YTD={total_ytd:.2f}")

    # ── Step 9: Recalculate % cont for all channels ──────────────────────────
    for month_num in sorted_months:
        sec = all_sections[month_num]
        total_actual = ws_out.cell(total_row, sec["actual_col"]).value if total_row else None
        total_ytd    = ws_out.cell(total_row, sec["ytd_col"]).value    if total_row else None

        for ch_name, ridx in channel_row_map.items():
            if ch_name.upper() in SKIP_ROWS:
                continue
            row_actual = ws_out.cell(ridx, sec["actual_col"]).value
            row_ytd    = ws_out.cell(ridx, sec["ytd_col"]).value

            if isinstance(row_actual, (int, float)) and isinstance(total_actual, (int, float)) and total_actual:
                pct = round(row_actual / total_actual, 4)
                ws_out.cell(row=ridx, column=sec["pct_cont_col"], value=pct)
                ws_out.cell(row=ridx, column=sec["pct_cont_col"]).number_format = "0%"

            if isinstance(row_ytd, (int, float)) and isinstance(total_ytd, (int, float)) and total_ytd:
                ytd_pct = round(row_ytd / total_ytd, 4)
                ws_out.cell(row=ridx, column=sec["ytd_pct_col"], value=ytd_pct)
                ws_out.cell(row=ridx, column=sec["ytd_pct_col"]).number_format = "0%"

    # ── Step 10: Validation ───────────────────────────────────────────────────
    max_template_month = max(template_month_sections.keys())
    if latest_sales_month > max_template_month:
        log(f"INFO: Template had {MONTH_NAMES[max_template_month-1]} as last month. "
            f"Created new blocks up to {MONTH_NAMES[latest_sales_month-1]}.")
    log(f"Offtake Report Summary complete. Months populated: "
        f"{[MONTH_NAMES[m-1] for m in sorted_months if m in sales_months or m in template_month_sections]}")






# ── SHEET 4: OFFTAKE PER SKU ─────────────────────────────────────────────────
def write_offtake_per_sku(ws_out, ws_sample, all_orders, valid_skus, unmapped_log):
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
            matched_sku = match_sku(normalized, valid_skus)
            
            if matched_sku:
                sku_qty[matched_sku][month] += qty
            else:
                # Log unmapped SKU
                unmapped_log.append({
                    "Platform": o["source"],
                    "Order No": o["order_no"],
                    "Raw Name": item_name,
                    "Normalized": normalized,
                    "Qty": qty,
                    "Date": dt.strftime("%Y-%m-%d")
                })

    for sku, month_data in sku_qty.items():
        row_idx = sku_row.get(sku)
        if not row_idx: continue
        for month, qty in month_data.items():
            col_idx = month_col.get(month)
            if col_idx:
                existing = ws_out.cell(row=row_idx, column=col_idx).value
                ws_out.cell(row=row_idx, column=col_idx, value=(existing or 0) + qty)

# ── SHEET 5: UNMAPPED SKUS ───────────────────────────────────────────────────
def write_unmapped_skus(ws_out, unmapped_log):
    headers = ["Platform", "Order No", "Date", "Raw Name", "Normalized", "Qty"]
    for j, h in enumerate(headers, 1):
        ws_out.cell(row=1, column=j, value=h)
    
    for i, log in enumerate(unmapped_log, 2):
        ws_out.cell(row=i, column=1, value=log["Platform"])
        ws_out.cell(row=i, column=2, value=log["Order No"])
        ws_out.cell(row=i, column=3, value=log["Date"])
        ws_out.cell(row=i, column=4, value=log["Raw Name"])
        ws_out.cell(row=i, column=5, value=log["Normalized"])
        ws_out.cell(row=i, column=6, value=log["Qty"])


# ── MAIN GENERATOR ───────────────────────────────────────────────────────────
def generate_full_report(all_orders, template_path, output_path, logger=None):
    if logger: logger("Loading template workbook...")
    wb_sample = load_workbook(template_path)
    
    valid_skus = extract_valid_skus_from_template(wb_sample["Offtake Report Summary per sku"])
    unmapped_log = []

    wb_out = openpyxl.Workbook()

    if logger: logger("Generating Daily Sales Summary...")
    ws1 = wb_out.active
    ws1.title = "Daily Sales Summary"
    write_daily_sales(ws1, wb_sample["Daily Sales Summary "], all_orders)

    if logger: logger("Generating Real Time Inventory Report...")
    ws2 = wb_out.create_sheet("Real Time Inventory Report")
    write_inventory(ws2, wb_sample["Real Time Inventory Report"])

    if logger: logger("Generating Offtake Report Summary...")
    ws3 = wb_out.create_sheet("Offtake Report Summary")
    write_offtake_summary(ws3, wb_sample["Offtake Report Summary"], all_orders, logger=logger)

    if logger: logger("Generating Offtake Report Summary per sku...")
    ws4 = wb_out.create_sheet("Offtake Report Summary per sku")
    write_offtake_per_sku(ws4, wb_sample["Offtake Report Summary per sku"], all_orders, valid_skus, unmapped_log)
    
    if unmapped_log:
        if logger: logger(f"Found {len(unmapped_log)} unmapped SKU entries. Creating 'Unmapped SKUs' sheet...")
        ws5 = wb_out.create_sheet("Unmapped SKUs")
        write_unmapped_skus(ws5, unmapped_log)

    if logger: logger(f"Saving output to {output_path}...")
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
