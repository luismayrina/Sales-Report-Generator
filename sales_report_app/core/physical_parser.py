import openpyxl
from datetime import datetime
import calendar
from collections import defaultdict

MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12
}

MONTH_ABBRS = {
    1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MAY", 6: "JUN",
    7: "JUL", 8: "AUG", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC"
}

SHEET_CHANNEL_MAP = {
    "COMMON ROOM": "Common Room",
    "FRANKIE": "Frankie and Friends",
    "SIMULAPH": "Simula PH",
    "9 Matters": "9 Matters",
    "TCC": "Craft Central",
    "TIKTOK": "Tiktok",
    "PICKAROO": "Pick-a-roo"
}

CHANNEL_ABBR = {
    "Common Room": "CR",
    "Frankie and Friends": "FF",
    "Craft Central": "CC",
    "Simula PH": "SP",
    "9 Matters": "9M",
    "Tiktok": "TT",
    "Pick-a-roo": "PR"
}

def load_py_data(filepath, py_year=2025):
    """
    Load historical previous year (PY) sales data from the dynamically named PY sheet (e.g. '2025').
    Returns dict: channel_name -> month_num -> sales_amount
    """
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        if str(py_year) not in wb.sheetnames:
            print(f"Warning: '{py_year}' sheet not found in {filepath}. PY data will be skipped.")
            return {}
        ws = wb[str(py_year)]
        
        headers = [str(ws.cell(1, c).value).strip().upper() for c in range(1, ws.max_column + 1)]
        months = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
        month_cols = {}
        for idx, h in enumerate(headers):
            if h in months:
                month_cols[months.index(h) + 1] = idx + 1
                
        py_data = defaultdict(dict)
        for r in range(2, ws.max_row + 1):
            channel = ws.cell(r, 1).value
            if not channel:
                continue
            channel_name = str(channel).strip()
            # Normalize some common names to match template names if they differ
            if channel_name == "Frankie & Friends":
                channel_name = "Frankie and Friends"
            
            for m_num, col_idx in month_cols.items():
                val = ws.cell(r, col_idx).value
                try:
                    py_data[channel_name][m_num] = float(val) if val is not None else 0.0
                except:
                    py_data[channel_name][m_num] = 0.0
        return py_data
    except Exception as e:
        print(f"Error loading PY ({py_year}) sheet from {filepath}: {e}")
        return {}

def parse_daily_sheet(ws, channel_name, current_year=2026):
    """
    Parse sheets where sales are listed by day (e.g., Tiktok, Pick-a-roo).
    Constructs an exact datetime for the sale using `current_year`.
    """
    row3 = [ws.cell(3, c).value for c in range(1, ws.max_column + 1)]
    row4 = [ws.cell(4, c).value for c in range(1, ws.max_column + 1)]
    
    transactions = []
    abbr = CHANNEL_ABBR.get(channel_name, channel_name[:2].upper())
    
    for col_idx in range(2, ws.max_column):
        cust_name = row3[col_idx]
        date_val = row4[col_idx]
        
        # If Row 3 has a customer name, it is a transaction column!
        if not cust_name or str(cust_name).strip() == "":
            continue
            
        cust_name = str(cust_name).strip()
        
        # Parse date
        order_date = None
        if isinstance(date_val, datetime):
            order_date = date_val
        elif date_val:
            try:
                order_date = datetime.strptime(str(date_val).strip()[:10], "%Y-%m-%d")
            except:
                pass
                
        # Fallback date if not found or invalid (find closest month start in row 2 to the left)
        if not order_date:
            for c in range(col_idx, 1, -1):
                r2_val = ws.cell(2, c + 1).value
                if isinstance(r2_val, datetime):
                    order_date = r2_val
                    break
                    
        if not order_date:
            order_date = datetime(current_year, 1, 1)
            
        # Construct datetime for the daily record
        try:
            order_date = datetime(current_year, order_date.month, order_date.day)
        except ValueError:
            # fallback to start of month if invalid day
            order_date = datetime(current_year, order_date.month, 1)
                
        # Parse items
        items = []
        for r_idx in range(5, ws.max_row + 1):
            variant = ws.cell(r_idx, 1).value
            price = ws.cell(r_idx, 2).value
            qty = ws.cell(r_idx, col_idx + 1).value
            
            if variant is None:
                continue
                
            variant_str = str(variant).strip().lower()
            if (any(kw in variant_str for kw in ["shipping", "discount", "paymongo", "payment fee", "remitted"])
                or variant_str.isdigit() 
                or variant_str.replace('.', '', 1).isdigit() 
                or variant_str == "total"
                or variant_str.startswith("popjunklove")):
                break
                
            if qty and isinstance(qty, (int, float)) and qty > 0:
                price = float(price) if price is not None and isinstance(price, (int, float)) else 0.0
                items.append((str(variant).strip(), int(qty), price))
                
        if items:
            gross = sum(qty * price for v, qty, price in items)
            transactions.append({
                "source": channel_name,
                "date": order_date,
                "customer": cust_name,
                "order_no": f"{abbr}-{order_date.strftime('%Y%m%d')}-{col_idx+1}",
                "gross": gross,
                "shipping": 0.0,
                "discount": 0.0,
                "paymongo_fee": None,
                "payment_fee": None,
                "net": gross,
                "deposit_date": order_date,
                "bank": channel_name.upper(),
                "ref": None,
                "status": "paid",
                "items": items
            })
            
    return transactions

def parse_monthly_sheet(ws, channel_name, current_year=2026):
    """
    Parse typical physical sheets (e.g., Common Room, Frankie and Friends)
    where data is organized in monthly blocks.
    Constructs an order matching the first day of that month in `current_year`.
    """
    row2 = [ws.cell(2, c).value for c in range(1, ws.max_column + 1)]
    row3 = [ws.cell(3, c).value for c in range(1, ws.max_column + 1)]
    
    current_location = channel_name
    orders_by_month_loc = {} # (month_num, location) -> list of (variant, qty, price)
    
    # Identify active columns
    active_cols = []
    for col_idx in range(2, ws.max_column):
        loc_val = row2[col_idx]
        month_val = row3[col_idx]
        if loc_val:
            current_location = str(loc_val).strip()
        if not month_val:
            continue
        m_str = str(month_val).strip().lower()
        if m_str in MONTH_MAP:
            active_cols.append((col_idx + 1, MONTH_MAP[m_str], current_location))
            
    # Iterate through rows
    for r_idx in range(5, ws.max_row + 1):
        variant = ws.cell(r_idx, 1).value
        price = ws.cell(r_idx, 2).value
        
        if variant is None:
            continue
            
        variant_str = str(variant).strip().lower()
        if (any(kw in variant_str for kw in ["shipping", "discount", "paymongo", "payment fee", "remitted"])
            or variant_str.isdigit() 
            or variant_str.replace('.', '', 1).isdigit() 
            or variant_str == "total" 
            or variant_str.startswith("popjunklove")):
            break
            
        price = float(price) if price is not None and isinstance(price, (int, float)) else 0.0
        
        # Read the quantities for active columns
        for col_num, month_num, loc in active_cols:
            qty = ws.cell(r_idx, col_num).value
            if qty and isinstance(qty, (int, float)) and qty > 0:
                key = (month_num, loc)
                if key not in orders_by_month_loc:
                    orders_by_month_loc[key] = []
                orders_by_month_loc[key].append((str(variant).strip(), int(qty), price))
                
    # Convert groupings to summary orders
    transactions = []
    abbr = CHANNEL_ABBR.get(channel_name, channel_name[:2].upper())
    
    for (month_num, loc), items in sorted(orders_by_month_loc.items()):
        if not items:
            continue
            
        # Get last day of the month in current_year
        last_day = calendar.monthrange(current_year, month_num)[1]
        order_date = datetime(current_year, month_num, last_day)
        
        gross = sum(qty * price for v, qty, price in items)
        loc_clean = "".join(c for c in loc if c.isalnum())
        loc_abbr = loc_clean[:8].upper()
        month_abbr = MONTH_ABBRS[month_num]
        
        transactions.append({
            "source": channel_name,
            "date": order_date,
            "customer": f"{channel_name} - {loc} Customer",
            "order_no": f"{abbr}-{loc_abbr}-{month_abbr}-{current_year}",
            "gross": gross,
            "shipping": 0.0,
            "discount": 0.0,
            "paymongo_fee": None,
            "payment_fee": None,
            "net": gross,
            "deposit_date": order_date,
            "bank": channel_name.upper(),
            "ref": None,
            "status": "paid",
            "items": items
        })
        
    return transactions

def load_physical_channels(filepath, current_year=2026):
    """
    Load data from the OFFTAKE REPORT.xlsx.
    Iterates over all sheets. Skips summary sheets and the previous year sheet.
    Uses 'parse_daily_sheet' for specific channels, else 'parse_monthly_sheet'.
    """
    all_transactions = []
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        py_year = str(current_year - 1)
        
        for sheetname in wb.sheetnames:
            sheet_upper = sheetname.upper().strip()
            
            # Skip summary/reporting sheets and the Previous Year sheet
            if (sheet_upper in (py_year, "ACCOUNT PERFORMANCE", "ONLINE SUMMARY", "LAZADA", "SHOPEE", "SHOPIFY")
                or sheet_upper.endswith("OFFTAKE REPORT")
                or "SUMMARY" in sheet_upper):
                continue
            
            # Map sheetname to target channel
            channel_name = SHEET_CHANNEL_MAP.get(sheetname)
            if not channel_name:
                # Fallback to sheetname if not mapped
                channel_name = sheetname
                
            ws = wb[sheetname]
            
            # Decide parsing method: TIKTOK and PICKAROO are daily, others are monthly summary
            if sheet_upper in ("TIKTOK", "PICK-A-ROO"):
                txns = parse_daily_sheet(ws, channel_name=channel_name, current_year=current_year)
            else:
                txns = parse_monthly_sheet(ws, channel_name=channel_name, current_year=current_year)
                
            all_transactions.extend(txns)
            
    except Exception as e:
        print(f"Error parsing physical channels from {filepath}: {e}")
        
    return all_transactions
