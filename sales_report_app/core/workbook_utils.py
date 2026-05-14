import copy
from datetime import datetime

def fval(s):
    """Safely convert a string to float."""
    try:
        return float(str(s).strip().replace(",","")) if s not in (None,"","–","-") else 0.0
    except:
        return 0.0

def parse_shopify_date(s):
    if not s or not str(s).strip():
        return None
    try:
        return datetime.strptime(str(s).strip()[:19], "%Y-%m-%d %H:%M:%S")
    except:
        return None

def parse_shopee_date(s):
    if not s or str(s).strip() in ("", "-"):
        return None
    try:
        return datetime.strptime(str(s).strip()[:16], "%Y-%m-%d %H:%M")
    except:
        return None

def parse_lazada_date(s):
    if not s or str(s).strip() == "":
        return None
    for fmt in ("%d %b %Y %H:%M", "%d %b %Y"):
        try:
            return datetime.strptime(str(s).strip(), fmt)
        except:
            pass
    return None

def copy_style(src, dst):
    """Copy formatting and styling from one cell to another."""
    if src.has_style:
        dst.font      = copy.copy(src.font)
        dst.fill      = copy.copy(src.fill)
        dst.border    = copy.copy(src.border)
        dst.alignment = copy.copy(src.alignment)
        dst.number_format = src.number_format

def bank_label(method, gateway="", status="paid"):
    """Determine the bank/payment method label for the report."""
    if status != "paid":
        return "PENDING"
    combined = (method + " " + gateway).lower()
    if "paymongo" in combined:
        return "UBP"
    if "bank deposit" in combined:
        return "BANK DEPOSIT"
    if "manual" in combined:
        return "MANUAL"
    return method or ""
