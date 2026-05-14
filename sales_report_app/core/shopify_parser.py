import csv
from .workbook_utils import fval, parse_shopify_date, bank_label
from collections import defaultdict

def load_shopify_transactions(path):
    txns = {}
    with open(path, newline="", encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            name   = row["Name"].strip()
            status = row["Status"].strip().lower()
            entry  = txns.get(name)
            if entry is None or (status == "success" and entry["status"] != "success"):
                txns[name] = {
                    "gateway":  row["Gateway"].strip(),
                    "status":   status,
                    "amount":   fval(row["Amount"]),
                    "paid_at":  parse_shopify_date(row["Created At"]),
                }
    return txns

def load_shopify_orders(path, txns, paymongo_rate=0.015):
    rows, seen = [], set()
    with open(path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            name = r["Name"].strip()
            if name in seen or not r["Financial Status"].strip():
                continue
            seen.add(name)
            txn     = txns.get(name, {})
            status  = r["Financial Status"].strip().lower()
            paid_at = parse_shopify_date(r["Paid at"]) or txn.get("paid_at")
            gateway = txn.get("gateway", r["Payment Method"].strip())
            is_pm   = "paymongo" in gateway.lower() or "paymongo" in r["Payment Method"].lower()
            total   = fval(r["Total"])
            disc    = fval(r["Discount Amount"])
            ship    = fval(r["Shipping"])
            subtotal= fval(r["Subtotal"])
            gross   = subtotal + disc   # before discount
            dep_dt  = paid_at if status == "paid" else None
            rows.append({
                "source":       "Shopify",
                "date":         parse_shopify_date(r["Created at"]),
                "customer":     r["Billing Name"].strip(),
                "order_no":     name,
                "gross":        gross,
                "shipping":     ship or None,
                "discount":     disc or None,
                "paymongo_fee": round(total * paymongo_rate, 2) if is_pm else None,
                "payment_fee":  None,
                "net":          total,
                "deposit_date": dep_dt,
                "bank":         bank_label(r["Payment Method"], gateway, status),
                "ref":          r["Payment Reference"].strip() or None,
                "status":       status,
                "items":        [],   # filled later
            })
    return rows

def load_shopify_items(orders_path):
    """Return {order_name: [(product_name, qty, price)]}"""
    items = defaultdict(list)
    with open(orders_path, newline="", encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            items[r["Name"].strip()].append((
                r["Lineitem name"].strip(),
                int(fval(r["Lineitem quantity"])),
                fval(r["Lineitem price"]),
            ))
    return items

def process_shopify(orders_path, txns_path):
    """High-level function to return a unified list of shopify orders."""
    txns = load_shopify_transactions(txns_path)
    shopify_orders = load_shopify_orders(orders_path, txns)
    shopify_items  = load_shopify_items(orders_path)
    for o in shopify_orders:
        o["items"] = shopify_items.get(o["order_no"], [])
    return shopify_orders
