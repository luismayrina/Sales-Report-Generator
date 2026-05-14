from openpyxl import load_workbook
from collections import defaultdict
from .workbook_utils import fval, parse_shopee_date

def load_shopee(path):
    wb = load_workbook(path)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    hi = {h: i for i, h in enumerate(headers)}

    orders = []
    all_rows = list(ws.iter_rows(min_row=2, values_only=True))

    # Group multi-product orders
    grouped = defaultdict(list)
    for row in all_rows:
        oid = row[hi["Order ID"]]
        grouped[oid].append(row)

    for oid, grp in grouped.items():
        first    = grp[0]
        status   = str(first[hi["Order Status"]]).strip()
        paid_time = parse_shopee_date(str(first[hi["Order Paid Time"]]))
        created   = parse_shopee_date(str(first[hi["Order Creation Date"]]))

        # Sum amounts across line items
        total_payment = sum(fval(r[hi["Total Buyer Payment"]]) for r in grp)
        total_fee     = sum(fval(r[hi["Service Fee"]]) for r in grp)
        total_disc    = sum(fval(r[hi["Total Discount(PHP)"]]) for r in grp)
        total_ship    = sum(fval(r[hi["Buyer Paid Shipping Fee"]]) for r in grp)

        # Grand Total from first row (it's shared per order)
        grand = fval(first[hi["Grand Total"]])

        is_cancelled = status.lower() == "cancelled"
        dep_dt = paid_time if not is_cancelled else None

        items = [(str(r[hi["Product Name"]]).strip(),
                  int(fval(r[hi["Quantity"]])),
                  fval(r[hi["Original Price"]])) for r in grp]

        orders.append({
            "source":       "Shopee",
            "date":         created,
            "customer":     str(first[hi["Receiver Name"]]).strip(),
            "order_no":     oid,
            "gross":        total_payment,
            "shipping":     total_ship or None,
            "discount":     total_disc or None,
            "paymongo_fee": None,
            "payment_fee":  total_fee or None,   # Shopee commission
            "net":          grand,
            "deposit_date": dep_dt,
            "bank":         "UBP" if not is_cancelled else "CANCELLED",
            "ref":          None,
            "status":       "cancelled" if is_cancelled else "paid",
            "items":        items,
        })
    return orders
