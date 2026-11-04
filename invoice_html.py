"""Invoice HTML template. One logo (Inputs/logo.png) is used for every invoice."""
from html import escape


def fmt_inr(x):
    """Indian digit grouping, 2 decimals: 1,00,000.00"""
    neg = x < 0
    s = f"{abs(x):.2f}"
    whole, dec = s.split(".")
    if len(whole) > 3:
        head, tail = whole[:-3], whole[-3:]
        parts = []
        while len(head) > 2:
            parts.insert(0, head[-2:])
            head = head[:-2]
        if head:
            parts.insert(0, head)
        whole = ",".join(parts) + "," + tail
    return ("-" if neg else "") + whole + "." + dec


def fmt_intl(x):
    return f"{x:,.2f}"


def money(x, cur):
    return fmt_inr(x) if cur == "INR" else fmt_intl(x)


_ONES = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine", "Ten", "Eleven",
         "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
_TENS = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]


def _below_1000(n):
    out = []
    if n >= 100:
        out.append(_ONES[n // 100] + " Hundred")
        n %= 100
    if n >= 20:
        out.append(_TENS[n // 10] + (("-" + _ONES[n % 10]) if n % 10 else ""))
    elif n > 0:
        out.append(_ONES[n])
    return " ".join(out)


def words_indian(n):
    if n == 0:
        return "Zero"
    parts = []
    for div, name in ((10**7, "Crore"), (10**5, "Lakh"), (1000, "Thousand")):
        if n >= div:
            parts.append(_below_1000(n // div) + " " + name)
            n %= div
    if n:
        parts.append(_below_1000(n))
    return " ".join(parts)


def words_intl(n):
    if n == 0:
        return "Zero"
    parts = []
    for div, name in ((10**9, "Billion"), (10**6, "Million"), (1000, "Thousand")):
        if n >= div:
            parts.append(_below_1000(n // div) + " " + name)
            n %= div
    if n:
        parts.append(_below_1000(n))
    return " ".join(parts)


def amount_in_words(amount, cur):
    whole = int(amount)
    cents = int(round((amount - whole) * 100))
    if cur == "INR":
        txt = "Indian Rupees " + words_indian(whole)
        if cents:
            txt += " and " + words_indian(cents) + " Paise"
    else:
        names = {"USD": ("US Dollars", "Cents"), "EUR": ("Euros", "Cents"), "GBP": ("Pounds", "Pence")}
        main, sub = names.get(cur, (cur, "Cents"))
        txt = words_intl(whole) + " " + main
        if cents:
            txt += " and " + words_intl(cents) + " " + sub
    return txt + " Only"


CSS = """
@page { size: A4; margin: 0 }
* { box-sizing: border-box; -webkit-print-color-adjust: exact; print-color-adjust: exact }
body { margin: 0; font-family: "Segoe UI", Calibri, Arial, sans-serif; font-size: 10pt; color: #1c2540 }
.page { width: 210mm; min-height: 295mm; display: flex; flex-direction: column; background: #fff }
.bar { height: 5mm; background: linear-gradient(90deg, #14285f 0%, #1f6fd6 72%, #f2a900 100%) }
.inner { flex: 1; padding: 8mm 14mm 0; display: flex; flex-direction: column }
.sp { flex: 1 1 0; min-height: 0 }
.head { display: flex; justify-content: space-between; align-items: flex-start }
.logo { height: 14mm }
.logo.sq { height: 21mm; margin: -2mm 0 -3mm }
.title { text-align: right }
.title h1 { margin: 0; font-size: 31pt; font-weight: 300; letter-spacing: 7px; color: #14285f; line-height: 1 }
.title .no { display: inline-block; margin-top: 2.5mm; padding: 1mm 4mm; border-radius: 10mm; background: #e8f0fc;
  color: #1f6fd6; font-weight: 700; font-size: 10.5pt; letter-spacing: .5px }
.tiles { display: flex; gap: 3mm; margin: 4mm 0 3mm }
.tile { flex: 1; background: #f4f7fc; border: 1px solid #dfe7f5; border-radius: 3mm; padding: 3mm 4mm }
.k { font-size: 7.5pt; text-transform: uppercase; letter-spacing: 1.2px; color: #6b7896; font-weight: 600 }
.tile .v { font-size: 11.5pt; font-weight: 700; color: #14285f; margin-top: 1.2mm }
.tile.due { background: #14285f; border-color: #14285f }
.tile.due .k { color: #b8c7ee }
.tile.due .v { color: #fff }
.parties { display: flex; gap: 8mm; margin-bottom: 5mm }
.party { flex: 1; border-left: 3px solid #f2a900; padding-left: 4.5mm; line-height: 1.55; font-size: 9.5pt }
.party .n { font-weight: 700; font-size: 12pt; color: #14285f; margin: .8mm 0 .6mm }
.party .meta { margin-top: 1.5mm; color: #3b476a }
.party .meta b { color: #14285f }
.items { width: 100%; border-collapse: separate; border-spacing: 0; border: 1px solid #dfe7f5; border-radius: 3mm; overflow: hidden }
.items th { background: #14285f; color: #fff; text-align: left; padding: 3mm 4mm; font-size: 8.5pt; font-weight: 600;
  letter-spacing: .8px; text-transform: uppercase }
.items td { padding: 3mm 4mm; vertical-align: top; background: #fcfdff }
.items .r { text-align: right; white-space: nowrap }
.items .n { width: 11mm; color: #6b7896 }
.items .hsn { width: 27mm }
.items .rate, .items .amt { width: 33mm }
.items .desc { font-weight: 600; color: #14285f; line-height: 1.45 }
.lower { display: flex; gap: 8mm; margin-top: 4mm; flex: 2 1 auto }
.bottom { display: flex; gap: 6mm; margin-top: 2mm }
.bottom .card { flex: 1; margin-bottom: 0 }
.left { flex: 1.15 }
.right { width: 79mm }
.card { background: #f4f7fc; border: 1px solid #e3eaf6; border-radius: 3mm; padding: 2.6mm 4.5mm; margin-bottom: 2.5mm; line-height: 1.45 }
.card .t { font-weight: 700; color: #14285f; font-size: 10pt }
.pay { display: grid; grid-template-columns: 1fr 1fr; gap: 1.1mm 5mm; margin-top: 1.5mm }
.pi.wide { grid-column: 1 / 3 }
.pi span { display: block; font-size: 7.5pt; color: #6b7896; text-transform: uppercase; letter-spacing: .8px }
.pi b { font-size: 9.5pt; color: #14285f }
.tot { width: 100%; border-collapse: collapse }
.tot td { padding: 2.2mm 1mm; border-bottom: 1px solid #e6ecf6 }
.tot td:last-child { text-align: right; white-space: nowrap; font-weight: 600 }
.tot .wh td { color: #c62828 }
.tot .tt td { font-weight: 700; border-top: 1px solid #c9d2e3 }
.grand { margin-top: 3mm; padding: 3.5mm 1mm 0; border-top: 2px solid #14285f; display: flex; justify-content: space-between; align-items: baseline }
.grand .k { color: #14285f; font-weight: 700; font-size: 9pt }
.grand .amt { font-size: 17pt; font-weight: 800; color: #14285f; letter-spacing: .2px }
.right { display: flex; flex-direction: column }
.signbox { margin-top: auto; padding-top: 6mm; text-align: center }
.sign { height: 16mm; max-width: 58mm; object-fit: contain; display: block; margin: 0 auto 1mm }
.signlbl { border-top: 1px solid #9aa6c4; padding-top: 1.6mm; font-size: 9pt; color: #4a567a }
.items .qty { width: 16mm; text-align: center }
.contact { margin-top: 1.5mm; color: #3b476a }
.tc { font-size: 8.3pt; color: #4a567a; line-height: 1.5; margin: 0; padding-left: 4mm }
.signco { font-size: 8.5pt; color: #6b7896; margin-top: .6mm }
.foot { display: flex; justify-content: space-between; align-items: center; margin: 3mm 14mm 0; padding: 3mm 0 4.5mm;
  border-top: 1px solid #dfe7f5; font-size: 8.5pt; color: #6b7896 }
.foot .thanks { font-size: 11pt; font-weight: 700; color: #14285f }
"""


CURRENCY_SYMBOLS = {"INR": "₹", "USD": "$", "EUR": "€", "GBP": "£"}
CURRENCY_NAMES = {"USD": "US Dollar", "EUR": "Euro", "GBP": "Pound Sterling"}


def _is_squarish(uri):
    """True for a near-square PNG logo (it is shown taller than a wide wordmark)."""
    try:
        import base64, struct
        head = base64.b64decode(uri.split(",", 1)[1][:48])
        w, h = struct.unpack(">II", head[16:24])
        return h > 0 and w / h < 1.6
    except Exception:
        return False


def render_invoice(d, logo_uri, sign_uri):
    e = escape
    cur = d["currency"]
    m = lambda v: money(v, cur)
    p = {k: ("" if str(v).strip().upper().startswith("FILL") else v) for k, v in d["provider"].items()}
    c = d["client"]

    from_lines = [f'<div class="n">{e(p["legal"])}</div>']
    for key in ("address",):
        if p.get(key):
            from_lines.append(f"<div>{e(p[key])}</div>")
    if p.get("gstin"):
        from_lines.append(f'<div class="meta">GSTIN: <b>{e(p["gstin"])}</b></div>')
    contact_lines = [x for x in (p.get("phone"), p.get("email"), p.get("website")) if x]
    if contact_lines:
        from_lines.append('<div class="contact">' + "<br>".join(e(x) for x in contact_lines) + "</div>")

    to_lines = [f'<div class="n">{e(c["legal"])}</div>']
    if c.get("address"):
        to_lines.append(f'<div>{e(c["address"])}</div>')
    if c.get("gstin") and c["gstin"].lower() != "n/a":
        to_lines.append(f'<div class="meta">GSTIN: <b>{e(c["gstin"])}</b></div>')
    meta = [f'Place of supply: <b>{e(d["place"])}</b>']
    if d.get("po"):
        meta.append(f'PO number: <b>{e(d["po"])}</b>')
    to_lines.append('<div class="meta">' + "<br>".join(meta) + "</div>")

    sym = CURRENCY_SYMBOLS.get(cur, cur + " ")
    rows = [f"<tr><td>Sub Total</td><td>{m(d['subtotal'])}</td></tr>"]
    for t in d["tax"]:
        rows.append(f"<tr><td>{e(t['label'])}</td><td>{m(t['amount'])}</td></tr>")
    if d["wh_on"]:
        rows.append(f'<tr class="wh"><td>{e(d["wh_label"])}</td><td>(-) {m(d["withheld"])}</td></tr>')
    rows.append(f'<tr class="tt"><td>Total</td><td>{sym}{m(d["balance"])}</td></tr>')

    notes = (f'<div class="card"><div class="t">Notes</div><div>{e(d["notes"])}</div></div>' if d.get("notes") else "")

    bank = []
    items = []
    if cur != "INR":
        items.append(("Payment in", CURRENCY_NAMES.get(cur, cur), False))
    for label, key, wide in (("Bank name", "bank", False), ("Branch", "branch", False),
                             ("Account name", "account_name", True), ("Account number", "account", False),
                             ("IFSC", "ifsc", False), ("Swift code", "swift", False), ("Account type", "acc_type", False)):
        if p.get(key):
            items.append((label, p[key], wide))
    for label, val, wide in items:
        bank.append(f'<div class="pi{" wide" if wide else ""}"><span>{label}</span><b>{e(val)}</b></div>')
    acct = (p.get("account_text_intl") if cur != "INR" else "") or p.get("account_text")  # non-INR invoices use the Non_India account details
    if acct:  # free-text account details, printed exactly as written in the Excel
        txt_lines = "<br>".join(e(x.strip()) for x in acct.splitlines() if x.strip())
        bank.append(f'<div class="pi wide"><b style="font-weight:600;line-height:1.55">{txt_lines}</b></div>')
    bank_html = f'<div class="card"><div class="t">Payment details</div><div class="pay">{"".join(bank)}</div></div>' if bank else ""
    tl = d.get("terms_list") or []
    terms_html = ('<div class="card"><div class="t">Terms &amp; Conditions</div><ol class="tc">'
                  + "".join(f"<li>{e(x)}</li>" for x in tl) + "</ol></div>") if tl else ""

    qty = d.get("qty") or 1
    qty_txt = f"{qty:g}"
    sign = f'<img class="sign" src="{sign_uri}">' if sign_uri else '<div class="sign"></div>'
    contact = " &nbsp;|&nbsp; ".join(e(x) for x in (p.get("email"), p.get("phone")) if x)

    return f"""<!doctype html><html><head><meta charset="utf-8"><title>{e(d["number"])}</title><style>{CSS}</style></head>
<body><div class="page">
<div class="bar"></div>
<div class="inner">
 <div class="head"><img class="logo{" sq" if _is_squarish(logo_uri) else ""}" src="{logo_uri}">
  <div class="title"><h1>INVOICE</h1><span class="no">{e(d["number"])}</span></div></div>
 <div class="sp"></div>
 <div class="tiles">
  <div class="tile"><div class="k">Invoice Date</div><div class="v">{e(d["date"])}</div></div>
  <div class="tile"><div class="k">Terms of Payment</div><div class="v">{e(d["terms"])}</div></div>
  <div class="tile due"><div class="k">Due Date</div><div class="v">{e(d["due"])}</div></div>
  <div class="tile"><div class="k">Currency</div><div class="v">{e(cur)}</div></div>
 </div>
 <div class="sp"></div>
 <div class="parties">
  <div class="party"><div class="k">Billed from</div>{"".join(from_lines)}</div>
  <div class="party"><div class="k">Billed to</div>{"".join(to_lines)}</div>
 </div>
 <div class="sp"></div>
 <table class="items"><thead><tr><th class="n">#</th><th>Item &amp; Description</th><th class="hsn">HSN/SAC</th><th class="qty">Qty</th><th class="rate r">Rate ({cur})</th><th class="amt r">Amount ({cur})</th></tr></thead>
 <tbody><tr><td class="n">1</td><td class="desc">{e(d["item"])}</td><td>{e(d["hsn"]) if d.get("hsn") else "&ndash;"}</td><td class="qty">{qty_txt}</td><td class="r">{m(d["amount"] / qty)}</td><td class="r">{m(d["amount"])}</td></tr></tbody></table>
 <div class="lower">
  <div class="left">
   <div class="card"><div class="k">Total due in words</div><div class="t" style="margin-top:1mm">{e(d["words"])}</div></div>
   {notes}
  </div>
  <div class="right">
   <table class="tot">{"".join(rows)}</table>
   <div class="grand"><div class="k">Balance Due</div><div class="amt">{sym}{m(d["balance"])}</div></div>
   <div class="signbox">{sign}<div class="signlbl">{e(d["sign_label"])}</div><div class="signco">{e(p.get("legal", ""))}</div></div>
  </div>
 </div>
 <div class="bottom">{bank_html}{terms_html}</div>
 <div class="sp"></div>
</div>
<div class="foot"><span><span class="thanks">Thank you for your business!</span><br>We look forward to continuing our partnership.</span><span>{contact}</span></div>
</div></body></html>"""
