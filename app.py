"""Invoice tool: local web app. Run with start.bat (or: python app.py).

Everything lives in this folder:
  Inputs/   Invoice Generation Inputs.xlsx   tabs: Invoices To Generate, Client Master, Service Provider, Settings
            last_invoice_numbers.txt         last number issued per provider (no database)
            logo.png, signature.png (optional)
  Outputs/  generated PDFs + Invoices-<Month>-<timestamp>.zip
The Excel file is read again every time the invoices are loaded and every time Generate is clicked.
"""
import base64
import datetime as dt
import io
import json
import mimetypes
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import urllib.parse
import urllib.request
import uuid
import webbrowser
from decimal import ROUND_HALF_UP, Decimal
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from openpyxl import load_workbook

from invoice_html import amount_in_words, render_invoice

BASE = Path(__file__).resolve().parent
INPUTS = BASE / "Inputs"
OUTPUTS = BASE / "Outputs"
WORK = OUTPUTS / ".work"
STATIC = BASE / "static"
INPUT_XLSX = INPUTS / "Invoice Generation Inputs.xlsx"
LAST_FILE = INPUTS / "last_invoice_numbers.txt"
LOG_FILE = INPUTS / "issued_numbers.log"
PORT = int(os.environ.get("INVOICE_TOOL_PORT", "8765"))

LOCK = threading.Lock()
STATE = {"rows": {}, "batches": {}}


class UserError(Exception):
    pass


# ----------------------------------------------------------------- helpers
def norm(s):
    return re.sub(r"[^a-z0-9]", "", str(s if s is not None else "").lower())


def txt(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def nums(s):
    return [float(x) for x in re.findall(r"\d+(?:\.\d+)?", str(s or ""))]


def yes(v):
    return norm(v) in ("yes", "y", "true", "1")


def colmap(headers, aliases):
    """aliases: {key: [normalised header names]} -> {key: column index}"""
    hn = [norm(h) for h in headers]
    out = {}
    for key, names in aliases.items():
        for n in names:
            if n in hn:
                out[key] = hn.index(n)
                break
    return out


CM_ALIASES = {
    "short": ["clientnameshort", "clientshort", "shortname"],
    "legal": ["clientnamelegal", "clientnameofficial", "legalname"],
    "provider": ["serviceprovider", "serviceprovidername", "provider"],
    "gstin": ["gstin"],
    "place": ["placeofsupply", "place", "placeofsupply7"],
    "gsttype": ["gsttype"],
    "gstpct": ["gst", "gstpct", "gstpercent"],
    "terms": ["paymentterms", "terms"],
    "withheld": ["amountwithheld"],
    "whpct": ["withheld", "withheldpct", "whpct", "percentwithheld"],
    "po": ["poneeded", "poneeeded"],
    "active": ["active"],
    "address": ["addressoptional", "address"],
}
SP_ALIASES = {
    "short": ["companynameshort", "serviceprovidershort", "shortname"],
    "legal": ["companynamelegal", "serviceproviderlegal", "legalname"],
    "address": ["address"],
    "gstin": ["gstin"],
    "phone": ["phone"],
    "email": ["email"],
    "account": ["accountnumber", "accountno"],
    "ifsc": ["ifsc", "ifsccode"],
    "account_name": ["nameonaccount", "accountname"],
    "website": ["website", "web"],
    "bank": ["bankname", "bank"],
    "branch": ["branch", "bankbranch"],
    "swift": ["swiftcode", "swift"],
    "acc_type": ["accounttype"],
    "account_text": ["accountdetails"],
}
INV_ALIASES = {
    "client": ["clientname", "client"],
    "currency": ["currency"],
    "amount": ["amount"],
    "item": ["itemdescription", "item"],
    "notes": ["notestoputininvoice", "notes"],
    "month": ["includeininvoicename", "monthforinvoicename", "customstring", "month"],
    "provider": ["serviceprovidername", "serviceprovider"],
    "po": ["ponumber", "po"],
    "hsn": ["hsnsac", "hsn", "sac"],
    "qty": ["qty", "quantity"],
}


def cell_text(v, key=""):
    if isinstance(v, (dt.datetime, dt.date)):
        if key == "gstpct":       # "9, 9" typed in Excel becomes 9-Sep
            return f"{v.month}, {v.day}"
        if key == "month":        # "Oct2026" typed in Excel becomes 1-Oct-2026
            return v.strftime("%b%Y")
        return v.strftime("%d/%m/%Y")
    return txt(v)


def read_sheet_rows(ws):
    return [list(row) for row in ws.iter_rows(values_only=True)]


def table_from_grid(grid, aliases, required=(), explicit=None):
    """Find the header row (first row containing the required headers). Returns (colmap, data rows, header index).
    explicit: optional {field: header text or None} chosen by the user in the mapping wizard."""
    if explicit:
        aliases = {k: ([norm(explicit[k])] if explicit.get(k) else []) if k in explicit else v for k, v in aliases.items()}
    for i, row in enumerate(grid):
        cm = colmap(row, aliases)
        if all(k in cm for k in required):
            return cm, grid[i + 1:], i
    return None, [], -1


def open_workbook(path=None):
    path = Path(path) if path else source_path()
    if not path.exists():
        raise UserError("'Invoice Generation Inputs.xlsx' was not found in the Inputs folder. Upload it on the first screen.")
    tmp = Path(tempfile.gettempdir()) / f"inputs_{uuid.uuid4().hex}.xlsx"
    shutil.copy(path, tmp)  # copy, so the file can stay open in Excel
    try:
        wb = load_workbook(tmp, data_only=True)
    except Exception:
        tmp.unlink(missing_ok=True)
        raise UserError("That file is not a valid .xlsx workbook.")
    return wb, tmp


def sheet_by_name(wb, *names):
    want = [norm(n) for n in names]
    for ws in wb.worksheets:
        if norm(ws.title) in want:
            return ws
    raise UserError(f"Tab '{names[0]}' not found in the Excel file.")


def pick_sheet(wb, configured, *fallbacks):
    names = ([configured] if configured else []) + list(fallbacks)
    return sheet_by_name(wb, *names)


def load_inputs(path=None, with_invoices=True, cfg=None):
    cfg = cfg if cfg is not None else load_config()
    tabs, cols = cfg.get("tabs") or {}, cfg.get("columns") or {}
    wb, tmp = open_workbook(path)
    try:
        grid = read_sheet_rows(pick_sheet(wb, tabs.get("clients"), "Client Master", "Clients"))
        cm, data, _ = table_from_grid(grid, CM_ALIASES, required=("short", "legal"), explicit=cols.get("clients"))
        if cm is None:
            raise UserError("Client Master tab: could not find the columns for client short and legal name. Check the column mapping.")
        clients = []
        for r in data:
            g = lambda k: cell_text(r[cm[k]], k) if k in cm and cm[k] < len(r) else ""
            if g("short"):
                clients.append({k: g(k) for k in CM_ALIASES})

        grid = read_sheet_rows(pick_sheet(wb, tabs.get("providers"), "Service Provider", "Service Providers", "Service Provider Master"))
        sp, data, _ = table_from_grid(grid, SP_ALIASES, required=("short", "legal"), explicit=cols.get("providers"))
        if sp is None:
            raise UserError("Service Provider tab: could not find the columns for company short and legal name. Check the column mapping.")
        providers = {}
        for r in data:
            g = lambda k: cell_text(r[sp[k]], k) if k in sp and sp[k] < len(r) else ""
            if g("short"):
                providers[g("short").upper()] = {k: g(k) for k in SP_ALIASES}

        settings = {"gst_rate": 18.0, "wh_label": "Amount Withheld (Section 393(1) Sl6(iii)D(b)-Fees)", "sign_label": "Authorized Signature", "terms_list": []}
        try:
            for r in read_sheet_rows(pick_sheet(wb, tabs.get("settings"), "Settings")):
                if not r or r[0] is None:
                    continue
                k, v = norm(r[0]), (r[1] if len(r) > 1 else None)
                if v is None or txt(v) == "":
                    continue
                if k.startswith("gstrate"):
                    settings["gst_rate"] = float(nums(v)[0])
                elif k.startswith("withholding"):
                    settings["wh_label"] = txt(v)
                elif k.startswith("signature"):
                    settings["sign_label"] = txt(v)
                elif k.startswith("termsandconditions") or k.startswith("termsconditions"):
                    settings["terms_list"].append(txt(v))
        except UserError:
            pass

        invoices = []
        if with_invoices:
            grid = read_sheet_rows(pick_sheet(wb, tabs.get("invoices"), "Invoices To Generate", "Invoice Table", "Invoices"))
            ic, data, hdr = table_from_grid(grid, INV_ALIASES, required=("client", "currency", "amount"), explicit=cols.get("invoices"))
            if ic is None:
                raise UserError("Invoices tab: could not find the Client, Currency and Amount columns. Check the column mapping.")
            for off, r in enumerate(data):
                g = lambda k: cell_text(r[ic[k]], k) if k in ic and ic[k] < len(r) else ""
                if not g("client") and not g("amount"):
                    continue
                invoices.append({"sheet_row": hdr + off + 2, **{k: g(k) for k in INV_ALIASES}})
        return clients, providers, settings, invoices
    finally:
        tmp.unlink(missing_ok=True)


# ------------------------------------------------------------ rules / calc
def find_client(clients, name):
    n = norm(name)
    for c in clients:
        if norm(c["short"]) == n:
            return c
    for c in clients:
        if norm(c["legal"]) == n:
            return c
    return None


def find_provider(providers, text):
    n = norm(text)
    if not n:
        return None
    for k in providers:
        if norm(k) == n:
            return k
    for k, p in providers.items():
        ln = norm(p["legal"])
        if n in ln or ln.startswith(n):
            return k
    return None


def terms_days(terms):
    m = re.search(r"(\d+)", terms or "")
    return int(m.group(1)) if m else None


def gst_plan(currency, client, settings):
    """Returns (list of (label, pct), error)."""
    if currency != "INR":
        return [("IGST", 0.0)], None
    t = norm(client["gsttype"])
    pcts = nums(client["gstpct"])
    if "igst" in t:
        return [("IGST", pcts[0] if pcts else settings["gst_rate"])], None
    if "cgst" in t or "sgst" in t:
        if len(pcts) >= 2:
            return [("CGST", pcts[0]), ("SGST", pcts[1])], None
        half = (pcts[0] if pcts else settings["gst_rate"]) / 2
        return [("CGST", half), ("SGST", half)], None
    return [("IGST", settings["gst_rate"])], "GST type missing for an INR client in the Client Master."


def to_amount(s):
    try:
        return float(str(s).replace(",", "").replace("$", "").replace("\u20b9", "").strip())
    except ValueError:
        return None


def providers_have_bank(p):
    ok = lambda v: bool(v) and not v.upper().startswith("FILL")
    return ok(p.get("account_text", "")) or (ok(p.get("account", "")) and ok(p.get("ifsc", "")))


def analyze(row, clients, providers, settings, last):
    errors, warnings = [], []
    client = find_client(clients, row["client"])
    cur = row["currency"].strip().upper()
    amount = to_amount(row["amount"])
    pkey = None
    if client:
        pkey = find_provider(providers, client["provider"]) if client["provider"] else None
    if pkey is None and row.get("provider"):
        pkey = find_provider(providers, row["provider"])
    out = {
        "id": row["id"], "sheet_row": row["sheet_row"], "client": row["client"], "currency": cur, "amount": amount,
        "item": row["item"], "notes": row["notes"], "month": row["month"], "hsn": row["hsn"],
        "provider": pkey, "po": row["po"], "client_missing": client is None,
    }
    if client is None:
        errors.append("Client is not in the Client Master. Add it to continue.")
    if not cur:
        errors.append("Currency is missing.")
    if amount is None or amount <= 0:
        errors.append("Amount is missing or not a number.")
    if not row["item"]:
        errors.append("Item description is missing.")
    if not row["month"]:
        errors.append("'Include in Invoice Name' is missing.")
    if not row["hsn"] and cur == "INR":  # HSN/SAC is only needed for Indian (INR) invoices
        errors.append("HSN/SAC is missing (needed for INR invoices).")
    qty = 1.0
    if row.get("qty"):
        qty = to_amount(row["qty"])
        if qty is None or qty <= 0:
            errors.append("Qty is not a valid number.")
            qty = 1.0
    out["qty"] = qty
    if pkey is None:
        if client:
            errors.append("Service provider is missing for this client in the Client Master (IAC or ITGS).")
    else:
        if pkey not in last:
            errors.append(f"Last invoice number for {pkey} is not set in Inputs/last_invoice_numbers.txt.")
        miss = []
        if not providers[pkey]["gstin"] or providers[pkey]["gstin"].upper().startswith("FILL"):
            miss.append("GSTIN")
        if not providers_have_bank(providers[pkey]):
            miss.append("account details")
        if miss:
            warnings.append(f"Service Provider tab: {', '.join(miss)} not filled for {pkey}.")
    defaults = None
    if client:
        needs_po = yes(client["po"])
        if needs_po and not row["po"]:
            errors.append("PO number is required for this client but missing.")
        if not needs_po and row["po"]:
            warnings.append("PO number given, but this client is not marked 'PO needed'.")
        if terms_days(client["terms"]) is None:
            errors.append("Payment terms missing in the Client Master.")
        plan, perr = gst_plan(cur, client, settings) if cur else ([], None)
        if perr:
            errors.append(perr)
        places = [p.strip() for p in re.split(r"[;\n|]", client["place"]) if p.strip()]
        whp = nums(client["whpct"])
        defaults = {
            "places": places, "place": places[0] if places else "",
            "terms": client["terms"], "wh_on": yes(client["withheld"]), "wh_pct": whp[0] if whp else None,
            "wh_label": settings["wh_label"], "tax": [{"label": l, "pct": p} for l, p in plan],
        }
        if not places:
            errors.append("Place of supply missing in the Client Master.")
        out["legal"] = client["legal"]
    out["defaults"] = defaults
    out["errors"], out["warnings"] = errors, warnings
    return out


def _money(x):
    return Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def calc(amount, tax_plan, wh_on, wh_pct):
    """Tax and withholding are both worked out on the sub total (before GST). Balance = sub total + GST - withheld."""
    sub = _money(amount)
    tax, tax_sum = [], Decimal(0)
    for name, pct in tax_plan:
        amt = _money(sub * Decimal(str(pct)) / 100)
        tax.append({"label": f"{name}{pct:g} ({pct:g}%)", "pct": pct, "amount": float(amt)})
        tax_sum += amt
    wh = _money(sub * Decimal(str(wh_pct)) / 100) if wh_on else Decimal(0)
    total = sub + tax_sum
    return {"subtotal": float(sub), "tax": tax, "withheld": float(wh), "total": float(total), "balance": float(total - wh)}


# ----------------------------------------------------------- invoice numbers
# Inputs/last_invoice_numbers.txt holds one line per provider:  IAC=INV-000152
# The tool counts up the last group of digits and keeps the rest of the format as it is.
def read_last():
    d = {}
    if LAST_FILE.exists():
        for line in LAST_FILE.read_text(encoding="utf-8-sig").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            if v.strip():
                d[k.strip().upper()] = v.strip()
    return d


def write_last(d):
    lines = ["# Last invoice number issued for each service provider (one per line, PROVIDER=NUMBER).",
             "# The tool updates this file every time invoices are generated.",
             "# To continue from a different number, edit the value after the = sign."]
    lines += [f"{k}={v}" for k, v in d.items()]
    LAST_FILE.write_text("\n".join(lines) + "\n", encoding="utf-8")


def bump(num):
    groups = list(re.finditer(r"\d+", num))
    if not groups:
        raise UserError(f"The last invoice number '{num}' has no digits to count up from.")
    g = groups[-1]
    return num[:g.start()] + str(int(g.group()) + 1).zfill(len(g.group())) + num[g.end():]


def issued_set():
    out = set()
    if LOG_FILE.exists():
        for line in LOG_FILE.read_text(encoding="utf-8").splitlines():
            parts = line.split(",")
            if len(parts) >= 2:
                out.add(parts[1].strip())
    return out


def next_numbers(providers):
    last, issued, out = read_last(), issued_set(), {}
    for k in providers:
        if k not in last:
            continue
        seq, cur = [], last[k]
        for _ in range(60):
            cur = bump(cur)
            while cur in issued:
                cur = bump(cur)
            seq.append(cur)
        out[k] = {"last": last[k], "seq": seq}
    return out


# ---------------------------------------------------------------- PDF output
def find_browser():
    cands = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    ]
    for c in cands:
        if os.path.exists(c):
            return c
    for name in ("msedge", "chrome", "google-chrome", "chromium"):
        p = shutil.which(name)
        if p:
            return p
    raise UserError("Microsoft Edge or Google Chrome is required to create PDFs, but none was found.")


def html_to_pdf(html_path, pdf_path):
    exe = find_browser()
    prof = tempfile.mkdtemp(prefix="inv_prof_")
    try:
        cmd = [exe, "--headless=new", "--disable-gpu", "--no-pdf-header-footer", "--print-to-pdf-no-header",
               f"--user-data-dir={prof}", f"--print-to-pdf={pdf_path}", Path(html_path).as_uri()]
        subprocess.run(cmd, timeout=90, capture_output=True)
        for _ in range(30):
            if os.path.exists(pdf_path) and os.path.getsize(pdf_path) > 0:
                return
            time.sleep(0.3)
        raise UserError("PDF was not created.")
    finally:
        shutil.rmtree(prof, ignore_errors=True)


def data_uri(path):
    if not path.exists():
        return ""
    mime = mimetypes.guess_type(str(path))[0] or "image/png"
    return f"data:{mime};base64," + base64.b64encode(path.read_bytes()).decode()


def safe_name(s):
    return re.sub(r'[\\/:*?"<>|]', "", s).strip()


def parse_date(s):
    return dt.datetime.strptime(s, "%Y-%m-%d").date()


# ----------------------------------------------------------------- API logic
def api_load(body=None):
    clients, providers, settings, raw = load_inputs()
    if not raw:
        raise UserError("The tab 'Invoices To Generate' has no rows.")
    last = read_last()
    rows = []
    with LOCK:
        STATE["rows"] = {}
        for i, r in enumerate(raw):
            r["id"] = i
            STATE["rows"][i] = r
            rows.append(analyze(r, clients, providers, settings, last))
    return {"ok": True, "rows": rows, "numbers": next_numbers(providers), "today": dt.date.today().isoformat(),
            "providers": sorted(providers)}


def api_add_client(body):
    if (load_config().get("source") or {}).get("type") == "google":
        raise UserError("Your data comes from a Google Sheet. Add the client in the Client Master tab of that sheet, then click Reload.")
    wb = load_workbook(INPUT_XLSX)
    ws = sheet_by_name(wb, "Client Master", "Clients")
    grid = read_sheet_rows(ws)
    cm, _, hdr = table_from_grid(grid, CM_ALIASES, required=("short", "legal"))
    if cm is None:
        raise UserError("Client Master header not found.")
    last_row = max((i for i, r in enumerate(grid) if any(v is not None for v in r)), default=hdr)
    new_row = last_row + 2
    sno = [r[0] for r in grid[hdr + 1:] if r and isinstance(r[0], (int, float))]
    for k, ci in cm.items():
        ws.cell(row=new_row, column=ci + 1, value=body.get(k, ""))
    ws.cell(row=new_row, column=1, value=(int(max(sno)) + 1 if sno else 1))
    try:
        wb.save(INPUT_XLSX)
    except PermissionError:
        raise UserError("'Invoice Generation Inputs.xlsx' is open in Excel. Close it and try again.")
    return {"ok": True}


def provider_logo(pkey):
    """Inputs/logo-<PROVIDER>.png (e.g. logo-IAC.png) if present, otherwise the shared Inputs/logo.png."""
    own = INPUTS / f"logo-{pkey}.png" if pkey else None
    return data_uri(own if own and own.exists() else INPUTS / "logo.png")


def api_generate(body):
    clients, providers, settings, _ = load_inputs(with_invoices=False)  # Excel is re-read on every Generate
    last = read_last()
    issued = issued_set()
    sign = data_uri(INPUTS / "signature.png")
    batch =dt.datetime.now().strftime("%Y%m%d-%H%M%S") + "-" + uuid.uuid4().hex[:4]
    bdir = WORK / batch
    bdir.mkdir(parents=True, exist_ok=True)
    results, used_in_batch, details = [], set(), {}

    for sel in body["rows"]:
        raw = STATE["rows"].get(sel["id"])
        res = {"id": sel["id"]}
        results.append(res)
        if raw is None:
            res["error"] = "Row not found. Load the invoices again."
            continue
        a = analyze(raw, clients, providers, settings, last)
        if a["errors"]:
            res["error"] = " ".join(a["errors"])
            continue
        client = find_client(clients, raw["client"])
        d0 = a["defaults"]
        place = sel.get("place") or d0["place"]
        terms = sel.get("terms") or d0["terms"]
        days = terms_days(terms)
        if days is None:
            res["error"] = "Terms must look like Net30."
            continue
        try:
            inv_date = parse_date(sel["date"])
        except Exception:
            res["error"] = "Invoice date is not valid."
            continue
        wh_on = bool(sel.get("wh_on"))
        wh_pct = to_amount(sel.get("wh_pct")) if wh_on else 0
        if wh_on and (wh_pct is None or wh_pct <= 0):
            res["error"] = "Withholding is ticked but the % is missing."
            continue
        pkey = a["provider"]
        number = (sel.get("number") or "").strip()
        manual = bool(sel.get("number_edited"))
        if not manual or not number:
            number = bump(last[pkey])
            while number in issued or number in used_in_batch:
                number = bump(number)
        elif number in issued or number in used_in_batch or number == last[pkey]:
            res["error"] = f"Invoice number {number} is already used."
            continue
        used_in_batch.add(number)
        c = calc(a["amount"], [(x["label"], x["pct"]) for x in d0["tax"]], wh_on, wh_pct or 0)
        cur = a["currency"]
        d = {
            "pkey": pkey, "number": number, "date": inv_date.strftime("%d/%m/%Y"), "terms": terms,
            "due": (inv_date + dt.timedelta(days=days)).strftime("%d/%m/%Y"), "place": place, "po": raw["po"],
            "currency": cur, "item": raw["item"], "notes": raw["notes"], "amount": a["amount"], "hsn": raw["hsn"],
            "qty": a["qty"], "terms_list": [],
            "wh_on": wh_on, "wh_label": sel.get("wh_label") or settings["wh_label"], "wh_pct": wh_pct or 0,
            "words": amount_in_words(c["balance"], cur), "sign_label": settings["sign_label"],
            "provider": dict(providers[pkey]), "client": client, **c,
        }
        fname = safe_name(f'{number} {client["short"]} {raw["month"]}') + ".pdf"
        html_path = bdir / f'{sel["id"]}.html'
        pdf_path = bdir / fname
        html_path.write_text(render_invoice(d, provider_logo(pkey), sign), encoding="utf-8")
        try:
            html_to_pdf(html_path, pdf_path)
        except Exception as ex:
            res["error"] = f"PDF failed: {ex}"
            continue
        # record the number only after the PDF exists
        issued.add(number)
        last[pkey] = number
        with open(LOG_FILE, "a", encoding="utf-8") as lf:
            lf.write(f"{pkey},{number},{dt.datetime.now():%Y-%m-%d %H:%M:%S}\n")
        details[sel["id"]] = {"d": d, "plan": [(x["label"], x["pct"]) for x in d0["tax"]], "html": str(html_path), "pdf": str(pdf_path)}
        res.update({"ok": True, "number": number, "file": fname, "balance": c["balance"], "currency": cur,
                    "due": d["due"], "date": d["date"], "terms": terms, "place": place,
                    "subtotal": c["subtotal"], "tax": c["tax"], "withheld": c["withheld"], "wh_on": wh_on,
                    "wh_pct": wh_pct or 0, "wh_label": d["wh_label"], "client": client["short"], "client_legal": client["legal"],
                    "provider": pkey, "month": raw["month"], "po": raw["po"], "hsn": raw["hsn"]})
    write_last(last)
    STATE["batches"][batch] = {r["id"]: r for r in results if r.get("ok")}
    STATE.setdefault("detail", {})[batch] = details
    return {"ok": True, "batch": batch, "results": results, "numbers": next_numbers(providers)}


def api_regenerate(body):
    """Edit the withholding line, terms, date or place of one reviewed invoice and rebuild its PDF. The invoice number stays."""
    batch, iid = body.get("batch"), body.get("id")
    det = (STATE.get("detail", {}).get(batch) or {}).get(iid)
    if det is None:
        raise UserError("This invoice is no longer available. Generate again.")
    d = det["d"]
    wh_on = bool(body.get("wh_on"))
    wh_pct = to_amount(body.get("wh_pct")) if wh_on else 0
    if wh_on and (wh_pct is None or wh_pct <= 0):
        raise UserError("Withholding is ticked but the % is missing.")
    terms = (body.get("terms") or "").strip()
    days = terms_days(terms)
    if days is None:
        raise UserError("Terms of payment must look like Net30.")
    try:
        inv_date = parse_date(body.get("date"))
    except Exception:
        raise UserError("Invoice date is not valid.")
    place = (body.get("place") or "").strip()
    if not place:
        raise UserError("Place of supply is missing.")
    label = (body.get("wh_label") or "").strip() or load_inputs(with_invoices=False)[2]["wh_label"]
    c = calc(d["amount"], det["plan"], wh_on, wh_pct or 0)
    d.update(c)
    d.update({"wh_on": wh_on, "wh_pct": wh_pct or 0, "wh_label": label, "terms": terms, "place": place,
              "date": inv_date.strftime("%d/%m/%Y"), "due": (inv_date + dt.timedelta(days=days)).strftime("%d/%m/%Y"),
              "words": amount_in_words(c["balance"], d["currency"])})
    html = render_invoice(d, provider_logo(d.get("pkey")), data_uri(INPUTS / "signature.png"))
    Path(det["html"]).write_text(html, encoding="utf-8")
    Path(det["pdf"]).unlink(missing_ok=True)
    html_to_pdf(det["html"], det["pdf"])
    upd = {"place": place, "date": d["date"], "terms": terms, "due": d["due"], "tax": c["tax"], "subtotal": c["subtotal"],
           "withheld": c["withheld"], "wh_on": wh_on, "wh_pct": wh_pct or 0, "wh_label": label, "balance": c["balance"]}
    STATE["batches"][batch][iid].update(upd)
    return {"ok": True, "result": upd}


def api_finalize(body):
    batch = body["batch"]
    info = STATE["batches"].get(batch)
    if info is None:
        raise UserError("Batch not found. Generate again.")
    stamp = dt.datetime.now()
    base = f"Invoices-{stamp.strftime('%B')}-{stamp.strftime('%Y%m%d-%H%M%S')}"
    folder = OUTPUTS / base
    folder.mkdir(parents=True, exist_ok=True)
    files = []
    for i in body["correct"]:
        r = info.get(i) or info.get(str(i))
        if not r:
            continue
        shutil.copy(WORK / batch / r["file"], folder / r["file"])
        files.append(r["file"])
    if not files:
        raise UserError("No invoices marked correct.")
    zpath = OUTPUTS / f"{base}.zip"
    with zipfile.ZipFile(zpath, "w", zipfile.ZIP_DEFLATED) as z:
        for f in files:
            z.write(folder / f, f)
    shutil.rmtree(WORK / batch, ignore_errors=True)
    return {"ok": True, "zip": zpath.name, "folder": base, "files": files, "outputs": str(OUTPUTS)}


# ------------------------------------------- data source setup: upload or Google Sheet, tab + column mapping
CONFIG_FILE = INPUTS / "config.json"
GOOGLE_CACHE = INPUTS / "google_source.xlsx"
INSPECT_FILE = INPUTS / "_inspect.xlsx"
ROLES = ("invoices", "clients", "providers")
ROLE_ALIASES = {"invoices": INV_ALIASES, "clients": CM_ALIASES, "providers": SP_ALIASES}
ROLE_TITLES = {"invoices": "Invoices to generate", "clients": "Client Master", "providers": "Service Providers"}
ROLE_REQUIRED = {
    "invoices": ("client", "currency", "amount", "item", "month"),
    "clients": ("short", "legal", "provider", "place", "terms"),
    "providers": ("short", "legal"),
}
FIELD_LABELS = {
    "invoices": {"client": "Client name", "currency": "Currency", "amount": "Amount", "item": "Item description",
                 "notes": "Notes to put in invoice", "month": "Include in invoice name", "hsn": "HSN/SAC (needed for INR)",
                 "po": "PO number", "provider": "Service provider (optional)", "qty": "Qty (optional)"},
    "clients": {"short": "Client name (short)", "legal": "Client name (legal)", "provider": "Service provider",
                "gstin": "GSTIN", "place": "Place of supply", "gsttype": "GST type", "gstpct": "GST %",
                "terms": "Payment terms", "withheld": "Amount withheld?", "whpct": "% withheld", "po": "PO needed?",
                "active": "Active?", "address": "Address (optional)"},
    "providers": {"short": "Company name (short)", "legal": "Company name (legal)", "address": "Address",
                  "gstin": "GSTIN", "phone": "Phone", "email": "Email", "account_text": "Account details (free text)",
                  "account": "Account number", "ifsc": "IFSC", "account_name": "Name on account",
                  "website": "Website", "bank": "Bank name", "branch": "Branch", "swift": "Swift code",
                  "acc_type": "Account type"},
}


def load_config():
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding="utf-8-sig"))
        except Exception:
            return {}
    return {}


def save_config(cfg):
    CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8")


def fetch_google_xlsx(url, dest):
    m = re.search(r"/spreadsheets/d/([\w-]+)", url or "")
    if not m:
        raise UserError("That does not look like a Google Sheet link.")
    xurl = f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=xlsx"
    try:
        req = urllib.request.Request(xurl, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=60) as resp:
            data = resp.read()
    except Exception as ex:
        raise UserError(f"Could not open the Google Sheet ({ex}). Share it as 'Anyone with the link can view'.")
    if data[:2] != b"PK":
        raise UserError("The Google Sheet is private. Share it as 'Anyone with the link can view' and try again.")
    Path(dest).write_bytes(data)


def source_path():
    """The workbook to read: the uploaded file, or a fresh download of the saved Google Sheet."""
    src = load_config().get("source") or {}
    if src.get("type") == "google" and src.get("url"):
        fetch_google_xlsx(src["url"], GOOGLE_CACHE)
        return GOOGLE_CACHE
    return INPUT_XLSX


def _best_header(grid, aliases):
    best = (0, -1)
    for i, row in enumerate(grid[:40]):
        sc = len(colmap(row, aliases))
        if sc > best[0]:
            best = (sc, i)
    return best


def _first_filled(grid):
    for i, row in enumerate(grid[:40]):
        if sum(1 for v in row if txt(v)) >= 2:
            return i
    return 0


def _inspect_grids():
    ins = STATE.get("inspect")
    if not ins or not INSPECT_FILE.exists():
        raise UserError("Please choose your Excel file or paste a Google Sheet link first.")
    if "grids" not in ins:
        wb, tmp = open_workbook(INSPECT_FILE)
        try:
            ins["grids"] = {ws.title: read_sheet_rows(ws)[:80] for ws in wb.worksheets}
        finally:
            tmp.unlink(missing_ok=True)
    return ins["grids"]


def suggest_tabs(grids):
    out = {"invoices": None, "clients": None, "providers": None, "settings": None}
    keys = {"invoices": ("invoice",), "clients": ("client",), "providers": ("provider",)}
    for role in ROLES:
        best = (0, None)
        for name, grid in grids.items():
            sc = _best_header(grid, ROLE_ALIASES[role])[0]
            if any(k in norm(name) for k in keys[role]):
                sc += 1
            if sc > best[0]:
                best = (sc, name)
        out[role] = best[1]
    for name in grids:
        if "setting" in norm(name):
            out["settings"] = name
    return out


def suggest_columns(grid, role):
    aliases = ROLE_ALIASES[role]
    sc, i = _best_header(grid, aliases)
    if i < 0:
        i = _first_filled(grid)
    row = grid[i] if grid else []
    headers = [txt(v) for v in row]
    cm = colmap(row, aliases)
    nxt = grid[i + 1] if i + 1 < len(grid) else []
    return {
        "headers": [h for h in headers if h],
        "map": {k: (headers[cm[k]] if k in cm and cm[k] < len(headers) and headers[cm[k]] else None) for k in aliases},
        "samples": {h: cell_text(nxt[j], "") for j, h in enumerate(headers) if h and j < len(nxt)},
        "header_row": i + 1,
    }


def inspect_payload(source):
    STATE["inspect"] = {"source": source}
    grids = _inspect_grids()
    tabs = []
    for name, grid in grids.items():
        i = _first_filled(grid)
        tabs.append({"name": name, "rows": max(0, len(grid) - i - 1),
                     "headers": [txt(v) for v in (grid[i] if grid else []) if txt(v)][:8]})
    return {"ok": True, "source": source, "tabs": tabs, "suggest": suggest_tabs(grids),
            "titles": ROLE_TITLES,
            "fields": {r: [{"key": k, "label": FIELD_LABELS[r][k], "required": k in ROLE_REQUIRED[r]}
                           for k in ROLE_ALIASES[r]] for r in ROLES}}


def api_source_workbook(data, filename):
    if Path(filename).suffix.lower() != ".xlsx":
        raise UserError("Please choose an Excel .xlsx file.")
    INSPECT_FILE.write_bytes(data)
    try:
        return inspect_payload({"type": "file", "name": filename})
    except UserError:
        raise
    except Exception:
        raise UserError("That file could not be read as an Excel workbook.")


def api_source_google(body):
    url = (body.get("url") or "").strip()
    fetch_google_xlsx(url, INSPECT_FILE)
    return inspect_payload({"type": "google", "url": url})


def api_source_columns(body):
    grids = _inspect_grids()
    tabs = body.get("tabs") or {}
    out = {}
    for role in ROLES:
        name = tabs.get(role)
        if not name or name not in grids:
            raise UserError(f"Choose a tab for '{ROLE_TITLES[role]}'.")
        out[role] = suggest_columns(grids[name], role)
    return {"ok": True, "columns": out}


def api_source_save(body):
    _inspect_grids()
    source, tabs, cols = body.get("source") or {}, body.get("tabs") or {}, body.get("columns") or {}
    for role in ROLES:
        if not tabs.get(role):
            raise UserError(f"Choose a tab for '{ROLE_TITLES[role]}'.")
        missing = [FIELD_LABELS[role][k] for k in ROLE_REQUIRED[role] if not (cols.get(role) or {}).get(k)]
        if missing:
            raise UserError(f"{ROLE_TITLES[role]}: choose a column for {', '.join(missing)}.")
    cfg = {"source": source, "tabs": tabs, "columns": cols}
    load_inputs(INSPECT_FILE, cfg=cfg)  # trial run: raises a clear error if the mapping does not work
    if source.get("type") == "google":
        shutil.copy(INSPECT_FILE, GOOGLE_CACHE)
    else:
        if INPUT_XLSX.exists():
            (INPUTS / "backup").mkdir(exist_ok=True)
            shutil.copy(INPUT_XLSX, INPUTS / "backup" / f"Invoice Generation Inputs-{dt.datetime.now():%Y%m%d-%H%M%S}.xlsx")
        shutil.copy(INSPECT_FILE, INPUT_XLSX)
    save_config(cfg)
    INSPECT_FILE.unlink(missing_ok=True)
    STATE.pop("inspect", None)
    return {"ok": True, "summary": inputs_summary()}


def api_last_numbers(body):
    """Save the last invoice number per provider, typed by the user. Keeps providers that are not mentioned."""
    nums_in = body.get("numbers") or {}
    clean = {}
    for k, v in nums_in.items():
        k, v = str(k).strip().upper(), str(v).strip()
        if not k or not v:
            continue
        if not re.search(r"\d", v):
            raise UserError(f"The number for {k} needs at least one digit to count up from, for example INV-000152.")
        clean[k] = v
    if not clean:
        raise UserError("Type at least one last invoice number.")
    write_last({**read_last(), **clean})
    return {"ok": True, "numbers": read_last()}


# ------------------------------------------------------- uploaded input files
def inputs_summary(path=None):
    clients, providers, settings, invoices = load_inputs(path)
    fill = sum((not p["gstin"] or p["gstin"].upper().startswith("FILL")) + (not providers_have_bank(p))
               for p in providers.values())
    return {"clients": len(clients), "providers": sorted(providers), "unfilled": fill, "invoices": len(invoices),
            "last_numbers": read_last()}


def api_status():
    out = {"inputs": None, "inputs_error": "", "source": load_config().get("source") or {"type": "file"},
           "sign_v": int((INPUTS / "signature.png").stat().st_mtime) if (INPUTS / "signature.png").exists() else 0,
           "logo": (INPUTS / "logo.png").exists(),
           "signature": (INPUTS / "signature.png").exists(),
           "logo_v": int((INPUTS / "logo.png").stat().st_mtime) if (INPUTS / "logo.png").exists() else 0}
    try:
        out["inputs"] = inputs_summary()
    except UserError as ex:
        out["inputs_error"] = str(ex)
    return {"ok": True, **out}


def _image_to_png(data, dest, white_bg):
    from PIL import Image, ImageChops
    try:
        im = Image.open(io.BytesIO(data))
        im.load()
    except Exception:
        raise UserError("That file is not a picture (use PNG, JPG or WEBP).")
    im = im.convert("RGBA")
    if white_bg:
        bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
        im = Image.alpha_composite(bg, im).convert("RGB")
        bbox = ImageChops.difference(im, Image.new("RGB", im.size, (255, 255, 255))).getbbox()
    else:
        bbox = im.getchannel("A").getbbox()
    if bbox:
        im = im.crop(bbox)
    im.save(dest, "PNG")


def api_upload(kind, filename, data):
    if not data:
        raise UserError("The file is empty.")
    if len(data) > 25 * 1024 * 1024:
        raise UserError("File is too large (limit 25 MB).")
    if kind == "inputs":
        if Path(filename).suffix.lower() != ".xlsx":
            raise UserError("Use an .xlsx file (tabs: Invoices To Generate, Client Master, Service Provider).")
        tmp = Path(tempfile.gettempdir()) / f"up_{uuid.uuid4().hex}.xlsx"
        tmp.write_bytes(data)
        try:
            summary = inputs_summary(tmp)
            if summary["clients"] == 0 or not summary["providers"]:
                raise UserError("The file has no clients or no service providers.")
            if INPUT_XLSX.exists():
                (INPUTS / "backup").mkdir(exist_ok=True)
                shutil.copy(INPUT_XLSX, INPUTS / "backup" / f"Invoice Generation Inputs-{dt.datetime.now():%Y%m%d-%H%M%S}.xlsx")
            shutil.copy(tmp, INPUT_XLSX)
        finally:
            tmp.unlink(missing_ok=True)
        return {"ok": True, "summary": summary}
    if kind == "workbook":
        return api_source_workbook(data, filename)
    if kind == "last":
        pairs = {}
        for line in data.decode("utf-8-sig", errors="replace").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            k, v = line.split("=", 1)
            if not re.search(r"\d", v):
                raise UserError(f"'{line}' has no number to count up from.")
            pairs[k.strip().upper()] = v.strip()
        if not pairs:
            raise UserError("No lines like IAC=INV-000152 were found in that file.")
        write_last({**read_last(), **pairs})
        return {"ok": True, "providers": sorted(pairs)}
    if kind == "logo":
        _image_to_png(data, INPUTS / "logo.png", True)
        return {"ok": True}
    if kind == "signature":
        _image_to_png(data, INPUTS / "signature.png", False)
        return {"ok": True}
    raise UserError("Unknown upload type.")


# ------------------------------------------------------- transparent logo for the page header
_LOGO_UI = {}


def logo_ui_png():
    """The logo with its white background turned into transparency (page UI only, invoices keep the original)."""
    p = INPUTS / "logo.png"
    if not p.exists():
        return None
    key = p.stat().st_mtime
    if _LOGO_UI.get("k") != key:
        from PIL import Image
        im = Image.open(p).convert("RGB")
        if im.width > 900:
            im = im.resize((900, round(im.height * 900 / im.width)), Image.LANCZOS)
        px = im.load()
        out = Image.new("RGBA", im.size)
        op = out.load()
        for y in range(im.height):
            for x in range(im.width):
                r, g, b = px[x, y]
                a = 255 - min(r, g, b)
                if a <= 2:
                    op[x, y] = (0, 0, 0, 0)
                    continue
                k = a / 255.0
                op[x, y] = tuple(max(0, min(255, int((c - 255 * (1 - k)) / k))) for c in (r, g, b)) + (a,)
        buf = io.BytesIO()
        out.save(buf, "PNG")
        _LOGO_UI.update(k=key, data=buf.getvalue())
    return _LOGO_UI["data"]


# ------------------------------------------------------------------- server
class Handler(BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def send_bytes(self, data, ctype, code=200, extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(data)

    def send_json(self, obj, code=200):
        self.send_bytes(json.dumps(obj).encode("utf-8"), "application/json", code)

    def send_file(self, path, ctype=None, extra=None):
        if not path.is_file():
            return self.send_bytes(b"Not found", "text/plain", 404)
        ctype = ctype or mimetypes.guess_type(str(path))[0] or "application/octet-stream"
        self.send_bytes(path.read_bytes(), ctype, extra=extra)

    def do_GET(self):
        path = urllib.parse.unquote(urllib.parse.urlparse(self.path).path)
        if path in ("/", "/index.html"):
            return self.send_file(STATIC / "index.html", "text/html; charset=utf-8")
        if path == "/logo":
            return self.send_file(INPUTS / "logo.png")
        if path == "/signature":
            return self.send_file(INPUTS / "signature.png")
        if path == "/logo-ui":
            data = logo_ui_png()
            if data:
                return self.send_bytes(data, "image/png")
            return self.send_bytes(b"Not found", "text/plain", 404)
        if path == "/api/status":
            return self.send_json(api_status())
        if path == "/download-input/inputs":
            return self.send_file(INPUT_XLSX, extra={
                "Content-Disposition": 'attachment; filename="Invoice Generation Inputs.xlsx"'})
        m = re.fullmatch(r"/pdf/([\w-]+)/(\d+)", path)
        if m:
            r = STATE["batches"].get(m.group(1), {}).get(int(m.group(2)))
            if r:
                return self.send_file(WORK / m.group(1) / r["file"], "application/pdf",
                                      {"Content-Disposition": "inline"})
        m = re.fullmatch(r"/download/(.+)", path)
        if m:
            f = (OUTPUTS / m.group(1)).resolve()
            if OUTPUTS.resolve() in f.parents and f.is_file():
                return self.send_file(f, extra={"Content-Disposition": f'attachment; filename="{f.name}"'})
        self.send_bytes(b"Not found", "text/plain", 404)

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        try:
            n = int(self.headers.get("Content-Length", 0))
            raw = self.rfile.read(n)
            if path == "/api/upload":
                q = urllib.parse.parse_qs(parsed.query)
                fname = urllib.parse.unquote(self.headers.get("X-Filename", "file"))
                return self.send_json(api_upload(q.get("kind", [""])[0], fname, raw))
            body = json.loads(raw or b"{}")
            if path == "/api/load":
                return self.send_json(api_load(body))
            if path == "/api/last-numbers":
                return self.send_json(api_last_numbers(body))
            if path == "/api/source/google":
                return self.send_json(api_source_google(body))
            if path == "/api/source/columns":
                return self.send_json(api_source_columns(body))
            if path == "/api/source/save":
                return self.send_json(api_source_save(body))
            if path == "/api/add-client":
                return self.send_json(api_add_client(body))
            if path == "/api/generate":
                return self.send_json(api_generate(body))
            if path == "/api/regenerate":
                return self.send_json(api_regenerate(body))
            if path == "/api/finalize":
                return self.send_json(api_finalize(body))
            if path == "/api/open-outputs":
                if hasattr(os, "startfile"):
                    os.startfile(str(OUTPUTS))
                else:
                    subprocess.Popen(["open" if sys.platform == "darwin" else "xdg-open", str(OUTPUTS)])
                return self.send_json({"ok": True})
            self.send_bytes(b"Not found", "text/plain", 404)
        except UserError as ex:
            self.send_json({"ok": False, "error": str(ex)})
        except Exception as ex:
            import traceback
            traceback.print_exc()
            self.send_json({"ok": False, "error": f"Unexpected error: {ex}"})


def main():
    INPUTS.mkdir(exist_ok=True)
    OUTPUTS.mkdir(exist_ok=True)
    WORK.mkdir(exist_ok=True)
    srv = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    url = f"http://127.0.0.1:{PORT}/"
    print(f"Invoice tool running at {url}  (close this window to stop)")
    if "--no-browser" not in sys.argv and not os.environ.get("INVOICE_TOOL_NO_BROWSER"):
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
