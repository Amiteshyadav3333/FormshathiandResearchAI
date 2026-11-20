import csv
import io
import re
import json
import sys
import base64
import zipfile
from urllib.parse import urlparse, parse_qs
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from flask import Flask, jsonify, request, send_from_directory, send_file
from werkzeug.exceptions import BadRequest, RequestEntityTooLarge
from werkzeug.utils import secure_filename

if getattr(sys, "frozen", False):
    frozen_root = Path(getattr(sys, "_MEIPASS", Path(sys.executable).resolve().parent))
    mac_resources = Path(sys.executable).resolve().parent.parent / "Resources"
    BASE_DIR = next((path for path in (frozen_root, mac_resources, Path(sys.executable).resolve().parent)
                     if (path / "index.html").is_file()), frozen_root)
else:
    BASE_DIR = Path(__file__).resolve().parent
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 12 * 1024 * 1024
app.config["JSON_AS_ASCII"] = False
ALLOWED = {"csv", "tsv", "xlsx", "xls", "txt", "json"}


def validate_archive(raw, filename):
    if filename.rsplit(".", 1)[-1].lower() not in {"xlsx", "docx"}:
        return
    try:
        with zipfile.ZipFile(io.BytesIO(raw)) as archive:
            entries = archive.infolist()
            total = sum(item.file_size for item in entries)
            if len(entries) > 5000 or total > 100 * 1024 * 1024:
                raise ValueError("File bahut badi ya complex hai. 100 MB se chhoti, standard Excel/Word file try karein.")
            if any(item.file_size > 1024 * 1024 and item.compress_size and item.file_size / item.compress_size > 100 for item in entries):
                raise ValueError("File ka compression pattern safe limits se zyada hai. File ko dobara export karke try karein.")
    except zipfile.BadZipFile as exc:
        raise ValueError("Excel/Word file khul nahi paayi. File ko dobara save karke try karein.") from exc


def safe_sheet_text(value):
    if isinstance(value, str) and value.lstrip("\ufeff\t\r\n ").startswith(("=", "+", "-", "@")):
        return "'" + value
    return value


def safe_error(exc, fallback):
    if isinstance(exc, ValueError):
        return str(exc)
    app.logger.exception("FormSaathi request failed")
    return fallback


@app.before_request
def restrict_local_requests():
    host = (urlparse("//" + request.host).hostname or "").lower()
    if host not in {"localhost", "127.0.0.1", "::1"}:
        return jsonify(error="FormSaathi sirf isi computer par khulta hai."), 403
    if request.method in {"POST", "PUT", "PATCH", "DELETE"}:
        origin = request.headers.get("Origin")
        referer = request.headers.get("Referer")
        source = urlparse(origin) if origin else (urlparse(referer) if referer else None)
        if not source or source.scheme != request.scheme or source.netloc.lower() != request.host.lower():
            return jsonify(error="Security check pass nahi hua. FormSaathi ko isi window mein dobara kholen."), 403


@app.after_request
def set_security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "no-referrer"
    response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; script-src 'self'; style-src 'self'; "
        "img-src 'self' data:; connect-src 'self'; object-src 'none'; base-uri 'self'; "
        "form-action 'self'; frame-ancestors 'none'"
    )
    if request.path.startswith("/api/"):
        response.headers["Cache-Control"] = "no-store"
    return response


@app.errorhandler(RequestEntityTooLarge)
def upload_too_large(_error):
    return jsonify(error="File 12 MB se badi hai. Chhoti file upload karein."), 413


@app.errorhandler(BadRequest)
def malformed_request(_error):
    return jsonify(error="Request format samajh nahi aaya. Dobara try karein."), 400


def rows_from_upload(file):
    name = secure_filename(file.filename or "")
    ext = name.rsplit(".", 1)[-1].lower() if "." in name else ""
    if ext not in ALLOWED:
        raise ValueError("CSV, TSV, XLSX, XLS, TXT, ya JSON file upload karein.")
    raw = file.read()
    validate_archive(raw, name)
    if ext in {"csv", "tsv", "txt"}:
        text = raw.decode("utf-8-sig", errors="replace")
        dialect = csv.excel_tab if ext == "tsv" else csv.Sniffer().sniff(text[:4096], delimiters=",;\t|")
        reader = csv.DictReader(io.StringIO(text), dialect=dialect)
        records = []
        for row in reader:
            records.append(row)
            if len(records) > 5000:
                raise ValueError("Ek baar mein adhiktam 5,000 data rows upload karein.")
        return records
    if ext == "json":
        import json
        data = json.loads(raw.decode("utf-8-sig"))
        if isinstance(data, dict):
            data = data.get("data", [data])
        if not isinstance(data, list):
            raise ValueError("JSON mein records ki list honi chahiye.")
        if len(data) > 5000:
            raise ValueError("Ek baar mein adhiktam 5,000 data rows upload karein.")
        return [dict(x) if isinstance(x, dict) else {"value": x} for x in data]
    try:
        import pandas as pd
    except ImportError:
        raise ValueError("Excel upload ke liye pandas aur openpyxl install karein.")
    frame = pd.read_excel(io.BytesIO(raw), nrows=5001)
    if len(frame) > 5000:
        raise ValueError("Ek baar mein adhiktam 5,000 data rows upload karein.")
    return frame.fillna("").to_dict(orient="records")


SALES_KEYS = {
    "product": ["product", "item", "productname", "itemname", "producttitle", "description", "sku"],
    "quantity": ["quantity", "qty", "unitssold", "sold", "units", "count"],
    "price": ["unitprice", "price", "sellingprice", "saleprice", "rate", "priceperunit"],
    "amount": ["total", "totalamount", "sales", "revenue", "amount", "linetotal", "netamount"],
    "total_cost": ["totalcost", "costtotal", "totalpurchase", "purchaseamount"],
    "cost": ["unitcost", "costperunit", "costprice", "buyingprice", "purchaseprice"],
    "date": ["date", "saledate", "transactiondate", "orderdate", "datetime"],
}


def norm(value):
    return re.sub(r"[^a-z0-9]", "", str(value or "").lower())


def choose_column(columns, candidates):
    is_amount = candidates == SALES_KEYS["amount"]
    is_price = candidates == SALES_KEYS["price"]
    is_cost = candidates == SALES_KEYS["cost"]
    is_total_cost = candidates == SALES_KEYS["total_cost"]
    def suitable(column):
        name = norm(column)
        if is_amount and any(x in name for x in ("cost", "profit", "margin", "cogs")):
            return False
        if is_price and any(x in name for x in ("total", "amount", "cost", "sales", "revenue")):
            return False
        if is_cost and "total" in name:
            return False
        if is_total_cost and not any(x in name for x in ("totalcost", "costtotal", "totalpurchase", "purchaseamount")):
            return False
        return True
    cleaned = {norm(column): column for column in columns}
    for key in candidates:
        if norm(key) in cleaned and suitable(cleaned[norm(key)]):
            return cleaned[norm(key)]
    for column in columns:
        name = norm(column)
        if suitable(column) and any(norm(key) in name or name in norm(key) for key in candidates if len(norm(key)) > 3):
            return column
    return None


def number(value):
    if value is None:
        return None
    text = str(value).strip().replace(",", "")
    if not text:
        return None
    text = re.sub(r"[^0-9.\-]", "", text)
    try:
        return float(text)
    except ValueError:
        return None


def analyze_sales(records):
    if not isinstance(records, list) or not records:
        raise ValueError("Sales data ki kam se kam ek row chahiye.")
    records = records[:10000]
    columns = list(records[0].keys())
    chosen = {key: choose_column(columns, keys) for key, keys in SALES_KEYS.items()}
    if not chosen["product"]:
        raise ValueError("Product/Item naam ka column nahin mila. Column ka naam Product ya Item rakhein.")
    if not chosen["quantity"]:
        raise ValueError("Quantity/Qty ka column nahin mila. Column ka naam Quantity rakhein.")
    normalized, issues = [], []
    for line, raw in enumerate(records, 2):
        product = str(raw.get(chosen["product"], "")).strip()
        qty = number(raw.get(chosen["quantity"]))
        if not product or qty is None:
            issues.append({"row": line, "reason": "Product ya quantity missing/invalid"})
            continue
        price = number(raw.get(chosen["price"])) if chosen["price"] else None
        amount = number(raw.get(chosen["amount"])) if chosen["amount"] else None
        cost = number(raw.get(chosen["cost"])) if chosen["cost"] else None
        total_cost = number(raw.get(chosen["total_cost"])) if chosen["total_cost"] else None
        if amount is None and price is not None:
            amount = qty * price
        if price is None and amount is not None and qty:
            price = amount / qty
        date = str(raw.get(chosen["date"], "")).strip() if chosen["date"] else ""
        normalized.append({"date": date, "product": product, "quantity": qty, "unit_price": price,
                           "sales": amount, "unit_cost": cost, "total_cost": total_cost,
                           "profit": amount - total_cost if amount is not None and total_cost is not None else (amount - qty * cost if amount is not None and cost is not None else None)})
    if not normalized:
        raise ValueError("Kisi row mein valid Product aur Quantity dono nahin mile.")
    products = defaultdict(lambda: {"product": "", "quantity": 0, "sales": 0, "profit": 0, "profit_rows": 0, "units_priced": 0})
    for row in normalized:
        entry = products[row["product"].casefold()]
        entry["product"] = row["product"]
        entry["quantity"] += row["quantity"]
        if row["sales"] is not None:
            entry["sales"] += row["sales"]
            entry["units_priced"] += row["quantity"]
        if row["profit"] is not None:
            entry["profit"] += row["profit"]
            entry["profit_rows"] += 1
    product_rows = []
    for entry in products.values():
        units_priced = entry.pop("units_priced")
        profit_rows = entry.pop("profit_rows")
        entry["average_price"] = entry["sales"] / units_priced if units_priced else None
        entry["profit"] = entry["profit"] if profit_rows else None
        product_rows.append(entry)
    product_rows.sort(key=lambda row: row["sales"], reverse=True)
    has_sales = any(row["sales"] is not None for row in normalized)
    profit_values = [row["profit"] for row in normalized if row["profit"] is not None]
    monthly = defaultdict(lambda: {"month": "", "sales": 0, "units": 0, "transactions": 0, "has_sales": False})
    for row in normalized:
        try:
            month = datetime.fromisoformat(row["date"][:10]).strftime("%Y-%m")
        except (ValueError, TypeError):
            continue
        entry = monthly[month]
        entry["month"] = month
        entry["units"] += row["quantity"]
        entry["transactions"] += 1
        if row["sales"] is not None:
            entry["sales"] += row["sales"]
            entry["has_sales"] = True
    monthly_rows = [{**row, "sales": row["sales"] if row["has_sales"] else None} for _, row in sorted(monthly.items())]
    return {
        "columns": chosen,
        "rows": normalized,
        "products": product_rows,
        "monthly": monthly_rows,
        "issues": issues,
        "summary": {"transactions": len(normalized), "products": len(product_rows),
                    "units": sum(row["quantity"] for row in normalized),
                    "sales": sum(row["sales"] or 0 for row in normalized) if has_sales else None,
                    "profit": sum(profit_values) if profit_values else None,
                    "top_product": product_rows[0]["product"] if product_rows else ""},
    }


def resume_profile(file):
    filename = secure_filename(file.filename or "")
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    raw = file.read()
    if ext == "txt":
        text = raw.decode("utf-8-sig", errors="replace")
    elif ext == "pdf":
        from pypdf import PdfReader
        reader = PdfReader(io.BytesIO(raw))
        if len(reader.pages) > 100:
            raise ValueError("Resume PDF mein 100 se kam pages rakhein.")
        chunks, total = [], 0
        for page in reader.pages:
            chunk = page.extract_text() or ""
            total += len(chunk)
            if total > 1_000_000:
                raise ValueError("Resume ka text bahut lamba hai. Chhoti file try karein.")
            chunks.append(chunk)
        text = "\n".join(chunks)
    elif ext == "docx":
        validate_archive(raw, filename)
        from docx import Document
        text = "\n".join(p.text for p in Document(io.BytesIO(raw)).paragraphs)
    else:
        raise ValueError("Resume .pdf, .docx ya .txt format mein upload karein.")
    email = re.search(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", text, re.I)
    phone = re.search(r"(?:\+?\d[\d ()-]{7,}\d)", text)
    lines = [re.sub(r"\s+", " ", line).strip() for line in text.splitlines() if line.strip()]
    ignored = re.compile(r"resume|curriculum|vitae|email|phone|mobile|address|linkedin|github", re.I)
    name = next((line for line in lines[:8] if len(line) < 65 and not ignored.search(line) and "@" not in line), "")
    skills = []
    skill_match = re.search(r"(?:skills|technical skills)\s*[:\n]\s*([^\n]{1,300})", text, re.I)
    if skill_match:
        skills = [x.strip(" •,-") for x in re.split(r"[,|•]", skill_match.group(1)) if x.strip(" •,-")]
    profile = {"Full Name": name, "Email": email.group(0) if email else "", "Phone": phone.group(0).strip() if phone else "", "Skills": ", ".join(skills)}
    def section(pattern):
        match = re.search(r"(?:^|\n)\s*(?:" + pattern + r")\s*:?\s*\n(.*?)(?=\n\s*(?:education|experience|work experience|employment|skills|technical skills|projects|certifications|summary|profile|achievements)\b|\Z)", text, re.I | re.S)
        return re.sub(r"\s+", " ", match.group(1)).strip()[:500] if match else ""
    profile["Education"] = section("education|academic background|qualification")
    profile["Work Experience"] = section("work experience|professional experience|experience|employment history")
    profile = {k: v for k, v in profile.items() if v}
    return {k: v for k, v in profile.items() if v}, {"filename": filename, "text_length": len(text), "fields_found": sum(bool(v) for v in profile.values())}


@app.get("/")
def index():
    return send_from_directory(BASE_DIR, "index.html")


@app.get("/<path:asset>")
def local_asset(asset):
    if asset not in {"app.js", "styles.css"}:
        return jsonify(error="Page nahi mili."), 404
    return send_from_directory(BASE_DIR, asset)


@app.post("/api/upload")
def upload():
    uploaded = request.files.get("file")
    if not uploaded:
        return jsonify(error="File chunein."), 400
    try:
        records = rows_from_upload(uploaded)
        records = [{str(k): (v.isoformat() if hasattr(v, "isoformat") else str(v)) for k, v in row.items()} for row in records]
        if not records:
            raise ValueError("File mein koi data row nahin mili.")
        return jsonify(filename=secure_filename(uploaded.filename), columns=list(records[0].keys()), records=records[:5000], count=len(records))
    except Exception as exc:
        return jsonify(error=safe_error(exc, "File padh nahi paayi. File format aur size check karke dobara try karein.")), 400


@app.post("/api/sales/analyze")
def sales_analyze():
    try:
        data = request.get_json(force=True)
        if not isinstance(data, dict):
            raise ValueError("Sales data samajh nahi aaya. File dobara upload karein.")
        result = analyze_sales(data.get("records", []))
        return jsonify(result)
    except Exception as exc:
        return jsonify(error=safe_error(exc, "Sales data process nahi ho paayi. Columns check karke dobara try karein.")), 400


@app.post("/api/sales/workbook")
def sales_workbook():
    try:
        data = request.get_json(force=True)
        if not isinstance(data, dict):
            raise ValueError("Sales data samajh nahi aaya. File dobara upload karein.")
        result = analyze_sales(data.get("records", []))
        company_name = str(data.get("company_name", "")).strip()[:100]
        currency = str(data.get("currency", "INR"))[:3]
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment
        from openpyxl.utils import get_column_letter
        wb = Workbook()
        summary = wb.active
        summary.title = "Overview"
        summary.append([company_name or "SALES OVERVIEW"])
        summary.append(["Transactions", result["summary"]["transactions"]])
        summary.append(["Products", result["summary"]["products"]])
        summary.append(["Units sold", result["summary"]["units"]])
        summary.append(["Gross sales", result["summary"]["sales"]])
        summary.append(["Currency", currency])
        summary.append(["Gross profit (when cost supplied)", result["summary"]["profit"]])
        summary.append(["Top product by sales", result["summary"]["top_product"]])
        summary.append([])
        summary.append(["Source column mapping"])
        summary.append(["Business field", "Source column"])
        for key, value in result["columns"].items():
            summary.append([key.replace("_", " ").title(), value or "Not found"])
        for sheet_name, headers, rows in [
            ("Product Summary", ["Product", "Units sold", "Sales", "Average unit price", "Profit"],
             [[p["product"], p["quantity"], p["sales"], p["average_price"], p["profit"]] for p in result["products"]]),
            ("Sales Detail", ["Date", "Product", "Quantity", "Unit price", "Sales", "Unit cost", "Profit"],
             [[r["date"], r["product"], r["quantity"], r["unit_price"], r["sales"], r["unit_cost"] if r["unit_cost"] is not None else r["total_cost"], r["profit"]] for r in result["rows"]]),
            ("Monthly Trend", ["Month", "Transactions", "Units sold", "Sales"],
             [[m["month"], m["transactions"], m["units"], m["sales"]] for m in result["monthly"]]),
        ]:
            ws = wb.create_sheet(sheet_name)
            ws.append(headers)
            for row in rows:
                ws.append(row)
            if sheet_name == "Sales Detail" and result["issues"]:
                ws.append([])
                ws.append(["Rows skipped", len(result["issues"])])
        for ws in wb.worksheets:
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for cell in ws[1]:
                cell.font = Font(color="FFFFFF", bold=True)
                cell.fill = PatternFill("solid", fgColor="167553")
            for column in ws.columns:
                letter = get_column_letter(column[0].column)
                width = min(max(max(len(str(c.value or "")) for c in column) + 2, 14), 42)
                ws.column_dimensions[letter].width = width
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    cell.value = safe_sheet_text(cell.value)
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = '#,##0.00;[Red]-#,##0.00'
            for row in ws.iter_rows(min_row=1, max_row=1):
                for cell in row:
                    cell.value = safe_sheet_text(cell.value)
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name="company-sales-report.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as exc:
        return jsonify(error=safe_error(exc, "Sales workbook nahi ban paayi. Data check karke dobara try karein.")), 400


@app.post("/api/resume")
def resume_upload():
    uploaded = request.files.get("file")
    if not uploaded:
        return jsonify(error="Resume file chunein."), 400
    try:
        profile, meta = resume_profile(uploaded)
        return jsonify(profile=profile, meta=meta)
    except Exception as exc:
        return jsonify(error=safe_error(exc, "Resume nahi padh paaya. PDF, DOCX ya TXT file check karein.")), 400


@app.post("/api/chat")
def assistant_chat():
    data = request.get_json(force=True)
    if not isinstance(data, dict):
        return jsonify(reply="Sawaal dobara bhejein; mujhe text samajh nahi aaya."), 400
    message = str(data.get("message", "")).strip()[:2000]
    context = data.get("context")
    context = context if isinstance(context, dict) else {}
    if not isinstance(context.get("sales"), dict):
        context.pop("sales", None)
    q = message.lower()
    if not message:
        return jsonify(reply="Apna sawaal likhiye. Main sales sheet, resume aur form fields mein madad karunga.")
    if any(word in q for word in ["sales", "bikri", "sell", "revenue", "product", "excel", "sheet", "report"]):
        if context.get("sales"):
            summary = context["sales"]
            sales = summary.get("sales")
            reply = f"Aapki sheet mein {summary.get('products', 0)} products aur {summary.get('units', 0):g} units dikhe."
            if sales is not None:
                reply += f" Kul sales {sales:,.2f} hain."
            if summary.get("profit") is not None:
                reply += f" Available cost ke hisaab se profit {summary['profit']:,.2f} hai."
            else:
                reply += " Profit nikalne ke liye Unit Cost/Cost Price column bhi chahiye."
            return jsonify(reply=reply)
        return jsonify(reply="Sales report banane ke liye Excel/CSV upload karein. Product aur Quantity columns zaroori hain; Unit Price ya Total se sales nikalunga, aur Cost Price mile to profit bhi.")
    if any(word in q for word in ["resume", "job", "naukri", "application", "career", "experience", "skills"]):
        return jsonify(reply="Resume PDF/DOCX/TXT upload karein aur job form ka address kholen. Main naam, email, phone aur skills jaise pehchane gaye fields map karunga; baaki field ka label samajh na aaye to uska text yahan bhejein, main arth samjha kar aapse sahi jawab poochhunga. Har jawab check karke form aap submit karein.")
    definitions = [("cover letter", "Cover letter ek chhota parichay hai jo batata hai ki aap is role ke liye kyun apply kar rahe hain."),
                   ("notice period", "Notice period ka matlab hai current job chhodne ke baad kaam khatam hone tak ka samay. Agar aap job mein nahi hain to form ke options ke hisaab se 'Not applicable' ya 'Immediate' chun sakte hain."),
                   ("expected salary", "Expected salary woh vetan hai jo aap is role ke liye maang rahe hain. Apni sthiti ke hisaab se amount aur period (mahina/saal) confirm karein."),
                   ("current ctc", "Current CTC aapki maujooda company ka saalana total compensation hota hai. Payslip/offer letter se verify karein.")]
    for phrase, explanation in definitions:
        if phrase in q:
            return jsonify(reply=explanation + " Yeh detail resume se confirm na ho to guess na karein; aapse poochh kar hi bharein.")
    if context.get("sales"):
        return jsonify(reply=f"Sales sheet mein abhi {context['sales'].get('transactions', 0)} sales rows aur {context['sales'].get('products', 0)} products process hue hain. Aap 'top product', 'total sales' ya missing price/cost ke baare mein poochh sakte hain.")
    return jsonify(reply="Samajh gaya. Thoda aur batayein—yeh company sales sheet ke baare mein hai, job application ke baare mein, ya kisi form ke specific field ke baare mein? Main na samajh aaye field par guess nahi karunga.")


@app.post("/api/research")
def research_web():
    """Search the public web and return source-linked results for review."""
    data = request.get_json(force=True)
    if not isinstance(data, dict):
        return jsonify(error="Search request samajh nahi aayi."), 400
    query = str(data.get("query", "")).strip()[:300]
    if len(query) < 3:
        return jsonify(error="Search ke liye thoda aur detail likhein."), 400
    try:
        import httpx
        from bs4 import BeautifulSoup
        response = httpx.get("https://www.bing.com/search", params={"q": query, "count": 10},
                             headers={"User-Agent": "Mozilla/5.0 FormSaathi/1.0"}, timeout=20,
                             follow_redirects=True)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
        results, seen = [], set()
        for item in soup.select("li.b_algo"):
            heading = item.select_one("h2 a")
            href = heading.get("href", "") if heading else ""
            if href.startswith("https://www.bing.com/ck/"):
                token = parse_qs(urlparse(href).query).get("u", [""])[0]
                if token.startswith("a1"):
                    try:
                        href = base64.urlsafe_b64decode(token[2:] + "==").decode("utf-8", "ignore")
                    except Exception:
                        pass
            if not heading or not href.startswith(("http://", "https://")) or "bing.com/" in href:
                continue
            snippet_node = item.select_one(".b_caption p, p")
            snippet = " ".join((snippet_node or item).get_text(" ", strip=True).split())[:700]
            if href in seen:
                continue
            seen.add(href)
            results.append({"title": heading.get_text(" ", strip=True), "url": href, "snippet": snippet})
            if len(results) >= 10:
                break
        if not results:
            raise ValueError("Search results nahi mile. Search dobara karein ya website ka URL dein.")
        return jsonify(query=query, searched_at=datetime.now().astimezone().isoformat(timespec="seconds"),
                       sources=results)
    except Exception as exc:
        return jsonify(error=safe_error(exc, "Online search abhi nahi ho paayi. Internet check karke dobara try karein.")), 502


@app.post("/api/research/workbook")
def research_workbook():
    data = request.get_json(force=True)
    if not isinstance(data, dict):
        return jsonify(error="Source list samajh nahi aayi."), 400
    sources = data.get("sources") or []
    if not isinstance(sources, list) or not sources:
        return jsonify(error="Pehle online sources search karein."), 400
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill
        from openpyxl.utils import get_column_letter
        wb = Workbook()
        ws = wb.active
        ws.title = "Search Sources"
        ws.append(["Search query", str(data.get("query", ""))])
        ws.append(["Searched at", str(data.get("searched_at", ""))])
        ws.append(["Review note", "Search snippets are leads, not verified facts. Open each source and confirm date and claims."])
        ws.append([])
        ws.append(["Source title", "Snippet", "URL"])
        for item in sources[:50]:
            ws.append([str(item.get("title", ""))[:300], str(item.get("snippet", ""))[:1000], str(item.get("url", ""))[:1000]])
        for cell in ws[5]:
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = PatternFill("solid", fgColor="167553")
        for row in ws.iter_rows():
            for cell in row:
                cell.value = safe_sheet_text(cell.value)
        ws.freeze_panes = "A6"
        for col, width in enumerate([38, 90, 64], 1):
            ws.column_dimensions[get_column_letter(col)].width = width
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return send_file(output, as_attachment=True, download_name="formsaathi-online-research.xlsx",
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    except Exception as exc:
        return jsonify(error=safe_error(exc, "Excel source sheet nahi ban paayi. Dobara try karein.")), 400


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8000, debug=False)
