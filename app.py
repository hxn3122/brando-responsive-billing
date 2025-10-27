from flask import Flask, render_template, request, send_file, redirect, url_for, flash, session
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import Table, TableStyle, Paragraph, SimpleDocTemplate, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
import datetime, os, json, re, functools, csv

app = Flask(__name__)
app.secret_key = "replace-this-with-a-random-secret"

COMPANY_NAME = "BRANDO"
BASE_DIR = os.path.dirname(__file__)
DATA_DIR = os.path.join(BASE_DIR, "invoices")
UPLOAD_DIR = os.path.join(BASE_DIR, "static", "uploads")
DEFAULT_LOGO_PATH = os.path.join(UPLOAD_DIR, "logo.png")

LOADSHEETS_DIR = os.path.join(DATA_DIR, "loadsheets")
def loadsheets_index_path(username):
    return os.path.join(DATA_DIR, f"loadsheets_{username}.json")

def load_loadsheets(username):
    p = loadsheets_index_path(username)
    if not os.path.exists(p):
        return {"items": []}
    with open(p, "r", encoding="utf-8") as f:
        return json.load(f)

def save_loadsheets(username, data):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(loadsheets_index_path(username), "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

def prune_old_loadsheets(username, days=7):
    # Delete entries older than N days and remove files
    data = load_loadsheets(username)
    now = datetime.datetime.now()
    keep = []
    for item in data["items"]:
        try:
            created = datetime.datetime.strptime(item["created_at"], "%Y-%m-%d %H:%M:%S")
        except Exception:
            created = now
        if (now - created).days > days:
            # remove files
            for k in ("pdf_path", "csv_path", "xlsx_path"):
                fp = item.get(k)
                if fp and os.path.exists(fp):
                    try: os.remove(fp)
                    except: pass
        else:
            keep.append(item)
    data["items"] = keep
    save_loadsheets(username, data)

def export_history_csv(username, dest_path):
    hist = load_history(username)["items"]
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    with open(dest_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Invoice #", "Customer", "Phone (Primary)", "Phone (Optional)", "Address", "Total (PKR)", "Created At"])
        for r in hist:
            w.writerow([r.get("invoice_no"), r.get("customer_name"), r.get("phone_primary"), r.get("phone_secondary"), r.get("customer_address"), f'{float(r.get("total",0)):.2f}', r.get("created_at")])

def export_history_xlsx(username, dest_path):
    hist = load_history(username)["items"]
    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "History"
    ws.append(["Invoice #", "Customer", "Phone (Primary)", "Phone (Optional)", "Address", "Total (PKR)", "Created At"])
    for r in hist:
        ws.append([r.get("invoice_no"), r.get("customer_name"), r.get("phone_primary"), r.get("phone_secondary"), r.get("customer_address"), float(r.get("total",0)), r.get("created_at")])
    wb.save(dest_path)

def selected_invoices(username, invoice_nos):
    # returns list of history rows for given invoice numbers (as strings)
    hist = load_history(username)["items"]
    chosen = []
    found_set = set(invoice_nos)
    for r in hist:
        if str(r.get("invoice_no")) in found_set:
            chosen.append(r)
    return chosen

def any_invoice_already_in_loadsheet(username, invoice_nos):
    ls = load_loadsheets(username)["items"]
    s = set(invoice_nos)
    for item in ls:
        if s.intersection(set(item.get("invoice_nos", []))):
            return True
    return False


def generate_loadsheet_files(username, invoice_rows, ls_code=None):
    # Create files in LOADSHEETS_DIR for this user; returns dict with paths and id
    os.makedirs(LOADSHEETS_DIR, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = f"{username}_loadsheet_{stamp}"
    # CSV
    csv_path = os.path.join(LOADSHEETS_DIR, base_name + ".csv")
    with open(csv_path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["Invoice #", "Customer", "Phone", "Address", "Total (PKR)"])
        per_day = {}
        for r in invoice_rows:
            created = (r.get("created_at") or "")[:10]
            per_day.setdefault(created, 0.0)
            per_day[created] += float(r.get("total",0))
            w.writerow([r.get("invoice_no"), r.get("customer_name"), r.get("phone_primary"), r.get("customer_address"), f'{float(r.get("total",0)):.2f}'])
        w.writerow([]); w.writerow(["Daily Totals"])
        for d, s in sorted(per_day.items()):
            w.writerow([d, f"{s:.2f}"])
    # XLSX
    xlsx_path = os.path.join(LOADSHEETS_DIR, base_name + ".xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "LoadSheet"
    ws.append(["Invoice #", "Customer", "Phone", "Address", "Total (PKR)"])
    total_sum = 0.0
    per_day = {}
    for r in invoice_rows:
        val = float(r.get("total",0))
        total_sum += val
        created = (r.get("created_at") or "")[:10]
        per_day.setdefault(created, 0.0)
        per_day[created] += val
        ws.append([r.get("invoice_no"), r.get("customer_name"), r.get("phone_primary"), r.get("customer_address"), val])
    ws.append([]); ws.append(["", "", "Grand Total", total_sum, ""])
    ws.append([]); ws.append(["Daily Totals"])
    for d, s in sorted(per_day.items()):
        ws.append([d, s])
    wb.save(xlsx_path)
    # PDF (6 columns: Invoice, Customer, Phone, Address, Total, Created)
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=24, leftMargin=24, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()
    title_txt = f"Load Sheet {'['+ls_code+'] ' if ls_code else ''}— {username}"
    title = Paragraph(f"<para align='center'><b>{title_txt}</b></para>", styles['Title'])
    dt = Paragraph(f"<para align='center'><font size=9>Generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</font></para>", styles['Normal'])
    story = [title, dt, Spacer(1, 12)]
    wrap = ParagraphStyle("wrap", parent=styles["Normal"], fontSize=9, leading=12, wordWrap="CJK")
    header = ParagraphStyle("wrapHeader", parent=wrap, textColor=colors.white, fontName="Helvetica-Bold")
    data = [[Paragraph("Invoice #", header), Paragraph("Customer", header), Paragraph("Phone", header), Paragraph("Address", header), Paragraph("Total (PKR)", header)]]
    total_sum = 0.0
    wrap = ParagraphStyle('wrap', parent=styles['Normal'], fontSize=9, leading=12, wordWrap='CJK')
    for r in invoice_rows:
        v = float(r.get('total',0))
        total_sum += v
        inv = Paragraph(str(r.get('invoice_no') or ''), wrap)
        cust = Paragraph((r.get('customer_name') or ''), wrap)
        phone = Paragraph((r.get('phone_primary') or ''), wrap)
        addr = Paragraph((r.get('customer_address') or ''), wrap)
        created = Paragraph((r.get('created_at') or ''), wrap)
        total_cell = Paragraph(f"{v:,.2f}", wrap)
        data.append([inv, cust, phone, addr, total_cell])
    data.append(["", "", "", Paragraph("<b>Grand Total</b>", styles['Normal']), Paragraph(f"<b>{total_sum:,.2f}</b>", styles['Normal'])])
    table = Table(data, colWidths=[18*mm, 42*mm, 26*mm, 82*mm, 18*mm])
    table.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("BACKGROUND", (0,0), (-1,0), colors.black),
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("ALIGN", (4,0), (4,-1), "RIGHT"),
        ("GRID", (0,0), (-1,-2), 0.25, colors.grey),
        ("LINEABOVE", (0,-1), (-1,-1), 0.75, colors.black),
        ("FONTNAME", (0,-1), (-1,-1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("BOTTOMPADDING", (0,0), (-1,0), 6),
        ("TOPPADDING", (0,0), (-1,0), 6),
    ]))
    story.append(table)
    # daily totals section
    per_day = {}
    for r in invoice_rows:
        d = (r.get("created_at") or "")[:10]
        per_day.setdefault(d, 0.0)
        per_day[d] += float(r.get("total",0))
    story.append(Spacer(1, 10))
    story.append(Paragraph("<b>Daily Totals</b>", styles['Normal']))
    for d in sorted(per_day.keys()):
        story.append(Paragraph(f"{d}: {per_day[d]:,.2f} PKR", styles['Normal']))
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"<font size=8>{COMPANY_NAME} — Load Sheet generated by BRANDO Billing</font>", styles['Normal']))
    doc.build(story)
    pdf_path = os.path.join(LOADSHEETS_DIR, base_name + ".pdf")
    with open(pdf_path, "wb") as f:
        f.write(buf.getbuffer())
    return {"id": base_name, "csv_path": csv_path, "xlsx_path": xlsx_path, "pdf_path": pdf_path}

# Users and history live under DATA_DIR
USERS_PATH = os.path.join(DATA_DIR, "users.json")

def load_users():
    os.makedirs(DATA_DIR, exist_ok=True)
    if not os.path.exists(USERS_PATH):
        # bootstrap default admin
        admin = {
            "name": "Administrator",
            "username": "admin",
            "password_hash": generate_password_hash("admin123"),
            "next_number": 1000,   # first invoice will be 1000
            "is_admin": True,
            "is_active": True
        }
        with open(USERS_PATH, "w", encoding="utf-8") as f:
            json.dump({"users":[admin]}, f, indent=2)
    with open(USERS_PATH, "r", encoding="utf-8") as f:
        return json.load(f)

def save_users(data):
    with open(USERS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

def get_current_user():
    uname = session.get("user")
    if not uname: return None
    users = load_users()["users"]
    for u in users:
        if u["username"] == uname:
            return u
    return None

def login_required(view):
    @functools.wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("user"):
            flash("Please log in first.", "error")
            return redirect(url_for("login"))
        return view(*args, **kwargs)
    return wrapped

def admin_required(view):
    @functools.wraps(view)
    def wrapped(*args, **kwargs):
        u = get_current_user()
        if not u or not u.get("is_admin"):
            flash("Admin access required.", "error")
            return redirect(url_for("index"))
        return view(*args, **kwargs)
    return wrapped

def human_now():
    return datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

def safe_float(x):
    try: return float(x)
    except: return 0.0

def per_user_history_path(username):
    return os.path.join(DATA_DIR, f"history_{username}.json")

def load_history(username):
    path = per_user_history_path(username)
    if not os.path.exists(path):
        return {"items": []}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)

def save_history(username, data):
    path = per_user_history_path(username)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)

def next_invoice_number_for_user(username, manual=None):
    data = load_users()
    for u in data["users"]:
        if u["username"] == username:
            if manual:
                return str(manual)
            # Use and increment this user's next_number
            num = int(u.get("next_number", 1000))
            u["next_number"] = num + 1
            save_users(data)
            return str(num)
    # Fallback if user not found
    return str(manual or 1000)

def make_invoice_pdf(company_name, invoice, items, logo_path=None, currency="PKR"):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=24, leftMargin=24, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('title', parent=styles['Title'], alignment=TA_CENTER, fontSize=20, leading=24, spaceAfter=6)
    tiny = ParagraphStyle('tiny', parent=styles['Normal'], fontSize=8)
    story = []
    if logo_path and os.path.exists(logo_path):
        try:
            im = Image(logo_path, width=40*mm, height=15*mm, kind='proportional')
            story.append(im)
        except Exception:
            pass
    story.append(Paragraph(f"<b>{company_name}</b>", title_style))
    story.append(Paragraph("Official Bill / Tax Invoice", styles['Normal']))
    story.append(Spacer(1, 6))
    meta = [
        ["Invoice #", invoice["invoice_no"]],
        ["Date/Time", invoice["date"]],
        ["Customer", invoice["customer_name"] or "-"],
        ["Phone (Primary)", invoice.get("phone_primary") or "-"],
        ["Phone (Optional)", invoice.get("phone_secondary") or "-"],
        ["Address", (invoice["customer_address"] or "-").replace('\\n','<br/>').replace('\\r','')]
    ]
    meta_table = Table(meta, colWidths=[32*mm, 126*mm])
    meta_table.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]))
    story += [meta_table, Spacer(1, 10)]
    data = [["#", "Product", f"Price ({currency})"]]
    total = 0.0
    for idx, it in enumerate(items, start=1):
        price = safe_float(it.get("price", 0))
        total += price
        data.append([str(idx), (it.get("name") or "").strip(), f"{price:,.2f}"])
    data.append(["", "Total", f"{total:,.2f}"])
    table = Table(data, colWidths=[18*mm, 42*mm, 26*mm, 82*mm, 18*mm])
    table.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0B3D91")),
        ("ALIGN", (-1,0), (-1,-1), "RIGHT"),
        ("ALIGN", (0,0), (0,-1), "CENTER"),
        ("GRID", (0,0), (-1,-2), 0.25, colors.grey),
        ("LINEABOVE", (0,-1), (-1,-1), 0.75, colors.black),
        ("FONTNAME", (0,-1), (-1,-1), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 10),
        ("BOTTOMPADDING", (0,0), (-1,0), 8),
        ("TOPPADDING", (0,0), (-1,0), 8),
    ]))
    story.append(table)
    story.append(Spacer(1, 12))
    story.append(Paragraph("<font size=8>Thank you for your business.</font>", tiny))
    doc.build(story)
    buf.seek(0)
    return buf

# ---------------- Auth Routes ----------------
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = (request.form.get("username") or "").strip()
        password = (request.form.get("password") or "").strip()
        users = load_users()["users"]
        user = next((u for u in users if u["username"] == username), None)
        if not user or not check_password_hash(user["password_hash"], password):
            flash("Invalid username or password.", "error")
            return redirect(url_for("login"))
        if not user.get("is_active", True):
            flash("This account is disabled. Contact an admin.", "error")
            return redirect(url_for("login"))
        session["user"] = user["username"]
        flash(f"Welcome, {user['name']}!", "info")
        return redirect(url_for("index"))
    return render_template("login.html", company_name=COMPANY_NAME)

@app.route("/logout")
def logout():
    session.clear()
    flash("Logged out.", "info")
    return redirect(url_for("login"))

# ---------------- Admin ----------------
@app.route("/admin/users", methods=["GET", "POST"])
@login_required
@admin_required
def admin_users():
    data = load_users()
    if request.method == "POST":
        action = request.form.get("action")
        if action == "delete":
            del_username = (request.form.get("del_username") or "").strip()
            # prevent deleting yourself
            me = get_current_user()["username"]
            if del_username == me:
                flash("You cannot delete yourself while logged in.", "error")
                return redirect(url_for("admin_users"))
            # ensure at least one admin remains
            admins = [u for u in data["users"] if u.get("is_admin")]
            target = next((u for u in data["users"] if u["username"] == del_username), None)
            if not target:
                flash("User not found.", "error")
                return redirect(url_for("admin_users"))
            if target.get("is_admin") and len(admins) <= 1:
                flash("Cannot delete the last admin.", "error")
                return redirect(url_for("admin_users"))
            # delete user
            data["users"] = [u for u in data["users"] if u["username"] != del_username]
            save_users(data)
            flash("User removed.", "info")
            return redirect(url_for("admin_users"))
        else:
            name = (request.form.get("name") or "").strip()
            username = (request.form.get("username") or "").strip()
            password = (request.form.get("password") or "").strip()
            start_from = int((request.form.get("start_from") or "1000").strip())
            is_admin = True if request.form.get("is_admin") == "on" else False
            if not name or not username or not password:
                flash("All fields except 'is admin' are required.", "error")
                return redirect(url_for("admin_users"))
            # ensure unique username
            for u in data["users"]:
                if u["username"] == username:
                    flash("Username already exists.", "error")
                    return redirect(url_for("admin_users"))

        # ensure unique username (case-insensitive) and unique starting invoice number
        for u in data["users"]:
            if u["username"].lower() == username.lower():
                flash("Username already exists.", "error")
                return redirect(url_for("admin_users"))
            # Prevent same starting invoice for two different users
            try:
                existing_next = int(u.get("next_number", -1))
            except Exception:
                existing_next = -1
            if existing_next == start_from:
                flash("This invoice number is already assigned to another user. Please choose a different one.", "error")
                return redirect(url_for("admin_users"))
            rec = {
                "name": name,
                "username": username,
                "password_hash": generate_password_hash(password),
                "next_number": start_from,
                "is_admin": is_admin
            }
            data["users"].append(rec)
            save_users(data)
            flash("User added.", "info")
            return redirect(url_for("admin_users"))
    return render_template("admin_users.html", users=load_users()["users"], company_name=COMPANY_NAME)

# ---------------- Core App ----------------
@app.route("/", methods=["GET"])
@login_required
def index():
    user = get_current_user()
    has_logo = os.path.exists(DEFAULT_LOGO_PATH)
    history = load_history(user["username"])
    return render_template("index.html", company_name=COMPANY_NAME, history_count=len(history["items"]), has_logo=has_logo)

@app.route("/history", methods=["GET"])
@login_required
def history():
    user = get_current_user()
    items = load_history(user["username"])["items"]
    # Filters
    q = (request.args.get("q") or "").strip().lower()
    sd = (request.args.get("start_date") or "").strip()
    ed = (request.args.get("end_date") or "").strip()
    def in_range(dt_str):
        if (not sd and not ed):
            return True
        try:
            dt = datetime.datetime.strptime(dt_str, "%Y-%m-%d %H:%M:%S")
        except Exception:
            return True
        ok = True
        if sd:
            try:
                sdt = datetime.datetime.strptime(sd + " 00:00:00", "%Y-%m-%d %H:%M:%S")
                ok = ok and (dt >= sdt)
            except: pass
        if ed:
            try:
                edt = datetime.datetime.strptime(ed + " 23:59:59", "%Y-%m-%d %H:%M:%S")
                ok = ok and (dt <= edt)
            except: pass
        return ok
    filtered = []
    for r in items:
        if q:
            blob = f"{r.get('invoice_no','')} {r.get('customer_name','')} {r.get('customer_address','')} {r.get('phone_primary','')} {r.get('phone_secondary','')}".lower()
            if q not in blob:
                continue
        if not in_range(r.get('created_at','')):
            continue
        filtered.append(r)
    filtered = sorted(filtered, key=lambda x: x.get('created_at',''), reverse=True)
    return render_template("history.html", items=filtered, company_name=COMPANY_NAME, q=q, start_date=sd, end_date=ed)

@app.route("/invoice/<invoice_no>", methods=["GET"])
@login_required
def serve_invoice(invoice_no):
    user = get_current_user()
    pdf_path = os.path.join(DATA_DIR, f"{user['username']}_{invoice_no}.pdf")
    if not os.path.exists(pdf_path):
        flash("Invoice not found", "error")
        return redirect(url_for("index"))
    resp = send_file(pdf_path, mimetype="application/pdf", as_attachment=False, download_name=f"{invoice_no}.pdf")
    resp.headers["Content-Disposition"] = f'inline; filename="{invoice_no}.pdf"'
    resp.headers["Cache-Control"] = "no-store"
    return resp

@app.route("/viewer/<invoice_no>", methods=["GET"])
@login_required
def viewer(invoice_no):
    return render_template("viewer.html", invoice_no=invoice_no, company_name=COMPANY_NAME)

@app.route("/generate", methods=["POST"])
@login_required
def generate():
    user = get_current_user()
    cust_name = (request.form.get("customer_name") or "").strip()
    cust_addr = (request.form.get("customer_address") or "").strip()
    manual_no = (request.form.get("invoice_no") or "").strip()
    phone_primary = (request.form.get("phone_primary") or "").strip()
    phone_secondary = (request.form.get("phone_secondary") or "").strip()

    # Validate primary phone (11 digits starting with 03)
    if not re.fullmatch(r"03[0-9]{9}", phone_primary):
        flash("Primary phone must be exactly 11 digits and start with 03 (e.g., 03XXXXXXXXX).", "error")
        return redirect(url_for("index"))

    logo_path = DEFAULT_LOGO_PATH if os.path.exists(DEFAULT_LOGO_PATH) else None

    names = request.form.getlist("name[]")
    prices = request.form.getlist("price[]")
    items = []
    for n, p in zip(names, prices):
        if (n or "").strip()=="" and (p or "").strip()=="":
            continue
        items.append({"name": n, "price": p})

    if not items:
        flash("Please add at least one product.", "error")
        return redirect(url_for('index'))

    invoice_no = next_invoice_number_for_user(user["username"], manual_no if manual_no else None)
    meta = {
        "invoice_no": invoice_no,
        "customer_name": cust_name,
        "customer_address": cust_addr,
        "phone_primary": phone_primary,
        "phone_secondary": phone_secondary,
        "date": human_now()
    }

    os.makedirs(DATA_DIR, exist_ok=True)
    pdf_io = make_invoice_pdf(COMPANY_NAME, meta, items, logo_path=logo_path)
    pdf_path = os.path.join(DATA_DIR, f"{user['username']}_{invoice_no}.pdf")
    with open(pdf_path, "wb") as f:
        f.write(pdf_io.getbuffer())

    hist = load_history(user["username"])
    total = sum([safe_float(p) for p in prices])
    hist["items"].append({
        "invoice_no": invoice_no,
        "customer_name": cust_name,
        "customer_address": cust_addr,
        "phone_primary": phone_primary,
        "phone_secondary": phone_secondary,
        "total": total,
        "created_at": meta["date"]
    })
    save_history(user["username"], hist)

    share = request.args.get("share")
    if share:
        return redirect(url_for("viewer", invoice_no=invoice_no, share="1"))
    return redirect(url_for("viewer", invoice_no=invoice_no))

@app.route("/admin/user/<username>", methods=["GET", "POST"])
@login_required
@admin_required
def admin_user_edit(username):
    data = load_users()
    target = next((u for u in data["users"] if u["username"] == username), None)
    if not target:
        flash("User not found.", "error")
        return redirect(url_for("admin_users"))
    if request.method == "POST":
        # Update fields
        target["name"] = (request.form.get("name") or target["name"]).strip()
        target["next_number"] = int((request.form.get("start_from") or target.get("next_number", 1000)))
        target["is_admin"] = True if request.form.get("is_admin") == "on" else False
        target["is_active"] = True if request.form.get("is_active") == "on" else False

        # Reset password if provided
        new_pw = (request.form.get("new_password") or "").strip()
        if new_pw:
            target["password_hash"] = generate_password_hash(new_pw)

        # Ensure you cannot demote/delete the last admin
        admins = [u for u in data["users"] if u.get("is_admin")]
        # If turning off admin for the only admin
        if target.get("is_admin") and len(admins) == 0:
            flash("There must be at least one admin.", "error")
            target["is_admin"] = True

        save_users(data)
        flash("User updated.", "info")
        return redirect(url_for("admin_users"))
    return render_template("admin_user_edit.html", u=target, company_name=COMPANY_NAME)

@app.route("/history/export")
@login_required
def history_export():
    user = get_current_user()
    fmt = (request.args.get("format") or "csv").lower()
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    fname = f"{user['username']}_history_{stamp}"
    if fmt == "xlsx":
        path = os.path.join(DATA_DIR, fname + ".xlsx")
        export_history_xlsx(user["username"], path)
        return send_file(path, as_attachment=True, download_name=fname + ".xlsx")
    else:
        path = os.path.join(DATA_DIR, fname + ".csv")
        export_history_csv(user["username"], path)
        return send_file(path, as_attachment=True, download_name=fname + ".csv")

@app.route("/loadsheets")
@login_required
def loadsheets():
    user = get_current_user()
    prune_old_loadsheets(user["username"])
    data = load_loadsheets(user["username"])
    items = sorted(data["items"], key=lambda x: x.get("created_at",""), reverse=True)
    return render_template("loadsheets.html", items=items, company_name=COMPANY_NAME)

@app.route("/loadsheets/generate", methods=["POST"])
@login_required
def loadsheets_generate():
    user = get_current_user()
    invoice_nos = request.form.getlist("invoice_no")
    if not invoice_nos:
        flash("Please select at least one invoice for the load sheet.", "error")
        return redirect(url_for("history"))
    # Check if any of these invoices already included in an existing loadsheet
    if any_invoice_already_in_loadsheet(user["username"], invoice_nos):
        flash("One or more selected invoices are already included in a previous load sheet. Use the load sheet history instead.", "error")
        return redirect(url_for("history"))
    rows = selected_invoices(user["username"], invoice_nos)
    if not rows:
        flash("Selected invoices not found.", "error")
        return redirect(url_for("history"))
    files = generate_loadsheet_files(user["username"], rows)
    # Record
    rec = {
        "id": files["id"],
        "invoice_nos": [str(x.get("invoice_no")) for x in rows],
        "created_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "pdf_path": files["pdf_path"],
        "csv_path": files["csv_path"],
        "xlsx_path": files["xlsx_path"]
    }
    data = load_loadsheets(user["username"])
    data["items"].append(rec)
    save_loadsheets(user["username"], data)
    flash("Load sheet generated and saved to history.", "info")
    return redirect(url_for("loadsheets"))

@app.route("/loadsheets/<ls_id>/<fmt>")
@login_required
def loadsheets_download(ls_id, fmt):
    user = get_current_user()
    data = load_loadsheets(user["username"])
    item = next((x for x in data["items"] if x["id"] == ls_id), None)
    if not item:
        flash("Load sheet not found.", "error")
        return redirect(url_for("loadsheets"))
    fmt = fmt.lower()
    if fmt == "pdf":
        return send_file(item["pdf_path"], as_attachment=True, download_name=ls_id + ".pdf")
    if fmt == "csv":
        return send_file(item["csv_path"], as_attachment=True, download_name=ls_id + ".csv")
    if fmt == "xlsx":
        return send_file(item["xlsx_path"], as_attachment=True, download_name=ls_id + ".xlsx")
    flash("Unknown format.", "error")
    return redirect(url_for("loadsheets"))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
