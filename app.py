# -*- coding: utf-8 -*-
"""
WarungQ - Aplikasi Kasir (POS) Flask + SQLite3
Fitur: admin & kasir, password + reset, print struk, export/import Excel,
barcode (scan kamera/USB & buat barcode untuk item homemade),
pembayaran tunai/debit/QRIS/e-wallet, laporan harian/mingguan/bulanan.
"""
import io
import os
import secrets
import sqlite3
import zipfile
from datetime import datetime, timedelta, date
from functools import wraps

from flask import (Flask, abort, flash, jsonify, redirect, render_template,
                   request, send_file, session, url_for)
from werkzeug.security import check_password_hash, generate_password_hash
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "warungq.db")
UPLOAD_DIR = os.path.join(BASE_DIR, "static", "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# ============================================================
# MODE DEMO
# True  = demo untuk kasir: admin dikunci, kasir boleh bertransaksi,
#         melihat laporan (baca-saja) dan menambah produk simulasi.
# False = aplikasi penuh seperti biasa (admin & kasir normal).
DEMO_MODE = True
# ============================================================

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 8 * 1024 * 1024  # 8 MB

# Kunci sesi disimpan di file agar login tidak reset setiap restart
_key_file = os.path.join(BASE_DIR, "secret.key")
if not os.path.exists(_key_file):
    with open(_key_file, "w") as f:
        f.write(secrets.token_hex(32))
with open(_key_file) as f:
    app.secret_key = f.read().strip()


# ---------------------------------------------------------------- database
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL DEFAULT 'kasir',
    security_question TEXT DEFAULT '',
    security_answer_hash TEXT DEFAULT ''
);
CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    barcode TEXT UNIQUE,
    category TEXT DEFAULT '',
    price INTEGER NOT NULL DEFAULT 0,
    stock INTEGER NOT NULL DEFAULT 0,
    min_stock INTEGER NOT NULL DEFAULT 5,
    is_homemade INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS transactions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT UNIQUE NOT NULL,
    created_at TEXT NOT NULL,
    user_id INTEGER NOT NULL REFERENCES users(id),
    payment_method TEXT NOT NULL,
    subtotal INTEGER NOT NULL DEFAULT 0,
    discount INTEGER NOT NULL DEFAULT 0,
    total INTEGER NOT NULL DEFAULT 0,
    paid INTEGER NOT NULL DEFAULT 0,
    change_amount INTEGER NOT NULL DEFAULT 0
);
CREATE TABLE IF NOT EXISTS transaction_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    transaction_id INTEGER NOT NULL REFERENCES transactions(id) ON DELETE CASCADE,
    product_id INTEGER,
    name TEXT NOT NULL,
    price INTEGER NOT NULL,
    qty INTEGER NOT NULL,
    subtotal INTEGER NOT NULL
);
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT DEFAULT ''
);
"""


def init_db():
    conn = get_db()
    conn.executescript(SCHEMA)
    # Akun awal: admin/admin123 dan kasir/kasir123 (segera diganti!)
    cur = conn.execute("SELECT COUNT(*) c FROM users").fetchone()
    if cur["c"] == 0:
        conn.execute(
            "INSERT INTO users(name, username, password_hash, role) VALUES (?,?,?,?)",
            ("Admin", "admin", generate_password_hash("admin123"), "admin"))
        conn.execute(
            "INSERT INTO users(name, username, password_hash, role) VALUES (?,?,?,?)",
            ("Kasir", "kasir", generate_password_hash("kasir123"), "kasir"))
    if DEMO_MODE:
        pc = conn.execute("SELECT COUNT(*) c FROM products").fetchone()["c"]
        if pc == 0:
            demo_products = [
                ("Teh Kotak 250ml", "8991002101234", "Minuman", 4000, 50, 5, 0),
                ("Kopi Sachet", "8991002101241", "Minuman", 2000, 100, 10, 0),
                ("Mie Instan Goreng", "8991002101258", "Makanan", 3500, 80, 10, 0),
                ("Gula Pasir 1kg", "8991002101265", "Sembako", 15000, 30, 5, 0),
                ("Roti Homemade Coklat", None, "Makanan", 10000, 20, 5, 1),
                ("Keripik Pisang Homemade", None, "Makanan", 12000, 15, 5, 1),
            ]
            conn.executemany(
                "INSERT INTO products(name, barcode, category, price, stock, "
                "min_stock, is_homemade) VALUES (?,?,?,?,?,?,?)", demo_products)
    defaults = {
        "store_name": "WarungQ",
        "store_address": "",
        "store_wa": "",
        "receipt_footer": "Terima kasih atas kunjungan Anda!",
        "qris_image": "",
    }
    for k, v in defaults.items():
        conn.execute("INSERT OR IGNORE INTO settings(key, value) VALUES (?,?)", (k, v))
    conn.commit()
    conn.close()


def get_settings():
    conn = get_db()
    rows = conn.execute("SELECT key, value FROM settings").fetchall()
    conn.close()
    return {r["key"]: r["value"] for r in rows}


# ---------------------------------------------------------------- auth
def login_required(f):
    @wraps(f)
    def wrapper(*a, **kw):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return f(*a, **kw)
    return wrapper


def admin_required(demo_ok=False):
    """Khusus admin. Bila demo_ok=True, di mode DEMO kasir juga boleh."""
    def deco(f):
        @wraps(f)
        def wrapper(*a, **kw):
            if "user_id" not in session:
                return redirect(url_for("login"))
            if session.get("role") != "admin" and not (demo_ok and DEMO_MODE):
                abort(403)
            return f(*a, **kw)
        return wrapper
    return deco


def current_user():
    conn = get_db()
    u = conn.execute("SELECT * FROM users WHERE id=?", (session["user_id"],)).fetchone()
    conn.close()
    return u


@app.template_filter("rp")
def rp_filter(n):
    try:
        return "Rp" + f"{int(n or 0):,}".replace(",", ".")
    except (TypeError, ValueError):
        return "Rp0"


@app.context_processor
def inject_globals():
    return {"settings": get_settings(), "role": session.get("role", ""),
            "user_name": session.get("name", ""), "demo": DEMO_MODE}


# ---------------------------------------------------------------- beranda
@app.route("/")
@login_required
def dashboard():
    conn = get_db()
    today = date.today().isoformat()
    sales_today = conn.execute(
        "SELECT COALESCE(SUM(total),0) s, COUNT(*) c FROM transactions "
        "WHERE date(created_at)=?", (today,)).fetchone()
    n_products = conn.execute("SELECT COUNT(*) c FROM products").fetchone()["c"]
    low_stock = conn.execute(
        "SELECT name, stock, min_stock FROM products WHERE stock <= min_stock "
        "ORDER BY stock ASC LIMIT 10").fetchall()
    recent = conn.execute(
        "SELECT t.code, t.created_at, t.total, t.payment_method, u.name kasir "
        "FROM transactions t JOIN users u ON u.id=t.user_id "
        "ORDER BY t.id DESC LIMIT 8").fetchall()
    # grafik 7 hari
    chart = []
    for i in range(6, -1, -1):
        d = date.today() - timedelta(days=i)
        r = conn.execute("SELECT COALESCE(SUM(total),0) s FROM transactions "
                         "WHERE date(created_at)=?", (d.isoformat(),)).fetchone()
        chart.append({"label": d.strftime("%d/%m"), "value": r["s"]})
    conn.close()
    return render_template("dashboard.html", sales_today=sales_today,
                           n_products=n_products, low_stock=low_stock,
                           recent=recent, chart=chart)


# ---------------------------------------------------------------- login
@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        conn = get_db()
        u = conn.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
        conn.close()
        if u and check_password_hash(u["password_hash"], password):
            if DEMO_MODE and u["role"] == "admin":
                flash("Ini mode demo: hanya akun Kasir yang bisa masuk "
                      "(kasir / kasir123).", "error")
            else:
                session["user_id"] = u["id"]
                session["name"] = u["name"]
                session["role"] = u["role"]
                return redirect(url_for("dashboard"))
        flash("Username atau password salah.", "error")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/lupa-password", methods=["GET", "POST"])
def lupa_password():
    """Reset password lewat pertanyaan keamanan."""
    step = 1
    user = None
    if request.method == "POST":
        conn = get_db()
        try:
            step = int(request.form.get("step", 1))
            if step == 1:
                username = request.form.get("username", "").strip()
                user = conn.execute("SELECT * FROM users WHERE username=?",
                                    (username,)).fetchone()
                if not user or not user["security_question"]:
                    flash("Username tidak ditemukan atau pertanyaan keamanan belum "
                          "diatur. Minta admin untuk mereset password Anda.", "error")
                else:
                    step = 2
            elif step == 2:
                user = conn.execute("SELECT * FROM users WHERE username=?",
                                    (request.form.get("username", "").strip(),)).fetchone()
                answer = request.form.get("answer", "").strip().lower()
                new1 = request.form.get("new_password", "")
                new2 = request.form.get("new_password2", "")
                if not user or not user["security_answer_hash"] or \
                   not check_password_hash(user["security_answer_hash"], answer):
                    flash("Jawaban keamanan salah.", "error")
                    step = 2
                elif len(new1) < 6:
                    flash("Password baru minimal 6 karakter.", "error")
                    step = 2
                elif new1 != new2:
                    flash("Konfirmasi password tidak sama.", "error")
                    step = 2
                else:
                    conn.execute("UPDATE users SET password_hash=? WHERE id=?",
                                 (generate_password_hash(new1), user["id"]))
                    conn.commit()
                    flash("Password berhasil direset. Silakan login.", "ok")
                    return redirect(url_for("login"))
        finally:
            conn.close()
    return render_template("lupa_password.html", step=step,
                           user=dict(user) if user else None)


@app.route("/profil", methods=["GET", "POST"])
@login_required
def profil():
    u = current_user()
    if request.method == "POST":
        if DEMO_MODE:
            flash("Fitur ini dimatikan dalam mode demo.", "error")
            return redirect(url_for("profil"))
        act = request.form.get("act")
        conn = get_db()
        if act == "password":
            old = request.form.get("old_password", "")
            new1 = request.form.get("new_password", "")
            new2 = request.form.get("new_password2", "")
            if not check_password_hash(u["password_hash"], old):
                flash("Password lama salah.", "error")
            elif len(new1) < 6:
                flash("Password baru minimal 6 karakter.", "error")
            elif new1 != new2:
                flash("Konfirmasi password tidak sama.", "error")
            else:
                conn.execute("UPDATE users SET password_hash=? WHERE id=?",
                             (generate_password_hash(new1), u["id"]))
                conn.commit()
                flash("Password berhasil diganti.", "ok")
        elif act == "security":
            q = request.form.get("security_question", "").strip()
            a = request.form.get("security_answer", "").strip()
            if q and a:
                conn.execute("UPDATE users SET security_question=?, security_answer_hash=? "
                             "WHERE id=?", (q, generate_password_hash(a.lower()), u["id"]))
                conn.commit()
                flash("Pertanyaan keamanan disimpan.", "ok")
            else:
                flash("Pertanyaan dan jawaban wajib diisi.", "error")
        conn.close()
        return redirect(url_for("profil"))
    return render_template("profil.html", u=u)


# ---------------------------------------------------------------- POS / kasir
@app.route("/pos")
@login_required
def pos():
    return render_template("pos.html")


@app.route("/api/products")
@login_required
def api_products():
    q = request.args.get("q", "").strip().lower()
    conn = get_db()
    if q:
        rows = conn.execute(
            "SELECT * FROM products WHERE lower(name) LIKE ? OR barcode LIKE ? "
            "ORDER BY name LIMIT 100", (f"%{q}%", f"%{q}%")).fetchall()
    else:
        rows = conn.execute("SELECT * FROM products ORDER BY name LIMIT 100").fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/product/<int:pid>")
@login_required
def api_product(pid):
    conn = get_db()
    p = conn.execute("SELECT * FROM products WHERE id=?", (pid,)).fetchone()
    conn.close()
    if not p:
        return jsonify({"error": "not found"}), 404
    return jsonify(dict(p))


@app.route("/api/checkout", methods=["POST"])
@login_required
def api_checkout():
    data = request.get_json(force=True)
    items = data.get("items", [])
    method = data.get("payment_method", "Tunai")
    discount = max(0, int(data.get("discount", 0)))
    paid = max(0, int(data.get("paid", 0)))
    if not items:
        return jsonify({"ok": False, "error": "Keranjang kosong."}), 400
    if method not in ("Tunai", "Debit", "QRIS", "E-Wallet"):
        return jsonify({"ok": False, "error": "Metode pembayaran tidak valid."}), 400

    conn = get_db()
    try:
        subtotal = 0
        prepared = []
        for it in items:
            pid = int(it["id"])
            qty = int(it["qty"])
            if qty <= 0:
                continue
            p = conn.execute("SELECT * FROM products WHERE id=?", (pid,)).fetchone()
            if not p:
                raise ValueError(f"Produk id {pid} tidak ada.")
            if p["stock"] < qty:
                raise ValueError(f"Stok '{p['name']}' tinggal {p['stock']}.")
            subtotal += p["price"] * qty
            prepared.append((p, qty))
        if not prepared:
            raise ValueError("Tidak ada item valid.")
        if discount > subtotal:
            discount = subtotal
        total = subtotal - discount
        if method == "Tunai" and paid < total:
            raise ValueError("Uang bayar kurang dari total.")
        change = (paid - total) if method == "Tunai" else 0

        now = datetime.now()
        code = "TRX" + now.strftime("%Y%m%d%H%M%S") + f"{secrets.randbelow(900)+100}"
        cur = conn.execute(
            "INSERT INTO transactions(code, created_at, user_id, payment_method, "
            "subtotal, discount, total, paid, change_amount) VALUES (?,?,?,?,?,?,?,?,?)",
            (code, now.strftime("%Y-%m-%d %H:%M:%S"), session["user_id"], method,
             subtotal, discount, total, paid if method == "Tunai" else total, change))
        tid = cur.lastrowid
        for p, qty in prepared:
            conn.execute("INSERT INTO transaction_items(transaction_id, product_id, "
                         "name, price, qty, subtotal) VALUES (?,?,?,?,?,?)",
                         (tid, p["id"], p["name"], p["price"], qty, p["price"] * qty))
            conn.execute("UPDATE products SET stock=stock-? WHERE id=?", (qty, p["id"]))
        conn.commit()
        return jsonify({"ok": True, "transaction_id": tid, "code": code})
    except ValueError as e:
        conn.rollback()
        return jsonify({"ok": False, "error": str(e)}), 400
    finally:
        conn.close()


@app.route("/struk/<int:tid>")
@login_required
def struk(tid):
    conn = get_db()
    t = conn.execute("SELECT t.*, u.name kasir FROM transactions t "
                     "JOIN users u ON u.id=t.user_id WHERE t.id=?", (tid,)).fetchone()
    if not t:
        conn.close()
        abort(404)
    items = conn.execute("SELECT * FROM transaction_items WHERE transaction_id=?",
                         (tid,)).fetchall()
    conn.close()
    return render_template("struk.html", t=t, items=items)


# ---------------------------------------------------------------- produk
@app.route("/produk")
@admin_required(demo_ok=True)
def produk():
    conn = get_db()
    rows = conn.execute("SELECT * FROM products ORDER BY name").fetchall()
    conn.close()
    return render_template("produk.html", rows=[dict(r) for r in rows])


def _product_from_form(conn, product_id=None):
    name = request.form.get("name", "").strip()
    barcode = request.form.get("barcode", "").strip()
    category = request.form.get("category", "").strip()
    price = int(request.form.get("price", 0) or 0)
    stock = int(request.form.get("stock", 0) or 0)
    min_stock = int(request.form.get("min_stock", 5) or 5)
    homemade = 1 if request.form.get("is_homemade") == "on" else 0
    if not name:
        raise ValueError("Nama produk wajib diisi.")
    if price < 0 or stock < 0:
        raise ValueError("Harga dan stok tidak boleh negatif.")
    if barcode:
        dup = conn.execute("SELECT id FROM products WHERE barcode=? AND id IS NOT ?",
                           (barcode, product_id)).fetchone()
        if dup:
            raise ValueError("Barcode sudah dipakai produk lain.")
    return (name, barcode or None, category, price, stock, min_stock, homemade)


@app.route("/produk/tambah", methods=["POST"])
@admin_required(demo_ok=True)
def produk_tambah():
    conn = get_db()
    try:
        v = _product_from_form(conn)
        conn.execute("INSERT INTO products(name, barcode, category, price, stock, "
                     "min_stock, is_homemade) VALUES (?,?,?,?,?,?,?)", v)
        conn.commit()
        flash("Produk ditambahkan.", "ok")
    except ValueError as e:
        flash(str(e), "error")
    finally:
        conn.close()
    return redirect(url_for("produk"))


@app.route("/produk/ubah/<int:pid>", methods=["POST"])
@admin_required()
def produk_ubah(pid):
    conn = get_db()
    try:
        v = _product_from_form(conn, pid)
        conn.execute("UPDATE products SET name=?, barcode=?, category=?, price=?, "
                     "stock=?, min_stock=?, is_homemade=? WHERE id=?",
                     v + (pid,))
        conn.commit()
        flash("Produk diperbarui.", "ok")
    except ValueError as e:
        flash(str(e), "error")
    finally:
        conn.close()
    return redirect(url_for("produk"))


@app.route("/produk/hapus/<int:pid>", methods=["POST"])
@admin_required()
def produk_hapus(pid):
    conn = get_db()
    conn.execute("DELETE FROM products WHERE id=?", (pid,))
    conn.commit()
    conn.close()
    flash("Produk dihapus.", "ok")
    return redirect(url_for("produk"))


@app.route("/produk/barcode/<int:pid>")
@admin_required(demo_ok=True)
def produk_barcode(pid):
    """Halaman label barcode siap print."""
    conn = get_db()
    p = conn.execute("SELECT * FROM products WHERE id=?", (pid,)).fetchone()
    conn.close()
    if not p:
        abort(404)
    return render_template("barcode_label.html", p=p,
                           copies=int(request.args.get("copies", 24)))


# ---- Excel produk
@app.route("/produk/export.xlsx")
@admin_required()
def produk_export():
    conn = get_db()
    rows = conn.execute("SELECT name, barcode, category, price, stock, min_stock, "
                        "is_homemade FROM products ORDER BY name").fetchall()
    conn.close()
    wb = Workbook()
    ws = wb.active
    ws.title = "Produk"
    headers = ["Nama", "Barcode", "Kategori", "Harga", "Stok", "Stok Minimum", "Homemade"]
    ws.append(headers)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="D9EAD3")
    for r in rows:
        ws.append([r["name"], r["barcode"] or "", r["category"], r["price"],
                   r["stock"], r["min_stock"], "Ya" if r["is_homemade"] else "Tidak"])
    widths = [30, 18, 16, 12, 8, 13, 11]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    fname = f"produk_warungq_{date.today().isoformat()}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/produk/import", methods=["POST"])
@admin_required()
def produk_import():
    f = request.files.get("file")
    if not f or not f.filename:
        flash("Pilih file Excel (.xlsx) atau CSV terlebih dahulu.", "error")
        return redirect(url_for("produk"))
    ext = os.path.splitext(f.filename)[1].lower()
    if ext not in (".xlsx", ".xlsm", ".csv"):
        flash("Format tidak dikenali. Gunakan file .xlsx atau .csv. "
              "Jika file Anda .xls (Excel lama), buka di Excel lalu Save As "
              "dengan format 'Excel Workbook (.xlsx)'.", "error")
        return redirect(url_for("produk"))
    try:
        if ext == ".csv":
            rows_in = _read_csv(f)
        else:
            wb = load_workbook(f, data_only=True)
            ws = wb.active
            rows_in = list(ws.iter_rows(min_row=2, values_only=True))
        added, updated = _import_rows(rows_in)
        flash(f"Import selesai: {added} produk baru, {updated} diperbarui.", "ok")
    except zipfile.BadZipFile:
        flash("File bukan Excel .xlsx yang valid. Kemungkinan format Excel lama "
              "(.xls) atau CSV yang di-rename. Buka di Excel/LibreOffice lalu "
              "Save As 'Excel Workbook (.xlsx)', atau simpan sebagai .csv.", "error")
    except Exception as e:
        flash(f"Gagal import: {e}", "error")
    return redirect(url_for("produk"))


def _read_csv(f):
    """Baca CSV (koma, titik-koma, atau tab) menjadi list baris."""
    import csv
    sample = f.stream.read(4096).decode("utf-8-sig", errors="replace")
    f.stream.seek(0)
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel
    text = f.stream.read().decode("utf-8-sig", errors="replace")
    return list(csv.reader(io.StringIO(text), dialect))


def _import_rows(rows_in):
    conn = get_db()
    added, updated = 0, 0
    rows_in = list(rows_in)
    # Lewati baris judul bila ada (misal CSV: Nama,Barcode,Kategori,...)
    if rows_in:
        first = str(rows_in[0][0]).strip().lower() if rows_in[0] else ""
        if first in ("nama", "name", "produk", "product", "nama produk"):
            rows_in = rows_in[1:]
    for row in rows_in:
        if not row or not row[0]:
            continue
        name = str(row[0]).strip()
        barcode = str(row[1]).strip() if len(row) > 1 and row[1] is not None else ""
        category = str(row[2]).strip() if len(row) > 2 and row[2] is not None else ""
        price = _to_int(row[3]) if len(row) > 3 else 0
        stock = _to_int(row[4]) if len(row) > 4 else 0
        min_stock = _to_int(row[5]) if len(row) > 5 and row[5] is not None else 5
        homemade = 1 if (len(row) > 6 and str(row[6]).strip().lower()
                         in ("ya", "yes", "1", "true")) else 0
        existing = None
        if barcode:
            existing = conn.execute("SELECT id FROM products WHERE barcode=?",
                                    (barcode,)).fetchone()
        if existing:
            conn.execute("UPDATE products SET name=?, category=?, price=?, stock=?, "
                         "min_stock=?, is_homemade=? WHERE id=?",
                         (name, category, price, stock, min_stock, homemade,
                          existing["id"]))
            updated += 1
        else:
            conn.execute("INSERT INTO products(name, barcode, category, price, "
                         "stock, min_stock, is_homemade) VALUES (?,?,?,?,?,?,?)",
                         (name, barcode or None, category, price, stock,
                          min_stock, homemade))
            added += 1
    conn.commit()
    conn.close()
    return added, updated


def _to_int(v):
    """Ubah '15.000', '15000.0', 15000 menjadi bilangan bulat."""
    if v is None:
        return 0
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v).strip().replace("Rp", "").replace(".", "").replace(",", "")
    digits = "".join(ch for ch in s if ch.isdigit() or ch == "-")
    return int(digits) if digits else 0


# ---------------------------------------------------------------- transaksi
@app.route("/transaksi")
@login_required
def transaksi():
    q_from = request.args.get("from", "")
    q_to = request.args.get("to", "")
    q = request.args.get("q", "").strip()
    sql = ("SELECT t.*, u.name kasir, "
           "(SELECT COALESCE(SUM(qty),0) FROM transaction_items ti "
           " WHERE ti.transaction_id=t.id) item_count "
           "FROM transactions t JOIN users u ON u.id=t.user_id WHERE 1=1")
    args = []
    if q_from:
        sql += " AND date(t.created_at) >= ?"
        args.append(q_from)
    if q_to:
        sql += " AND date(t.created_at) <= ?"
        args.append(q_to)
    if q:
        sql += " AND (t.code LIKE ? OR u.name LIKE ?)"
        args += [f"%{q}%", f"%{q}%"]
    sql += " ORDER BY t.id DESC LIMIT 500"
    conn = get_db()
    rows = conn.execute(sql, args).fetchall()
    conn.close()
    return render_template("transaksi.html", rows=rows,
                           q_from=q_from, q_to=q_to, q=q)


@app.route("/transaksi/<int:tid>")
@login_required
def transaksi_detail(tid):
    conn = get_db()
    t = conn.execute("SELECT t.*, u.name kasir FROM transactions t "
                     "JOIN users u ON u.id=t.user_id WHERE t.id=?", (tid,)).fetchone()
    items = conn.execute("SELECT * FROM transaction_items WHERE transaction_id=?",
                         (tid,)).fetchall()
    conn.close()
    if not t:
        abort(404)
    return render_template("transaksi_detail.html", t=t, items=items)


# ---------------------------------------------------------------- laporan
def _report_data(d_from, d_to):
    conn = get_db()
    summary = conn.execute(
        "SELECT COALESCE(SUM(total),0) omzet, COUNT(*) ntrx, "
        "COALESCE(SUM((SELECT SUM(qty) FROM transaction_items ti WHERE ti.transaction_id=t.id)),0) qty "
        "FROM transactions t WHERE date(created_at) BETWEEN ? AND ?",
        (d_from, d_to)).fetchone()
    per_day = conn.execute(
        "SELECT date(created_at) d, COUNT(*) ntrx, SUM(total) omzet "
        "FROM transactions WHERE date(created_at) BETWEEN ? AND ? "
        "GROUP BY date(created_at) ORDER BY d", (d_from, d_to)).fetchall()
    per_method = conn.execute(
        "SELECT payment_method, COUNT(*) ntrx, SUM(total) omzet FROM transactions "
        "WHERE date(created_at) BETWEEN ? AND ? GROUP BY payment_method "
        "ORDER BY omzet DESC", (d_from, d_to)).fetchall()
    top = conn.execute(
        "SELECT ti.name, SUM(ti.qty) qty, SUM(ti.subtotal) omzet "
        "FROM transaction_items ti JOIN transactions t ON t.id=ti.transaction_id "
        "WHERE date(t.created_at) BETWEEN ? AND ? "
        "GROUP BY ti.name ORDER BY qty DESC LIMIT 10", (d_from, d_to)).fetchall()
    per_kasir = conn.execute(
        "SELECT u.name kasir, COUNT(*) ntrx, SUM(t.total) omzet FROM transactions t "
        "JOIN users u ON u.id=t.user_id WHERE date(t.created_at) BETWEEN ? AND ? "
        "GROUP BY u.name ORDER BY omzet DESC", (d_from, d_to)).fetchall()
    conn.close()
    return summary, per_day, per_method, top, per_kasir


@app.route("/laporan")
@admin_required(demo_ok=True)
def laporan():
    preset = request.args.get("preset", "bulan")
    today = date.today()
    if preset == "hari":
        d_from = d_to = today
    elif preset == "minggu":  # Senin s.d. hari ini
        d_from = today - timedelta(days=today.weekday())
        d_to = today
    elif preset == "bulan":
        d_from = today.replace(day=1)
        d_to = today
    elif preset == "bulan_lalu":
        first_this = today.replace(day=1)
        d_to = first_this - timedelta(days=1)
        d_from = d_to.replace(day=1)
    else:
        try:
            d_from = date.fromisoformat(request.args.get("from") or today.isoformat())
            d_to = date.fromisoformat(request.args.get("to") or today.isoformat())
        except ValueError:
            d_from = d_to = today
        preset = "custom"
    summary, per_day, per_method, top, per_kasir = _report_data(
        d_from.isoformat(), d_to.isoformat())
    return render_template("laporan.html", preset=preset,
                           d_from=d_from.isoformat(), d_to=d_to.isoformat(),
                           summary=summary, per_day=per_day,
                           per_method=per_method, top=top, per_kasir=per_kasir)


@app.route("/laporan/export.xlsx")
@admin_required(demo_ok=True)
def laporan_export():
    d_from = request.args.get("from", date.today().isoformat())
    d_to = request.args.get("to", date.today().isoformat())
    summary, per_day, per_method, top, per_kasir = _report_data(d_from, d_to)
    wb = Workbook()

    ws = wb.active
    ws.title = "Ringkasan"
    ws.append(["Laporan Penjualan WarungQ"])
    ws.append(["Periode", d_from, "s.d.", d_to])
    ws.append([])
    ws.append(["Total Omzet", summary["omzet"]])
    ws.append(["Jumlah Transaksi", summary["ntrx"]])
    ws.append(["Item Terjual", summary["qty"]])
    ws["A1"].font = Font(bold=True, size=14)

    def sheet(title, headers, rows, widths):
        w = wb.create_sheet(title)
        w.append(headers)
        for c in w[1]:
            c.font = Font(bold=True)
            c.fill = PatternFill("solid", fgColor="D9EAD3")
        for r in rows:
            w.append(list(r))
        for i, wd in enumerate(widths, 1):
            w.column_dimensions[get_column_letter(i)].width = wd

    sheet("Per Hari", ["Tanggal", "Transaksi", "Omzet"],
          [(r["d"], r["ntrx"], r["omzet"]) for r in per_day], [14, 12, 14])
    sheet("Per Metode", ["Metode", "Transaksi", "Omzet"],
          [(r["payment_method"], r["ntrx"], r["omzet"]) for r in per_method], [14, 12, 14])
    sheet("Terlaris", ["Produk", "Qty", "Omzet"],
          [(r["name"], r["qty"], r["omzet"]) for r in top], [30, 8, 14])
    sheet("Per Kasir", ["Kasir", "Transaksi", "Omzet"],
          [(r["kasir"], r["ntrx"], r["omzet"]) for r in per_kasir], [20, 12, 14])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(buf, as_attachment=True,
                     download_name=f"laporan_{d_from}_{d_to}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------- pengguna
@app.route("/pengguna")
@admin_required()
def pengguna():
    conn = get_db()
    rows = conn.execute("SELECT id, name, username, role FROM users ORDER BY id").fetchall()
    conn.close()
    return render_template("pengguna.html", rows=rows)


@app.route("/pengguna/tambah", methods=["POST"])
@admin_required()
def pengguna_tambah():
    name = request.form.get("name", "").strip()
    username = request.form.get("username", "").strip().lower()
    password = request.form.get("password", "")
    role = request.form.get("role", "kasir")
    if role not in ("admin", "kasir"):
        role = "kasir"
    if not name or not username or len(password) < 6:
        flash("Nama, username wajib diisi dan password minimal 6 karakter.", "error")
        return redirect(url_for("pengguna"))
    conn = get_db()
    try:
        conn.execute("INSERT INTO users(name, username, password_hash, role) "
                     "VALUES (?,?,?,?)",
                     (name, username, generate_password_hash(password), role))
        conn.commit()
        flash("Pengguna ditambahkan.", "ok")
    except sqlite3.IntegrityError:
        flash("Username sudah dipakai.", "error")
    finally:
        conn.close()
    return redirect(url_for("pengguna"))


@app.route("/pengguna/reset/<int:uid>", methods=["POST"])
@admin_required()
def pengguna_reset(uid):
    new = request.form.get("new_password", "")
    if len(new) < 6:
        flash("Password baru minimal 6 karakter.", "error")
    else:
        conn = get_db()
        conn.execute("UPDATE users SET password_hash=? WHERE id=?",
                     (generate_password_hash(new), uid))
        conn.commit()
        conn.close()
        flash("Password pengguna direset.", "ok")
    return redirect(url_for("pengguna"))


@app.route("/pengguna/hapus/<int:uid>", methods=["POST"])
@admin_required()
def pengguna_hapus(uid):
    if uid == session["user_id"]:
        flash("Tidak bisa menghapus akun sendiri.", "error")
        return redirect(url_for("pengguna"))
    conn = get_db()
    n = conn.execute("SELECT COUNT(*) c FROM users WHERE role='admin'").fetchone()["c"]
    target = conn.execute("SELECT role FROM users WHERE id=?", (uid,)).fetchone()
    if target and target["role"] == "admin" and n <= 1:
        flash("Minimal harus ada satu admin.", "error")
    else:
        conn.execute("DELETE FROM users WHERE id=?", (uid,))
        conn.commit()
        flash("Pengguna dihapus.", "ok")
    conn.close()
    return redirect(url_for("pengguna"))


# ---------------------------------------------------------------- pengaturan
@app.route("/pengaturan", methods=["GET", "POST"])
@admin_required()
def pengaturan():
    if request.method == "POST":
        conn = get_db()
        for k in ("store_name", "store_address", "store_wa", "receipt_footer"):
            conn.execute("UPDATE settings SET value=? WHERE key=?",
                         (request.form.get(k, "").strip(), k))
        f = request.files.get("qris_file")
        if f and f.filename:
            ext = os.path.splitext(f.filename)[1].lower()
            if ext in (".png", ".jpg", ".jpeg", ".webp"):
                fname = "qris" + ext
                f.save(os.path.join(UPLOAD_DIR, fname))
                conn.execute("UPDATE settings SET value=? WHERE key=?",
                             (fname, "qris_image"))
        conn.commit()
        conn.close()
        flash("Pengaturan disimpan.", "ok")
        return redirect(url_for("pengaturan"))
    return render_template("pengaturan.html", s=get_settings())


# ---------------------------------------------------------------- main
init_db()

# Untuk PythonAnywhere: arahkan "WSGI configuration file" di tab Web
# langsung ke file ini, misal: /home/warungq/warungq-flask/app.py
application = app

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
