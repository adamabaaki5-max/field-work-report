"""
FIELD WORK REPORT SYSTEM  –  Flask version for Vercel
-----------------------------------------------------
* Database : PostgreSQL (Neon via the Vercel Marketplace) – read from DATABASE_URL / POSTGRES_URL.
             When no DATABASE_URL is set and the app runs on your own computer, a local SQLite
             file (local_dev.db) is used so you can test before deploying.
* Excel    : built on the fly with openpyxl (Vercel has no permanent disk, so nothing is stored as files).
* No login : as requested.
"""
import io
import os
import re
import sqlite3
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal

from flask import (Flask, abort, flash, g, redirect, render_template, request,
                   send_file, session, url_for)
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# CONFIG
# ----------------------------------------------------------------------
DATABASE_URL = (os.environ.get("DATABASE_URL")
                or os.environ.get("POSTGRES_URL")
                or os.environ.get("POSTGRES_URL_NON_POOLING") or "")
IS_PG = DATABASE_URL.startswith(("postgres://", "postgresql://"))
ON_VERCEL = bool(os.environ.get("VERCEL"))
LOCAL_DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "local_dev.db")

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-this-secret-in-vercel-settings")
app.permanent_session_lifetime = timedelta(days=30)
app.config.update(SESSION_COOKIE_SAMESITE="Lax", SESSION_COOKIE_SECURE=ON_VERCEL,
                  SESSION_COOKIE_HTTPONLY=True)

# ----------------------------------------------------------------------
# CATEGORY DEFINITIONS  (key, label, input type)
# ----------------------------------------------------------------------
CATS = {
    "tax": {
        "table": "tax_reports",
        "title": "Report on Tax Payment",
        "sheet": "Tax Payments",
        "fields": [
            ("taxpayer_name",      "Name of Taxpayer",                     "text"),
            ("tin",                "TIN",                                  "text"),
            ("nature_of_business", "Nature of Business",                   "text"),
            ("phone",              "Phone Number",                         "text"),
            ("tax_charged",        "Tax Charged (GH₵)",                    "money"),
            ("tax_paid",           "Tax Paid (GH₵)",                       "money"),
            ("outstanding_amount", "Outstanding / Remaining Amount (GH₵)", "money"),
            ("remarks",            "Remarks",                              "longtext"),
        ],
    },
    "invite": {
        "table": "invitation_reports",
        "title": "Report on Invitation Letters",
        "sheet": "Invitation Letters",
        "fields": [
            ("taxpayer_name",      "Name of Taxpayer",               "text"),
            ("tin",                "TIN",                            "text"),
            ("phone",              "Telephone / Phone Number",       "text"),
            ("document_type",      "Type of Document",               "fixed"),
            ("person_served",      "Name / Status of Person Served", "text"),
            ("nature_of_business", "Nature of Business",             "text"),
            ("served_date",        "Date Served",                    "date"),
            ("served_time",        "Time Served",                    "time"),
            ("report_date",        "Date to Report at Office",       "date"),
            ("remarks",            "Remarks",                        "longtext"),
        ],
    },
}
FIXED = {"document_type": "INVITATION LETTER"}

# TIN formats accepted: Ghana Card (GHA-123456789-0)  or  1 letter + 10 digits.
TIN_RE = re.compile(r"^([A-Z]\d{10}|GHA-\d{9}-\d)$")
TIN_HTML_PATTERN = r"[A-Za-z][0-9]{10}|[Gg][Hh][Aa]-?[0-9]{9}-?[0-9]"
TEXT_RE = re.compile(r"^[^\W\d_](?:[^\W_]|[ .,'’&()/\-])*$")
TIME_RE = re.compile(r"^([01]\d|2[0-3]):[0-5]\d$")


# ----------------------------------------------------------------------
# SMALL HELPERS
# ----------------------------------------------------------------------
def utc_now():
    return datetime.now(timezone.utc)      # Ghana is on UTC all year


def today_str():
    return utc_now().strftime("%Y-%m-%d")


def now_str():
    return utc_now().strftime("%Y-%m-%d %H:%M:%S")


def up(x):
    return str(x if x is not None else "").upper()


def valid_date(s):
    if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", s or ""):
        return False
    try:
        datetime.strptime(s, "%Y-%m-%d")
        return True
    except ValueError:
        return False


def normalise_tin(x):
    x = re.sub(r"\s+", "", x or "").upper()
    m = re.fullmatch(r"GHA-?(\d{9})-?(\d)", x)
    if m:
        x = f"GHA-{m.group(1)}-{m.group(2)}"
    return x


# ----------------------------------------------------------------------
# DATABASE LAYER  (PostgreSQL on Vercel, SQLite for local testing)
# ----------------------------------------------------------------------
class DBConfigError(Exception):
    pass


def get_db():
    if "db" not in g:
        if IS_PG:
            import psycopg
            from psycopg.rows import dict_row
            try:
                # prepare_threshold=None keeps it compatible with Neon's pooled connections
                g.db = psycopg.connect(DATABASE_URL, row_factory=dict_row,
                                       prepare_threshold=None, connect_timeout=10)
            except psycopg.OperationalError as ex:
                raise DBConfigError(f"Could not connect to the database: {ex}")
        elif ON_VERCEL:
            raise DBConfigError("No database is connected. In Vercel open Storage, add a Postgres "
                                "database (Neon) and connect it to this project, then redeploy.")
        else:
            g.db = sqlite3.connect(LOCAL_DB)
            g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(_exc):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def norm(v):
    if v is None:
        return ""
    if isinstance(v, Decimal):
        return float(v)
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, date):
        return v.isoformat()
    return v


def db_query(sql, params=()):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(sql.replace("?", "%s") if IS_PG else sql, params)
    rows = [dict(r) for r in cur.fetchall()]
    cur.close()
    return [{k: norm(v) for k, v in r.items()} for r in rows]


def db_exec(sql, params=()):
    conn = get_db()
    cur = conn.cursor()
    cur.execute(sql.replace("?", "%s") if IS_PG else sql, params)
    cur.close()
    conn.commit()


def integrity_errors():
    if IS_PG:
        import psycopg
        return (psycopg.IntegrityError,)
    return (sqlite3.IntegrityError,)


def sql_type(t):
    if t == "money":
        return "NUMERIC(15,2) NOT NULL DEFAULT 0" if IS_PG else "REAL NOT NULL DEFAULT 0"
    if t == "date":
        return "DATE" if IS_PG else "TEXT"
    return "TEXT"


_ready = False


def init_db():
    """Create tables (and add any missing columns) – runs once per server start."""
    global _ready
    if _ready:
        return
    conn = get_db()
    for cfg in CATS.values():
        cols = ", ".join(f'"{k}" {sql_type(t)}' for k, _l, t in cfg["fields"])
        pk = "SERIAL PRIMARY KEY" if IS_PG else "INTEGER PRIMARY KEY AUTOINCREMENT"
        dt = "DATE" if IS_PG else "TEXT"
        ts = "TIMESTAMP" if IS_PG else "TEXT"
        conn.execute(f"CREATE TABLE IF NOT EXISTS {cfg['table']} (id {pk}, field_date {dt} NOT NULL, "
                     f"{cols}, created_at {ts}, updated_at {ts}, UNIQUE (tin))")
        conn.execute(f"CREATE INDEX IF NOT EXISTS idx_{cfg['table']}_day ON {cfg['table']} (field_date)")
        # add columns that are missing from a table created by an older version
        if IS_PG:
            for k, _l, t in cfg["fields"]:
                conn.execute(f'ALTER TABLE {cfg["table"]} ADD COLUMN IF NOT EXISTS "{k}" {sql_type(t)}')
        else:
            have = {r[1] for r in conn.execute(f"PRAGMA table_info({cfg['table']})").fetchall()}
            for k, _l, t in cfg["fields"]:
                if k not in have:
                    conn.execute(f'ALTER TABLE {cfg["table"]} ADD COLUMN "{k}" {sql_type(t)}')
    conn.commit()
    _ready = True


@app.before_request
def _startup():
    if request.endpoint != "static":
        init_db()


@app.errorhandler(DBConfigError)
def db_problem(ex):
    return (f"<div style='font-family:Arial;max-width:640px;margin:40px auto;padding:20px;"
            f"border:1px solid #e3b4b4;background:#fde3e3;border-radius:8px'>"
            f"<h3>Database problem</h3><p>{ex}</p></div>"), 500


@app.after_request
def no_index(resp):
    resp.headers["X-Robots-Tag"] = "noindex, nofollow"
    resp.headers["Cache-Control"] = "no-store"
    return resp


# ----------------------------------------------------------------------
# RECORD FUNCTIONS
# ----------------------------------------------------------------------
def load_records(cat, tin=None, day=None, rec_id=None):
    sql = f"SELECT * FROM {CATS[cat]['table']} WHERE 1=1"
    p = []
    if tin:
        sql += " AND UPPER(tin) LIKE ?"
        p.append(f"%{tin.strip().upper()}%")
    if day:
        sql += " AND field_date = ?"
        p.append(day)
    if rec_id is not None:
        sql += " AND id = ?"
        p.append(rec_id)
    sql += " ORDER BY field_date ASC, id ASC"          # oldest first
    return db_query(sql, p)


def insert_record(cat, vals, day):
    cols = list(vals)
    names = ", ".join(f'"{c}"' for c in cols)
    marks = ", ".join(["?"] * (len(cols) + 3))
    now = now_str()
    db_exec(f"INSERT INTO {CATS[cat]['table']} ({names}, field_date, created_at, updated_at) VALUES ({marks})",
            [*vals.values(), day, now, now])


def update_record(cat, rec_id, vals):
    sets = ", ".join(f'"{k}" = ?' for k in vals)
    db_exec(f"UPDATE {CATS[cat]['table']} SET {sets}, updated_at = ? WHERE id = ?",
            [*vals.values(), now_str(), rec_id])


def delete_record(cat, rec_id):
    db_exec(f"DELETE FROM {CATS[cat]['table']} WHERE id = ?", [rec_id])


# ----------------------------------------------------------------------
# FORM HANDLING + VALIDATION
# ----------------------------------------------------------------------
def collect_post(cat, form):
    v = {}
    for k, _l, t in CATS[cat]["fields"]:
        x = FIXED[k] if t == "fixed" else (form.get(k) or "")
        if t == "longtext":
            x = x.replace("\r\n", "\n").strip()
        else:
            x = re.sub(r"\s+", " ", x).strip()
        if t != "money":
            x = up(x)                                   # everything in CAPITALS
        else:
            c = x.replace(",", "").replace(" ", "")
            if re.fullmatch(r"\d{1,12}(\.\d{1,2})?", c):
                x = f"{Decimal(c):.2f}"
        if k == "tin":
            x = normalise_tin(x)
        v[k] = x
    if cat == "tax":
        if v["tax_paid"] == "":                          # nothing paid yet
            v["tax_paid"] = "0.00"
        money = re.compile(r"^\d+\.\d{2}$")
        if v["outstanding_amount"] == "" and money.match(v["tax_charged"]) and money.match(v["tax_paid"]):
            d = int(Decimal(v["tax_charged"]) * 100) - int(Decimal(v["tax_paid"]) * 100)
            if d >= 0:
                v["outstanding_amount"] = f"{Decimal(d) / 100:.2f}"
    return v


def validate(cat, v, exclude_id=0):
    """Returns {field_key: error message}. Empty dict = valid."""
    e = {}
    for k, label, t in CATS[cat]["fields"]:
        val = v.get(k, "") or ""
        lbl = re.sub(r"\s*\(GH₵\)", "", label)

        if k in ("taxpayer_name", "person_served", "nature_of_business"):
            if val == "":
                e[k] = f"{lbl} is required."
            elif not 2 <= len(val) <= 100:
                e[k] = f"{lbl} must be between 2 and 100 characters."
            elif not TEXT_RE.match(val):
                e[k] = f"{lbl} must start with a letter and contain only letters, numbers, spaces and . , ' - & ( ) /"
        elif k == "tin":
            if val == "":
                e[k] = "TIN is required."
            elif not TIN_RE.match(val):
                e[k] = "TIN is invalid. Enter a valid Ghana Card number, e.g. GHA-123456789-0."
            else:                                         # a TIN may be recorded only once per report
                dup = db_query(f"SELECT taxpayer_name FROM {CATS[cat]['table']} "
                               f"WHERE UPPER(tin) = ? AND id <> ? LIMIT 1", (val, exclude_id))
                if dup:
                    e[k] = f"This TIN has already been entered ({dup[0]['taxpayer_name']}). A TIN can only be entered once."
        elif k == "phone":
            if val == "":
                e[k] = "Phone number is required."
            elif not re.fullmatch(r"\d{10}", val):
                e[k] = "Phone number must be exactly 10 digits – numbers only (no letters, spaces or symbols)."
        elif k == "remarks":
            if len(val) > 500:
                e[k] = "Remarks must not exceed 500 characters."
            elif re.search(r"[\x00-\x08\x0B\x0C\x0E-\x1F]", val):
                e[k] = "Remarks contain invalid characters."
        elif t == "money":
            if val == "":
                e[k] = f"{lbl} is required (enter 0 if none)."
            elif not re.fullmatch(r"\d{1,12}\.\d{2}", val):
                e[k] = f"{lbl} must be a valid amount – numbers only, up to 2 decimal places (e.g. 1500.00)."
        elif k == "report_date":
            if not valid_date(val):
                e[k] = f"{lbl} is required and must be a valid date."
            elif valid_date(v.get("served_date", "")) and val < v["served_date"]:
                e[k] = f"{lbl} cannot be earlier than the Date Served."
        elif t == "date":
            if not valid_date(val):
                e[k] = f"{lbl} is required and must be a valid date."
            elif val > today_str():
                e[k] = f"{lbl} cannot be in the future."
        elif t == "time":
            if not TIME_RE.match(val):
                e[k] = f"{lbl} is required and must be a valid time (HH:MM)."
        elif t == "fixed":
            if val != FIXED[k]:
                e[k] = f"{lbl} must be '{FIXED[k]}'."

    if cat == "tax" and not {"tax_charged", "tax_paid", "outstanding_amount"} & set(e):
        c, p, o = (int(Decimal(v[x]) * 100) for x in ("tax_charged", "tax_paid", "outstanding_amount"))
        if p > c:
            e["tax_paid"] = "Tax Paid cannot be more than Tax Charged."
        elif o != c - p:
            e["outstanding_amount"] = ("Outstanding amount must equal Tax Charged minus Tax Paid "
                                       f"(GH₵ {Decimal(c - p) / 100:,.2f}).")
    return e


def form_values(cat, rec):
    """Turn a database row into strings suitable for the form inputs."""
    out = {}
    for k, _l, t in CATS[cat]["fields"]:
        v = rec.get(k, "")
        out[k] = f"{float(v):.2f}" if t == "money" and v != "" else str(v)
    return out


# ----------------------------------------------------------------------
# EXCEL
# ----------------------------------------------------------------------
def sheet_data(cat, rows):
    cfg = CATS[cat]
    headers = ["No.", "Field Work Date"] + [l for _k, l, _t in cfg["fields"]] + ["Created At", "Last Updated"]
    out = []
    for i, r in enumerate(rows, start=1):                # numbering 1,2,3,...
        line = [i, r["field_date"]]
        for k, _l, t in cfg["fields"]:
            line.append(float(r[k] or 0) if t == "money" else up(r[k]))
        line += [r["created_at"], r["updated_at"]]
        out.append(line)
    return headers, out


def build_xlsx(sheets):
    """sheets = {sheet name: (headers, rows)} -> bytes"""
    wb = Workbook()
    wb.remove(wb.active)
    for name, (headers, rows) in sheets.items():
        ws = wb.create_sheet(name[:31])
        ws.append(headers)
        for r in rows:
            ws.append(r)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="16325C")
            c.alignment = Alignment(wrap_text=True, vertical="center")
        for i, h in enumerate(headers, start=1):
            width = max([len(str(h))] + [len(str(r[i - 1])) for r in rows]) if rows else len(str(h))
            ws.column_dimensions[get_column_letter(i)].width = min(max(width + 3, 12), 50)
        ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def send_xlsx(filename, sheets):
    return send_file(build_xlsx(sheets), as_attachment=True, download_name=filename,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ----------------------------------------------------------------------
# TEMPLATE CONTEXT
# ----------------------------------------------------------------------
app.jinja_env.globals.update(today_str=today_str, TIN_HTML_PATTERN=TIN_HTML_PATTERN, up=up)


@app.context_processor
def inject():
    day = session.get("field_day")
    label = datetime.strptime(day, "%Y-%m-%d").strftime("%a, %d %b %Y") if day else None
    return {"day": day, "day_label": label}


def need_day():
    return session.get("field_day")


# ----------------------------------------------------------------------
# ROUTES
# ----------------------------------------------------------------------
@app.route("/")
def index():
    if not need_day():
        return render_template("gate.html", active="", default_day=today_str())
    return redirect(url_for("report", cat="tax"))


@app.route("/set_day", methods=["POST"])
def set_day():
    d = request.form.get("field_day", "")
    if valid_date(d):
        session.permanent = True
        session["field_day"] = d
    return redirect(url_for("index"))


@app.route("/change_day")
def change_day():
    session.pop("field_day", None)
    return redirect(url_for("index"))


@app.route("/report/<cat>", methods=["GET", "POST"])
def report(cat):
    if cat not in CATS:
        abort(404)
    day = need_day()
    if not day:
        return redirect(url_for("index"))
    cfg = CATS[cat]
    tab = request.args.get("tab", "add")
    errors, form_vals = {}, None
    rec_id = request.args.get("id", type=int)
    pending_all = dict(session.get("pending", {}))

    if request.method == "POST":
        act = request.form.get("action", "")
        if act == "review":
            v = collect_post(cat, request.form)
            errors = validate(cat, v)
            if errors:
                form_vals = v
            else:
                pending_all[cat] = v
                session["pending"] = pending_all
        elif act == "back":                               # return to form with the entered data
            form_vals = pending_all.pop(cat, None)
            session["pending"] = pending_all
        elif act == "save":
            p = pending_all.pop(cat, None)
            session["pending"] = pending_all
            if p and validate(cat, p):
                flash("The record failed validation and was not saved. Please check the details.", "err")
            elif p:
                try:
                    insert_record(cat, p, day)
                    flash("Record saved successfully.", "ok")
                except integrity_errors():
                    get_db().rollback()
                    flash("The record could not be saved (this TIN may already exist).", "err")
            return redirect(url_for("report", cat=cat, tab="add"))
        elif act == "update":
            rec_id = request.form.get("id", type=int) or 0
            v = collect_post(cat, request.form)
            errors = validate(cat, v, rec_id)
            if not errors:
                try:
                    update_record(cat, rec_id, v)
                    flash("Record updated.", "ok")
                except integrity_errors():
                    get_db().rollback()
                    flash("The record could not be updated (this TIN may already exist).", "err")
                return redirect(url_for("report", cat=cat, tab="edit", id=rec_id))
            form_vals, tab = v, "edit"
        elif act == "delete":
            rid = request.form.get("id", type=int) or 0
            rec = (load_records(cat, rec_id=rid) or [None])[0]
            if rec:
                delete_record(cat, rid)
                flash(f"The record of {up(rec['taxpayer_name'])} ({up(rec['tin'])}) has been deleted.", "ok")
            else:
                flash("Record not found.", "err")
            return redirect(url_for("report", cat=cat, tab="edit", tin=request.args.get("tin") or None))

    ctx = dict(cat=cat, cfg=cfg, tab=tab, errors=errors, active=cat,
               pending=pending_all.get(cat), tin_q="", rows=[], rec=None, form_vals=form_vals)

    if tab == "add":
        ctx["defaults"] = form_vals or {"served_date": day, "served_time": utc_now().strftime("%H:%M")}
    elif tab == "view":
        tin_q = request.args.get("tin", "").strip()
        only = bool(request.args.get("onlyday"))
        ctx.update(tin_q=tin_q, only=only, rows=load_records(cat, tin_q or None, day if only else None))
    else:  # edit
        tin_q = request.args.get("tin", "").strip()
        ctx["tin_q"] = tin_q
        ctx["rows"] = load_records(cat, tin_q or None)
        if rec_id:
            rec = (load_records(cat, rec_id=rec_id) or [None])[0]
            ctx["rec"] = rec
            if rec:
                ctx["edit_vals"] = form_vals or form_values(cat, rec)
    return render_template("category.html", **ctx)


@app.route("/report/<cat>/download")
def download(cat):
    if cat not in CATS:
        abort(404)
    day = need_day()
    if not day:
        return redirect(url_for("index"))
    tin = request.args.get("tin", "").strip()
    only = bool(request.args.get("onlyday"))
    rows = load_records(cat, tin or None, day if only else None)
    return send_xlsx(f"{cat}_report_{utc_now():%Y%m%d_%H%M}.xlsx", {CATS[cat]["sheet"]: sheet_data(cat, rows)})


@app.route("/search")
def search():
    if not need_day():
        return redirect(url_for("index"))
    tin = request.args.get("tin", "").strip()
    results = []
    if tin:
        for cat, cfg in CATS.items():
            rows = load_records(cat, tin)
            if rows:
                results.append((cat, cfg, rows))
    return render_template("search.html", active="search", tin=tin, results=results)


@app.route("/search/download/<cat>")
def search_download(cat):
    if cat not in CATS or not need_day():
        abort(404)
    tin = request.args.get("tin", "").strip()
    safe = re.sub(r"\W", "", tin) or "all"
    return send_xlsx(f"{cat}_{safe}.xlsx", {CATS[cat]["sheet"]: sheet_data(cat, load_records(cat, tin))})


@app.route("/stats")
def stats():
    if not need_day():
        return redirect(url_for("index"))
    period = request.args.get("period", "daily")
    first = db_query("SELECT tin, MIN(field_date) AS d FROM ("
                     "SELECT tin, field_date FROM tax_reports UNION ALL "
                     "SELECT tin, field_date FROM invitation_reports) t GROUP BY tin")
    tax_n = db_query("SELECT COUNT(*) AS n FROM tax_reports")[0]["n"]
    inv_n = db_query("SELECT COUNT(*) AS n FROM invitation_reports")[0]["n"]
    counts = {}
    for r in first:                    # each taxpayer is counted once, on the first day recorded
        d = date.fromisoformat(str(r["d"])[:10])
        if period == "weekly":
            y, w, _ = d.isocalendar()
            key = f"{y}-W{w:02d}"
        elif period == "monthly":
            key = d.strftime("%Y-%m")
        elif period == "yearly":
            key = d.strftime("%Y")
        else:
            period, key = "daily", d.isoformat()
        counts[key] = counts.get(key, 0) + 1
    counts = dict(sorted(counts.items()))
    return render_template("stats.html", active="stats", period=period, counts=counts,
                           total=len(first), tax_n=tax_n, inv_n=inv_n,
                           max_count=max(counts.values()) if counts else 1)


@app.route("/daily")
def daily():
    day = need_day()
    if not day:
        return redirect(url_for("index"))
    days = {}
    for cat, cfg in CATS.items():
        for r in db_query(f"SELECT field_date AS d, COUNT(*) AS n FROM {cfg['table']} GROUP BY field_date"):
            days.setdefault(str(r["d"])[:10], {"tax": 0, "invite": 0})[cat] = r["n"]
    return render_template("daily.html", active="daily", days=sorted(days.items(), reverse=True))


@app.route("/daily/download/<d>")
def daily_download(d):
    if not need_day() or not valid_date(d):
        abort(404)
    sheets = {cfg["sheet"]: sheet_data(cat, load_records(cat, day=d)) for cat, cfg in CATS.items()}
    return send_xlsx(f"FieldWork_{d}.xlsx", sheets)


if __name__ == "__main__":
    app.run(debug=True, port=5000)
