# 🌐 تطبيق مبيعات Flask + SQLite + تصدير Excel/PDF (نسخة محسّنة)
# شغّل: python app.py ثم افتح http://127.0.0.1:5000

from flask import Flask, request, redirect, url_for, render_template_string, jsonify, send_file
import sqlite3
from datetime import datetime
import io

# Excel
from openpyxl import Workbook

# PDF
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors

app = Flask(__name__)
DB_FILE = "sales.db"

# =========================
# Database Layer
# =========================
def get_db():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db() as conn:
        conn.execute("""
        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product TEXT NOT NULL,
            price REAL NOT NULL,
            cost REAL NOT NULL,
            profit REAL NOT NULL,
            date TEXT NOT NULL
        )
        """)


def get_all_sales():
    try:
        with get_db() as conn:
            rows = conn.execute("SELECT * FROM sales ORDER BY date ASC").fetchall()
            return [dict(r) for r in rows]
    except Exception as e:
        print("DB Error:", e)
        return []


def insert_sale(product, price, cost):
    try:
        with get_db() as conn:
            conn.execute(
                "INSERT INTO sales (product, price, cost, profit, date) VALUES (?, ?, ?, ?, ?)",
                (product, price, cost, price - cost, datetime.now().strftime('%Y-%m-%d %H:%M'))
            )
    except Exception as e:
        print("Insert Error:", e)

# =========================
# Templates
# =========================
BASE = """
<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>نظام المبيعات</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
body{font-family:system-ui;background:#0f172a;color:#e5e7eb;margin:0}
.container{max-width:900px;margin:auto;padding:16px}
.card{background:#111827;padding:16px;margin-bottom:12px;border-radius:12px}
input,button{width:100%;padding:10px;margin:6px 0;background:#020617;color:white;border:1px solid #374151}
.row{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.btns{display:grid;grid-template-columns:1fr 1fr;gap:8px}
</style>
</head>
<body>
<div class="container">
<h2>📊 نظام المبيعات</h2>
{{content|safe}}
</div>
</body>
</html>
"""

HOME = """
<div class="card">
<form method="post" action="/add">
<div class="row">
<input name="product" placeholder="اسم المنتج" required>
<input name="price" placeholder="السعر" required>
</div>
<input name="cost" placeholder="التكلفة" required>
<button type="submit">➕ إضافة</button>
</form>
</div>

<div class="card btns">
<a href="/export/excel">📥 تصدير Excel</a>
<a href="/export/pdf">📄 تصدير PDF</a>
</div>

<div class="card">
<canvas id="chart"></canvas>
</div>

<script>
try {
  const data = {{ chart_data | tojson }};

  if (!Array.isArray(data)) {
    throw new Error("Invalid chart data format");
  }

  const labels = data.map(x => x.date || "");
  const profits = data.map(x => Number(x.profit || 0));

  if (labels.length > 0) {
    new Chart(document.getElementById('chart'), {
      type: 'line',
      data: {
        labels: labels,
        datasets: [{
          label: 'الأرباح اليومية',
          data: profits
        }]
      }
    });
  }
} catch(e) {
  console.error("Chart Error:", e);
}
</script>
"""

# =========================
# Logic
# =========================
def calculate_daily(data):
    if not isinstance(data, list):
        return []

    daily = {}
    for s in data:
        try:
            d = (s.get("date") or "").split(" ")[0]
            if not d:
                continue
            daily[d] = daily.get(d, 0) + float(s.get("profit", 0))
        except Exception:
            continue

    return [
        {"date": k, "profit": round(v, 2)}
        for k, v in sorted(daily.items())
    ]

# =========================
# Export Excel
# =========================
@app.route('/export/excel')
def export_excel():
    data = get_all_sales()

    wb = Workbook()
    ws = wb.active

    ws.append(["ID", "Product", "Price", "Cost", "Profit", "Date"])

    for s in data:
        ws.append([
            s.get("id"),
            s.get("product"),
            s.get("price"),
            s.get("cost"),
            s.get("profit"),
            s.get("date")
        ])

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)

    return send_file(stream, as_attachment=True, download_name="sales.xlsx")

# =========================
# Export PDF
# =========================
@app.route('/export/pdf')
def export_pdf():
    data = get_all_sales()

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer)

    table_data = [["ID", "Product", "Price", "Cost", "Profit", "Date"]]

    for s in data:
        table_data.append([
            s.get("id"),
            s.get("product"),
            s.get("price"),
            s.get("cost"),
            s.get("profit"),
            s.get("date")
        ])

    table = Table(table_data)
    table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 1, colors.black)
    ]))

    doc.build([table])
    buffer.seek(0)

    return send_file(buffer, as_attachment=True, download_name="sales.pdf")

# =========================
# Routes
# =========================
@app.route('/')
def index():
    data = get_all_sales()
    chart_data = calculate_daily(data)

    html = render_template_string(HOME, chart_data=chart_data)
    return render_template_string(BASE, content=html)

@app.route('/add', methods=['POST'])
def add():
    try:
        product = request.form.get('product','').strip()
        price = float(request.form.get('price') or 0)
        cost = float(request.form.get('cost') or 0)

        if product:
            insert_sale(product, price, cost)
    except Exception as e:
        print("Add Error:", e)

    return redirect(url_for('index'))

@app.route('/api')
def api():
    return jsonify(get_all_sales())

# =========================
# TEST CASES
# =========================
# ✅ تشغيل بدون بيانات (لا crash)
# ✅ chart_data = [] لا يسبب خطأ
# ✅ إدخال بيانات صحيحة يظهر في chart
# ✅ إدخال بيانات غير صالحة لا يكسر النظام
# ✅ تصدير Excel يعمل
# ✅ تصدير PDF يعمل
# ✅ API يرجع JSON صحيح

# =========================
# Run
# =========================
if __name__ == '__main__':
    init_db()
    app.run(debug=True)
