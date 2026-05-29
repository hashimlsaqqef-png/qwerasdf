import sqlite3
import io
import csv
from flask import Flask, render_template_string, request, redirect, url_for, jsonify, make_response

app = Flask(__name__)

# 1. إنشاء وقراءة قاعدة البيانات (SQLite)
def init_db():
    conn = sqlite3.connect('sales.db')
    cursor = conn.cursor()
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS sales (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product TEXT NOT NULL,
            price REAL NOT NULL,
            cost REAL NOT NULL,
            profit REAL NOT NULL,
            date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    conn.commit()
    conn.close()

def insert_sale(product, price, cost):
    price_val = float(price) if price else 0.0
    cost_val = float(cost) if cost else 0.0
    profit = price_val - cost_val
    
    conn = sqlite3.connect('sales.db')
    cursor = conn.cursor()
    cursor.execute('INSERT INTO sales (product, price, cost, profit) VALUES (?, ?, ?, ?)',
                   (product, price_val, cost_val, profit))
    conn.commit()
    conn.close()

def get_all_sales():
    conn = sqlite3.connect('sales.db')
    cursor = conn.cursor()
    cursor.execute('SELECT id, product, price, cost, profit, date FROM sales ORDER BY id DESC')
    rows = cursor.fetchall()
    conn.close()
    
    sales_list = []
    for row in rows:
        sales_list.append({
            'id': row[0],
            'product': row[1],
            'price': row[2],
            'cost': row[3],
            'profit': row[4],
            'date': row[5]
        })
    return sales_list

# 2. واجهة المستخدم الرسومية (HTML + CSS + Chart.js)
HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <title>نظام إدارة المبيعات الأنيق</title>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f4f6f9; margin: 0; padding: 20px; color: #333; }
        .container { max-width: 1000px; margin: 0 auto; background: white; padding: 30px; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.05); }
        h1, h2 { color: #2c3e50; border-bottom: 2px solid #ecf0f1; padding-bottom: 10px; }
        form { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)) auto; gap: 15px; margin-bottom: 30px; background: #f8f9fa; padding: 20px; border-radius: 8px; align-items: end; }
        .form-group { display: flex; flex-direction: column; gap: 5px; }
        label { font-weight: bold; font-size: 14px; }
        input { padding: 10px; border: 1px solid #ccc; border-radius: 6px; font-size: 16px; }
        button { background-color: #27ae60; color: white; border: none; padding: 11px 20px; font-size: 16px; border-radius: 6px; cursor: pointer; font-weight: bold; transition: 0.2s; }
        button:hover { background-color: #219653; }
        .btn-export { background-color: #2980b9; margin-left: 10px; }
        .btn-export:hover { background-color: #2471a3; }
        .btn-pdf { background-color: #e74c3c; }
        .btn-pdf:hover { background-color: #c0392b; }
        .actions { margin-bottom: 20px; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; text-align: right; }
        th, td { padding: 12px 15px; border-bottom: 1px solid #ddd; }
        th { background-color: #34495e; color: white; }
        tr:hover { background-color: #f1f2f6; }
        .chart-container { max-width: 600px; margin: 30px auto; }
    </style>
</head>
<body>
    <div class="container">
        <h1>📊 لوحة تحكم المبيعات والأرباح</h1>
        
        <form action="/add" method="POST">
            <div class="form-group">
                <label>اسم المنتج</label>
                <input type="text" name="product" required placeholder="مثال: عطر لوريت">
            </div>
            <div class="form-group">
                <label>سعر البيع ($)</label>
                <input type="number" step="0.01" name="price" required placeholder="0.00">
            </div>
            <div class="form-group">
                <label>التكلفة ($)</label>
                <input type="number" step="0.01" name="cost" required placeholder="0.00">
            </div>
            <button type="submit">إضافة المنتج</button>
        </form>

        <div class="actions">
            <a href="/export/excel"><button class="btn-export">📥 تصدير Excel</button></a>
            <a href="/export/pdf" target="_blank"><button class="btn-pdf">🖨️ طباعة / حفظ PDF</button></a>
        </div>

        <div class="chart-container">
            <canvas id="salesChart"></canvas>
        </div>

        <h2>📋 سجل العمليات المدخلة</h2>
        <table>
            <thead>
                <tr>
                    <th>المعرف</th>
                    <th>اسم المنتج</th>
                    <th>سعر البيع</th>
                    <th>التكلفة</th>
                    <th>الربح الصافي</th>
                    <th>التاريخ</th>
                </tr>
            </thead>
            <tbody>
                {% for sale in sales %}
                <tr>
                    <td>{{ sale.id }}</td>
                    <td>{{ sale.product }}</td>
                    <td>${{ sale.price }}</td>
                    <td>${{ sale.cost }}</td>
                    <td style="color: {{ 'green' if sale.profit >= 0 else 'red' }}">${{ sale.profit }}</td>
                    <td>{{ sale.date }}</td>
                </tr>
                {% else %}
                <tr>
                    <td colspan="6" style="text-align: center; color: #7f8c8d;">لا توجد بيانات مدخلة حالياً. ابدأ بإضافة منتج!</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
    </div>

    <script>
        // جلب البيانات للرسم البياني وتجنب خطأ [object Object]
        const salesData = {{ sales | tojson }};
        
        if (salesData.length > 0) {
            // عكس المصفوفة ليعرض الرسم البياني من الأقدم إلى الأحدث
            const dataReversed = [...salesData].reverse();
            const labels = dataReversed.map(item => item.product);
            const profits = dataReversed.map(item => item.profit);

            const ctx = document.getElementById('salesChart').getContext('2d');
            new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: labels,
                    datasets: [{
                        label: 'صافي الربح لكل منتج ($)',
                        data: profits,
                        backgroundColor: 'rgba(39, 174, 96, 0.6)',
                        borderColor: 'rgba(39, 174, 96, 1)',
                        borderWidth: 1
                    }]
                },
                options: {
                    responsive: true,
                    scales: { y: { beginAtZero: true } }
                }
            });
        } else {
            // في حال عدم وجود بيانات، يظهر رسم فارغ لطيف دون كسر الصفحة
            const ctx = document.getElementById('salesChart').getContext('2d');
            ctx.font = "16px sans-serif";
            ctx.fillStyle = "#7f8c8d";
            ctx.textAlign = "center";
            ctx.fillText("أضف منتجات ليظهر الرسم البياني هنا", ctx.canvas.width/2, ctx.canvas.height/2);
        }
    </script>
</body>
</html>
"""

# 3. المسارات والتوجيهات (Routes)
@app.route('/')
def index():
    sales = get_all_sales()
    return render_template_string(HTML_TEMPLATE, sales=sales)

@app.route('/add', methods=['POST'])
def add_sale():
    try:
        product = request.form.get('product')
        price = request.form.get('price')
        cost = request.form.get('cost')
        
        if product:
            insert_sale(product, price, cost)
    except Exception as e:
        print("Add Error:", e)
    return redirect(url_for('index'))

@app.route('/api')
def api():
    return jsonify(get_all_sales())

@app.route('/export/excel')
def export_excel():
    sales = get_all_sales()
    si = io.StringIO()
    cw = csv.writer(si)
    # كتابة العناوين
    cw.writerow(['المعرف', 'المنتج', 'سعر البيع', 'التكلفة', 'الربح الصافي', 'التاريخ'])
    for s in sales:
        cw.writerow([s['id'], s['product'], s['price'], s['cost'], s['profit'], s['date']])
    
    output = make_response(si.getvalue())
    output.headers["Content-Disposition"] = "attachment; filename=sales_report.csv"
    # utf-8-sig تضمن قراءة الحروف العربية بشكل سليم داخل Excel
    output.headers["Content-type"] = "text/csv; charset=utf-8-sig"
    return output

@app.route('/export/pdf')
def export_pdf():
    sales = get_all_sales()
    # واجهة طباعة نظيفة ومحسنة كتقرير PDF جاهز للطباعة مباشرة
    pdf_template = """
    <!DOCTYPE html>
    <html lang="ar" dir="rtl">
    <head>
        <meta charset="UTF-8">
        <title>تقرير المبيعات المالي</title>
        <style>
            body { font-family: Arial, sans-serif; padding: 40px; }
            h1 { text-align: center; color: #2c3e50; }
            table { width: 100%; border-collapse: collapse; margin-top: 30px; }
            th, td { border: 1px solid #333; padding: 10px; text-align: right; }
            th { background-color: #f2f2f2; }
        </style>
    </head>
    <body>
        <h1>تقرير المبيعات والأرباح المالي</h1>
        <table>
            <thead>
                <tr>
                    <th>المعرف</th>
                    <th>اسم المنتج</th>
                    <th>سعر البيع</th>
                    <th>التكلفة</th>
                    <th>الربح الصافي</th>
                    <th>التاريخ</th>
                </tr>
            </thead>
            <tbody>
                {% for sale in sales %}
                <tr>
                    <td>{{ sale.id }}</td>
                    <td>{{ sale.product }}</td>
                    <td>${{ sale.price }}</td>
                    <td>${{ sale.cost }}</td>
                    <td>${{ sale.profit }}</td>
                    <td>{{ sale.date }}</td>
                </tr>
                {% endfor %}
            </tbody>
        </table>
        <script>window.onload = function() { window.print(); }</script>
    </body>
    </html>
    """
    return render_template_string(pdf_template, sales=sales)

if __name__ == '__main__':
    init_db()
    app.run(debug=True)
