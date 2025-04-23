import json
import os
from jinja2 import Environment, FileSystemLoader
import pdfkit
from openpyxl import Workbook
#json 暫用
data = {
  "title": "銷售報表",
  "items": [
    {"name": "滑鼠", "quantity": 10, "price": 250},
    {"name": "鍵盤", "quantity": 5, "price": 500},
    {"name": "螢幕", "quantity": 2, "price": 3500}
  ]
}

# 使用 Jinja2 渲染 HTML 模板
env = Environment(loader=FileSystemLoader(""))
template = env.get_template("template.html")
html_out = template.render(data)

# 建立輸出資料夾
os.makedirs('output', exist_ok=True)

# 輸出 HTML 檔
html_path = 'output/report.html'
with open(html_path, 'w', encoding='utf-8') as f:
    f.write(html_out)

# 輸出 PDF 檔（需安裝 wkhtmltopdf）
pdf_path = 'output/report.pdf'
pdfkit.from_file(html_path, pdf_path)

# 使用 openpyxl 寫 Excel 檔
wb = Workbook()
ws = wb.active
ws.title = "銷售報表"

# 寫入標題列
ws.append(["產品名稱", "數量", "價格"])

# 寫入資料列
for item in data["items"]:
    ws.append([item["name"], item["quantity"], item["price"]])

# 儲存 Excel 檔
excel_path = 'output/report.xlsx'
wb.save(excel_path)

print("報表已完成輸出：HTML、PDF、Excel")