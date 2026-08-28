import os
import io
import re
import base64
import sqlite3
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from PIL import Image

try:
    from google import genai
    from google.genai import types
except ImportError:
    genai = None

app = Flask(__name__)
DB_NAME = 'inventory.db'

# 取得 Gemini API Key (可從環境變數讀取)
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

def get_db():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def format_date_str(val):
    """將 ISO 時間 (如 2026-08-27T16:00:00.000Z) 自動清洗為 YYYY-MM-DD"""
    if not val:
        return ""
    val_str = str(val).strip()
    if 'T' in val_str:
        return val_str.split('T')[0]
    if len(val_str) >= 10 and val_str[4] == '-' and val_str[7] == '-':
        return val_str[:10]
    return val_str

def init_db():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS inventory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            category TEXT,
            mfg_date TEXT,
            exp_date TEXT,
            in_date TEXT,
            quantity INTEGER DEFAULT 1,
            unit TEXT DEFAULT '個',
            notes TEXT,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    
    # 自動清洗並覆蓋資料庫中已存在的長字串日期
    try:
        cursor.execute("UPDATE inventory SET mfg_date = SUBSTR(mfg_date, 1, 10) WHERE mfg_date LIKE '%T%'")
        cursor.execute("UPDATE inventory SET exp_date = SUBSTR(exp_date, 1, 10) WHERE exp_date LIKE '%T%'")
        cursor.execute("UPDATE inventory SET in_date = SUBSTR(in_date, 1, 10) WHERE in_date LIKE '%T%'")
        conn.commit()
    except Exception as e:
        print("Date sanitize error:", e)
    finally:
        conn.close()

# 程式啟動時初始化與清洗資料庫
init_db()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/items', methods=['GET'])
def get_items():
    conn = get_db()
    items = conn.execute("SELECT * FROM inventory ORDER BY id DESC").fetchall()
    conn.close()
    
    result = []
    for item in items:
        d = dict(item)
        d['mfg_date'] = format_date_str(d.get('mfg_date'))
        d['exp_date'] = format_date_str(d.get('exp_date'))
        d['in_date'] = format_date_str(d.get('in_date'))
        result.append(d)
    return jsonify(result)

@app.route('/api/items', methods=['POST'])
def add_item():
    data = request.json or {}
    name = data.get('name', '').strip()
    category = data.get('category', '')
    mfg_date = format_date_str(data.get('mfg_date', ''))
    exp_date = format_date_str(data.get('exp_date', ''))
    in_date = format_date_str(data.get('in_date', ''))
    quantity = int(data.get('quantity', 1))
    unit = data.get('unit', '個')
    notes = data.get('notes', '')

    if not name:
        return jsonify({"error": "請輸入品名"}), 400

    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO inventory (name, category, mfg_date, exp_date, in_date, quantity, unit, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (name, category, mfg_date, exp_date, in_date, quantity, unit, notes))
    conn.commit()
    new_id = cursor.lastrowid
    conn.close()
    return jsonify({"success": True, "id": new_id})

@app.route('/api/items/<int:item_id>', methods=['PUT'])
def update_item(item_id):
    data = request.json or {}
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE inventory
        SET name = ?, category = ?, mfg_date = ?, exp_date = ?, in_date = ?, quantity = ?, unit = ?, notes = ?
        WHERE id = ?
    """, (
        data.get('name', '').strip(),
        data.get('category', ''),
        format_date_str(data.get('mfg_date', '')),
        format_date_str(data.get('exp_date', '')),
        format_date_str(data.get('in_date', '')),
        int(data.get('quantity', 1)),
        data.get('unit', '個'),
        data.get('notes', ''),
        item_id
    ))
    conn.commit()
    conn.close()
    return jsonify({"success": True})

@app.route('/api/items/<int:item_id>', methods=['DELETE'])
def delete_item(item_id):
    conn = get_db()
    conn.execute("DELETE FROM inventory WHERE id = ?", (item_id,))
    conn.commit()
    conn.close()
    return jsonify({"success": True})

@app.route('/api/recognize', methods=['POST'])
def recognize_image():
    """使用 Google GenAI 辨識食品包裝照片"""
    if 'image' not in request.files:
        return jsonify({"error": "未提供圖片檔案"}), 400
    
    file = request.files['image']
    if file.filename == '':
        return jsonify({"error": "檔案名稱為空"}), 400

    image_bytes = file.read()
    
    api_key = os.environ.get("GEMINI_API_KEY", GEMINI_API_KEY)
    if not api_key or not genai:
        return jsonify({
            "name": "食品樣品",
            "mfg_date": "",
            "exp_date": "",
            "notes": "未設定 GEMINI_API_KEY"
        })

    try:
        client = genai.Client(api_key=api_key)
        prompt = (
            "請分析這張食品包裝圖片，以繁體中文提取以下資訊：\n"
            "1. 食品完整中文品名 (例如：桂格大燕麥片、愛之味牛奶花生)\n"
            "2. 製造日期 (格式 YYYY-MM-DD，若無請留空)\n"
            "3. 有效日期 / 保存期限 (格式 YYYY-MM-DD，若無請留空)\n"
            "請僅輸出 JSON 格式如下：\n"
            "{\"name\": \"品名\", \"mfg_date\": \"YYYY-MM-DD\", \"exp_date\": \"YYYY-MM-DD\"}"
        )
        
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=[
                types.Part.from_bytes(data=image_bytes, mime_type=file.mimetype or "image/jpeg"),
                prompt
            ]
        )
        
        text = response.text or ""
        json_match = re.search(r'\{.*\}', text, re.DOTALL)
        if json_match:
            import json
            result = json.loads(json_match.group(0))
            result['mfg_date'] = format_date_str(result.get('mfg_date', ''))
            result['exp_date'] = format_date_str(result.get('exp_date', ''))
            return jsonify(result)
        else:
            return jsonify({"name": text.strip()[:30], "mfg_date": "", "exp_date": ""})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/export')
def export_excel():
    """匯出格式美化、無多餘時間字元的食品物資存放盤點表"""
    conn = get_db()
    items = conn.execute("SELECT * FROM inventory ORDER BY id ASC").fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "盤點表"

    # 樣式定義
    current_month_str = datetime.now().strftime("%Y-%m")
    title_font = Font(name="微軟正黑體", size=16, bold=True)
    header_font = Font(name="微軟正黑體", size=11, bold=True)
    body_font = Font(name="微軟正黑體", size=11)
    
    header_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid") # 淡綠色表頭
    thin_border = Border(
        left=Side(style='thin', color='A0A0A0'),
        right=Side(style='thin', color='A0A0A0'),
        top=Side(style='thin', color='A0A0A0'),
        bottom=Side(style='thin', color='A0A0A0')
    )
    center_align = Alignment(horizontal='center', vertical='center')
    left_align = Alignment(horizontal='left', vertical='center')

    # 大標題
    ws.merge_cells("A1:H1")
    ws["A1"] = f"食品物資存放盤點表 ({current_month_str})"
    ws["A1"].font = title_font
    ws["A1"].alignment = center_align
    ws.row_dimensions[1].height = 40

    # 表頭欄位
    headers = ["序號", "食品名稱", "類別", "製造日期", "有效日期", "入庫日期", "數量/單位", "備註"]
    ws.append([]) # 空行
    ws.append(headers)
    ws.row_dimensions[3].height = 26

    for col_idx in range(1, 9):
        cell = ws.cell(row=3, column=col_idx)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border

    # 填入資料 (清洗日期)
    for idx, row in enumerate(items, 1):
        mfg = format_date_str(row['mfg_date'])
        exp = format_date_str(row['exp_date'])
        in_d = format_date_str(row['in_date'])
        qty_unit = f"{row['quantity']} {row['unit'] or ''}".strip()
        
        row_data = [
            idx,
            row['name'] or "",
            row['category'] or "",
            mfg,
            exp,
            in_d,
            qty_unit,
            row['notes'] or ""
        ]
        ws.append(row_data)
        
        current_row = ws.max_row
        ws.row_dimensions[current_row].height = 22
        for col_idx in range(1, 9):
            c = ws.cell(row=current_row, column=col_idx)
            c.font = body_font
            c.border = thin_border
            # 品名與備註靠左，其餘置中
            if col_idx in [2, 8]:
                c.alignment = left_align
            else:
                c.alignment = center_align

    # 自動調整欄寬
    col_widths = {1: 8, 2: 24, 3: 14, 4: 15, 5: 15, 6: 15, 7: 12, 8: 20}
    for col_idx, width in col_widths.items():
        ws.column_dimensions[chr(64 + col_idx)].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    filename = f"食品盤點表_{current_month_str}.xlsx"
    return send_file(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename
    )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
