# -*- coding: utf-8 -*-
import io
import os
import json
import requests
from datetime import datetime, timedelta
from flask import Flask, render_template, request, send_file, redirect, url_for, jsonify
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from google import genai
from PIL import Image

app = Flask(__name__)

# 已配置您的專屬 Google Apps Script 雲端網址
GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwWzg4xz5YwHwTc4K7KFNQqCy_kfe73h9rQ4MWZMdvlMEfmF5Vf7PSshhRbaqrOfp8O/exec"

ORG_FULL_NAME = "財團法人私立天主教中華聖母社會福利慈善事業基金會附設嘉義縣私立隆興社區長照機構(團體家屋)"

API_KEY = "AQ.Ab8RN6LbOPya1X3Ax44ZGGdQPRJE36da0GU5bYXhZxq_1b6u8w"
client = genai.Client(api_key=API_KEY)

def format_int(val):
    if val is None or val == '':
        return None
    try:
        return int(round(float(val)))
    except:
        return val

def fetch_sheet_data():
    try:
        res = requests.get(GOOGLE_SCRIPT_URL, timeout=10)
        data = res.json()
        raw_items = data.get('items', [])
        raw_logs = data.get('logs', [])
        
        items = []
        if len(raw_items) > 1:
            for row in raw_items[1:]:
                if len(row) >= 8 and row[0] != '':
                    items.append({
                        'id': int(row[0]),
                        'item_name': str(row[1]),
                        'in_qty': format_int(row[2]),
                        'in_unit': str(row[3]),
                        'mfg_date': str(row[4]),
                        'exp_date': str(row[5]),
                        'in_date': str(row[6]),
                        'in_signer': str(row[7]),
                        'current_stock': format_int(row[8]) if len(row) > 8 and row[8] != '' else format_int(row[2]),
                        'last_check_date': str(row[9]) if len(row) > 9 else '',
                        'last_check_signer': str(row[10]) if len(row) > 10 else '',
                        'last_remark': str(row[11]) if len(row) > 11 else ''
                    })

        logs = []
        if len(raw_logs) > 1:
            for row in raw_logs[1:]:
                if len(row) >= 10 and row[0] != '':
                    logs.append({
                        'id': int(row[0]),
                        'item_id': int(row[1]),
                        'item_name': str(row[2]),
                        'in_qty': format_int(row[3]),
                        'in_unit': str(row[4]),
                        'in_date': str(row[5]),
                        'check_date': str(row[6]),
                        'check_qty': format_int(row[7]),
                        'check_signer': str(row[8]),
                        'remark': str(row[9])
                    })
        return items, logs
    except Exception as e:
        print("讀取 Google Sheets 失敗:", e)
        return [], []

def post_sheet_data(payload):
    try:
        requests.post(GOOGLE_SCRIPT_URL, json=payload, timeout=10)
    except Exception as e:
        print("寫入 Google Sheets 失敗:", e)

@app.route('/')
def index():
    today_str = datetime.now().strftime('%Y-%m-%d')
    limit_30_days = (datetime.now() + timedelta(days=30)).strftime('%Y-%m-%d')

    items, logs = fetch_sheet_data()

    active_items = []
    distinct_names = set()
    expiring_count = 0
    expired_count = 0
    checked_count = 0

    for it in items:
        distinct_names.add((it['item_name'], it['in_unit']))
        if it['current_stock'] is None or it['current_stock'] > 0:
            exp = it.get('exp_date') or ''
            it['is_expiring'] = False
            it['is_expired'] = False
            if exp:
                if exp < today_str:
                    it['is_expired'] = True
                    expired_count += 1
                elif exp <= limit_30_days:
                    it['is_expiring'] = True
                    expiring_count += 1
            if it.get('last_check_date'):
                checked_count += 1
            active_items.append(it)

    active_items.reverse()

    checked_items = [l for l in reversed(logs) if l['check_qty'] > 0]
    finished_items = [l for l in reversed(logs) if l['check_qty'] == 0]

    distinct_items = [{"item_name": name, "in_unit": unit} for name, unit in distinct_names]

    return render_template(
        'index.html',
        items=active_items,
        checked_items=checked_items,
        finished_items=finished_items,
        distinct_items=distinct_items,
        today=today_str,
        expiring_count=expiring_count,
        expired_count=expired_count,
        checked_count=checked_count,
        org_name=ORG_FULL_NAME
    )

@app.route('/recognize_item', methods=['POST'])
def recognize_item():
    if 'image' not in request.files:
        return jsonify({'error': '未收到圖片檔案'}), 400
    
    file = request.files['image']
    try:
        image = Image.open(file.stream)
        prompt = (
            "這是一張長照機構入庫的食品照片。請仔細辨識包裝上的【品牌廠牌】與【完整食品名稱】。"
            "嚴格格式要求：請務必完整輸出「品牌 + 食品品名」（例如：桂格大燕麥片、愛之味脆瓜、統一肉燥麵、工研烏醋、同榮豆豉鰻魚）。"
            "嚴格禁止輸出任何問候語、解釋、英文、標點符號、包裝標示或日期，只回傳品名文字。"
        )
        response = client.models.generate_content(
            model='gemini-3.6-flash',
            contents=[image, prompt]
        )
        recognized_name = response.text.strip().replace("\n", "").replace("*", "").replace("`", "")
        return jsonify({'item_name': recognized_name})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/add', methods=['POST'])
def add_item():
    f = request.form
    in_qty = format_int(f.get('in_qty'))
    in_unit = f.get('in_unit', '')
    
    post_sheet_data({
        "action": "add_item",
        "item": {
            "item_name": f.get('item_name'),
            "in_qty": in_qty,
            "in_unit": in_unit,
            "mfg_date": f.get('mfg_date'),
            "exp_date": f.get('exp_date'),
            "in_date": f.get('in_date'),
            "in_signer": f.get('in_signer')
        }
    })
    return redirect(url_for('index'))

@app.route('/edit_item/<int:item_id>', methods=['POST'])
def edit_item(item_id):
    f = request.form
    post_sheet_data({
        "action": "edit_item",
        "item": {
            "id": item_id,
            "item_name": f.get('item_name'),
            "in_qty": format_int(f.get('in_qty')),
            "in_unit": f.get('in_unit'),
            "mfg_date": f.get('mfg_date'),
            "exp_date": f.get('exp_date'),
            "in_date": f.get('in_date'),
            "in_signer": f.get('in_signer')
        }
    })
    return redirect(url_for('index'))

@app.route('/update_check/<int:item_id>', methods=['POST'])
def update_check(item_id):
    f = request.form
    check_date = f.get('check_date')
    check_qty = format_int(f.get('check_qty')) or 0
    check_signer = f.get('check_signer')
    remark = f.get('remark', '')
    if check_qty == 0 and not remark:
        remark = "已用完"

    post_sheet_data({
        "action": "update_check",
        "check": {
            "item_id": item_id,
            "check_date": check_date,
            "check_qty": check_qty,
            "check_signer": check_signer,
            "remark": remark
        }
    })
    return redirect(url_for('index'))

@app.route('/delete/<int:item_id>', methods=['POST'])
def delete_item(item_id):
    post_sheet_data({
        "action": "delete_item",
        "item_id": item_id
    })
    return redirect(url_for('index'))

@app.route('/delete_log/<int:log_id>', methods=['POST'])
def delete_log(log_id):
    post_sheet_data({
        "action": "delete_log",
        "log_id": log_id
    })
    return redirect(url_for('index'))

@app.route('/export')
def export_excel():
    target_month = request.args.get('month', datetime.now().strftime('%Y-%m'))
    
    wb = Workbook()
    ws = wb.active
    ws.title = f"{target_month}盤點表"

    thin = Side(style='thin', color='000000')
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    font_title = Font(name='標楷體', size=15, bold=True)
    font_header = Font(name='標楷體', size=11, bold=True)
    font_data = Font(name='標楷體', size=11)
    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    header_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')

    ws.merge_cells('A1:L1')
    ws['A1'] = f"{ORG_FULL_NAME}\n食品物資存放盤點表 ({target_month})"
    ws['A1'].font = font_title
    ws['A1'].alignment = align_center
    ws.row_dimensions[1].height = 45

    headers = [
        "食品品名", "入庫數量", "單位", "製造日期", "有效日期", 
        "入庫日期", "入庫人員簽章", "盤點日期", "實存數量", "單位", 
        "盤點人員簽章", "說明/處置情形"
    ]
    ws.row_dimensions[2].height = 25
    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(row=2, column=col_idx, value=h_text)
        cell.font = font_header
        cell.alignment = align_center
        cell.border = border_all
        cell.fill = header_fill

    items, _ = fetch_sheet_data()
    export_rows = []
    for it in items:
        in_d = str(it.get('in_date', ''))
        chk_d = str(it.get('last_check_date', ''))
        if in_d.startswith(target_month) or chk_d.startswith(target_month):
            curr_chk_d = chk_d if chk_d.startswith(target_month) else ''
            curr_chk_q = format_int(it.get('current_stock')) if curr_chk_d else ''
            curr_chk_s = it.get('last_check_signer', '') if curr_chk_d else ''
            curr_chk_r = it.get('last_remark', '') if curr_chk_d else ''

            export_rows.append([
                it.get('item_name'), format_int(it.get('in_qty')), it.get('in_unit'),
                it.get('mfg_date'), it.get('exp_date'), in_d, it.get('in_signer'),
                curr_chk_d, curr_chk_q, it.get('in_unit'), curr_chk_s, curr_chk_r
            ])

    start_row = 3
    for i, row_data in enumerate(export_rows):
        curr_row = start_row + i
        ws.row_dimensions[curr_row].height = 22

        for col_idx, val in enumerate(row_data, start=1):
            cell = ws.cell(row=curr_row, column=col_idx, value=val if val is not None else '')
            cell.font = font_data
            cell.alignment = align_center
            cell.border = border_all

    col_widths = [24, 10, 8, 13, 13, 13, 14, 13, 10, 8, 14, 20]
    for col_idx, width in enumerate(col_widths, start=1):
        col_letter = ws.cell(row=2, column=col_idx).column_letter
        ws.column_dimensions[col_letter].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return send_file(
        output,
        as_attachment=True,
        download_name=f"隆興團體家屋食品物資存放盤點表_{target_month}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)