import os
import re
import webbrowser
import zipfile
from threading import Timer

from flask import Flask, render_template, request, send_file, redirect, url_for
from docx import Document

from filler import fill_template, set_cell_text
from utils import num_to_chinese

app = Flask(__name__)
UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

results_cache = []


def read_docx_data(path):
    """读取 docx 中的关键数据，用于预览和编辑。"""
    doc = Document(path)
    party = ''
    date_str = ''
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        if any(kw in text for kw in ['研究所', '大学', '公司', '学院', '中心']):
            party = text
        elif re.search(r'\d{4}年\d{1,2}月\d{1,2}日', text):
            date_str = text

    rows_data = []
    total = 0
    summary_text = ''
    if doc.tables:
        table = doc.tables[0]
        for ri, row in enumerate(table.rows):
            cells_text = [cell.text.strip() for cell in row.cells]
            if any(cells_text):
                rows_data.append({
                    'index': ri,
                    'cells': cells_text,
                })
                if '测试费用总共' in cells_text[0] or any('测试费用总共' in c for c in cells_text):
                    summary_text = cells_text[0]
                    m = re.search(r'人民币\s*(\d+)', summary_text)
                    if m:
                        total = int(m.group(1))
    return {
        'party': party,
        'date': date_str,
        'rows': rows_data,
        'total': total,
        'summary': summary_text,
    }


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/process', methods=['POST'])
def process():
    global results_cache
    excel_file = request.files.get('excel')
    template_file = request.files.get('template')

    if not excel_file or not template_file:
        return '需要上传 Excel 和 Word 模板两个文件', 400

    excel_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
    template_path = os.path.join(UPLOAD_FOLDER, template_file.filename)
    excel_file.save(excel_path)
    template_file.save(template_path)

    # 清空旧输出
    for f in os.listdir(OUTPUT_FOLDER):
        os.remove(os.path.join(OUTPUT_FOLDER, f))

    results_cache = fill_template(excel_path, template_path, OUTPUT_FOLDER)
    return redirect(url_for('result'))


@app.route('/result')
def result():
    return render_template('result.html', results=results_cache)


@app.route('/preview/<filename>')
def preview(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(path):
        return '文件不存在', 404
    data = read_docx_data(path)
    return render_template('preview.html', filename=filename, data=data)


@app.route('/edit/<filename>', methods=['GET', 'POST'])
def edit(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(path):
        return '文件不存在', 404

    if request.method == 'POST':
        doc = Document(path)
        new_party = request.form.get('party', '').strip()
        new_date = request.form.get('date', '').strip()
        services = request.form.getlist('service[]')
        quantities = request.form.getlist('quantity[]')
        prices = request.form.getlist('price[]')

        # 1. 更新段落
        for p in doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            if any(kw in text for kw in ['研究所', '大学', '公司', '学院', '中心']):
                p.text = new_party
            elif re.search(r'\d{4}年\d{1,2}月\d{1,2}日', text):
                p.text = new_date

        # 2. 更新表格
        if doc.tables:
            table = doc.tables[0]
            data_rows = []
            for i in range(len(services)):
                svc = services[i].strip()
                qty_str = quantities[i].strip()
                price_str = prices[i].strip()
                if not svc:
                    continue
                try:
                    qty = float(qty_str) if qty_str else 0
                except ValueError:
                    qty = 0
                try:
                    price = float(price_str) if price_str else 0
                except ValueError:
                    price = 0
                subtotal = qty * price
                data_rows.append({
                    '服务内容': svc,
                    '数量': qty,
                    '单价': price,
                    '合计': subtotal,
                })

            # 找到数据行和汇总行
            from filler import find_table_rows_with_yellow
            yellow_rows = find_table_rows_with_yellow(table)
            if len(yellow_rows) >= 2:
                data_row_idx = yellow_rows[0]
                summary_row_idx = yellow_rows[-1]

                # 清空所有数据行
                for ri in range(data_row_idx, summary_row_idx):
                    row = table.rows[ri]
                    for cell in row.cells:
                        set_cell_text(cell, '')

                # 填充数据
                total_sum = 0
                for i, rec in enumerate(data_rows):
                    target_row = table.rows[data_row_idx + i]
                    cells = target_row.cells
                    if len(cells) > 0:
                        set_cell_text(cells[0], rec['服务内容'])
                    if len(cells) > 1:
                        set_cell_text(cells[1], str(int(rec['数量'])) if rec['数量'] == int(rec['数量']) else str(rec['数量']))
                    if len(cells) > 2:
                        set_cell_text(cells[2], str(int(rec['单价'])) if rec['单价'] == int(rec['单价']) else str(rec['单价']))
                    if len(cells) > 3:
                        set_cell_text(cells[3], str(int(rec['合计'])) if rec['合计'] == int(rec['合计']) else str(rec['合计']))
                    if len(cells) > 4:
                        set_cell_text(cells[4], '')
                    total_sum += rec['合计']

                # 填充总计列（纵向合并）
                total_col = 4
                if total_col < len(table.rows[0].cells):
                    first_total_cell = None
                    last_total_cell = None
                    for ri in range(data_row_idx, data_row_idx + len(data_rows)):
                        row = table.rows[ri]
                        if total_col < len(row.cells):
                            cell = row.cells[total_col]
                            if ri == data_row_idx:
                                set_cell_text(cell, str(int(total_sum)) if total_sum == int(total_sum) else str(total_sum))
                                first_total_cell = cell
                            else:
                                set_cell_text(cell, '')
                            last_total_cell = cell
                    if first_total_cell and last_total_cell and last_total_cell != first_total_cell:
                        first_total_cell.merge(last_total_cell)

                # 填充汇总行
                summary_row = table.rows[summary_row_idx]
                if summary_row.cells:
                    seen_tc = set()
                    chinese_total = num_to_chinese(int(total_sum))
                    summary_text = f"测试费用总共为：人民币  {int(total_sum)}元（大写：人民币{chinese_total}）"
                    for cell in summary_row.cells:
                        tc = cell._tc
                        if tc in seen_tc:
                            continue
                        seen_tc.add(tc)
                        set_cell_text(cell, summary_text)

        doc.save(path)

        # 更新缓存
        for r in results_cache:
            if r['filename'] == filename:
                r['total'] = int(total_sum)
                r['records'] = len(data_rows)

        return redirect(url_for('preview', filename=filename))

    data = read_docx_data(path)
    # 提取可编辑的数据行（去掉表头和汇总行）
    edit_rows = []
    for row in data['rows']:
        cells = row['cells']
        if '测试费用总共' in cells[0]:
            continue
        if any(cells):
            edit_rows.append({
                'service': cells[0] if len(cells) > 0 else '',
                'quantity': cells[1] if len(cells) > 1 else '',
                'price': cells[2] if len(cells) > 2 else '',
            })
    return render_template('edit.html', filename=filename, data=data, rows=edit_rows)


@app.route('/download/<filename>')
def download(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    return '文件不存在', 404


@app.route('/download/all')
def download_all():
    zip_path = os.path.join(OUTPUT_FOLDER, 'all_documents.zip')
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for r in results_cache:
            if r.get('path') and os.path.exists(r['path']):
                zf.write(r['path'], arcname=r['filename'])
    return send_file(zip_path, as_attachment=True, download_name='all_documents.zip')


def open_browser():
    webbrowser.open('http://localhost:5001/')


if __name__ == '__main__':
    import socket
    hostname = socket.gethostname()
    local_ip = socket.getaddrinfo(hostname, None)[0][4][0]
    print(f"Flask 应用已启动!")
    print(f"本地访问: http://localhost:5001")
    print(f"局域网访问: http://{local_ip}:5001")
    Timer(1.5, open_browser).start()
    app.run(host='0.0.0.0', port=5001, debug=False)
