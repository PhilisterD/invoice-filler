import logging
import os
import re
import sys
import time
import traceback
import webbrowser
import zipfile
import shutil
from logging.handlers import RotatingFileHandler
from threading import Timer
import subprocess

from flask import Flask, g, render_template, request, send_file, redirect, url_for
from docx import Document
from werkzeug.exceptions import HTTPException

from filler import (
    build_summary_text,
    fill_template,
    get_row_text,
    is_summary_row_text,
    has_numeric_value,
    merge_total_column_span,
    set_cell_text,
)


def get_resource_path(relative_path):
    """Get absolute path to a resource, works for dev and PyInstaller."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative_path)


app = Flask(__name__)
# Bundled (read-only) resource uploads included in the exe by PyInstaller
BUNDLED_UPLOADS = get_resource_path('uploads')
TEMPLATE_PATH = get_resource_path('template.docx')

# Runtime (writable) directories — when frozen, place next to the executable; otherwise next to source
if getattr(sys, '_MEIPASS', False):
    runtime_root = os.path.dirname(sys.executable)
else:
    runtime_root = os.path.dirname(os.path.abspath(__file__))

UPLOAD_FOLDER = os.path.join(runtime_root, 'uploads')
OUTPUT_FOLDER = os.path.join(runtime_root, 'output')
LOG_FOLDER = os.path.join(runtime_root, 'log')
LOG_FILE = os.path.join(LOG_FOLDER, 'invoice-filler.log')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)

# Built-in templates that ship with the application
BUILTIN_TEMPLATES = [
    '早鸟天筹明细-模板1.docx',
    '上海辉研明细-模板1.docx',
    '狮山辉研明细-模板1.docx',
]


def configure_logging():
    logger = logging.getLogger('invoice-filler')
    logger.setLevel(logging.INFO)
    logger.propagate = False
    if not logger.handlers:
        formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')

        file_handler = RotatingFileHandler(
            LOG_FILE,
            maxBytes=1024 * 1024,
            backupCount=5,
            encoding='utf-8',
        )
        file_handler.setFormatter(formatter)

        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)

        logger.addHandler(file_handler)
        logger.addHandler(console_handler)
    return logger


logger = configure_logging()


def log_step(message, **fields):
    if fields:
        extra = ' | '.join(f'{key}={value}' for key, value in fields.items())
        logger.info('%s | %s', message, extra)
    else:
        logger.info(message)


def sync_builtin_templates():
    """Ensure all built-in templates exist in the runtime uploads folder.

    This copies missing templates from the bundled (read-only) directory so that
    users always have access to the full set even if the runtime uploads folder
    was previously populated with only a subset (e.g. an older version)."""
    if not os.path.exists(BUNDLED_UPLOADS):
        return
    for name in BUILTIN_TEMPLATES:
        src = os.path.join(BUNDLED_UPLOADS, name)
        dst = os.path.join(UPLOAD_FOLDER, name)
        if os.path.isfile(src):
            try:
                shutil.copy2(src, dst)
            except Exception:
                # best-effort copy; don't fail startup if copy fails
                pass


sync_builtin_templates()

results_cache = []


@app.before_request
def log_request_start():
    g.request_start_time = time.time()
    logger.info(
        'HTTP %s %s from %s',
        request.method,
        request.path,
        request.remote_addr or '-',
    )


@app.after_request
def log_request_end(response):
    start_time = getattr(g, 'request_start_time', None)
    if start_time is not None:
        duration_ms = (time.time() - start_time) * 1000
        logger.info(
            'HTTP %s %s -> %s in %.1fms',
            request.method,
            request.path,
            response.status_code,
            duration_ms,
        )
    return response


@app.errorhandler(Exception)
def log_unhandled_exception(exc):
    if isinstance(exc, HTTPException):
        return exc
    logger.error(
        'Unhandled error on %s %s\n%s',
        request.method,
        request.path,
        traceback.format_exc(),
    )
    return '服务器发生内部错误，请查看日志', 500


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
            row_text = get_row_text(row)
            is_summary = is_summary_row_text(row_text) or ri == len(table.rows) - 1
            if any(cells_text):
                rows_data.append({
                    'index': ri,
                    'cells': cells_text,
                    'is_summary': is_summary,
                })
                if is_summary:
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
    template_name = request.form.get('template', '').strip()

    log_step(
        '开始处理上传请求',
        excel=excel_file.filename if excel_file else '',
        template=template_name or '默认模板',
    )

    if not excel_file:
        logger.warning('缺少 Excel 文件上传')
        return '需要上传 Excel 明细表', 400

    excel_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)
    excel_file.save(excel_path)
    log_step('Excel 已保存', path=excel_path)

    # 确定模板路径：自定义上传 > 预设选择 > 默认
    custom_template_file = request.files.get('custom_template')
    if template_name == '__custom__' and custom_template_file and custom_template_file.filename:
        custom_path = os.path.join(UPLOAD_FOLDER, custom_template_file.filename)
        custom_template_file.save(custom_path)
        selected_template = custom_path
        log_step('使用自定义模板', path=selected_template)
    elif template_name:
        # Prefer runtime (writable) uploads, fall back to bundled templates included in the exe
        runtime_candidate = os.path.join(UPLOAD_FOLDER, template_name)
        bundled_candidate = os.path.join(BUNDLED_UPLOADS, template_name) if os.path.exists(BUNDLED_UPLOADS) else None
        if os.path.exists(runtime_candidate):
            selected_template = runtime_candidate
            log_step('使用运行时模板', name=template_name, path=selected_template)
        elif bundled_candidate and os.path.exists(bundled_candidate):
            selected_template = bundled_candidate
            log_step('使用内置模板', name=template_name, path=selected_template)
        else:
            logger.error('找不到模板文件: %s (runtime=%s bundled=%s)', template_name, runtime_candidate, bundled_candidate)
            return f'找不到模板文件: {template_name}', 400
    else:
        selected_template = TEMPLATE_PATH
        log_step('使用默认模板', path=selected_template)

    if not os.path.exists(selected_template):
        logger.warning('模板文件不存在: %s', selected_template)
        return f'模板文件不存在: {template_name}', 400

    # 清空旧输出
    log_step('清空旧输出目录', output_dir=OUTPUT_FOLDER)
    for f in os.listdir(OUTPUT_FOLDER):
        os.remove(os.path.join(OUTPUT_FOLDER, f))

    log_step('开始调用填充逻辑')
    results_cache = fill_template(excel_path, selected_template, OUTPUT_FOLDER, logger=logger)
    log_step('填充逻辑完成', results=len(results_cache))
    return redirect(url_for('result'))


@app.route('/result')
def result():
    return render_template('result.html', results=results_cache)


@app.route('/log')
def log_view():
    if os.path.exists(LOG_FILE):
        with open(LOG_FILE, 'r', encoding='utf-8', errors='replace') as f:
            log_text = ''.join(f.readlines()[-300:])
    else:
        log_text = '暂无日志'
    return render_template('log.html', log_text=log_text, log_file=LOG_FILE)


@app.route('/preview/<filename>')
def preview(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(path):
        return '文件不存在', 404
    data = read_docx_data(path)
    return render_template('preview.html', filename=filename, data=data)


@app.route('/rename/<filename>', methods=['POST'])
def rename(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(path):
        return '文件不存在', 404

    new_name = request.form.get('new_name', '').strip()
    if not new_name:
        return '文件名不能为空', 400
    if not new_name.lower().endswith('.docx'):
        new_name += '.docx'

    safe_name = re.sub(r'[\n\r\\/:*?"<>|]', '_', new_name).strip()
    new_path = os.path.join(OUTPUT_FOLDER, safe_name)

    if os.path.exists(new_path) and new_path != path:
        return '文件名已存在', 400

    os.rename(path, new_path)
    log_step('重命名文件', old=filename, new=safe_name)

    for r in results_cache:
        if r['filename'] == filename:
            r['filename'] = safe_name
            r['path'] = new_path
            break

    return redirect(url_for('result'))


@app.route('/edit/<filename>', methods=['GET', 'POST'])
def edit(filename):
    path = os.path.join(OUTPUT_FOLDER, filename)
    if not os.path.exists(path):
        return '文件不存在', 404

    if request.method == 'POST':
        log_step('开始编辑文档', filename=filename)
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

            from filler import detect_table_layout
            data_row_idx, summary_row_idx, _ = detect_table_layout(table)
            if data_row_idx is None or summary_row_idx is None:
                return '无法识别模板中的数据行和汇总行', 400

            # 清空所有数据行
            for ri in range(data_row_idx, summary_row_idx):
                row = table.rows[ri]
                for cell in row.cells:
                    set_cell_text(cell, '')

            # 填充数据
            total_sum = 0
            merge_row_indices = []
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
                if has_numeric_value(rec['合计']):
                    merge_row_indices.append(data_row_idx + i)

            # 填充总计列（安全地尝试纵向合并）
            total_col = 4
            if total_col < len(table.rows[0].cells):
                for ri in range(data_row_idx, data_row_idx + len(data_rows)):
                    row = table.rows[ri]
                    if total_col < len(row.cells):
                        cell = row.cells[total_col]
                        if ri == data_row_idx:
                            set_cell_text(cell, str(int(total_sum)) if total_sum == int(total_sum) else str(total_sum))
                        else:
                            set_cell_text(cell, '')
                merge_total_column_span(table, merge_row_indices, total_col)

            # 填充汇总行
            summary_row = table.rows[summary_row_idx]
            if summary_row.cells:
                seen_tc = set()
                summary_text = build_summary_text(total_sum, summary_row)
                for cell in summary_row.cells:
                    tc = cell._tc
                    if tc in seen_tc:
                        continue
                    seen_tc.add(tc)
                    set_cell_text(cell, summary_text)

        doc.save(path)
        log_step('编辑文档已保存', filename=filename, total=total_sum, records=len(data_rows))

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
        if row.get('is_summary'):
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
        log_step('下载文件', filename=filename)
        return send_file(path, as_attachment=True)
    return '文件不存在', 404


@app.route('/download/all')
def download_all():
    zip_path = os.path.join(OUTPUT_FOLDER, 'all_documents.zip')
    log_step('打包下载全部文件', zip_path=zip_path)
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for r in results_cache:
            if r.get('path') and os.path.exists(r['path']):
                zf.write(r['path'], arcname=r['filename'])
    return send_file(zip_path, as_attachment=True, download_name='all_documents.zip')


def open_browser():
    webbrowser.open('http://localhost:5001/')


def kill_previous_instances():
    """When running as a frozen exe, try to kill any previously-running instances
    of the same executable (excluding current PID). Best-effort; non-fatal."""
    if not getattr(sys, 'frozen', False):
        return
    exe_name = os.path.basename(sys.executable)
    own_pid = os.getpid()
    try:
        import psutil
        for proc in psutil.process_iter(['pid', 'name', 'exe']):
            try:
                pid = proc.info.get('pid')
                name = proc.info.get('name')
                exe = proc.info.get('exe')
                if pid == own_pid:
                    continue
                if name == exe_name or (exe and os.path.basename(exe) == exe_name):
                    proc.kill()
            except Exception:
                continue
    except Exception:
        # Fallback: use platform tools
        if os.name == 'nt':
            try:
                out = subprocess.check_output(['tasklist', '/FI', f'IMAGENAME eq {exe_name}', '/FO', 'CSV', '/NH'], text=True, stderr=subprocess.DEVNULL)
                for line in out.splitlines():
                    parts = [p.strip().strip('"') for p in line.split(',')]
                    if len(parts) >= 2:
                        try:
                            pid = int(parts[1])
                        except Exception:
                            continue
                        if pid != own_pid:
                            subprocess.run(['taskkill', '/PID', str(pid), '/F'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception:
                pass
        else:
            try:
                out = subprocess.check_output(['pgrep', '-f', exe_name], text=True)
                for pid in out.split():
                    try:
                        pid_i = int(pid)
                        if pid_i != own_pid:
                            subprocess.run(['kill', '-9', str(pid_i)], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    except Exception:
                        continue
            except Exception:
                pass


@app.route('/shutdown', methods=['POST'])
def shutdown():
    # Only allow local requests
    remote = request.remote_addr or ''
    if remote not in ('127.0.0.1', '::1', 'localhost') and not remote.startswith('192.') and remote != '0.0.0.0':
        return 'Forbidden', 403
    func = request.environ.get('werkzeug.server.shutdown')
    if func:
        func()
        return '服务器已关闭'
    # If werkzeug shutdown not available (packaged differently), exit after short delay
    Timer(0.5, lambda: os._exit(0)).start()
    return '服务器即将关闭（不可用 werkzeug.shutdown）'


if __name__ == '__main__':
    import socket
    hostname = socket.gethostname()
    local_ip = socket.getaddrinfo(hostname, None)[0][4][0]
    logger.info('Flask 应用已启动')
    logger.info('本地访问: http://localhost:5001')
    logger.info('日志页面: http://localhost:5001/log')
    logger.info('局域网访问: http://%s:5001', local_ip)
    # 如果是打包后的 exe，优先尝试终止之前运行的同名实例，避免多个后台进程
    try:
        kill_previous_instances()
    except Exception:
        pass
    logger.info('已执行旧实例清理检查')
    Timer(1.5, open_browser).start()
    app.run(host='0.0.0.0', port=5001, debug=False)
