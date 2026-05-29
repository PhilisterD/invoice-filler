import os
import re
from copy import deepcopy
from datetime import datetime

import pandas as pd
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt

from utils import num_to_chinese, format_date


def is_yellow_highlight(run) -> bool:
    """判断一个 run 是否是黄色高亮 (highlight_color == 7)"""
    if run.font.highlight_color == 7:
        return True
    rPr = run._element.find(qn('w:rPr'))
    if rPr is not None:
        shd = rPr.find(qn('w:shd'))
        if shd is not None:
            fill = shd.get(qn('w:fill'))
            if fill and fill.upper() in ('FFFF00', 'FF0', 'YELLOW'):
                return True
    return False


def find_yellow_runs(paragraph):
    """返回段落中所有黄色高亮的 run 索引"""
    return [i for i, run in enumerate(paragraph.runs) if is_yellow_highlight(run)]


def find_table_rows_with_yellow(table):
    """找出表格中包含黄色高亮的行索引"""
    rows = []
    for ri, row in enumerate(table.rows):
        for cell in row.cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    if is_yellow_highlight(run):
                        rows.append(ri)
                        break
                if ri in rows:
                    break
            if ri in rows:
                break
    return sorted(set(rows))


def set_cell_text(cell, text):
    """安全设置单元格文本，保留第一个 run 的样式"""
    p = cell.paragraphs[0]
    if p.runs:
        p.runs[0].text = text
        for r in p.runs[1:]:
            r.text = ""
    else:
        p.add_run(text)


def remove_highlight_from_run(run):
    """清除 run 的高亮和底纹"""
    run.font.highlight_color = None
    rPr = run._element.find(qn('w:rPr'))
    if rPr is not None:
        for hl in rPr.findall(qn('w:highlight')):
            rPr.remove(hl)
        for shd in rPr.findall(qn('w:shd')):
            rPr.remove(shd)


def remove_all_highlights(doc):
    """清除文档中所有段落和表格的高亮"""
    for p in doc.paragraphs:
        for run in p.runs:
            remove_highlight_from_run(run)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for run in p.runs:
                        remove_highlight_from_run(run)


def copy_row_xml(table, source_row_idx):
    """复制表格中的一行并插入到其后（XML层面）。"""
    tbl = table._tbl
    source_tr = table.rows[source_row_idx]._tr
    new_tr = deepcopy(source_tr)
    actual_idx = list(tbl).index(source_tr)
    tbl.insert(actual_idx + 1, new_tr)


def fill_template(excel_path: str, template_path: str, output_dir: str) -> list[dict]:
    """
    读取Excel，按受托方分组，为每组生成一个Word文档。
    返回生成文件的信息列表。
    """
    df = pd.read_excel(excel_path)
    rename_map = {}
    for col in df.columns:
        c = str(col).strip()
        if '受托方' in c or '委托方' in c:
            rename_map[col] = '受托方'
        elif '姓名' in c:
            rename_map[col] = '姓名'
        elif '服务内容' in c:
            rename_map[col] = '服务内容'
        elif '数量' in c and '单价' not in c and '合计' not in c:
            rename_map[col] = '数量'
        elif '单价' in c:
            rename_map[col] = '单价'
        elif '合计' in c and '总计' not in c:
            rename_map[col] = '合计'
        elif '总计' in c:
            rename_map[col] = '总计'
        elif '日期' in c:
            rename_map[col] = '日期'
        elif '序号' in c:
            rename_map[col] = '序号'
    df.rename(columns=rename_map, inplace=True)

    for col in ['数量', '单价', '合计', '总计']:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    grouped = df.groupby('受托方', sort=False)
    results = []

    for party, group in grouped:
        group = group.sort_values('序号') if '序号' in group.columns else group
        doc = Document(template_path)

        # ---- 1. 替换文档中黄色高亮的文本（委托方、日期） ----
        for p in doc.paragraphs:
            yellow_indices = find_yellow_runs(p)
            for idx in yellow_indices:
                run = p.runs[idx]
                text = run.text
                if any(kw in text for kw in ['研究所', '大学', '公司', '学院', '中心']):
                    run.text = str(party)
                elif re.search(r'\d{4}年\d{1,2}月\d{1,2}日', text):
                    date_val = group['日期'].iloc[0]
                    run.text = format_date(date_val)

        # ---- 2. 处理表格 ----
        if not doc.tables:
            results.append({
                'party': party,
                'filename': f"{party}_明细.docx",
                'path': None,
                'error': '模板中没有表格'
            })
            continue

        table = doc.tables[0]
        yellow_rows = find_table_rows_with_yellow(table)

        if len(yellow_rows) < 2:
            results.append({
                'party': party,
                'filename': f"{party}_明细.docx",
                'path': None,
                'error': '表格中标黄行不足2行（数据行+汇总行）'
            })
            continue

        data_row_idx = yellow_rows[0]

        # 获取列结构
        header_row = table.rows[0]
        headers = [cell.text.strip() for cell in header_row.cells]
        col_map = {}
        for i, h in enumerate(headers):
            if '服务' in h or '内容' in h:
                col_map['服务内容'] = i
            elif '数量' in h:
                col_map['数量'] = i
            elif '单价' in h:
                col_map['单价'] = i
            elif '合计' in h and '总计' not in h:
                col_map['合计'] = i
            elif '总计' in h:
                col_map['总计'] = i

        records = group.to_dict('records')

        # 清空原始数据行
        data_row = table.rows[data_row_idx]
        for cell in data_row.cells:
            set_cell_text(cell, "")

        # 插入副本（需要记录数 - 1 个新行）
        for _ in range(len(records) - 1):
            copy_row_xml(table, data_row_idx)

        # 重新打开文档以刷新 table.rows（python-docx 缓存问题）
        temp_path = os.path.join(output_dir, '_temp.docx')
        doc.save(temp_path)
        doc = Document(temp_path)
        os.remove(temp_path)
        table = doc.tables[0]

        # 重新定位行索引
        yellow_rows = find_table_rows_with_yellow(table)
        data_row_idx = yellow_rows[0]
        summary_row_idx = yellow_rows[-1]

        # 填充数据行
        total_sum = 0
        for i, rec in enumerate(records):
            target_row = table.rows[data_row_idx + i]
            service = str(rec.get('服务内容', ''))
            qty = rec.get('数量', 0)
            price = rec.get('单价', 0)
            subtotal = rec.get('合计', 0)
            if pd.isna(subtotal) or subtotal == 0:
                subtotal = (qty or 0) * (price or 0)
            total_sum += subtotal

            cells = target_row.cells
            if '服务内容' in col_map and col_map['服务内容'] < len(cells):
                set_cell_text(cells[col_map['服务内容']], service)
            if '数量' in col_map and col_map['数量'] < len(cells):
                set_cell_text(cells[col_map['数量']], str(int(qty)) if qty == int(qty) else str(qty))
            if '单价' in col_map and col_map['单价'] < len(cells):
                set_cell_text(cells[col_map['单价']], str(int(price)) if price == int(price) else str(price))
            if '合计' in col_map and col_map['合计'] < len(cells):
                set_cell_text(cells[col_map['合计']], str(int(subtotal)) if subtotal == int(subtotal) else str(subtotal))

        # 填充总计列（只在第一行设置，然后纵向合并单元格）
        total_col = col_map.get('总计', len(table.rows[0].cells) - 1)
        first_total_cell = None
        last_total_cell = None
        for ri in range(data_row_idx, summary_row_idx):
            row = table.rows[ri]
            if total_col < len(row.cells):
                cell = row.cells[total_col]
                if ri == data_row_idx:
                    set_cell_text(cell, str(int(total_sum)) if total_sum == int(total_sum) else str(total_sum))
                    first_total_cell = cell
                else:
                    set_cell_text(cell, "")
                last_total_cell = cell

        # 纵向合并总计列单元格（多于一行数据时才合并）
        if first_total_cell and last_total_cell and last_total_cell != first_total_cell:
            first_total_cell.merge(last_total_cell)

        # 填充汇总行（合并单元格中多个 cell 可能共享同一个 XML 元素，需要去重）
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

        # ---- 4. 清除所有黄色高亮 ----
        remove_all_highlights(doc)

        # ---- 5. 保存 ----
        person = str(group['姓名'].iloc[0]) if '姓名' in group.columns else ''
        safe_person = re.sub(r'[\\/:*?"<>|]', '_', person)
        safe_party = re.sub(r'[\\/:*?"<>|]', '_', str(party))
        filename = f"{safe_person}_{safe_party}_{int(total_sum)}.docx"
        out_path = os.path.join(output_dir, filename)
        doc.save(out_path)
        results.append({
            'party': party,
            'filename': filename,
            'path': out_path,
            'records': len(records),
            'total': int(total_sum)
        })

    return results


if __name__ == '__main__':
    import sys
    if len(sys.argv) >= 3:
        excel = sys.argv[1]
        template = sys.argv[2]
        out = sys.argv[3] if len(sys.argv) > 3 else './output'
    else:
        excel = '../明细-excel.xlsx'
        template = '../早鸟天筹明细-模板1.docx'
        out = './output'
    os.makedirs(out, exist_ok=True)
    res = fill_template(excel, template, out)
    for r in res:
        print(r)
