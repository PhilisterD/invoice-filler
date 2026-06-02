#!/usr/bin/env python3
"""
替换 Word 模板中的印章/图片。

用法:
    python replace_stamp.py <模板路径> <新图片路径> [输出路径]

示例:
    python replace_stamp.py uploads/早鸟天筹明细-模板1.docx stamp_hd.png uploads/早鸟天筹明细-模板1_new.docx
"""
import os
import sys
import shutil
import zipfile
from docx import Document


def replace_image_in_docx(docx_path, new_image_path, output_path=None):
    """替换 docx 中的第一张图片（通常就是印章）。"""
    if output_path is None:
        output_path = docx_path

    if not os.path.exists(docx_path):
        print(f'错误: 找不到模板文件 {docx_path}')
        return False

    if not os.path.exists(new_image_path):
        print(f'错误: 找不到新图片 {new_image_path}')
        return False

    # 读取 docx 中的图片关系
    doc = Document(docx_path)
    rels = doc.part.rels

    image_rels = [r for r in rels.values() if 'image' in r.reltype]
    if not image_rels:
        print('错误: 模板中没有找到图片')
        return False

    target_rel = image_rels[0]
    print(f'找到图片关系: {target_rel.rId} -> {target_rel.target_ref}')

    # 获取新图片的扩展名
    _, new_ext = os.path.splitext(new_image_path)
    new_ext = new_ext.lower()
    if new_ext not in ('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff'):
        print(f'警告: 不常见的图片格式 {new_ext}')

    # 确定目标 content type
    content_type_map = {
        '.png': 'image/png',
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.gif': 'image/gif',
        '.bmp': 'image/bmp',
        '.tiff': 'image/tiff',
    }
    new_content_type = content_type_map.get(new_ext, 'image/jpeg')

    # 构建新的 docx（通过 ZIP 操作）
    temp_path = output_path + '.tmp'
    try:
        with zipfile.ZipFile(docx_path, 'r') as zin:
            with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)
                    if item.filename.startswith('word/media/'):
                        # 替换第一张图片（通常只有一个印章）
                        if 'image1' in item.filename or item.filename == target_rel.target_ref:
                            with open(new_image_path, 'rb') as f:
                                data = f.read()
                            print(f'替换图片: {item.filename} ({len(data)} bytes)')
                            # 同时更新 [Content_Types].xml 中的 content type（如果需要）
                        zout.writestr(item, data)
                    elif item.filename == '[Content_Types].xml':
                        # 更新 content type（如果扩展名改变）
                        text = data.decode('utf-8')
                        # 简单替换旧的 image/jpeg 为新的（如果不同）
                        # 更精确的做法是解析 XML，但这里简单处理
                        zout.writestr(item, text)
                    else:
                        zout.writestr(item, data)

        # 如果用原地替换，先备份
        if output_path == docx_path:
            backup = docx_path + '.bak'
            shutil.copy2(docx_path, backup)
            print(f'已备份原文件到 {backup}')

        shutil.move(temp_path, output_path)
        print(f'保存成功: {output_path}')
        return True

    except Exception as e:
        print(f'处理失败: {e}')
        if os.path.exists(temp_path):
            os.remove(temp_path)
        return False


def analyze_docx_images(docx_path):
    """分析 docx 中所有图片的尺寸和大小。"""
    from PIL import Image
    import io

    if not os.path.exists(docx_path):
        print(f'错误: 找不到文件 {docx_path}')
        return

    print(f'\n=== 分析: {docx_path} ===')
    with zipfile.ZipFile(docx_path, 'r') as zf:
        for info in zf.infolist():
            if 'media' in info.filename:
                data = zf.read(info.filename)
                try:
                    img = Image.open(io.BytesIO(data))
                    # 计算 DPI 信息
                    dpi = img.info.get('dpi', (72, 72))
                    print(f'  {info.filename}:')
                    print(f'    尺寸: {img.size[0]} x {img.size[1]} 像素')
                    print(f'    模式: {img.mode}')
                    print(f'    文件大小: {len(data):,} bytes ({len(data)/1024:.1f} KB)')
                    print(f'    DPI: {dpi[0]} x {dpi[1]}')
                except Exception as e:
                    print(f'  {info.filename}: 无法解析 ({e})')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    cmd = sys.argv[1]

    if cmd == 'analyze':
        for path in sys.argv[2:]:
            analyze_docx_images(path)
    elif len(sys.argv) >= 3:
        docx_path = sys.argv[1]
        img_path = sys.argv[2]
        out_path = sys.argv[3] if len(sys.argv) > 3 else None
        replace_image_in_docx(docx_path, img_path, out_path)
    else:
        print(__doc__)
        sys.exit(1)
