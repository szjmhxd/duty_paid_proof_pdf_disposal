# -*- coding: utf-8 -*-
"""算法4：基于PyMuPDF(fitz)进行PDF结构解析"""
import os, sys, json, time
import fitz  # PyMuPDF

# 修复Windows控制台UTF-8编码问题
sys.stdout.reconfigure(encoding='utf-8')

# 路径配置
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
# PDF数据目录：请通过命令行参数指定，或修改此处的默认路径
PDF_DIR = os.path.join(BASE_DIR, '..', '.')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 获取前3个PDF文件（排除带(1)的重复文件）
TEST_FILES = [f for f in os.listdir(PDF_DIR) if f.endswith('.pdf') and '(1)' not in f][:3]


def extract_with_pymupdf(pdf_path):
    """使用PyMuPDF提取PDF的文本、结构和元数据信息"""
    result = {
        '文件名': os.path.basename(pdf_path),
        '算法': 'pymupdf_structured_extract',
        '页数': 0,
        '元数据': {},
        '总文本块数': 0,
        '总图片数': 0,
        '总链接数': 0,
        '页面详情': []
    }

    doc = fitz.open(pdf_path)

    # 提取PDF元数据
    result['元数据'] = {
        '标题': doc.metadata.get('title', ''),
        '作者': doc.metadata.get('author', ''),
        '主题': doc.metadata.get('subject', ''),
        '创建程序': doc.metadata.get('creator', ''),
        '页数': doc.page_count
    }

    result['页数'] = doc.page_count

    for page_idx in range(doc.page_count):
        page = doc.load_page(page_idx)

        page_info = {
            '页码': page_idx + 1,
            '宽度': page.rect.width,
            '高度': page.rect.height,
            '旋转角度': page.rotation,
            '文本': '',
            '文本块详情': [],
            '图片列表': [],
            '链接列表': [],
            '表格信息': []
        }

        # 提取完整文本
        full_text = page.get_text("text")
        if full_text.strip():
            page_info['文本'] = full_text
            result['总文本块数'] += 1

        # 提取文本块详情（带位置信息）
        blocks = page.get_text("dict")["blocks"]
        for block in blocks:
            if block["type"] == 0:  # 文本块
                block_text = ""
                for line in block.get("lines", []):
                    for span in line.get("spans", []):
                        block_text += span["text"]
                if block_text.strip():
                    page_info['文本块详情'].append({
                        '位置': {
                            'x0': round(block["bbox"][0], 2),
                            'y0': round(block["bbox"][1], 2),
                            'x1': round(block["bbox"][2], 2),
                            'y1': round(block["bbox"][3], 2)
                        },
                        '文本': block_text.strip()
                    })

        # 提取图片信息
        image_list = page.get_images(full=True)
        for img in image_list:
            page_info['图片列表'].append({
                'xref': img[0],
                '宽度': img[2],
                '高度': img[3]
            })
        result['总图片数'] += len(image_list)

        # 提取链接
        links = page.get_links()
        for link in links:
            page_info['链接列表'].append({
                '类型': link.get('kind', ''),
                '目标': link.get('uri', link.get('to', ''))
            })
        result['总链接数'] += len(links)

        # 使用PyMuPDF的表格检测功能
        try:
            tab_finder = page.find_tables()
            tables = tab_finder.tables
            for tbl_idx, table in enumerate(tables):
                tbl_data = table.extract()
                page_info['表格信息'].append({
                    '表格序号': tbl_idx + 1,
                    '行列数': f'{table.rows}行x{table.cols}列',
                    '边界': {
                        'x0': round(table.bbox[0], 2),
                        'y0': round(table.bbox[1], 2),
                        'x1': round(table.bbox[2], 2),
                        'y1': round(table.bbox[3], 2)
                    },
                    '数据': tbl_data
                })
        except Exception:
            pass  # 部分版本不支持find_tables

        result['页面详情'].append(page_info)

    doc.close()
    return result


def save_result(result, output_dir):
    """保存提取结果为JSON和可读文本文件"""
    fn = result['文件名'].replace('.pdf', '')

    # 保存JSON结果
    json_path = os.path.join(output_dir, f'{fn}_alg4.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    # 保存可读文本结果
    txt_path = os.path.join(output_dir, f'{fn}_alg4.txt')
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write(f"算法: PyMuPDF结构解析\n")
        f.write(f"文件: {result['文件名']}\n")
        f.write(f"页数: {result['页数']}\n")
        f.write(f"元数据: {json.dumps(result['元数据'], ensure_ascii=False)}\n")
        f.write(f"总图片数: {result['总图片数']}\n")
        f.write(f"总链接数: {result['总链接数']}\n\n")

        for page in result['页面详情']:
            f.write(f"{'='*40}\n")
            f.write(f"第{page['页码']}页 (尺寸: {page['宽度']:.1f} x {page['高度']:.1f}, 旋转: {page['旋转角度']}°)\n")
            f.write(f"{'='*40}\n")

            if page['文本']:
                f.write(f"\n[文本内容]\n{page['文本']}\n")

            if page['文本块详情']:
                f.write(f"\n[文本块详情 ({len(page['文本块详情'])}个)]\n")
                for block in page['文本块详情']:
                    f.write(f"  位置{block['位置']}: {block['文本']}\n")

            if page['图片列表']:
                f.write(f"\n[图片 ({len(page['图片列表'])}个)]\n")
                for img in page['图片列表']:
                    f.write(f"  xref={img['xref']}, {img['宽度']}x{img['高度']}\n")

            if page['表格信息']:
                f.write(f"\n[表格 ({len(page['表格信息'])}个)]\n")
                for tbl in page['表格信息']:
                    f.write(f"  表格{tbl['表格序号']}: {tbl['行列数']}\n")
                    for row in tbl['数据']:
                        f.write(f"    {row}\n")

    return json_path, txt_path


# ========== 主程序入口 ==========
if __name__ == '__main__':
    print("=" * 60)
    print("算法4：PyMuPDF结构解析")
    print(f"源目录: {PDF_DIR}")
    print(f"输出目录: {OUTPUT_DIR}")
    print(f"测试文件数: {len(TEST_FILES)}")
    print("=" * 60)

    total_start = time.time()

    for i, fn in enumerate(TEST_FILES):
        print(f"\n[{i+1}/{len(TEST_FILES)}] 处理: {fn}")
        start = time.time()
        try:
            result = extract_with_pymupdf(os.path.join(PDF_DIR, fn))
            json_path, txt_path = save_result(result, OUTPUT_DIR)
            elapsed = time.time() - start
            print(f"  耗时: {elapsed:.2f}秒")
            print(f"  页数: {result['页数']}, 图片: {result['总图片数']}, 链接: {result['总链接数']}")
            print(f"  JSON输出: {json_path}")
        except Exception as e:
            print(f"  错误: {e}")

    total_elapsed = time.time() - total_start
    print(f"\n{'='*60}")
    print(f"全部完成！总耗时: {total_elapsed:.2f}秒")
    print(f"结果保存在: {OUTPUT_DIR}")
    print("=" * 60)
