# -*- coding: utf-8 -*-
"""算法2：基于pdfplumber表格提取功能的PDF解析"""
import os, sys, json, time
import pdfplumber

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


def extract_with_pdfplumber(pdf_path):
    """使用pdfplumber提取PDF的文本和表格信息"""
    result = {
        '文件名': os.path.basename(pdf_path),
        '算法': 'pdfplumber_table_extract',
        '页数': 0,
        '总表格数': 0,
        '总文本块数': 0,
        '页面详情': []
    }

    with pdfplumber.open(pdf_path) as pdf:
        result['页数'] = len(pdf.pages)

        for page_idx, page in enumerate(pdf.pages):
            page_info = {
                '页码': page_idx + 1,
                '宽度': float(page.width),
                '高度': float(page.height),
                '文本': '',
                '表格': []
            }

            # 提取文本
            text = page.extract_text()
            if text:
                page_info['文本'] = text
                result['总文本块数'] += 1

            # 提取表格（核心功能）
            tables = page.extract_tables()
            result['总表格数'] += len(tables)

            for tbl_idx, table in enumerate(tables):
                table_data = {
                    '表格序号': tbl_idx + 1,
                    '行数': len(table),
                    '列数': len(table[0]) if table else 0,
                    '数据': []
                }
                # 将每行数据转为字典格式（用首行作为表头）
                if table and len(table) > 1:
                    headers = table[0] if table[0] else [f'列{i+1}' for i in range(len(table[1]))]
                    for row in table[1:]:
                        row_dict = {}
                        for col_idx, cell in enumerate(row):
                            col_name = headers[col_idx] if col_idx < len(headers) else f'列{col_idx+1}'
                            row_dict[col_name] = cell
                        table_data['数据'].append(row_dict)
                # 同时保留原始二维数组格式
                table_data['原始数据'] = table

                page_info['表格'].append(table_data)

            result['页面详情'].append(page_info)

    return result


def save_result(result, output_dir):
    """保存提取结果为JSON和可读文本文件"""
    fn = result['文件名'].replace('.pdf', '')

    # 保存JSON结果
    json_path = os.path.join(output_dir, f'{fn}_alg2.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    # 保存可读文本结果
    txt_path = os.path.join(output_dir, f'{fn}_alg2.txt')
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write(f"算法: pdfplumber表格提取\n")
        f.write(f"文件: {result['文件名']}\n")
        f.write(f"页数: {result['页数']}\n")
        f.write(f"总表格数: {result['总表格数']}\n")
        f.write(f"总文本块数: {result['总文本块数']}\n\n")

        for page in result['页面详情']:
            f.write(f"{'='*40}\n")
            f.write(f"第{page['页码']}页 (尺寸: {page['宽度']:.1f} x {page['高度']:.1f})\n")
            f.write(f"{'='*40}\n")

            if page['文本']:
                f.write(f"\n[文本内容]\n{page['文本']}\n")

            for tbl in page['表格']:
                f.write(f"\n[表格{tbl['表格序号']} - {tbl['行数']}行x{tbl['列数']}列]\n")
                for row in tbl['数据']:
                    f.write(f"  {row}\n")

    return json_path, txt_path


# ========== 主程序入口 ==========
if __name__ == '__main__':
    print("=" * 60)
    print("算法2：pdfplumber表格提取")
    print(f"源目录: {PDF_DIR}")
    print(f"输出目录: {OUTPUT_DIR}")
    print(f"测试文件数: {len(TEST_FILES)}")
    print("=" * 60)

    total_start = time.time()

    for i, fn in enumerate(TEST_FILES):
        print(f"\n[{i+1}/{len(TEST_FILES)}] 处理: {fn}")
        start = time.time()
        try:
            result = extract_with_pdfplumber(os.path.join(PDF_DIR, fn))
            json_path, txt_path = save_result(result, OUTPUT_DIR)
            elapsed = time.time() - start
            print(f"  耗时: {elapsed:.2f}秒")
            print(f"  页数: {result['页数']}, 表格数: {result['总表格数']}, 文本块: {result['总文本块数']}")
            print(f"  JSON输出: {json_path}")
        except Exception as e:
            print(f"  错误: {e}")

    total_elapsed = time.time() - total_start
    print(f"\n{'='*60}")
    print(f"全部完成！总耗时: {total_elapsed:.2f}秒")
    print(f"结果保存在: {OUTPUT_DIR}")
    print("=" * 60)
