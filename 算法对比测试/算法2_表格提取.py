﻿# -*- coding: utf-8 -*-
"""
算法2：基于表格提取的方式
使用pdfplumber的extract_tables()方法提取PDF中的表格数据
"""
import os
import sys
import json
import pdfplumber

# 设置编码
sys.stdout.reconfigure(encoding='utf-8')

# 配置路径
# PDF数据目录：请通过命令行参数指定，或修改此处的默认路径
PDF_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '.')
OUTPUT_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
os.makedirs(OUTPUT_DIR, exist_ok=True)

# 选择一个测试PDF文件
TEST_FILES = [f for f in os.listdir(PDF_DIR) if f.endswith('.pdf') and '(1)' not in f][:3]  # 取前3个测试

def extract_tables_from_pdf(pdf_path):
    """使用表格提取方式解析PDF"""
    result = {
        '文件名': os.path.basename(pdf_path),
        '页数': 0,
        '表格数量': 0,
        '表格数据': [],
        '文本数据': []
    }
    
    with pdfplumber.open(pdf_path) as pdf:
        result['页数'] = len(pdf.pages)
        
        for page_idx, page in enumerate(pdf.pages):
            # 提取文本
            text = page.extract_text()
            if text:
                result['文本数据'].append({
                    '页码': page_idx + 1,
                    '文本内容': text
                })
            
            # 提取表格
            tables = page.extract_tables()
            result['表格数量'] += len(tables)
            
            for table_idx, table in enumerate(tables):
                table_data = {
                    '页码': page_idx + 1,
                    '表格序号': table_idx + 1,
                    '行数': len(table),
                    '列数': len(table[0]) if table else 0,
                    '内容': table
                }
                result['表格数据'].append(table_data)
    
    return result

def save_result(result, output_dir):
    """保存结果到文件"""
    filename = result['文件名'].replace('.pdf', '')
    
    # 保存JSON格式（便于查看结构）
    json_path = os.path.join(output_dir, f'{filename}_表格提取.json')
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    
    # 保存文本格式
    txt_path = os.path.join(output_dir, f'{filename}_表格提取.txt')
    with open(txt_path, 'w', encoding='utf-8') as f:
        f.write(f"文件: {result['文件名']}\n")
        f.write(f"页数: {result['页数']}\n")
        f.write(f"表格数量: {result['表格数量']}\n\n")
        
        # 写入文本数据
        f.write("=" * 60 + "\n")
        f.write("文本数据:\n")
        f.write("=" * 60 + "\n")
        for text_item in result['文本数据']:
            f.write(f"\n--- 第{text_item['页码']}页 ---\n")
            f.write(text_item['文本内容'] + "\n")
        
        # 写入表格数据
        f.write("\n" + "=" * 60 + "\n")
        f.write("表格数据:\n")
        f.write("=" * 60 + "\n")
        for table_item in result['表格数据']:
            f.write(f"\n--- 第{table_item['页码']}页 - 表格{table_item['表格序号']} ---\n")
            f.write(f"行数: {table_item['行数']}, 列数: {table_item['列数']}\n")
            for row_idx, row in enumerate(table_item['内容']):
                f.write(f"  行{row_idx+1}: {row}\n")
    
    return json_path, txt_path

def main():
    print("=" * 60)
    print("算法2：表格提取方式测试")
    print("=" * 60)
    print(f"测试文件夹: {PDF_DIR}")
    print(f"测试文件数量: {len(TEST_FILES)}")
    print()
    
    all_results = []
    
    for i, filename in enumerate(TEST_FILES):
        print(f"[{i+1}/{len(TEST_FILES)}] 处理文件: {filename}")
        pdf_path = os.path.join(PDF_DIR, filename)
        
        try:
            result = extract_tables_from_pdf(pdf_path)
            json_path, txt_path = save_result(result, OUTPUT_DIR)
            
            print(f"  页数: {result['页数']}")
            print(f"  表格数量: {result['表格数量']}")
            print(f"  文本块数量: {len(result['文本数据'])}")
            print(f"  JSON保存到: {json_path}")
            print(f"  文本保存到: {txt_path}")
            
            all_results.append(result)
            
        except Exception as e:
            print(f"  错误: {e}")
    
    # 生成汇总报告
    summary_path = os.path.join(OUTPUT_DIR, '表格提取_汇总.txt')
    with open(summary_path, 'w', encoding='utf-8') as f:
        f.write("算法2：表格提取方式 - 汇总报告\n")
        f.write("=" * 60 + "\n\n")
        
        total_pages = sum(r['页数'] for r in all_results)
        total_tables = sum(r['表格数量'] for r in all_results)
        total_text_blocks = sum(len(r['文本数据']) for r in all_results)
        
        f.write(f"测试文件数: {len(all_results)}\n")
        f.write(f"总页数: {total_pages}\n")
        f.write(f"总表格数: {total_tables}\n")
        f.write(f"总文本块数: {total_text_blocks}\n\n")
        
        for result in all_results:
            f.write(f"文件: {result['文件名']}\n")
            f.write(f"  页数: {result['页数']}, 表格: {result['表格数量']}, 文本块: {len(result['文本数据'])}\n")
    
    print(f"\n汇总报告保存到: {summary_path}")
    print("\n测试完成！")

if __name__ == '__main__':
    main()
