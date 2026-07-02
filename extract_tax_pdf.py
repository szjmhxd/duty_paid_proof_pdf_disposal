# -*- coding: utf-8 -*-
"""
完税凭证PDF批量提取脚本
功能：将PDF完税凭证中的文本/表格提取出来，转为XML、CSV
"""
import os
import sys
import csv
import xml.etree.ElementTree as ET

def install_deps():
    for dep in ['pdfplumber', 'pandas']:
        try:
            __import__(dep)
        except ImportError:
            os.system(f"pip install {dep}")

install_deps()

import pdfplumber

# PDF数据目录：请通过命令行参数指定，或修改此处的默认路径
PDF_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.')
OUTPUT_DIR = os.path.join(PDF_DIR, 'output')
os.makedirs(OUTPUT_DIR, exist_ok=True)


def extract_from_pdf(pdf_path):
    result = {'文件名': os.path.basename(pdf_path), '文本': '', '表格': []}
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ''
            result['文本'] += text + '\n'
            for table in page.extract_tables():
                result['表格'].append(table)
    return result


def parse_tax_fields(text):
    fields = {}
    for line in text.split('\n'):
        line = line.strip()
        sep = '：' if '：' in line else ':' if ':' in line else None
        if not sep:
            continue
        key, val = line.split(sep, 1)
        key, val = key.strip(), val.strip()
        mapping = {
            '纳税人姓名': '纳税人姓名', '纳税人': '纳税人姓名',
            '身份证': '身份证号', '证件号码': '身份证号',
            '税种': '税种', '税款所属期': '税款所属期', '所属期': '税款所属期',
            '实缴税额': '实缴税额', '实纳': '实缴税额',
            '纳税金额': '税额', '税额': '税额',
            '完税凭证号': '凭证号', '凭证号': '凭证号',
            '征收机关': '征收机关', '缴款单位': '扣缴义务人',
            '扣缴义务人': '扣缴义务人', '品目名称': '品目名称',
            '入库日期': '入库日期', '开具单位': '税务机关', '税务机关': '税务机关',
        }
        for pattern, field in mapping.items():
            if pattern in key:
                fields[field] = val
                break
    return fields


def save_as_xml(all_data, output_path):
    root = ET.Element('完税凭证汇总')
    for item in all_data:
        record = ET.SubElement(root, '记录')
        ET.SubElement(record, '文件名').text = item.get('文件名', '')
        fields = item.get('字段', {})
        if fields:
            fields_elem = ET.SubElement(record, '字段')
            for k, v in fields.items():
                ET.SubElement(fields_elem, k).text = v
        ET.SubElement(record, '原始文本').text = item.get('文本', '')
        tables = item.get('表格', [])
        if tables:
            tables_elem = ET.SubElement(record, '表格')
            for t_idx, table in enumerate(tables):
                table_elem = ET.SubElement(tables_elem, f'表格{t_idx+1}')
                for row in table:
                    row_elem = ET.SubElement(table_elem, '行')
                    for cell in row:
                        ET.SubElement(row_elem, '单元格').text = str(cell) if cell else ''
    xml_str = '<?xml version="1.0" encoding="utf-8"?>\n' + ET.tostring(root, encoding='unicode')
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(xml_str)
    print(f"XML已保存: {output_path}")


def save_as_csv(all_data, output_path):
    all_field_names = set()
    for item in all_data:
        all_field_names.update(item.get('字段', {}).keys())
    field_names = sorted(all_field_names)
    headers = ['文件名'] + field_names + ['原始文本']
    with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(headers)
        for item in all_data:
            fields = item.get('字段', {})
            row = [item.get('文件名', '')] + [fields.get(fn, '') for fn in field_names]
            row.append(item.get('文本', '').replace('\n', ' | '))
            writer.writerow(row)
    print(f"CSV已保存: {output_path}")


def main():
    pdf_files = sorted([f for f in os.listdir(PDF_DIR) if f.endswith('.pdf') and '(1)' not in f])
    print(f"找到 {len(pdf_files)} 个PDF文件（已去重）")

    all_data = []
    for i, fname in enumerate(pdf_files):
        fpath = os.path.join(PDF_DIR, fname)
        print(f"[{i+1}/{len(pdf_files)}] {fname}")
        result = extract_from_pdf(fpath)
        fields = parse_tax_fields(result['文本'])
        all_data.append({'文件名': fname, '文本': result['文本'], '字段': fields, '表格': result['表格']})

    save_as_xml(all_data, os.path.join(OUTPUT_DIR, '完税凭证汇总.xml'))
    save_as_csv(all_data, os.path.join(OUTPUT_DIR, '完税凭证汇总.csv'))

    txt_path = os.path.join(OUTPUT_DIR, '完税凭证纯文本.txt')
    with open(txt_path, 'w', encoding='utf-8') as f:
        for item in all_data:
            f.write(f"===== {item['文件名']} =====\n{item['文本']}\n\n")
    print(f"纯文本已保存: {txt_path}")
    print(f"\n处理完成！输出目录: {OUTPUT_DIR}")


if __name__ == '__main__':
    main()
