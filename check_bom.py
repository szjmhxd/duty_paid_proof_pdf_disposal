import os

# 检查源仓库中的.py文件是否有BOM
repo_dir = r'e:\test\python\pdf完税凭证文档'
files_to_check = [
    os.path.join(repo_dir, 'extract_tax_pdf.py'),
    os.path.join(repo_dir, '可视化提取', 'app.py'),
    os.path.join(repo_dir, '算法对比测试', 'alg2_table_extract.py'),
    os.path.join(repo_dir, '算法对比测试', 'alg4_pymupdf.py'),
    os.path.join(repo_dir, '算法对比测试', 'merge_extract.py'),
    os.path.join(repo_dir, '算法对比测试', '算法2_表格提取.py'),
]

for f in files_to_check:
    with open(f, 'rb') as fp:
        first3 = fp.read(3)
        if first3 == b'\xef\xbb\xbf':
            print(f'BOM: {f}')
        else:
            print(f'OK: {f}')
