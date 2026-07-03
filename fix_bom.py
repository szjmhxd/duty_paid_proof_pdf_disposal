import os

# 修复源仓库中的BOM编码
repo_dir = r'e:\test\python\pdf完税凭证文档'
files_to_fix = [
    os.path.join(repo_dir, 'extract_tax_pdf.py'),
    os.path.join(repo_dir, '可视化提取', 'app.py'),
    os.path.join(repo_dir, '算法对比测试', 'alg2_table_extract.py'),
    os.path.join(repo_dir, '算法对比测试', 'alg4_pymupdf.py'),
    os.path.join(repo_dir, '算法对比测试', 'merge_extract.py'),
    os.path.join(repo_dir, '算法对比测试', '算法2_表格提取.py'),
]

for f in files_to_fix:
    with open(f, 'rb') as fp:
        content = fp.read()
    # 移除BOM
    if content.startswith(b'\xef\xbb\xbf'):
        content = content[3:]
        with open(f, 'wb') as fp:
            fp.write(content)
        print(f'Fixed: {f}')
    else:
        print(f'OK: {f}')
