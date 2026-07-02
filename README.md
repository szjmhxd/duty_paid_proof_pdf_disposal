# 完税证明PDF批量提取工具

## 概述

从国家税务总局完税证明PDF文件中批量提取税务信息，支持多层文件夹递归扫描、MD5去重、忽略合并文件，输出Excel和SQL。

---

## 目录结构

```
pdf完税凭证文档/
├── README.md
├── extract_tax_pdf.py                 # 基础提取脚本（PDF→XML/CSV/TXT）
├── 可视化提取/
│   ├── app.py                         # GUI可视化版（主程序）
│   └── dist/TaxExtractor.exe          # 打包好的可执行文件
├── 算法对比测试/
│   ├── alg2_table_extract.py          # 算法2：pdfplumber表格提取
│   ├── alg4_pymupdf.py                # 算法4：PyMuPDF结构解析
│   ├── alg5_pdfminer.py               # 算法5：pdfminer.six布局分析
│   ├── merge_extract.py               # 三算法合并提取脚本
│   └── output/                        # 算法对比测试结果
├── PDF数据目录/                       # PDF数据目录（请替换为实际路径）
└── 2026年社保完税证明/                  # PDF数据目录（多层子文件夹）
```

---

## 三种PDF解析算法对比

| 算法 | 库 | 耗时 | 优点 | 缺点 |
|------|-----|------|------|------|
| 算法2 | pdfplumber | ~0.22s | 表格结构提取，输出结构化数据 | 合并单元格字段映射混乱 |
| 算法4 | PyMuPDF | ~0.48s | 文本按行合并清晰，含坐标字体 | 无法区分表格行列 |
| 算法5 | pdfminer.six | ~0.18s | 布局分析最详细，按阅读顺序排列 | 输出格式较乱 |

最终选择 **算法2（pdfplumber）** 作为主提取算法，配合正则解析。

---

## 功能特性

### 1. 基础脚本 (extract_tax_pdf.py)

PDF→XML/CSV/TXT，无GUI。

```bash
python extract_tax_pdf.py
```

### 2. 命令行版 (算法对比测试/merge_extract.py)

支持递归子文件夹、MD5去重、忽略合并文件，输出Excel+SQL。

```bash
python merge_extract.py
python merge_extract.py "E:\path\to\pdfs"
```

### 3. GUI可视化版 (可视化提取/app.py)

图形界面：选择文件夹、递归/去重/忽略合并选项、实时日志、自定义输出。

```bash
python app.py
```

### 4. 打包为exe（无需Python环境）

```bash
pyinstaller --noconfirm --onefile --windowed --name "TaxExtractor" app.py
```

---

## 去重/递归/忽略合并演进

### 去重

- **初版**：文件名`(1)`后缀判断（不可靠）
- **最终版**：MD5 hash判断内容相同（可靠）

### 递归子文件夹

- **初版**：`glob("*.pdf")` 只扫描当前目录
- **最终版**：`glob("**/*.pdf", recursive=True)` 递归所有子文件夹

### 忽略合并文件

```python
if '合并' in os.path.basename(f):
    continue
```

---

## 输出字段

### 主表 invoice_main

| 字段 | 类型 | 说明 |
|------|------|------|
| invoice_number | varchar(64) | 完税证明编号 |
| amount_tax | decimal(20,2) | 完税金额合计 |
| 文件名称 | varchar(255) | PDF源文件名 |
| 文件路径 | varchar(255) | 相对路径（递归时显示子目录） |

### 附加表 invoice_detail

| 字段 | 类型 | 说明 |
|------|------|------|
| voucher_number | varchar(255) | 原凭证号 |
| tax_categories | varchar(255) | 税种 |
| items_name | varchar(255) | 品目名称 |
| start_date | date | 税款所属时期(始) |
| end_date | date | 税款所属时期(终) |
| storage_date | date | 入(退)库日期 |
| amount | decimal(20,2) | 实缴(退)金额 |

---

## 技术栈

- Python 3.12+
- pdfplumber / PyMuPDF / pdfminer.six
- openpyxl（Excel读写）
- tkinter（GUI界面）

## 依赖安装

```bash
pip install pdfplumber PyMuPDF pdfminer.six openpyxl
```
