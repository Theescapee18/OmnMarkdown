# OmnMarkdown

A powerful, multi-format document batch converter that transforms documents into clean Markdown. No Python environment required — just download and run.

## Features

- **Multi-Format Support**: Convert `.docx`, `.pdf`, `.pptx`, `.xlsx`, and `.txt` files to Markdown
- **Batch Processing**: Convert up to 10 files simultaneously
- **Image Extraction**: Automatically extract embedded images with original quality (no compression)
- **Table Preservation**: Tables in Word, PPT, and Excel are converted to Markdown tables
- **Portable Output**: Each document gets its own image folder, making it easy to share and move
- **Zero Dependencies**: Standalone `.exe` — no Python, pip, or system environment needed
- **GUI & CLI**: Desktop GUI for easy use, command-line version for automation

## Supported Formats

| Format | Extension | Features |
|--------|-----------|----------|
| Word | `.docx` | Headings, lists, tables, bold/italic, hyperlinks, images |
| PDF | `.pdf` | Page-by-page text extraction, embedded images |
| PowerPoint | `.pptx` | Slide-by-slide, text boxes, tables, images |
| Excel | `.xlsx` | Multi-sheet, Markdown tables (first row as header) |
| Text | `.txt` | Auto-detect encoding (UTF-8 / GBK / GB2312) |

## Quick Start

### For Non-Technical Users

1. Download `OmnMarkdown.exe` from the [Releases](../../releases) page
2. Double-click to run
3. Click **Add Files** to select documents
4. Choose an output folder
5. Click **Start Conversion**

### Command Line Usage

```bash
# Convert a single file
OmnMarkdown-CLI.exe input.docx -o output/

# Batch convert a folder
OmnMarkdown-CLI.exe ./documents/ -o ./output/

# Without image extraction
OmnMarkdown-CLI.exe input.docx --no-images

# Parallel processing (4 workers)
OmnMarkdown-CLI.exe ./documents/ -o ./output/ -w 4
```

### Install from Source (for Developers)

```bash
pip install -e .
```

Then use the CLI:

```bash
omnmarkdown input.docx -o output/
```

## Output Structure

Each converted document produces a `.md` file and a same-named folder for images:

```
output/
├── Report.md
├── Report/
│   └── images/
│       ├── img_0001_a1b2c3d4.png
│       └── img_0002_e5f6g7h8.png
├── Data.md
└── Data/
    └── images/
        └── img_0001_x9y8z7w6.png
```

> **Note**: The `.md` file and its same-named folder are a pair. Always move them together to keep image references intact.

## GUI Preview

The desktop app features:
- Drag-and-drop file selection
- Real-time task progress bar
- Format hint banner showing supported file types
- Resizable window (defaults to 2/3 screen)

## Building from Source

Requirements:
- Python 3.8+
- PyInstaller (for building `.exe`)

```bash
# Install dependencies
pip install python-docx Pillow PyMuPDF python-pptx openpyxl

# Build GUI version
pyinstaller --onefile --name "OmnMarkdown" --windowed \
  --add-data "word2md;word2md" \
  --hidden-import=word2md --hidden-import=word2md.converter --hidden-import=word2md.gui \
  --hidden-import=fitz --hidden-import=pptx --hidden-import=openpyxl \
  word2md/gui.py

# Build CLI version
pyinstaller --onefile --name "OmnMarkdown-CLI" --console \
  --add-data "word2md;word2md" \
  --hidden-import=word2md --hidden-import=word2md.converter --hidden-import=word2md.cli \
  --hidden-import=fitz --hidden-import=pptx --hidden-import=openpyxl \
  word2md/cli.py
```

## Project Structure

```
word2md-tool/
├── word2md/
│   ├── __init__.py
│   ├── __main__.py
│   ├── converter.py    # Core conversion engine (multi-format)
│   ├── cli.py          # Command-line interface
│   └── gui.py          # Desktop GUI (tkinter)
├── dist/
│   ├── OmnMarkdown.exe       # GUI executable (~90 MB)
│   ├── OmnMarkdown-CLI.exe   # CLI executable (~87 MB)
│   └── 使用说明书.txt         # User guide (Chinese)
├── pyproject.toml
└── README.md
```

## FAQ

**Q: Does it compress images?**  
No. All images are extracted at original resolution.

**Q: Can I share the output with others?**  
Yes. Package the `.md` file and its same-named folder together (e.g., as a `.zip`).

**Q: Does it support old `.doc` format?**  
No. Please save `.doc` files as `.docx` first using Microsoft Word.

**Q: Will image paths break if I move the files?**  
No — as long as you keep the `.md` and its same-named folder together, relative paths remain valid.

## License

MIT License

## Author

YNAU Team
# Word2MD Batch

批量将 Word 文档转换为 Markdown，支持图片提取、表格保留、多进程并行。

## 安装

```bash
# 从源码安装（推荐）
pip install .

# 或直接运行
python -m word2md <输入路径>
```

## 使用方法

```bash
# 转换单个文件
word2md report.docx

# 批量转换目录（递归）
word2md ./docs/

# 指定输出目录
word2md ./docs/ -o ./output/

# 8 线程并行（适合大量文件）
word2md ./docs/ -w 8

# 不提取图片
word2md ./docs/ --no-images

# 图片最大宽度 800px（自动缩放）
word2md ./docs/ --max-image-width 800

# 不递归子目录
word2md ./docs/ --no-recursive
```

## 参数说明

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `input` | 输入文件或目录 | 必填 |
| `-o, --output` | 输出目录 | 与源文件同目录 |
| `-w, --workers` | 并行进程数 | 4 |
| `--no-images` | 不提取图片 | false |
| `--image-dir` | 图片子目录名 | images |
| `--image-format` | 图片格式 (png/jpg/webp) | png |
| `--max-image-width` | 图片最大宽度 (0=不缩放) | 0 |
| `--heading-offset` | 标题级别偏移 | 0 |
| `--no-recursive` | 不递归子目录 | false |

## 支持特性

- ✓ 标题 (H1-H6)
- ✓ 有序/无序列表（含嵌套）
- ✓ 表格（Markdown 表格格式）
- ✓ 图片自动提取
- ✓ 加粗/斜体/下划线/删除线
- ✓ 超链接
- ✓ 引用块
- ✓ 代码样式
- ✓ 多进程并行处理
- ✓ 进度条显示

## 依赖

- Python >= 3.8
- python-docx
- Pillow

## License

MIT
