"""
Word2MD 批量转换 CLI
用法: word2md <输入路径> [选项]
"""

import argparse
import sys
import time
import os
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed
from typing import List, Tuple

from word2md.converter import Word2MarkdownConverter


def find_docx_files(input_path: Path, recursive: bool = True) -> List[Path]:
    """查找所有 .docx 文件"""
    if input_path.is_file():
        if input_path.suffix.lower() in (".docx", ".doc"):
            return [input_path]
        else:
            print(f"[错误] 不支持的文件格式: {input_path.suffix}")
            return []

    if input_path.is_dir():
        pattern = "**/*.docx" if recursive else "*.docx"
        files = sorted(input_path.glob(pattern))
        if not files:
            print(f"[警告] 目录中未找到 .docx 文件: {input_path}")
        return files

    print(f"[错误] 路径不存在: {input_path}")
    return []


def convert_single_file(args: Tuple[str, str, dict]) -> Tuple[str, bool, str]:
    """转换单个文件（用于多进程）"""
    input_path, output_path, options = args
    try:
        converter = Word2MarkdownConverter(
            extract_images=options.get("extract_images", True),
            image_format=options.get("image_format", "png"),
            image_dir=options.get("image_dir", "images"),
            heading_offset=options.get("heading_offset", 0),
            max_image_width=options.get("max_image_width", 0),
            extract_colors=options.get("extract_colors", True),
            extract_highlights=options.get("extract_highlights", True),
            extract_footnotes=options.get("extract_footnotes", True),
            extract_comments=options.get("extract_comments", True),
        )
        result = converter.convert_file(input_path, output_path)
        return (input_path, True, result)
    except Exception as e:
        return (input_path, False, str(e))


def progress_bar(current: int, total: int, width: int = 40) -> str:
    """生成进度条"""
    pct = current / total if total > 0 else 1
    filled = int(width * pct)
    bar = "█" * filled + "░" * (width - filled)
    return f"[{bar}] {current}/{total} ({pct:.0%})"


def run_batch(files: List[Path], output_dir: Path, options: dict, workers: int):
    """批量转换"""
    tasks = []
    for f in files:
        if output_dir:
            # 保持子目录结构
            rel = f.relative_to(files[0].parent) if len(files) > 1 else Path(f.name)
            out = output_dir / rel.with_suffix(".md")
        else:
            out = str(f.with_suffix(".md"))
        tasks.append((str(f), str(out), options))

    total = len(tasks)
    success = 0
    failed = 0
    start_time = time.time()

    print(f"\n{'='*60}")
    print(f"  Word2MD 批量转换")
    print(f"  文件数: {total}  |  并行数: {workers}")
    print(f"{'='*60}\n")

    if workers > 1 and total > 1:
        with ProcessPoolExecutor(max_workers=workers) as executor:
            futures = {executor.submit(convert_single_file, t): t[0] for t in tasks}
            done_count = 0
            for future in as_completed(futures):
                done_count += 1
                input_path, ok, result = future.result()
                fname = Path(input_path).name

                if ok:
                    success += 1
                    print(f"  ✓ {fname}")
                else:
                    failed += 1
                    print(f"  ✗ {fname} -> {result}")

                print(f"    {progress_bar(done_count, total)}")
    else:
        for i, task in enumerate(tasks):
            input_path, ok, result = convert_single_file(task)
            fname = Path(input_path).name

            if ok:
                success += 1
                print(f"  ✓ {fname}")
            else:
                failed += 1
                print(f"  ✗ {fname} -> {result}")

            print(f"    {progress_bar(i + 1, total)}")

    elapsed = time.time() - start_time
    print(f"\n{'='*60}")
    print(f"  完成! 成功: {success}  失败: {failed}  耗时: {elapsed:.1f}s")
    print(f"{'='*60}\n")


def main():
    parser = argparse.ArgumentParser(
        prog="word2md",
        description="批量将 Word 文档转换为 Markdown（支持图片提取和表格保留）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  word2md report.docx                    # 转换单个文件
  word2md ./docs/                        # 批量转换目录
  word2md ./docs/ -o ./output/           # 指定输出目录
  word2md ./docs/ -w 8                   # 8 线程并行
  word2md ./docs/ --no-images            # 不提取图片
  word2md ./docs/ --max-image-width 800  # 图片最大宽度 800px
        """,
    )

    parser.add_argument("input", help="输入文件或目录路径")
    parser.add_argument("-o", "--output", help="输出目录（默认与源文件同目录）")
    parser.add_argument("-w", "--workers", type=int, default=4, help="并行处理线程数（默认 4）")
    parser.add_argument("--no-images", action="store_true", help="不提取图片")
    parser.add_argument("--image-dir", default="images", help="图片保存子目录名（默认 images）")
    parser.add_argument("--image-format", default="png", choices=["png", "jpg", "webp"], help="图片格式（默认 png）")
    parser.add_argument("--max-image-width", type=int, default=0, help="图片最大宽度（0=不缩放）")
    parser.add_argument("--heading-offset", type=int, default=0, help="标题级别偏移（如 1 表示 H1→H2）")
    parser.add_argument("--no-recursive", action="store_true", help="不递归扫描子目录")
    parser.add_argument("--no-color", action="store_true", help="不保留文字颜色（默认保留）")
    parser.add_argument("--no-highlight", action="store_true", help="不保留高亮/背景色（默认保留）")
    parser.add_argument("--no-footnotes", action="store_true", help="不提取脚注/尾注（默认提取）")
    parser.add_argument("--no-comments", action="store_true", help="不提取批注/旁注（默认提取）")
    parser.add_argument("-v", "--version", action="version", version="OmnMarkdown 2.1.0")

    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output) if args.output else None

    if output_dir:
        output_dir.mkdir(parents=True, exist_ok=True)

    files = find_docx_files(input_path, recursive=not args.no_recursive)
    if not files:
        sys.exit(1)

    options = {
        "extract_images": not args.no_images,
        "image_format": args.image_format,
        "image_dir": args.image_dir,
        "heading_offset": args.heading_offset,
        "max_image_width": args.max_image_width,
        "extract_colors": not args.no_color,
        "extract_highlights": not args.no_highlight,
        "extract_footnotes": not args.no_footnotes,
        "extract_comments": not args.no_comments,
    }

    if len(files) == 1 and not output_dir:
        # 单文件直接转换
        print(f"\n转换: {files[0].name}")
        input_path_str, ok, result = convert_single_file(
            (str(files[0]), None, options)
        )
        if ok:
            print(f"  ✓ 输出: {result}")
        else:
            print(f"  ✗ 错误: {result}")
            sys.exit(1)
    else:
        run_batch(files, output_dir, options, args.workers)


if __name__ == "__main__":
    main()
