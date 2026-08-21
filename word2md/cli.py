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
from word2md.scanner import SensitiveWordScanner


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


def handle_scan(args):
    """处理 scan 子命令：批量扫描敏感词"""
    report_dir = args.output if hasattr(args, 'output') and args.output else None
    scanner = SensitiveWordScanner(
        wordlist_path=args.wordlist if hasattr(args, 'wordlist') and args.wordlist else None,
        report_dir=report_dir,
    )

    if scanner.get_word_count() == 0:
        print("\n[警告] 敏感词库为空，请先添加敏感词：")
        print("  word2md words add 敏感词1 敏感词2 ...")
        sys.exit(1)

    target = Path(args.target)
    if not target.exists():
        print(f"[错误] 路径不存在: {target}")
        sys.exit(1)

    extensions = args.ext.split(",") if args.ext else None

    print(f"\n{'='*60}")
    print(f"  OmnMarkdown 敏感词扫描")
    print(f"  目标: {target}")
    print(f"  词库: {scanner.get_word_count()} 个敏感词")
    print(f"{'='*60}\n")

    report = scanner.scan_directory(
        str(target),
        extensions=extensions,
        recursive=not args.no_recursive,
    )

    if "error" in report:
        print(f"[错误] {report['error']}")
        sys.exit(1)

    # 输出结果
    print(f"  扫描文件: {report['total_files_scanned']} 个")
    print(f"  命中文件: {report['files_with_hits']} 个")
    print(f"  总命中数: {report['total_hits']} 次\n")

    if report['file_summary']:
        print("  命中文件详情:")
        for item in report['file_summary']:
            words_str = ", ".join(item['unique_words'][:5])
            if len(item['unique_words']) > 5:
                words_str += f"... (+{len(item['unique_words'])-5})"
            print(f"    [{item['hit_count']}次] {item['filename']}")
            print(f"           命中词: {words_str}")
        print()

    if report['details'] and args.show_details:
        print("  详细记录:")
        for hit in report['details'][:50]:
            print(f"    {hit['filename']}:{hit['line']} [{hit['word']}] {hit['context']}")
        if len(report['details']) > 50:
            print(f"    ... 还有 {len(report['details'])-50} 条记录")
        print()

    # 保存报告
    json_path = scanner.save_report(report)
    csv_path = scanner.save_report_csv(report)
    print(f"  报告已保存:")
    print(f"    JSON: {json_path}")
    print(f"    CSV:  {csv_path}")
    print(f"{'='*60}\n")


def handle_words(args):
    """处理 words 子命令：管理敏感词库"""
    scanner = SensitiveWordScanner()

    if args.action == "add":
        if not args.words:
            print("[错误] 请指定要添加的敏感词")
            sys.exit(1)
        category = args.category or "default"
        added = scanner.add_words(args.words, category=category)
        print(f"  ✓ 新增 {added} 个敏感词（分类: {category}），词库共 {scanner.get_word_count()} 个")

    elif args.action == "remove":
        if not args.words:
            print("[错误] 请指定要删除的敏感词")
            sys.exit(1)
        removed = scanner.remove_words(args.words)
        print(f"  ✓ 删除 {removed} 个敏感词，词库剩余 {scanner.get_word_count()} 个")

    elif args.action == "list":
        categories = scanner.list_words()
        if not categories or all(len(v) == 0 for v in categories.values()):
            print("  敏感词库为空")
            return
        total = 0
        for cat, words in sorted(categories.items()):
            if words:
                print(f"\n  [{cat}] ({len(words)} 个):")
                for i, w in enumerate(words, 1):
                    print(f"    {i}. {w}")
                total += len(words)
        print(f"\n  共 {total} 个敏感词")
        print(f"  词库路径: {scanner.wordlist_path}")


def main():
    # 检测是否为子命令模式
    if len(sys.argv) > 1 and sys.argv[1] in ("scan", "words"):
        # 子命令模式
        parser = argparse.ArgumentParser(
            prog="word2md",
            description="OmnMarkdown - 文档转换 + 敏感词扫描工具",
        )
        subparsers = parser.add_subparsers(dest="command")

        # scan 子命令
        scan_parser = subparsers.add_parser("scan", help="批量扫描文件中的敏感词")
        scan_parser.add_argument("target", help="要扫描的文件或目录路径")
        scan_parser.add_argument("--wordlist", help="自定义词库路径（JSON）")
        scan_parser.add_argument("--ext", help="文件扩展名，逗号分隔（如 .docx,.md,.txt）")
        scan_parser.add_argument("--no-recursive", action="store_true", help="不递归扫描子目录")
        scan_parser.add_argument("-d", "--show-details", action="store_true", help="显示详细命中记录")
        scan_parser.add_argument("-o", "--output", help="报告保存目录（默认 ~/.omnmarkdown/scan_reports）")

        # words 子命令
        words_parser = subparsers.add_parser("words", help="管理敏感词库")
        words_parser.add_argument("action", choices=["add", "remove", "list"], help="操作类型")
        words_parser.add_argument("words", nargs="*", help="敏感词列表")
        words_parser.add_argument("-c", "--category", help="分类名称（默认 default）")

        args = parser.parse_args()

        if args.command == "scan":
            handle_scan(args)
        elif args.command == "words":
            handle_words(args)
        else:
            parser.print_help()
        return

    # 默认模式：文档转换
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
  word2md scan ./docs/ -d                # 扫描敏感词
  word2md words add 机密 内部            # 添加敏感词
  word2md words list                     # 查看词库
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
    parser.add_argument("-v", "--version", action="version", version="OmnMarkdown 2.2.0")

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
