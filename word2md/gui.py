"""
OmnMarkdown 桌面版 - 批量文档转 Markdown 图形界面
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from typing import List

# 添加项目路径
sys.path.insert(0, str(Path(__file__).parent.parent))

try:
    from word2md.converter import Word2MarkdownConverter
except ImportError:
    from converter import Word2MarkdownConverter

try:
    from word2md.scanner import SensitiveWordScanner
except ImportError:
    from scanner import SensitiveWordScanner


class OmnMarkdownApp:
    """OmnMarkdown 桌面应用"""

    MAX_TASKS = 10

    # 支持的格式扩展名
    SUPPORTED_EXTS = {".docx", ".pdf", ".pptx", ".xlsx", ".txt"}
    FILETYPES = [
        ("支持的文档", "*.docx;*.pdf;*.pptx;*.xlsx;*.txt"),
        ("Word 文档", "*.docx"),
        ("PDF 文档", "*.pdf"),
        ("PPT 演示文稿", "*.pptx"),
        ("Excel 表格", "*.xlsx"),
        ("文本文件", "*.txt"),
        ("所有文件", "*.*"),
    ]

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("OmnMarkdown - 文档转换 + 敏感词扫描")

        # 获取屏幕尺寸，窗口设为 2/3
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        win_w = int(screen_w * 2 / 3)
        win_h = int(screen_h * 2 / 3)
        x = (screen_w - win_w) // 2
        y = (screen_h - win_h) // 2
        self.root.geometry(f"{win_w}x{win_h}+{x}+{y}")
        self.root.minsize(600, 450)  # 最小尺寸限制

        # 变量
        self.tasks: List[Path] = []
        self.output_dir = tk.StringVar()
        self.extract_images = tk.BooleanVar(value=True)
        self.max_image_width = tk.IntVar(value=0)
        self.extract_colors = tk.BooleanVar(value=True)
        self.extract_highlights = tk.BooleanVar(value=True)
        self.extract_footnotes = tk.BooleanVar(value=True)
        self.extract_comments = tk.BooleanVar(value=True)
        self.is_converting = False

        # 敏感词扫描变量
        self.scan_target = tk.StringVar()
        self.scan_report_dir = tk.StringVar(value=str(Path.home() / ".omnmarkdown" / "scan_reports"))
        self.new_word_entry = tk.StringVar()
        self.new_word_category = tk.StringVar(value="default")
        self.is_scanning = False

        # 构建界面
        self._build_ui()
        # 加载词库显示
        self._refresh_wordlist()

    def _build_ui(self):
        """构建界面"""
        # 主容器
        main = ttk.Frame(self.root, padding=15)
        main.pack(fill=tk.BOTH, expand=True)

        # ===== 标题 =====
        title_frame = ttk.Frame(main)
        title_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 8))
        ttk.Label(
            title_frame,
            text="OmnMarkdown",
            font=("Microsoft YaHei", 20, "bold"),
        ).pack(side=tk.LEFT)
        ttk.Label(
            title_frame,
            text="v2.2",
            foreground="gray",
            font=("Microsoft YaHei", 11),
        ).pack(side=tk.LEFT, padx=(8, 0))

        # ===== 选项卡 =====
        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # 选项卡1: 文档转换
        convert_tab = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(convert_tab, text="  文档转换  ")
        self._build_convert_tab(convert_tab)

        # 选项卡2: 敏感词扫描
        scan_tab = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(scan_tab, text="  敏感词扫描  ")
        self._build_scan_tab(scan_tab)

    def _build_convert_tab(self, parent):
        """构建文档转换选项卡"""
        # 格式提示
        format_hint = ttk.Label(
            parent,
            text="支持格式：Word (.docx)  |  PDF (.pdf)  |  PPT (.pptx)  |  Excel (.xlsx)  |  文本 (.txt)",
            foreground="#2196F3",
            font=("Microsoft YaHei", 11),
            background="#E3F2FD",
            padding=(12, 6),
        )
        format_hint.pack(side=tk.TOP, fill=tk.X, pady=(0, 8))

        # 文件列表区域
        file_frame = ttk.LabelFrame(parent, text=" 待转换文件 ", padding=12)
        file_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        btn_frame = ttk.Frame(file_frame)
        btn_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 8))
        ttk.Button(btn_frame, text="  添加文件  ", command=self._add_files, width=12).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_frame, text="  添加文件夹  ", command=self._add_folder, width=12).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_frame, text="  移除选中  ", command=self._remove_selected, width=12).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(btn_frame, text="  清空  ", command=self._clear_files, width=8).pack(side=tk.RIGHT)

        self.task_count_label = ttk.Label(file_frame, text="当前任务: 0 / 10", font=("Microsoft YaHei", 12))
        self.task_count_label.pack(side=tk.TOP, anchor=tk.E, pady=(0, 6))

        list_frame = ttk.Frame(file_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.file_listbox = tk.Listbox(list_frame, selectmode=tk.EXTENDED, yscrollcommand=scrollbar.set, font=("Microsoft YaHei", 12), height=5)
        self.file_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)

        # 输出设置
        output_frame = ttk.LabelFrame(parent, text=" 输出设置 ", padding=10)
        output_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 6))

        dir_row = ttk.Frame(output_frame)
        dir_row.pack(fill=tk.X)
        ttk.Label(dir_row, text="保存位置:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        self.output_entry = ttk.Entry(dir_row, textvariable=self.output_dir, width=40, font=("Microsoft YaHei", 11))
        self.output_entry.pack(side=tk.LEFT, padx=(6, 6), fill=tk.X, expand=True)
        ttk.Button(dir_row, text="浏览", command=self._select_output_dir, width=6).pack(side=tk.LEFT)

        opt_row = ttk.Frame(output_frame)
        opt_row.pack(fill=tk.X, pady=(8, 0))
        ttk.Checkbutton(opt_row, text="提取图片", variable=self.extract_images).pack(side=tk.LEFT, padx=(0, 14))
        ttk.Checkbutton(opt_row, text="保留文字颜色", variable=self.extract_colors).pack(side=tk.LEFT, padx=(0, 14))
        ttk.Checkbutton(opt_row, text="保留高亮背景", variable=self.extract_highlights).pack(side=tk.LEFT, padx=(0, 14))

        opt_row2 = ttk.Frame(output_frame)
        opt_row2.pack(fill=tk.X, pady=(4, 0))
        ttk.Checkbutton(opt_row2, text="提取脚注/尾注", variable=self.extract_footnotes).pack(side=tk.LEFT, padx=(0, 14))
        ttk.Checkbutton(opt_row2, text="提取批注/旁注", variable=self.extract_comments).pack(side=tk.LEFT, padx=(0, 14))

        # 进度条和按钮
        action_frame = ttk.Frame(parent)
        action_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))
        self.progress = ttk.Progressbar(action_frame, mode="determinate", length=400)
        self.progress.pack(fill=tk.X, pady=(0, 6))
        self.status_label = ttk.Label(action_frame, text="就绪", foreground="gray", font=("Microsoft YaHei", 11))
        self.status_label.pack(anchor=tk.W, pady=(0, 6))
        self.start_btn = ttk.Button(action_frame, text="  开始转换  ", command=self._start_conversion, style="Start.TButton")
        self.start_btn.pack(fill=tk.X, ipady=8)

    def _build_scan_tab(self, parent):
        """构建敏感词扫描选项卡"""
        # 上半部分：词库管理
        wordlist_frame = ttk.LabelFrame(parent, text=" 敏感词库管理 ", padding=12)
        wordlist_frame.pack(fill=tk.X, pady=(0, 8))

        # 添加词行
        add_row = ttk.Frame(wordlist_frame)
        add_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(add_row, text="添加敏感词:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        ttk.Entry(add_row, textvariable=self.new_word_entry, width=30, font=("Microsoft YaHei", 11)).pack(side=tk.LEFT, padx=(6, 6))
        ttk.Label(add_row, text="分类:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        ttk.Entry(add_row, textvariable=self.new_word_category, width=10, font=("Microsoft YaHei", 11)).pack(side=tk.LEFT, padx=(6, 6))
        ttk.Button(add_row, text="添加", command=self._add_sensitive_word, width=6).pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(add_row, text="删除选中", command=self._remove_sensitive_word, width=8).pack(side=tk.LEFT, padx=(6, 0))

        # 词库列表
        list_frame = ttk.Frame(wordlist_frame)
        list_frame.pack(fill=tk.X)
        wl_scroll = ttk.Scrollbar(list_frame)
        wl_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.wordlist_box = tk.Listbox(list_frame, yscrollcommand=wl_scroll.set, font=("Microsoft YaHei", 11), height=5, selectmode=tk.EXTENDED)
        self.wordlist_box.pack(fill=tk.X)
        wl_scroll.config(command=self.wordlist_box.yview)
        ttk.Button(wordlist_frame, text="刷新词库", command=self._refresh_wordlist, width=10).pack(anchor=tk.E, pady=(6, 0))

        # 下半部分：扫描操作
        scan_frame = ttk.LabelFrame(parent, text=" 扫描操作 ", padding=12)
        scan_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        target_row = ttk.Frame(scan_frame)
        target_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(target_row, text="扫描目标:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        ttk.Entry(target_row, textvariable=self.scan_target, width=40, font=("Microsoft YaHei", 11)).pack(side=tk.LEFT, padx=(6, 6), fill=tk.X, expand=True)
        ttk.Button(target_row, text="选择文件", command=self._select_scan_files, width=8).pack(side=tk.LEFT)
        ttk.Button(target_row, text="选择目录", command=self._select_scan_target, width=8).pack(side=tk.LEFT, padx=(4, 0))

        # 扫描结果
        result_frame = ttk.Frame(scan_frame)
        result_frame.pack(fill=tk.BOTH, expand=True)
        res_scroll = ttk.Scrollbar(result_frame)
        res_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.scan_result_text = tk.Text(result_frame, yscrollcommand=res_scroll.set, font=("Microsoft YaHei", 11), height=10, state=tk.DISABLED)
        self.scan_result_text.pack(fill=tk.BOTH, expand=True)
        res_scroll.config(command=self.scan_result_text.yview)

        # 报告保存路径
        report_row = ttk.Frame(scan_frame)
        report_row.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(report_row, text="报告保存:", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT)
        ttk.Entry(report_row, textvariable=self.scan_report_dir, width=40, font=("Microsoft YaHei", 10)).pack(side=tk.LEFT, padx=(6, 6), fill=tk.X, expand=True)
        ttk.Button(report_row, text="选择目录", command=self._select_report_dir, width=8).pack(side=tk.LEFT)

        # 扫描按钮
        scan_action = ttk.Frame(parent)
        scan_action.pack(side=tk.BOTTOM, fill=tk.X)
        self.scan_btn = ttk.Button(scan_action, text="  开始扫描  ", command=self._start_scan, style="Start.TButton")
        self.scan_btn.pack(fill=tk.X, ipady=8)

    def _update_task_count(self):
        """更新任务计数"""
        count = len(self.tasks)
        self.task_count_label.config(
            text=f"当前任务: {count} / {self.MAX_TASKS}",
            foreground="red" if count >= self.MAX_TASKS else "black",
        )

    def _add_files(self):
        """添加文件"""
        files = filedialog.askopenfilenames(
            title="选择文档",
            filetypes=self.FILETYPES,
        )
        for f in files:
            path = Path(f)
            if path.suffix.lower() not in self.SUPPORTED_EXTS:
                messagebox.showwarning("不支持的格式", f"{path.name}\n暂不支持 {path.suffix} 格式，目前支持: .docx")
                continue
            if path not in self.tasks and len(self.tasks) < self.MAX_TASKS:
                self.tasks.append(path)
                self.file_listbox.insert(tk.END, str(path.name))

        if len(self.tasks) >= self.MAX_TASKS:
            messagebox.showwarning("提示", f"已达到最大任务数 {self.MAX_TASKS}")

        self._update_task_count()

    def _add_folder(self):
        """添加整个文件夹"""
        folder = filedialog.askdirectory(title="选择文件夹")
        if not folder:
            return

        added = 0
        for ext in self.SUPPORTED_EXTS:
            for f in Path(folder).rglob(f"*{ext}"):
                if f not in self.tasks and len(self.tasks) < self.MAX_TASKS:
                    self.tasks.append(f)
                    self.file_listbox.insert(tk.END, f"{f.parent.name}/{f.name}")
                    added += 1

        if added == 0:
            ext_list = ", ".join(self.SUPPORTED_EXTS)
            messagebox.showinfo("提示", f"文件夹中未找到支持的文件 ({ext_list}) 或已达到上限")
        elif len(self.tasks) >= self.MAX_TASKS:
            messagebox.showwarning("提示", f"已达到最大任务数 {self.MAX_TASKS}，添加了 {added} 个文件")

        self._update_task_count()

    def _remove_selected(self):
        """移除选中的文件"""
        selected = list(self.file_listbox.curselection())
        for i in reversed(selected):
            self.file_listbox.delete(i)
            del self.tasks[i]
        self._update_task_count()

    def _clear_files(self):
        """清空所有文件"""
        self.file_listbox.delete(0, tk.END)
        self.tasks.clear()
        self._update_task_count()

    def _select_output_dir(self):
        """选择输出目录"""
        folder = filedialog.askdirectory(title="选择输出文件夹")
        if folder:
            self.output_dir.set(folder)

    def _start_conversion(self):
        """开始转换"""
        if not self.tasks:
            messagebox.showwarning("提示", "请先添加要转换的文件")
            return

        if not self.output_dir.get():
            messagebox.showwarning("提示", "请选择输出文件夹")
            return

        output_path = Path(self.output_dir.get())
        if not output_path.exists():
            output_path.mkdir(parents=True, exist_ok=True)

        self.is_converting = True
        self.start_btn.config(state=tk.DISABLED, text="转换中...")
        self.progress["value"] = 0
        self.progress["maximum"] = len(self.tasks)

        # 在后台线程执行转换
        thread = threading.Thread(target=self._do_conversion, daemon=True)
        thread.start()

    def _do_conversion(self):
        """执行转换（后台线程）"""
        converter = Word2MarkdownConverter(
            extract_images=self.extract_images.get(),
            max_image_width=self.max_image_width.get(),
            extract_colors=self.extract_colors.get(),
            extract_highlights=self.extract_highlights.get(),
            extract_footnotes=self.extract_footnotes.get(),
            extract_comments=self.extract_comments.get(),
        )

        output_path = Path(self.output_dir.get())
        success = 0
        failed = 0
        total = len(self.tasks)

        for i, task in enumerate(self.tasks):
            try:
                out_file = output_path / task.with_suffix(".md").name
                self.root.after(
                    0,
                    lambda t=task.name, idx=i: self.status_label.config(
                        text=f"[{idx+1}/{total}] 正在转换: {t}"
                    ),
                )
                converter.convert_file(str(task), str(out_file))
                success += 1
            except Exception as e:
                failed += 1
                self.root.after(
                    0,
                    lambda t=task.name, err=str(e): self.status_label.config(
                        text=f"✗ 失败: {t} - {err}", foreground="red"
                    ),
                )

            self.root.after(0, lambda v=i+1: self.progress.config(value=v))

        # 完成
        self.root.after(
            0,
            lambda: self._conversion_done(success, failed, total, str(output_path)),
        )

    def _conversion_done(self, success: int, failed: int, total: int, output_dir: str):
        """转换完成"""
        self.is_converting = False
        self.start_btn.config(state=tk.NORMAL, text="▶ 开始转换")
        self.status_label.config(
            text=f"✓ 完成! 成功: {success}  失败: {failed}  总计: {total}",
            foreground="green" if failed == 0 else "orange",
        )

        result_msg = f"转换完成!\n\n成功: {success}\n失败: {failed}\n输出目录: {output_dir}"
        if messagebox.askyesno("完成", result_msg + "\n\n是否打开输出文件夹?"):
            os.startfile(output_dir)

    # ================================================================
    # 敏感词扫描功能
    # ================================================================

    def _add_sensitive_word(self):
        """添加敏感词"""
        text = self.new_word_entry.get().strip()
        if not text:
            messagebox.showwarning("提示", "请输入敏感词")
            return

        # 支持逗号/空格分隔多个词
        words = [w.strip() for w in text.replace("，", ",").split(",") if w.strip()]
        category = self.new_word_category.get().strip() or "default"

        scanner = SensitiveWordScanner()
        added = scanner.add_words(words, category=category)
        self.new_word_entry.set("")
        messagebox.showinfo("成功", f"新增 {added} 个敏感词（分类: {category}）\n词库共 {scanner.get_word_count()} 个")
        self._refresh_wordlist()

    def _remove_sensitive_word(self):
        """删除选中的敏感词"""
        selected = list(self.wordlist_box.curselection())
        if not selected:
            messagebox.showwarning("提示", "请先选择要删除的词")
            return

        words_to_remove = [self.wordlist_box.get(i) for i in selected]
        # 去掉序号前缀
        clean_words = []
        for w in words_to_remove:
            parts = w.split(". ", 1)
            if len(parts) == 2:
                clean_words.append(parts[1])
            else:
                clean_words.append(w)

        scanner = SensitiveWordScanner()
        removed = scanner.remove_words(clean_words)
        messagebox.showinfo("成功", f"删除 {removed} 个敏感词")
        self._refresh_wordlist()

    def _refresh_wordlist(self):
        """刷新词库显示"""
        self.wordlist_box.delete(0, tk.END)
        scanner = SensitiveWordScanner()
        categories = scanner.list_words()
        idx = 1
        for cat, words in sorted(categories.items()):
            for w in words:
                self.wordlist_box.insert(tk.END, f"{idx}. {w}  [{cat}]")
                idx += 1
        if idx == 1:
            self.wordlist_box.insert(tk.END, "（词库为空，请添加敏感词）")

    def _select_scan_target(self):
        """选择扫描目标目录"""
        folder = filedialog.askdirectory(title="选择要扫描的目录")
        if folder:
            self.scan_target.set(folder)

    def _select_scan_files(self):
        """选择扫描目标文件"""
        files = filedialog.askopenfilenames(
            title="选择要扫描的文件",
            filetypes=[
                ("支持的文档", "*.docx;*.pdf;*.pptx;*.xlsx;*.txt;*.md"),
                ("Word 文档", "*.docx"),
                ("Markdown", "*.md"),
                ("PDF 文档", "*.pdf"),
                ("所有文件", "*.*"),
            ],
        )
        if files:
            self.scan_target.set(";".join(files))

    def _select_report_dir(self):
        """选择报告保存目录"""
        folder = filedialog.askdirectory(title="选择报告保存目录")
        if folder:
            self.scan_report_dir.set(folder)

    def _start_scan(self):
        """开始扫描"""
        target = self.scan_target.get().strip()
        if not target:
            messagebox.showwarning("提示", "请选择要扫描的文件或目录")
            return

        # 多文件模式（分号分隔）跳过路径检查
        if ";" not in target:
            if not Path(target).exists():
                messagebox.showerror("错误", f"路径不存在: {target}")
                return

        scanner = SensitiveWordScanner()
        if scanner.get_word_count() == 0:
            messagebox.showwarning("提示", "敏感词库为空，请先在词库管理中添加敏感词")
            return

        self.is_scanning = True
        self.scan_btn.config(state=tk.DISABLED, text="扫描中...")
        self.scan_result_text.config(state=tk.NORMAL)
        self.scan_result_text.delete("1.0", tk.END)
        self.scan_result_text.insert(tk.END, "正在扫描，请稍候...\n")
        self.scan_result_text.config(state=tk.DISABLED)

        report_dir = self.scan_report_dir.get().strip()
        thread = threading.Thread(target=self._do_scan, args=(target, report_dir), daemon=True)
        thread.start()

    def _do_scan(self, target: str, report_dir: str = None):
        """执行扫描（后台线程）"""
        scanner = SensitiveWordScanner(report_dir=report_dir if report_dir else None)

        # 支持多文件（分号分隔）或目录
        if ";" in target:
            file_list = [f.strip() for f in target.split(";") if f.strip()]
            # 逐个文件扫描
            all_hits = []
            scanned_files = 0
            for fp in file_list:
                if Path(fp).exists():
                    scanned_files += 1
                    hits = scanner.scan_file(fp)
                    all_hits.extend(hits)

            # 构建报告
            file_summary = {}
            for hit in all_hits:
                fname = hit["filename"]
                if fname not in file_summary:
                    file_summary[fname] = {"file": hit["file"], "hits": [], "word_set": set()}
                file_summary[fname]["hits"].append(hit)
                file_summary[fname]["word_set"].add(hit["word"])

            word_freq = {}
            for hit in all_hits:
                w = hit["word"]
                word_freq[w] = word_freq.get(w, 0) + 1

            report = {
                "scan_time": __import__('datetime').datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "directory": f"{len(file_list)} 个文件",
                "total_files_scanned": scanned_files,
                "files_with_hits": len(file_summary),
                "total_hits": len(all_hits),
                "unique_words_found": list(word_freq.keys()),
                "word_frequency": dict(sorted(word_freq.items(), key=lambda x: -x[1])),
                "file_summary": [{"filename": k, "file": v["file"], "hit_count": len(v["hits"]), "unique_words": sorted(v["word_set"])} for k, v in sorted(file_summary.items(), key=lambda x: -len(x[1]["hits"]))],
                "details": all_hits,
            }
        else:
            report = scanner.scan_directory(target)

        # 保存报告
        json_path = scanner.save_report(report)
        csv_path = scanner.save_report_csv(report)

        # 构建显示文本
        lines = []
        lines.append(f"扫描完成：{report['scan_time']}")
        lines.append(f"扫描文件: {report['total_files_scanned']} 个")
        lines.append(f"命中文件: {report['files_with_hits']} 个")
        lines.append(f"总命中数: {report['total_hits']} 次")
        lines.append("")

        if report['file_summary']:
            lines.append("━━━ 命中文件详情 ━━━")
            for item in report['file_summary']:
                lines.append(f"  [{item['hit_count']}次] {item['filename']}")
                lines.append(f"    命中词: {', '.join(item['unique_words'])}")
            lines.append("")

        if report['details']:
            lines.append("━━━ 详细记录 ━━━")
            for hit in report['details'][:100]:
                lines.append(f"  {hit['filename']}:{hit['line']}  [{hit['word']}]")
                lines.append(f"    {hit['context']}")
            if len(report['details']) > 100:
                lines.append(f"  ... 还有 {len(report['details'])-100} 条记录")
            lines.append("")

        lines.append(f"报告已保存:")
        lines.append(f"  JSON: {json_path}")
        lines.append(f"  CSV:  {csv_path}")

        result_text = "\n".join(lines)

        self.root.after(0, lambda: self._scan_done(result_text, report))

    def _scan_done(self, result_text: str, report: dict):
        """扫描完成"""
        self.is_scanning = False
        self.scan_btn.config(state=tk.NORMAL, text="  开始扫描  ")
        self.scan_result_text.config(state=tk.NORMAL)
        self.scan_result_text.delete("1.0", tk.END)
        self.scan_result_text.insert(tk.END, result_text)
        self.scan_result_text.config(state=tk.DISABLED)

        if report['total_hits'] > 0:
            messagebox.showwarning(
                "扫描完成",
                f"发现 {report['total_hits']} 处敏感词命中！\n"
                f"涉及 {report['files_with_hits']} 个文件\n\n"
                f"报告已保存至本地，详见结果面板。"
            )
        else:
            messagebox.showinfo("扫描完成", "未发现敏感词，所有文件安全。")


def main():
    root = tk.Tk()

    # 设置 DPI 感知（Windows 高 DPI 屏幕）
    try:
        from ctypes import windll
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    # 设置 ttk 主题
    style = ttk.Style()
    style.theme_use("clam")

    # 全局字体配置
    FONT_LABEL = ("Microsoft YaHei", 12)
    FONT_BUTTON = ("Microsoft YaHei", 12)
    FONT_TITLE = ("Microsoft YaHei", 22, "bold")
    FONT_LIST = ("Microsoft YaHei", 13)
    FONT_STATUS = ("Microsoft YaHei", 12)
    FONT_START_BTN = ("Microsoft YaHei", 16, "bold")

    style.configure("TLabel", font=FONT_LABEL)
    style.configure("TButton", font=FONT_BUTTON, padding=8)
    style.configure("TCheckbutton", font=FONT_LABEL)
    style.configure("TLabelframe.Label", font=("Microsoft YaHei", 13, "bold"))
    style.configure("TEntry", font=FONT_LABEL, padding=6)
    style.configure("TSpinbox", font=FONT_LABEL, padding=6)
    style.configure("Horizontal.TProgressbar", thickness=24)
    style.configure("Start.TButton", font=("Microsoft YaHei", 16, "bold"), padding=15)

    app = OmnMarkdownApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
