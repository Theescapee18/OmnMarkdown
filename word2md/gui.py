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
        self.root.title("OmnMarkdown - 批量转换工具")

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

        # 构建界面
        self._build_ui()

    def _build_ui(self):
        """构建界面"""
        # 主容器
        main = ttk.Frame(self.root, padding=25)
        main.pack(fill=tk.BOTH, expand=True)

        # ===== 底部：进度和按钮区域（先 pack，固定在底部） =====
        action_frame = ttk.Frame(main)
        action_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0))

        # 进度条
        self.progress = ttk.Progressbar(action_frame, mode="determinate", length=400)
        self.progress.pack(fill=tk.X, pady=(0, 8))

        # 状态标签
        self.status_label = ttk.Label(action_frame, text="就绪", foreground="gray", font=("Microsoft YaHei", 12))
        self.status_label.pack(anchor=tk.W, pady=(0, 8))

        # 开始按钮
        self.start_btn = ttk.Button(
            action_frame,
            text="  开始转换  ",
            command=self._start_conversion,
            style="Start.TButton",
        )
        self.start_btn.pack(fill=tk.X, ipady=12)

        # ===== 输出设置区域（固定在按钮上方） =====
        output_frame = ttk.LabelFrame(main, text=" 输出设置 ", padding=15)
        output_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(0, 10))

        # 输出目录
        dir_row = ttk.Frame(output_frame)
        dir_row.pack(fill=tk.X)
        ttk.Label(dir_row, text="保存位置:", font=("Microsoft YaHei", 12)).pack(side=tk.LEFT)
        self.output_entry = ttk.Entry(dir_row, textvariable=self.output_dir, width=50, font=("Microsoft YaHei", 12))
        self.output_entry.pack(side=tk.LEFT, padx=(8, 8), fill=tk.X, expand=True)
        ttk.Button(dir_row, text="  浏览  ", command=self._select_output_dir, width=8).pack(
            side=tk.LEFT
        )

        # 选项行1
        opt_row = ttk.Frame(output_frame)
        opt_row.pack(fill=tk.X, pady=(12, 0))
        ttk.Checkbutton(opt_row, text="提取图片", variable=self.extract_images).pack(
            side=tk.LEFT, padx=(0, 20)
        )
        ttk.Checkbutton(opt_row, text="保留文字颜色", variable=self.extract_colors).pack(
            side=tk.LEFT, padx=(0, 20)
        )
        ttk.Checkbutton(opt_row, text="保留高亮背景", variable=self.extract_highlights).pack(
            side=tk.LEFT, padx=(0, 20)
        )

        # 选项行2
        opt_row2 = ttk.Frame(output_frame)
        opt_row2.pack(fill=tk.X, pady=(6, 0))
        ttk.Checkbutton(opt_row2, text="提取脚注/尾注", variable=self.extract_footnotes).pack(
            side=tk.LEFT, padx=(0, 20)
        )
        ttk.Checkbutton(opt_row2, text="提取批注/旁注", variable=self.extract_comments).pack(
            side=tk.LEFT, padx=(0, 20)
        )
        ttk.Label(opt_row2, text="图片最大宽度:", font=("Microsoft YaHei", 12)).pack(side=tk.LEFT)
        ttk.Spinbox(
            opt_row2, from_=0, to=4000, increment=100, textvariable=self.max_image_width, width=8, font=("Microsoft YaHei", 12)
        ).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Label(opt_row2, text="px (0=不缩放)", foreground="gray", font=("Microsoft YaHei", 11)).pack(side=tk.LEFT, padx=(5, 0))

        # ===== 标题（固定在顶部） =====
        title_frame = ttk.Frame(main)
        title_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 8))
        ttk.Label(
            title_frame,
            text="OmnMarkdown",
            font=("Microsoft YaHei", 22, "bold"),
        ).pack(side=tk.LEFT)
        ttk.Label(
            title_frame,
            text=f"最多 {self.MAX_TASKS} 个任务",
            foreground="gray",
            font=("Microsoft YaHei", 12),
        ).pack(side=tk.RIGHT)

        # 支持格式提示
        format_hint = ttk.Label(
            main,
            text="支持格式：Word (.docx)  |  PDF (.pdf)  |  PPT (.pptx)  |  Excel (.xlsx)  |  文本 (.txt)",
            foreground="#2196F3",
            font=("Microsoft YaHei", 11),
            background="#E3F2FD",
            padding=(12, 6),
        )
        format_hint.pack(side=tk.TOP, fill=tk.X, pady=(0, 12))

        # ===== 文件选择区域（填充剩余空间） =====
        file_frame = ttk.LabelFrame(main, text=" 待转换文件 ", padding=15)
        file_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # 按钮行（先 pack 到顶部，固定位置）
        btn_frame = ttk.Frame(file_frame)
        btn_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))

        ttk.Button(btn_frame, text="  添加文件  ", command=self._add_files, width=14).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(btn_frame, text="  添加文件夹  ", command=self._add_folder, width=14).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(btn_frame, text="  移除选中  ", command=self._remove_selected, width=14).pack(
            side=tk.LEFT, padx=(0, 8)
        )
        ttk.Button(btn_frame, text="  清空  ", command=self._clear_files, width=10).pack(
            side=tk.RIGHT
        )

        # 任务计数（固定在按钮下方）
        self.task_count_label = ttk.Label(file_frame, text="当前任务: 0 / 10", font=("Microsoft YaHei", 12))
        self.task_count_label.pack(side=tk.TOP, anchor=tk.E, pady=(0, 8))

        # 文件列表（填充剩余空间）
        list_frame = ttk.Frame(file_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.file_listbox = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED,
            yscrollcommand=scrollbar.set,
            font=("Microsoft YaHei", 13),
            height=6,
        )
        self.file_listbox.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.file_listbox.yview)

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
