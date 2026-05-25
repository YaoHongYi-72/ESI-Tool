from __future__ import annotations

import os
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from windows_tool.esi_pipeline import PipelineSettings, load_settings, run_pipeline, save_settings


class EsiToolApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("ESI高水平论文统计工具")
        self.root.geometry("860x620")

        settings = load_settings()
        self.input_dir = tk.StringVar(value=settings.input_dir)
        self.template_docx = tk.StringVar(value=settings.template_docx)
        self.workspace_dir = tk.StringVar(value=settings.workspace_dir)
        self.status = tk.StringVar(value="请先选择数据文件夹、2024模板和输出目录。")

        self._build_form()
        self._build_log()

    def _build_form(self) -> None:
        wrapper = ttk.Frame(self.root, padding=16)
        wrapper.pack(fill=tk.BOTH, expand=True)

        intro = (
            "用途：把 6 期 ESI 附表自动整理成统计工作簿、详细去重总表、正式报告、统计口径说明和作者流程文档。\n"
            "当前版本按 2025 年统计口径输出，适合老师在 Windows 电脑上点选后直接生成。"
        )
        ttk.Label(wrapper, text=intro, justify=tk.LEFT).pack(anchor=tk.W, pady=(0, 12))

        self._add_picker(wrapper, "附表文件夹", self.input_dir, self.pick_input_dir, "选择包含 6 期附表的文件夹")
        self._add_picker(wrapper, "2024模板报告", self.template_docx, self.pick_template_docx, "选择 2024 年正式报告 docx")
        self._add_picker(wrapper, "输出工作目录", self.workspace_dir, self.pick_workspace_dir, "选择生成结果要保存到哪里")

        btn_row = ttk.Frame(wrapper)
        btn_row.pack(fill=tk.X, pady=(8, 12))
        ttk.Button(btn_row, text="保存路径设置", command=self.save_current_settings).pack(side=tk.LEFT)
        ttk.Button(btn_row, text="开始生成", command=self.start_pipeline).pack(side=tk.LEFT, padx=8)
        ttk.Button(btn_row, text="打开输出目录", command=self.open_workspace_dir).pack(side=tk.LEFT)

        ttk.Label(wrapper, textvariable=self.status, foreground="#1F4E78").pack(anchor=tk.W, pady=(0, 8))

    def _add_picker(self, parent, label_text, variable, command, hint_text) -> None:
        frame = ttk.Frame(parent)
        frame.pack(fill=tk.X, pady=4)
        ttk.Label(frame, text=label_text, width=12).pack(side=tk.LEFT)
        ttk.Entry(frame, textvariable=variable).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        ttk.Button(frame, text="浏览", command=command).pack(side=tk.LEFT)
        ttk.Label(parent, text=hint_text, foreground="#666666").pack(anchor=tk.W, padx=88)

    def _build_log(self) -> None:
        box = ttk.LabelFrame(self.root, text="运行日志", padding=12)
        box.pack(fill=tk.BOTH, expand=True, padx=16, pady=(0, 16))
        self.log = tk.Text(box, wrap=tk.WORD, height=20)
        self.log.pack(fill=tk.BOTH, expand=True)

    def pick_input_dir(self) -> None:
        path = filedialog.askdirectory(title="选择包含 6 期附表的文件夹")
        if path:
            self.input_dir.set(path)

    def pick_template_docx(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 2024 年模板报告",
            filetypes=[("Word 文档", "*.docx")],
        )
        if path:
            self.template_docx.set(path)

    def pick_workspace_dir(self) -> None:
        path = filedialog.askdirectory(title="选择输出工作目录")
        if path:
            self.workspace_dir.set(path)

    def save_current_settings(self) -> None:
        settings = PipelineSettings(
            input_dir=self.input_dir.get().strip(),
            template_docx=self.template_docx.get().strip(),
            workspace_dir=self.workspace_dir.get().strip(),
        )
        save_settings(settings)
        self.status.set("路径设置已保存。")

    def open_workspace_dir(self) -> None:
        target = self.workspace_dir.get().strip()
        if not target:
            messagebox.showinfo("提示", "请先选择输出工作目录。")
            return
        path = Path(target)
        if not path.exists():
            messagebox.showinfo("提示", "输出工作目录还不存在，请先运行生成。")
            return
        os.startfile(path)  # type: ignore[attr-defined]

    def append_log(self, message: str) -> None:
        self.log.insert(tk.END, message + "\n")
        self.log.see(tk.END)
        self.root.update_idletasks()

    def start_pipeline(self) -> None:
        if not self.input_dir.get().strip() or not self.template_docx.get().strip() or not self.workspace_dir.get().strip():
            messagebox.showwarning("缺少信息", "请先把三个路径都选好。")
            return
        self.save_current_settings()
        self.log.delete("1.0", tk.END)
        self.status.set("正在生成，请稍候……")

        settings = PipelineSettings(
            input_dir=self.input_dir.get().strip(),
            template_docx=self.template_docx.get().strip(),
            workspace_dir=self.workspace_dir.get().strip(),
        )

        def worker() -> None:
            try:
                result = run_pipeline(settings, self.append_log)
            except Exception as exc:  # noqa: BLE001
                self.append_log(f"[失败] {exc}")
                self.status.set("生成失败，请查看日志。")
                return
            self.append_log("生成完成。")
            self.append_log(f"统计工作簿：{result['stats_xlsx']}")
            self.append_log(f"详细去重总表：{result['detailed_xlsx']}")
            self.append_log(f"正式报告：{result['report_docx']}")
            self.append_log(f"统计口径说明：{result['process_docx']}")
            self.append_log(f"作者统计流程文档：{result['author_flow_docx']}")
            self.status.set("生成完成，可以打开输出目录查看。")

        threading.Thread(target=worker, daemon=True).start()


def main() -> None:
    root = tk.Tk()
    style = ttk.Style()
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = EsiToolApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
