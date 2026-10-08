from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .excel_to_markdown import create_markdown, find_header_row
from .markdown_to_excel import process_markdown
from openpyxl import load_workbook


class ConverterWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("md-xlsx-convertor")
        self.root.geometry("760x430")
        self.root.minsize(650, 400)
        self.mode = tk.StringVar(value="xlsx_to_md")
        self.input_path = tk.StringVar()
        self.workbook_path = tk.StringVar()
        self.sheet_name = tk.StringVar()
        self.start_row = tk.StringVar()
        self.overwrite = tk.BooleanVar(value=False)
        self.status = tk.StringVar(value="Select an operation and input file.")
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self._build_ui()
        self._update_mode()
        self.root.after(100, self._poll_events)

    def _build_ui(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("vista" if "vista" in style.theme_names() else "clam")
        style.configure("App.TFrame", background="#f1f4f2")
        style.configure("Panel.TFrame", background="#ffffff")
        style.configure("Title.TLabel", background="#f1f4f2", foreground="#18302a", font=("Yu Gothic UI", 19, "bold"))
        style.configure("Subtitle.TLabel", background="#f1f4f2", foreground="#53645e", font=("Yu Gothic UI", 10))
        style.configure("Panel.TLabel", background="#ffffff", foreground="#24332e", font=("Yu Gothic UI", 10))
        style.configure("PanelTitle.TLabel", background="#ffffff", foreground="#18302a", font=("Yu Gothic UI", 11, "bold"))
        style.configure("Accent.TButton", font=("Yu Gothic UI", 10, "bold"), padding=(18, 9))

        outer = ttk.Frame(self.root, style="App.TFrame", padding=(28, 24))
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="テスト仕様書変換ツール", style="Title.TLabel").pack(anchor="w")
        ttk.Label(outer, text="Excelとmael Markdownを変換・更新します。", style="Subtitle.TLabel").pack(anchor="w", pady=(4, 18))

        panel = ttk.Frame(outer, style="Panel.TFrame", padding=20)
        panel.pack(fill="both", expand=True)
        panel.columnconfigure(1, weight=1)
        ttk.Label(panel, text="処理", style="PanelTitle.TLabel").grid(row=0, column=0, sticky="w", columnspan=3)
        choices = ttk.Frame(panel, style="Panel.TFrame")
        choices.grid(row=1, column=0, columnspan=3, sticky="w", pady=(10, 18))
        ttk.Radiobutton(choices, text="ExcelからMarkdownを生成", value="xlsx_to_md", variable=self.mode, command=self._update_mode).pack(side="left", padx=(0, 24))
        ttk.Radiobutton(choices, text="MarkdownからExcelを更新", value="md_to_xlsx", variable=self.mode, command=self._update_mode).pack(side="left")

        self.input_label = ttk.Label(panel, text="対象Excel", style="Panel.TLabel")
        self.input_label.grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(panel, textvariable=self.input_path).grid(row=2, column=1, sticky="ew", padx=(14, 8), pady=6)
        self.browse_input = ttk.Button(panel, text="参照...", command=self._browse_input)
        self.browse_input.grid(row=2, column=2, sticky="ew", pady=6)

        ttk.Label(panel, text="更新先Excel（任意）", style="Panel.TLabel").grid(row=3, column=0, sticky="w", pady=6)
        self.workbook_entry = ttk.Entry(panel, textvariable=self.workbook_path)
        self.workbook_entry.grid(row=3, column=1, sticky="ew", padx=(14, 8), pady=6)
        self.browse_workbook = ttk.Button(panel, text="参照...", command=self._browse_workbook)
        self.browse_workbook.grid(row=3, column=2, sticky="ew", pady=6)

        ttk.Label(panel, text="シート名（任意）", style="Panel.TLabel").grid(row=4, column=0, sticky="w", pady=6)
        ttk.Entry(panel, textvariable=self.sheet_name).grid(row=4, column=1, sticky="ew", padx=(14, 8), pady=6)
        ttk.Label(panel, text="空欄なら自動判定", style="Panel.TLabel").grid(row=4, column=2, sticky="w", pady=6)

        self.start_label = ttk.Label(panel, text="変換開始行（任意）", style="Panel.TLabel")
        self.start_label.grid(row=5, column=0, sticky="w", pady=6)
        ttk.Entry(panel, textvariable=self.start_row, width=14).grid(row=5, column=1, sticky="w", padx=(14, 8), pady=6)
        self.start_hint = ttk.Label(panel, text="Excelから読み込む最初の行", style="Panel.TLabel")
        self.start_hint.grid(row=5, column=2, sticky="w", pady=6)

        self.overwrite_check = ttk.Checkbutton(panel, text="同名Markdownを上書き", variable=self.overwrite)
        self.overwrite_check.grid(row=6, column=1, sticky="w", padx=(14, 8), pady=(10, 4))
        self.run_button = ttk.Button(panel, text="実行", style="Accent.TButton", command=self._run)
        self.run_button.grid(row=7, column=2, sticky="e", pady=(18, 0))

        footer = ttk.Frame(outer, style="App.TFrame")
        footer.pack(fill="x", pady=(14, 0))
        ttk.Label(footer, textvariable=self.status, style="Subtitle.TLabel").pack(side="left", fill="x", expand=True)
        self.progress = ttk.Progressbar(footer, mode="indeterminate", length=150)
        self.progress.pack(side="right")

    def _update_mode(self) -> None:
        export = self.mode.get() == "xlsx_to_md"
        self.input_label.configure(text="対象Excel" if export else "対象Markdown")
        self.start_label.configure(text="変換開始行（任意）" if export else "コピー開始行（任意）")
        self.start_hint.configure(text="Excelから読み込む最初の行" if export else "更新先Excelの行番号")
        self.workbook_entry.configure(state="disabled" if export else "normal")
        self.browse_workbook.configure(state="disabled" if export else "normal")
        self.overwrite_check.configure(state="normal" if export else "disabled")

    def _browse_input(self) -> None:
        export = self.mode.get() == "xlsx_to_md"
        filetypes = [("Excelブック", "*.xlsx")] if export else [("Markdown", "*.md")]
        path = filedialog.askopenfilename(title="対象ファイルを選択", filetypes=filetypes)
        if path:
            self.input_path.set(path)
            if not export:
                workbooks = list(Path(path).parent.glob("*.xlsx"))
                if len(workbooks) == 1:
                    self.workbook_path.set(str(workbooks[0]))

    def _browse_workbook(self) -> None:
        path = filedialog.askopenfilename(title="更新先Excelを選択", filetypes=[("Excelブック", "*.xlsx")])
        if path:
            self.workbook_path.set(path)

    def _run(self) -> None:
        source = Path(self.input_path.get().strip().strip('"')).expanduser()
        suffix = ".xlsx" if self.mode.get() == "xlsx_to_md" else ".md"
        if not source.is_file() or source.suffix.lower() != suffix:
            messagebox.showerror("入力エラー", f"有効な{suffix}ファイルを指定してください。")
            return
        raw_row = self.start_row.get().strip()
        if raw_row and (not raw_row.isdigit() or int(raw_row) < 1):
            messagebox.showerror("行番号エラー", "開始行は1以上の整数で指定してください。")
            return
        start_row = int(raw_row) if raw_row else None
        raw_workbook = self.workbook_path.get().strip().strip('"')
        workbook = Path(raw_workbook).expanduser() if raw_workbook else None
        sheet = self.sheet_name.get().strip() or None

        if self.mode.get() == "xlsx_to_md":
            try:
                output_path, _, _ = create_markdown(source, sheet, start_row)
            except Exception as error:
                messagebox.showerror("Excelを確認してください", str(error))
                return
            if output_path.exists() and not self.overwrite.get():
                messagebox.showerror("出力先にファイルがあります", f"上書きを有効にしてください。\n{output_path}")
                return

        operation = self.mode.get()
        self.run_button.configure(state="disabled")
        self.progress.start(12)
        self.status.set("処理中...")
        threading.Thread(target=self._worker, args=(operation, source, workbook, sheet, start_row), daemon=True).start()

    def _worker(self, operation, source, workbook, sheet, start_row) -> None:
        try:
            if operation == "xlsx_to_md":
                output, count, content = create_markdown(source, sheet, start_row)
                output.write_text(content, encoding="utf-8", newline="\n")
                message = f"Markdownを作成しました。\n{output}\n{count}行"
            else:
                output, count = process_markdown(
                    source, workbook, sheet, start_row,
                    progress=lambda text: self.events.put(("status", text)),
                )
                message = f"仕様書を更新しました。\n{output}\n{count}行"
            self.events.put(("success", message))
        except Exception as error:
            self.events.put(("error", str(error)))

    def _poll_events(self) -> None:
        try:
            while True:
                event, payload = self.events.get_nowait()
                if event == "status":
                    self.status.set(str(payload))
                else:
                    self.progress.stop()
                    self.run_button.configure(state="normal")
                    self.status.set("完了" if event == "success" else "失敗")
                    if event == "success":
                        messagebox.showinfo("完了", str(payload))
                    else:
                        messagebox.showerror("処理に失敗しました", str(payload))
        except queue.Empty:
            pass
        self.root.after(100, self._poll_events)


def main() -> None:
    root = tk.Tk()
    ConverterWindow(root)
    root.mainloop()
