"""Tkinter desktop interface for the CSV/Excel merger."""

from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Dict, Optional

from merge_core import MergeResult, merge_csv_files


APP_NAME = "CSV文件合并工具"
DISPLAY_NAME = "CSV / Excel 文件合并工具"


def get_settings_path() -> Path:
    """Use the per-user application data directory, including in a packaged EXE."""

    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming"))
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    return base / APP_NAME / "settings.json"


class SettingsStore:
    def __init__(self, path: Optional[Path] = None) -> None:
        self.path = path or get_settings_path()

    def load(self) -> Dict[str, Any]:
        try:
            with self.path.open("r", encoding="utf-8") as handle:
                value = json.load(handle)
            return value if isinstance(value, dict) else {}
        except (FileNotFoundError, OSError, ValueError, TypeError):
            return {}

    def save(self, settings: Dict[str, Any]) -> None:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        temporary_path = self.path.with_suffix(".tmp")
        with temporary_path.open("w", encoding="utf-8") as handle:
            json.dump(settings, handle, ensure_ascii=False, indent=2)
        temporary_path.replace(self.path)


class MergeApplication:
    POLL_INTERVAL_MS = 100

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(DISPLAY_NAME)
        self.root.geometry("820x590")
        self.root.minsize(720, 520)

        self.settings_store = SettingsStore()
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.last_output_folder = ""

        settings = self.settings_store.load()
        self.input_folder = tk.StringVar(value=str(settings.get("input_folder", "")))
        self.template_path = tk.StringVar(value=str(settings.get("template_path", "")))
        self.output_folder = tk.StringVar(value=str(settings.get("output_folder", "")))
        self.output_filename = tk.StringVar(
            value=str(settings.get("output_filename", "merged_output.csv"))
        )
        self.add_source_column = tk.BooleanVar(
            value=bool(settings.get("add_source_column", True))
        )
        self.status_text = tk.StringVar(value="请选择文件夹后开始合并")

        self._configure_style()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _configure_style(self) -> None:
        style = ttk.Style(self.root)
        if sys.platform == "win32" and "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("Subtitle.TLabel", foreground="#5f6368")
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 10, "bold"))

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=(24, 20, 24, 18))
        container.grid(row=0, column=0, sticky="nsew")
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        container.columnconfigure(0, weight=1)
        container.rowconfigure(4, weight=1)

        ttk.Label(container, text=DISPLAY_NAME, style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            container,
            text="按列名对齐批量合并 CSV、XLSX；模板第一行决定目标列名和顺序。",
            style="Subtitle.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(4, 18))

        form = ttk.LabelFrame(container, text="合并设置", padding=(16, 14))
        form.grid(row=2, column=0, sticky="ew")
        form.columnconfigure(1, weight=1)

        self._add_path_row(
            form,
            row=0,
            label="输入文件夹",
            variable=self.input_folder,
            command=self._choose_input_folder,
        )
        self._add_path_row(
            form,
            row=1,
            label="列名模板",
            variable=self.template_path,
            command=self._choose_template,
            optional=True,
        )
        self._add_path_row(
            form,
            row=2,
            label="输出文件夹",
            variable=self.output_folder,
            command=self._choose_output_folder,
        )

        ttk.Label(form, text="输出文件名").grid(
            row=3, column=0, sticky="w", padx=(0, 12), pady=7
        )
        ttk.Entry(form, textvariable=self.output_filename).grid(
            row=3, column=1, sticky="ew", pady=7
        )
        ttk.Label(form, text="可省略 .csv", foreground="#5f6368").grid(
            row=3, column=2, sticky="w", padx=(8, 0), pady=7
        )

        ttk.Checkbutton(
            form,
            text="添加 Source_File 来源文件名列",
            variable=self.add_source_column,
        ).grid(row=4, column=1, columnspan=2, sticky="w", pady=(8, 2))

        actions = ttk.Frame(container)
        actions.grid(row=3, column=0, sticky="ew", pady=(16, 12))
        actions.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(actions, mode="indeterminate", length=180)
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 12))
        self.open_button = ttk.Button(
            actions,
            text="打开输出文件夹",
            command=self._open_output_folder,
            state="disabled",
        )
        self.open_button.grid(row=0, column=1, padx=(0, 8))
        self.merge_button = ttk.Button(
            actions,
            text="开始合并",
            command=self._start_merge,
            style="Primary.TButton",
            width=14,
        )
        self.merge_button.grid(row=0, column=2)

        log_frame = ttk.LabelFrame(container, text="运行信息", padding=(10, 8))
        log_frame.grid(row=4, column=0, sticky="nsew")
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log = tk.Text(
            log_frame,
            height=9,
            wrap="word",
            state="disabled",
            borderwidth=0,
            background="#fafafa",
            font=("Microsoft YaHei UI", 9),
        )
        scrollbar = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=scrollbar.set)
        self.log.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        ttk.Label(container, textvariable=self.status_text, style="Subtitle.TLabel").grid(
            row=5, column=0, sticky="w", pady=(10, 0)
        )

    def _add_path_row(
        self,
        parent: ttk.LabelFrame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command: object,
        optional: bool = False,
    ) -> None:
        label_text = f"{label}（可选）" if optional else label
        ttk.Label(parent, text=label_text).grid(
            row=row, column=0, sticky="w", padx=(0, 12), pady=7
        )
        ttk.Entry(parent, textvariable=variable).grid(
            row=row, column=1, sticky="ew", pady=7
        )
        ttk.Button(parent, text="浏览…", command=command, width=10).grid(
            row=row, column=2, padx=(8, 0), pady=7
        )

    def _initial_directory(self, value: str) -> str:
        if value:
            path = Path(value)
            candidate = path if path.is_dir() else path.parent
            if candidate.is_dir():
                return str(candidate)
        return str(Path.home())

    def _choose_input_folder(self) -> None:
        selected = filedialog.askdirectory(
            title="选择包含 CSV 或 Excel 的输入文件夹",
            initialdir=self._initial_directory(self.input_folder.get()),
        )
        if selected:
            self.input_folder.set(selected)
            if not self.output_folder.get().strip():
                self.output_folder.set(selected)
            self._save_settings(show_error=False)

    def _choose_template(self) -> None:
        selected = filedialog.askopenfilename(
            title="选择列名模板",
            initialdir=self._initial_directory(self.template_path.get()),
            filetypes=[
                ("支持的模板", "*.csv *.xlsx *.xlsm"),
                ("CSV 文件", "*.csv"),
                ("Excel 文件", "*.xlsx *.xlsm"),
                ("所有文件", "*.*"),
            ],
        )
        if selected:
            self.template_path.set(selected)
            self._save_settings(show_error=False)

    def _choose_output_folder(self) -> None:
        selected = filedialog.askdirectory(
            title="选择输出文件夹",
            initialdir=self._initial_directory(self.output_folder.get()),
        )
        if selected:
            self.output_folder.set(selected)
            self._save_settings(show_error=False)

    def _settings(self) -> Dict[str, Any]:
        return {
            "input_folder": self.input_folder.get().strip(),
            "template_path": self.template_path.get().strip(),
            "output_folder": self.output_folder.get().strip(),
            "output_filename": self.output_filename.get().strip(),
            "add_source_column": self.add_source_column.get(),
        }

    def _save_settings(self, show_error: bool = True) -> None:
        try:
            self.settings_store.save(self._settings())
        except OSError as exc:
            if show_error:
                messagebox.showwarning("设置未保存", f"无法保存上次选择路径：\n{exc}")

    def _validate(self) -> Optional[Dict[str, Any]]:
        settings = self._settings()
        input_folder_text = settings["input_folder"]
        if not input_folder_text:
            messagebox.showerror("输入有误", "请选择输入文件夹。")
            return None

        input_folder = Path(input_folder_text)
        output_folder_text = settings["output_folder"] or settings["input_folder"]
        template_text = settings["template_path"]

        if not input_folder.is_dir():
            messagebox.showerror("输入有误", "请选择有效的输入文件夹。")
            return None
        if template_text and not Path(template_text).is_file():
            messagebox.showerror("输入有误", "请选择有效的列名模板，或将模板路径留空。")
            return None
        if not settings["output_filename"]:
            messagebox.showerror("输入有误", "请输入输出文件名。")
            return None

        settings["output_folder"] = output_folder_text
        return settings

    def _start_merge(self) -> None:
        settings = self._validate()
        if settings is None:
            return

        self.output_folder.set(settings["output_folder"])
        self._save_settings()
        self._clear_log()
        self._append_log("开始执行合并任务……")
        self.status_text.set("正在合并，请稍候……")
        self.merge_button.configure(state="disabled")
        self.open_button.configure(state="disabled")
        self.progress.start(12)

        worker = threading.Thread(
            target=self._merge_worker,
            args=(settings,),
            daemon=True,
            name="csv-merge-worker",
        )
        worker.start()
        self.root.after(self.POLL_INTERVAL_MS, self._poll_events)

    def _merge_worker(self, settings: Dict[str, Any]) -> None:
        try:
            result = merge_csv_files(
                input_folder=settings["input_folder"],
                output_folder=settings["output_folder"],
                template_path=settings["template_path"] or None,
                output_filename=settings["output_filename"],
                add_source_column=settings["add_source_column"],
                progress_callback=lambda message: self.events.put(("progress", message)),
            )
            self.events.put(("success", result))
        except Exception as exc:
            self.events.put(("error", exc))

    def _poll_events(self) -> None:
        finished = False
        while True:
            try:
                event, payload = self.events.get_nowait()
            except queue.Empty:
                break

            if event == "progress":
                message = str(payload)
                self._append_log(message)
                self.status_text.set(message)
            elif event == "success":
                finished = True
                self._merge_succeeded(payload)
            elif event == "error":
                finished = True
                self._merge_failed(payload)

        if not finished:
            self.root.after(self.POLL_INTERVAL_MS, self._poll_events)

    def _merge_succeeded(self, result_object: object) -> None:
        result = result_object
        if not isinstance(result, MergeResult):
            self._merge_failed(RuntimeError("程序返回了无效的合并结果。"))
            return

        self._finish_busy_state()
        self.last_output_folder = str(result.output_path.parent)
        self.open_button.configure(state="normal")
        summary = f"合并完成：{result.file_count} 个文件，共 {result.row_count} 行。"
        self.status_text.set(summary)
        self._append_log(f"输出文件：{result.output_path}")
        if result.warning_path:
            self._append_log(f"检查报告：{result.warning_path}")
        messagebox.showinfo("合并完成", f"{summary}\n\n输出文件：\n{result.output_path}")

    def _merge_failed(self, error: object) -> None:
        self._finish_busy_state()
        message = str(error)
        self.status_text.set("合并失败，请检查运行信息。")
        self._append_log(f"错误：{message}")
        messagebox.showerror("合并失败", message)

    def _finish_busy_state(self) -> None:
        self.progress.stop()
        self.merge_button.configure(state="normal")

    def _append_log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert("end", message + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")

    def _open_output_folder(self) -> None:
        folder = self.last_output_folder or self.output_folder.get().strip()
        if not folder or not Path(folder).is_dir():
            messagebox.showerror("无法打开", "输出文件夹不存在。")
            return
        try:
            if sys.platform == "win32":
                os.startfile(folder)  # type: ignore[attr-defined]
            elif sys.platform == "darwin":
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])
        except OSError as exc:
            messagebox.showerror("无法打开", str(exc))

    def _on_close(self) -> None:
        self._save_settings(show_error=False)
        self.root.destroy()


def enable_windows_dpi_awareness() -> None:
    if sys.platform != "win32":
        return
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass


def main() -> None:
    enable_windows_dpi_awareness()
    root = tk.Tk()
    MergeApplication(root)
    root.mainloop()


if __name__ == "__main__":
    main()
