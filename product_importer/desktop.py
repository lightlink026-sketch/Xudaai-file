from __future__ import annotations

import json
import mimetypes
import os
import threading
import tkinter as tk
import urllib.error
import urllib.request
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any

import httpx

from core import make_product_package, scan_product_folder

APP_TITLE = "Xudaai 产品资料整理工具"
CONFIG_PATH = Path.home() / ".xudaai-product-importer.json"
FIELDS = [
    ("name", "产品名称", "基本信息"), ("category", "产品分类", "基本信息"),
    ("brand", "品牌", "基本信息"), ("model", "产品型号", "基本信息"),
    ("sku", "自有 SKU", "基本信息"), ("supplier_name", "供应商", "基本信息"),
    ("supplier_sku", "供应商型号/SKU", "基本信息"), ("unit", "销售单位", "基本信息"),
    ("description", "产品描述", "基本信息"), ("material", "材质", "基本信息"),
    ("colors", "颜色/变体", "规格与价格"), ("sizes", "尺码/规格", "规格与价格"),
    ("currency", "价格币种", "规格与价格"), ("unit_price", "采购单价", "规格与价格"),
    ("moq", "MOQ", "规格与价格"), ("carton_spec", "箱规（长×宽×高及单位）", "包装运输"),
    ("units_per_carton", "单箱装量", "包装运输"), ("carton_quantity", "箱数", "包装运输"),
    ("carton_volume_cbm", "单箱体积 CBM", "包装运输"),
    ("gross_weight_kg_per_carton", "单箱毛重 kg", "包装运输"),
    ("net_weight_kg_per_carton", "单箱净重 kg", "包装运输"),
    ("lead_time", "生产/交期", "包装运输"), ("notes", "其他备注", "包装运输"),
]
AI_SCHEMA = {key: "string; unknown values must be empty" for key, _, _ in FIELDS}


def read_config() -> dict[str, Any]:
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


class ProductImporter(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1120x860")
        self.minsize(950, 700)
        self.info: dict[str, Any] | None = None
        self.package_dir: Path | None = None
        self.product: dict[str, Any] | None = None
        self.image_names: list[str] = []
        self.vars: dict[str, tk.StringVar] = {}
        self.config_data = read_config()
        self._build()

    def _build(self) -> None:
        outer = ttk.Frame(self, padding=14)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="本地产品资料整理", font=("Microsoft YaHei UI", 18, "bold")).pack(anchor="w")
        ttk.Label(outer, text="选择一个已整理好的产品文件夹（一个文件夹对应一款产品），本地解析、人工确认后再上传到现有资料库。", wraplength=1050).pack(anchor="w", pady=(4, 12))

        paths = ttk.LabelFrame(outer, text="1. 选择资料", padding=10)
        paths.pack(fill="x", pady=4)
        self.source_var = tk.StringVar()
        self.excel_var = tk.StringVar()
        self.output_var = tk.StringVar(value=self.config_data.get("output_dir", str(Path.home() / "Documents" / "Xudaai产品档案")))
        self.server_var = tk.StringVar(value=self.config_data.get("server_url", "http://127.0.0.1:8002"))
        self.user_var = tk.StringVar(value=self.config_data.get("server_user", "admin"))
        self.password_var = tk.StringVar()
        self.source_var.trace_add("write", self._invalidate_package)
        self._path_row(paths, 0, "产品文件夹", self.source_var, self._choose_source, "选择文件夹")
        self._path_row(paths, 1, "补充 Excel（可选）", self.excel_var, self._choose_excel, "选择 Excel")
        self._path_row(paths, 2, "标准档案输出位置", self.output_var, self._choose_output, "选择位置")
        self.scan_button = ttk.Button(paths, text="扫描并读取资料", command=self._scan)
        self.scan_button.grid(row=3, column=2, sticky="e", pady=(8, 0))
        paths.columnconfigure(1, weight=1)

        notebook = ttk.Notebook(outer)
        notebook.pack(fill="both", expand=True, pady=8)
        for tab_name in ("基本信息", "规格与价格", "包装运输"):
            pane = ttk.Frame(notebook, padding=12)
            notebook.add(pane, text=tab_name)
            tab_fields = [(key, label) for key, label, section in FIELDS if section == tab_name]
            for index, (key, label) in enumerate(tab_fields):
                row, col = divmod(index, 2)
                cell = ttk.Frame(pane)
                cell.grid(row=row, column=col, sticky="ew", padx=(0 if col == 0 else 14, 0), pady=5)
                ttk.Label(cell, text=label).pack(anchor="w")
                var = tk.StringVar()
                var.trace_add("write", self._invalidate_package)
                self.vars[key] = var
                if key in {"description", "notes"}:
                    text = tk.Text(cell, height=3, wrap="word")
                    text.pack(fill="x")
                    text.bind("<<Modified>>", self._text_modified)
                    self.vars[key] = var
                    setattr(self, f"text_{key}", text)
                else:
                    ttk.Entry(cell, textvariable=var).pack(fill="x", ipady=3)
            pane.columnconfigure(0, weight=1)
            pane.columnconfigure(1, weight=1)

        media = ttk.LabelFrame(outer, text="主图选择", padding=9)
        media.pack(fill="x")
        self.main_image_var = tk.StringVar()
        self.main_image_var.trace_add("write", self._invalidate_package)
        self.image_select = ttk.Combobox(media, textvariable=self.main_image_var, state="readonly")
        self.image_select.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.scan_summary = tk.StringVar(value="尚未扫描")
        ttk.Label(media, textvariable=self.scan_summary).pack(side="left")

        ai_box = ttk.LabelFrame(outer, text="可选：AI 提取（会把解析出的文字发送到指定模型服务）", padding=9)
        ai_box.pack(fill="x", pady=(8, 0))
        self.ai_url_var = tk.StringVar(value=self.config_data.get("ai_url", "https://api.siliconflow.cn/v1"))
        self.ai_model_var = tk.StringVar(value=self.config_data.get("ai_model", ""))
        self.ai_key_var = tk.StringVar()
        self._compact_entry(ai_box, "API 地址", self.ai_url_var, 0, 0)
        self._compact_entry(ai_box, "模型名称", self.ai_model_var, 0, 2)
        self._compact_entry(ai_box, "API Key（仅本次运行保存在内存）", self.ai_key_var, 1, 0, secret=True)
        ttk.Button(ai_box, text="AI 提取并填入草稿", command=self._ai_extract).grid(row=1, column=3, sticky="e", padx=5, pady=4)
        for col in (1, 3): ai_box.columnconfigure(col, weight=1)

        upload = ttk.LabelFrame(outer, text="2. 上传到现有产品资料库", padding=9)
        upload.pack(fill="x", pady=(8, 0))
        self._compact_entry(upload, "软件地址", self.server_var, 0, 0)
        self._compact_entry(upload, "账号", self.user_var, 0, 2)
        self._compact_entry(upload, "密码（本次运行）", self.password_var, 1, 0, secret=True)
        buttons = ttk.Frame(upload)
        buttons.grid(row=1, column=3, sticky="e", padx=5)
        ttk.Button(buttons, text="生成标准产品档案", command=self._generate_package).pack(side="left", padx=4)
        ttk.Button(buttons, text="上传并建立索引", command=self._upload).pack(side="left", padx=4)
        upload.columnconfigure(1, weight=1)
        upload.columnconfigure(3, weight=1)
        self.status_var = tk.StringVar(value="准备就绪。建议先扫描、复核字段和主图，再生成档案并上传。")
        ttk.Label(outer, textvariable=self.status_var, wraplength=1050).pack(anchor="w", pady=(9, 0))

    def _path_row(self, parent: ttk.LabelFrame, row: int, label: str, variable: tk.StringVar, callback, button: str) -> None:
        ttk.Label(parent, text=label, width=22).grid(row=row, column=0, sticky="w", pady=3)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=7, pady=3)
        ttk.Button(parent, text=button, command=callback).grid(row=row, column=2, sticky="e", pady=3)

    def _compact_entry(self, parent: ttk.LabelFrame, label: str, variable: tk.StringVar, row: int, col: int, secret: bool = False) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="w", padx=5, pady=3)
        entry = ttk.Entry(parent, textvariable=variable, show="•" if secret else "")
        entry.grid(row=row, column=col + 1, sticky="ew", padx=5, pady=3)

    def _choose_source(self) -> None:
        path = filedialog.askdirectory(title="选择一个产品资料文件夹")
        if path: self.source_var.set(path)

    def _choose_excel(self) -> None:
        path = filedialog.askopenfilename(title="选择补充 Excel", filetypes=[("Excel", "*.xlsx *.xlsm"), ("所有文件", "*.*")])
        if path: self.excel_var.set(path)

    def _choose_output(self) -> None:
        path = filedialog.askdirectory(title="选择标准档案输出位置")
        if path: self.output_var.set(path)

    def _scan(self) -> None:
        source_value = self.source_var.get().strip()
        if not source_value:
            messagebox.showinfo(APP_TITLE, "请先选择产品资料文件夹。")
            return
        self.info = None
        self.package_dir = None
        self.product = None
        self.ai_sources = {}
        extra = Path(self.excel_var.get()) if self.excel_var.get().strip() else None
        self.scan_button.configure(state="disabled")
        self.status_var.set("正在本地扫描并读取文件…")
        threading.Thread(target=self._run_scan, args=(Path(source_value), extra), daemon=True).start()

    def _run_scan(self, source: Path, extra: Path | None) -> None:
        try:
            info = scan_product_folder(source, extra)
            self.after(0, lambda: self._scan_finished(info, ""))
        except Exception as exc:
            error_text = str(exc)
            self.after(0, lambda message=error_text: self._scan_finished(None, message))

    def _scan_finished(self, info: dict[str, Any] | None, error: str) -> None:
        self.scan_button.configure(state="normal")
        if error:
            messagebox.showerror(APP_TITLE, error)
            return
        self.info = info
        assert self.info is not None
        self.image_names = [path.relative_to(Path(self.info["folder"])).as_posix() for path in self.info["images"]]
        self.image_select["values"] = self.image_names
        self.main_image_var.set(self.image_names[0] if self.image_names else "")
        self.scan_summary.set(f"{len(self.info['files'])} 个文件 · {len(self.image_names)} 张图 · {len(self.info['videos'])} 个视频")
        self.vars["name"].set(self.info["folder_name"])
        self.status_var.set(f"已读取 {len(self.info['source_documents'])} 份文字资料；可直接填写，或运行 AI 提取。")

    def _all_values(self) -> dict[str, Any]:
        result = {key: var.get().strip() for key, var in self.vars.items()}
        for key in ("description", "notes"):
            widget = getattr(self, f"text_{key}", None)
            if widget: result[key] = widget.get("1.0", "end").strip()
        return result

    def _invalidate_package(self, *_args) -> None:
        self.package_dir = None
        self.product = None

    def _text_modified(self, event) -> None:
        widget = event.widget
        if widget.edit_modified():
            self._invalidate_package()
            widget.edit_modified(False)

    def _ai_extract(self) -> None:
        if not self.info: self._scan()
        if not self.info: return
        if not self.info["source_text"].strip():
            messagebox.showinfo(APP_TITLE, "所选资料没有可提取文字。请手工填写字段；AI 暂无文字可供整理。")
            return
        api_url, model, key = self.ai_url_var.get().strip(), self.ai_model_var.get().strip(), self.ai_key_var.get().strip()
        if not api_url or not model or not key:
            messagebox.showinfo(APP_TITLE, "请填写 API 地址、模型名称和 API Key；不填 Key 也可手动整理并生成档案。")
            return
        self.status_var.set("AI 正在根据文字资料提取字段…")
        thread = threading.Thread(target=self._run_ai_extract, args=(api_url, model, key, self.info["source_text"], self.info["folder_name"]), daemon=True)
        thread.start()

    def _run_ai_extract(self, api_url: str, model: str, key: str, source_text: str, folder_name: str) -> None:
        prompt = (
            "你是供应商产品资料整理助手。只根据给定文件内容提取字段，不得猜测。未知填空字符串。"
            "返回一个纯 JSON 对象，字段必须为：" + json.dumps(AI_SCHEMA, ensure_ascii=False) +
            "。另加 field_sources 对象，记录每个有值字段对应的来源文件名/表格位置；不确定字段标记在 notes。"
            "不要把F货/F类商品误解为F价。货币、重量、箱规单位必须按原文。\n"
            f"产品目录名：{folder_name}\n来源资料文字：\n{source_text[:65000]}"
        )
        body = json.dumps({"model": model, "temperature": 0, "messages": [{"role": "user", "content": prompt}]}, ensure_ascii=False).encode("utf-8")
        request = urllib.request.Request(api_url.rstrip("/") + "/chat/completions", data=body, headers={"Content-Type": "application/json", "Authorization": f"Bearer {key}"}, method="POST")
        try:
            with urllib.request.urlopen(request, timeout=120) as response:
                payload = json.loads(response.read().decode("utf-8"))
            content = payload["choices"][0]["message"]["content"]
            content = content.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
            result = json.loads(content)
            self.after(0, lambda: self._apply_ai_result(result))
        except Exception as exc:
            error_text = str(exc)
            self.after(0, lambda message=error_text: self.status_var.set(f"AI 提取失败：{message}"))

    def _apply_ai_result(self, result: dict[str, Any]) -> None:
        if not isinstance(result, dict):
            self.status_var.set("AI 返回格式不是 JSON 对象，请手工填写。")
            return
        for key, value in result.items():
            if key in self.vars and key not in {"description", "notes"}:
                self.vars[key].set(str(value or ""))
            elif key in {"description", "notes"}:
                widget = getattr(self, f"text_{key}")
                widget.delete("1.0", "end")
                widget.insert("1.0", str(value or ""))
        if isinstance(result.get("field_sources"), dict):
            self.ai_sources = result["field_sources"]
        self.status_var.set("AI 草稿已填入。请逐项核对供应商、价格、单位、箱规和重量后再生成正式档案。")

    def _generate_package(self) -> None:
        if not self.info: self._scan()
        if not self.info: return
        try:
            output_text = self.output_var.get().strip()
            if not output_text: raise ValueError("请选择档案输出位置。")
            output = Path(output_text)
            values = self._all_values()
            if hasattr(self, "ai_sources"): values["field_sources"] = self.ai_sources
            self.package_dir, self.product = make_product_package(self.info, Path(self.excel_var.get()) if self.excel_var.get().strip() else None, output, values, self.main_image_var.get())
            self._save_config()
            self.status_var.set(f"标准档案已生成：{self.package_dir}。检查输出内容后，点击“上传并建立索引”。")
            messagebox.showinfo(APP_TITLE, f"档案已生成：\n{self.package_dir}")
        except Exception as exc:
            messagebox.showerror(APP_TITLE, str(exc))

    def _save_config(self) -> None:
        config = {"output_dir": self.output_var.get().strip(), "server_url": self.server_var.get().strip(), "server_user": self.user_var.get().strip(), "ai_url": self.ai_url_var.get().strip(), "ai_model": self.ai_model_var.get().strip()}
        CONFIG_PATH.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")

    def _upload(self) -> None:
        if not self.package_dir or not self.product:
            messagebox.showinfo(APP_TITLE, "请先生成并检查标准档案。")
            return
        if not messagebox.askyesno(APP_TITLE, "确认将该产品档案上传到现有软件资料库并建立索引？"):
            return
        self.status_var.set("正在上传文件并建立索引，请稍候…")
        threading.Thread(target=self._run_upload, daemon=True).start()

    def _run_upload(self) -> None:
        files = []
        streams = []
        try:
            for path in sorted(item for item in self.package_dir.rglob("*") if item.is_file()):
                relative = path.relative_to(self.package_dir).as_posix()
                stream = path.open("rb")
                streams.append(stream)
                files.append(("files", (relative, stream, mimetypes.guess_type(path.name)[0] or "application/octet-stream")))
            auth = (self.user_var.get().strip(), self.password_var.get()) if self.password_var.get() else None
            url = self.server_var.get().strip().rstrip("/") + "/api/products/import"
            response = httpx.post(url, data={"manifest": json.dumps(self.product, ensure_ascii=False)}, files=files, auth=auth, timeout=900)
            payload = response.json()
            if response.status_code >= 400:
                raise RuntimeError(payload.get("detail", f"服务器返回 HTTP {response.status_code}"))
            self.after(0, lambda: self._upload_finished(payload))
        except Exception as exc:
            error_text = str(exc)
            self.after(0, lambda message=error_text: self.status_var.set(f"上传失败：{message}"))
            self.after(0, lambda message=error_text: messagebox.showerror(APP_TITLE, f"上传失败：{message}"))
        finally:
            for stream in streams:
                stream.close()

    def _upload_finished(self, result: dict[str, Any]) -> None:
        product = result.get("product", {})
        self.status_var.set(f"上传完成：{product.get('name', self.product.get('name'))} · 已索引 {result.get('indexed', 0)} 个文件 · {result.get('folder', '')}")
        messagebox.showinfo(APP_TITLE, self.status_var.get())


if __name__ == "__main__":
    ProductImporter().mainloop()
