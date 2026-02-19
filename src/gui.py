from __future__ import annotations

import sys
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .config import XmlConfig, load_gui_defaults, save_gui_defaults
from .models import MeasurementRecord, WorkbookParseResult
from .parser import parse_folder, parse_workbook
from .xml_exporter import write_addon_xml


class ExcelParserApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Chromsystems Universal Leaflet Parser 0.1.0")
        self.root.geometry("1600x920")
        self._window_icon = None
        self._apply_window_icon()

        self.results: list[WorkbookParseResult] = []

        self._filter_source = tk.StringVar(value="")
        self._filter_sample = tk.StringVar(value="")
        self._filter_analyte = tk.StringVar(value="")
        self._filter_unit = tk.StringVar(value="")
        self._filter_metric = tk.StringVar(value="")

        self._cfg_method_id = tk.StringVar()
        self._cfg_method_version = tk.StringVar()
        self._cfg_sample_tube_types = tk.StringVar()
        self._cfg_measurement_lists = tk.StringVar()
        self._cfg_run_results_path = tk.StringVar()

        self._build_ui()
        self._load_default_config()
        self._refresh_preview()

    def _build_ui(self) -> None:
        self.root.rowconfigure(3, weight=1)
        self.root.rowconfigure(4, weight=1)
        self.root.columnconfigure(0, weight=1)

        control_frame = ttk.Frame(self.root, padding=10)
        control_frame.grid(row=0, column=0, sticky="ew")
        for idx in range(6):
            control_frame.columnconfigure(idx, weight=0)
        control_frame.columnconfigure(6, weight=1)

        ttk.Button(control_frame, text="Import File", command=self.import_file).grid(
            row=0, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Button(control_frame, text="Import Folder", command=self.import_folder).grid(
            row=0, column=1, padx=5, pady=5, sticky="w"
        )
        ttk.Button(
            control_frame, text="Preview Normalized Records", command=self._refresh_preview
        ).grid(row=0, column=2, padx=5, pady=5, sticky="w")
        ttk.Button(control_frame, text="Export XML", command=self.export_xml).grid(
            row=0, column=3, padx=5, pady=5, sticky="w"
        )
        ttk.Button(control_frame, text="Save Defaults", command=self.save_defaults).grid(
            row=0, column=4, padx=5, pady=5, sticky="w"
        )
        ttk.Button(control_frame, text="Clear", command=self.clear_results).grid(
            row=0, column=5, padx=5, pady=5, sticky="w"
        )

        config_frame = ttk.LabelFrame(self.root, text="XML Config", padding=10)
        config_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))
        for idx in range(4):
            config_frame.columnconfigure(idx, weight=1)

        ttk.Label(config_frame, text="MethodId").grid(row=0, column=0, sticky="w", padx=4, pady=4)
        ttk.Entry(config_frame, textvariable=self._cfg_method_id).grid(
            row=0, column=1, columnspan=3, sticky="ew", padx=4, pady=4
        )
        ttk.Label(config_frame, text="MethodVersion").grid(
            row=1, column=0, sticky="w", padx=4, pady=4
        )
        ttk.Entry(config_frame, textvariable=self._cfg_method_version).grid(
            row=1, column=1, sticky="ew", padx=4, pady=4
        )
        ttk.Label(config_frame, text="RunResultsExportPath").grid(
            row=1, column=2, sticky="w", padx=4, pady=4
        )
        ttk.Entry(config_frame, textvariable=self._cfg_run_results_path).grid(
            row=1, column=3, sticky="ew", padx=4, pady=4
        )
        ttk.Label(config_frame, text="SampleTubeTypes (comma-separated)").grid(
            row=2, column=0, sticky="w", padx=4, pady=4
        )
        ttk.Entry(config_frame, textvariable=self._cfg_sample_tube_types).grid(
            row=2, column=1, sticky="ew", padx=4, pady=4
        )
        ttk.Label(config_frame, text="MeasurementSampleLists (comma-separated)").grid(
            row=2, column=2, sticky="w", padx=4, pady=4
        )
        ttk.Entry(config_frame, textvariable=self._cfg_measurement_lists).grid(
            row=2, column=3, sticky="ew", padx=4, pady=4
        )

        filter_frame = ttk.LabelFrame(self.root, text="Preview Filters", padding=10)
        filter_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=(0, 10))
        for idx in range(10):
            filter_frame.columnconfigure(idx, weight=1)

        self._source_combo = self._make_filter_combo(
            parent=filter_frame, label="File", variable=self._filter_source, row=0, col=0
        )
        self._sample_combo = self._make_filter_combo(
            parent=filter_frame, label="Sample", variable=self._filter_sample, row=0, col=2
        )
        self._analyte_combo = self._make_filter_combo(
            parent=filter_frame, label="Analyte", variable=self._filter_analyte, row=0, col=4
        )
        self._unit_combo = self._make_filter_combo(
            parent=filter_frame, label="Unit", variable=self._filter_unit, row=0, col=6
        )
        self._metric_combo = self._make_filter_combo(
            parent=filter_frame, label="Metric", variable=self._filter_metric, row=0, col=8
        )

        preview_frame = ttk.LabelFrame(self.root, text="Normalized Records", padding=10)
        preview_frame.grid(row=3, column=0, sticky="nsew", padx=10, pady=(0, 10))
        preview_frame.rowconfigure(0, weight=1)
        preview_frame.columnconfigure(0, weight=1)

        columns = (
            "source_file",
            "sample_label",
            "sample_code",
            "unit",
            "analyte_name",
            "group_name",
            "metric_role",
            "raw_value",
            "numeric_value",
            "value_status",
            "sheet_row",
            "sheet_col",
        )
        self._tree = ttk.Treeview(preview_frame, columns=columns, show="headings", height=18)
        for col in columns:
            self._tree.heading(col, text=col)
            width = 110
            if col in {"source_file", "analyte_name"}:
                width = 220
            if col in {"raw_value", "group_name"}:
                width = 160
            self._tree.column(col, width=width, anchor="w")
        self._tree.grid(row=0, column=0, sticky="nsew")

        y_scroll = ttk.Scrollbar(preview_frame, orient="vertical", command=self._tree.yview)
        y_scroll.grid(row=0, column=1, sticky="ns")
        self._tree.configure(yscrollcommand=y_scroll.set)

        log_frame = ttk.LabelFrame(self.root, text="Status Log", padding=10)
        log_frame.grid(row=4, column=0, sticky="nsew", padx=10, pady=(0, 10))
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)

        self._log_text = tk.Text(log_frame, height=8, wrap="word")
        self._log_text.grid(row=0, column=0, sticky="nsew")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self._log_text.yview)
        log_scroll.grid(row=0, column=1, sticky="ns")
        self._log_text.configure(yscrollcommand=log_scroll.set, state="disabled")

    def _make_filter_combo(
        self,
        parent: ttk.LabelFrame,
        label: str,
        variable: tk.StringVar,
        row: int,
        col: int,
    ) -> ttk.Combobox:
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="w", padx=4, pady=4)
        combo = ttk.Combobox(parent, textvariable=variable, state="readonly")
        combo.grid(row=row, column=col + 1, sticky="ew", padx=4, pady=4)
        combo["values"] = ("",)
        combo.bind("<<ComboboxSelected>>", lambda _event: self._refresh_preview())
        return combo

    def _load_default_config(self) -> None:
        cfg = load_gui_defaults()
        self._cfg_method_id.set(cfg.method_id)
        self._cfg_method_version.set(cfg.method_version)
        self._cfg_sample_tube_types.set(", ".join(cfg.sample_tube_types))
        self._cfg_measurement_lists.set(", ".join(cfg.measurement_sample_lists))
        self._cfg_run_results_path.set(cfg.run_results_export_path)

    def _collect_config(self) -> XmlConfig:
        return XmlConfig(
            method_id=self._cfg_method_id.get().strip() or "Beta Therapeutic Drug Monitoring",
            method_version=self._cfg_method_version.get().strip() or "1.2",
            sample_tube_types=_split_csv(self._cfg_sample_tube_types.get()),
            measurement_sample_lists=_split_csv(self._cfg_measurement_lists.get()),
            run_results_export_path=self._cfg_run_results_path.get().strip(),
        )

    def save_defaults(self) -> None:
        cfg = self._collect_config()
        save_gui_defaults(cfg)
        self._log("Saved GUI defaults to config/gui_defaults.json")

    def import_file(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select Excel workbook",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not selected:
            return
        result = parse_workbook(Path(selected))
        self._upsert_results([result])
        self._log(f"Imported file: {result.source_file}")
        self._log_warnings(result)
        self._refresh_preview()

    def import_folder(self) -> None:
        selected = filedialog.askdirectory(title="Select folder with Excel workbooks")
        if not selected:
            return
        results = parse_folder(Path(selected))
        if not results:
            messagebox.showinfo("No files", "No .xlsx files found in selected folder.")
            return
        self._upsert_results(results)
        self._log(f"Imported folder: {selected} ({len(results)} workbook(s))")
        for result in results:
            self._log_warnings(result)
        self._refresh_preview()

    def export_xml(self) -> None:
        if not self.results:
            messagebox.showwarning("No data", "Import at least one workbook first.")
            return
        out_dir = filedialog.askdirectory(title="Select output folder for XML export")
        if not out_dir:
            return
        cfg = self._collect_config()
        save_gui_defaults(cfg)
        exported = []
        for result in self.results:
            out_path = write_addon_xml(result=result, cfg=cfg, out_dir=Path(out_dir))
            exported.append(out_path)
        self._log(f"Exported {len(exported)} XML file(s) to: {out_dir}")
        messagebox.showinfo("XML export complete", f"Exported {len(exported)} XML file(s).")

    def clear_results(self) -> None:
        self.results = []
        self._filter_source.set("")
        self._filter_sample.set("")
        self._filter_analyte.set("")
        self._filter_unit.set("")
        self._filter_metric.set("")
        self._refresh_preview()
        self._log("Cleared loaded results.")

    def _upsert_results(self, new_results: list[WorkbookParseResult]) -> None:
        by_name = {result.source_file: result for result in self.results}
        for result in new_results:
            by_name[result.source_file] = result
        self.results = [by_name[key] for key in sorted(by_name.keys())]

    def _refresh_preview(self) -> None:
        self._refresh_filter_values()
        records = self._filtered_measurements()

        self._tree.delete(*self._tree.get_children())
        for rec in records:
            self._tree.insert(
                "",
                "end",
                values=(
                    rec.source_file,
                    rec.sample_label or "",
                    rec.sample_code or "",
                    rec.unit or "",
                    rec.analyte_name,
                    rec.group_name or "",
                    rec.metric_role,
                    rec.raw_value or "",
                    "" if rec.numeric_value is None else rec.numeric_value,
                    rec.value_status,
                    rec.sheet_row,
                    rec.sheet_col,
                ),
            )
        self._log(f"Preview rows: {len(records)}")

    def _filtered_measurements(self) -> list[MeasurementRecord]:
        source_filter = self._filter_source.get().strip()
        sample_filter = self._filter_sample.get().strip()
        analyte_filter = self._filter_analyte.get().strip()
        unit_filter = self._filter_unit.get().strip()
        metric_filter = self._filter_metric.get().strip()

        out: list[MeasurementRecord] = []
        for result in self.results:
            for rec in result.normalized_values:
                if source_filter and rec.source_file != source_filter:
                    continue
                if sample_filter and (rec.sample_label or "") != sample_filter:
                    continue
                if analyte_filter and rec.analyte_name != analyte_filter:
                    continue
                if unit_filter and (rec.unit or "") != unit_filter:
                    continue
                if metric_filter and rec.metric_role != metric_filter:
                    continue
                out.append(rec)
        return out

    def _refresh_filter_values(self) -> None:
        source_values = sorted({result.source_file for result in self.results})
        sample_values = sorted(
            {
                rec.sample_label
                for result in self.results
                for rec in result.normalized_values
                if rec.sample_label
            }
        )
        analyte_values = sorted(
            {rec.analyte_name for result in self.results for rec in result.normalized_values}
        )
        unit_values = sorted(
            {rec.unit for result in self.results for rec in result.normalized_values if rec.unit}
        )
        metric_values = sorted(
            {rec.metric_role for result in self.results for rec in result.normalized_values}
        )

        self._source_combo["values"] = ("", *source_values)
        self._sample_combo["values"] = ("", *sample_values)
        self._analyte_combo["values"] = ("", *analyte_values)
        self._unit_combo["values"] = ("", *unit_values)
        self._metric_combo["values"] = ("", *metric_values)

        if self._filter_source.get() not in self._source_combo["values"]:
            self._filter_source.set("")
        if self._filter_sample.get() not in self._sample_combo["values"]:
            self._filter_sample.set("")
        if self._filter_analyte.get() not in self._analyte_combo["values"]:
            self._filter_analyte.set("")
        if self._filter_unit.get() not in self._unit_combo["values"]:
            self._filter_unit.set("")
        if self._filter_metric.get() not in self._metric_combo["values"]:
            self._filter_metric.set("")

    def _log_warnings(self, result: WorkbookParseResult) -> None:
        for warning in result.warnings:
            self._log(f"[WARN] {result.source_file}: {warning}")

    def _log(self, message: str) -> None:
        self._log_text.configure(state="normal")
        self._log_text.insert("end", message + "\n")
        self._log_text.see("end")
        self._log_text.configure(state="disabled")

    def _apply_window_icon(self) -> None:
        icon_path = _resource_path(Path("assets") / "favicon.ico")
        if not icon_path.exists():
            return
        try:
            self.root.iconbitmap(default=str(icon_path))
            return
        except tk.TclError:
            pass

        try:
            self._window_icon = tk.PhotoImage(file=str(icon_path))
            self.root.iconphoto(True, self._window_icon)
        except tk.TclError:
            # Some Tk builds cannot read .ico via PhotoImage.
            pass


def _split_csv(raw: str) -> list[str]:
    return [item.strip() for item in raw.split(",") if item.strip()]


def _resource_path(relative_path: str | Path) -> Path:
    if hasattr(sys, "_MEIPASS"):
        base = Path(getattr(sys, "_MEIPASS"))
    else:
        base = Path(__file__).resolve().parents[1]
    return base / Path(relative_path)
