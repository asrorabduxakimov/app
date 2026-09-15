"""File operations for the Android UI. Processing stays in processor_core."""
import json
import math
import re
import uuid
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
from processor_core import ExcelProcessor, detect_id_segments


class MobileProcessor(ExcelProcessor):
    def read_file(self, path):
        # The desktop chooser advertises .xls, but openpyxl cannot read it.
        if Path(path).suffix.lower() == ".xls":
            self.log(f"  Чтение: {Path(path).name}")
            return pd.read_excel(path, dtype=str, engine="xlrd")
        return super().read_file(path)


def inspect_file(path, sort_column=""):
    processor = MobileProcessor(log_callback=lambda _: None)
    df = processor.drop_duplicate_header(processor.read_file(path))
    columns = list(df.columns)
    chosen = sort_column if sort_column in columns else next(
        (c for c in columns if any(s in str(c).lower() for s in ("savol", "вопрос"))),
        columns[min(3, len(columns) - 1)] if columns else None,
    )
    segments = detect_id_segments(df[chosen].tolist()) if chosen is not None else []
    return json.dumps({"headers": [str(c) for c in columns], "segments": segments,
                       "sort_column": str(chosen or ""), "rows": len(df)}, ensure_ascii=False)


def validate_options(options):
    allowed = {"add_gender", "add_scores", "multiplier", "sort_column", "header_column",
               "prefix_column", "suffixes", "size_rules"}
    if set(options) - allowed:
        raise ValueError("Неизвестные настройки обработки")
    multiplier = float(options.get("multiplier", 2))
    if not math.isfinite(multiplier) or multiplier < 0:
        raise ValueError("Коэффициент должен быть конечным неотрицательным числом")
    options["multiplier"] = multiplier
    for key in ("sort_column", "header_column", "prefix_column"):
        options[key] = options.get(key) or None
    for rule in options.get("suffixes", []):
        lo, hi = float(rule["from"]), float(rule["to"])
        if not rule["prefix"].strip() or not (math.isfinite(lo) and math.isfinite(hi)) or lo > hi:
            raise ValueError("Проверьте префикс и границы диапазона")
    for rule in options.get("size_rules", []):
        if not rule["prefix"].strip() or int(rule["count"]) != rule["count"] or rule["count"] <= 0:
            raise ValueError("Размер группы должен быть целым положительным числом")
    return options


def output_name(name, used):
    stem = Path(name.replace("\\", "/")).stem
    stem = re.sub(r'[\x00-\x1f<>:"/\\|?*]', "_", stem).strip(" .")[:100] or "file"
    candidate = f"{stem}_processed.xlsx"
    number = 2
    while candidate.casefold() in used:
        candidate = f"{stem}_processed_{number}.xlsx"
        number += 1
    used.add(candidate.casefold())
    return candidate


def process_batch(files_json, options_json, results_root, callback=None):
    files = json.loads(files_json)
    options = validate_options(json.loads(options_json))
    if not files:
        raise ValueError("Добавьте хотя бы один файл")
    run_dir = Path(results_root) / (datetime.now().strftime("%Y%m%d_%H%M%S_") + uuid.uuid4().hex[:8])
    run_dir.mkdir(parents=True)
    results, logs, used = [], [], set()

    def log(message):
        logs.append(str(message))

    for index, item in enumerate(files):
        if callback is not None:
            callback.onProgress(index, len(files), item["name"])
        target = run_dir / output_name(item["name"], used)
        result = {"name": item["name"], "path": "", "participants": 0, "error_rows": 0, "error": ""}
        log(f"\n[{index + 1}/{len(files)}] {item['name']}")
        try:
            processor = MobileProcessor(log_callback=log, **options)
            count, errors = processor.process(item["path"], str(target))
            if not target.is_file():
                raise ValueError("Нет корректных данных: файл результата не создан")
            result.update(path=str(target), participants=count, error_rows=errors)
        except Exception as exc:
            if target.exists():
                target.unlink()  # Only our incomplete output in this new run directory.
            result["error"] = str(exc)
            log(f"ОШИБКА: {exc}")
        results.append(result)

    successful = [r for r in results if r["path"]]
    archive = run_dir / "Attestation_results.zip"
    if successful:
        with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as output:
            for result in successful:
                output.write(result["path"], Path(result["path"]).name)
    payload = {"results": results, "archive": str(archive) if successful else "",
               "log": "\n".join(logs), "success_count": len(successful), "total": len(files)}
    encoded = json.dumps(payload, ensure_ascii=False)
    (run_dir / "report.json").write_text(encoded, encoding="utf-8")
    root = Path(results_root)
    temporary = root / "latest.tmp"
    temporary.write_text(encoded, encoding="utf-8")
    temporary.replace(root / "latest.json")
    if callback is not None:
        callback.onProgress(len(files), len(files), "Готово")
    return encoded
