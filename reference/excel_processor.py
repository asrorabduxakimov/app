"""
Excel Processor — обработка тестовых данных
Использует: pandas, openpyxl, tkinter

Структура выходного листа "Данные":
  A    : ПИНФЛ (уникальный список)
  B    : ФИО   (уникальный список)
  C    : Пол   (опционально, определяется по первой цифре ПИНФЛ)
  D…   : матрица бинарных результатов (1/0)
  D+N… : матрица ABCD-ответов (те же заголовки)
  …    : Jami ball (опционально, сумма баллов × коэффициент)
  …    : To'g'ri javoblar soni (опционально, сумма баллов)
"""

import re
import threading
import os
import sys
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side
from openpyxl.utils import get_column_letter


# ─────────────────────────────────────────────
#  ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ─────────────────────────────────────────────

def question_sort_key(q, suffix_order=None):
    """
    Натуральная сортировка с приоритетом суффиксов.
    Сначала числа без префикса, потом префиксы по алфавиту, потом остальные.
    Внутри каждой группы — по возрастанию числа.
    """
    s = str(q).strip()

    # Чистое число
    try:
        num = float(s)
        if num.is_integer():
            return (0, "", int(num), "", s)
        else:
            return (0, "", num, "", s)
    except (ValueError, TypeError):
        pass

    # Строка с префиксом и числом (например "K-36", "PM-41", "K-6226144654")
    m = re.match(r'^([A-Za-zА-Яа-яЁё\-]+)[^\d]*(\d+)', s)
    if m:
        detected_prefix = m.group(1).upper()
        number_part = int(m.group(2))
        return (1, detected_prefix, number_part, s, "")

    # Всё остальное — в конец
    return (2, "", 0, s, "")



def detect_id_segments(values) -> list:
    """
    Определяет сегменты (непрерывные последовательности подряд идущих чисел)
    в столбце Savol ID. Разрыв — место, где следующее число не равно
    предыдущему + 1.
    Возвращает список {"start": int, "end": int, "count": int}, отсортированный
    по возрастанию.
    """
    nums = []
    for v in values:
        try:
            nums.append(int(float(str(v).strip())))
        except (ValueError, TypeError):
            continue
    nums = sorted(set(nums))
    if not nums:
        return []

    segments = []
    seg_start = prev = nums[0]
    count = 1
    for n in nums[1:]:
        if n == prev + 1:
            count += 1
        else:
            segments.append({"start": seg_start, "end": prev, "count": count})
            seg_start = n
            count = 1
        prev = n
    segments.append({"start": seg_start, "end": prev, "count": count})
    return segments


# ─────────────────────────────────────────────
#  КЛАСС ОБРАБОТКИ
# ─────────────────────────────────────────────

class ExcelProcessor:
    """Бизнес-логика без GUI."""

    def __init__(self, log_callback=None, add_gender=False, add_scores=False, multiplier=2.0,
                 sort_column=None, header_column=None, prefix_column=None, suffixes=None,
                 size_rules=None):
        self.log = log_callback or print
        self.add_gender = add_gender
        self.add_scores = add_scores
        self.multiplier = multiplier
        self.sort_column = sort_column      # столбец для сортировки (Savol ID)
        self.header_column = header_column  # столбец для текста вопроса
        self.prefix_column = prefix_column  # столбец для числового ключа префикса (Savol Tr)
        self.suffixes = list(suffixes) if suffixes else []
        # Сохраняем порядок префиксов как они заданы в GUI
        self.suffixes_ordered = [rule.get("prefix", "") for rule in self.suffixes]
        # Правила "префикс по размеру сегмента Savol ID" (например: 10 шт. -> "PM-").
        # Диапазоны from/to для каждого файла вычисляются заново (см. _apply_size_rules),
        # т.к. в разных файлах абсолютные Savol ID отличаются, но размеры сегментов совпадают.
        self.size_rules = list(size_rules) if size_rules else []

    # ── Чтение ──────────────────────────────────────────────────────
    def read_file(self, path: str) -> pd.DataFrame:
        self.log(f"  Чтение: {Path(path).name}")
        ext = Path(path).suffix.lower()
        if ext in (".xlsx", ".xls", ".xlsm"):
            # Читаем все столбцы как строки, чтобы сохранить формат
            df = pd.read_excel(path, header=0, dtype=str, engine="openpyxl")
        elif ext == ".csv":
            df = pd.read_csv(path, header=0, dtype=str, encoding='utf-8')
        else:
            raise ValueError(f"Неподдерживаемый формат: {ext}")
        self.log(f"     Строк: {len(df):,}")
        return df

    # ── Двойной заголовок ───────────────────────────────────────────
    def drop_duplicate_header(self, df: pd.DataFrame) -> pd.DataFrame:
        if df.empty:
            return df
        first_row = df.iloc[0].astype(str).str.strip().tolist()
        header    = [str(c).strip() for c in df.columns]
        if first_row == header:
            self.log("     Дублирующий заголовок удалён.")
            df = df.iloc[1:].reset_index(drop=True)
        return df

    # ── Получение заголовков для анализа ────────────────────────────
    def get_headers(self, df: pd.DataFrame) -> list:
        """Возвращает список заголовков столбцов"""
        return list(df.columns)

    # ── Выбор нужных столбцов ───────────────────────────────────────
    def select_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Выбирает нужные столбцы на основе конфигурации.
        sort_column   — столбец для сортировки вопросов (числовой ключ).
        header_column — столбец из которого берём текст заголовка вопроса.
        prefix_column — столбец для определения префикса (числовой ключ).
        """
        cols = list(df.columns)
        self.log(f"     Найдено столбцов: {len(cols)}")

        pinfl_idx   = None
        fio_idx     = None
        sort_idx    = None   # столбец-ключ сортировки
        header_idx  = None   # столбец-текст заголовка
        prefix_idx  = None   # столбец-ключ префикса
        answer_idx  = None
        result_idx  = None

        # Автопоиск по именам
        for i, col in enumerate(cols):
            col_lower = str(col).lower().strip()
            if 'pinfl' in col_lower or 'пинфл' in col_lower:
                pinfl_idx = i
            elif 'fio' in col_lower or 'фио' in col_lower:
                fio_idx = i
            elif 'yuklangan' in col_lower and 'variant' in col_lower:
                answer_idx = i
            elif "to'g'riligi" in col_lower or 'natija' in col_lower or 'результат' in col_lower:
                result_idx = i

        # Столбец сортировки (выбранный пользователем)
        if self.sort_column is not None:
            for i, col in enumerate(cols):
                if col == self.sort_column:
                    sort_idx = i
                    break

        # Столбец заголовков (выбранный пользователем)
        if self.header_column is not None:
            for i, col in enumerate(cols):
                if col == self.header_column:
                    header_idx = i
                    break

        # Столбец префикса (выбранный пользователем)
        if self.prefix_column is not None:
            for i, col in enumerate(cols):
                if col == self.prefix_column:
                    prefix_idx = i
                    break

        # Если sort_column не выбран — ищем автоматически
        if sort_idx is None:
            for i, col in enumerate(cols):
                col_lower = str(col).lower().strip()
                if 'savol tr' in col_lower or 'вопрос' in col_lower or 'savol' in col_lower:
                    sort_idx = i
                    break

        # Если header_column не выбран — по умолчанию равен sort_column
        if header_idx is None:
            header_idx = sort_idx

        # Если prefix_column не выбран — по умолчанию None (будет использоваться sort_val)
        # prefix_idx остаётся None

        # Дефолтные значения
        if pinfl_idx  is None: pinfl_idx  = 0;  self.log("     ВНИМАНИЕ: PINFL не найден, используется столбец 0")
        if fio_idx    is None: fio_idx    = 1;  self.log("     ВНИМАНИЕ: ФИО не найден, используется столбец 1")
        if sort_idx   is None: sort_idx   = 3;  self.log("     ВНИМАНИЕ: Столбец сортировки не найден, используется столбец 3")
        if header_idx is None: header_idx = sort_idx
        if answer_idx is None: answer_idx = 9;  self.log("     ВНИМАНИЕ: Ответ не найден, используется столбец 9")
        if result_idx is None: result_idx = 10; self.log("     ВНИМАНИЕ: Результат не найден, используется столбец 10")

        max_idx = len(cols) - 1
        pinfl_idx   = min(pinfl_idx,   max_idx)
        fio_idx     = min(fio_idx,     max_idx)
        sort_idx    = min(sort_idx,    max_idx)
        header_idx  = min(header_idx,  max_idx)
        answer_idx  = min(answer_idx,  max_idx)
        result_idx  = min(result_idx,  max_idx)

        self.log(f"     Используемые столбцы:")
        self.log(f"       PINFL:           {cols[pinfl_idx]}  (индекс {pinfl_idx})")
        self.log(f"       ФИО:             {cols[fio_idx]}  (индекс {fio_idx})")
        self.log(f"       Сортировка:      {cols[sort_idx]}  (индекс {sort_idx})")
        self.log(f"       Заголовок:       {cols[header_idx]}  (индекс {header_idx})")
        if prefix_idx is not None:
            self.log(f"       Префикс (ключ):  {cols[prefix_idx]}  (индекс {prefix_idx})")
        else:
            self.log(f"       Префикс (ключ):  используется столбец сортировки")
        self.log(f"       Ответ:           {cols[answer_idx]}  (индекс {answer_idx})")
        self.log(f"       Результат:       {cols[result_idx]}  (индекс {result_idx})")

        # Собираем нужные столбцы
        needed = [pinfl_idx, fio_idx, sort_idx, answer_idx, result_idx]
        col_names = ["PINFL", "ФИО", "СортКлюч", "Ответ", "Результат"]

        # Вставляем header_idx после sort_idx
        if header_idx not in needed:
            needed.insert(3, header_idx)
            col_names.insert(3, "Заголовок")
        else:
            # если header_idx уже есть (например, равен sort_idx), не дублируем
            pass

        # Вставляем prefix_idx, если задан и ещё не добавлен
        if prefix_idx is not None and prefix_idx not in needed:
            # вставляем после header_idx (перед "Ответ")
            insert_pos = col_names.index("Ответ") if "Ответ" in col_names else len(needed)
            needed.insert(insert_pos, prefix_idx)
            col_names.insert(insert_pos, "ПрефиксКлюч")

        df_selected = df.iloc[:, needed].copy()
        df_selected.columns = col_names

        # ВАЖНО: Преобразуем СортКлюч в строку явно
        df_selected["СортКлюч"] = df_selected["СортКлюч"].astype(str).str.strip()
        
        # Если есть отдельный Заголовок, тоже преобразуем в строку
        if "Заголовок" in df_selected.columns:
            df_selected["Заголовок"] = df_selected["Заголовок"].astype(str).str.strip()
        else:
            df_selected["Заголовок"] = df_selected["СортКлюч"]
            
        # Если есть ПрефиксКлюч, преобразуем в строку
        if "ПрефиксКлюч" in df_selected.columns:
            df_selected["ПрефиксКлюч"] = df_selected["ПрефиксКлюч"].astype(str).str.strip()

        # Анализ сегментов Savol ID (разрывы порядка)
        segments = self._log_id_segments(df_selected["СортКлюч"])

        # Применяем правила "префикс по размеру сегмента" (если заданы через
        # окно "Группы Savol ID") — диапазоны вычисляются заново для каждого файла
        self._apply_size_rules(segments)

        # Строим итоговый столбец «Вопрос» с учётом суффиксов
        # Используем ПрефиксКлюч если есть, иначе СортКлюч
        if "ПрефиксКлюч" in df_selected.columns:
            df_selected["Вопрос"] = df_selected.apply(
                lambda row: self._apply_suffix(
                    row["СортКлюч"],
                    row["Заголовок"],
                    row["ПрефиксКлюч"]
                ), axis=1
            )
        else:
            df_selected["Вопрос"] = df_selected.apply(
                lambda row: self._apply_suffix(
                    row["СортКлюч"],
                    row["Заголовок"],
                    None
                ), axis=1
            )

        # Нормализация: один СортКлюч → одно значение Вопрос (первое встреченное).
        # Без этого один и тот же вопрос может получить разные имена у разных участников
        # (например, если ПрефиксКлюч варьируется по строкам для одного СортКлюч).
        first_label = df_selected.groupby("СортКлюч", sort=False)["Вопрос"].first()
        df_selected["Вопрос"] = df_selected["СортКлюч"].map(first_label)

        # Возвращаем минимально необходимые столбцы
        return_cols = ["PINFL", "ФИО", "СортКлюч", "Вопрос", "Ответ", "Результат"]
        return df_selected[return_cols]

    # ── Анализ сегментов Savol ID ────────────────────────────────────
    def _log_id_segments(self, sort_key_series):
        """
        Находит сегменты (непрерывные подряд идущие числа) в столбце Savol ID
        и пишет результат в лог. Если сегментов больше 3 — предупреждает.
        Возвращает список сегментов для дальнейшего использования (см. _apply_size_rules).
        """
        segments = detect_id_segments(sort_key_series.tolist())
        if not segments:
            return []

        n_seg = len(segments)
        gaps  = max(n_seg - 1, 0)
        sizes = [s["count"] for s in segments]

        self.log(f"     Savol ID — сегментов: {n_seg} (разрывов: {gaps}), размеры: {sizes}")
        for i, s in enumerate(segments, 1):
            self.log(f"       Сегмент {i}: {s['start']}–{s['end']}  ({s['count']} шт.)")

        if n_seg > 3:
            self.log(f"     ВНИМАНИЕ: обнаружено {n_seg} сегментов Savol ID (больше 3) — проверьте данные!")

        return segments

    # ── Применение правил "префикс по размеру сегмента" ──────────────
    def _apply_size_rules(self, segments):
        """
        Сопоставляет сегменты ТЕКУЩЕГО файла с правилами self.size_rules
        (заданными по количеству вопросов в сегменте, а не по абсолютным Savol ID —
        абсолютные значения в разных файлах разные, а количество вопросов в
        группе одинаковое). Найденные соответствия добавляются в self.suffixes
        с диапазоном from/to именно этого файла.
        """
        if not self.size_rules or not segments:
            return

        for rule in self.size_rules:
            size   = rule.get("count")
            prefix = rule.get("prefix", "")
            matches = [s for s in segments if s["count"] == size]

            if not matches:
                self.log(f"     ВНИМАНИЕ: в этом файле нет сегмента размером {size} шт. "
                          f"— префикс '{prefix}' не применён")
                continue

            if len(matches) > 1:
                self.log(f"     ВНИМАНИЕ: найдено {len(matches)} сегментов размером {size} шт. "
                          f"— префикс '{prefix}' применён ко всем")

            for seg in matches:
                self.suffixes.append({"prefix": prefix, "from": seg["start"], "to": seg["end"]})
                self.log(f"     Группа {size} шт. в этом файле: {seg['start']}–{seg['end']} "
                          f"-> префикс '{prefix}'")

    def _apply_suffix(self, sort_val, header_val, prefix_val) -> str:
        """
        Применяет префикс к значению заголовка.
        Ключ для диапазонов берётся из prefix_val (если задан), иначе из sort_val.
        """
        header_str = str(header_val).strip()
        
        # Определяем источник числового ключа
        key_source = prefix_val if prefix_val is not None else sort_val

        try:
            key = float(str(key_source).strip())
        except (ValueError, TypeError):
            return header_str

        # Проверяем правила суффиксов
        for rule in self.suffixes:
            try:
                lo = float(rule.get("from", 0))
                hi = float(rule.get("to", 0))
                if lo <= key <= hi:
                    prefix = str(rule.get("prefix", ""))
                    return f"{prefix}{header_str}"
            except (ValueError, TypeError):
                continue

        return header_str

    # ── Очистка ─────────────────────────────────────────────────────
    def clean(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.dropna(how="all")

        df["PINFL"] = (df["PINFL"].astype(str).str.strip()
                                  .str.replace(r"\D", "", regex=True))
        df = df[df["PINFL"] != ""]
        df["PINFL"] = pd.to_numeric(df["PINFL"], errors="coerce")
        df = df.dropna(subset=["PINFL"])
        df["PINFL"] = df["PINFL"].astype(np.int64)

        df["Вопрос"] = df["Вопрос"].astype(str).str.strip()
        df = df[df["Вопрос"].notna() & (df["Вопрос"] != "") & (df["Вопрос"] != "nan")]
        df["Результат"] = (pd.to_numeric(df["Результат"].astype(str).str.strip(),
                                         errors="coerce")
                           .fillna(0).astype(np.int8))
        df["ФИО"]   = df["ФИО"].astype(str).str.strip()
        df["Ответ"] = df["Ответ"].astype(str).str.strip().str.upper()
        
        # Убеждаемся, что СортКлюч остаётся строкой
        df["СортКлюч"] = df["СортКлюч"].astype(str).str.strip()

        self.log(f"     Строк после очистки: {len(df):,}")
        return df

    # ── Целостность ─────────────────────────────────────────────────
    def check_integrity(self, df: pd.DataFrame):
        counts       = df.groupby("PINFL", sort=False).size()
        if len(counts) == 0:
            return df, pd.DataFrame(), 0
        normal_count = int(counts.mode().iloc[0])
        self.log(f"     Норма строк / участник: {normal_count}")

        bad = counts[counts != normal_count].index
        if len(bad):
            self.log(f"     Аномалий: {len(bad)} PINFL -> {list(bad[:5])}"
                     f"{'...' if len(bad) > 5 else ''}")
        else:
            self.log("     Все участники имеют одинаковое число строк.")

        mask = df["PINFL"].isin(bad)
        return df[~mask].copy(), df[mask].copy(), normal_count

    # ── Сортировка ──────────────────────────────────────────────────
    def sort_data(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        
        # Убеждаемся, что СортКлюч - строка
        df["СортКлюч"] = df["СортКлюч"].astype(str).str.strip()
        
        suffix_order = self.suffixes_ordered

        def get_sort_key(val):
            return question_sort_key(val, suffix_order)

        df["_sort_key"] = df["Вопрос"].apply(get_sort_key)
        df = (df.sort_values(["PINFL", "_sort_key"])
                .drop(columns=["_sort_key"])
                .reset_index(drop=True))
        return df

    # ── Участники ───────────────────────────────────────────────────
    def get_participants(self, df: pd.DataFrame) -> pd.DataFrame:
        parts = (df[["PINFL", "ФИО"]]
                 .drop_duplicates(subset=["PINFL", "ФИО"])
                 .sort_values("PINFL")
                 .reset_index(drop=True))
        
        if self.add_gender:
            parts["Пол"] = parts["PINFL"].apply(self._determine_gender)
            self.log(f"     Добавлен столбец 'Пол'")
        
        self.log(f"     Уникальных участников: {len(parts):,}")
        return parts

    def _determine_gender(self, pinfl):
        try:
            pinfl_str = str(int(pinfl))
            first_digit = int(pinfl_str[0])
            return "Ayol" if first_digit % 2 == 0 else "Erkak"
        except:
            return ""

    # ── Список вопросов ─────────────────────────────────────────────
    def get_sorted_questions(self, df: pd.DataFrame) -> list:
        df_unique = df[["СортКлюч", "Вопрос"]].drop_duplicates()
        
        # Убеждаемся, что СортКлюч - строка
        df_unique["СортКлюч"] = df_unique["СортКлюч"].astype(str).str.strip()
        
        suffix_order = self.suffixes_ordered

        def get_sort_key(val):
            return question_sort_key(val, suffix_order)

        df_unique["_sort_key"] = df_unique["Вопрос"].apply(get_sort_key)
        df_sorted = df_unique.sort_values("_sort_key")
        
        questions_list = df_sorted["Вопрос"].tolist()
        if questions_list:
            self.log(f"     Первые 5 вопросов: {questions_list[:5]}")
        
        return questions_list

    # ── Матрицы ─────────────────────────────────────────────────────
    def build_matrices(self, df, participants, questions):
        pinfl_order = participants["PINFL"].tolist()

        bin_mx = (df.pivot_table(index="PINFL", columns="Вопрос",
                                 values="Результат", aggfunc="first")
                    .reindex(columns=questions).reindex(pinfl_order))

        abcd_mx = (df.pivot_table(index="PINFL", columns="Вопрос",
                                  values="Ответ", aggfunc="first")
                     .reindex(columns=questions).reindex(pinfl_order))

        self.log(f"     Матрица: {len(pinfl_order)} x {len(questions)}")
        return bin_mx, abcd_mx

    # ── Сохранение ──────────────────────────────────────────────────
    def save(self, df_errors, participants, bin_mx, abcd_mx, questions, out_path):
        self.log("  Запись файла...")
        n_q = len(questions)
        n_p = len(participants)

        pinfl_arr = participants["PINFL"].to_numpy()
        fio_arr   = participants["ФИО"].to_numpy()
        bin_arr   = bin_mx.to_numpy()
        abcd_arr  = abcd_mx.to_numpy()

        participant_cols = ["PINFL", "ФИО"]
        if self.add_gender:
            gender_arr = participants["Пол"].to_numpy()
            participant_cols.append("Пол")

        wb = Workbook()

        # ── Лист "Данные" ──────────────────────────────────────────
        ws = wb.active
        ws.title = "Данные"
        bold = Font(bold=True)
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        col_idx = 1
        ws.cell(row=1, column=col_idx, value="ПИНФЛ").font = bold
        col_idx += 1
        ws.cell(row=1, column=col_idx, value="ФИО").font = bold
        col_idx += 1
        
        if self.add_gender:
            ws.cell(row=1, column=col_idx, value="Пол").font = bold
            col_idx += 1

        abcd_start_col = col_idx
        for j, q in enumerate(questions):
            ws.cell(row=1, column=col_idx + j, value=q).font = bold

        bin_start_col = col_idx + n_q
        for j, q in enumerate(questions):
            ws.cell(row=1, column=bin_start_col + j, value=q).font = bold

        scores_start_col = bin_start_col + n_q
        if self.add_scores:
            ws.cell(row=1, column=scores_start_col, value="Jami ball").font = bold
            ws.cell(row=1, column=scores_start_col + 1, value="To'g'ri javoblar soni").font = bold

        for i in range(n_p):
            r = i + 2
            col_idx = 1
            
            cell = ws.cell(row=r, column=col_idx, value=int(pinfl_arr[i]))
            cell.border = thin_border
            col_idx += 1
            
            cell = ws.cell(row=r, column=col_idx, value=str(fio_arr[i]))
            cell.border = thin_border
            col_idx += 1
            
            if self.add_gender:
                cell = ws.cell(row=r, column=col_idx, value=gender_arr[i])
                cell.border = thin_border
                col_idx += 1

            for j in range(n_q):
                av = abcd_arr[i, j]
                val = str(av) if (pd.notna(av) and str(av) not in ("nan", "")) else ""
                cell = ws.cell(row=r, column=abcd_start_col + j, value=val)
                cell.border = thin_border

            row_sum = 0
            for j in range(n_q):
                bv = bin_arr[i, j]
                if isinstance(bv, float) and np.isnan(bv):
                    val = 0
                else:
                    val = int(bv) if bv is not None else 0

                row_sum += val
                cell = ws.cell(row=r, column=bin_start_col + j, value=val)
                cell.border = thin_border

            if self.add_scores:
                cell = ws.cell(row=r, column=scores_start_col, value=row_sum * self.multiplier)
                cell.border = thin_border
                cell = ws.cell(row=r, column=scores_start_col + 1, value=row_sum)
                cell.border = thin_border

        last_col = scores_start_col + (1 if self.add_scores else -1)
        if not self.add_scores:
            last_col = bin_start_col + n_q - 1
            
        for col in range(1, last_col + 1):
            cell = ws.cell(row=1, column=col)
            cell.border = thin_border

        for column in ws.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if cell.value:
                        if isinstance(cell.value, (int, float)):
                            length = len(str(cell.value))
                        else:
                            length = len(str(cell.value))
                        if length > max_length:
                            max_length = length
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = max(adjusted_width, 10)

        # ── Лист "Ошибки" ─────────────────────────────────────────
        if not df_errors.empty:
            ws_err = wb.create_sheet("Ошибки")
            err_cols = list(df_errors.columns)
            for c_idx, col_name in enumerate(err_cols, start=1):
                cell = ws_err.cell(row=1, column=c_idx, value=col_name)
                cell.font = bold
                cell.border = thin_border
            for r_idx, row_data in enumerate(df_errors.itertuples(index=False), start=2):
                for c_idx, val in enumerate(row_data, start=1):
                    cell = ws_err.cell(row=r_idx, column=c_idx,
                                      value=None if pd.isna(val) else val)
                    cell.border = thin_border
            for column in ws_err.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        if cell.value:
                            length = len(str(cell.value))
                            if length > max_length:
                                max_length = length
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                ws_err.column_dimensions[column_letter].width = max(adjusted_width, 10)
            self.log(f"     Лист 'Ошибки': {len(df_errors):,} строк")

        wb.save(out_path)
        self.log(f"  Сохранён: {Path(out_path).name}")

    # ── Главный метод для одного файла ──────────────────────────────
    def process(self, in_path: str, out_path: str):
        df               = self.read_file(in_path)
        df               = self.drop_duplicate_header(df)
        df               = self.select_columns(df)
        df               = self.clean(df)
        df_ok, df_err, _ = self.check_integrity(df)
        if df_ok.empty:
            self.log("     ВНИМАНИЕ: нет корректных данных для обработки")
            return 0, len(df_err)
        df_ok            = self.sort_data(df_ok)
        participants     = self.get_participants(df_ok)
        questions        = self.get_sorted_questions(df_ok)
        bin_mx, abcd_mx  = self.build_matrices(df_ok, participants, questions)
        self.save(df_err, participants, bin_mx, abcd_mx, questions, out_path)
        return len(participants), len(df_err)


# ─────────────────────────────────────────────
#  GUI
# ─────────────────────────────────────────────

DARK_BG   = "#0F172A"
CARD_BG   = "#1E293B"
ACCENT    = "#3B82F6"
ACCENT2   = "#6366F1"
TEXT      = "#F1F5F9"
TEXT_DIM  = "#94A3B8"
SUCCESS   = "#22C55E"
ERROR_COL = "#EF4444"
WARNING   = "#F59E0B"
BORDER    = "#334155"
ROW_ALT   = "#162032"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel Processor")
        self.geometry("900x800")
        self.minsize(800, 650)
        self.configure(bg=DARK_BG)

        self._file_rows: list[dict] = []
        self._out_folder = tk.StringVar()
        self._status     = tk.StringVar(value="Добавьте файлы для обработки")
        self._progress   = tk.DoubleVar(value=0)
        
        self._add_gender = tk.BooleanVar(value=False)
        self._add_scores = tk.BooleanVar(value=False)
        self._multiplier = tk.DoubleVar(value=2.0)
        self._sort_column   = tk.StringVar(value="")
        self._header_column = tk.StringVar(value="")
        self._prefix_column = tk.StringVar(value="")
        
        self._headers = []
        self._headers_loaded = False

        # Правила "префикс по размеру сегмента Savol ID" — [{"count":..,"prefix":..}, ...].
        # Хранятся отдельно от суффиксов, т.к. работают по количеству вопросов
        # в группе, а не по абсолютным Savol ID (которые различаются между файлами).
        self._size_group_rules = []
        self._size_rules_summary = tk.StringVar(value="Группы Savol ID: не заданы")

        self._build_ui()
        self._center()

    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        # Шапка
        hdr = tk.Frame(self, bg=CARD_BG)
        hdr.pack(fill="x")
        tk.Frame(hdr, bg=ACCENT, height=3).pack(fill="x")
        ih = tk.Frame(hdr, bg=CARD_BG, padx=24, pady=14)
        ih.pack(fill="x")
        tk.Label(ih, text="Excel Processor", bg=CARD_BG, fg=TEXT,
                 font=("Helvetica", 17, "bold")).pack(side="left")
        tk.Label(ih, text="Обработка тестовых данных", bg=CARD_BG, fg=TEXT_DIM,
                 font=("Helvetica", 10)).pack(side="left", padx=(10,0), pady=(4,0))

        body = tk.Frame(self, bg=DARK_BG, padx=20, pady=16)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(4, weight=1)

        # ── Панель инструкции ────────────────────────────────────
        self._info_open = tk.BooleanVar(value=False)
        info_outer = tk.Frame(body, bg=CARD_BG)
        info_outer.grid(row=0, column=0, sticky="ew", pady=(0, 8))

        info_hdr = tk.Frame(info_outer, bg=CARD_BG, padx=12, pady=7, cursor="hand2")
        info_hdr.pack(fill="x")
        self._info_arrow = tk.Label(info_hdr, text="▶", bg=CARD_BG, fg=ACCENT,
                                    font=("Helvetica", 9, "bold"))
        self._info_arrow.pack(side="left")
        tk.Label(info_hdr, text="  Инструкция: формат входного файла и логика обработки",
                 bg=CARD_BG, fg=TEXT_DIM, font=("Helvetica", 9, "bold")
                 ).pack(side="left")

        self._info_body = tk.Frame(info_outer, bg="#0D1117", padx=14, pady=10)

        INST = (
            ("ВХОДНОЙ ФАЙЛ — Столбцы определяются автоматически или вручную", True),
            ("", False),
            ("  PINFL         Идентификатор участника (определяется по имени)", False),
            ("  ФИО           Фамилия Имя Отчество (определяется по имени)", False),
            ("  Вопрос        Номер вопроса (можно выбрать вручную)", False),
            ("  Ответ         Вариант ответа (определяется по имени)", False),
            ("  Результат     Балл: 1 (верно) или 0 (неверно)", False),
            ("", False),
            ("ВЫХОДНОЙ ФАЙЛ — лист «Данные»", True),
            ("", False),
            ("  A  ПИНФЛ        Уникальный список участников (по возрастанию)", False),
            ("  B  ФИО          ФИО участника", False),
            ("  C  Пол          (опционально) Erkak/Ayol по 1-й цифре ПИНФЛ", False),
            ("  D… Вопрос 1…N   Бинарная матрица результатов (1 / 0)", False),
            ("  …  Вопрос 1…N   ABCD-матрица ответов (те же заголовки)", False),
            ("  …  Jami ball    (опционально) сумма баллов × коэффициент", False),
            ("  …  To'g'ri javoblar soni (опционально) сумма баллов", False),
            ("", False),
            ("ЛОГИКА ОБРАБОТКИ", True),
            ("", False),
            ("  1. Если 2-я строка дублирует заголовок — удаляется автоматически.", False),
            ("  2. Столбцы определяются по именам или выбираются вручную.", False),
            ("  3. PINFL очищается от нецифровых символов и приводится к int.", False),
            ("  4. Вопросы сортируются: числовые → префиксы по порядку → остальные.", False),
            ("  5. Участники с аномальным числом строк → лист «Ошибки».", False),
        )

        for text, is_title in INST:
            if not text:
                tk.Label(self._info_body, text="", bg="#0D1117", height=0).pack(anchor="w")
                continue
            fg    = ACCENT if is_title else TEXT_DIM
            font  = ("Courier", 9, "bold") if is_title else ("Courier", 9)
            tk.Label(self._info_body, text=text, bg="#0D1117", fg=fg,
                     font=font, anchor="w", justify="left").pack(anchor="w")

        def _toggle_info(_event=None):
            if self._info_open.get():
                self._info_body.pack_forget()
                self._info_arrow.configure(text="▶")
                self._info_open.set(False)
            else:
                self._info_body.pack(fill="x")
                self._info_arrow.configure(text="▼")
                self._info_open.set(True)

        info_hdr.bind("<Button-1>", _toggle_info)
        for child in info_hdr.winfo_children():
            child.bind("<Button-1>", _toggle_info)

        # ── Панель настроек ──────────────────────────────────────
        settings_frame = tk.Frame(body, bg=CARD_BG, padx=12, pady=8)
        settings_frame.grid(row=1, column=0, sticky="ew", pady=(0, 8))

        tk.Label(settings_frame, text="Настройки обработки:",
                 bg=CARD_BG, fg=TEXT_DIM, font=("Helvetica", 9, "bold")
                 ).pack(anchor="w", pady=(0, 6))

        # Первая строка (чекбоксы и множитель)
        settings_row1 = tk.Frame(settings_frame, bg=CARD_BG)
        settings_row1.pack(fill="x", pady=(0, 5))

        cb_gender = tk.Checkbutton(settings_row1, text="Добавить столбец 'Пол'",
                                   variable=self._add_gender,
                                   bg=CARD_BG, fg=TEXT, selectcolor=CARD_BG,
                                   activebackground=CARD_BG,
                                   font=("Helvetica", 9))
        cb_gender.pack(side="left", padx=(0, 20))

        cb_scores = tk.Checkbutton(settings_row1, text="Добавить расчетные столбцы",
                                   variable=self._add_scores,
                                   command=self._toggle_multiplier,
                                   bg=CARD_BG, fg=TEXT, selectcolor=CARD_BG,
                                   activebackground=CARD_BG,
                                   font=("Helvetica", 9))
        cb_scores.pack(side="left")

        mult_frame = tk.Frame(settings_row1, bg=CARD_BG)
        mult_frame.pack(side="left", padx=(20, 0))
        tk.Label(mult_frame, text="Множитель:", bg=CARD_BG, fg=TEXT_DIM,
                 font=("Helvetica", 9)).pack(side="left", padx=(0, 5))
        self._mult_entry = tk.Entry(mult_frame, textvariable=self._multiplier,
                                    bg="#0D1117", fg=TEXT, width=5,
                                    relief="flat", font=("Helvetica", 9),
                                    highlightthickness=1, highlightbackground=BORDER,
                                    state="disabled")
        self._mult_entry.pack(side="left")

        # Вторая строка - выбор столбцов
        settings_row2 = tk.Frame(settings_frame, bg=CARD_BG)
        settings_row2.pack(fill="x", pady=(5, 0))

        # Сортировать по
        tk.Label(settings_row2, text="Сортировать по:", bg=CARD_BG, fg=TEXT_DIM,
                 font=("Helvetica", 9)).pack(side="left", padx=(0, 6))
        self._sort_combo = ttk.Combobox(settings_row2, textvariable=self._sort_column,
                                        state="readonly", width=18)
        self._sort_combo.pack(side="left", padx=(0, 12))

        # Заголовок из
        tk.Label(settings_row2, text="Заголовок из:", bg=CARD_BG, fg=TEXT_DIM,
                 font=("Helvetica", 9)).pack(side="left", padx=(0, 6))
        self._header_combo = ttk.Combobox(settings_row2, textvariable=self._header_column,
                                          state="readonly", width=18)
        self._header_combo.pack(side="left", padx=(0, 12))

        # Префикс из
        tk.Label(settings_row2, text="Префикс из:", bg=CARD_BG, fg=TEXT_DIM,
                 font=("Helvetica", 9)).pack(side="left", padx=(0, 6))
        self._prefix_combo = ttk.Combobox(settings_row2, textvariable=self._prefix_column,
                                          state="readonly", width=18)
        self._prefix_combo.pack(side="left", padx=(0, 12))

        # Кнопка загрузить заголовки
        btn_load_headers = tk.Button(settings_row2, text="Загрузить заголовки",
                                     command=self._load_headers_from_file,
                                     bg=BORDER, fg=TEXT, activebackground=ACCENT,
                                     activeforeground="white", relief="flat",
                                     cursor="hand2", font=("Helvetica", 9), padx=10, pady=2)
        btn_load_headers.pack(side="left")
        self._hover(btn_load_headers, BORDER, ACCENT)

        btn_id_groups = tk.Button(settings_row2, text="Группы Savol ID",
                                  command=self._open_id_groups_window,
                                  bg=BORDER, fg=TEXT, activebackground=ACCENT,
                                  activeforeground="white", relief="flat",
                                  cursor="hand2", font=("Helvetica", 9), padx=10, pady=2)
        btn_id_groups.pack(side="left", padx=(8, 0))
        self._hover(btn_id_groups, BORDER, ACCENT)

        # Сводка активных правил "префикс по размеру группы" (действуют одинаково
        # на все файлы в списке, т.к. привязаны к количеству вопросов, а не к
        # абсолютным Savol ID конкретного файла)
        tk.Label(settings_row2, textvariable=self._size_rules_summary,
                 bg=CARD_BG, fg=TEXT_DIM, font=("Helvetica", 8, "italic")
                 ).pack(side="left", padx=(12, 0))

        # Третья строка — суффиксы
        suf_outer = tk.Frame(settings_frame, bg=CARD_BG)
        suf_outer.pack(fill="x", pady=(8, 0))

        tk.Label(suf_outer, text="Суффиксы (префикс + диапазон значений столбца «Префикс из» или «Сортировать по»):",
                 bg=CARD_BG, fg=TEXT_DIM, font=("Helvetica", 9)).pack(anchor="w", pady=(0, 4))

        self._suffix_rows = []
        self._suffix_frames_container = tk.Frame(suf_outer, bg=CARD_BG)
        self._suffix_frames_container.pack(fill="x")

        for i in range(3):
            self._add_suffix_row(i + 1)

        # ── Папка вывода ──────────────────────────────────────────
        out_f = tk.Frame(body, bg=CARD_BG, padx=12, pady=8)
        out_f.grid(row=2, column=0, sticky="ew", pady=(0, 10))

        tk.Label(out_f, text="Папка для сохранения результатов:",
                 bg=CARD_BG, fg=TEXT_DIM, font=("Helvetica", 9, "bold")
                 ).pack(anchor="w", pady=(0, 4))

        row_out = tk.Frame(out_f, bg=CARD_BG)
        row_out.pack(fill="x")
        tk.Entry(row_out, textvariable=self._out_folder,
                 bg="#0D1117", fg=TEXT, insertbackground=TEXT,
                 relief="flat", font=("Helvetica", 10),
                 highlightthickness=1, highlightbackground=BORDER,
                 highlightcolor=ACCENT
                 ).pack(side="left", fill="x", expand=True, padx=(0,10), ipady=5)
        btn_of = tk.Button(row_out, text="Выбрать...",
                           command=self._browse_out_folder,
                           bg=BORDER, fg=TEXT, activebackground=ACCENT,
                           activeforeground="white", relief="flat",
                           cursor="hand2", font=("Helvetica", 9), padx=12, pady=5)
        btn_of.pack(side="right")
        self._hover(btn_of, BORDER, ACCENT)

        # ── Тулбар файлов ────────────────────────────────────────
        tb = tk.Frame(body, bg=DARK_BG)
        tb.grid(row=3, column=0, sticky="ew", pady=(0, 6))

        tk.Label(tb, text="Файлы для обработки:", bg=DARK_BG, fg=TEXT_DIM,
                 font=("Helvetica", 9, "bold")).pack(side="left")

        for txt, cmd, bg in [("+ Добавить файлы",    self._add_files,    ACCENT),
                              ("+ Добавить папку",    self._add_folder,   BORDER),
                              ("✕ Удалить выбранное", self._remove_files, "#4B1C1C")]:
            b = tk.Button(tb, text=txt, command=cmd,
                          bg=bg, fg=TEXT, activebackground=ACCENT2,
                          activeforeground="white", relief="flat",
                          cursor="hand2", font=("Helvetica", 9), padx=10, pady=4)
            b.pack(side="left", padx=(6, 0))
            self._hover(b, bg, ACCENT2)

        # ── Таблица файлов ────────────────────────────────────────
        tbl_wrap = tk.Frame(body, bg=BORDER, padx=1, pady=1)
        tbl_wrap.grid(row=4, column=0, sticky="nsew")

        self._canvas = tk.Canvas(tbl_wrap, bg=DARK_BG, highlightthickness=0)
        vsb = tk.Scrollbar(tbl_wrap, orient="vertical",
                           command=self._canvas.yview,
                           bg=CARD_BG, troughcolor=DARK_BG)
        self._canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self._canvas.pack(side="left", fill="both", expand=True)

        hdr_row = tk.Frame(tbl_wrap, bg=CARD_BG)
        hdr_row.place(relx=0, rely=0, relwidth=1, height=28)
        for txt, w in [("#", 30), ("Входной файл", 0), ("Выходной файл", 220), ("Статус", 110)]:
            tk.Label(hdr_row, text=txt, bg=CARD_BG, fg=TEXT_DIM,
                     font=("Helvetica", 8, "bold"),
                     width=w if w else 0, anchor="w"
                     ).pack(side="left", padx=(8,0))

        self._rows_frame = tk.Frame(self._canvas, bg=DARK_BG)
        self._canvas_window = self._canvas.create_window(
            (0, 28), window=self._rows_frame, anchor="nw")
        self._rows_frame.bind("<Configure>", self._on_frame_configure)
        self._canvas.bind("<Configure>",     self._on_canvas_configure)
        self._canvas.bind_all("<MouseWheel>",
                              lambda e: self._canvas.yview_scroll(
                                  -1 if e.delta > 0 else 1, "units"))

        # ── Нижняя панель ─────────────────────────────────────────
        bot = tk.Frame(body, bg=DARK_BG)
        bot.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        bot.columnconfigure(0, weight=1)

        self._run_btn = tk.Button(
            bot, text="▶  Запустить обработку",
            command=self._run,
            bg=ACCENT, fg="white",
            activebackground=ACCENT2, activeforeground="white",
            font=("Helvetica", 12, "bold"),
            relief="flat", cursor="hand2", padx=28, pady=11)
        self._run_btn.grid(row=0, column=0, sticky="ew")
        self._hover(self._run_btn, ACCENT, ACCENT2)

        self._results = []  # список (имя_файла, кол-во участников) в порядке обработки
        self._copy_counts_btn = tk.Button(
            bot, text="📋 Копировать файл + участники",
            command=self._copy_participant_counts,
            bg=BORDER, fg=TEXT,
            activebackground=ACCENT2, activeforeground="white",
            font=("Helvetica", 10), relief="flat", cursor="hand2",
            padx=14, pady=11, state="disabled")
        self._copy_counts_btn.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self._hover(self._copy_counts_btn, BORDER, ACCENT2)

        style = ttk.Style(); style.theme_use("clam")
        style.configure("TProgressbar", troughcolor=CARD_BG, background=ACCENT,
                        bordercolor=BORDER, lightcolor=ACCENT, darkcolor=ACCENT)
        self._pbar = ttk.Progressbar(bot, variable=self._progress,
                                     maximum=100, mode="determinate")
        self._pbar.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))

        self._status_lbl = tk.Label(bot, textvariable=self._status,
                                    bg=DARK_BG, fg=TEXT_DIM,
                                    font=("Helvetica", 10), anchor="w")
        self._status_lbl.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(4, 0))

        # ── Лог ───────────────────────────────────────────────────
        tk.Label(body, text="Журнал", bg=DARK_BG, fg=TEXT_DIM,
                 font=("Helvetica", 9, "bold"), anchor="w"
                 ).grid(row=6, column=0, pady=(14, 4), sticky="ew")
        body.rowconfigure(7, weight=1)

        log_f = tk.Frame(body, bg=BORDER, padx=1, pady=1)
        log_f.grid(row=7, column=0, sticky="nsew")

        self._log_box = tk.Text(
            log_f, bg="#0D1117", fg=TEXT, insertbackground=TEXT,
            font=("Courier", 9), relief="flat", wrap="word",
            state="disabled", height=7)
        sb2 = tk.Scrollbar(log_f, command=self._log_box.yview,
                           bg=CARD_BG, troughcolor=DARK_BG)
        self._log_box.configure(yscrollcommand=sb2.set)
        self._log_box.pack(side="left", fill="both", expand=True)
        sb2.pack(side="right", fill="y")
        self._log_box.tag_config("ok",    foreground=SUCCESS)
        self._log_box.tag_config("warn",  foreground=WARNING)
        self._log_box.tag_config("error", foreground=ERROR_COL)
        self._log_box.tag_config("info",  foreground=TEXT_DIM)

    def _load_headers_from_file(self):
        """Загружает заголовки из первого файла в списке"""
        if not self._file_rows:
            messagebox.showwarning("Предупреждение", "Сначала добавьте файлы для обработки")
            return
        
        try:
            first_file = self._file_rows[0]["in_path"]
            df = pd.read_excel(first_file, header=0, nrows=0)
            self._headers = list(df.columns)

            self._sort_combo['values']   = self._headers
            self._header_combo['values'] = self._headers
            self._prefix_combo['values'] = self._headers
            
            # Автовыбор: ищем столбцы по именам
            for col in self._headers:
                col_lower = str(col).lower().strip()
                if 'savol tr' in col_lower or 'вопрос' in col_lower:
                    self._sort_column.set(col)
                    self._header_column.set(col)
                    self._prefix_column.set(col)
            
            self._headers_loaded = True
            self._log("Заголовки загружены успешно")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить заголовки:\n{str(e)}")

    def _open_id_groups_window(self):
        """
        Отдельное окно: анализирует столбец Savol ID («Сортировать по») из
        первого файла, находит сегменты (непрерывные диапазоны чисел) и
        позволяет назначить каждому сегменту префикс. Применение переносит
        сегменты в существующий список суффиксов (от/до).
        """
        if not self._file_rows:
            messagebox.showwarning("Предупреждение", "Сначала добавьте файлы для обработки")
            return

        sort_col = self._sort_column.get().strip()
        if not sort_col:
            messagebox.showwarning("Предупреждение",
                                   "Сначала выберите столбец «Сортировать по» (Savol ID)")
            return

        try:
            first_file = self._file_rows[0]["in_path"]
            df = pd.read_excel(first_file, header=0, dtype=str, usecols=[sort_col])
            values = df[sort_col].tolist()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось прочитать столбец «{sort_col}»:\n{e}")
            return

        segments = detect_id_segments(values)
        if not segments:
            messagebox.showwarning("Предупреждение",
                                   f"В столбце «{sort_col}» не найдено числовых значений")
            return

        n_seg = len(segments)
        segments_by_size = sorted(segments, key=lambda s: s["count"], reverse=True)

        def _size_label(rank, total):
            if total == 1:
                return "Единственный"
            if rank == 0:
                return "Наибольший"
            if rank == total - 1:
                return "Наименьший"
            if total == 3:
                return "Средний"
            return f"Средний {rank}"

        win = tk.Toplevel(self)
        win.title("Группы Savol ID")
        win.configure(bg=DARK_BG)
        win.geometry(f"480x{140 + 34 * n_seg}")
        win.transient(self)
        win.grab_set()

        warn = n_seg > 3
        summary = f"Найдено сегментов: {n_seg} (разрывов: {max(n_seg - 1, 0)})"
        if warn:
            summary += "  ⚠ больше 3 — проверьте данные"
        tk.Label(win, text=summary, bg=DARK_BG, fg=(WARNING if warn else TEXT),
                 font=("Helvetica", 10, "bold"), wraplength=450, justify="left"
                 ).pack(anchor="w", padx=14, pady=(12, 8))

        rows_frame = tk.Frame(win, bg=DARK_BG)
        rows_frame.pack(fill="both", expand=True, padx=14)

        counts = [seg["count"] for seg in segments_by_size]
        if len(set(counts)) != len(counts):
            tk.Label(win, text="⚠ Есть сегменты одинакового размера — им будет назначен один и тот же префикс",
                     bg=DARK_BG, fg=WARNING, font=("Helvetica", 8), wraplength=450, justify="left"
                     ).pack(anchor="w", padx=14, pady=(0, 6))

        prefix_vars = []
        existing_by_count = {r["count"]: r["prefix"] for r in self._size_group_rules}
        for i, seg in enumerate(segments_by_size):
            label = _size_label(i, n_seg)
            row = tk.Frame(rows_frame, bg=DARK_BG)
            row.pack(fill="x", pady=3)
            tk.Label(row, text=f"{label} ({seg['count']} шт.): {seg['start']}–{seg['end']}",
                     bg=DARK_BG, fg=TEXT, font=("Helvetica", 9), width=32, anchor="w"
                     ).pack(side="left")
            tk.Label(row, text="Префикс:", bg=DARK_BG, fg=TEXT_DIM,
                     font=("Helvetica", 9)).pack(side="left", padx=(6, 4))
            pv = tk.StringVar(value=existing_by_count.get(seg["count"], ""))
            tk.Entry(row, textvariable=pv, width=8, bg="#0D1117", fg=TEXT,
                     insertbackground=TEXT, relief="flat", font=("Helvetica", 9),
                     highlightthickness=1, highlightbackground=BORDER
                     ).pack(side="left", ipady=2)
            prefix_vars.append(pv)

        tk.Label(win, text="Правила запоминаются по количеству вопросов в группе — "
                            "будут работать одинаково для всех файлов, даже если "
                            "у них разные Savol ID.",
                 bg=DARK_BG, fg=TEXT_DIM, font=("Helvetica", 8), wraplength=450, justify="left"
                 ).pack(anchor="w", padx=14, pady=(6, 0))

        def _apply():
            self._size_group_rules = [
                {"count": seg["count"], "prefix": pv.get().strip()}
                for seg, pv in zip(segments_by_size, prefix_vars)
                if pv.get().strip()
            ]
            if self._size_group_rules:
                parts = [f"{r['count']}→{r['prefix']}" for r in self._size_group_rules]
                self._size_rules_summary.set("Группы Savol ID: " + "  ".join(parts))
            else:
                self._size_rules_summary.set("Группы Savol ID: не заданы")
            self._log(f"Сохранены группы Savol ID по размеру: {self._size_group_rules}")
            win.destroy()

        btn_frame = tk.Frame(win, bg=DARK_BG)
        btn_frame.pack(fill="x", padx=14, pady=12)
        tk.Button(btn_frame, text="Сохранить правила", command=_apply,
                  bg=ACCENT, fg="white", activebackground=ACCENT2, activeforeground="white",
                  relief="flat", cursor="hand2", font=("Helvetica", 9, "bold"), padx=10, pady=6
                  ).pack(side="right")
        tk.Button(btn_frame, text="Отмена", command=win.destroy,
                  bg=BORDER, fg=TEXT, activebackground=ACCENT2, activeforeground="white",
                  relief="flat", cursor="hand2", font=("Helvetica", 9), padx=10, pady=6
                  ).pack(side="right", padx=(0, 8))


    def _add_suffix_row(self, num: int, prefix: str = "", from_val: str = "", to_val: str = ""):
        sf_frame = tk.Frame(self._suffix_frames_container, bg=CARD_BG)
        sf_frame.pack(fill="x", pady=2)

        tk.Label(sf_frame, text=f"  #{num}", bg=CARD_BG, fg=TEXT_DIM,
                 font=("Helvetica", 9), width=4).pack(side="left")

        tk.Label(sf_frame, text="Префикс:", bg=CARD_BG, fg=TEXT_DIM,
                 font=("Helvetica", 9)).pack(side="left", padx=(0, 4))
        prefix_var = tk.StringVar(value=prefix)
        tk.Entry(sf_frame, textvariable=prefix_var, width=8,
                 bg="#0D1117", fg=TEXT, insertbackground=TEXT, relief="flat",
                 font=("Helvetica", 9), highlightthickness=1, highlightbackground=BORDER
                 ).pack(side="left", padx=(0, 12), ipady=2)

        tk.Label(sf_frame, text="от:", bg=CARD_BG, fg=TEXT_DIM,
                 font=("Helvetica", 9)).pack(side="left", padx=(0, 4))
        from_var = tk.StringVar(value=from_val)
        tk.Entry(sf_frame, textvariable=from_var, width=7,
                 bg="#0D1117", fg=TEXT, insertbackground=TEXT, relief="flat",
                 font=("Helvetica", 9), highlightthickness=1, highlightbackground=BORDER
                 ).pack(side="left", padx=(0, 8), ipady=2)

        tk.Label(sf_frame, text="до:", bg=CARD_BG, fg=TEXT_DIM,
                 font=("Helvetica", 9)).pack(side="left", padx=(0, 4))
        to_var = tk.StringVar(value=to_val)
        tk.Entry(sf_frame, textvariable=to_var, width=7,
                 bg="#0D1117", fg=TEXT, insertbackground=TEXT, relief="flat",
                 font=("Helvetica", 9), highlightthickness=1, highlightbackground=BORDER
                 ).pack(side="left", ipady=2)

        tk.Label(sf_frame,
                 text="→ Префикс + значение из столбца «Заголовок»",
                 bg=CARD_BG, fg=TEXT_DIM, font=("Helvetica", 8)
                 ).pack(side="left", padx=(12, 0))

        self._suffix_rows.append({"prefix": prefix_var, "from": from_var, "to": to_var})

    def _toggle_multiplier(self):
        if self._add_scores.get():
            self._mult_entry.configure(state="normal")
        else:
            self._mult_entry.configure(state="disabled")

    def _on_frame_configure(self, event=None):
        bbox = self._canvas.bbox("all")
        if bbox:
            # Явно фиксируем верх scrollregion в 0, а не в bbox[1] (=28),
            # иначе при прокрутке в самый верх первая строка уезжает под
            # шапку таблицы (hdr_row наложена сверху через .place).
            self._canvas.configure(scrollregion=(0, 0, bbox[2], bbox[3]))

    def _on_canvas_configure(self, event=None):
        self._canvas.itemconfig(self._canvas_window, width=event.width)

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="Выберите файлы",
            filetypes=[("Excel / CSV", "*.xlsx *.xls *.xlsm *.csv"),
                       ("Все файлы", "*.*")])
        for p in paths:
            self._add_row(p)
        if self._file_rows and not self._headers_loaded:
            self._load_headers_from_file()

    def _add_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку с файлами")
        if not folder:
            return
        exts = {".xlsx", ".xls", ".xlsm", ".csv"}
        for p in sorted(Path(folder).glob("*")):
            if p.suffix.lower() in exts:
                self._add_row(str(p))
        if self._file_rows and not self._headers_loaded:
            self._load_headers_from_file()

    def _add_row(self, in_path: str):
        existing = [r["in_path"] for r in self._file_rows]
        if in_path in existing:
            return

        out_folder = self._out_folder.get().strip()
        if out_folder:
            out_path = str(Path(out_folder) / f"{Path(in_path).stem}_processed.xlsx")
        else:
            out_path = str(Path(in_path).parent / f"{Path(in_path).stem}_processed.xlsx")

        status_var = tk.StringVar(value="Ожидание")
        idx = len(self._file_rows)
        bg  = DARK_BG if idx % 2 == 0 else ROW_ALT

        row_frame = tk.Frame(self._rows_frame, bg=bg)
        row_frame.pack(fill="x", pady=0)

        var_check = tk.BooleanVar(value=False)
        tk.Checkbutton(row_frame, variable=var_check,
                       bg=bg, activebackground=bg,
                       selectcolor=CARD_BG, fg=TEXT_DIM
                       ).pack(side="left", padx=(4,0))

        tk.Label(row_frame, text=Path(in_path).name, bg=bg, fg=TEXT,
                 font=("Courier", 9), anchor="w", width=32
                 ).pack(side="left", padx=(4,0))

        out_var = tk.StringVar(value=out_path)
        e = tk.Entry(row_frame, textvariable=out_var, bg="#0D1117", fg=TEXT,
                     insertbackground=TEXT, relief="flat",
                     font=("Courier", 9), width=28,
                     highlightthickness=1, highlightbackground=BORDER)
        e.pack(side="left", padx=(8,0), ipady=2)

        status_lbl = tk.Label(row_frame, textvariable=status_var,
                               bg=bg, fg=TEXT_DIM,
                               font=("Helvetica", 8), width=50, anchor="w")
        status_lbl.pack(side="left", padx=(10,0))

        self._file_rows.append({
            "in_path":    in_path,
            "out_var":    out_var,
            "status_var": status_var,
            "status_lbl": status_lbl,
            "check_var":  var_check,
            "frame":      row_frame,
            "bg":         bg,
        })

    def _remove_files(self):
        to_keep = []
        for row in self._file_rows:
            if row["check_var"].get():
                row["frame"].destroy()
            else:
                to_keep.append(row)
        self._file_rows = to_keep
        for i, row in enumerate(self._file_rows):
            bg = DARK_BG if i % 2 == 0 else ROW_ALT
            row["bg"] = bg
            row["frame"].configure(bg=bg)
            for w in row["frame"].winfo_children():
                try: w.configure(bg=bg)
                except: pass

    def _set_row_status(self, row: dict, text: str, color: str):
        row["status_var"].set(text)
        row["status_lbl"].configure(fg=color)

    def _browse_out_folder(self):
        folder = filedialog.askdirectory(title="Выберите папку для результатов")
        if folder:
            self._out_folder.set(folder)
            for row in self._file_rows:
                old = Path(row["out_var"].get()).name
                row["out_var"].set(str(Path(folder) / old))

    def _hover(self, w, n, h):
        w.bind("<Enter>", lambda e: w.config(bg=h))
        w.bind("<Leave>", lambda e: w.config(bg=n))

    def _log(self, msg: str):
        self._log_box.after(0, self._append_log, msg)

    def _append_log(self, msg: str):
        self._log_box.configure(state="normal")
        tag = "default"
        stripped = msg.lstrip()
        if any(x in msg for x in ("Сохранён", "Завершено", "Матрица", "участников")):
            tag = "ok"
        elif any(x in msg for x in ("Аномалий", "Дублир", "ВНИМАНИЕ")):
            tag = "warn"
        elif "ОШИБКА" in msg or "Ошибка" in msg:
            tag = "error"
        elif stripped.startswith(("Строк", "Норма", "Уникальных", "Лист", "Выбраны", "Найдено", "Добавлен", "Используемые", "Заголовки", "Первые")):
            tag = "info"
        self._log_box.insert("end", msg + "\n", tag)
        self._log_box.configure(state="disabled")
        self._log_box.see("end")

    def _set_status(self, msg, color=TEXT_DIM):
        self._status.set(msg)
        self._status_lbl.configure(fg=color)

    def _set_progress(self, val):
        self._progress.set(val)
        self.update_idletasks()

    def _run(self):
        if not self._file_rows:
            messagebox.showerror("Ошибка", "Добавьте хотя бы один файл.")
            return

        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

        self._run_btn.configure(state="disabled", text="Обработка...")
        self._copy_counts_btn.configure(state="disabled")
        self._set_status("Обработка...", WARNING)
        self._set_progress(0)

        threading.Thread(target=self._worker, daemon=True).start()

    def _worker(self):
        rows   = list(self._file_rows)
        total  = len(rows)
        errors = []
        self._results = []

        sort_col   = self._sort_column.get().strip() or None
        header_col = self._header_column.get().strip() or None
        prefix_col = self._prefix_column.get().strip() or None

        suffixes = []
        for sf in self._suffix_rows:
            prefix = sf["prefix"].get().strip()
            from_v = sf["from"].get().strip()
            to_v   = sf["to"].get().strip()
            if prefix and from_v and to_v:
                try:
                    suffixes.append({"prefix": prefix,
                                     "from": float(from_v),
                                     "to":   float(to_v)})
                except ValueError:
                    pass

        for idx, row in enumerate(rows):
            in_p  = row["in_path"]
            out_p = row["out_var"].get().strip()

            if not out_p:
                out_p = str(Path(in_p).parent / f"{Path(in_p).stem}_processed.xlsx")

            self.after(0, self._set_row_status, row, "Обработка...", WARNING)
            self._log(f"\n[{idx+1}/{total}] {Path(in_p).name}")

            try:
                proc = ExcelProcessor(
                    log_callback=self._log,
                    add_gender=self._add_gender.get(),
                    add_scores=self._add_scores.get(),
                    multiplier=self._multiplier.get(),
                    sort_column=sort_col,
                    header_column=header_col,
                    prefix_column=prefix_col,
                    suffixes=suffixes,
                    size_rules=self._size_group_rules
                )
                n_p, n_e = proc.process(in_p, out_p)
                label = f"OK ({n_p} уч.)" + (f" /{n_e} с ошибкой" if n_e else "")
                self.after(0, self._set_row_status, row, label, SUCCESS)
                self._results.append((Path(in_p).name, n_p))
            except Exception as exc:
                self._log(f"  ОШИБКА: {exc}")
                self._log(traceback.format_exc())
                self.after(0, self._set_row_status, row, "Ошибка", ERROR_COL)
                errors.append((Path(in_p).name, str(exc)))
                self._results.append((Path(in_p).name, "ОШИБКА"))

            self.after(0, self._set_progress, (idx + 1) / total * 100)

        self.after(0, self._done, total, errors)

    def _done(self, total, errors):
        self._set_progress(100)
        if self._results:
            self._copy_counts_btn.configure(state="normal")
        if errors:
            self._set_status(f"Завершено с ошибками: {len(errors)}/{total}", WARNING)
            err_list = "\n".join(f"• {n}: {e}" for n, e in errors)
            messagebox.showwarning(
                "Завершено с ошибками",
                f"Обработано: {total}\nОшибок: {len(errors)}\n\n{err_list}")
        else:
            self._set_status(f"Все {total} файлов обработаны успешно!", SUCCESS)
            self._log(f"\nГотово! Обработано файлов: {total}")

            out_folder = self._out_folder.get().strip()
            if not out_folder and self._file_rows:
                out_folder = str(Path(self._file_rows[0]["out_var"].get()).parent)

            if messagebox.askyesno("Готово",
                                   f"Все {total} файлов обработаны!\n\n"
                                   f"Открыть папку с результатами?"):
                folder = out_folder or str(Path(self._file_rows[0]["out_var"].get()).parent)
                if sys.platform   == "win32":  os.startfile(folder)
                elif sys.platform == "darwin": os.system(f'open "{folder}"')
                else:                          os.system(f'xdg-open "{folder}"')

        self._run_btn.configure(state="normal", text="▶  Запустить обработку")

    def _failed(self, msg):
        self._set_status(f"Ошибка: {msg}", ERROR_COL)
        self._run_btn.configure(state="normal", text="▶  Запустить обработку")
        messagebox.showerror("Ошибка", f"Ошибка:\n\n{msg}")

    def _copy_participant_counts(self):
        """Копирует «имя файла<TAB>кол-во участников» по одной строке на файл,
        чтобы можно было вставить сразу двумя столбцами в Excel."""
        if not self._results:
            messagebox.showinfo("Список участников", "Сначала запустите обработку файлов.")
            return
        text = "\n".join(f"{name}\t{count}" for name, count in self._results)
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update()
        self._set_status(f"Скопировано {len(self._results)} строк в буфер обмена", SUCCESS)


# ─────────────────────────────────────────────
#  ТОЧКА ВХОДА
# ─────────────────────────────────────────────

def main():
    missing = []
    for pkg in ("pandas", "openpyxl", "numpy"):
        try:
            __import__(pkg)
        except ImportError:
            missing.append(pkg)
    if missing:
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("Зависимости",
                             f"Установите пакеты:\n\npip install {' '.join(missing)}")
        root.destroy()
        return
    App().mainloop()


if __name__ == "__main__":
    main()