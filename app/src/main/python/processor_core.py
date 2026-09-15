import re
from pathlib import Path
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side
from openpyxl.utils import get_column_letter

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

