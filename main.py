# -*- coding: utf-8 -*-
"""1000BASE-T / 100BASE-TX / 10BASE-T: извлечение PDF -> предварительный просмотр/правка -> Word."""
import os
import re
import tempfile
from decimal import Decimal, ROUND_HALF_UP
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
import fitz
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

FONT_NAME = "Arial"


def round_text(value, digits, decimal_mark=","):
    """Decimal ROUND_HALF_UP; всегда выводит ровно `digits` знаков после запятой."""
    quantum = Decimal("1").scaleb(-digits)
    rounded = Decimal(str(value).replace(",", ".")).quantize(quantum, rounding=ROUND_HALF_UP)
    result = format(rounded, ".{}f".format(digits))
    return result.replace(".", decimal_mark)


def norm(s):
    return re.sub(r"\s+", " ", s or "").strip().lower()


def first_value(text, pattern, label):
    m = re.search(pattern, text, re.I)
    if not m:
        raise ValueError("Не найдено значение в PDF: " + label)
    return m.group(1).strip()


# ----------------------------------------------------------------------------
# 100BASE-TX
# ----------------------------------------------------------------------------
def parse_pdf(path):
    pdf = fitz.open(path)
    if len(pdf) < 12:
        raise ValueError("Ожидался PDF-отчёт минимум на 12 страниц.")
    pages = [pdf[i].get_text("text") for i in range(12)]
    p2, p4, p5, p6, p7, p8, p9, p10, p11, p12 = (
        pages[1], pages[3], pages[4], pages[5], pages[6],
        pages[7], pages[8], pages[9], pages[10], pages[11])

    def get(pattern, text, label):
        return first_value(text, pattern, label)

    vals = {}
    vals["amp_pos"] = round_text(get(r"Positive Amplitude\(\+Vout\)\s*:\s*([0-9.]+)mV", p2, "амплитуда +"), 0)
    vals["amp_neg"] = round_text(abs(Decimal(get(r"Negative Amplitude\(-Vout\)\s*:\s*(-?[0-9.]+)mV", p2, "амплитуда −"))), 0)
    
    vals["overshoot_pos"] = round_text(get(r"Overshoot\(\+ve\)\s*:?\s*([0-9.]+)%", p4, "выброс +"), 2)
    vals["overshoot_neg"] = round_text(get(r"Overshoot\(-ve\)\s*:?\s*([0-9.]+)%", p4, "выброс −"), 2)
    vals["amp_sym"] = round_text(get(r"Amplitude Symmetry\s*:?\s*([0-9.]+)", p5, "симметрия амплитуды"), 3)

    vals["edges"] = [
        round_text(get(r"Rise Time\(\+ve\)\s*:\s*([0-9.]+)ns", p6, "передний фронт +"), 2),
        round_text(get(r"Rise Time\(-ve\)\s*:\s*([0-9.]+)ns", p6, "передний фронт −"), 2),
        round_text(get(r"Fall Time\(\+ve\)\s*:\s*([0-9.]+)ns", p7, "задний фронт +"), 2),
        round_text(get(r"Fall Time\(-ve\)\s*:\s*([0-9.]+)ns", p7, "задний фронт −"), 2)]

    def symmetry_to_ns(pattern, label):
        raw = get(pattern, p8, label)
        match = re.fullmatch(r"([0-9.]+)\s*(ps|fs)", raw, re.I)
        if not match:
            raise ValueError("Не удалось распознать единицу симметрии: " + raw)
        amount = Decimal(match.group(1))
        unit = match.group(2).lower()
        ns = amount / Decimal("1000.0" if unit == "ps" else "1000000.0")
        return round_text(ns, 3)

    vals["edge_sym"] = [
        symmetry_to_ns(r"Rise/Fall Symmetry\(\+ve\)\s*:\s*([0-9.]+\s*(?:ps|fs))", "симметрия +"),
        symmetry_to_ns(r"Rise/Fall Symmetry\(-ve\)\s*:\s*([0-9.]+\s*(?:ps|fs))", "симметрия −")]
    
    jitter_pos = Decimal(get(r"Transmit Jitter\(\+ve\)\s*:?\s*([0-9.]+)ps", p11, "джиттер +")) / Decimal("1000.0")
    jitter_neg = Decimal(get(r"Transmit Jitter\(-ve\)\s*:?\s*([0-9.]+)ps", p11, "джиттер −")) / Decimal("1000.0")
    vals["jitter"] = [round_text(jitter_pos, 2), round_text(jitter_neg, 2)]

    duty = get(r"Distortion\(Duty Cycle\)\s*:?\s*([0-9.]+)ps", p9, "искажение коэффициента заполнения")
    duty_ns = Decimal(duty) / Decimal("1000.0")
    vals["duty"] = "{} (± {})".format(round_text(duty_ns, 3), round_text(duty_ns / 2, 3))

    def return_loss(page_text, title):
        rows = []
        for line in page_text.splitlines():
            nums = re.findall(r"-\s*\d+(?:\.\d+)?(?=dB)", line)
            if len(nums) >= 3:
                rows.append([abs(float(x.replace(" ", ""))) for x in nums[:3]])
        if len(rows) < 9:
            raise ValueError("Не удалось извлечь 9 строк затухания отражения ({})".format(title))
        if title == "приёмник":
            rows = list(reversed(rows[:9]))
        return [[round_text(x, 1) for x in row] for row in rows[:9]]

    vals["tx_rl"] = return_loss(p10, "передатчик")
    vals["rx_rl"] = return_loss(p12, "приёмник")
    return vals


# ----------------------------------------------------------------------------
# 10BASE-T
# ----------------------------------------------------------------------------
def parse_10baset_pdf(path):
    pdf = fitz.open(path)
    if len(pdf) < 19:
        raise ValueError("Ожидался PDF-отчёт 10BASE-T минимум на 19 страниц.")
    pages = [page.get_text("text") for page in pdf]

    def find(pattern, page_no, label, flags=re.I):
        m = re.search(pattern, pages[page_no - 1], flags)
        if not m:
            raise ValueError("10BASE-T: не найден параметр «{}» на странице {}.".format(label, page_no))
        return m

    v = {}
    t = pages[12]
    for polarity, prefix in (("Pos", "amp_pos"), ("Neg", "amp_neg")):
        mx = find(r"Max\s+" + polarity + r"\s+Peak\s+Diff\s+Voltage\s*:\s*([-+]?\d+(?:\.\d+)?)", 13, prefix + " max").group(1)
        mn_match = re.search(r"Min\s+" + polarity + r"\s+Peak\s+Diff\s+Voltage\s*:\s*([-+]?\d+(?:\.\d+)?)", t, re.I)
        max_value = abs(Decimal(mx))
        min_value = abs(Decimal(mn_match.group(1))) if mn_match else None
        
        if min_value is None or min_value == max_value:
            v[prefix] = round_text(max_value, 3)
        else:
            lo, hi = sorted((min_value, max_value))
            v[prefix] = "От {} до {}".format(round_text(lo, 3), round_text(hi, 3))

    m = find(r"Fundamental\s*:\s*([\d.]+)\s*MHz\s*([-+]?\s*[\d.]+)\s*dB", 14, "Fundamental")
    fundamental_db = Decimal(m.group(2).replace(" ", ""))
    harmonics = []
    harmonic_pattern = r"Harmonic\s*\d+\s*:?\s*([\d.]+)\s*MHz\s*([-+]?\s*[\d.]+)\s*dB"
    for page_text in [pages[13]] + [t for i, t in enumerate(pages) if i != 13]:
        found = re.findall(harmonic_pattern, page_text, re.I)
        if found:
            harmonics = [(Decimal(db.replace(" ", "")), Decimal(freq)) for freq, db in found]
            break

    if not harmonics:
        raise ValueError("10BASE-T: не удалось распознать строки Harmonic на странице 14.")
    
    harmonic_db, harmonic_freq = max(harmonics, key=lambda item: item[0])
    v["harmonic_diff"] = "{} дБ; {} МГц".format(round_text(fundamental_db - harmonic_db, 1), round_text(harmonic_freq, 0))

    v["period"] = round_text(Decimal(find(r"Average\s+of\s+50\s+Link\s+Intervals\s*=\s*([\d.]+)\s*ms", 6, "Average of 50 Link Intervals").group(1)), 3)

    def impedance_table(page_no, title):
        text = pages[page_no - 1]
        rows = []
        for line in text.splitlines():
            nums = re.findall(r"-?\s*\d+(?:\.\d+)?\s*dB", line, re.I)
            if len(nums) >= 3:
                rows.append([Decimal(re.search(r"\d+(?:\.\d+)?", x).group()) for x in nums[:3]])
        if len(rows) < 2:
            raise ValueError("10BASE-T: не удалось извлечь 2 строки по трём нагрузкам: {} (стр. {}).".format(title, page_no))
        return rows[:2]

    v["imp_tx"] = impedance_table(15, "передатчик")
    v["imp_rx"] = impedance_table(19, "приёмник")

    without = Decimal(find(r"Pk-Pk\s+Jitter\s+without\s+Cable\s*:?\s*([\d.]+)\s*n", 17, "Jitter without Cable").group(1))
    with_cable = Decimal(find(r"Pk-Pk\s+Jitter\s+with\s+Cable\s*:?\s*([\d.]+)\s*n", 18, "Jitter with Cable").group(1))
    v["jitter_without"] = "± {} ({})".format(round_text(without / 2, 2), round_text(without, 2))
    v["jitter_with"] = "± {} ({})".format(round_text(with_cable / 2, 2), round_text(with_cable, 2))

    cm = Decimal(find(r"Common-mode\s+10BaseT\s+output\s+voltage\s*:?\s*([\d.]+)", 16, "Common-mode output voltage").group(1))
    v["common_mode"] = round_text(cm, 1)
    return v


# ----------------------------------------------------------------------------
# 1000BASE-T (Улучшенный парсинг для таблиц)
# ----------------------------------------------------------------------------
def parse_1000baset_pdf(path, label):
    pdf = fitz.open(path)
    if len(pdf) < 1:
        raise ValueError("Пустой PDF-файл.")
    text = pdf[0].get_text("text")
    
    # Ищем все значения вида -16.00dB или −16.00dB (учитывая возможные пробелы и спецсимволы)
    # В таблице 11 строк по 4 значения (Spec + 3 измерения) = 44 значения
    db_values = re.findall(r"([-\u2212]\s*\d+(?:[.,]\d+)?)\s*dB", text, re.I)
    
    if len(db_values) < 44:
        raise ValueError(f"1000BASE-T ({label}): не удалось найти 44 значения dB (найдено {len(db_values)}).")
    
    rows = []
    # Обрабатываем первые 44 значения, разбивая их по 4 штуки на строку
    for i in range(0, 44, 4):
        try:
            # Пропускаем db_values[i] (это Spec. Value)
            # Извлекаем три значения (85, 100, 115 Ом)
            vals = []
            for j in range(1, 4):
                num_match = re.search(r"\d+(?:[.,]\d+)?", db_values[i+j])
                if num_match:
                    vals.append(abs(float(num_match.group().replace(",", "."))))
            
            if len(vals) == 3:
                rows.append(vals)
        except Exception:
            continue
            
    if len(rows) < 11:
        raise ValueError(f"1000BASE-T ({label}): удалось извлечь только {len(rows)} строк из 11.")
    
    return [[round_text(x, 1) for x in row] for row in rows[:11]]


# ----------------------------------------------------------------------------
# Заполнение Word — 10BASE-T
# ----------------------------------------------------------------------------
def fill_10baset_docx(doc, v):
    errors = []

    def table(clause, min_rows, min_cols, title):
        matches = []
        for t in doc.tables:
            txt = table_text(t)
            if clause.lower() in txt and len(t.rows) >= min_rows and len(t.columns) >= min_cols:
                matches.append(t)
        if not matches:
            raise ValueError("В Word-форме не найдена таблица 10BASE-T для {}.".format(title))
        return matches[0]

    try:
        t_amp = table("7.6.2.1", 5, 5, "1.2")
        t_harm = next((t for t in doc.tables if "7.6.2.2" in table_text(t) and len(t.rows) >= 4), None)
        t_period = next((t for t in doc.tables if "результат измерения, мс" in table_text(t) and "16 ± 8" in table_text(t)), None)
        t_tx = next((t for t in doc.tables if "7.6.2.4" in table_text(t) and "импеданс, ом" in table_text(t)), None)
        t_jit = next((t for t in doc.tables if "7.6.2.5" in table_text(t) and "выходной временной джиттер" not in table_text(t)), None)
        t_cm = next((t for t in doc.tables if "7.6.2.7" in table_text(t) and "результат измерения, мв" in table_text(t)), None)
        candidates = [t for t in doc.tables if "7.6.2.4" in table_text(t) and "импеданс, ом" in table_text(t)]
        t_rx = candidates[1] if len(candidates) > 1 else None
        if not all([t_amp, t_harm, t_period, t_tx, t_jit, t_cm, t_rx]):
            raise ValueError("Не удалось однозначно определить все таблицы 10BASE-T в Word-форме.")

        put(t_amp, 3, 2, v["amp_pos"])
        put(t_amp, 4, 2, v["amp_neg"])
        harmonic_difference, harmonic_frequency = [part.strip() for part in v["harmonic_diff"].split(";")]
        put(t_harm, 3, 0, harmonic_frequency.replace("МГц", "").strip())
        put(t_harm, 3, 1, harmonic_difference.replace("дБ", "").strip())
        put(t_period, 3, 1, v["period"])
        
        for table_obj, key in ((t_tx, "imp_tx"), (t_rx, "imp_rx")):
            row = 3
            for frequency_row in v[key]:
                for value in frequency_row:
                    put(table_obj, row, 3, round_text(value, 1))
                    row += 1
                    
        put(t_jit, 3, 1, v["jitter_without"])
        put(t_jit, 4, 1, v["jitter_with"])
        put(t_cm, 3, 1, v["common_mode"])
    except Exception as exc:
        errors.append(str(exc))
    return errors


# ----------------------------------------------------------------------------
# Заполнение Word — 1000BASE-T (Поиск по 7.6.4.8)
# ----------------------------------------------------------------------------
def fill_1000baset_docx(doc, v_a, v_b):
    errors = []
    try:
        # Ищем таблицы по пункту СТБ 7.6.4.8
        rl_1000 = find_tables(doc, "7.6.4.8", ("частота",))
        if len(rl_1000) < 2:
            # Резервный поиск только по номеру пункта
            rl_1000 = find_tables(doc, "7.6.4.8")
        
        if len(rl_1000) < 2:
            raise ValueError("В Word-форме не найдены 2 таблицы для 1000BASE-T (пункт 7.6.4.8).")
        
        # Сортируем таблицы, чтобы Порт A был первым, если в тексте есть "порт a" или "port a"
        tables_a = [t for t in rl_1000 if "порт a" in table_text(t) or "port a" in table_text(t)]
        tables_b = [t for t in rl_1000 if "порт b" in table_text(t) or "port b" in table_text(t)]
        
        if tables_a and tables_b:
            rl_1000 = [tables_a[0], tables_b[0]]
        else:
            rl_1000 = rl_1000[:2] # Берем первые две найденные

        for table_obj, values_matrix in zip(rl_1000, [v_a, v_b]):
            if not values_matrix:
                continue
            row = 3
            for freq_row in values_matrix:
                for ci, value in enumerate(freq_row):
                    put(table_obj, row, 2 + ci, value)
                row += 1
    except Exception as exc:
        errors.append("1000BASE-T: " + str(exc))
    return errors


def table_text(table):
    return norm(" ".join(cell.text for row in table.rows for cell in row.cells))


def find_tables(doc, clause, required=()):
    found = []
    for t in doc.tables:
        txt = table_text(t)
        if clause.lower() in txt and all(norm(x) in txt for x in required):
            found.append(t)
    return found


def require_one(tables, clause):
    if not tables:
        raise ValueError("В выбранной Word-форме не найдена таблица для пункта {}.".format(clause))
    return tables[0]


def set_cell(cell, value):
    cell.text = str(value)
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in paragraph.runs:
            run.font.name = FONT_NAME
            rpr = run._element.get_or_add_rPr()
            rpr.rFonts.set(qn("w:eastAsia"), FONT_NAME)
            rpr.rFonts.set(qn("w:ascii"), FONT_NAME)
            rpr.rFonts.set(qn("w:hAnsi"), FONT_NAME)


def put(table, row, col, value):
    if row >= len(table.rows) or col >= len(table.columns):
        raise ValueError("В Word-форме не найдена ячейка: строка {}, столбец {}.".format(row + 1, col + 1))
    set_cell(table.cell(row, col), value)


# ----------------------------------------------------------------------------
# Заполнение Word — 100BASE-TX
# ----------------------------------------------------------------------------
def fill_docx(template, output, v):
    doc = Document(template)
    errors = []
    try:
        t_amp = require_one(find_tables(doc, "7.6.3.1", ("полярность импульса",)), "2.2")
        t_mask = require_one(find_tables(doc, "7.6.3.1", ("номер сегмента маски",)), "2.3")
        t_over = require_one(find_tables(doc, "7.6.3.2", ("полярность импульса",)), "2.4")
        t_sym = require_one(find_tables(doc, "7.6.3.3", ("результат измерения",)), "2.5")
        rl = find_tables(doc, "7.6.3.4", ("частота, мгц", "импеданс, ом"))
        if len(rl) < 2:
            raise ValueError("Не найдены обе таблицы затухания отражения 100BASE-TX.")
        t_edges = require_one(find_tables(doc, "7.6.3.5", ("фронт импульса",)), "2.7")
        edge_sym_tables = [t for t in find_tables(doc, "7.6.3.5", ("полярность импульса", "результат измерения, нс"))
                           if "фронт импульса" not in table_text(t)]
        t_edgesym = require_one(edge_sym_tables, "2.8")
        t_duty = require_one(find_tables(doc, "7.6.3.6", ("результат измерения, нс",)), "2.9")
        t_jitter = require_one(find_tables(doc, "7.6.3.7", ("полярность импульса",)), "2.10")

        put(t_amp, 3, 2, v["amp_pos"])
        put(t_amp, 4, 2, v["amp_neg"])
        
        for r in range(4, 14):
            put(t_mask, r, 1, "0")
            put(t_mask, r, 2, "0")
            
        put(t_over, 3, 2, v["overshoot_pos"])
        put(t_over, 4, 2, v["overshoot_neg"])
        put(t_sym, 3, 1, v["amp_sym"])

        r = 3
        for row in v["tx_rl"]:
            for value in row:
                put(rl[0], r, 3, value)
                r += 1

        for row, value in zip([3, 5, 4, 6], v["edges"]):
            put(t_edges, row, 3, value)

        for row, value in zip([3, 4], v["edge_sym"]):
            put(t_edgesym, row, 2, value)
            
        put(t_duty, 3, 1, v["duty"])
        
        for row, value in zip([3, 4], v["jitter"]):
            put(t_jitter, row, 2, value)

        r = 3
        for row in v["rx_rl"]:
            for value in row:
                put(rl[1], r, 3, value)
                r += 1
    except Exception as exc:
        errors.append(str(exc))
    doc.save(output)
    return errors


# ----------------------------------------------------------------------------
# Окно предпросмотра
# ----------------------------------------------------------------------------
FIELDS_100 = [
    ("amp_pos", "2.2 Амплитуда положительная"),
    ("amp_neg", "2.2 Амплитуда отрицательная"),
    ("overshoot_pos", "2.4 Выброс положительный"),
    ("overshoot_neg", "2.4 Выброс отрицательный"),
    ("amp_sym", "2.5 Симметрия амплитуды"),
    ("edges.0", "2.7 Положительный передний фронт"),
    ("edges.1", "2.7 Отрицательный передний фронт"),
    ("edges.2", "2.7 Положительный задний фронт"),
    ("edges.3", "2.7 Отрицательный задний фронт"),
    ("edge_sym.0", "2.8 Симметрия положительного импульса, нс"),
    ("edge_sym.1", "2.8 Симметрия отрицательного импульса, нс"),
    ("duty", "2.9 Искажение коэффициента заполнения"),
    ("jitter.0", "2.10 Джиттер положительный, нс"),
    ("jitter.1", "2.10 Джиттер отрицательный, нс"),
]
FIELDS_10 = [
    ("t10.amp_pos", "1.2 Пиковое напряжение положительное, В"),
    ("t10.amp_neg", "1.2 Пиковое напряжение отрицательное, В"),
    ("t10.harmonic_diff", "1.3 Разность гармоник и частота"),
    ("t10.period", "1.8 Период повторения, мс"),
    ("t10.jitter_without", "1.11 Джиттер без TPM, нс"),
    ("t10.jitter_with", "1.11 Джиттер с TPM, нс"),
    ("t10.common_mode", "1.13 Общее напряжение, мВ"),
]


def get_path(data, path):
    keys = path.split(".")
    value = data
    for item in keys:
        if isinstance(value, list):
            value = value[int(item)]
        else:
            value = value[item]
    return value


def set_path(data, path, value):
    keys = path.split(".")
    target = data
    for item in keys[:-1]:
        if isinstance(target, list):
            target = target[int(item)]
        else:
            target = target[item]

    last_key = keys[-1]
    if isinstance(target, list):
        target[int(last_key)] = value
    else:
        target[last_key] = value


def preview_and_edit(root, values, has_100, has_10, has_1000_a, has_1000_b):
    win = tk.Toplevel(root)
    win.title("Проверка и корректировка значений")
    win.geometry("820x720")
    win.minsize(650, 500)
    win.attributes("-topmost", True)
    win.update_idletasks()
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    ww, wh = min(820, sw - 80), min(720, sh - 100)
    x, y = max(0, (sw - ww) // 2), max(0, (sh - wh) // 2)
    win.geometry("{}x{}+{}+{}".format(ww, wh, x, y))
    win.deiconify()
    win.lift()
    win.focus_force()
    win.after(900, lambda: win.attributes("-topmost", False))

    outer = ttk.Frame(win, padding=10)
    outer.pack(fill="both", expand=True)
    ttk.Label(outer, text="Предварительный просмотр результатов",
              font=("Arial", 13, "bold")).pack(anchor="w", pady=(0, 4))
    ttk.Label(outer, text="Проверьте найденные значения. Любое поле можно изменить до переноса в Word.",
              font=("Arial", 9)).pack(anchor="w", pady=(0, 8))

    canvas = tk.Canvas(outer, highlightthickness=0)
    scrollbar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    inner = ttk.Frame(canvas)
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    inner.columnconfigure(0, weight=1)
    inner.columnconfigure(1, weight=1)
    ttk.Label(inner, text="Показатель", font=("Arial", 9, "bold")).grid(row=0, column=0, sticky="w", padx=5, pady=4)
    ttk.Label(inner, text="Найденное значение (можно изменить)", font=("Arial", 9, "bold")).grid(row=0, column=1, sticky="w", padx=5, pady=4)
    entries = {}

    def add_field(path, label, row):
        ttk.Label(inner, text=label, font=("Arial", 9)).grid(row=row, column=0, sticky="w", padx=5, pady=3)
        ent = ttk.Entry(inner, font=("Arial", 10), width=28)
        ent.insert(0, str(get_path(values, path)))
        ent.grid(row=row, column=1, sticky="ew", padx=5, pady=3)
        entries[path] = ent

    def add_matrix(key, title, row, labels):
        ttk.Separator(inner).grid(row=row, column=0, columnspan=2, sticky="ew", pady=(10, 4))
        row += 1
        ttk.Label(inner, text=title, font=("Arial", 10, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", padx=5, pady=4)
        row += 1

        matrix_values = get_path(values, key)

        for ri, vals in enumerate(matrix_values):
            label = labels[ri] if ri < len(labels) else str(ri + 1)
            ttk.Label(inner, text="Частота {}".format(label), font=("Arial", 9)).grid(row=row, column=0, sticky="w", padx=5, pady=3)
            sub = ttk.Frame(inner)
            sub.grid(row=row, column=1, sticky="ew", padx=5, pady=3)
            for ci, value in enumerate(vals):
                path = "{}.{}.{}".format(key, ri, ci)
                ent = ttk.Entry(sub, width=9, font=("Arial", 10))
                ent.insert(0, str(value))
                ent.pack(side="left", padx=(0, 5))
                entries[path] = ent
            row += 1
        return row

    row = 1

    # --- 10BASE-T ---
    if has_10:
        ttk.Separator(inner).grid(row=row, column=0, columnspan=2, sticky="ew", pady=(6, 4))
        row += 1
        ttk.Label(inner, text="10BASE-T", font=("Arial", 10, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", padx=5, pady=4)
        row += 1
        for path, label in FIELDS_10:
            add_field(path, label, row)
            row += 1

        row = add_matrix("t10.imp_tx",
                         "10BASE-T — импеданс передатчика (5/10 МГц; нагрузки 85/100/111 Ом)",
                         row, ["5 МГц", "10 МГц"])
        row = add_matrix("t10.imp_rx",
                         "10BASE-T — импеданс приёмника (5/10 МГц; нагрузки 85/100/111 Ом)",
                         row, ["5 МГц", "10 МГц"])

    # --- 100BASE-TX ---
    if has_100:
        ttk.Separator(inner).grid(row=row, column=0, columnspan=2, sticky="ew", pady=(10, 4))
        row += 1
        ttk.Label(inner, text="100BASE-TX", font=("Arial", 10, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", padx=5, pady=4)
        row += 1
        for path, label in FIELDS_100:
            add_field(path, label, row)
            row += 1

        freq_labels = [
            "1 МГц", "10 МГц", "20 МГц", "30 МГц", "40 МГц",
            "50 МГц", "60 МГц", "70 МГц", "80 МГц"
        ]

        row = add_matrix("tx_rl", "2.6 Затухание отражения цепей передатчика", row, freq_labels)
        row = add_matrix("rx_rl", "2.11 Затухание отражения цепей приёмника", row, freq_labels)

        ttk.Separator(inner).grid(row=row, column=0, columnspan=2, sticky="ew", pady=(10, 4))
        row += 1
        ttk.Label(inner, text="2.3 Форма сигнала: при PASS число пересечений сегментов = 0",
                  font=("Arial", 9)).grid(row=row, column=0, columnspan=2, sticky="w", padx=5, pady=4)
        row += 1

    # --- 1000BASE-T ---
    if has_1000_a or has_1000_b:
        ttk.Separator(inner).grid(row=row, column=0, columnspan=2, sticky="ew", pady=(10, 4))
        row += 1
        ttk.Label(inner, text="1000BASE-T", font=("Arial", 10, "bold")).grid(row=row, column=0, columnspan=2, sticky="w", padx=5, pady=4)
        row += 1
        
        freq_labels_1000 = [
            "1 МГц", "10 МГц", "20 МГц", "30 МГц", "40 МГц",
            "50 МГц", "60 МГц", "70 МГц", "80 МГц", "90 МГц", "100 МГц"
        ]

        if has_1000_a:
            row = add_matrix("t1000_a", "1000BASE-T — Порт A (нагрузки 85/100/115 Ом)", row, freq_labels_1000)
        if has_1000_b:
            row = add_matrix("t1000_b", "1000BASE-T — Порт B (нагрузки 85/100/115 Ом)", row, freq_labels_1000)

    result = {"ok": False}
    buttons = ttk.Frame(outer)
    buttons.pack(fill="x", pady=(10, 0))

    def accept():
        try:
            for path, ent in entries.items():
                raw = ent.get().strip()
                if not raw:
                    raise ValueError("Пустое значение: {}".format(path))
                set_path(values, path, raw)

            def number(value):
                return Decimal(str(value).replace(",", ".").replace("−", "-")
                             .replace("mV", "").replace("ps", "").replace("ns", "").strip())

            if has_100:
                values["amp_pos"] = round_text(number(values["amp_pos"]), 0)
                values["amp_neg"] = round_text(abs(number(values["amp_neg"])), 0)
                
                for matrix in (values["tx_rl"], values["rx_rl"]):
                    for ri, row_values in enumerate(matrix):
                        for ci, value in enumerate(row_values):
                            matrix[ri][ci] = round_text(abs(number(value)), 1)

                for i, value in enumerate(values["edges"]):
                    values["edges"][i] = round_text(number(value), 2)

                for i, value in enumerate(values["edge_sym"]):
                    values["edge_sym"][i] = round_text(number(value), 3)
                    
                for i, value in enumerate(values["jitter"]):
                    values["jitter"][i] = round_text(number(value), 2)

                duty_match = re.search(r"([+-]?[0-9.,]+)\s*\(\s*±\s*([+-]?[0-9.,]+)\s*\)", str(values["duty"]))
                if duty_match:
                    values["duty"] = "{} (± {})".format(
                        round_text(Decimal(duty_match.group(1).replace(",", ".")), 3),
                        round_text(Decimal(duty_match.group(2).replace(",", ".")), 3))

            if has_10:
                for matrix_key in ("imp_tx", "imp_rx"):
                    for ri, row_values in enumerate(values["t10"][matrix_key]):
                        for ci, value in enumerate(row_values):
                            values["t10"][matrix_key][ri][ci] = round_text(abs(number(value)), 1)
                            
            if has_1000_a:
                for ri, row_values in enumerate(values["t1000_a"]):
                    for ci, value in enumerate(row_values):
                        values["t1000_a"][ri][ci] = round_text(abs(number(value)), 1)
            if has_1000_b:
                for ri, row_values in enumerate(values["t1000_b"]):
                    for ci, value in enumerate(row_values):
                        values["t1000_b"][ri][ci] = round_text(abs(number(value)), 1)

            result["ok"] = True
            win.destroy()
        except Exception as exc:
            messagebox.showerror("Проверка значений", str(exc), parent=win)

    def cancel():
        win.destroy()

    ttk.Button(buttons, text="Перенести в Word", command=accept).pack(side="right", padx=(5, 0))
    ttk.Button(buttons, text="Отмена", command=cancel).pack(side="right")
    win.protocol("WM_DELETE_WINDOW", cancel)
    win.grab_set()
    win.wait_visibility()
    win.lift()
    root.wait_window(win)
    return result["ok"]


# ----------------------------------------------------------------------------
# main
# ----------------------------------------------------------------------------
def main():
    root = tk.Tk()
    root.withdraw()

    template = filedialog.askopenfilename(
        title="Шаг 1: выберите Word-форму протокола",
        filetypes=[("Word", "*.docx")])
    if not template:
        root.destroy()
        return

    pdf_10 = pdf_100 = pdf_1000_a = pdf_1000_b = None

    if messagebox.askyesno("Выбор PDF", "Добавить PDF-отчёт 10BASE-T?", parent=root):
        pdf_10 = filedialog.askopenfilename(title="Выберите PDF-отчёт 10BASE-T", filetypes=[("PDF", "*.pdf")])
    if messagebox.askyesno("Выбор PDF", "Добавить PDF-отчёт 100BASE-TX?", parent=root):
        pdf_100 = filedialog.askopenfilename(title="Выберите PDF-отчёт 100BASE-TX", filetypes=[("PDF", "*.pdf")])
    if messagebox.askyesno("Выбор PDF", "Добавить PDF-отчёт 1000BASE-T Порт A?", parent=root):
        pdf_1000_a = filedialog.askopenfilename(title="Выберите PDF-отчёт 1000BASE-T Порт A", filetypes=[("PDF", "*.pdf")])
    if messagebox.askyesno("Выбор PDF", "Добавить PDF-отчёт 1000BASE-T Порт B?", parent=root):
        pdf_1000_b = filedialog.askopenfilename(title="Выберите PDF-отчёт 1000BASE-T Порт B", filetypes=[("PDF", "*.pdf")])

    if not any([pdf_10, pdf_100, pdf_1000_a, pdf_1000_b]):
        messagebox.showinfo("PDF не выбран", "Выберите хотя бы один PDF-документ.", parent=root)
        root.destroy()
        return

    values = {}
    has_100 = has_10 = has_1000_a = has_1000_b = False
    issues = []

    if pdf_10:
        try:
            values["t10"] = parse_10baset_pdf(pdf_10)
            has_10 = True
        except Exception as exc:
            issues.append("10BASE-T PDF: " + str(exc))
            
    if pdf_100:
        try:
            values.update(parse_pdf(pdf_100))
            has_100 = True
        except Exception as exc:
            issues.append("100BASE-TX PDF: " + str(exc))

    if pdf_1000_a:
        try:
            values["t1000_a"] = parse_1000baset_pdf(pdf_1000_a, "Порт A")
            has_1000_a = True
        except Exception as exc:
            issues.append("1000BASE-T A PDF: " + str(exc))

    if pdf_1000_b:
        try:
            values["t1000_b"] = parse_1000baset_pdf(pdf_1000_b, "Порт B")
            has_1000_b = True
        except Exception as exc:
            issues.append("1000BASE-T B PDF: " + str(exc))

    # Проверка Word-формы
    if has_10:
        try:
            probe = Document(template)
            errors = fill_10baset_docx(probe, values["t10"])
            if errors:
                raise ValueError("; ".join(errors))
        except Exception as exc:
            has_10 = False
            values.pop("t10", None)
            issues.append("10BASE-T Word: " + str(exc))

    if has_100:
        fd, probe_path = tempfile.mkstemp(suffix=".docx")
        os.close(fd)
        try:
            errors = fill_docx(template, probe_path, values)
            if errors:
                raise ValueError("; ".join(errors))
        except Exception as exc:
            has_100 = False
            for key in ("amp_pos", "amp_neg", "overshoot_pos", "overshoot_neg",
                        "amp_sym", "edges", "edge_sym", "jitter", "duty", "tx_rl", "rx_rl"):
                values.pop(key, None)
            issues.append("100BASE-TX Word: " + str(exc))
        finally:
            try:
                os.remove(probe_path)
            except OSError:
                pass

    if has_1000_a or has_1000_b:
        try:
            probe = Document(template)
            errors = fill_1000baset_docx(probe, values.get("t1000_a"), values.get("t1000_b"))
            if errors:
                raise ValueError("; ".join(errors))
        except Exception as exc:
            has_1000_a = False
            has_1000_b = False
            values.pop("t1000_a", None)
            values.pop("t1000_b", None)
            issues.append("1000BASE-T Word: " + str(exc))

    if not any([has_100, has_10, has_1000_a, has_1000_b]):
        messagebox.showerror("Нет данных", "Ни один выбранный PDF не прошёл проверку.\n\n" + "\n\n".join(issues), parent=root)
        root.destroy()
        return
    if issues:
        messagebox.showwarning("Часть данных пропущена", "Ошибочные PDF/разделы Word пропущены.\n\n" + "\n\n".join(issues), parent=root)

    if not preview_and_edit(root, values, has_100, has_10, has_1000_a, has_1000_b):
        root.destroy()
        return

    output = filedialog.asksaveasfilename(title="Сохранить заполненный протокол", defaultextension=".docx",
                                          filetypes=[("Word", "*.docx")], initialfile="Protocol.docx")
    if not output:
        root.destroy()
        return

    temp_names = []
    try:
        doc = Document(template)
        
        if has_100:
            temp100 = output + ".100.tmp.docx"
            temp_names.append(temp100)
            errors = fill_docx(template, temp100, values)
            if errors:
                raise ValueError("100BASE-TX: " + "; ".join(errors))
            doc = Document(temp100)
            
        if has_10:
            temp10 = output + ".10.tmp.docx"
            temp_names.append(temp10)
            doc.save(temp10)
            doc10 = Document(temp10)
            errors = fill_10baset_docx(doc10, values["t10"])
            if errors:
                raise ValueError("10BASE-T: " + "; ".join(errors))
            doc10.save(temp10)
            doc = Document(temp10)

        if has_1000_a or has_1000_b:
            temp1000 = output + ".1000.tmp.docx"
            temp_names.append(temp1000)
            doc.save(temp1000)
            doc1000 = Document(temp1000)
            errors = fill_1000baset_docx(doc1000, values.get("t1000_a"), values.get("t1000_b"))
            if errors:
                raise ValueError("1000BASE-T: " + "; ".join(errors))
            doc1000.save(temp1000)
            doc = Document(temp1000)

        doc.save(output)
        messagebox.showinfo("Готово", "Данные перенесены в Word.\n\nФайл сохранён:\n" + output, parent=root)
    except Exception as exc:
        messagebox.showerror("Ошибка заполнения Word", "Итоговый документ не сохранён.\n\n" + str(exc), parent=root)
    finally:
        for name in temp_names:
            try:
                if os.path.exists(name):
                    os.remove(name)
            except OSError:
                pass
        root.destroy()

if __name__ == "__main__":
    main()