"""Подстановка данных из Excel в Word-шаблон ({{Колонка}})."""
import re
from datetime import date, datetime
from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from openpyxl import load_workbook

PLACEHOLDER = re.compile(r"\{\{\s*(.+?)\s*\}\}")


# ---------- Excel ----------
def _to_text(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, date):
        return v.strftime("%d.%m.%Y")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def read_excel(path):
    """Возвращает (заголовки, список словарей). Первая строка — заголовки."""
    wb = load_workbook(path, data_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    try:
        header = [_to_text(h) for h in next(rows)]
    except StopIteration:
        return [], []
    records = []
    for r in rows:
        if all(c is None or str(c).strip() == "" for c in r):
            continue
        records.append({h: _to_text(c) for h, c in zip(header, r) if h})
    return [h for h in header if h], records


# ---------- Word ----------
def _iter_paragraphs(container):
    """Все параграфы контейнера, включая таблицы (рекурсивно)."""
    for p in container.paragraphs:
        yield p
    for t in container.tables:
        yield from _iter_table(t)


def _iter_table(table: Table):
    for row in table.rows:
        seen = set()
        for cell in row.cells:
            if id(cell._tc) in seen:  # объединённые ячейки
                continue
            seen.add(id(cell._tc))
            yield from _iter_paragraphs(cell)


def _all_paragraphs(doc):
    yield from _iter_paragraphs(doc)
    for s in doc.sections:
        for part in (s.header, s.footer, s.first_page_header,
                     s.first_page_footer, s.even_page_header, s.even_page_footer):
            if part is not None and not part.is_linked_to_previous:
                yield from _iter_paragraphs(part)


def find_placeholders(template_path):
    doc = Document(template_path)
    found = []
    for p in _all_paragraphs(doc):
        for m in PLACEHOLDER.finditer(p.text):
            if m.group(1) not in found:
                found.append(m.group(1))
    return found


def _replace_in_paragraph(p: Paragraph, data, missing):
    runs = p.runs
    full = "".join(r.text for r in runs)
    if "{{" not in full:
        return
    matches = list(PLACEHOLDER.finditer(full))
    # идём с конца, чтобы смещения не сбивались
    for m in reversed(matches):
        key = m.group(1)
        if key not in data:
            missing.add(key)
            continue
        start, end = m.span()
        pos = 0
        first = True
        for r in runs:
            r_start, r_end = pos, pos + len(r.text)
            pos = r_end
            if r_end <= start or r_start >= end:
                continue
            s = max(start, r_start) - r_start
            e = min(end, r_end) - r_start
            if first:
                r.text = r.text[:s] + data[key] + r.text[e:]
                first = False
            else:
                r.text = r.text[:s] + r.text[e:]


def fill_template(template_path, data, out_path):
    """Подставляет data в шаблон; возвращает множество меток без колонки."""
    doc = Document(template_path)
    missing = set()
    for p in _all_paragraphs(doc):
        _replace_in_paragraph(p, data, missing)
    doc.save(out_path)
    return missing


# ---------- Генерация ----------
def safe_name(s):
    s = re.sub(r'[\\/:*?"<>|\r\n\t]+', " ", s).strip(" .")
    return re.sub(r"\s+", " ", s)[:120] or "документ"


def generate_all(xlsx_path, template_path, out_dir, name_column):
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    header, records = read_excel(xlsx_path)
    if name_column not in header:
        raise ValueError(f"В таблице нет колонки «{name_column}»")
    used = {}
    files, warnings = [], []
    for i, rec in enumerate(records, start=2):  # номер строки в Excel
        base = safe_name(rec.get(name_column, "")) if rec.get(name_column) else f"строка {i}"
        n = used.get(base.lower(), 0) + 1
        used[base.lower()] = n
        fname = f"{base}.docx" if n == 1 else f"{base}_{n}.docx"
        missing = fill_template(template_path, rec, out_dir / fname)
        files.append(fname)
        empty = [k for k, v in rec.items() if v == ""]
        if empty:
            warnings.append(f"Строка {i} ({fname}): пустые ячейки — {', '.join(empty)}")
        if missing:
            warnings.append(f"{fname}: нет колонки для меток — {', '.join(sorted(missing))}")
    return files, warnings
