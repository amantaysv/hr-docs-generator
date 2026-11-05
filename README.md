# Генератор документов

Excel (ФИО, паспорт…) + Word-шаблон с метками `{{Название колонки}}` → отдельный .docx на каждого человека.

    python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
    .venv/bin/python app.py      # открыть http://127.0.0.1:5050

Готовые файлы по умолчанию в `output/` (имя файла = значение выбранной колонки, обычно ФИО).
Совет: колонку с паспортом в Excel делайте текстовой, иначе Excel съест ведущие нули.

## Версия для GitHub Pages (папка `site/`)
Статическая страница: Excel и Word обрабатываются в браузере, ничего не отправляется на сервер.
Деплой — через GitHub Actions (`.github/workflows/pages.yml`) при пуше в `main`.
Локально: `cd site && python3 -m http.server 8000`.
