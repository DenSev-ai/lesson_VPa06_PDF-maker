#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_invoice.py

Генератор PDF-чеков из CSV/JSON-данных по HTML-шаблонам с помощью WeasyPrint.
Поддерживает кириллицу, работает на Windows и macOS.
"""

import csv
import json
import os
import platform
import subprocess
import sys
from pathlib import Path
from typing import Any

import pandas as pd

# Временно подавляем stderr при импорте WeasyPrint, чтобы скрыть
# не критичные предупреждения GLib/GTK на Windows (UWP-приложения и т.п.).
_old_stderr = sys.stderr
sys.stderr = open(os.devnull, "w")
try:
    from weasyprint import HTML, CSS
except OSError as e:
    sys.stderr = _old_stderr
    print("\nОшибка: WeasyPrint не может загрузить системные библиотеки.")
    print("Скорее всего, не установлены GTK/Pango.")
    print("\nWindows:")
    print("  1. Скачайте GTK для Windows:")
    print("     https://github.com/tschoonj/GTK-for-Windows-Runtime-Environment-Installer/releases")
    print("  2. Установите и добавьте путь к bin в переменную PATH.")
    print("\nmacOS:")
    print("  brew install pango gdk-pixbuf libffi")
    print("\nПосле установки перезапустите терминал и повторите запуск.\n")
    sys.exit(1)
finally:
    sys.stderr.close()
    sys.stderr = _old_stderr


# === Конфигурация ===
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
TEMPLATES_DIR = BASE_DIR / "templates"
OUTPUT_DIR = BASE_DIR / "output"

# Создаём нужные директории, если их нет
DATA_DIR.mkdir(exist_ok=True)
TEMPLATES_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


# === Вспомогательные функции ===
def clear_screen() -> None:
    """Очистка консоли."""
    os.system("cls" if platform.system() == "Windows" else "clear")


def print_header(title: str) -> None:
    """Красивый заголовок."""
    print("\n" + "=" * 60)
    print(f"  {title}")
    print("=" * 60 + "\n")


def list_files(directory: Path, extensions: tuple[str, ...]) -> list[Path]:
    """Возвращает список файлов с нужными расширениями."""
    files = sorted(
        f for f in directory.iterdir()
        if f.is_file() and f.suffix.lower() in extensions
    )
    return files


def choose_from_menu(items: list[Any], title: str) -> Any:
    """Выводит нумерованное меню и возвращает выбранный элемент."""
    print_header(title)
    if not items:
        print("Нет доступных вариантов. Проверьте директорию.")
        sys.exit(1)

    for idx, item in enumerate(items, start=1):
        label = item.name if isinstance(item, Path) else str(item)
        print(f"  {idx}. {label}")

    print()
    while True:
        try:
            choice = int(input("Введите номер: "))
            if 1 <= choice <= len(items):
                return items[choice - 1]
            print(f"Введите число от 1 до {len(items)}.")
        except ValueError:
            print("Пожалуйста, введите корректное число.")


# === Парсеры данных ===
def load_csv(path: Path) -> list[dict[str, Any]]:
    """Загружает CSV через pandas."""
    df = pd.read_csv(path, dtype=str)
    return df.to_dict(orient="records")


def load_json(path: Path) -> list[dict[str, Any]]:
    """Загружает JSON стандартной библиотекой."""
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    if isinstance(data, dict):
        return [data]
    return data


def detect_invoice_id_field(record: dict[str, Any]) -> str:
    """Пытается найти поле с идентификатором чека."""
    candidates = ["invoice_id", "id", "order_id", "номер", "invoice", "order"]
    for key in record.keys():
        lower_key = key.lower()
        if lower_key in candidates:
            return key
    return list(record.keys())[0]


def add_invoice_ids(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """Убеждается, что у каждой записи есть invoice_id."""
    if not records:
        return records

    field = detect_invoice_id_field(records[0])

    for idx, record in enumerate(records, start=1):
        if "invoice_id" not in record:
            # Если найдено подходящее поле — переименуем его
            record["invoice_id"] = str(record.get(field, idx))
        else:
            record["invoice_id"] = str(record["invoice_id"])
    return records


# === Генерация PDF ===
def render_html(template_path: Path, record: dict[str, Any]) -> str:
    """Подставляет данные в HTML-шаблон."""
    with open(template_path, "r", encoding="utf-8") as f:
        html = f.read()

    # Простая подстановка {{key}} на значения
    for key, value in record.items():
        placeholder = f"{{{{{key}}}}}"
        html = html.replace(placeholder, str(value))

    return html


def build_css() -> CSS:
    """Возвращает CSS со шрифтом, поддерживающим кириллицу."""
    # Пытаемся найти подходящий системный шрифт
    font_family = find_cyrillic_font()

    css_text = f"""
    @page {{
        size: A4;
        margin: 2cm;
    }}
    body {{
        font-family: "{font_family}", "DejaVu Sans", "Roboto", sans-serif;
        font-size: 12pt;
        color: #222;
        line-height: 1.5;
    }}
    h1 {{
        font-size: 20pt;
        text-align: center;
        margin-bottom: 0.5cm;
        color: #1a1a1a;
    }}
    .invoice-box {{
        max-width: 100%;
        border: 1px solid #ccc;
        padding: 1.5cm;
        box-sizing: border-box;
    }}
    .details {{
        margin-top: 1cm;
        width: 100%;
        border-collapse: collapse;
    }}
    .details th, .details td {{
        border: 1px solid #ddd;
        padding: 8px;
        text-align: left;
    }}
    .details th {{
        background-color: #f2f2f2;
        font-weight: bold;
    }}
    .footer {{
        margin-top: 1cm;
        font-size: 10pt;
        color: #555;
        text-align: center;
    }}
    """
    return CSS(string=css_text)


def find_cyrillic_font() -> str:
    """Ищет системный шрифт с поддержкой кириллицы."""
    system = platform.system()

    if system == "Windows":
        candidates = [
            "Segoe UI",
            "Arial",
            "Times New Roman",
            "Calibri",
            "DejaVu Sans",
            "Roboto",
        ]
    elif system == "Darwin":  # macOS
        candidates = [
            "Helvetica Neue",
            "Arial",
            "San Francisco",
            "DejaVu Sans",
            "Roboto",
        ]
    else:
        candidates = ["DejaVu Sans", "Roboto", "Arial"]

    return ", ".join(candidates)


def generate_pdf(template_path: Path, record: dict[str, Any]) -> Path:
    """Генерирует PDF и сохраняет в output/."""
    invoice_id = str(record.get("invoice_id", "unknown")).replace(" ", "_")
    output_path = OUTPUT_DIR / f"invoice_{invoice_id}.pdf"

    html_content = render_html(template_path, record)
    html_doc = HTML(string=html_content, base_url=str(BASE_DIR))
    css = build_css()

    html_doc.write_pdf(str(output_path), stylesheets=[css])
    return output_path


def open_pdf(path: Path) -> None:
    """Открывает PDF в стандартной программе системы."""
    system = platform.system()
    try:
        if system == "Windows":
            os.startfile(str(path))
        elif system == "Darwin":
            subprocess.run(["open", str(path)], check=True)
        else:
            subprocess.run(["xdg-open", str(path)], check=True)
    except Exception as e:
        print(f"Не удалось автоматически открыть PDF: {e}")


# === Главный процесс ===
def main() -> None:
    clear_screen()
    print_header("Генератор PDF-чеков")

    # 1. Список файлов данных
    data_files = list_files(DATA_DIR, (".csv", ".json"))
    print("Доступные файлы с данными:")
    for idx, file in enumerate(data_files, start=1):
        print(f"  {idx}. {file.name}")

    if not data_files:
        print("\nВ директории data/ нет CSV или JSON файлов.")
        sys.exit(1)

    # 2. Список шаблонов
    template_files = list_files(TEMPLATES_DIR, (".html",))
    print("\nДоступные HTML-шаблоны:")
    for idx, file in enumerate(template_files, start=1):
        print(f"  {idx}. {file.name}")

    if not template_files:
        print("\nВ директории templates/ нет HTML-шаблонов.")
        sys.exit(1)

    # 3. Выбор файла данных
    data_file = choose_from_menu(data_files, "Выберите файл данных")

    # 4. Выбор шаблона
    template_file = choose_from_menu(template_files, "Выберите HTML-шаблон")

    # 5. Загрузка данных
    if data_file.suffix.lower() == ".csv":
        records = load_csv(data_file)
    else:
        records = load_json(data_file)

    records = add_invoice_ids(records)

    # 6. Выбор invoice_id
    invoice_ids = [r["invoice_id"] for r in records]
    chosen_id = choose_from_menu(invoice_ids, "Выберите invoice id")

    selected_record = next(r for r in records if r["invoice_id"] == chosen_id)

    # 7. Генерация PDF
    print(f"\nГенерация PDF для invoice_id = {chosen_id}...")
    output_path = generate_pdf(template_file, selected_record)

    print(f"PDF сохранён: {output_path}")

    # 8. Открытие PDF
    open_pdf(output_path)
    print("PDF открыт в системной программе.")


if __name__ == "__main__":
    main()
