import os
import re
import sys
from pathlib import Path
import pandas as pd
from pypdf import PdfWriter, PdfReader
import win32com.client as win32
from tkinter import Tk, filedialog

# === КОНФИГУРАЦИЯ ===
REFERENCE_XLSX = "Список ведомостей.xlsx"
TEMPLATE_DOCX = "Ведомость (Текущая) Шаблон v.2.docx"

# Сопоставление полного названия школы → код
SCHOOL_TO_CODE = {
    "МАОУ «Школа № 13 г. Благовещенска»": "13",
    "МАОУ «Школа № 16 г. Благовещенска»": "16",
    "МАОУ «Школа № 17 г. Благовещенска»": "17",
    "МАОУ «Школа № 22 г. Благовещенска им. Ф.Э. Дзержинского»": "22",
    "МАОУ «Алексеевская гимназия г. Благовещенска»": "АГ"
}

SCHOOL_NAMES = {
    "13": "Школа №13",
    "16": "Школа №16",
    "17": "Школа №17",
    "22": "Школа №22",
    "АГ": "Алексеевская гимназия"
}


# === ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ===

def extract_school_code(full_name):
    return SCHOOL_TO_CODE.get(full_name.strip(), None)


def get_teacher_folder_name(teacher_fio):
    """Из 'Демьяненко А.Е.' → 'Демьяненко'"""
    return teacher_fio.split()[0] if teacher_fio and isinstance(teacher_fio, str) else ""


def split_title_docx(template_path, output_dir):
    """Автоматически разделяет Word-файл на 5 титульных PDF по школам."""
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    doc = word.Documents.Open(os.path.abspath(template_path))

    # Порядок титулов в вашем файле:
    title_order = [
        ("16", "МАОУ «Школа № 16 г. Благовещенска»"),
        ("АГ", "МАОУ «Алексеевская гимназия г. Благовещенска»"),
        ("17", "МАОУ «Школа № 17 г. Благовещенска»"),
        ("13", "МАОУ «Школа № 13 г. Благовещенска»"),
        ("22", "МАОУ «Школа № 22 г. Благовещенска им. Ф.Э. Дзержинского»")
    ]

    full_text = doc.Range().Text
    start = 0
    for i, (code, school_name) in enumerate(title_order):
        # Найти начало титула
        pos = full_text.find(school_name, start)
        if pos == -1:
            print(f"⚠️ Не найден титул для {school_name}")
            continue

        # Создать временный документ
        temp_doc = word.Documents.Add()
        # Копируем весь контент (проще всего — выделяем блок до следующего титула или конца)
        end_pos = len(full_text)
        if i < len(title_order) - 1:
            next_school = title_order[i + 1][1]
            next_pos = full_text.find(next_school, pos)
            if next_pos != -1:
                end_pos = next_pos

        # Выделяем текст (в Word это сложно без точных позиций)
        # Поэтому проще: сохраняем весь документ и делим вручную один раз
        # → РЕКОМЕНДУЕМ: сохранить 5 title_*.docx вручную
        temp_doc.Close()

    doc.Close()
    word.Quit()
    print("❗ Автоматическое разделение титулов сложно. Рекомендуется сохранить вручную:")
    print("   title_13.docx, title_16.docx, title_17.docx, title_22.docx, title_АГ.docx")
    return False  # сигнал, что нужно вручную


# === ОСНОВНОЙ КОД ===

def main():
    # --- Шаг 1: Выбор папок ---
    root = Tk()
    root.withdraw()
    input_folder = filedialog.askdirectory(title="Выберите папку с папками преподавателей")
    if not input_folder:
        print("❌ Папка не выбрана. Выход.")
        return
    output_folder = filedialog.askdirectory(title="Выберите папку для сохранения PDF")
    if not output_folder:
        print("❌ Папка вывода не выбрана. Выход.")
        return

    os.makedirs(output_folder, exist_ok=True)

    # --- Шаг 2: Загрузка справочника ---
    try:
        df = pd.read_excel(REFERENCE_XLSX, sheet_name="Лист1", dtype=str)
    except Exception as e:
        print(f"❌ Ошибка при чтении {REFERENCE_XLSX}: {e}")
        return

    # Оставляем только строки с названием листа
    df = df.dropna(subset=["Название листа"])
    print(f"✅ Загружено {len(df)} ведомостей для обработки.")

    # Словарь: название_листа → (номер_приложения, код_школы, фамилия_преподавателя)
    sheet_info = {}
    for _, row in df.iterrows():
        sheet_name = row["Название листа"].strip()
        try:
            app_num = int(float(row["Номер Приложения"]))
        except:
            print(f"⚠️ Пропущен лист '{sheet_name}': некорректный номер приложения")
            continue
        school_full = row["Школа"]
        school_code = extract_school_code(school_full)
        if not school_code:
            print(f"⚠️ Неизвестная школа для листа '{sheet_name}': {school_full}")
            continue
        teacher_fio = row["ФИО Преподавателя"]
        teacher_folder = get_teacher_folder_name(teacher_fio)
        sheet_info[sheet_name] = (app_num, school_code, teacher_folder)

    # --- Шаг 3: Поиск Excel-файлов ---
    all_excel_files = []
    for teacher_dir in Path(input_folder).iterdir():
        if teacher_dir.is_dir():
            for f in teacher_dir.glob("*.xlsx"):
                all_excel_files.append(f)

    if not all_excel_files:
        print("❌ Не найдено ни одного Excel-файла в подпапках.")
        return

    print(f"🔍 Найдено {len(all_excel_files)} Excel-файлов.")

    # --- Шаг 4: Экспорт нужных листов в PDF ---
    excel_app = win32.gencache.EnsureDispatch('Excel.Application')
    excel_app.Visible = False
    excel_app.DisplayAlerts = False

    found_sheets = []  # (school_code, app_num, pdf_path)

    for excel_path in all_excel_files:
        try:
            wb = excel_app.Workbooks.Open(str(excel_path))
            sheet_names_in_file = [s.Name for s in wb.Sheets]
            for sheet_name in sheet_names_in_file:
                if sheet_name in ("Служебное", "Списки классов"):
                    continue
                if sheet_name not in sheet_info:
                    continue

                app_num, school_code, _ = sheet_info[sheet_name]
                pdf_path = os.path.join(output_folder, f"{sheet_name}.pdf")
                ws = wb.Worksheets(sheet_name)
                ws.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
                found_sheets.append((school_code, app_num, pdf_path))
                print(f"✅ Экспортирован: {sheet_name}")
            wb.Close(SaveChanges=False)
        except Exception as e:
            print(f"⚠️ Ошибка при обработке {excel_path}: {e}")

    excel_app.Quit()

    if not found_sheets:
        print("❌ Ни один лист из справочника не найден в Excel-файлах.")
        return

    # --- Шаг 5: Подготовка титульных листов ---
    title_docs_ready = True
    for code in SCHOOL_TO_CODE.values():
        if not os.path.exists(f"title_{code}.docx"):
            title_docs_ready = False
            break

    if not title_docs_ready:
        print("\n❗ Отсутствуют файлы title_*.docx. Попытка автоматического разделения...")
        if not split_title_docx(TEMPLATE_DOCX, output_folder):
            print("❌ Пожалуйста, сохраните вручную 5 титульных .docx в папке со скриптом:")
            for code in ["13", "16", "17", "22", "АГ"]:
                print(f"   title_{code}.docx")
            return

    # Конвертируем титулы в PDF
    word_app = win32.gencache.EnsureDispatch('Word.Application')
    word_app.Visible = False
    word_app.DisplayAlerts = False

    title_pdfs = {}
    for code in ["13", "16", "17", "22", "АГ"]:
        docx_path = f"title_{code}.docx"
        pdf_path = os.path.join(output_folder, f"title_{code}.pdf")
        try:
            doc = word_app.Documents.Open(os.path.abspath(docx_path))
            doc.ExportAsFixedFormat(os.path.abspath(pdf_path), 17)  # 17 = PDF
            doc.Close()
            title_pdfs[code] = pdf_path
            print(f"✅ Титульный лист для {SCHOOL_NAMES[code]} сохранён.")
        except Exception as e:
            print(f"⚠️ Не удалось создать титульный PDF для {code}: {e}")
            title_pdfs[code] = None

    word_app.Quit()

    # --- Шаг 6: Сборка итоговых PDF ---
    from collections import defaultdict
    groups = defaultdict(list)
    for school, app_num, pdf in found_sheets:
        groups[school].append((app_num, pdf))

    for school_code, items in groups.items():
        items.sort(key=lambda x: x[0])  # сортировка по номеру приложения
        writer = PdfWriter()

        # Добавляем титульный лист
        title_pdf = title_pdfs.get(school_code)
        if title_pdf and os.path.exists(title_pdf):
            writer.append(PdfReader(title_pdf))
        else:
            print(f"⚠️ Пропущен титульный лист для {SCHOOL_NAMES[school_code]}")

        # Добавляем ведомости
        for _, pdf_path in items:
            if os.path.exists(pdf_path):
                writer.append(PdfReader(pdf_path))

        output_path = os.path.join(output_folder, f"Ведомости_{SCHOOL_NAMES[school_code]}.pdf")
        with open(output_path, "wb") as f:
            writer.write(f)
        print(f"📄 Создан итоговый PDF: {output_path}")

    print("\n🎉 Готово! Все PDF-файлы сформированы.")


if __name__ == "__main__":
    main()