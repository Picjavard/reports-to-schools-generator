import os
import sys
from pathlib import Path
import pandas as pd
from pypdf import PdfWriter, PdfReader
import win32com.client as win32
from tkinter import Tk, filedialog
import re

# === КОНФИГУРАЦИЯ ===
REFERENCE_XLSX = "C:/Users/alexd/OneDrive/Документы/Тестовая папка с ведомостями/Список ведомостей.xlsx"
# Сопоставление школы → код
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


def get_teacher_folder_name(teacher_fio):
    """Из 'Демьяненко А.Е.' → 'Демьяненко'"""
    if not isinstance(teacher_fio, str) or not teacher_fio.strip():
        return ""
    return teacher_fio.split()[0]


def update_app_number_and_set_print_area(ws, app_num_from_ref):
    """Находит ячейку с 'ПРИЛОЖЕНИЕ' в первой строке, обновляет номер и устанавливает область печати"""
    try:
        app_column = None
        app_cell = None

        # Ищем ячейку с "ПРИЛОЖЕНИЕ" в первой строке
        for col in range(1, 50):  # проверяем первые 50 столбцов
            cell_value = ws.Cells(1, col).Value
            if cell_value and isinstance(cell_value, str) and "ПРИЛОЖЕНИЕ" in cell_value:
                app_column = col
                app_cell = ws.Cells(1, col)
                break

        if not app_cell:
            print(f"  ⚠️ Не найдена ячейка с 'ПРИЛОЖЕНИЕ' в первой строке")
            # Попробуем найти в других строках
            for row in range(1, 5):
                for col in range(1, 50):
                    cell_value = ws.Cells(row, col).Value
                    if cell_value and isinstance(cell_value, str) and "ПРИЛОЖЕНИЕ" in cell_value:
                        app_column = col
                        app_cell = ws.Cells(row, col)
                        break
                if app_cell:
                    break

        if not app_cell:
            print(f"  ⚠️ Не найдена ячейка с 'ПРИЛОЖЕНИЕ' во всем заголовке")
            return False

        # Обновляем номер приложения в ячейке на основе справочника
        if app_num_from_ref:
            # Обрабатываем случаи, когда номер приложения может быть не числом (например, "1-2")
            if isinstance(app_num_from_ref, str) and '-' in app_num_from_ref:
                new_text = f"ПРИЛОЖЕНИЕ №{app_num_from_ref}"
            else:
                new_text = f"ПРИЛОЖЕНИЕ №{app_num_from_ref}"

            app_cell.Value = new_text
            print(f"  ✏️ Обновлен номер приложения: {new_text}")

        # Устанавливаем область печати - от A1 до столбца с ПРИЛОЖЕНИЕ
        if app_column:
            # Находим последнюю строку с данными (ищем в столбце A)
            last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # -4162 = xlUp

            # Устанавливаем область печати до столбца с ПРИЛОЖЕНИЕ
            end_col = app_column
            print_area_range = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, end_col))
            ws.PageSetup.PrintArea = print_area_range.Address

            print(f"  📏 Область печати установлена: A1:{chr(64 + end_col)}{last_row}")
            return True
        else:
            print("  ⚠️ Не удалось определить столбец для области печати")
            return False

    except Exception as e:
        print(f"  ⚠️ Ошибка при обновлении номера приложения и установке области печати: {e}")
        return False


def main():
    # --- Выбор папок ---
    root = Tk()
    root.withdraw()
    input_folder = filedialog.askdirectory(
        title="Выберите папку с папками преподавателей",
        initialdir=r"C:/Users/alexd/OneDrive/Документы/Тестовая папка с ведомостями"
    )
    if not input_folder:
        print("❌ Не выбрана входная папка")
        return

    output_folder = filedialog.askdirectory(
        title="Выберите папку для сохранения итоговых PDF",
        initialdir=input_folder
    )
    if not output_folder:
        print("❌ Не выбрана выходная папка")
        return

    os.makedirs(output_folder, exist_ok=True)

    # --- Загрузка справочника ---
    try:
        df = pd.read_excel(REFERENCE_XLSX, sheet_name="Лист1", dtype=str)
    except Exception as e:
        print(f"❌ Ошибка при чтении справочника:\n{e}")
        return

    # Оставляем только строки с названием листа
    df = df.dropna(subset=["Название листа"])
    print(f"✅ Загружено {len(df)} ведомостей для обработки.")

    # Словарь: фамилия_преподавателя → {название_листа → (номер_приложения, код_школы)}
    sheet_info_by_teacher = {}

    for _, row in df.iterrows():
        sheet_name = row["Название листа"].strip()
        app_num = row["Номер Приложения"].strip() if pd.notna(row["Номер Приложения"]) else ""

        school_full = row["Школа"]
        school_code = SCHOOL_TO_CODE.get(school_full.strip())
        if not school_code:
            print(f"⚠️ Неизвестная школа для листа '{sheet_name}': {school_full}")
            continue

        teacher_fio = row["ФИО Преподавателя"]
        teacher_folder = get_teacher_folder_name(teacher_fio)

        if teacher_folder not in sheet_info_by_teacher:
            sheet_info_by_teacher[teacher_folder] = {}

        # Сохраняем информацию о листе для этого преподавателя
        sheet_info_by_teacher[teacher_folder][sheet_name] = (app_num, school_code)

    # --- Поиск Excel-файлов ---
    all_excel_files = []
    found_teachers = set()

    for item in Path(input_folder).iterdir():
        if item.is_dir():
            teacher_name = item.name
            found_teachers.add(teacher_name)
            for f in item.glob("*.xlsx"):
                all_excel_files.append((f, teacher_name))

    if not all_excel_files:
        print("❌ Не найдено ни одного Excel-файла в подпапках.")
        return

    print(f"🔍 Найдено {len(all_excel_files)} Excel-файлов у {len(found_teachers)} преподавателей.")

    # --- Экспорт нужных листов в PDF ---
    excel_app = win32.gencache.EnsureDispatch('Excel.Application')
    excel_app.Visible = False
    excel_app.DisplayAlerts = False

    found_sheets = []  # (school_code, app_num, pdf_path, teacher_name, sheet_name)

    for excel_path, teacher_name in all_excel_files:
        try:
            # Проверяем, есть ли информация о листах для этого преподавателя
            if teacher_name not in sheet_info_by_teacher:
                print(f"  ⚠️ Нет данных о листах для преподавателя: {teacher_name}")
                continue

            wb = excel_app.Workbooks.Open(str(excel_path))
            sheet_names_in_file = [s.Name for s in wb.Sheets]

            for sheet_name in sheet_names_in_file:
                if sheet_name in ("Служебное", "Списки классов"):
                    continue

                # Проверяем, есть ли такой лист у этого преподавателя в справочнике
                if sheet_name not in sheet_info_by_teacher[teacher_name]:
                    continue

                app_num, school_code = sheet_info_by_teacher[teacher_name][sheet_name]
                pdf_path = os.path.join(output_folder, f"{teacher_name}_{sheet_name}.pdf")

                ws = wb.Worksheets(sheet_name)
                # Обновляем номер приложения и устанавливаем область печати
                update_app_number_and_set_print_area(ws, app_num)

                # Экспортируем только область печати
                ws.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
                found_sheets.append((school_code, app_num, pdf_path, teacher_name, sheet_name))
                print(f"✅ Экспортирован: {teacher_name}/{sheet_name} (Приложение №{app_num})")

            wb.Close(SaveChanges=True)  # Сохраняем изменения (обновленные номера приложений)
        except Exception as e:
            print(f"⚠️ Ошибка при обработке {excel_path} ({teacher_name}): {e}")

    excel_app.Quit()

    if not found_sheets:
        print("❌ Ни один лист из справочника не найден в Excel-файлах.")
        return

    print(f"✅ Найдено {len(found_sheets)} ведомостей для включения в итоговые PDF.")

    # --- Проверка наличия титульных PDF ---
    title_pdfs = {}
    for code in ["13", "16", "17", "22", "АГ"]:
        pdf_path = os.path.join(output_folder, f"title_{code}.pdf")
        if os.path.exists(pdf_path):
            title_pdfs[code] = pdf_path
        else:
            print(f"❌ Отсутствует титульный PDF: {pdf_path}")
            print("❗ Пожалуйста, поместите подписанные титулы в папку вывода с именами:")
            print("    title_13.pdf, title_16.pdf, title_17.pdf, title_22.pdf, title_АГ.pdf")
            return

    # --- Сборка итоговых PDF ---
    from collections import defaultdict
    groups = defaultdict(list)
    for school, app_num, pdf_path, teacher_name, sheet_name in found_sheets:
        groups[school].append((app_num, pdf_path, teacher_name, sheet_name))

    # Функция для преобразования номера приложения в числовой формат для сортировки
    def app_num_to_sort_key(app_num_str):
        if not app_num_str:
            return (999, 0)  # Последнее место для пустых значений

        # Если номер приложения - диапазон (например, "1-2")
        if '-' in app_num_str:
            parts = app_num_str.split('-')
            try:
                first_num = float(parts[0].strip())
                return (first_num, 0)
            except ValueError:
                return (999, 0)

        # Пробуем преобразовать в число
        try:
            return (float(app_num_str.strip()), 0)
        except ValueError:
            # Если не число, разбиваем на текстовую и числовую части
            match = re.match(r'(\D*)(\d+)', app_num_str.strip())
            if match:
                text_part = match.group(1)
                num_part = int(match.group(2))
                return (num_part, text_part)
            return (999, app_num_str)

    for school_code, items in groups.items():
        # Сортируем по номеру приложения, а при одинаковых номерах - по имени преподавателя
        items.sort(key=lambda x: (app_num_to_sort_key(x[0]), x[2]))

        writer = PdfWriter()

        # Титульный лист
        writer.append(PdfReader(title_pdfs[school_code]))

        # Ведомости
        for app_num, pdf_path, teacher_name, sheet_name in items:
            if os.path.exists(pdf_path):
                writer.append(PdfReader(pdf_path))
                print(f"  ➕ Добавлено: Приложение №{app_num} ({teacher_name}/{sheet_name})")

        output_path = os.path.join(output_folder, f"Ведомости_{SCHOOL_NAMES[school_code]}.pdf")
        with open(output_path, "wb") as f:
            writer.write(f)
        print(f"📄 Создан итоговый PDF: {output_path} ({len(items)} приложений)")

    print("\n🎉 Готово! Все PDF-файлы сформированы.")


if __name__ == "__main__":
    main()