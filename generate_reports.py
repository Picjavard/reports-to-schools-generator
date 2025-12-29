import os
import sys
from pathlib import Path
import pandas as pd
from pypdf import PdfWriter, PdfReader
import win32com.client as win32
from tkinter import Tk, filedialog
import re
import pythoncom
import time
import logging
from datetime import datetime

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


def setup_logging(output_folder):
    """Настройка логирования"""
    log_file = os.path.join(output_folder, f"processing_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    return log_file


def get_teacher_folder_name(teacher_fio):
    """Извлекает только фамилию из ФИО (первое слово до пробела)"""
    if not isinstance(teacher_fio, str) or not teacher_fio.strip():
        return ""
    return teacher_fio.split()[0].strip()


def get_excel_cell_value(ws, row, col):
    """Безопасное получение значения ячейки Excel"""
    try:
        return ws.Cells(row, col).Value
    except:
        return None


def update_app_number_and_set_print_area(ws, app_num_from_ref):
    """Находит ячейку с 'ПРИЛОЖЕНИЕ' в первой строке, обновляет номер и устанавливает область печати"""
    try:
        app_column = None
        app_cell = None

        # Ищем ячейку с "ПРИЛОЖЕНИЕ" в первых 5 строках и 50 столбцах
        for row in range(1, 6):  # Проверяем первые 5 строк
            for col in range(1, 51):  # Проверяем первые 50 столбцов
                cell_value = get_excel_cell_value(ws, row, col)
                if cell_value and isinstance(cell_value, str) and "ПРИЛОЖЕНИЕ" in cell_value:
                    app_column = col
                    app_cell = ws.Cells(row, col)
                    logging.info(f"  📌 Найдена ячейка с 'ПРИЛОЖЕНИЕ' в строке {row}, столбце {col}: {cell_value}")
                    break
            if app_cell:
                break

        if not app_cell:
            logging.warning(f"  ⚠️ Не найдена ячейка с 'ПРИЛОЖЕНИЕ' в заголовке")
            return False

        # Обновление номера приложения
        if app_num_from_ref:
            if isinstance(app_num_from_ref, str) and '-' in app_num_from_ref:
                new_text = f"ПРИЛОЖЕНИЕ №{app_num_from_ref}"
            else:
                new_text = f"ПРИЛОЖЕНИЕ №{app_num_from_ref}"

            app_cell.Value = new_text
            logging.info(f"  ✏️ Обновлен номер приложения: {new_text}")

        # --- ИСПРАВЛЕНО: ЭФФЕКТИВНЫЙ ПОИСК ПОСЛЕДНЕЙ СТРОКИ ---
        # Вместо перебора всех строк используем метод Excel End(xlUp)
        try:
            last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # -4162 = xlUp
            # Ограничиваем максимальное количество строк (на случай, если в столбце A нет данных)
            if last_row > 200:  # Ограничение на 200 строк (достаточно для 2 страниц)
                last_row = 200
            logging.info(f"  📏 Последняя строка с данными: {last_row}")
        except Exception as e:
            logging.warning(f"  ⚠️ Ошибка при определении последней строки, используем значение по умолчанию: {e}")
            last_row = 200  # Значение по умолчанию (2 страницы)

        # Установка области печати
        if app_column:
            end_col = app_column
            # Ограничиваем количество столбцов (максимум 30 столбцов - достаточно для 2 страниц вправо)
            if end_col > 30:
                end_col = 30

            print_area_range = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, end_col))
            ws.PageSetup.PrintArea = print_area_range.Address

            logging.info(f"  📐 Область печати установлена: A1:{chr(64 + end_col)}{last_row}")
            return True
        else:
            logging.warning("  ⚠️ Не удалось определить столбец для области печати")
            return False

    except Exception as e:
        logging.error(f"  ⚠️ Ошибка при обновлении номера приложения и установке области печати: {e}")
        return False


def initialize_excel():
    """Пытаемся инициализировать Excel несколькими способами"""
    try:
        logging.info("🔄 Попытка подключения к Excel...")
        excel_app = win32.Dispatch('Excel.Application')
        excel_app.Visible = False
        excel_app.DisplayAlerts = False
        logging.info("✅ Excel успешно подключен (метод Dispatch)")
        return excel_app
    except Exception as e1:
        logging.warning(f"⚠️ Первая попытка не удалась: {e1}")
        try:
            logging.info("🔄 Попытка пересоздания кэша...")
            pythoncom.CoInitialize()
            excel_app = win32.gencache.EnsureDispatch('Excel.Application')
            excel_app.Visible = False
            excel_app.DisplayAlerts = False
            logging.info("✅ Excel успешно подключен (метод gencache)")
            return excel_app
        except Exception as e2:
            logging.error(f"❌ Все попытки подключения к Excel не удалась: {e2}")
            logging.error("🔧 Рекомендуемые действия:")
            logging.error("1. Удалите папку C:\\Users\\alexd\\AppData\\Local\\Temp\\gen_py")
            logging.error("2. Перезапустите компьютер")
            logging.error("3. Убедитесь, что Excel установлен и работает")
            return None


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

    # Настройка логирования
    log_file = setup_logging(output_folder)
    logging.info("=== НАЧАЛО ОБРАБОТКИ ===")
    logging.info(f"Входная папка: {input_folder}")
    logging.info(f"Выходная папка: {output_folder}")
    logging.info(f"Файл лога: {log_file}")

    os.makedirs(output_folder, exist_ok=True)

    # --- Загрузка справочника ---
    try:
        df = pd.read_excel(REFERENCE_XLSX, sheet_name="Лист1", dtype=str)
        logging.info(f"✅ Загружен справочник из файла: {REFERENCE_XLSX}")
    except Exception as e:
        logging.error(f"❌ Ошибка при чтении справочника:\n{e}")
        if "FileNotFoundError" in str(e):
            logging.error(f"💡 Проверьте путь к файлу: {REFERENCE_XLSX}")
        return

    # Оставляем только строки с названием листа
    df = df.dropna(subset=["Название листа"])
    logging.info(f"✅ Загружено {len(df)} ведомостей для обработки.")

    # Словарь: фамилия_преподавателя → {название_листа → (номер_приложения, код_школы)}
    sheet_info_by_teacher = {}
    teacher_stats = {}

    logging.info("\n📊 Анализ данных из справочника:")
    for _, row in df.iterrows():
        sheet_name = row["Название листа"].strip()
        app_num = row["Номер Приложения"].strip() if pd.notna(row["Номер Приложения"]) else ""

        school_full = row["Школа"]
        school_code = SCHOOL_TO_CODE.get(school_full.strip())
        if not school_code:
            logging.warning(f"⚠️ Неизвестная школа для листа '{sheet_name}': {school_full}")
            continue

        teacher_fio = row["ФИО Преподавателя"]
        teacher_surname = get_teacher_folder_name(teacher_fio)
        logging.debug(f"  👨‍🏫 Обработка записи: {teacher_fio} → фамилия: '{teacher_surname}'")

        if teacher_surname not in teacher_stats:
            teacher_stats[teacher_surname] = {
                "original_name": teacher_fio,
                "count": 0
            }
        teacher_stats[teacher_surname]["count"] += 1

        if teacher_surname not in sheet_info_by_teacher:
            sheet_info_by_teacher[teacher_surname] = {}

        sheet_info_by_teacher[teacher_surname][sheet_name] = (app_num, school_code)
        logging.debug(f"    📄 Добавлен лист '{sheet_name}' для {teacher_surname} (Приложение №{app_num})")

    # Вывод статистики по преподавателям
    logging.info("\n👨‍🏫 Статистика по преподавателям в справочнике:")
    for surname, stats in teacher_stats.items():
        logging.info(f"  👨‍🏫 {surname} (из: {stats['original_name']}): {stats['count']} листов")

    # --- Поиск Excel-файлов ---
    all_excel_files = []
    found_teachers = set()

    logging.info("\n📂 Поиск Excel-файлов в подпапках:")
    for item in Path(input_folder).iterdir():
        if item.is_dir():
            folder_name = item.name
            found_teachers.add(folder_name)
            logging.info(f"  📂 Найдена папка: {folder_name}")

            # Извлекаем фамилию из названия папки
            folder_surname = get_teacher_folder_name(folder_name)
            logging.info(f"    🔍 Фамилия из названия папки: '{folder_surname}'")

            for f in item.glob("*.xlsx"):
                all_excel_files.append((f, folder_name, folder_surname))
                logging.info(f"    📄 Найден Excel-файл: {f.name}")

    if not all_excel_files:
        logging.error("❌ Не найдено ни одного Excel-файла в подпапках.")
        return

    logging.info(f"\n🔍 Найдено {len(all_excel_files)} Excel-файлов у {len(found_teachers)} преподавателей.")

    # --- Инициализация Excel ---
    excel_app = initialize_excel()
    if not excel_app:
        logging.error("❌ Невозможно продолжить без Excel")
        return

    found_sheets = []  # (school_code, app_num, pdf_path, teacher_name, sheet_name)

    logging.info("\n⚙️ Обработка Excel-файлов:")
    for excel_path, folder_name, folder_surname in all_excel_files:
        try:
            logging.info(f"\n📁 Обработка файла: {excel_path.name} (папка: {folder_name}, фамилия: {folder_surname})")

            # Проверяем, есть ли информация о листах для этой фамилии
            if folder_surname not in sheet_info_by_teacher:
                logging.warning(f"  ⚠️ Нет данных о листах для фамилии: {folder_surname} (папка: {folder_name})")
                continue

            wb = excel_app.Workbooks.Open(str(excel_path))
            logging.info(f"  ✅ Открыт файл: {excel_path.name}")

            sheet_names_in_file = [s.Name for s in wb.Sheets]
            logging.info(f"  📋 Листы в файле: {', '.join(sheet_names_in_file)}")

            for sheet_name in sheet_names_in_file:
                if sheet_name in ("Служебное", "Списки классов"):
                    logging.info(f"  ⚪ Пропускаем служебный лист: {sheet_name}")
                    continue

                # Проверяем, есть ли такой лист для этой фамилии в справочнике
                if sheet_name not in sheet_info_by_teacher[folder_surname]:
                    logging.debug(f"  🔍 Лист '{sheet_name}' не найден в справочнике для {folder_surname}")
                    continue

                app_num, school_code = sheet_info_by_teacher[folder_surname][sheet_name]
                pdf_path = os.path.join(output_folder, f"{folder_surname}_{sheet_name}.pdf")

                logging.info(f"  📄 Найден нужный лист: {sheet_name} (Приложение №{app_num})")
                ws = wb.Worksheets(sheet_name)

                # Обновляем номер приложения и устанавливаем область печати
                if update_app_number_and_set_print_area(ws, app_num):
                    logging.info(f"  ✅ Успешно настроена область печати для {sheet_name}")
                else:
                    logging.warning(f"  ⚠️ Не удалось настроить область печати для {sheet_name}, используем весь лист")

                # Экспортируем только область печати
                logging.info(f"  🖨️ Экспорт в PDF: {sheet_name} → {os.path.basename(pdf_path)}")
                ws.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
                found_sheets.append((school_code, app_num, pdf_path, folder_name, sheet_name))
                logging.info(f"✅ Экспортирован: {folder_name}/{sheet_name} (Приложение №{app_num})")

            wb.Close(SaveChanges=True)  # Сохраняем изменения (обновленные номера приложений)
        except Exception as e:
            logging.error(f"⚠️ Ошибка при обработке {excel_path} ({folder_name}): {e}")
            try:
                # Пытаемся закрыть книгу, даже если произошла ошибка
                wb.Close(SaveChanges=False)
            except:
                pass

    excel_app.Quit()

    if not found_sheets:
        logging.error("❌ Ни один лист из справочника не найден в Excel-файлах.")
        logging.error("\n🔍 Детальная информация для отладки:")
        logging.error("1. Фамилии в справочнике:")
        for surname in sheet_info_by_teacher.keys():
            logging.error(f"   - {surname}")
        logging.error("2. Фамилии из папок:")
        for _, _, surname in all_excel_files:
            logging.error(f"   - {surname}")
        logging.error("3. Примеры названий листов в справочнике:")
        for surname, sheets in sheet_info_by_teacher.items():
            for sheet_name in list(sheets.keys())[:3]:  # Первые 3 листа
                logging.error(f"   - {surname}: {sheet_name}")

        # Выводим в консоль краткую информацию для быстрой отладки
        print("\n🔧 КРАТКАЯ ИНФОРМАЦИЯ ДЛЯ ОТЛАДКИ:")
        print(f"Фамилии в справочнике: {list(sheet_info_by_teacher.keys())}")
        print(f"Фамилии из папок: {list(set([surname for _, _, surname in all_excel_files]))}")
        return

    logging.info(f"\n✅ Найдено {len(found_sheets)} ведомостей для включения в итоговые PDF.")

    # --- Проверка наличия титульных PDF ---
    title_pdfs = {}
    for code in ["13", "16", "17", "22", "АГ"]:
        pdf_path = os.path.join(output_folder, f"title_{code}.pdf")
        if os.path.exists(pdf_path):
            title_pdfs[code] = pdf_path
            logging.info(f"✅ Найден титульный PDF для школы {code}: {pdf_path}")
        else:
            logging.error(f"❌ Отсутствует титульный PDF: {pdf_path}")
            logging.error("❗ Пожалуйста, поместите подписанные титулы в папку вывода с именами:")
            logging.error("    title_13.pdf, title_16.pdf, title_17.pdf, title_22.pdf, title_АГ.pdf")
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

    logging.info("\n📄 Сборка итоговых PDF по школам:")
    for school_code, items in groups.items():
        # Сортируем по номеру приложения, а при одинаковых номерах - по имени преподавателя
        items.sort(key=lambda x: (app_num_to_sort_key(x[0]), x[2]))

        writer = PdfWriter()

        # Титульный лист
        logging.info(f"  📑 Добавление титульного листа для {SCHOOL_NAMES[school_code]}")
        writer.append(PdfReader(title_pdfs[school_code]))

        # Ведомости
        for app_num, pdf_path, teacher_name, sheet_name in items:
            if os.path.exists(pdf_path):
                logging.info(f"  ➕ Добавлено: Приложение №{app_num} ({teacher_name}/{sheet_name})")
                writer.append(PdfReader(pdf_path))

        output_path = os.path.join(output_folder, f"Ведомости_{SCHOOL_NAMES[school_code]}.pdf")
        with open(output_path, "wb") as f:
            writer.write(f)
        logging.info(f"✅ Создан итоговый PDF: {output_path} ({len(items)} приложений)")

    logging.info("\n🎉 ГОТОВО! Все PDF-файлы сформированы.")
    logging.info(f"📄 Список созданных файлов:")
    for school_code in groups.keys():
        output_path = os.path.join(output_folder, f"Ведомости_{SCHOOL_NAMES[school_code]}.pdf")
        logging.info(f"  - {output_path}")

    # Выводим в консоль краткую информацию о результатах
    print("\n✅ Обработка завершена успешно!")
    print(f"📄 Создано PDF-файлов для школ: {len(groups)}")
    for school_code in groups.keys():
        print(f"  - {SCHOOL_NAMES[school_code]} ({len(groups[school_code])} приложений)")
    print(f"\n📜 Полный лог сохранен в файл: {log_file}")


if __name__ == "__main__":
    main()