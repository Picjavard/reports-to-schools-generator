import os
import sys
import re
import logging
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
from pypdf import PdfWriter, PdfReader
import win32com.client as win32
import pythoncom

# === КОНФИГУРАЦИЯ ===
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


class VedomostiApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор ведомостей для школ")
        self.root.geometry("800x600")
        self.root.resizable(True, True)

        # Переменные для хранения путей
        self.reference_path = tk.StringVar()
        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.mode = tk.StringVar(value="current")  # "current" или "final"

        self.setup_ui()
        self.load_last_paths()

    def setup_ui(self):
        # Главный фрейм
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Режим работы
        mode_frame = ttk.LabelFrame(main_frame, text="Режим работы", padding=10)
        mode_frame.pack(fill=tk.X, pady=5)

        ttk.Radiobutton(mode_frame, text="Текущая успеваемость (левая часть таблицы)",
                        variable=self.mode, value="current").pack(anchor=tk.W)
        ttk.Radiobutton(mode_frame, text="Итоговая успеваемость (правая часть таблицы)",
                        variable=self.mode, value="final").pack(anchor=tk.W)

        # Путь к справочнику
        ref_frame = ttk.LabelFrame(main_frame, text="Справочник ведомостей", padding=10)
        ref_frame.pack(fill=tk.X, pady=5)

        ttk.Entry(ref_frame, textvariable=self.reference_path, width=70).pack(side=tk.LEFT, fill=tk.X, expand=True,
                                                                              padx=(0, 5))
        ttk.Button(ref_frame, text="Выбрать файл...", command=self.select_reference).pack(side=tk.RIGHT)

        # Входная папка
        input_frame = ttk.LabelFrame(main_frame, text="Папка с ведомостями преподавателей", padding=10)
        input_frame.pack(fill=tk.X, pady=5)

        ttk.Entry(input_frame, textvariable=self.input_path, width=70).pack(side=tk.LEFT, fill=tk.X, expand=True,
                                                                            padx=(0, 5))
        ttk.Button(input_frame, text="Выбрать папку...", command=self.select_input).pack(side=tk.RIGHT)

        # Выходная папка
        output_frame = ttk.LabelFrame(main_frame, text="Папка для сохранения результатов", padding=10)
        output_frame.pack(fill=tk.X, pady=5)

        ttk.Entry(output_frame, textvariable=self.output_path, width=70).pack(side=tk.LEFT, fill=tk.X, expand=True,
                                                                              padx=(0, 5))
        ttk.Button(output_frame, text="Выбрать папку...", command=self.select_output).pack(side=tk.RIGHT)

        # Кнопка запуска
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=10)

        self.start_btn = ttk.Button(btn_frame, text="🚀 Запустить обработку", command=self.start_processing,
                                    style="Accent.TButton")
        self.start_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.open_btn = ttk.Button(btn_frame, text="📂 Открыть папку с результатами", command=self.open_output_folder,
                                   state=tk.DISABLED)
        self.open_btn.pack(side=tk.LEFT)

        # Прогресс-бар
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.pack(fill=tk.X, pady=5)

        # Лог
        log_frame = ttk.LabelFrame(main_frame, text="Лог обработки", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=15, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.configure(state=tk.DISABLED)

        # Настройка стиля кнопки
        style = ttk.Style()
        try:
            style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))
        except:
            pass

    def log(self, message, level="INFO"):
        self.log_text.configure(state=tk.NORMAL)
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefix = {
            "INFO": "ℹ️",
            "WARNING": "⚠️",
            "ERROR": "❌",
            "SUCCESS": "✅"
        }.get(level, "ℹ️")

        self.log_text.insert(tk.END, f"[{timestamp}] {prefix} {message}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)
        self.root.update()

    def select_reference(self):
        path = filedialog.askopenfilename(
            title="Выберите файл 'Список ведомостей.xlsx'",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            self.reference_path.set(path)
            self.save_last_paths()

    def select_input(self):
        path = filedialog.askdirectory(title="Выберите папку с папками преподавателей")
        if path:
            self.input_path.set(path)
            self.save_last_paths()

    def select_output(self):
        path = filedialog.askdirectory(title="Выберите папку для сохранения результатов")
        if path:
            self.output_path.set(path)
            self.save_last_paths()

    def save_last_paths(self):
        try:
            paths = {
                "reference": self.reference_path.get(),
                "input": self.input_path.get(),
                "output": self.output_path.get()
            }
            with open("last_paths.txt", "w", encoding="utf-8") as f:
                for key, value in paths.items():
                    if value:
                        f.write(f"{key}={value}\n")
        except:
            pass

    def load_last_paths(self):
        try:
            with open("last_paths.txt", "r", encoding="utf-8") as f:
                for line in f:
                    if "=" in line:
                        key, value = line.strip().split("=", 1)
                        if key == "reference":
                            self.reference_path.set(value)
                        elif key == "input":
                            self.input_path.set(value)
                        elif key == "output":
                            self.output_path.set(value)
        except:
            pass

    def get_teacher_folder_name(self, teacher_fio):
        if not isinstance(teacher_fio, str) or not teacher_fio.strip():
            return ""
        return teacher_fio.split()[0].strip()

    def get_excel_cell_value(self, ws, row, col):
        try:
            return ws.Cells(row, col).Value
        except:
            return None

    def update_app_number_and_set_print_area(self, ws, app_num_from_ref, mode="current"):
        try:
            app_cells = []
            for row in range(1, 6):
                for col in range(1, 61):
                    cell_value = self.get_excel_cell_value(ws, row, col)
                    if cell_value and isinstance(cell_value, str) and "ПРИЛОЖЕНИЕ" in cell_value:
                        app_cells.append((row, col, cell_value))
                        if mode == "final" and len(app_cells) >= 2:
                            break
                if mode == "final" and len(app_cells) >= 2:
                    break

            if not app_cells:
                self.log("Не найдены ячейки 'ПРИЛОЖЕНИЕ'", "WARNING")
                return False

            if mode == "current":
                target_row, target_col, _ = app_cells[0]
                start_col = 1
                end_col = target_col
                self.log(f"ТЕКУЩИЕ: столбцы A-{chr(64 + end_col)}", "INFO")

            elif mode == "final":
                if len(app_cells) < 2:
                    self.log("Для итоговых требуется 2 ячейки 'ПРИЛОЖЕНИЕ'", "WARNING")
                    return False

                _, first_col, _ = app_cells[0]
                target_row, target_col, _ = app_cells[1]
                start_col = first_col + 1
                end_col = target_col
                self.log(f"ИТОГОВЫЕ: столбцы {chr(64 + start_col)}-{chr(64 + end_col)}", "INFO")

            # Обновляем номер приложения
            cell_to_update = ws.Cells(target_row, target_col)
            new_text = f"ПРИЛОЖЕНИЕ №{app_num_from_ref}"
            cell_to_update.Value = new_text

            # Последняя строка
            try:
                last_row = ws.Cells(ws.Rows.Count, start_col).End(-4162).Row
                last_row = min(last_row, 200)
            except:
                last_row = 200

            # Устанавливаем область печати
            ws.PageSetup.PrintArea = ws.Range(
                ws.Cells(1, start_col),
                ws.Cells(last_row, end_col)
            ).Address

            return True

        except Exception as e:
            self.log(f"Ошибка настройки печати: {e}", "ERROR")
            return False

    def open_output_folder(self):
        path = self.output_path.get()
        if path and os.path.exists(path):
            os.startfile(path)

    def start_processing(self):
        # Валидация путей
        if not self.reference_path.get():
            messagebox.showerror("Ошибка", "Не выбран файл справочника 'Список ведомостей.xlsx'")
            return

        if not os.path.exists(self.reference_path.get()):
            messagebox.showerror("Ошибка", "Файл справочника не найден")
            return

        if not self.input_path.get():
            messagebox.showerror("Ошибка", "Не выбрана папка с ведомостями преподавателей")
            return

        if not os.path.exists(self.input_path.get()):
            messagebox.showerror("Ошибка", "Входная папка не найдена")
            return

        if not self.output_path.get():
            messagebox.showerror("Ошибка", "Не выбрана папка для сохранения результатов")
            return

        os.makedirs(self.output_path.get(), exist_ok=True)

        # Отключаем кнопку и запускаем прогресс
        self.start_btn.config(state=tk.DISABLED)
        self.progress.start(10)
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.configure(state=tk.DISABLED)

        try:
            self.process_files()
            self.log("✅ Обработка завершена успешно!", "SUCCESS")
            self.open_btn.config(state=tk.NORMAL)
        except Exception as e:
            self.log(f"❌ КРИТИЧЕСКАЯ ОШИБКА: {e}", "ERROR")
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{e}")
        finally:
            self.progress.stop()
            self.start_btn.config(state=tk.NORMAL)

    def process_files(self):
        # Настройка логирования в файл
        mode_name = "итоговые" if self.mode.get() == "final" else "текущие"
        log_file = os.path.join(self.output_path.get(),
                                f"processing_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file, encoding='utf-8'),
                logging.StreamHandler(sys.stdout)
            ]
        )

        self.log(f"Начало обработки ({mode_name} ведомости)", "INFO")
        self.log(f"Справочник: {self.reference_path.get()}", "INFO")
        self.log(f"Входная папка: {self.input_path.get()}", "INFO")
        self.log(f"Выходная папка: {self.output_path.get()}", "INFO")

        # Загрузка справочника
        try:
            df = pd.read_excel(self.reference_path.get(), sheet_name="Лист1", dtype=str)
            self.log(f"Загружено {len(df)} записей из справочника", "INFO")
        except Exception as e:
            self.log(f"Ошибка загрузки справочника: {e}", "ERROR")
            raise

        # Фильтрация и подготовка данных
        df = df.dropna(subset=["Название листа"])
        sheet_info_by_teacher = {}
        teacher_stats = {}

        for _, row in df.iterrows():
            sheet_name = row["Название листа"].strip()
            app_num = row["Номер Приложения"].strip() if pd.notna(row["Номер Приложения"]) else ""

            school_full = row["Школа"]
            school_code = SCHOOL_TO_CODE.get(school_full.strip())
            if not school_code:
                continue

            teacher_fio = row["ФИО Преподавателя"]
            teacher_surname = self.get_teacher_folder_name(teacher_fio)

            if teacher_surname not in sheet_info_by_teacher:
                sheet_info_by_teacher[teacher_surname] = {}

            sheet_info_by_teacher[teacher_surname][sheet_name] = (app_num, school_code)

        # Поиск Excel-файлов
        all_excel_files = []
        for item in Path(self.input_path.get()).iterdir():
            if item.is_dir():
                folder_name = item.name
                folder_surname = self.get_teacher_folder_name(folder_name)
                for f in item.glob("*.xlsx"):
                    all_excel_files.append((f, folder_name, folder_surname))

        self.log(f"Найдено {len(all_excel_files)} Excel-файлов", "INFO")

        if not all_excel_files:
            self.log("Не найдено ни одного Excel-файла", "ERROR")
            return

        # Инициализация Excel
        self.log("Подключение к Excel...", "INFO")
        pythoncom.CoInitialize()
        excel_app = win32.Dispatch('Excel.Application')
        excel_app.Visible = False
        excel_app.DisplayAlerts = False

        found_sheets = []

        # Обработка файлов
        for idx, (excel_path, folder_name, folder_surname) in enumerate(all_excel_files, 1):
            self.log(f"[{idx}/{len(all_excel_files)}] Обработка: {folder_name}/{excel_path.name}", "INFO")

            if folder_surname not in sheet_info_by_teacher:
                self.log(f"  Пропускаем: нет данных для {folder_surname}", "WARNING")
                continue

            try:
                wb = excel_app.Workbooks.Open(str(excel_path))
                sheet_names_in_file = [s.Name for s in wb.Sheets]

                for sheet_name in sheet_names_in_file:
                    if sheet_name in ("Служебное", "Списки классов"):
                        continue

                    if sheet_name not in sheet_info_by_teacher[folder_surname]:
                        continue

                    app_num, school_code = sheet_info_by_teacher[folder_surname][sheet_name]
                    suffix = "_итоговые" if self.mode.get() == "final" else ""
                    pdf_path = os.path.join(self.output_path.get(), f"{folder_surname}_{sheet_name}{suffix}.pdf")

                    ws = wb.Worksheets(sheet_name)
                    if self.update_app_number_and_set_print_area(ws, app_num, mode=self.mode.get()):
                        try:
                            ws.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
                            found_sheets.append((school_code, app_num, pdf_path, folder_name, sheet_name))
                            self.log(f"  ✅ Экспортирован: {sheet_name} (Приложение №{app_num})", "INFO")
                        except Exception as e:
                            self.log(f"  ❌ Ошибка экспорта {sheet_name}: {e}", "ERROR")
                    else:
                        if self.mode.get() == "final":
                            self.log(f"  ⚠️ Пропущен (нет данных для итоговых): {sheet_name}", "WARNING")

                wb.Close(SaveChanges=True)
            except Exception as e:
                self.log(f"Ошибка обработки {excel_path}: {e}", "ERROR")
                try:
                    wb.Close(SaveChanges=False)
                except:
                    pass

        excel_app.Quit()

        if not found_sheets:
            self.log("Не найдено ни одной подходящей ведомости", "ERROR")
            return

        self.log(f"Найдено {len(found_sheets)} ведомостей для сборки", "INFO")

        # Проверка титульных листов (не останавливает работу)
        title_pdfs = {}
        for code in ["13", "16", "17", "22", "АГ"]:
            title_filename = f"title_{code}_итоговые.pdf" if self.mode.get() == "final" else f"title_{code}.pdf"
            pdf_path = os.path.join(self.output_path.get(), title_filename)

            if not os.path.exists(pdf_path) and self.mode.get() == "final":
                pdf_path = os.path.join(self.output_path.get(), f"title_{code}.pdf")

            if os.path.exists(pdf_path):
                title_pdfs[code] = pdf_path
                self.log(f"Найден титульный для школы {code}", "INFO")
            else:
                self.log(f"Титульный для школы {code} отсутствует (продолжаем без него)", "WARNING")

        # Сборка итоговых PDF
        from collections import defaultdict
        groups = defaultdict(list)
        for school, app_num, pdf_path, teacher_name, sheet_name in found_sheets:
            groups[school].append((app_num, pdf_path, teacher_name, sheet_name))

        def app_num_to_sort_key(app_num_str):
            if not app_num_str:
                return (999, 0)
            try:
                return (float(app_num_str.strip()), 0)
            except:
                return (999, app_num_str)

        for school_code, items in groups.items():
            items.sort(key=lambda x: (app_num_to_sort_key(x[0]), x[2]))
            writer = PdfWriter()

            if school_code in title_pdfs:
                writer.append(PdfReader(title_pdfs[school_code]))
            else:
                self.log(f"Сборка без титульного для {SCHOOL_NAMES[school_code]}", "WARNING")

            for _, pdf_path, _, _ in items:
                if os.path.exists(pdf_path):
                    writer.append(PdfReader(pdf_path))

            suffix = "_итоговые" if self.mode.get() == "final" else ""
            output_path = os.path.join(self.output_path.get(), f"Ведомости_{SCHOOL_NAMES[school_code]}{suffix}.pdf")
            with open(output_path, "wb") as f:
                writer.write(f)

            self.log(f"Создан файл: {os.path.basename(output_path)} ({len(items)} приложений)", "SUCCESS")

        self.log(f"Лог сохранён в: {log_file}", "INFO")


if __name__ == "__main__":
    root = tk.Tk()
    app = VedomostiApp(root)
    root.mainloop()