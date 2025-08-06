import os
import re
from collections import defaultdict
import openpyxl
from docx import Document
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext

# --- КОНФИГУРАЦИЯ ---
# ИЗМЕНЕНИЕ 1: Вместо списка материалов теперь используется регулярное выражение.
# Оно находит шаблоны "число_х_число" или "число_х_число_х_число".
# \b - граница слова (чтобы не находить шаблон внутри другого слова)
# \d+(?:,\d+)? - число, целое или с запятой
# (?:\s*[хx]\s*\d+(?:,\d+)?){1,2} - группа "разделитель + число", которая повторяется 1 или 2 раза
MATERIAL_REGEX_PATTERN = r'\b(\d+(?:,\d+)?(?:\s*[хx]\s*\d+(?:,\d+)?){1,2})\b'

NAME_KEYWORDS = ['наим', 'материал', 'позиция']
LENGTH_KEYWORDS = ['длин', 'метр']
QUANTITY_KEYWORDS = ['кол', 'шт', 'колич']
CONTINGENCY_PERCENTAGE = 10
FILENAME_FILTER_KEYWORD = 'журнал'
# ------------------------------------

# --- Вспомогательные функции ---
def find_columns_indices(header_row):
    indices = {'name': None, 'length': None, 'quantity': None}
    for i, cell_text in enumerate(header_row):
        if not cell_text: continue
        lower_cell_text = str(cell_text).lower()
        if any(kw in lower_cell_text for kw in NAME_KEYWORDS): indices['name'] = i
        if any(kw in lower_cell_text for kw in LENGTH_KEYWORDS): indices['length'] = i
        if any(kw in lower_cell_text for kw in QUANTITY_KEYWORDS): indices['quantity'] = i
    return indices

def parse_value(value):
    if isinstance(value, (int, float)): return value
    if isinstance(value, str):
        try:
            return float(value.replace(',', '.').strip())
        except (ValueError, TypeError):
            return 0
    return 0

# ИЗМЕНЕНИЕ 2: Функция process_row теперь использует регулярное выражение
def process_row(row_data, column_indices, file_specific_data):
    """Находит материал по шаблону в строке и обновляет данные."""
    name_idx, length_idx, quantity_idx = column_indices['name'], column_indices['length'], column_indices['quantity']
    if len(row_data) <= max(name_idx, length_idx, quantity_idx): return

    material_cell_content = str(row_data[name_idx]).strip()

    # Ищем первое соответствие шаблону в ячейке
    match = re.search(MATERIAL_REGEX_PATTERN, material_cell_content)

    if match:
        # Получаем найденное имя материала (например, "50 х 50 х 3")
        found_name = match.group(1)
        # Нормализуем его: убираем пробелы и меняем запятую на точку для единообразия
        normalized_name = found_name.replace(',', '.').replace(' ', '')

        length = parse_value(row_data[length_idx])
        quantity = parse_value(row_data[quantity_idx])

        if length > 0 and quantity > 0:
            file_specific_data[normalized_name] += length * quantity

# --- Основной класс приложения с GUI (без изменений в структуре) ---
class ParserApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Универсальный парсер журналов v4.0")
        self.geometry("800x600")
        self.folder_path = tk.StringVar()

        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill="both", expand=True)
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill="x", pady=5)

        folder_label = ttk.Label(top_frame, text="Папка для поиска (включая подпапки):")
        folder_label.pack(side="left", padx=(0, 10))
        self.folder_entry = ttk.Entry(top_frame, textvariable=self.folder_path, state="readonly", width=50)
        self.folder_entry.pack(side="left", fill="x", expand=True)
        browse_button = ttk.Button(top_frame, text="Выбрать...", command=self.select_folder)
        browse_button.pack(side="left", padx=(10, 0))

        self.run_button = ttk.Button(main_frame, text="Запустить анализ", command=self.run_parser)
        self.run_button.pack(fill="x", pady=10)
        self.results_text = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, height=20, state="disabled")
        self.results_text.pack(fill="both", expand=True)

    def log(self, message):
        self.results_text.config(state="normal")
        self.results_text.insert(tk.END, message + "\n")
        self.results_text.config(state="disabled")
        self.results_text.see(tk.END)
        self.update_idletasks()

    def select_folder(self):
        path = filedialog.askdirectory(title="Выберите папку для сканирования")
        if path:
            self.folder_path.set(path)
            self.log(f"Выбрана папка: {path}")

    def run_parser(self):
        start_path = self.folder_path.get()
        if not start_path:
            self.log("Ошибка: Папка не выбрана.")
            return

        self.results_text.config(state="normal"); self.results_text.delete('1.0', tk.END); self.results_text.config(state="disabled")
        self.run_button.config(state="disabled")

        master_data = defaultdict(lambda: defaultdict(float))

        self.log(f"Начинаю поиск файлов c '{FILENAME_FILTER_KEYWORD}' в названии...")
        self.log(f"Стартовая директория: {start_path}\n")

        files_with_data_count = 0
        try:
            for dirpath, _, filenames in os.walk(start_path):
                for filename in filenames:
                    if FILENAME_FILTER_KEYWORD not in filename.lower(): continue

                    file_path = os.path.join(dirpath, filename)
                    relative_path = os.path.relpath(file_path, start_path)

                    file_specific_data = None
                    if filename.lower().endswith('.xlsx'):
                        self.log(f"[XLSX] Обработка: {relative_path}")
                        file_specific_data = self.parse_xlsx(file_path)
                    elif filename.lower().endswith('.docx'):
                        self.log(f"[DOCX] Обработка: {relative_path}")
                        file_specific_data = self.parse_docx(file_path)

                    if file_specific_data:
                        files_with_data_count += 1
                        master_data[relative_path] = file_specific_data

            self.log("\n-------------------------------------------")
            self.log("--- ИТОГОВЫЙ РАСЧЕТ ПО КАЖДОМУ ФАЙЛУ ---")

            if files_with_data_count == 0:
                 self.log(f"Файлы с '{FILENAME_FILTER_KEYWORD}' в названии найдены, но в них нет материалов нужного формата.")
            else:
                for file_path, file_data in sorted(master_data.items()):
                    self.log(f"\n======== Результаты для файла: {file_path} ========")
                    if not file_data:
                        self.log("  > В этом файле не найдено материалов, соответствующих шаблону.")
                        continue
                    for material, total_length in sorted(file_data.items()):
                        final_length = (total_length * (1 + CONTINGENCY_PERCENTAGE / 100)) / 1000
                        self.log(f"  > Наименование: {material}")
                        self.log(f"    Суммарная длина: {total_length:.3f}м")
                        self.log(f"    Общая длина с запасом ({CONTINGENCY_PERCENTAGE}%): {final_length:.3f}м")

            self.log("\n--- Анализ завершен. ---")

        except Exception as e:
            self.log(f"КРИТИЧЕСКАЯ ОШИБКА: {e}")
        finally:
            self.run_button.config(state="normal")

    def parse_xlsx(self, file_path):
        file_data = defaultdict(float)
        try:
            workbook = openpyxl.load_workbook(file_path, data_only=True)
            for sheet in workbook.worksheets:
                for row_idx in range(1, sheet.max_row + 1):
                    header_row_values = [cell.value for cell in sheet[row_idx]]
                    column_indices = find_columns_indices(header_row_values)
                    if all(idx is not None for idx in column_indices.values()):
                        for data_row_idx in range(row_idx + 1, sheet.max_row + 1):
                            data_row_values = [cell.value for cell in sheet[data_row_idx]]
                            if any(data_row_values):
                                process_row(data_row_values, column_indices, file_data)
                        break
        except Exception as e:
            self.log(f"  > Ошибка при чтении файла: {e}")
        return file_data

    def parse_docx(self, file_path):
        file_data = defaultdict(float)
        try:
            document = Document(file_path)
            for table in document.tables:
                header_row_values = [cell.text for cell in table.rows[0].cells]
                column_indices = find_columns_indices(header_row_values)
                if all(idx is not None for idx in column_indices.values()):
                    for i in range(1, len(table.rows)):
                        data_row_values = [cell.text for cell in table.rows[i].cells]
                        if any(data_row_values):
                            process_row(data_row_values, column_indices, file_data)
        except Exception as e:
            self.log(f"  > Ошибка при чтении файла: {e}")
        return file_data


if __name__ == "__main__":
    # Не забудьте импортировать re в начале файла!
    app = ParserApp()
    app.mainloop()
