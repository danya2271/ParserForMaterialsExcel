import os
import re
from collections import defaultdict
import openpyxl
from docx import Document
import tkinter as tk
from tkinter import filedialog, ttk, scrolledtext
import win32com.client as win32  # Для конвертации .doc -> .docx
from tkinter import messagebox

# --- КОНФИГУРАЦИЯ ---

# Для стандартных профилей (трубы, уголки) ищет "число x число"
MATERIAL_REGEX_PATTERN = r'(\d+(?:,\d+)?(?:\s*[хx]\s*\d+(?:,\d+)?){1,2})'
# Для арматуры, ищет "L=число"
REINFORCEMENT_REGEX_PATTERN = r'[Aa][54]00[СсСc]?.*?[⌀ø]\s*\d+.*?L\s*=\s*(\d+)'


NAME_KEYWORDS = ['наим', 'материал', 'позиция']
LENGTH_KEYWORDS = ['длин', 'метр']
QUANTITY_KEYWORDS = ['кол', 'шт', 'колич']
CONTINGENCY_PERCENTAGE = 10
FILENAME_FILTER_KEYWORD = 'журнал'
EXCLUDE_KEYWORD = 'лист'
# ------------------------------------


# --- Вспомогательные функции ---
def find_columns_indices(header_row):
    """Находит индексы столбцов по ключевым словам."""
    indices = {'name': None, 'length': None, 'quantity': None}
    for i, cell_text in enumerate(header_row):
        if not cell_text: continue
        lower_cell_text = str(cell_text).lower()
        if any(kw in lower_cell_text for kw in NAME_KEYWORDS): indices['name'] = i
        if any(kw in lower_cell_text for kw in LENGTH_KEYWORDS): indices['length'] = i
        if any(kw in lower_cell_text for kw in QUANTITY_KEYWORDS): indices['quantity'] = i
    return indices

def parse_value(value):
    """Преобразует значение в число, если возможно."""
    if isinstance(value, (int, float)): return value
    if isinstance(value, str):
        try:
            return float(value.replace(',', '.').strip())
        except (ValueError, TypeError):
            return 0
    return 0

def natural_sort_key(s):
    """Ключ для "естественной" сортировки строк с числами."""
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]


def process_row(row_data, column_indices, file_specific_data):
    """Обрабатывает строку данных, ищет материалы и обновляет подсчеты."""
    name_idx = column_indices['name']
    length_idx = column_indices.get('length') # .get() чтобы не было ошибки, если столбец не найден
    quantity_idx = column_indices['quantity']

    if not all(idx is not None for idx in [name_idx, quantity_idx]):
        return

    if len(row_data) <= max(filter(None, [name_idx, length_idx, quantity_idx])):
        return

    material_cell_content = str(row_data[name_idx]).strip()

    if EXCLUDE_KEYWORD in material_cell_content.lower():
        return

    quantity = parse_value(row_data[quantity_idx])
    if quantity <= 0: return

    reinforcement_match = re.search(REINFORCEMENT_REGEX_PATTERN, material_cell_content, re.IGNORECASE)
    if reinforcement_match:
        length_mm = parse_value(reinforcement_match.group(1))
        if length_mm > 0:
            normalized_name = f"Арматура {material_cell_content}"
            total_length_m = (length_mm / 1000) * quantity
            file_specific_data[normalized_name] += total_length_m
            return

    material_match = re.search(MATERIAL_REGEX_PATTERN, material_cell_content)
    if material_match and length_idx is not None:
        found_name = material_match.group(1)
        normalized_name = found_name.replace(',', '.').replace(' ', '')

        length_m = parse_value(row_data[length_idx])
        if length_m > 0:
            file_specific_data[normalized_name] += length_m * quantity


# --- Основной класс приложения с GUI (с добавлением обработки .doc) ---
class ParserApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Универсальный парсер журналов v6.1 - Естественная сортировка")
        self.geometry("800x600")
        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        self.word_app = None

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
        """Выводит сообщение в текстовое поле лога."""
        self.results_text.config(state="normal")
        self.results_text.insert(tk.END, message + "\n")
        self.results_text.config(state="disabled")
        self.results_text.see(tk.END)
        self.update_idletasks()

    def select_folder(self):
        """Открывает диалог выбора папки."""
        path = filedialog.askdirectory(title="Выберите папку для сканирования")
        if path:
            self.folder_path.set(path)
            self.log(f"Выбрана папка: {path}")

    def on_closing(self):
        """Корректно закрывает приложение и COM-объект."""
        if self.word_app:
            try:
                self.word_app.Quit(False)
            except Exception as e:
                print(f"Не удалось корректно закрыть MS Word: {e}")
        self.destroy()

    def run_parser(self):
        """Основная функция запуска парсинга."""
        start_path = self.folder_path.get()
        if not start_path:
            self.log("Ошибка: Папка не выбрана.")
            return

        self.results_text.config(state="normal"); self.results_text.delete('1.0', tk.END); self.results_text.config(state="disabled")
        self.run_button.config(state="disabled")

        master_data = defaultdict(lambda: defaultdict(float))

        self.log(f"Начинаю поиск файлов c '{FILENAME_FILTER_KEYWORD}' в названии...")
        self.log(f"Стартовая директория: {start_path}\n")

        processed_files = []
        try:
            for dirpath, _, filenames in os.walk(start_path):
                for filename in filenames:
                    if FILENAME_FILTER_KEYWORD not in filename.lower(): continue
                    processed_files.append(os.path.join(dirpath, filename))

            # ИЗМЕНЕНИЕ 1: Сортируем файлы используя "естественную" сортировку
            processed_files.sort(key=natural_sort_key)

            for file_path in processed_files:
                relative_path = os.path.relpath(file_path, start_path)
                file_ext = os.path.splitext(file_path)[1].lower()

                file_specific_data = None
                if file_ext == '.xlsx':
                    self.log(f"[XLSX] Обработка: {relative_path}")
                    file_specific_data = self.parse_xlsx(file_path)
                elif file_ext == '.docx':
                    self.log(f"[DOCX] Обработка: {relative_path}")
                    file_specific_data = self.parse_docx(file_path)
                elif file_ext == '.doc':
                    self.log(f"[DOC] Обработка: {relative_path}")
                    file_specific_data = self.parse_doc(file_path)

                if file_specific_data:
                    master_data[relative_path] = file_specific_data

            self.log("\n-------------------------------------------")
            self.log("--- ИТОГОВЫЙ РАСЧЕТ ПО ВСЕМ ФАЙЛАМ ---")

            if not master_data:
                 self.log(f"Файлы с '{FILENAME_FILTER_KEYWORD}' в названии найдены, но в них нет подходящих материалов или данных.")
            else:
                # ИЗМЕНЕНИЕ 2: Сортируем итоговый вывод по имени файла "естественным" способом
                sorted_files = sorted(master_data.items(), key=lambda item: natural_sort_key(item[0]))

                for filename, file_data in sorted_files:
                    self.log(f"\n===========================================")
                    self.log(f"ФАЙЛ: {filename}")
                    self.log(f"-------------------------------------------")

                    if not file_data:
                        self.log("  > В этом файле не найдено подходящих материалов.")
                        continue

                    sorted_materials = sorted(file_data.items(), key=lambda item: natural_sort_key(item[0]))

                    total_file_length = 0
                    total_file_length_contingency = 0

                    for i, (material, total_length) in enumerate(sorted_materials, 1):
                        final_length_with_contingency = (total_length * (1 + CONTINGENCY_PERCENTAGE / 100)) / 1000
                        total_file_length += total_length / 1000
                        total_file_length_contingency += final_length_with_contingency

                        self.log(f"{i}. Наименование: {material}")
                        self.log(f"   - Суммарная длина (без запаса): {total_length:.3f} м")
                        self.log(f"   - Итоговая длина с запасом ({CONTINGENCY_PERCENTAGE}%): {final_length_with_contingency:.3f} м")



            self.log("\n\n--- Анализ завершен. ---")

        except Exception as e:
            self.log(f"КРИТИЧЕСКАЯ ОШИБКА: {e}")
            if "pywintypes.com_error" in str(e):
                messagebox.showerror("Ошибка COM",
                                     "Не удалось запустить Microsoft Word. Убедитесь, что он установлен и доступен. "
                                     "Возможно, потребуется запустить это приложение от имени администратора.")
        finally:
            self.run_button.config(state="normal")


    def parse_xlsx(self, file_path):
        """Парсит XLSX файл."""
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
            self.log(f"  > Ошибка при чтении файла XLSX: {e}")
        return file_data

    def parse_docx_tables(self, document):
        """Общий парсер таблиц для DOCX документа."""
        file_data = defaultdict(float)
        for table in document.tables:
            header_row_values = [cell.text for cell in table.rows[0].cells]
            column_indices = find_columns_indices(header_row_values)
            if all(idx is not None for idx in column_indices.values()):
                for i in range(1, len(table.rows)):
                    data_row_values = [cell.text for cell in table.rows[i].cells]
                    if any(data_row_values):
                        process_row(data_row_values, column_indices, file_data)
        return file_data

    def parse_docx(self, file_path):
        """Парсит DOCX файл."""
        try:
            document = Document(file_path)
            return self.parse_docx_tables(document)
        except Exception as e:
            self.log(f"  > Ошибка при чтении файла DOCX: {e}")
            return None

    def parse_doc(self, file_path):
        """Конвертирует DOC в DOCX и затем парсит."""
        if not self.word_app:
            try:
                self.word_app = win32.Dispatch("Word.Application")
                self.word_app.Visible = False
            except Exception as e:
                raise Exception(f"Не удалось запустить MS Word для конвертации: {e}")

        try:
            # Создаем временный путь для DOCX файла
            docx_path = os.path.splitext(file_path)[0] + "._temp_converted.docx"
            doc = self.word_app.Documents.Open(file_path, ReadOnly=True)
            # 16 = wdFormatXMLDocument (формат .docx)
            doc.SaveAs2(docx_path, FileFormat=16)
            doc.Close(False)
            data = self.parse_docx(docx_path)
            os.remove(docx_path)
            return data
        except Exception as e:
            self.log(f"  > Ошибка при конвертации/чтении файла DOC: {e}")
            return None


if __name__ == "__main__":
    try:
        import win32com.client
    except ImportError:
        messagebox.showerror("Отсутствует библиотека",
                             "Для обработки .doc файлов необходима библиотека pywin32.\n"
                             "Пожалуйста, установите ее командой: pip install pywin32")
        exit()

    app = ParserApp()
    app.mainloop()
