"""
GUI приложение для автоматизации обработки Excel файлов

Объединяет все функции:
1. Конвертация .xls в .xlsx
2. Удаление пустых строк и обработка дублей
3. Сортировка по толщине материала
4. Конвертация в TXT файлы

Автор: Автоматизация обработки Excel
Дата: 2025-08-12
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import sys
import os
from pathlib import Path
import logging
import pandas as pd
import re
from datetime import datetime
import requests
import zipfile
import io
import shutil

# Импортируем наши модули
from automation_tool_fixed import ExcelProcessor
from material_sorter import MaterialSorter
from excel_to_txt_converter import ExcelToTxtConverter

GITHUB_REPO = "PSymmpplee17/TopazMyshkin"  # Укажите свой репозиторий (без .git и https)
APP_VERSION = "1.0.0"  # Текущая версия приложения

# Настройка логирования для GUI
class GUILogHandler(logging.Handler):
    """Обработчик для вывода логов в GUI"""
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        
    def emit(self, record):
        msg = self.format(record)
        # Добавляем сообщение в текстовое поле в главном потоке
        self.text_widget.after(0, self._append_log, msg)
        
    def _append_log(self, msg):
        self.text_widget.insert(tk.END, msg + '\n')
        self.text_widget.see(tk.END)
        self.text_widget.update()


class ExcelAutomationGUI:
    """Главный класс GUI приложения"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Automation Tool - Полная обработка")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # Переменные
        self.input_file = tk.StringVar()
        self.order_number = tk.StringVar()
        self.current_step = tk.StringVar(value="Выберите файл")
        
        # Настройка интерфейса
        self.setup_ui()
        
        # Настройка логирования
        self.setup_logging()
        
        # Центрируем окно
        self.center_window()
    
    def center_window(self):
        """Центрирует окно на экране"""
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")
    
    def setup_ui(self):
        """Настройка пользовательского интерфейса"""
        # Главный фрейм
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настройка сетки
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Заголовок
        title_label = ttk.Label(main_frame, text="Excel Automation Tool", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # Выбор файла
        ttk.Label(main_frame, text="Входной файл:").grid(row=1, column=0, sticky=tk.W, pady=5)
        
        file_frame = ttk.Frame(main_frame)
        file_frame.grid(row=1, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        file_frame.columnconfigure(0, weight=1)
        
        self.file_entry = ttk.Entry(file_frame, textvariable=self.input_file, state='readonly')
        self.file_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(file_frame, text="Обзор...", command=self.select_file).grid(row=0, column=1)
        
        # Номер заказа
        ttk.Label(main_frame, text="Номер круга:").grid(row=2, column=0, sticky=tk.W, pady=5)
        
        order_frame = ttk.Frame(main_frame)
        order_frame.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.order_entry = ttk.Entry(order_frame, textvariable=self.order_number, width=10)
        self.order_entry.grid(row=0, column=0, padx=(0, 10))
        
        ttk.Label(order_frame, text="(например: 72 или 1 или 113)").grid(row=0, column=1, sticky=tk.W)
        
        # Текущий шаг
        ttk.Label(main_frame, text="Состояние:").grid(row=3, column=0, sticky=tk.W, pady=5)
        
        self.status_label = ttk.Label(main_frame, textvariable=self.current_step, 
                                     font=('Arial', 10, 'bold'))
        self.status_label.grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # Прогресс-бар
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=10)
        
        # Кнопка запуска
        self.start_button = ttk.Button(main_frame, text="Начать обработку", 
                                      command=self.start_processing, style='Accent.TButton')
        self.start_button.grid(row=5, column=0, columnspan=3, pady=10)
        
        # Лог
        ttk.Label(main_frame, text="Журнал обработки:").grid(row=6, column=0, sticky=tk.W, pady=(10, 5))
        
        log_frame = ttk.Frame(main_frame)
        log_frame.grid(row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=80)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Настройка весов для растягивания
        main_frame.rowconfigure(7, weight=1)
        
        # Кнопка очистки лога
        ttk.Button(main_frame, text="Очистить лог", 
                  command=self.clear_log).grid(row=8, column=0, pady=5)
        
        # Кнопка выхода
        ttk.Button(main_frame, text="Выход", 
                  command=self.root.quit).grid(row=8, column=2, pady=5, sticky=tk.E)
    
    def setup_logging(self):
        """Настройка системы логирования"""
        # Создаем обработчик для GUI
        self.gui_handler = GUILogHandler(self.log_text)
        self.gui_handler.setLevel(logging.INFO)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        self.gui_handler.setFormatter(formatter)
        
        # Настраиваем корневой логгер
        logging.basicConfig(level=logging.INFO, handlers=[self.gui_handler])
        
        # Добавляем обработчик к существующим логгерам модулей
        for module_name in ['__main__', 'automation_tool_fixed', 'material_sorter', 'excel_to_txt_converter']:
            logger = logging.getLogger(module_name)
            logger.addHandler(self.gui_handler)
            logger.setLevel(logging.INFO)
    
    def clear_log(self):
        """Очищает лог"""
        self.log_text.delete(1.0, tk.END)
    
    def select_file(self):
        """Выбор входного файла"""
        filetypes = [
            ('Excel files', '*.xls *.xlsm *.xlsx'),
            ('XLS files', '*.xls'),
            ('XLSM files', '*.xlsm'),
            ('XLSX files', '*.xlsx'),
            ('All files', '*.*')
        ]
        
        filename = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=filetypes,
            initialdir=Path('.').parent
        )
        
        if filename:
            self.input_file.set(filename)
            self.current_step.set("Файл выбран")
            logging.info(f"Выбран файл: {filename}")
    
    def validate_inputs(self):
        """Проверка входных данных"""
        if not self.input_file.get():
            messagebox.showerror("Ошибка", "Выберите входной файл")
            return False
        
        if not Path(self.input_file.get()).exists():
            messagebox.showerror("Ошибка", "Выбранный файл не существует")
            return False
        
        if not self.order_number.get().strip():
            messagebox.showerror("Ошибка", "Введите номер круга")
            return False
        
        try:
            int(self.order_number.get().strip())
        except ValueError:
            messagebox.showerror("Ошибка", "Номер круга должен быть числом")
            return False
        
        return True
    
    def start_processing(self):
        """Запуск обработки файла"""
        if not self.validate_inputs():
            return
        
        # Блокируем интерфейс
        self.start_button.config(state='disabled')
        self.progress.start()
        
        # Запускаем обработку в отдельном потоке
        thread = threading.Thread(target=self.process_file)
        thread.daemon = True
        thread.start()
    
    def process_file(self):
        """Полная обработка файла"""
        try:
            input_path = Path(self.input_file.get())
            order_num = self.order_number.get().strip()
            
            # Формируем OrderID
            current_year = datetime.now().year
            year_suffix = str(current_year)[-2:]
            formatted_number = f"{int(order_num):03d}"
            order_id = f"{year_suffix}-{formatted_number}"
            
            logging.info(f"=== Начало обработки файла ===")
            logging.info(f"Входной файл: {input_path.name}")
            logging.info(f"Номер заказа: {order_num} -> OrderID: {order_id}")
            
            # ШАГ 1: Обработка исходного файла
            self.current_step.set("Шаг 1: Удаление пустых строк и обработка дублей")
            logging.info("ШАГ 1: Удаление пустых строк и обработка дублей")
            
            processor = ExcelProcessor(str(input_path))
            if not processor.load_data():
                raise Exception("Ошибка загрузки данных")
            
            # Удаление пустых строк
            if not processor.remove_empty_rows(col1_idx=3, col2_idx=4):
                raise Exception("Ошибка удаления пустых строк")
            
            # Обработка дублей
            if not processor.process_duplicates_with_order_preservation():
                raise Exception("Ошибка обработки дублей")
            
            # Сохранение результата первого шага
            if not processor.save_data_with_formatting():
                raise Exception("Ошибка сохранения данных")
            
            processed_file = processor.output_file
            logging.info(f"Шаг 1 завершен: {processed_file}")
            
            # ШАГ 2: Сортировка по толщине материала
            self.current_step.set("Шаг 2: Сортировка по толщине материала")
            logging.info("ШАГ 2: Сортировка по толщине материала")
            
            sorter = MaterialSorter(str(processed_file))
            if not sorter.load_data():
                raise Exception("Ошибка загрузки обработанных данных")
            
            if not sorter.sort_data_by_thickness():
                raise Exception("Ошибка сортировки по толщине")
            
            # Создание файла с листами по толщине
            # Передаем OrderID для имени файла
            output_file = processed_file.parent / f"{order_id}_by_thickness.xlsx"
            if not sorter.create_sorted_workbook_auto(str(output_file), order_id):
                raise Exception("Ошибка создания файла по толщине")
            
            thickness_file = sorter.output_file
            logging.info(f"Шаг 2 завершен: {thickness_file}")
            
            # ШАГ 3: Конвертация в TXT файлы
            self.current_step.set("Шаг 3: Конвертация в TXT файлы")
            logging.info("ШАГ 3: Конвертация в TXT файлы")
            
            converter = ExcelToTxtConverter(str(thickness_file))
            if not converter.load_workbook():
                raise Exception("Ошибка загрузки файла для конвертации")
            
            txt_files = converter.convert_all_sheets()
            if not txt_files:
                raise Exception("Ошибка конвертации в TXT")
            
            # Завершение
            self.current_step.set("Обработка завершена успешно!")
            logging.info("=== ОБРАБОТКА ЗАВЕРШЕНА УСПЕШНО ===")
            logging.info(f"Создано TXT файлов: {len(txt_files)}")
            for txt_file in txt_files:
                logging.info(f"  • {txt_file.name}")
            
            # Показываем сообщение об успехе
            self.root.after(0, self.show_success, txt_files)
            
        except Exception as e:
            logging.error(f"ОШИБКА ОБРАБОТКИ: {e}")
            self.current_step.set("Ошибка обработки")
            self.root.after(0, self.show_error, str(e))
        
        finally:
            # Разблокируем интерфейс
            self.root.after(0, self.finish_processing)
    
    def show_success(self, txt_files):
        """Показывает сообщение об успешном завершении"""
        message = f"Обработка завершена успешно!\n\nСоздано файлов: {len(txt_files)}\n\nФайлы:\n"
        for txt_file in txt_files[:5]:  # Показываем максимум 5 файлов
            message += f"• {txt_file.name}\n"
        if len(txt_files) > 5:
            message += f"... и еще {len(txt_files) - 5} файлов"
        
        messagebox.showinfo("Успех", message)
    
    def show_error(self, error_message):
        """Показывает сообщение об ошибке"""
        messagebox.showerror("Ошибка", f"Произошла ошибка:\n{error_message}")
    
    def finish_processing(self):
        """Завершение обработки - разблокировка интерфейса"""
        self.progress.stop()
        self.start_button.config(state='normal')
    
    def check_update(self):
        """Проверяет наличие новой версии на GitHub и предлагает обновиться"""
        try:
            self.current_step.set("Проверка обновлений...")
            url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            resp = requests.get(url, timeout=10)
            if resp.status_code != 200:
                messagebox.showerror("Ошибка", "Не удалось получить информацию о релизе")
                return
            data = resp.json()
            latest_version = data.get("tag_name", "")
            if not latest_version:
                messagebox.showerror("Ошибка", "Не удалось определить версию релиза")
                return
            if latest_version == APP_VERSION:
                messagebox.showinfo("Обновление", "У вас последняя версия приложения")
                return
            # Есть новая версия
            assets = data.get("assets", [])
            download_url = None
            for asset in assets:
                if asset["name"].endswith(".zip"):
                    download_url = asset["browser_download_url"]
                    break
            if not download_url:
                messagebox.showerror("Ошибка", "В релизе не найден архив приложения")
                return
            if messagebox.askyesno("Обновление", f"Доступна новая версия: {latest_version}. Скачать и обновить?"):
                self.download_and_update(download_url, latest_version)
        except Exception as e:
            messagebox.showerror("Ошибка обновления", str(e))

    def download_and_update(self, url, new_version):
        """Скачивает и обновляет приложение"""
        try:
            self.current_step.set("Скачивание обновления...")
            resp = requests.get(url, stream=True, timeout=30)
            if resp.status_code != 200:
                messagebox.showerror("Ошибка", "Не удалось скачать архив обновления")
                return
            with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
                for member in zf.namelist():
                    # Перезаписываем файлы приложения
                    zf.extract(member, Path(__file__).parent)
            messagebox.showinfo("Обновление", f"Обновление до версии {new_version} завершено! Перезапустите приложение.")
            self.current_step.set("Обновление завершено")
        except Exception as e:
            messagebox.showerror("Ошибка обновления", str(e))
            self.current_step.set("Ошибка обновления")

    def auto_check_update(self):
        """Автоматическая проверка обновлений при запуске"""
        def check_in_background():
            try:
                url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
                resp = requests.get(url, timeout=10)
                if resp.status_code != 200:
                    return
                data = resp.json()
                latest_version = data.get("tag_name", "")
                if latest_version and latest_version != APP_VERSION:
                    # Показываем уведомление об обновлении
                    assets = data.get("assets", [])
                    download_url = None
                    for asset in assets:
                        if asset["name"].endswith(".zip") or asset["name"].endswith(".exe"):
                            download_url = asset["browser_download_url"]
                            break
                    
                    if download_url:
                        self.root.after(1000, lambda: self.show_update_notification(latest_version, download_url))
            except:
                pass
        
        # Запускаем проверку в отдельном потоке
        thread = threading.Thread(target=check_in_background)
        thread.daemon = True
        thread.start()
    
    def show_update_notification(self, latest_version, download_url):
        """Показывает уведомление об обновлении"""
        if messagebox.askyesno("Обновление", 
                              f"Доступна новая версия: {latest_version}\n\n"
                              f"Текущая версия: {APP_VERSION}\n\n"
                              "Скачать и обновить сейчас?"):
            self.download_and_update(download_url, latest_version)


# Добавляем метод для автоматической обработки без запроса OrderID
def create_sorted_workbook_auto(sorter, output_file, order_id):
    """
    Автоматическая версия create_sorted_workbook без запроса у пользователя
    """
    if not hasattr(sorter, 'thickness_groups'):
        logging.error("Данные не отсортированы. Сначала вызовите sort_data_by_thickness()")
        return False
    
    try:
        from openpyxl import Workbook
        
        sorter.output_file = Path(output_file)
        logging.info(f"Создаем файл с сортировкой по толщине: {sorter.output_file}")
        
        # Создаем новую книгу
        wb = Workbook()
        wb.remove(wb.active)
        
        # Создаем листы для каждой толщины в определенном порядке
        thickness_order = ["1mm", "1.5mm", "2mm", "3mm"]
        
        for thickness in thickness_order:
            if thickness in sorter.thickness_groups:
                ws = wb.create_sheet(thickness)
                sorter._populate_worksheet(ws, sorter.thickness_groups[thickness], order_id)
                logging.info(f"Создан лист '{thickness}' с {len(sorter.thickness_groups[thickness])} строками")
        
        # Добавляем листы для других толщин (если есть)
        for thickness in sorter.thickness_groups:
            if thickness not in thickness_order:
                ws = wb.create_sheet(thickness)
                sorter._populate_worksheet(ws, sorter.thickness_groups[thickness], order_id)
                logging.info(f"Создан лист '{thickness}' с {len(sorter.thickness_groups[thickness])} строками")
        
        # Создаем лист для неклассифицированных данных (если есть)
        if sorter.unmatched_rows:
            real_unmatched = []
            for row in sorter.unmatched_rows:
                first_value = row.iloc[0] if len(row) > 0 else ""
                if not (isinstance(first_value, str) and 
                       first_value in ['№', 'Порядковый номер', 'OrderID', 'PartName', 'Приоритет', 'nan']):
                    real_unmatched.append(row)
            
            if real_unmatched:
                ws = wb.create_sheet("Неопределенные")
                sorter._populate_worksheet(ws, real_unmatched, order_id)
                logging.info(f"Создан лист 'Неопределенные' с {len(real_unmatched)} строками")
        
        # Сохраняем файл
        wb.save(sorter.output_file)
        wb.close()
        
        logging.info(f"✓ Файл успешно сохранен: {sorter.output_file}")
        return True
        
    except Exception as e:
        logging.error(f"Ошибка при создании файла: {e}")
        return False

# Добавляем метод к классу MaterialSorter
MaterialSorter.create_sorted_workbook_auto = create_sorted_workbook_auto


def main():
    """Главная функция приложения"""
    # Создаем главное окно
    root = tk.Tk()
    
    # Применяем современную тему если доступна
    try:
        root.tk.call('source', 'azure.tcl')
        root.tk.call('set_theme', 'light')
    except:
        pass  # Если тема не доступна, используем стандартную
    
    # Создаем приложение
    app = ExcelAutomationGUI(root)
    app.auto_check_update()
    # Запускаем главный цикл
    root.mainloop()


if __name__ == "__main__":
    main()
