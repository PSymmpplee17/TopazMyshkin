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
import subprocess
from pathlib import Path
import logging
import pandas as pd
import re
from datetime import datetime
import requests
import zipfile
import io
import shutil
import subprocess

# Отключаем SSL warnings
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Импортируем наши модули
from automation_tool_fixed import ExcelProcessor
from material_sorter import MaterialSorter
from excel_to_txt_converter import ExcelToTxtConverter

GITHUB_REPO = "PSymmpplee17/TopazMyshkin"  # Укажите свой репозиторий (без .git и https)
APP_VERSION = "1.0.2"  # Текущая версия приложения

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
        self.root.title("MyshkinTOOL - Полная обработка")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # Переменные
        self.input_file = tk.StringVar()
        self.order_number = tk.StringVar()
        self.current_step = tk.StringVar(value="Запуск приложения...")
        
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
        
        # Заголовок и версия
        header_frame = ttk.Frame(main_frame)
        header_frame.grid(row=0, column=0, columnspan=3, pady=(0, 20), sticky=(tk.W, tk.E))
        header_frame.columnconfigure(1, weight=1)
        
        title_label = ttk.Label(header_frame, text="Excel Automation Tool", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, sticky=tk.W)
        
        version_label = ttk.Label(header_frame, text=f"v{APP_VERSION}", 
                                 font=('Arial', 10), foreground='gray')
        version_label.grid(row=0, column=2, sticky=tk.E)
        
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
            self.progress.start()
            
            # Определяем, запущены ли мы как exe или как скрипт
            is_exe = getattr(sys, 'frozen', False)
            current_exe = Path(sys.executable if is_exe else __file__)
            
            # Скачиваем обновление с обработкой SSL ошибок
            headers = {
                'User-Agent': 'ExcelAutomationTool/1.0.0',
                'Accept': 'application/octet-stream'
            }
            
            try:
                resp = requests.get(url, stream=True, timeout=30, headers=headers, verify=True)
            except requests.exceptions.SSLError:
                logging.info("SSL ошибка при скачивании, пробуем без верификации сертификата")
                resp = requests.get(url, stream=True, timeout=30, headers=headers, verify=False)
                
            if resp.status_code != 200:
                messagebox.showerror("Ошибка", "Не удалось скачать архив обновления")
                return
            
            self.current_step.set("Установка обновления...")
            
            if url.endswith('.exe'):
                # Скачиваем новый exe файл
                new_exe_path = current_exe.parent / f"ExcelAutomationTool_v{new_version}.exe"
                with open(new_exe_path, 'wb') as f:
                    for chunk in resp.iter_content(chunk_size=8192):
                        f.write(chunk)
                
                # Создаем батник для замены файла и перезапуска
                batch_script = current_exe.parent / "update.bat"
                batch_content = f"""@echo off
timeout /t 2 /nobreak >nul
del "{current_exe}" >nul 2>&1
move "{new_exe_path}" "{current_exe}" >nul 2>&1
start "" "{current_exe}"
del "%~f0" >nul 2>&1
"""
                with open(batch_script, 'w', encoding='cp1251') as f:
                    f.write(batch_content)
                
                messagebox.showinfo("Обновление", 
                                   f"Обновление до версии {new_version} загружено!\n\n"
                                   "Приложение перезапустится автоматически.")
                
                # Запускаем батник и закрываем приложение
                subprocess.Popen([str(batch_script)], shell=True)
                self.root.quit()
                
            else:
                # Обновление через zip архив (для разработки)
                with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
                    # Извлекаем только .py файлы
                    for member in zf.namelist():
                        if member.endswith('.py') or member == 'requirements.txt':
                            zf.extract(member, current_exe.parent)
                
                messagebox.showinfo("Обновление", 
                                   f"Обновление до версии {new_version} завершено!\n\n"
                                   "Перезапустите приложение для применения изменений.")
                self.current_step.set("Обновление завершено - перезапустите приложение")
                
        except Exception as e:
            messagebox.showerror("Ошибка обновления", f"Произошла ошибка при обновлении:\n{str(e)}")
            self.current_step.set("Ошибка обновления")
        finally:
            self.progress.stop()

    def auto_check_update(self):
        """Автоматическая проверка обновлений при запуске"""
        def check_in_background():
            try:
                # Показываем статус проверки
                self.root.after(0, lambda: self.current_step.set("Проверка обновлений..."))
                
                url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
                
                # Добавляем SSL контекст и заголовки для решения проблем с сертификатами
                headers = {
                    'User-Agent': 'ExcelAutomationTool/1.0.0',
                    'Accept': 'application/vnd.github.v3+json'
                }
                
                # Попробуем с верификацией SSL, если не получится - отключим её
                try:
                    resp = requests.get(url, timeout=10, headers=headers, verify=True)
                except requests.exceptions.SSLError:
                    # Если SSL ошибка, попробуем без верификации
                    logging.info("SSL ошибка, пробуем без верификации сертификата")
                    resp = requests.get(url, timeout=10, headers=headers, verify=False)
                
                if resp.status_code != 200:
                    self.root.after(0, lambda: self.current_step.set("Готов к работе"))
                    return
                    
                data = resp.json()
                latest_version = data.get("tag_name", "")
                
                if latest_version and latest_version != APP_VERSION:
                    # Есть обновление
                    assets = data.get("assets", [])
                    download_url = None
                    
                    # Ищем подходящий файл для скачивания
                    for asset in assets:
                        if asset["name"].endswith(".exe") or asset["name"].endswith(".zip"):
                            download_url = asset["browser_download_url"]
                            break
                    
                    if download_url:
                        # Показываем уведомление через 2 секунды после запуска
                        self.root.after(2000, lambda: self.show_update_notification(latest_version, download_url))
                    else:
                        self.root.after(0, lambda: self.current_step.set("Готов к работе"))
                else:
                    # Обновлений нет
                    self.root.after(0, lambda: self.current_step.set("Готов к работе (актуальная версия)"))
                    # Через 3 секунды убираем это сообщение
                    self.root.after(3000, lambda: self.current_step.set("Готов к работе"))
                    
            except Exception as e:
                # Если ошибка сети, просто продолжаем работу
                logging.info(f"Не удалось проверить обновления: {e}")
                self.root.after(0, lambda: self.current_step.set("Готов к работе"))
        
        # Запускаем проверку в отдельном потоке
        thread = threading.Thread(target=check_in_background)
        thread.daemon = True
        thread.start()
    
    def show_update_notification(self, latest_version, download_url):
        """Показывает уведомление об обновлении"""
        # Создаем красивое окно уведомления
        update_window = tk.Toplevel(self.root)
        update_window.title("Доступно обновление")
        update_window.geometry("400x250")
        update_window.resizable(False, False)
        update_window.grab_set()  # Модальное окно
        
        # Центрируем окно
        update_window.update_idletasks()
        x = (update_window.winfo_screenwidth() // 2) - (200)
        y = (update_window.winfo_screenheight() // 2) - (125)
        update_window.geometry(f"+{x}+{y}")
        
        # Фрейм для содержимого
        main_frame = ttk.Frame(update_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Иконка и заголовок
        ttk.Label(main_frame, text="🔄 Доступно обновление", 
                 font=('Arial', 14, 'bold')).pack(pady=(0, 10))
        
        # Информация о версиях
        info_text = f"""Найдена новая версия приложения!

Текущая версия: {APP_VERSION}
Новая версия: {latest_version}

Новая версия содержит улучшения и исправления.
Обновление произойдет автоматически."""
        
        ttk.Label(main_frame, text=info_text, justify=tk.CENTER).pack(pady=(0, 20))
        
        # Кнопки
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        def start_update():
            update_window.destroy()
            self.download_and_update(download_url, latest_version)
        
        def cancel_update():
            update_window.destroy()
            # Сохраняем информацию о пропущенном обновлении
            self.current_step.set("Обновление отложено")
        
        ttk.Button(button_frame, text="Обновить сейчас", 
                  command=start_update, 
                  style='Accent.TButton').pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(button_frame, text="Отложить", 
                  command=cancel_update).pack(side=tk.LEFT)
        
        # Автоматическое обновление через 10 секунд
        def auto_update():
            if update_window.winfo_exists():
                start_update()
        
        update_window.after(10000, auto_update)  # 10 секунд
        
        # Обратный отсчет
        countdown_label = ttk.Label(main_frame, text="Автоматическое обновление через 10 сек", 
                                   font=('Arial', 8), foreground='gray')
        countdown_label.pack()
        
        def update_countdown(seconds):
            if update_window.winfo_exists() and seconds > 0:
                countdown_label.config(text=f"Автоматическое обновление через {seconds} сек")
                update_window.after(1000, lambda: update_countdown(seconds-1))
        
        update_countdown(10)


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
