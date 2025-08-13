#!/usr/bin/env python3
"""
Автоматический тест системы обновлений GUI
"""

import tkinter as tk
import threading
import time
import sys
from pathlib import Path

# Импортируем наш GUI модуль
sys.path.append(str(Path(__file__).parent))
from excel_automation_gui import ExcelAutomationGUI

def test_update_button():
    """Тестирует кнопку обновления автоматически"""
    print("🧪 Запуск автоматического теста системы обновлений...")
    
    # Создаем GUI
    root = tk.Tk()
    root.withdraw()  # Скрываем главное окно для теста
    
    app = ExcelAutomationGUI(root)
    
    def auto_test():
        """Автоматически тестирует функцию обновления"""
        time.sleep(1)  # Ждем инициализации
        
        print("🔍 Тестируем функцию check_for_updates...")
        try:
            # Вызываем функцию проверки обновлений
            app.check_for_updates()
            print("✅ Функция check_for_updates вызвана успешно")
            
            # Ждем завершения проверки
            time.sleep(5)
            print("⏱️  Тест завершен")
            
        except Exception as e:
            print(f"❌ Ошибка в тесте: {e}")
        finally:
            root.after(100, root.quit)  # Закрываем приложение
    
    # Запускаем тест в отдельном потоке
    test_thread = threading.Thread(target=auto_test, daemon=True)
    test_thread.start()
    
    # Запускаем GUI на короткое время
    root.after(8000, root.quit)  # Автоматически закрываем через 8 секунд
    root.mainloop()

if __name__ == "__main__":
    test_update_button()
