#!/usr/bin/env python3
"""
Тестовый скрипт для проверки системы автообновлений
Имитирует GitHub Releases API и тестирует процесс обновления
"""

import json
import threading
import time
from pathlib import Path
import shutil
import tkinter as tk
from tkinter import messagebox

# Импортируем наш модуль
from excel_automation_gui import ExcelAutomationGUI

def create_test_release_info():
    """Создает тестовую информацию о релизе для имитации GitHub API"""
    return {
        "tag_name": "v1.0.9",
        "name": "Test Release v1.0.9",
        "body": "Тестовый релиз для проверки системы автообновлений",
        "assets": [
            {
                "name": "ExcelAutomationTool-v1.0.9.zip",
                "browser_download_url": "file:///" + str(Path("D:/Myshkin/JOB/ExcelAutomationTool-v1.0.9.zip").as_posix())
            }
        ]
    }

def test_version_comparison():
    """Тестируем функцию сравнения версий"""
    print("=== Тестирование сравнения версий ===")
    
    # Создаем тестовый экземпляр GUI
    root = tk.Tk()
    root.withdraw()  # Скрываем окно
    app = ExcelAutomationGUI(root)
    
    test_cases = [
        ("1.0.8", "1.0.9", 1),  # Новая версия доступна
        ("1.0.8", "1.0.8", 0),  # Одинаковые версии  
        ("1.0.9", "1.0.8", -1), # Текущая версия новее
        ("1.0.8", "v1.0.9", 1), # С префиксом v
        ("v1.0.8", "1.0.9", 1), # С префиксом v в current
    ]
    
    for current, latest, expected in test_cases:
        result = app.compare_versions(current, latest)
        status = "✓" if (result > 0) == (expected > 0) and (result == 0) == (expected == 0) and (result < 0) == (expected < 0) else "✗"
        print(f"{status} {current} vs {latest}: {result} (ожидалось: {expected})")
    
    root.destroy()

def test_zip_processing():
    """Тестируем обработку ZIP архивов"""
    print("\n=== Тестирование обработки ZIP архивов ===")
    
    zip_path = Path("D:/Myshkin/JOB/ExcelAutomationTool-v1.0.9.zip")
    
    if not zip_path.exists():
        print(f"✗ ZIP файл не найден: {zip_path}")
        return
    
    import zipfile
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            exe_files = [f for f in zf.namelist() if f.endswith('.exe')]
            print(f"✓ ZIP архив найден: {zip_path}")
            print(f"✓ EXE файлов в архиве: {len(exe_files)}")
            for exe_file in exe_files:
                file_info = zf.getinfo(exe_file)
                print(f"  - {exe_file}: {file_info.file_size:,} байт")
    except Exception as e:
        print(f"✗ Ошибка при обработке ZIP: {e}")

def test_manual_update():
    """Тестируем ручной вызов обновления"""
    print("\n=== Ручной тест обновления ===")
    print("ВНИМАНИЕ: Это протестирует реальное обновление!")
    print("Приложение может быть закрыто и заменено!")
    
    response = input("Продолжить? (y/N): ")
    if response.lower() != 'y':
        print("Тест отменен")
        return
    
    # Создаем GUI для тестирования
    root = tk.Tk()
    app = ExcelAutomationGUI(root)
    
    # Имитируем обнаружение обновления
    fake_url = str(Path("D:/Myshkin/JOB/ExcelAutomationTool-v1.0.9.zip").as_uri())
    print(f"Тестируем обновление с URL: {fake_url}")
    
    try:
        # Запускаем процесс обновления
        app.download_and_update(fake_url, "v1.0.9")
        print("✓ Обновление запущено успешно")
    except Exception as e:
        print(f"✗ Ошибка при обновлении: {e}")
    finally:
        root.destroy()

def main():
    """Главная функция тестирования"""
    print("🧪 ТЕСТИРОВАНИЕ СИСТЕМЫ АВТООБНОВЛЕНИЙ")
    print("=" * 50)
    
    # Тест 1: Сравнение версий
    test_version_comparison()
    
    # Тест 2: Обработка ZIP архивов
    test_zip_processing()
    
    # Тест 3: Показать доступные файлы
    print("\n=== Доступные файлы для тестирования ===")
    job_dir = Path("D:/Myshkin/JOB")
    for zip_file in job_dir.glob("*.zip"):
        print(f"📦 {zip_file.name} ({zip_file.stat().st_size:,} байт)")
    
    # Тест 4: Предложить ручной тест
    print("\n=== Опции тестирования ===")
    print("1. Ручной тест обновления (осторожно!)")
    print("2. Запустить приложение для автоматической проверки")
    print("3. Выход")
    
    choice = input("\nВыберите опцию (1-3): ")
    
    if choice == "1":
        test_manual_update()
    elif choice == "2":
        print("Запускаем приложение...")
        print("Приложение проверит обновления автоматически при запуске")
        import subprocess
        subprocess.run([str(job_dir / "dist" / "ExcelAutomationTool.exe")])
    else:
        print("Выход из тестирования")

if __name__ == "__main__":
    main()
