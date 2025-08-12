"""
Скрипт сборки исполняемого файла приложения

Использует PyInstaller для создания .exe файла
"""

import subprocess
import sys
from pathlib import Path
import shutil

def install_pyinstaller():
    """Устанавливает PyInstaller если не установлен"""
    try:
        import PyInstaller
        print("✓ PyInstaller уже установлен")
    except ImportError:
        print("🔄 Устанавливаю PyInstaller...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])
        print("✓ PyInstaller установлен")

def build_exe():
    """Собирает исполняемый файл"""
    app_name = "ExcelAutomationTool"
    spec_file = f"{app_name}.spec"
    
    # Проверяем наличие spec файла
    if not Path(spec_file).exists():
        print(f"❌ Не найден файл {spec_file}")
        return
    
    # Параметры сборки через spec файл
    cmd = [
        'pyinstaller',
        '--clean',                # Очистка кеша
        '--noconfirm',            # Не спрашивать подтверждений
        spec_file
    ]
    
    # Удаляем предыдущие сборки
    dist_dir = Path('dist')
    build_dir = Path('build')
    
    if dist_dir.exists():
        shutil.rmtree(dist_dir)
    if build_dir.exists():
        shutil.rmtree(build_dir)
    
    print("🔄 Собираю исполняемый файл через spec файл...")
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("✓ Сборка завершена успешно!")
        
        exe_file = dist_dir / f"{app_name}.exe"
        if exe_file.exists():
            print(f"✓ Исполняемый файл создан: {exe_file.absolute()}")
            print(f"📁 Размер файла: {exe_file.stat().st_size / 1024 / 1024:.1f} MB")
        else:
            print("❌ Исполняемый файл не найден")
            
    except subprocess.CalledProcessError as e:
        print(f"❌ Ошибка при сборке: {e}")
        if e.stdout:
            print("Вывод:", e.stdout[:1000])  # Ограничиваем вывод
        if e.stderr:
            print("Ошибки:", e.stderr[:1000])  # Ограничиваем вывод ошибок

def main():
    """Главная функция сборки"""
    print("=== Сборка Excel Automation Tool ===")
    
    # Проверяем, что мы в правильной директории
    if not Path('excel_automation_gui.py').exists():
        print("❌ Не найден файл excel_automation_gui.py")
        print("Запустите скрипт из директории с приложением")
        return
    
    # Устанавливаем PyInstaller
    install_pyinstaller()
    
    # Собираем приложение
    build_exe()
    
    print("\n=== Готово! ===")
    print("Исполняемый файл находится в папке 'dist'")

if __name__ == "__main__":
    main()
