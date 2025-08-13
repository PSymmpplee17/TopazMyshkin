#!/usr/bin/env python3
"""
Простой скрипт для тестирования автообновлений
"""

import subprocess
import sys
from pathlib import Path

def test_update_system():
    """Тестирует систему автообновлений"""
    print("🔍 Тестирование системы автообновлений...")
    
    # 1. Проверим текущую версию
    try:
        result = subprocess.run([
            sys.executable, "-m", "semantic_release", "version", "--print"
        ], capture_output=True, text=True, cwd=Path(__file__).parent)
        
        if result.returncode == 0:
            next_version = result.stdout.strip()
            print(f"📦 Следующая версия: {next_version}")
        else:
            print(f"❌ Ошибка получения версии: {result.stderr}")
            
    except Exception as e:
        print(f"❌ Ошибка: {e}")
    
    # 2. Проверим статус git
    try:
        result = subprocess.run([
            "git", "status", "--porcelain"
        ], capture_output=True, text=True, cwd=Path(__file__).parent)
        
        if result.stdout.strip():
            print("⚠️  Есть незафиксированные изменения:")
            print(result.stdout)
        else:
            print("✅ Git репозиторий чистый")
            
    except Exception as e:
        print(f"❌ Ошибка git: {e}")
    
    # 3. Проверим конфигурацию
    config_file = Path(__file__).parent / "pyproject.toml"
    if config_file.exists():
        print("✅ Конфигурация pyproject.toml найдена")
    else:
        print("❌ Конфигурация pyproject.toml не найдена")

if __name__ == "__main__":
    test_update_system()
