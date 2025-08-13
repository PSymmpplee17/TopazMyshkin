#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тестер системы обновлений в консоли
"""

from simple_updater import SimpleUpdater
import logging

# Настраиваем логирование
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def test_updater():
    """Тестирует систему обновлений"""
    print("=== ТЕСТИРОВАНИЕ СИСТЕМЫ ОБНОВЛЕНИЙ ===")
    
    # Текущая версия приложения
    current_version = "1.3.0"
    print(f"Текущая версия: {current_version}")
    
    # Создаем updater
    updater = SimpleUpdater(current_version)
    
    # Получаем информацию о коммите
    print("\n--- Информация о текущем коммите ---")
    commit_info = updater.get_commit_info()
    for key, value in commit_info.items():
        print(f"{key.capitalize()}: {value}")
    
    # Проверяем обновления
    print("\n--- Проверка обновлений ---")
    try:
        has_update, new_version = updater.check_for_updates()
        
        if has_update and new_version:
            print(f"✅ НАЙДЕНО ОБНОВЛЕНИЕ: {new_version}")
            print(f"   Текущая версия: {current_version}")
            print(f"   Новая версия: {new_version}")
            
            # Можно было бы обновить версию в файле
            # success = updater.update_version_file(new_version, Path("pyproject.toml"))
            # print(f"   Обновление файла версии: {'✅' if success else '❌'}")
            
        else:
            print("ℹ️  Обновлений не найдено")
            print(f"   У вас установлена последняя версия: {current_version}")
            
    except Exception as e:
        print(f"❌ ОШИБКА при проверке обновлений: {e}")
        logging.error(f"Ошибка проверки обновлений: {e}")
    
    print("\n=== ТЕСТИРОВАНИЕ ЗАВЕРШЕНО ===")

if __name__ == "__main__":
    test_updater()
