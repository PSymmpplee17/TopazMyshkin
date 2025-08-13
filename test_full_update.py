#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Тестовое приложение для проверки системы обновлений
"""

import logging
from pathlib import Path
from simple_updater import SimpleUpdater

# Текущая версия (специально ставим 1.3.0 для тестирования)
CURRENT_VERSION = "1.3.0"

def main():
    """Тест полного цикла обновления"""
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    
    print("="*50)
    print("ТЕСТИРОВАНИЕ СИСТЕМЫ ОБНОВЛЕНИЙ")
    print("="*50)
    
    print(f"Текущая версия: {CURRENT_VERSION}")
    
    updater = SimpleUpdater(CURRENT_VERSION, Path.cwd())
    
    # 1. Проверяем обновления
    print("\n1. Проверка обновлений...")
    has_update, new_version = updater.check_for_updates()
    
    if has_update:
        print(f"✅ Найдена новая версия: {new_version}")
        
        # 2. Показываем информацию о коммите
        print("\n2. Информация о текущем состоянии:")
        commit_info = updater.get_commit_info()
        print(f"   Коммит: {commit_info['hash']}")
        print(f"   Сообщение: {commit_info['message']}")
        print(f"   Дата: {commit_info['date']}")
        
        # 3. Спрашиваем пользователя
        print(f"\n3. Обновить до версии {new_version}? (y/n): ", end="")
        answer = input().lower().strip()
        
        if answer in ['y', 'yes', 'да', 'д']:
            print("\n4. Выполняем обновление...")
            success = updater.perform_git_pull_update(new_version)
            
            if success:
                print("✅ Обновление успешно завершено!")
                print("🔄 Перезапустите приложение для применения изменений")
                
                # Проверяем, что версия действительно обновилась
                print("\n5. Проверка после обновления:")
                new_commit_info = updater.get_commit_info()
                print(f"   Новый коммит: {new_commit_info['hash']}")
                print(f"   Новое сообщение: {new_commit_info['message']}")
                print(f"   Новая дата: {new_commit_info['date']}")
                
            else:
                print("❌ Ошибка при обновлении!")
        else:
            print("⏭️  Обновление отменено пользователем")
            
    else:
        print("ℹ️  Обновлений нет - у вас последняя версия")
    
    print("\n" + "="*50)

if __name__ == "__main__":
    main()
