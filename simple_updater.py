#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Простая система обновлений без Unicode проблем
Использует только Git теги для определения версий
Избегает проблем с emoji в semantic-release
"""

import subprocess
import sys
import logging
from pathlib import Path
from packaging import version
import json
from typing import Optional, Tuple
import os


class SimpleUpdater:
    """Простая система обновлений без Unicode проблем"""
    
    def __init__(self, current_version: str, repo_path: Optional[Path] = None):
        """
        Инициализация обновлятора
        
        Args:
            current_version: Текущая версия приложения
            repo_path: Путь к Git репозиторию (по умолчанию - текущая директория)
        """
        self.current_version = current_version
        self.repo_path = repo_path or Path.cwd()
        self.logger = logging.getLogger(__name__)
    
    def check_for_updates(self) -> Tuple[bool, Optional[str]]:
        """
        Проверяет наличие обновлений
        
        Returns:
            Tuple[bool, Optional[str]]: (есть_обновления, новая_версия)
        """
        try:
            # Получаем список удаленных тегов
            latest_version = self._get_latest_remote_version()
            if not latest_version:
                self.logger.warning("Не удалось получить удаленные версии")
                return False, None
            
            # Сравниваем версии
            current_ver = version.parse(self.current_version)
            latest_ver = version.parse(latest_version)
            
            has_update = latest_ver > current_ver
            return has_update, latest_version if has_update else None
            
        except Exception as e:
            self.logger.error(f"Ошибка при проверке обновлений: {e}")
            return False, None
    
    def _get_latest_remote_version(self) -> Optional[str]:
        """
        Получает последнюю версию из удаленного репозитория
        
        Returns:
            str: Последняя версия или None
        """
        try:
            # Обновляем информацию о удаленном репозитории
            fetch_result = subprocess.run([
                "git", "fetch", "--tags"
            ], 
            cwd=self.repo_path, 
            capture_output=True, 
            text=True,
            env=self._get_clean_env()
            )
            
            if fetch_result.returncode != 0:
                self.logger.error(f"Ошибка git fetch: {fetch_result.stderr}")
                return None
            
            # Получаем список всех тегов
            tags_result = subprocess.run([
                "git", "tag", "-l", "--sort=-version:refname"
            ], 
            cwd=self.repo_path, 
            capture_output=True, 
            text=True,
            env=self._get_clean_env()
            )
            
            if tags_result.returncode != 0:
                self.logger.error(f"Ошибка git tag: {tags_result.stderr}")
                return None
            
            # Парсим теги и находим последнюю семантическую версию
            tags = tags_result.stdout.strip().split('\n')
            valid_versions = []
            
            for tag in tags:
                tag = tag.strip()
                if not tag:
                    continue
                    
                # Убираем префикс 'v' если есть
                version_str = tag.lstrip('v')
                
                try:
                    # Проверяем, что это валидная семантическая версия
                    parsed_version = version.parse(version_str)
                    if not parsed_version.is_prerelease:  # Игнорируем pre-release
                        valid_versions.append(version_str)
                except Exception:
                    continue
            
            return valid_versions[0] if valid_versions else None
            
        except Exception as e:
            self.logger.error(f"Ошибка при получении удаленных версий: {e}")
            return None
    
    def create_new_version(self, bump_type: str = "patch") -> Optional[str]:
        """
        Создает новую версию и тег
        
        Args:
            bump_type: Тип увеличения версии ("major", "minor", "patch")
            
        Returns:
            str: Новая версия или None при ошибке
        """
        try:
            current_ver = version.parse(self.current_version)
            
            # Увеличиваем версию
            if bump_type == "major":
                new_version = f"{current_ver.major + 1}.0.0"
            elif bump_type == "minor":
                new_version = f"{current_ver.major}.{current_ver.minor + 1}.0"
            else:  # patch
                new_version = f"{current_ver.major}.{current_ver.minor}.{current_ver.micro + 1}"
            
            # Создаем тег
            tag_result = subprocess.run([
                "git", "tag", "-a", f"v{new_version}", "-m", f"Release version {new_version}"
            ],
            cwd=self.repo_path,
            capture_output=True,
            text=True,
            env=self._get_clean_env()
            )
            
            if tag_result.returncode != 0:
                self.logger.error(f"Ошибка создания тега: {tag_result.stderr}")
                return None
            
            # Пушим тег
            push_result = subprocess.run([
                "git", "push", "origin", f"v{new_version}"
            ],
            cwd=self.repo_path,
            capture_output=True,
            text=True,
            env=self._get_clean_env()
            )
            
            if push_result.returncode != 0:
                self.logger.error(f"Ошибка отправки тега: {push_result.stderr}")
                return None
                
            self.logger.info(f"Создана новая версия: {new_version}")
            return new_version
            
        except Exception as e:
            self.logger.error(f"Ошибка создания новой версии: {e}")
            return None
    
    def update_version_file(self, new_version: str, version_file: Path) -> bool:
        """
        Обновляет версию в файле
        
        Args:
            new_version: Новая версия
            version_file: Путь к файлу версии
            
        Returns:
            bool: Успешность операции
        """
        try:
            if version_file.name == "pyproject.toml":
                return self._update_pyproject_version(new_version, version_file)
            elif version_file.suffix == ".py":
                return self._update_python_version(new_version, version_file)
            else:
                self.logger.error(f"Неподдерживаемый тип файла: {version_file}")
                return False
                
        except Exception as e:
            self.logger.error(f"Ошибка обновления файла версии: {e}")
            return False
    
    def _update_pyproject_version(self, new_version: str, pyproject_file: Path) -> bool:
        """Обновляет версию в pyproject.toml"""
        try:
            content = pyproject_file.read_text(encoding='utf-8')
            
            # Простая замена версии в pyproject.toml
            import re
            pattern = r'version\s*=\s*"[^"]*"'
            replacement = f'version = "{new_version}"'
            new_content = re.sub(pattern, replacement, content)
            
            pyproject_file.write_text(new_content, encoding='utf-8')
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка обновления pyproject.toml: {e}")
            return False
    
    def _update_python_version(self, new_version: str, py_file: Path) -> bool:
        """Обновляет версию в Python файле"""
        try:
            content = py_file.read_text(encoding='utf-8')
            
            # Простая замена __version__
            import re
            pattern = r'__version__\s*=\s*["\'][^"\']*["\']'
            replacement = f'__version__ = "{new_version}"'
            new_content = re.sub(pattern, replacement, content)
            
            py_file.write_text(new_content, encoding='utf-8')
            return True
            
        except Exception as e:
            self.logger.error(f"Ошибка обновления Python файла: {e}")
            return False
    
    def _get_clean_env(self) -> dict:
        """
        Получает окружение без проблемных переменных для Git
        
        Returns:
            dict: Очищенное окружение
        """
        env = os.environ.copy()
        
        # Устанавливаем кодировку для избежания Unicode проблем
        env['PYTHONIOENCODING'] = 'utf-8'
        env['LC_ALL'] = 'en_US.UTF-8'
        
        # Отключаем цветной вывод Git
        env['GIT_CONFIG_GLOBAL'] = '/dev/null'
        env['GIT_CONFIG_SYSTEM'] = '/dev/null'
        
        return env
    
    def get_commit_info(self) -> dict:
        """
        Получает информацию о текущем коммите
        
        Returns:
            dict: Информация о коммите
        """
        try:
            # Получаем хеш коммита
            hash_result = subprocess.run([
                "git", "rev-parse", "--short", "HEAD"
            ],
            cwd=self.repo_path,
            capture_output=True,
            text=True,
            env=self._get_clean_env()
            )
            
            # Получаем сообщение коммита (только первую строку)
            msg_result = subprocess.run([
                "git", "log", "-1", "--pretty=format:%s"
            ],
            cwd=self.repo_path,
            capture_output=True,
            text=True,
            env=self._get_clean_env()
            )
            
            # Получаем дату коммита
            date_result = subprocess.run([
                "git", "log", "-1", "--pretty=format:%ci"
            ],
            cwd=self.repo_path,
            capture_output=True,
            text=True,
            env=self._get_clean_env()
            )
            
            return {
                'hash': hash_result.stdout.strip() if hash_result.returncode == 0 else 'unknown',
                'message': msg_result.stdout.strip() if msg_result.returncode == 0 else 'No message',
                'date': date_result.stdout.strip() if date_result.returncode == 0 else 'unknown'
            }
            
        except Exception as e:
            self.logger.error(f"Ошибка получения информации о коммите: {e}")
            return {'hash': 'unknown', 'message': 'Error', 'date': 'unknown'}


def main():
    """Тестирование updater"""
    logging.basicConfig(level=logging.INFO)
    
    updater = SimpleUpdater("1.2.0")
    
    print("Проверяем обновления...")
    has_update, new_version = updater.check_for_updates()
    
    if has_update:
        print(f"Найдена новая версия: {new_version}")
    else:
        print("Обновлений нет")
    
    # Информация о коммите
    commit_info = updater.get_commit_info()
    print(f"Текущий коммит: {commit_info}")


if __name__ == "__main__":
    main()
