# json_database.py
"""
Модуль для работы с JSON базой данных продуктов
Хранит данные в структурированных JSON файлах с поддержкой резервного копирования и истории изменений
"""

import json
import os
import shutil
import logging
from datetime import datetime
from typing import List, Dict, Any, Optional, Union
import hashlib
import re
from pathlib import Path

# Настройка логирования
logger = logging.getLogger('JSONDatabase')


class JSONDatabaseError(Exception):
    """Базовое исключение для ошибок JSON базы данных"""
    pass


class JSONDatabase:
    """
    Класс для работы с JSON базой данных продуктов

    Структура:
    - products.json - основной файл со всеми продуктами
    - history/ - папка с историей изменений
    - backups/ - папка с резервными копиями
    - stats.json - статистика и метаданные
    - schema.json - схема данных
    """

    # Версия схемы данных
    SCHEMA_VERSION = "1.0"

    # Обязательные поля для продукта
    REQUIRED_FIELDS = ['name']

    # Допустимые поля (основные)
    ALLOWED_FIELDS = [
        'id', 'name', 'version', 'gk', 'description', 'subsystem',
        'certificate', 'platform', 'language', 'owner', 'module',
        'sheet_name', 'source_file', 'created_at', 'last_updated',
        'is_active', 'tags', 'notes'
    ]

    def __init__(self, base_dir: str = 'json_database'):
        """
        Инициализация JSON базы данных

        Args:
            base_dir: базовая директория для хранения файлов
        """
        self.base_dir = base_dir
        self.products_file = os.path.join(base_dir, 'products.json')
        self.stats_file = os.path.join(base_dir, 'stats.json')
        self.schema_file = os.path.join(base_dir, 'schema.json')
        self.history_dir = os.path.join(base_dir, 'history')
        self.backups_dir = os.path.join(base_dir, 'backups')
        self.temp_dir = os.path.join(base_dir, 'temp')

        # Создаем структуру директорий
        self._ensure_directories()

        # Инициализируем файлы
        self._init_files()

        # Загружаем данные
        self.products = self._load_products()
        self.stats = self._load_stats()
        self.schema = self._load_schema()

        # Проверяем целостность данных
        self._validate_data()

        logger.info(f"JSON Database инициализирована в {base_dir}")
        logger.info(f"Загружено {len(self.products)} продуктов")

    def _ensure_directories(self):
        """Создание необходимых директорий"""
        directories = [
            self.base_dir,
            self.history_dir,
            self.backups_dir,
            self.temp_dir
        ]

        for directory in directories:
            os.makedirs(directory, exist_ok=True)
            logger.debug(f"Директория создана/проверена: {directory}")

    def _init_files(self):
        """Инициализация файлов если их нет"""
        # schema.json
        if not os.path.exists(self.schema_file):
            schema = {
                'version': self.SCHEMA_VERSION,
                'created_at': datetime.now().isoformat(),
                'description': 'Схема данных для продуктов ПО',
                'fields': self.ALLOWED_FIELDS,
                'required': self.REQUIRED_FIELDS
            }
            with open(self.schema_file, 'w', encoding='utf-8') as f:
                json.dump(schema, f, ensure_ascii=False, indent=2)
            logger.info(f"Создан файл схемы: {self.schema_file}")

        # products.json
        if not os.path.exists(self.products_file):
            with open(self.products_file, 'w', encoding='utf-8') as f:
                json.dump([], f, ensure_ascii=False, indent=2)
            logger.info(f"Создан файл продуктов: {self.products_file}")

        # stats.json
        if not os.path.exists(self.stats_file):
            initial_stats = {
                'total_products': 0,
                'by_subsystem': {},
                'with_certificate': 0,
                'without_certificate': 0,
                'last_update': None,
                'created_at': datetime.now().isoformat(),
                'version': self.SCHEMA_VERSION,
                'total_backups': 0,
                'last_backup': None,
                'database_size': 0
            }
            with open(self.stats_file, 'w', encoding='utf-8') as f:
                json.dump(initial_stats, f, ensure_ascii=False, indent=2)
            logger.info(f"Создан файл статистики: {self.stats_file}")

    def _load_products(self) -> List[Dict]:
        """Загрузка продуктов из JSON"""
        try:
            with open(self.products_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, list):
                    # Проверяем каждый продукт, но не отбрасываем невалидные
                    valid_products = []
                    for product in data:
                        if self._validate_product(product):
                            valid_products.append(product)
                        else:
                            # Пытаемся очистить продукт
                            cleaned_product = self._clean_product(product)
                            if cleaned_product:
                                valid_products.append(cleaned_product)
                                logger.info(f"Продукт очищен и добавлен: {cleaned_product.get('name', 'unknown')}")
                            else:
                                logger.warning(f"Продукт с ID {product.get('id', 'unknown')} не может быть очищен")
                    return valid_products
                else:
                    logger.error(f"products.json должен содержать список, получен {type(data)}")
                    return []
        except json.JSONDecodeError as e:
            logger.error(f"Ошибка парсинга products.json: {e}")
            # Пытаемся восстановить из резервной копии
            self._attempt_recovery()
            return []
        except FileNotFoundError:
            logger.error(f"Файл products.json не найден")
            return []

    def _clean_product(self, product: Dict) -> Optional[Dict]:
        """Очистка продукта от недопустимых полей"""
        try:
            cleaned = {}

            # Копируем только разрешенные поля
            for key, value in product.items():
                if key in self.ALLOWED_FIELDS:
                    cleaned[key] = value
                elif key == 'row_index':  # Игнорируем служебные поля
                    continue
                else:
                    logger.debug(f"Поле {key} удалено из продукта")

            # Проверяем наличие обязательных полей
            for field in self.REQUIRED_FIELDS:
                if field not in cleaned or not cleaned.get(field):
                    logger.warning(f"Отсутствует обязательное поле {field} после очистки")
                    return None

            return cleaned

        except Exception as e:
            logger.error(f"Ошибка очистки продукта: {e}")
            return None

    def _load_stats(self) -> Dict:
        """Загрузка статистики из JSON"""
        try:
            with open(self.stats_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, FileNotFoundError) as e:
            logger.error(f"Ошибка загрузки stats.json: {e}")
            return {
                'total_products': 0,
                'by_subsystem': {},
                'with_certificate': 0,
                'without_certificate': 0,
                'last_update': None,
                'created_at': datetime.now().isoformat(),
                'version': self.SCHEMA_VERSION,
                'total_backups': 0,
                'last_backup': None,
                'database_size': 0
            }

    def _load_schema(self) -> Dict:
        """Загрузка схемы данных"""
        try:
            with open(self.schema_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, FileNotFoundError) as e:
            logger.error(f"Ошибка загрузки schema.json: {e}")
            return {
                'version': self.SCHEMA_VERSION,
                'created_at': datetime.now().isoformat(),
                'fields': self.ALLOWED_FIELDS,
                'required': self.REQUIRED_FIELDS
            }

    def _validate_product(self, product: Dict) -> bool:
        """Валидация структуры продукта (более мягкая)"""
        try:
            # Проверка наличия обязательных полей
            for field in self.REQUIRED_FIELDS:
                if field not in product or not product.get(field):
                    logger.debug(f"Отсутствует обязательное поле: {field}")
                    return False

            # Проверка типа данных
            if not isinstance(product.get('name', ''), str):
                logger.debug(f"Поле 'name' должно быть строкой")
                return False

            if 'gk' in product and product['gk'] is not None and not isinstance(product['gk'], list):
                logger.debug(f"Поле 'gk' должно быть списком")
                return False

            # Не проверяем наличие неразрешенных полей - просто игнорируем их при сохранении
            return True

        except Exception as e:
            logger.error(f"Ошибка валидации: {e}")
            return False

    def _validate_data(self):
        """Проверка целостности данных"""
        # Проверяем уникальность ID
        ids = [p.get('id') for p in self.products if p.get('id') is not None]
        if len(ids) != len(set(ids)):
            logger.warning("Обнаружены дублирующиеся ID!")
            # Исправляем дубликаты
            self._fix_duplicate_ids()

        # Проверяем, что у всех продуктов есть ID
        for i, product in enumerate(self.products):
            if 'id' not in product or product['id'] is None:
                product['id'] = self._generate_id()
                logger.info(f"Продукту на позиции {i} присвоен новый ID: {product['id']}")

        if self.products:
            self._save_products(create_backup=False)

    def _fix_duplicate_ids(self):
        """Исправление дублирующихся ID"""
        seen_ids = set()
        for product in self.products:
            pid = product.get('id')
            if pid in seen_ids:
                new_id = self._generate_id()
                logger.info(f"Изменение дублирующегося ID {pid} на {new_id}")
                product['id'] = new_id
            else:
                if pid is not None:
                    seen_ids.add(pid)

    def _generate_id(self) -> int:
        """Генерация нового уникального ID"""
        if not self.products:
            return 1

        # Находим максимальный ID
        max_id = 0
        for p in self.products:
            pid = p.get('id', 0)
            if isinstance(pid, (int, float)) and pid > max_id:
                max_id = int(pid)

        return max_id + 1

    def _save_products(self, create_backup: bool = True):
        """
        Сохранение продуктов в JSON

        Args:
            create_backup: создать резервную копию перед сохранением
        """
        if create_backup:
            self.create_backup('before_save')

        # Сохраняем во временный файл
        temp_file = os.path.join(self.temp_dir, f'products_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json')

        try:
            # Очищаем продукты перед сохранением
            cleaned_products = []
            for product in self.products:
                cleaned_product = {}
                for key, value in product.items():
                    if key in self.ALLOWED_FIELDS:
                        cleaned_product[key] = value
                cleaned_products.append(cleaned_product)

            with open(temp_file, 'w', encoding='utf-8') as f:
                json.dump(cleaned_products, f, ensure_ascii=False, indent=2, default=str)

            # Если успешно сохранили во временный файл, заменяем основной
            shutil.move(temp_file, self.products_file)

            # Обновляем статистику
            self._update_stats()

            logger.info(f"Продукты сохранены: {len(cleaned_products)} записей")

        except Exception as e:
            logger.error(f"Ошибка сохранения products.json: {e}")
            # Удаляем временный файл в случае ошибки
            if os.path.exists(temp_file):
                os.remove(temp_file)
            raise JSONDatabaseError(f"Не удалось сохранить продукты: {e}")

    def _save_stats(self):
        """Сохранение статистики в JSON"""
        try:
            with open(self.stats_file, 'w', encoding='utf-8') as f:
                json.dump(self.stats, f, ensure_ascii=False, indent=2, default=str)
        except Exception as e:
            logger.error(f"Ошибка сохранения stats.json: {e}")

    def _update_stats(self):
        """Обновление статистики"""
        total = len(self.products)

        # Статистика по подсистемам
        by_subsystem = {}
        with_cert = 0
        active_count = 0

        for product in self.products:
            if product.get('is_active', True):
                active_count += 1

                subs = product.get('subsystem', 'Не определена')
                if subs:
                    by_subsystem[subs] = by_subsystem.get(subs, 0) + 1

                if product.get('certificate'):
                    with_cert += 1

        # Размер базы данных
        db_size = 0
        if os.path.exists(self.products_file):
            db_size = os.path.getsize(self.products_file)

        self.stats.update({
            'total_products': total,
            'active_products': active_count,
            'by_subsystem': by_subsystem,
            'with_certificate': with_cert,
            'without_certificate': active_count - with_cert,
            'last_update': datetime.now().isoformat(),
            'database_size': db_size
        })

        self._save_stats()

    def _add_to_history(self, product_id: int, action: str, old_data: Optional[Dict] = None,
                        new_data: Optional[Dict] = None, changed_by: str = 'system',
                        comment: str = ''):
        """
        Добавление записи в историю

        Args:
            product_id: ID продукта
            action: действие (CREATE, UPDATE, DELETE, RESTORE)
            old_data: старые данные
            new_data: новые данные
            changed_by: кто изменил
            comment: комментарий
        """
        try:
            history_file = os.path.join(self.history_dir, f'product_{product_id}.json')

            # Загружаем существующую историю
            history = []
            if os.path.exists(history_file):
                with open(history_file, 'r', encoding='utf-8') as f:
                    history = json.load(f)

            # Добавляем новую запись
            history.append({
                'action': action,
                'old_data': old_data,
                'new_data': new_data,
                'changed_by': changed_by,
                'changed_at': datetime.now().isoformat(),
                'comment': comment
            })

            # Ограничиваем историю последними 100 записями
            if len(history) > 100:
                history = history[-100:]

            with open(history_file, 'w', encoding='utf-8') as f:
                json.dump(history, f, ensure_ascii=False, indent=2, default=str)

            logger.debug(f"Добавлена запись в историю для продукта {product_id}: {action}")

        except Exception as e:
            logger.error(f"Ошибка сохранения истории для продукта {product_id}: {e}")

    def _attempt_recovery(self):
        """Попытка восстановления данных из резервной копии"""
        logger.info("Попытка восстановления данных из последней резервной копии")

        backups = self.get_backups_list()
        if backups:
            latest_backup = backups[0]  # Самая новая копия
            backup_path = latest_backup.get('path')

            if backup_path and os.path.exists(backup_path):
                backup_file = os.path.join(backup_path, 'products.json')
                if os.path.exists(backup_file):
                    try:
                        with open(backup_file, 'r', encoding='utf-8') as f:
                            data = json.load(f)
                            if isinstance(data, list):
                                self.products = data
                                logger.info(f"Данные восстановлены из резервной копии: {backup_path}")
                                self._save_products(create_backup=False)
                    except Exception as e:
                        logger.error(f"Ошибка восстановления из резервной копии: {e}")

    def create_backup(self, reason: str = 'manual') -> str:
        """
        Создание резервной копии базы данных

        Args:
            reason: причина создания копии

        Returns:
            str: путь к созданной копии
        """
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_name = f'backup_{timestamp}_{reason}'
        backup_dir = os.path.join(self.backups_dir, backup_name)

        try:
            # Создаем директорию для бэкапа
            os.makedirs(backup_dir, exist_ok=True)

            # Копируем основные файлы
            if os.path.exists(self.products_file):
                shutil.copy2(self.products_file, os.path.join(backup_dir, 'products.json'))

            if os.path.exists(self.stats_file):
                shutil.copy2(self.stats_file, os.path.join(backup_dir, 'stats.json'))

            if os.path.exists(self.schema_file):
                shutil.copy2(self.schema_file, os.path.join(backup_dir, 'schema.json'))

            # Копируем историю (опционально)
            history_backup = os.path.join(backup_dir, 'history')
            if os.path.exists(self.history_dir) and os.listdir(self.history_dir):
                shutil.copytree(self.history_dir, history_backup)

            # Создаем метаданные бэкапа
            metadata = {
                'created_at': timestamp,
                'reason': reason,
                'total_products': len(self.products),
                'files': ['products.json', 'stats.json', 'schema.json'],
                'has_history': os.path.exists(self.history_dir) and bool(os.listdir(self.history_dir)),
                'database_size': os.path.getsize(self.products_file) if os.path.exists(self.products_file) else 0
            }

            with open(os.path.join(backup_dir, 'metadata.json'), 'w', encoding='utf-8') as f:
                json.dump(metadata, f, ensure_ascii=False, indent=2)

            # Обновляем статистику
            self.stats['total_backups'] = len(self.get_backups_list())
            self.stats['last_backup'] = timestamp
            self._save_stats()

            logger.info(f"Создана резервная копия: {backup_dir}")

            # Очищаем старые бэкапы (оставляем последние 20)
            self._cleanup_old_backups()

            return backup_dir

        except Exception as e:
            logger.error(f"Ошибка создания резервной копии: {e}")
            return ''

    def _cleanup_old_backups(self, keep_last: int = 20):
        """Очистка старых резервных копий"""
        try:
            backups = self.get_backups_list()
            if len(backups) > keep_last:
                for backup in backups[keep_last:]:
                    backup_path = backup.get('path')
                    if backup_path and os.path.exists(backup_path):
                        shutil.rmtree(backup_path)
                        logger.info(f"Удалена старая резервная копия: {backup_path}")
        except Exception as e:
            logger.error(f"Ошибка очистки старых бэкапов: {e}")

    def restore_from_backup(self, backup_dir: str) -> bool:
        """
        Восстановление из резервной копии

        Args:
            backup_dir: путь к директории с бэкапом

        Returns:
            bool: успешность восстановления
        """
        try:
            if not os.path.exists(backup_dir):
                logger.error(f"Директория бэкапа не найдена: {backup_dir}")
                return False

            # Проверяем наличие файлов
            products_backup = os.path.join(backup_dir, 'products.json')
            if not os.path.exists(products_backup):
                logger.error(f"Файл products.json не найден в бэкапе")
                return False

            # Создаем бэкап текущего состояния перед восстановлением
            self.create_backup('before_restore')

            # Загружаем данные из бэкапа для проверки
            with open(products_backup, 'r', encoding='utf-8') as f:
                test_data = json.load(f)

            if not isinstance(test_data, list):
                logger.error(f"Неверный формат данных в бэкапе")
                return False

            # Восстанавливаем файлы
            shutil.copy2(products_backup, self.products_file)

            stats_backup = os.path.join(backup_dir, 'stats.json')
            if os.path.exists(stats_backup):
                shutil.copy2(stats_backup, self.stats_file)

            schema_backup = os.path.join(backup_dir, 'schema.json')
            if os.path.exists(schema_backup):
                shutil.copy2(schema_backup, self.schema_file)

            # Восстанавливаем историю
            history_backup = os.path.join(backup_dir, 'history')
            if os.path.exists(history_backup):
                # Сохраняем текущую историю
                current_history_backup = os.path.join(self.backups_dir,
                                                      f'history_before_restore_{datetime.now().strftime("%Y%m%d_%H%M%S")}')
                if os.path.exists(self.history_dir) and os.listdir(self.history_dir):
                    shutil.copytree(self.history_dir, current_history_backup)

                # Очищаем и восстанавливаем
                if os.path.exists(self.history_dir):
                    shutil.rmtree(self.history_dir)
                shutil.copytree(history_backup, self.history_dir)

            # Перезагружаем данные
            self.products = self._load_products()
            self.stats = self._load_stats()
            self.schema = self._load_schema()

            logger.info(f"Восстановлено из бэкапа: {backup_dir}")
            return True

        except Exception as e:
            logger.error(f"Ошибка восстановления из бэкапа: {e}")
            return False

    def get_backups_list(self) -> List[Dict]:
        """
        Получение списка доступных резервных копий

        Returns:
            List[Dict]: список бэкапов с метаданными
        """
        backups = []

        try:
            for item in os.listdir(self.backups_dir):
                backup_path = os.path.join(self.backups_dir, item)
                if os.path.isdir(backup_path) and item.startswith('backup_'):
                    metadata_file = os.path.join(backup_path, 'metadata.json')

                    if os.path.exists(metadata_file):
                        with open(metadata_file, 'r', encoding='utf-8') as f:
                            metadata = json.load(f)
                            metadata['path'] = backup_path
                            metadata['name'] = item
                            backups.append(metadata)
                    else:
                        # Если нет metadata.json, создаем базовую информацию
                        created = item.replace('backup_', '').split('_')[0] if item.startswith('backup_') else 'unknown'
                        backups.append({
                            'created_at': created,
                            'reason': 'unknown',
                            'path': backup_path,
                            'name': item,
                            'total_products': 'unknown',
                            'files': []
                        })

            # Сортируем по дате создания (новые сверху)
            backups.sort(key=lambda x: x.get('created_at', ''), reverse=True)

        except Exception as e:
            logger.error(f"Ошибка получения списка бэкапов: {e}")

        return backups

    def add_product(self, product: Dict) -> Optional[int]:
        """
        Добавление нового продукта

        Args:
            product: данные продукта

        Returns:
            Optional[int]: ID нового продукта или None
        """
        try:
            # Очищаем продукт от недопустимых полей
            cleaned_product = self._clean_product(product)
            if not cleaned_product:
                logger.error(f"Не удалось очистить продукт: {product.get('name', 'unknown')}")
                return None

            # Валидация
            if not self._validate_product(cleaned_product):
                logger.error(f"Невалидная структура продукта после очистки: {cleaned_product}")
                return None

            # Генерируем ID
            product_id = self._generate_id()

            # Добавляем метаданные
            new_product = cleaned_product.copy()
            new_product['id'] = product_id
            new_product['created_at'] = datetime.now().isoformat()
            new_product['last_updated'] = datetime.now().isoformat()
            new_product['is_active'] = True

            # Обеспечиваем наличие обязательных полей
            if 'gk' not in new_product or new_product['gk'] is None:
                new_product['gk'] = []
            if 'subsystem' not in new_product or not new_product['subsystem']:
                new_product['subsystem'] = 'Не определена'
            if 'tags' not in new_product:
                new_product['tags'] = []

            # Добавляем в список
            self.products.append(new_product)

            # Сохраняем
            self._save_products()

            # Добавляем в историю
            self._add_to_history(
                product_id,
                'CREATE',
                None,
                new_product,
                'manual',
                f"Создан продукт: {new_product.get('name')}"
            )

            logger.info(f"Добавлен продукт ID {product_id}: {product.get('name', '')}")
            return product_id

        except Exception as e:
            logger.error(f"Ошибка добавления продукта: {e}")
            return None

    def add_products_batch(self, products: List[Dict]) -> List[int]:
        """
        Добавление нескольких продуктов (batch mode)

        Args:
            products: список продуктов

        Returns:
            List[int]: список ID добавленных продуктов
        """
        added_ids = []

        try:
            for product in products:
                cleaned_product = self._clean_product(product)
                if cleaned_product and self._validate_product(cleaned_product):
                    product_id = self._generate_id()

                    new_product = cleaned_product.copy()
                    new_product['id'] = product_id
                    new_product['created_at'] = datetime.now().isoformat()
                    new_product['last_updated'] = datetime.now().isoformat()
                    new_product['is_active'] = True

                    if 'gk' not in new_product:
                        new_product['gk'] = []
                    if 'subsystem' not in new_product:
                        new_product['subsystem'] = 'Не определена'

                    self.products.append(new_product)
                    added_ids.append(product_id)

            if added_ids:
                self._save_products()
                logger.info(f"Добавлено {len(added_ids)} продуктов в batch режиме")

            return added_ids

        except Exception as e:
            logger.error(f"Ошибка batch добавления продуктов: {e}")
            return added_ids

    def update_product(self, product_id: int, updates: Dict) -> bool:
        """
        Обновление продукта

        Args:
            product_id: ID продукта
            updates: обновленные данные

        Returns:
            bool: успешность обновления
        """
        try:
            # Ищем продукт
            for i, product in enumerate(self.products):
                if product.get('id') == product_id:
                    # Сохраняем старые данные
                    old_data = product.copy()

                    # Очищаем обновления
                    cleaned_updates = {}
                    for key, value in updates.items():
                        if key in self.ALLOWED_FIELDS:
                            cleaned_updates[key] = value

                    # Применяем обновления
                    for key, value in cleaned_updates.items():
                        if key != 'id':
                            product[key] = value

                    product['last_updated'] = datetime.now().isoformat()

                    # Валидация после обновления
                    if not self._validate_product(product):
                        logger.error(f"Продукт ID {product_id} стал невалидным после обновления")
                        # Возвращаем старые данные
                        self.products[i] = old_data
                        return False

                    # Сохраняем
                    self._save_products()

                    # Добавляем в историю
                    self._add_to_history(
                        product_id,
                        'UPDATE',
                        old_data,
                        product,
                        'manual',
                        f"Обновлен продукт: {product.get('name')}"
                    )

                    logger.info(f"Обновлен продукт ID {product_id}")
                    return True

            logger.warning(f"Продукт ID {product_id} не найден")
            return False

        except Exception as e:
            logger.error(f"Ошибка обновления продукта: {e}")
            return False

    def delete_product(self, product_id: int, hard_delete: bool = False) -> bool:
        """
        Удаление продукта

        Args:
            product_id: ID продукта
            hard_delete: полное удаление (False - мягкое удаление)

        Returns:
            bool: успешность удаления
        """
        try:
            if hard_delete:
                # Полное удаление
                for i, product in enumerate(self.products):
                    if product.get('id') == product_id:
                        old_data = product.copy()
                        del self.products[i]

                        self._save_products()
                        self._add_to_history(
                            product_id,
                            'HARD_DELETE',
                            old_data,
                            None,
                            'manual',
                            f"Полное удаление продукта: {old_data.get('name')}"
                        )

                        logger.info(f"Полное удаление продукта ID {product_id}")
                        return True
            else:
                # Мягкое удаление (деактивация)
                return self.update_product(product_id, {'is_active': False})

            return False

        except Exception as e:
            logger.error(f"Ошибка удаления продукта: {e}")
            return False

    def restore_product(self, product_id: int) -> bool:
        """
        Восстановление мягко удаленного продукта

        Args:
            product_id: ID продукта

        Returns:
            bool: успешность восстановления
        """
        try:
            for product in self.products:
                if product.get('id') == product_id and not product.get('is_active', True):
                    return self.update_product(product_id, {'is_active': True})

            logger.warning(f"Продукт ID {product_id} не найден или уже активен")
            return False

        except Exception as e:
            logger.error(f"Ошибка восстановления продукта: {e}")
            return False

    def get_product(self, product_id: int) -> Optional[Dict]:
        """Получение продукта по ID"""
        for product in self.products:
            if product.get('id') == product_id:
                return product.copy()
        return None

    def get_all_products(self, include_inactive: bool = False) -> List[Dict]:
        """
        Получение всех продуктов

        Args:
            include_inactive: включать неактивные

        Returns:
            List[Dict]: список продуктов
        """
        if include_inactive:
            return [p.copy() for p in self.products]
        return [p.copy() for p in self.products if p.get('is_active', True)]

    def find_products(self,
                      subsystem: Optional[str] = None,
                      search: Optional[str] = None,
                      with_certificate: Optional[bool] = None,
                      gk_number: Optional[str] = None,
                      include_inactive: bool = False) -> List[Dict]:
        """
        Поиск продуктов с фильтрацией

        Args:
            subsystem: фильтр по подсистеме
            search: поиск по названию и описанию
            with_certificate: фильтр по наличию сертификата
            gk_number: фильтр по номеру ГК
            include_inactive: включать неактивные

        Returns:
            List[Dict]: отфильтрованный список продуктов
        """
        results = []

        for product in self.products:
            if not include_inactive and not product.get('is_active', True):
                continue

            # Фильтр по подсистеме
            if subsystem and subsystem != 'Все подсистемы' and subsystem != 'Все':
                if product.get('subsystem') != subsystem:
                    continue

            # Поиск по тексту
            if search:
                search_lower = search.lower()
                name = product.get('name', '').lower()
                description = product.get('description', '').lower() if product.get('description') else ''

                if search_lower not in name and search_lower not in description:
                    continue

            # Фильтр по сертификату
            if with_certificate is not None:
                has_cert = bool(product.get('certificate'))
                if with_certificate and not has_cert:
                    continue
                if not with_certificate and has_cert:
                    continue

            # Фильтр по ГК
            if gk_number:
                product_gk = product.get('gk', [])
                if not any(gk_number.upper() in gk.upper() for gk in product_gk):
                    continue

            results.append(product.copy())

        return results

    def get_product_history(self, product_id: int) -> List[Dict]:
        """
        Получение истории изменений продукта

        Args:
            product_id: ID продукта

        Returns:
            List[Dict]: история изменений
        """
        history_file = os.path.join(self.history_dir, f'product_{product_id}.json')

        if not os.path.exists(history_file):
            return []

        try:
            with open(history_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"Ошибка загрузки истории для продукта {product_id}: {e}")
            return []

    def get_subsystems(self) -> List[str]:
        """Получение списка подсистем"""
        subsystems = set()
        for product in self.products:
            if product.get('is_active', True):
                subs = product.get('subsystem', 'Не определена')
                if subs and subs != 'Не определена':
                    subsystems.add(subs)

        # Добавляем 'Не определена' если есть продукты без подсистемы
        has_undefined = any(p.get('subsystem') == 'Не определена' for p in self.products if p.get('is_active', True))
        if has_undefined:
            subsystems.add('Не определена')

        return sorted(list(subsystems))

    def get_stats(self) -> Dict:
        """Получение статистики"""
        return self.stats.copy()

    def search_by_name(self, name_part: str, case_sensitive: bool = False, exact: bool = False) -> List[Dict]:
        """
        Поиск продуктов по части названия

        Args:
            name_part: часть названия для поиска
            case_sensitive: учитывать регистр
            exact: точное совпадение

        Returns:
            List[Dict]: найденные продукты
        """
        if not case_sensitive:
            name_part = name_part.lower()

        results = []
        for product in self.products:
            if not product.get('is_active', True):
                continue

            product_name = product.get('name', '')
            if not case_sensitive:
                product_name = product_name.lower()

            if exact:
                if product_name == name_part:
                    results.append(product.copy())
            else:
                if name_part in product_name:
                    results.append(product.copy())

        return results

    def get_products_by_gk(self, gk_number: str) -> List[Dict]:
        """
        Получение продуктов по номеру ГК

        Args:
            gk_number: номер ГК (полный или частичный)

        Returns:
            List[Dict]: продукты с указанным ГК
        """
        results = []
        gk_upper = gk_number.upper()

        for product in self.products:
            if not product.get('is_active', True):
                continue

            product_gk = product.get('gk', [])
            if any(gk_upper in gk.upper() for gk in product_gk):
                results.append(product.copy())

        return results

    def get_products_by_subsystem(self, subsystem: str) -> List[Dict]:
        """
        Получение продуктов по подсистеме

        Args:
            subsystem: название подсистемы

        Returns:
            List[Dict]: продукты указанной подсистемы
        """
        results = []
        for product in self.products:
            if not product.get('is_active', True):
                continue

            if product.get('subsystem') == subsystem:
                results.append(product.copy())

        return results

    def export_to_file(self, file_path: str, format: str = 'json') -> bool:
        """
        Экспорт базы данных в файл

        Args:
            file_path: путь для сохранения
            format: формат (json, csv, txt)

        Returns:
            bool: успешность экспорта
        """
        try:
            if format == 'json':
                data = {
                    'export_date': datetime.now().isoformat(),
                    'schema_version': self.SCHEMA_VERSION,
                    'stats': self.stats,
                    'products': self.products
                }
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2, default=str)

            elif format == 'csv':
                import csv
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    # Определяем все возможные поля
                    all_fields = set()
                    for product in self.products:
                        all_fields.update(product.keys())

                    # Убираем сложные поля
                    exclude_fields = ['gk', 'tags']
                    fieldnames = [f for f in all_fields if f not in exclude_fields]

                    writer = csv.DictWriter(f, fieldnames=fieldnames, restval='', extrasaction='ignore')
                    writer.writeheader()

                    for product in self.products:
                        # Преобразуем списки в строки
                        row = product.copy()
                        if 'gk' in row:
                            row['gk'] = ', '.join(row['gk']) if row['gk'] else ''
                        if 'tags' in row:
                            row['tags'] = ', '.join(row['tags']) if row['tags'] else ''
                        writer.writerow(row)

            elif format == 'txt':
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(f"ЭКСПОРТ БАЗЫ ДАННЫХ ПРОДУКТОВ\n")
                    f.write(f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
                    f.write(f"Всего продуктов: {len(self.products)}\n")
                    f.write("=" * 80 + "\n\n")

                    for product in self.products:
                        if product.get('is_active', True):
                            f.write(f"ID: {product.get('id', 'N/A')}\n")
                            f.write(f"Наименование: {product.get('name', 'N/A')}\n")
                            f.write(f"Версия: {product.get('version', 'N/A')}\n")
                            f.write(f"Подсистема: {product.get('subsystem', 'N/A')}\n")
                            f.write(f"ГК: {', '.join(product.get('gk', []))}\n")
                            f.write(f"Сертификат: {product.get('certificate', 'N/A')}\n")
                            f.write(f"Описание: {product.get('description', 'N/A')}\n")
                            f.write("-" * 40 + "\n")

            logger.info(f"Экспорт в {format.upper()} завершен: {file_path}")
            return True

        except Exception as e:
            logger.error(f"Ошибка экспорта в {format}: {e}")
            return False

    def import_from_file(self, file_path: str, format: str = 'json') -> tuple[int, int]:
        """
        Импорт данных из файла

        Args:
            file_path: путь к файлу
            format: формат файла (json, csv)

        Returns:
            Tuple[int, int]: (добавлено, обновлено)
        """
        added = 0
        updated = 0

        try:
            if format == 'json':
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)

                products = data.get('products', [])
                for product in products:
                    # Ищем существующий продукт
                    existing = None
                    if 'id' in product:
                        existing = self.get_product(product['id'])

                    if existing:
                        # Обновляем
                        if self.update_product(product['id'], product):
                            updated += 1
                    else:
                        # Добавляем новый
                        if 'id' in product:
                            del product['id']  # ID будет сгенерирован автоматически
                        if self.add_product(product):
                            added += 1

            elif format == 'csv':
                import csv
                with open(file_path, 'r', encoding='utf-8-sig') as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        # Преобразуем строки обратно в списки
                        if 'gk' in row and row['gk']:
                            row['gk'] = [g.strip() for g in row['gk'].split(',') if g.strip()]
                        if 'tags' in row and row['tags']:
                            row['tags'] = [t.strip() for t in row['tags'].split(',') if t.strip()]

                        if self.add_product(row):
                            added += 1

            logger.info(f"Импорт из {format.upper()} завершен: добавлено {added}, обновлено {updated}")
            return added, updated

        except Exception as e:
            logger.error(f"Ошибка импорта из {format}: {e}")
            return added, updated

    def clear_database(self, create_backup: bool = True) -> bool:
        """
        Очистка базы данных

        Args:
            create_backup: создать резервную копию перед очисткой

        Returns:
            bool: успешность очистки
        """
        try:
            if create_backup:
                self.create_backup('before_clear')

            self.products = []
            self._save_products()

            # Очищаем историю
            if os.path.exists(self.history_dir):
                shutil.rmtree(self.history_dir)
                os.makedirs(self.history_dir)

            logger.info("База данных очищена")
            return True

        except Exception as e:
            logger.error(f"Ошибка очистки базы данных: {e}")
            return False

    def validate_database(self) -> Dict[str, Any]:
        """
        Проверка целостности базы данных

        Returns:
            Dict: результаты проверки
        """
        issues = []
        stats = {
            'total_products': len(self.products),
            'valid_products': 0,
            'invalid_products': 0,
            'duplicate_ids': [],
            'missing_ids': [],
            'validation_errors': []
        }

        seen_ids = set()

        for i, product in enumerate(self.products):
            # Проверка ID
            pid = product.get('id')
            if pid is None:
                stats['missing_ids'].append(i)
                issues.append(f"Продукт на позиции {i} не имеет ID")
                continue

            if pid in seen_ids:
                stats['duplicate_ids'].append(pid)
                issues.append(f"Дублирующийся ID: {pid}")
            else:
                seen_ids.add(pid)

            # Валидация структуры
            if self._validate_product(product):
                stats['valid_products'] += 1
            else:
                stats['invalid_products'] += 1
                issues.append(f"Невалидная структура у продукта ID {pid}")

        stats['issues'] = issues
        stats['is_valid'] = len(issues) == 0

        return stats

    def get_database_info(self) -> Dict[str, Any]:
        """
        Получение подробной информации о базе данных

        Returns:
            Dict: информация о БД
        """
        info = {
            'path': self.base_dir,
            'files': {
                'products': os.path.exists(self.products_file),
                'stats': os.path.exists(self.stats_file),
                'schema': os.path.exists(self.schema_file)
            },
            'sizes': {},
            'stats': self.stats,
            'history_count': len(os.listdir(self.history_dir)) if os.path.exists(self.history_dir) else 0,
            'backup_count': len(self.get_backups_list()),
            'validation': self.validate_database()
        }

        # Размеры файлов
        for name, path in [('products', self.products_file),
                           ('stats', self.stats_file),
                           ('schema', self.schema_file)]:
            if os.path.exists(path):
                info['sizes'][name] = os.path.getsize(path)

        return info

    def repair_database(self) -> bool:
        """
        Попытка восстановления/репарации базы данных

        Returns:
            bool: успешность восстановления
        """
        logger.info("Запуск процедуры восстановления базы данных")

        try:
            # Создаем бэкап перед восстановлением
            backup_path = self.create_backup('before_repair')

            # Проверяем текущее состояние
            validation = self.validate_database()

            if validation['is_valid']:
                logger.info("База данных не требует восстановления")
                return True

            # Исправляем дублирующиеся ID
            if validation['duplicate_ids']:
                self._fix_duplicate_ids()

            # Добавляем недостающие ID
            for idx in validation['missing_ids']:
                if idx < len(self.products):
                    self.products[idx]['id'] = self._generate_id()
                    logger.info(f"Добавлен ID для продукта на позиции {idx}")

            # Очищаем невалидные продукты
            cleaned_products = []
            for product in self.products:
                cleaned = self._clean_product(product)
                if cleaned:
                    cleaned_products.append(cleaned)
                else:
                    logger.warning(f"Продукт {product.get('id', 'unknown')} удален из-за невалидной структуры")

            self.products = cleaned_products

            # Сохраняем исправленную базу
            self._save_products()

            logger.info("База данных восстановлена")
            return True

        except Exception as e:
            logger.error(f"Ошибка восстановления базы данных: {e}")
            return False


# Пример использования
if __name__ == "__main__":
    # Настройка логирования для примера
    logging.basicConfig(level=logging.INFO)

    # Создаем базу данных
    db = JSONDatabase("test_database")

    # Добавляем тестовый продукт
    test_product = {
        'name': 'PostgreSQL',
        'version': '15.2',
        'subsystem': 'Базы данных',
        'gk': ['ФКУ0123/2024/РИС'],
        'certificate': 'да',
        'description': 'Реляционная база данных'
    }

    product_id = db.add_product(test_product)
    print(f"Добавлен продукт с ID: {product_id}")

    # Получаем статистику
    stats = db.get_stats()
    print(f"Статистика: {stats}")

    # Ищем продукты
    results = db.find_products(subsystem='Базы данных')
    print(f"Найдено продуктов: {len(results)}")

    # Создаем резервную копию
    backup = db.create_backup('test')
    print(f"Создана резервная копия: {backup}")

    # Получаем информацию о БД
    info = db.get_database_info()
    print(f"Информация о БД: {info['validation']}")