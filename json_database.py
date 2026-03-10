# json_database.py
"""
Модуль для работы с JSON базой данных продуктов и проверок
Хранит данные в структурированных JSON файлах с поддержкой резервного копирования и истории изменений
"""

import json
import os
import shutil
import logging
from datetime import datetime
from typing import List, Dict, Any, Optional, Union, Tuple
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
    Класс для работы с JSON базой данных продуктов и проверок

    Структура:
    - products.json - основной файл со всеми продуктами
    - checks.json - файл со всеми проверками
    - history/ - папка с историей изменений
    - backups/ - папка с резервными копиями
    - stats.json - статистика и метаданные
    - schema.json - схема данных
    """

    # Версия схемы данных
    SCHEMA_VERSION = "2.0"

    # Обязательные поля для продукта
    REQUIRED_FIELDS = ['name']

    # Допустимые поля (основные)
    ALLOWED_FIELDS = [
        'id', 'name', 'version', 'gk', 'description', 'subsystem',
        'certificate', 'platform', 'language', 'owner', 'module',
        'sheet_name', 'source_file', 'created_at', 'last_updated',
        'is_active', 'tags', 'notes'
    ]

    # Допустимые поля для проверки
    CHECK_FIELDS = [
        'id', 'name', 'group', 'type', 'aliases', 'without_aliases',
        'text', 'threshold', 'trust_threshold', 'description',
        'version_sections', 'required_total_indicators', 'strict_mode',
        'conditions', 'logic_operator', 'required_passed',
        'created_at', 'last_updated', 'is_enabled', 'is_deleted',
        'tags', 'notes', 'order'
    ]

    def __init__(self, base_dir: str = 'json_database'):
        """
        Инициализация JSON базы данных

        Args:
            base_dir: базовая директория для хранения файлов
        """
        self.base_dir = base_dir
        self.products_file = os.path.join(base_dir, 'products.json')
        self.checks_file = os.path.join(base_dir, 'checks.json')
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
        self.checks = self._load_checks()
        self.stats = self._load_stats()
        self.schema = self._load_schema()

        # Проверяем целостность данных
        self._validate_data()

        logger.info(f"JSON Database инициализирована в {base_dir}")
        logger.info(f"Загружено {len(self.products)} продуктов")
        logger.info(f"Загружено {len(self.checks)} проверок")

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
                'description': 'Схема данных для продуктов ПО и проверок',
                'product_fields': self.ALLOWED_FIELDS,
                'check_fields': self.CHECK_FIELDS,
                'product_required': self.REQUIRED_FIELDS
            }
            with open(self.schema_file, 'w', encoding='utf-8') as f:
                json.dump(schema, f, ensure_ascii=False, indent=2)
            logger.info(f"Создан файл схемы: {self.schema_file}")

        # products.json
        if not os.path.exists(self.products_file):
            with open(self.products_file, 'w', encoding='utf-8') as f:
                json.dump([], f, ensure_ascii=False, indent=2)
            logger.info(f"Создан файл продуктов: {self.products_file}")

        # checks.json
        if not os.path.exists(self.checks_file):
            # Создаем базовые проверки по умолчанию
            default_checks = self._get_default_checks()
            with open(self.checks_file, 'w', encoding='utf-8') as f:
                json.dump(default_checks, f, ensure_ascii=False, indent=2)
            logger.info(f"Создан файл проверок: {self.checks_file}")

        # stats.json
        if not os.path.exists(self.stats_file):
            initial_stats = {
                'total_products': 0,
                'total_checks': 0,
                'enabled_checks': 0,
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

    def _get_default_checks(self) -> List[Dict]:
        """Получение списка проверок по умолчанию"""
        return [
            {
                'id': 1,
                'name': 'Oracle',
                'group': 'Импортозамещение',
                'type': 'no_text_present',
                'aliases': ['Oracle', 'Oracle Database', 'Oracle DB', 'Oracle 11g', 'Oracle 12c', 'Oracle 19c'],
                'description': 'Проверка отсутствия упоминаний Oracle',
                'created_at': datetime.now().isoformat(),
                'last_updated': datetime.now().isoformat(),
                'is_enabled': True,
                'is_deleted': False,
                'order': 1
            },
            {
                'id': 2,
                'name': 'Запрещённое ПО',
                'group': 'Импортозамещение',
                'type': 'no_text_present',
                'aliases': ['Cisco', 'Juniper', 'Check Point', 'Palo Alto', 'Windows Server', 'Microsoft SQL', 'IBM',
                            'HP', 'Dell EMC'],
                'description': 'Проверка отсутствия запрещенного ПО',
                'created_at': datetime.now().isoformat(),
                'last_updated': datetime.now().isoformat(),
                'is_enabled': True,
                'is_deleted': False,
                'order': 2
            },
            {
                'id': 3,
                'name': 'Российское ПО',
                'group': 'Импортозамещение',
                'type': 'text_present',
                'aliases': ['Российское ПО', 'отечественное', 'реестр российского ПО', 'МойОфис', 'Астра Линукс',
                            'РЕД ОС'],
                'description': 'Проверка наличия упоминаний российского ПО',
                'created_at': datetime.now().isoformat(),
                'last_updated': datetime.now().isoformat(),
                'is_enabled': True,
                'is_deleted': False,
                'order': 3
            }
        ]

    def _load_products(self) -> List[Dict]:
        """Загрузка продуктов из JSON"""
        try:
            with open(self.products_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, list):
                    # Проверяем каждый продукт
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
            self._attempt_recovery('products')
            return []
        except FileNotFoundError:
            logger.error(f"Файл products.json не найден")
            return []

    def _load_checks(self) -> List[Dict]:
        """Загрузка проверок из JSON"""
        try:
            with open(self.checks_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if isinstance(data, list):
                    # Фильтруем только не удаленные проверки
                    active_checks = [check for check in data if not check.get('is_deleted', False)]
                    logger.info(f"Загружено {len(active_checks)} активных проверок из {len(data)} всего")
                    return data
                elif isinstance(data, dict):
                    # Старый формат конфигурации - конвертируем
                    logger.info("Обнаружен старый формат конфигурации, выполняем конвертацию")
                    converted = self._convert_old_config(data)
                    # Сохраняем сконвертированные данные
                    self.checks = converted
                    self._save_checks()
                    return converted
                else:
                    logger.error(f"checks.json должен содержать список или словарь, получен {type(data)}")
                    return []
        except json.JSONDecodeError as e:
            logger.error(f"Ошибка парсинга checks.json: {e}")
            self._attempt_recovery('checks')
            return []
        except FileNotFoundError:
            logger.info(f"Файл checks.json не найден, будут созданы проверки по умолчанию")
            return self._get_default_checks()

    def _convert_old_config(self, old_config: Dict) -> List[Dict]:
        """Конвертация старого формата конфигурации в новый"""
        new_checks = []
        check_id = 1

        # Проверяем разные возможные структуры старого формата
        checks_list = []

        # Формат 1: {'checks': [...]}
        if 'checks' in old_config and isinstance(old_config['checks'], list):
            checks_list = old_config['checks']
        # Формат 2: прямой список групп
        elif isinstance(old_config, dict) and any('group' in v for v in old_config.values() if isinstance(v, dict)):
            checks_list = [old_config]
        # Формат 3: список групп в корне
        elif isinstance(old_config, list):
            checks_list = old_config

        for group_item in checks_list:
            if not isinstance(group_item, dict):
                continue

            group_name = group_item.get('group', 'Без группы')
            subchecks = group_item.get('subchecks', [])

            if not isinstance(subchecks, list):
                continue

            for subcheck in subchecks:
                if not isinstance(subcheck, dict):
                    logger.warning(f"Пропущен элемент, не являющийся словарем: {subcheck}")
                    continue

                check_data = {
                    'id': check_id,
                    'name': subcheck.get('name', ''),
                    'group': group_name,
                    'type': subcheck.get('type', ''),
                    'description': subcheck.get('description', ''),
                    'created_at': datetime.now().isoformat(),
                    'last_updated': datetime.now().isoformat(),
                    'is_enabled': True,
                    'is_deleted': False,
                    'order': check_id
                }

                # Копируем специфичные поля
                if 'aliases' in subcheck and subcheck['aliases']:
                    if isinstance(subcheck['aliases'], list):
                        check_data['aliases'] = subcheck['aliases']
                    elif isinstance(subcheck['aliases'], str):
                        check_data['aliases'] = [a.strip() for a in subcheck['aliases'].split(',') if a.strip()]

                if 'without_aliases' in subcheck and subcheck['without_aliases']:
                    if isinstance(subcheck['without_aliases'], list):
                        check_data['without_aliases'] = subcheck['without_aliases']
                    elif isinstance(subcheck['without_aliases'], str):
                        check_data['without_aliases'] = [a.strip() for a in subcheck['without_aliases'].split(',') if
                                                         a.strip()]

                if 'text' in subcheck and subcheck['text']:
                    check_data['text'] = subcheck['text']

                if 'threshold' in subcheck:
                    try:
                        check_data['threshold'] = float(subcheck['threshold'])
                    except (ValueError, TypeError):
                        check_data['threshold'] = 70.0

                if 'trust_threshold' in subcheck:
                    try:
                        check_data['trust_threshold'] = float(subcheck['trust_threshold'])
                    except (ValueError, TypeError):
                        check_data['trust_threshold'] = 85.0

                if 'version_sections' in subcheck:
                    check_data['version_sections'] = subcheck['version_sections']

                if 'required_total_indicators' in subcheck:
                    check_data['required_total_indicators'] = subcheck['required_total_indicators']

                if 'strict_mode' in subcheck:
                    check_data['strict_mode'] = subcheck['strict_mode']

                if 'conditions' in subcheck:
                    check_data['conditions'] = subcheck['conditions']

                if 'logic_operator' in subcheck:
                    check_data['logic_operator'] = subcheck['logic_operator']

                if 'required_passed' in subcheck:
                    check_data['required_passed'] = subcheck['required_passed']

                new_checks.append(check_data)
                check_id += 1

        logger.info(f"Сконвертировано {len(new_checks)} проверок из старого формата")
        return new_checks

    def _load_stats(self) -> Dict:
        """Загрузка статистики из JSON"""
        try:
            with open(self.stats_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except (json.JSONDecodeError, FileNotFoundError) as e:
            logger.error(f"Ошибка загрузки stats.json: {e}")
            return {
                'total_products': 0,
                'total_checks': 0,
                'enabled_checks': 0,
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
                'product_fields': self.ALLOWED_FIELDS,
                'check_fields': self.CHECK_FIELDS,
                'product_required': self.REQUIRED_FIELDS
            }

    def _validate_product(self, product: Dict) -> bool:
        """Валидация структуры продукта"""
        try:
            for field in self.REQUIRED_FIELDS:
                if field not in product or not product.get(field):
                    logger.debug(f"Отсутствует обязательное поле: {field}")
                    return False
            if not isinstance(product.get('name', ''), str):
                logger.debug(f"Поле 'name' должно быть строкой")
                return False
            if 'gk' in product and product['gk'] is not None and not isinstance(product['gk'], list):
                logger.debug(f"Поле 'gk' должно быть списком")
                return False
            return True
        except Exception as e:
            logger.error(f"Ошибка валидации: {e}")
            return False

    def _validate_check(self, check: Dict) -> bool:
        """Валидация структуры проверки"""
        try:
            required_check_fields = ['id', 'name', 'group', 'type']
            for field in required_check_fields:
                if field not in check:
                    logger.debug(f"Отсутствует обязательное поле проверки: {field}")
                    return False
            if not isinstance(check.get('name', ''), str):
                logger.debug(f"Поле 'name' должно быть строкой")
                return False
            if 'is_enabled' in check and not isinstance(check['is_enabled'], bool):
                logger.debug(f"Поле 'is_enabled' должно быть boolean")
                return False
            return True
        except Exception as e:
            logger.error(f"Ошибка валидации проверки: {e}")
            return False

    def _clean_product(self, product: Dict) -> Optional[Dict]:
        """Очистка продукта от недопустимых полей"""
        try:
            cleaned = {}
            for key, value in product.items():
                if key in self.ALLOWED_FIELDS:
                    cleaned[key] = value
                elif key == 'row_index':
                    continue
                else:
                    logger.debug(f"Поле {key} удалено из продукта")
            for field in self.REQUIRED_FIELDS:
                if field not in cleaned or not cleaned.get(field):
                    logger.warning(f"Отсутствует обязательное поле {field} после очистки")
                    return None
            return cleaned
        except Exception as e:
            logger.error(f"Ошибка очистки продукта: {e}")
            return None

    def _validate_data(self):
        """Проверка целостности данных"""
        product_ids = [p.get('id') for p in self.products if p.get('id') is not None]
        if len(product_ids) != len(set(product_ids)):
            logger.warning("Обнаружены дублирующиеся ID продуктов!")
            self._fix_duplicate_product_ids()

        check_ids = [c.get('id') for c in self.checks if c.get('id') is not None]
        if len(check_ids) != len(set(check_ids)):
            logger.warning("Обнаружены дублирующиеся ID проверок!")
            self._fix_duplicate_check_ids()

        for i, product in enumerate(self.products):
            if 'id' not in product or product['id'] is None:
                product['id'] = self._generate_product_id()
                logger.info(f"Продукту на позиции {i} присвоен новый ID: {product['id']}")

        for i, check in enumerate(self.checks):
            if 'id' not in check or check['id'] is None:
                check['id'] = self._generate_check_id()
                logger.info(f"Проверке на позиции {i} присвоен новый ID: {check['id']}")

        if self.products or self.checks:
            self._save_all(create_backup=False)

    def _fix_duplicate_product_ids(self):
        """Исправление дублирующихся ID продуктов"""
        seen_ids = set()
        for product in self.products:
            pid = product.get('id')
            if pid in seen_ids:
                new_id = self._generate_product_id()
                logger.info(f"Изменение дублирующегося ID продукта {pid} на {new_id}")
                product['id'] = new_id
            else:
                if pid is not None:
                    seen_ids.add(pid)

    def _fix_duplicate_check_ids(self):
        """Исправление дублирующихся ID проверок"""
        seen_ids = set()
        for check in self.checks:
            cid = check.get('id')
            if cid in seen_ids:
                new_id = self._generate_check_id()
                logger.info(f"Изменение дублирующегося ID проверки {cid} на {new_id}")
                check['id'] = new_id
            else:
                if cid is not None:
                    seen_ids.add(cid)

    def _generate_product_id(self) -> int:
        """Генерация нового уникального ID продукта"""
        if not self.products:
            return 1
        max_id = 0
        for p in self.products:
            pid = p.get('id', 0)
            if isinstance(pid, (int, float)) and pid > max_id:
                max_id = int(pid)
        return max_id + 1

    def _generate_check_id(self) -> int:
        """Генерация нового уникального ID проверки"""
        if not self.checks:
            return 1
        max_id = 0
        for c in self.checks:
            cid = c.get('id', 0)
            if isinstance(cid, (int, float)) and cid > max_id:
                max_id = int(cid)
        return max_id + 1

    def _save_all(self, create_backup: bool = True):
        """Сохранение всех данных"""
        if create_backup:
            self.create_backup('before_save')
        self._save_products(create_backup=False)
        self._save_checks(create_backup=False)
        self._update_stats()
        self._save_stats()

    def _save_products(self, create_backup: bool = True):
        """Сохранение продуктов в JSON"""
        if create_backup:
            self.create_backup('before_save_products')
        temp_file = os.path.join(self.temp_dir, f'products_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json')
        try:
            cleaned_products = []
            for product in self.products:
                cleaned_product = {}
                for key, value in product.items():
                    if key in self.ALLOWED_FIELDS:
                        cleaned_product[key] = value
                cleaned_products.append(cleaned_product)
            with open(temp_file, 'w', encoding='utf-8') as f:
                json.dump(cleaned_products, f, ensure_ascii=False, indent=2, default=str)
            shutil.move(temp_file, self.products_file)
            logger.info(f"Продукты сохранены: {len(cleaned_products)} записей")
        except Exception as e:
            logger.error(f"Ошибка сохранения products.json: {e}")
            if os.path.exists(temp_file):
                os.remove(temp_file)
            raise JSONDatabaseError(f"Не удалось сохранить продукты: {e}")

    def _save_checks(self, create_backup: bool = True):
        """Сохранение проверок в JSON"""
        if create_backup:
            self.create_backup('before_save_checks')
        temp_file = os.path.join(self.temp_dir, f'checks_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json')
        try:
            cleaned_checks = []
            for check in self.checks:
                cleaned_check = {}
                for key, value in check.items():
                    if key in self.CHECK_FIELDS:
                        cleaned_check[key] = value
                cleaned_checks.append(cleaned_check)
            with open(temp_file, 'w', encoding='utf-8') as f:
                json.dump(cleaned_checks, f, ensure_ascii=False, indent=2, default=str)
            shutil.move(temp_file, self.checks_file)
            logger.info(f"Проверки сохранены: {len(cleaned_checks)} записей")
        except Exception as e:
            logger.error(f"Ошибка сохранения checks.json: {e}")
            if os.path.exists(temp_file):
                os.remove(temp_file)
            raise JSONDatabaseError(f"Не удалось сохранить проверки: {e}")

    def _save_stats(self):
        """Сохранение статистики в JSON"""
        try:
            with open(self.stats_file, 'w', encoding='utf-8') as f:
                json.dump(self.stats, f, ensure_ascii=False, indent=2, default=str)
        except Exception as e:
            logger.error(f"Ошибка сохранения stats.json: {e}")

    def _update_stats(self):
        """Обновление статистики"""
        total_products = len(self.products)
        total_checks = len(self.checks)
        enabled_checks = sum(1 for c in self.checks if c.get('is_enabled', True) and not c.get('is_deleted', False))

        by_subsystem = {}
        with_cert = 0
        active_products = 0

        for product in self.products:
            if product.get('is_active', True):
                active_products += 1
                subs = product.get('subsystem', 'Не определена')
                if subs:
                    by_subsystem[subs] = by_subsystem.get(subs, 0) + 1
                if product.get('certificate'):
                    with_cert += 1

        db_size = 0
        if os.path.exists(self.products_file):
            db_size += os.path.getsize(self.products_file)
        if os.path.exists(self.checks_file):
            db_size += os.path.getsize(self.checks_file)

        self.stats.update({
            'total_products': total_products,
            'total_checks': total_checks,
            'enabled_checks': enabled_checks,
            'active_products': active_products,
            'by_subsystem': by_subsystem,
            'with_certificate': with_cert,
            'without_certificate': active_products - with_cert,
            'last_update': datetime.now().isoformat(),
            'database_size': db_size
        })

    def _add_to_history(self, item_id: int, item_type: str, action: str, old_data: Optional[Dict] = None,
                        new_data: Optional[Dict] = None, changed_by: str = 'system', comment: str = ''):
        """Добавление записи в историю"""
        try:
            history_file = os.path.join(self.history_dir, f'{item_type}_{item_id}.json')
            history = []
            if os.path.exists(history_file):
                with open(history_file, 'r', encoding='utf-8') as f:
                    history = json.load(f)
            history.append({
                'item_type': item_type,
                'action': action,
                'old_data': old_data,
                'new_data': new_data,
                'changed_by': changed_by,
                'changed_at': datetime.now().isoformat(),
                'comment': comment
            })
            if len(history) > 100:
                history = history[-100:]
            with open(history_file, 'w', encoding='utf-8') as f:
                json.dump(history, f, ensure_ascii=False, indent=2, default=str)
            logger.debug(f"Добавлена запись в историю для {item_type} {item_id}: {action}")
        except Exception as e:
            logger.error(f"Ошибка сохранения истории для {item_type} {item_id}: {e}")

    def _attempt_recovery(self, item_type: str = 'all'):
        """Попытка восстановления данных из резервной копии"""
        logger.info(f"Попытка восстановления {item_type} из последней резервной копии")
        backups = self.get_backups_list()
        if backups:
            latest_backup = backups[0]
            backup_path = latest_backup.get('path')
            if backup_path and os.path.exists(backup_path):
                if item_type in ['products', 'all']:
                    products_backup = os.path.join(backup_path, 'products.json')
                    if os.path.exists(products_backup):
                        try:
                            with open(products_backup, 'r', encoding='utf-8') as f:
                                data = json.load(f)
                                if isinstance(data, list):
                                    self.products = data
                                    logger.info(f"Продукты восстановлены из резервной копии: {backup_path}")
                        except Exception as e:
                            logger.error(f"Ошибка восстановления продуктов из резервной копии: {e}")
                if item_type in ['checks', 'all']:
                    checks_backup = os.path.join(backup_path, 'checks.json')
                    if os.path.exists(checks_backup):
                        try:
                            with open(checks_backup, 'r', encoding='utf-8') as f:
                                data = json.load(f)
                                if isinstance(data, list):
                                    self.checks = data
                                    logger.info(f"Проверки восстановлены из резервной копии: {backup_path}")
                        except Exception as e:
                            logger.error(f"Ошибка восстановления проверок из резервной копии: {e}")

    def create_backup(self, reason: str = 'manual') -> str:
        """Создание резервной копии базы данных"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        backup_name = f'backup_{timestamp}_{reason}'
        backup_dir = os.path.join(self.backups_dir, backup_name)
        try:
            os.makedirs(backup_dir, exist_ok=True)
            if os.path.exists(self.products_file):
                shutil.copy2(self.products_file, os.path.join(backup_dir, 'products.json'))
            if os.path.exists(self.checks_file):
                shutil.copy2(self.checks_file, os.path.join(backup_dir, 'checks.json'))
            if os.path.exists(self.stats_file):
                shutil.copy2(self.stats_file, os.path.join(backup_dir, 'stats.json'))
            if os.path.exists(self.schema_file):
                shutil.copy2(self.schema_file, os.path.join(backup_dir, 'schema.json'))
            if os.path.exists(self.history_dir) and os.listdir(self.history_dir):
                shutil.copytree(self.history_dir, os.path.join(backup_dir, 'history'))
            metadata = {
                'created_at': timestamp,
                'reason': reason,
                'total_products': len(self.products),
                'total_checks': len(self.checks),
                'files': ['products.json', 'checks.json', 'stats.json', 'schema.json'],
                'has_history': os.path.exists(self.history_dir) and bool(os.listdir(self.history_dir))
            }
            with open(os.path.join(backup_dir, 'metadata.json'), 'w', encoding='utf-8') as f:
                json.dump(metadata, f, ensure_ascii=False, indent=2)
            self.stats['total_backups'] = len(self.get_backups_list())
            self.stats['last_backup'] = timestamp
            self._save_stats()
            logger.info(f"Создана резервная копия: {backup_dir}")
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
        """Восстановление из резервной копии"""
        try:
            if not os.path.exists(backup_dir):
                logger.error(f"Директория бэкапа не найдена: {backup_dir}")
                return False
            self.create_backup('before_restore')
            restored = False
            products_backup = os.path.join(backup_dir, 'products.json')
            if os.path.exists(products_backup):
                shutil.copy2(products_backup, self.products_file)
                restored = True
            checks_backup = os.path.join(backup_dir, 'checks.json')
            if os.path.exists(checks_backup):
                shutil.copy2(checks_backup, self.checks_file)
                restored = True
            stats_backup = os.path.join(backup_dir, 'stats.json')
            if os.path.exists(stats_backup):
                shutil.copy2(stats_backup, self.stats_file)
            schema_backup = os.path.join(backup_dir, 'schema.json')
            if os.path.exists(schema_backup):
                shutil.copy2(schema_backup, self.schema_file)
            history_backup = os.path.join(backup_dir, 'history')
            if os.path.exists(history_backup):
                if os.path.exists(self.history_dir) and os.listdir(self.history_dir):
                    current_history_backup = os.path.join(self.backups_dir,
                                                          f'history_before_restore_{datetime.now().strftime("%Y%m%d_%H%M%S")}')
                    shutil.copytree(self.history_dir, current_history_backup)
                if os.path.exists(self.history_dir):
                    shutil.rmtree(self.history_dir)
                shutil.copytree(history_backup, self.history_dir)
            if restored:
                self.products = self._load_products()
                self.checks = self._load_checks()
                self.stats = self._load_stats()
                self.schema = self._load_schema()
                logger.info(f"Восстановлено из бэкапа: {backup_dir}")
            return restored
        except Exception as e:
            logger.error(f"Ошибка восстановления из бэкапа: {e}")
            return False

    def get_backups_list(self) -> List[Dict]:
        """Получение списка доступных резервных копий"""
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
                        created = item.replace('backup_', '').split('_')[0] if item.startswith('backup_') else 'unknown'
                        backups.append({
                            'created_at': created,
                            'reason': 'unknown',
                            'path': backup_path,
                            'name': item,
                            'total_products': 'unknown',
                            'total_checks': 'unknown',
                            'files': []
                        })
            backups.sort(key=lambda x: x.get('created_at', ''), reverse=True)
        except Exception as e:
            logger.error(f"Ошибка получения списка бэкапов: {e}")
        return backups

    # ==================== МЕТОДЫ ДЛЯ РАБОТЫ С ПРОВЕРКАМИ ====================

    def _get_check_fingerprint(self, check: Dict) -> str:
        """
        Создает уникальный отпечаток проверки на основе её содержимого

        Args:
            check: данные проверки

        Returns:
            str: уникальный хеш проверки
        """
        # Копируем и очищаем от служебных полей
        check_copy = check.copy()
        check_copy.pop('id', None)
        check_copy.pop('created_at', None)
        check_copy.pop('last_updated', None)
        check_copy.pop('is_enabled', None)
        check_copy.pop('is_deleted', None)
        check_copy.pop('order', None)

        # Сортируем ключи для консистентности
        sorted_items = sorted(check_copy.items())

        # Создаем строковое представление
        content_str = json.dumps(sorted_items, sort_keys=True, default=str)

        # Создаем хеш
        return hashlib.md5(content_str.encode('utf-8')).hexdigest()

    def _compare_checks_content(self, check1: Dict, check2: Dict) -> Tuple[bool, List[str]]:
        """
        Сравнивает содержимое двух проверок и возвращает различия

        Args:
            check1: первая проверка
            check2: вторая проверка

        Returns:
            Tuple[bool, List[str]]: (одинаковые, список различий)
        """
        differences = []

        # Поля для сравнения (исключаем служебные)
        compare_fields = ['name', 'group', 'type', 'aliases', 'without_aliases', 'text',
                          'threshold', 'trust_threshold', 'description', 'version_sections',
                          'required_total_indicators', 'strict_mode', 'conditions',
                          'logic_operator', 'required_passed']

        for field in compare_fields:
            val1 = check1.get(field)
            val2 = check2.get(field)

            # Преобразуем списки в множества для сравнения
            if isinstance(val1, list) and isinstance(val2, list):
                set1 = set(str(v) for v in val1)
                set2 = set(str(v) for v in val2)
                if set1 != set2:
                    missing = set2 - set1
                    extra = set1 - set2
                    if missing:
                        differences.append(f"В поле '{field}' отсутствуют: {missing}")
                    if extra:
                        differences.append(f"В поле '{field}' лишние: {extra}")
            elif val1 != val2:
                differences.append(f"Поле '{field}': '{val1}' != '{val2}'")

        return len(differences) == 0, differences

    def _merge_checks(self, target: Dict, source: Dict) -> Dict:
        """
        Объединяет две проверки, добавляя недостающие параметры

        Args:
            target: целевая проверка (в которую добавляем)
            source: источник (из которого берем новые данные)

        Returns:
            Dict: объединенная проверка
        """
        merged = target.copy()
        changes = []

        # Поля для объединения
        merge_fields = ['aliases', 'without_aliases', 'conditions', 'version_sections']

        for field in merge_fields:
            target_val = target.get(field)
            source_val = source.get(field)

            if source_val and not target_val:
                # Если в целевой нет, а в источнике есть - добавляем
                merged[field] = source_val
                changes.append(f"Добавлено поле '{field}'")
            elif target_val and source_val:
                # Если есть в обоих, объединяем списки (для списков)
                if isinstance(target_val, list) and isinstance(source_val, list):
                    # Преобразуем в множества для уникальности
                    if field == 'aliases' or field == 'without_aliases':
                        # Для алиасов объединяем и убираем дубликаты
                        combined = list(set(target_val + source_val))
                        if len(combined) > len(target_val):
                            merged[field] = combined
                            changes.append(
                                f"В поле '{field}' добавлено {len(combined) - len(target_val)} новых элементов")
                elif field == 'conditions' and isinstance(target_val, list) and isinstance(source_val, list):
                    # Для условий - добавляем только уникальные по названию
                    existing_names = {c.get('name') for c in target_val if c.get('name')}
                    new_conditions = [c for c in source_val if c.get('name') not in existing_names]
                    if new_conditions:
                        merged[field] = target_val + new_conditions
                        changes.append(f"В поле '{field}' добавлено {len(new_conditions)} новых условий")

        # Для текстовых полей - если в целевой нет, а в источнике есть, добавляем
        text_fields = ['text', 'description']
        for field in text_fields:
            if not target.get(field) and source.get(field):
                merged[field] = source.get(field)
                changes.append(f"Добавлено поле '{field}'")

        # Для числовых полей - если в целевой нет, а в источнике есть, добавляем
        numeric_fields = ['threshold', 'trust_threshold', 'required_total_indicators', 'required_passed']
        for field in numeric_fields:
            if not target.get(field) and source.get(field) is not None:
                merged[field] = source.get(field)
                changes.append(f"Добавлено поле '{field}'")

        # Для булевых полей
        bool_fields = ['strict_mode']
        for field in bool_fields:
            if field not in target and field in source:
                merged[field] = source.get(field)
                changes.append(f"Добавлено поле '{field}'")

        if changes:
            logger.info(f"Объединение проверки '{target.get('name')}': {', '.join(changes)}")
            merged['last_updated'] = datetime.now().isoformat()

        return merged

    def find_matching_checks(self, check: Dict, threshold: float = 0.7) -> List[Tuple[Dict, float, List[str]]]:
        """
        Поиск проверок, похожих на данную по содержимому

        Args:
            check: проверка для поиска
            threshold: порог схожести (0-1)

        Returns:
            List[Tuple[Dict, float, List[str]]]: список (проверка, оценка схожести, различия)
        """
        from difflib import SequenceMatcher

        matches = []
        check_name = check.get('name', '').lower()
        check_group = check.get('group', '').lower()
        check_type = check.get('type', '')

        for existing in self.checks:
            if existing.get('is_deleted', False):
                continue

            existing_name = existing.get('name', '').lower()
            existing_group = existing.get('group', '').lower()
            existing_type = existing.get('type', '')

            # Сначала проверяем по названию и группе
            name_similarity = SequenceMatcher(None, check_name, existing_name).ratio()
            group_match = check_group == existing_group

            # Если тип разный - скорее всего это разные проверки
            if check_type != existing_type:
                continue

            # Оценка схожести
            similarity_score = 0.0

            if name_similarity > 0.8 and group_match:
                # Очень похожие названия и та же группа
                similarity_score = 0.9
            elif name_similarity > 0.6 and group_match:
                # Похожие названия и та же группа
                similarity_score = 0.7
            elif name_similarity > 0.8:
                # Очень похожие названия но другая группа
                similarity_score = 0.5
            else:
                # Проверяем по содержимому
                fingerprint1 = self._get_check_fingerprint(check)
                fingerprint2 = self._get_check_fingerprint(existing)
                if fingerprint1 == fingerprint2:
                    similarity_score = 1.0

            if similarity_score >= threshold:
                # Получаем детальные различия
                is_same, differences = self._compare_checks_content(check, existing)
                matches.append((existing, similarity_score, differences))

        # Сортируем по убыванию схожести
        matches.sort(key=lambda x: x[1], reverse=True)
        return matches

    def add_or_update_check(self, check_data: Dict, auto_merge: bool = True) -> Tuple[Optional[int], str, List[str]]:
        """
        Добавляет новую проверку или обновляет существующую если найдено совпадение

        Args:
            check_data: данные проверки
            auto_merge: автоматически объединять с существующей

        Returns:
            Tuple[Optional[int], str, List[str]]: (ID, действие, сообщения)
        """
        try:
            # Ищем точное совпадение по названию и группе
            exact_matches = []
            for check in self.checks:
                if check.get('is_deleted', False):
                    continue
                if (check.get('name', '').lower() == check_data.get('name', '').lower() and
                        check.get('group', '').lower() == check_data.get('group', '').lower()):
                    exact_matches.append(check)

            if exact_matches:
                # Найдено точное совпадение по названию и группе
                existing = exact_matches[0]

                if auto_merge:
                    # Объединяем проверки
                    merged = self._merge_checks(existing, check_data)
                    if self.update_check(existing['id'], merged, skip_duplicate_check=True):
                        return existing['id'], "merged", ["Проверка обновлена (объединение)"]
                    else:
                        return existing['id'], "update_failed", ["Ошибка обновления проверки"]
                else:
                    return existing['id'], "exists", ["Проверка уже существует"]

            # Ищем похожие по содержимому
            matches = self.find_matching_checks(check_data)

            if matches and auto_merge:
                # Найдены похожие проверки
                best_match, score, differences = matches[0]

                if score > 0.9:
                    # Очень похожие - объединяем
                    merged = self._merge_checks(best_match, check_data)
                    if self.update_check(best_match['id'], merged, skip_duplicate_check=True):
                        return best_match['id'], "merged_similar", [
                            f"Объединено с похожей проверкой (схожесть: {score:.2f})"]
                elif score > 0.7:
                    # Похожие, но не очень - спрашиваем пользователя
                    # В автоматическом режиме создаем новую
                    pass

            # Не найдено совпадений - создаем новую
            check_id = self._generate_check_id()
            new_check = check_data.copy()
            new_check['id'] = check_id
            new_check['created_at'] = datetime.now().isoformat()
            new_check['last_updated'] = datetime.now().isoformat()
            new_check['is_enabled'] = True
            new_check['is_deleted'] = False

            if 'order' not in new_check:
                max_order = max([c.get('order', 0) for c in self.checks], default=0)
                new_check['order'] = max_order + 1

            if not self._validate_check(new_check):
                return None, "invalid", ["Невалидная структура проверки"]

            self.checks.append(new_check)
            self._save_checks()
            self._add_to_history(
                check_id,
                'check',
                'CREATE',
                None,
                new_check,
                'manual',
                f"Создана проверка: {new_check.get('name')}"
            )
            self._update_stats()
            self._save_stats()

            return check_id, "created", ["Проверка успешно создана"]

        except Exception as e:
            logger.error(f"Ошибка добавления/обновления проверки: {e}")
            return None, "error", [f"Ошибка: {str(e)}"]

    def import_checks_from_yaml(self, file_path: str, auto_merge: bool = True) -> Tuple[int, int, int, List[Dict]]:
        """
        Импорт проверок из YAML файла с умным объединением

        Args:
            file_path: путь к файлу
            auto_merge: автоматически объединять с существующими

        Returns:
            Tuple[int, int, int, List[Dict]]: (добавлено, обновлено, пропущено, список деталей)
        """
        added = 0
        updated = 0
        skipped = 0
        details = []

        try:
            import yaml

            with open(file_path, 'r', encoding='utf-8') as f:
                data = yaml.safe_load(f)

            if data is None:
                logger.error("Файл пуст")
                return added, updated, skipped, details

            checks_to_process = []

            # Разбираем различные форматы
            if isinstance(data, dict):
                if 'checks' in data and isinstance(data['checks'], list):
                    # Формат: {'checks': [...]}
                    for group_item in data['checks']:
                        if not isinstance(group_item, dict):
                            continue
                        group_name = group_item.get('group', 'Без группы')
                        subchecks = group_item.get('subchecks', [])
                        for check in subchecks:
                            if isinstance(check, dict):
                                check['group'] = group_name
                                checks_to_process.append(check)
                elif any(isinstance(v, list) for v in data.values()):
                    # Формат: {group_name: [...]}
                    for group_name, subchecks in data.items():
                        if isinstance(subchecks, list):
                            for check in subchecks:
                                if isinstance(check, dict):
                                    check['group'] = group_name
                                    checks_to_process.append(check)
                else:
                    # Одиночная проверка
                    if 'name' in data:
                        checks_to_process.append(data)

            elif isinstance(data, list):
                # Формат: список проверок или список групп
                for item in data:
                    if isinstance(item, dict):
                        if 'group' in item and 'subchecks' in item:
                            # Это группа
                            group_name = item.get('group', 'Без группы')
                            for check in item.get('subchecks', []):
                                if isinstance(check, dict):
                                    check['group'] = group_name
                                    checks_to_process.append(check)
                        elif 'name' in item:
                            # Это проверка
                            if 'group' not in item:
                                item['group'] = 'Без группы'
                            checks_to_process.append(item)

            logger.info(f"Найдено {len(checks_to_process)} проверок для обработки")

            # Обрабатываем каждую проверку
            for check in checks_to_process:
                try:
                    # Очищаем от служебных полей
                    check.pop('id', None)
                    check.pop('created_at', None)
                    check.pop('last_updated', None)

                    # Добавляем или обновляем
                    check_id, action, messages = self.add_or_update_check(check, auto_merge)

                    if action == "created":
                        added += 1
                        details.append({
                            'check': check,
                            'action': 'created',
                            'id': check_id,
                            'messages': messages
                        })
                    elif action in ["merged", "merged_similar", "updated"]:
                        updated += 1
                        details.append({
                            'check': check,
                            'action': 'updated',
                            'id': check_id,
                            'messages': messages
                        })
                    elif action == "exists":
                        skipped += 1
                        details.append({
                            'check': check,
                            'action': 'skipped',
                            'reason': 'already_exists',
                            'messages': messages
                        })
                    else:
                        skipped += 1
                        details.append({
                            'check': check,
                            'action': 'skipped',
                            'reason': action,
                            'messages': messages
                        })

                except Exception as e:
                    logger.error(f"Ошибка обработки проверки {check.get('name', 'unknown')}: {e}")
                    skipped += 1
                    details.append({
                        'check': check,
                        'action': 'error',
                        'reason': str(e)
                    })

            logger.info(f"Импорт завершен: добавлено {added}, обновлено {updated}, пропущено {skipped}")
            return added, updated, skipped, details

        except Exception as e:
            logger.error(f"Ошибка импорта из YAML: {e}")
            return added, updated, skipped, details

    def check_for_duplicate(self, name: str, group: str, exclude_id: Optional[int] = None) -> Tuple[
        bool, Optional[Dict]]:
        """Проверка на наличие дубликата проверки"""
        for check in self.checks:
            if check.get('is_deleted', False):
                continue
            if exclude_id and check.get('id') == exclude_id:
                continue
            if check.get('name', '').lower() == name.lower() and check.get('group', '').lower() == group.lower():
                return True, check
        return False, None

    def add_check(self, check_data: Dict, skip_duplicate_check: bool = False) -> Optional[int]:
        """Добавление новой проверки"""
        try:
            if not skip_duplicate_check:
                name = check_data.get('name', '')
                group = check_data.get('group', '')
                is_duplicate, duplicate = self.check_for_duplicate(name, group)
                if is_duplicate:
                    logger.warning(f"Обнаружен дубликат проверки: {name} в группе {group}")
                    raise JSONDatabaseError(f"Проверка с названием '{name}' уже существует в группе '{group}'")

            check_id = self._generate_check_id()
            new_check = check_data.copy()
            new_check['id'] = check_id
            new_check['created_at'] = datetime.now().isoformat()
            new_check['last_updated'] = datetime.now().isoformat()
            new_check['is_enabled'] = True
            new_check['is_deleted'] = False

            if 'order' not in new_check:
                max_order = max([c.get('order', 0) for c in self.checks], default=0)
                new_check['order'] = max_order + 1

            if not self._validate_check(new_check):
                logger.error(f"Невалидная структура проверки: {new_check}")
                return None

            self.checks.append(new_check)
            self._save_checks()
            self._add_to_history(
                check_id,
                'check',
                'CREATE',
                None,
                new_check,
                'manual',
                f"Создана проверка: {new_check.get('name')}"
            )
            self._update_stats()
            self._save_stats()

            logger.info(f"Добавлена проверка ID {check_id}: {check_data.get('name', '')}")
            return check_id

        except JSONDatabaseError as e:
            raise e
        except Exception as e:
            logger.error(f"Ошибка добавления проверки: {e}")
            return None

    def update_check(self, check_id: int, updates: Dict, skip_duplicate_check: bool = False) -> bool:
        """Обновление проверки"""
        try:
            for i, check in enumerate(self.checks):
                if check.get('id') == check_id:
                    if not skip_duplicate_check:
                        new_name = updates.get('name', check.get('name'))
                        new_group = updates.get('group', check.get('group'))
                        if new_name != check.get('name') or new_group != check.get('group'):
                            is_duplicate, duplicate = self.check_for_duplicate(new_name, new_group, exclude_id=check_id)
                            if is_duplicate:
                                logger.warning(f"Обнаружен дубликат при обновлении: {new_name} в группе {new_group}")
                                raise JSONDatabaseError(
                                    f"Проверка с названием '{new_name}' уже существует в группе '{new_group}'")

                    old_data = check.copy()
                    for key, value in updates.items():
                        if key in self.CHECK_FIELDS and key != 'id' and key != 'created_at':
                            check[key] = value
                    check['last_updated'] = datetime.now().isoformat()

                    if not self._validate_check(check):
                        logger.error(f"Проверка ID {check_id} стала невалидной после обновления")
                        self.checks[i] = old_data
                        return False

                    self._save_checks()
                    self._add_to_history(
                        check_id,
                        'check',
                        'UPDATE',
                        old_data,
                        check,
                        'manual',
                        f"Обновлена проверка: {check.get('name')}"
                    )
                    self._update_stats()
                    self._save_stats()

                    logger.info(f"Обновлена проверка ID {check_id}")
                    return True

            logger.warning(f"Проверка ID {check_id} не найдена")
            return False

        except JSONDatabaseError as e:
            raise e
        except Exception as e:
            logger.error(f"Ошибка обновления проверки: {e}")
            return False

    def delete_check(self, check_id: int, hard_delete: bool = False) -> bool:
        """Удаление проверки"""
        try:
            if hard_delete:
                for i, check in enumerate(self.checks):
                    if check.get('id') == check_id:
                        old_data = check.copy()
                        del self.checks[i]
                        self._save_checks()
                        self._add_to_history(
                            check_id,
                            'check',
                            'HARD_DELETE',
                            old_data,
                            None,
                            'manual',
                            f"Полное удаление проверки: {old_data.get('name')}"
                        )
                        self._update_stats()
                        self._save_stats()
                        logger.info(f"Полное удаление проверки ID {check_id}")
                        return True
            else:
                return self.update_check(check_id, {'is_deleted': True, 'is_enabled': False})
            return False
        except Exception as e:
            logger.error(f"Ошибка удаления проверки: {e}")
            return False

    def restore_check(self, check_id: int) -> bool:
        """Восстановление мягко удаленной проверки"""
        try:
            for check in self.checks:
                if check.get('id') == check_id and check.get('is_deleted', False):
                    return self.update_check(check_id, {'is_deleted': False})
            logger.warning(f"Проверка ID {check_id} не найдена или не удалена")
            return False
        except Exception as e:
            logger.error(f"Ошибка восстановления проверки: {e}")
            return False

    def enable_check(self, check_id: int, enabled: bool = True) -> bool:
        """Включение/выключение проверки"""
        return self.update_check(check_id, {'is_enabled': enabled})

    def toggle_check(self, check_id: int) -> bool:
        """Переключение состояния проверки"""
        check = self.get_check(check_id)
        if check:
            return self.enable_check(check_id, not check.get('is_enabled', True))
        return False

    def get_check(self, check_id: int) -> Optional[Dict]:
        """Получение проверки по ID"""
        for check in self.checks:
            if check.get('id') == check_id:
                return check.copy()
        return None

    def get_all_checks(self, include_disabled: bool = True, include_deleted: bool = False) -> List[Dict]:
        """Получение всех проверок"""
        results = []
        for check in self.checks:
            if not include_deleted and check.get('is_deleted', False):
                continue
            if not include_disabled and not check.get('is_enabled', True):
                continue
            results.append(check.copy())
        results.sort(key=lambda x: (x.get('order', 999), x.get('name', '')))
        return results

    def get_enabled_checks(self) -> List[Dict]:
        """Получение только включенных проверок"""
        return self.get_all_checks(include_disabled=False, include_deleted=False)

    def find_checks(self,
                    group: Optional[str] = None,
                    search: Optional[str] = None,
                    check_type: Optional[str] = None,
                    only_enabled: bool = False,
                    include_deleted: bool = False) -> List[Dict]:
        """Поиск проверок с фильтрацией"""
        results = []
        for check in self.checks:
            if not include_deleted and check.get('is_deleted', False):
                continue
            if only_enabled and not check.get('is_enabled', True):
                continue
            if group and group != 'Все группы':
                if check.get('group') != group:
                    continue
            if check_type and check_type != 'Все типы':
                if check.get('type') != check_type:
                    continue
            if search:
                search_lower = search.lower()
                name = check.get('name', '').lower()
                description = check.get('description', '').lower() if check.get('description') else ''
                if search_lower not in name and search_lower not in description:
                    continue
            results.append(check.copy())
        results.sort(key=lambda x: (x.get('order', 999), x.get('name', '')))
        return results

    def get_groups(self) -> List[str]:
        """Получение списка групп проверок"""
        groups = set()
        for check in self.checks:
            if not check.get('is_deleted', False):
                group = check.get('group', 'Без группы')
                groups.add(group)
        return sorted(list(groups))

    def get_check_types(self) -> List[str]:
        """Получение списка типов проверок"""
        types = set()
        for check in self.checks:
            if not check.get('is_deleted', False):
                check_type = check.get('type', '')
                if check_type:
                    types.add(check_type)
        return sorted(list(types))

    def get_check_history(self, check_id: int) -> List[Dict]:
        """Получение истории изменений проверки"""
        history_file = os.path.join(self.history_dir, f'check_{check_id}.json')
        if not os.path.exists(history_file):
            return []
        try:
            with open(history_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            logger.error(f"Ошибка загрузки истории для проверки {check_id}: {e}")
            return []

    def reorder_checks(self, check_ids: List[int]) -> bool:
        """Изменение порядка проверок"""
        try:
            check_dict = {c['id']: c for c in self.checks}
            for order, check_id in enumerate(check_ids, 1):
                if check_id in check_dict:
                    check_dict[check_id]['order'] = order
            self._save_checks()
            return True
        except Exception as e:
            logger.error(f"Ошибка изменения порядка проверок: {e}")
            return False

    # ==================== МЕТОДЫ ДЛЯ ЭКСПОРТА/ИМПОРТА ====================

    def export_checks_to_yaml(self, file_path: str) -> bool:
        """Экспорт проверок в YAML файл"""
        try:
            import yaml
            active_checks = self.get_all_checks(include_disabled=True, include_deleted=False)
            grouped_checks = {}
            for check in active_checks:
                group = check.get('group', 'Без группы')
                if group not in grouped_checks:
                    grouped_checks[group] = []
                grouped_checks[group].append(check)
            export_data = {
                'checks': [
                    {
                        'group': group,
                        'subchecks': checks
                    }
                    for group, checks in grouped_checks.items()
                ]
            }
            with open(file_path, 'w', encoding='utf-8') as f:
                yaml.dump(export_data, f, allow_unicode=True, default_flow_style=False, indent=2)
            logger.info(f"Проверки экспортированы в YAML: {file_path}")
            return True
        except Exception as e:
            logger.error(f"Ошибка экспорта в YAML: {e}")
            return False

    # ==================== МЕТОДЫ ДЛЯ ПРОДУКТОВ ====================

    def add_product(self, product: Dict) -> Optional[int]:
        """Добавление нового продукта"""
        try:
            cleaned_product = self._clean_product(product)
            if not cleaned_product:
                logger.error(f"Не удалось очистить продукт: {product.get('name', 'unknown')}")
                return None
            if not self._validate_product(cleaned_product):
                logger.error(f"Невалидная структура продукта после очистки: {cleaned_product}")
                return None
            product_id = self._generate_product_id()
            new_product = cleaned_product.copy()
            new_product['id'] = product_id
            new_product['created_at'] = datetime.now().isoformat()
            new_product['last_updated'] = datetime.now().isoformat()
            new_product['is_active'] = True
            if 'gk' not in new_product or new_product['gk'] is None:
                new_product['gk'] = []
            if 'subsystem' not in new_product or not new_product['subsystem']:
                new_product['subsystem'] = 'Не определена'
            if 'tags' not in new_product:
                new_product['tags'] = []
            self.products.append(new_product)
            self._save_products()
            self._add_to_history(
                product_id,
                'product',
                'CREATE',
                None,
                new_product,
                'manual',
                f"Создан продукт: {new_product.get('name')}"
            )
            self._update_stats()
            self._save_stats()
            logger.info(f"Добавлен продукт ID {product_id}: {product.get('name', '')}")
            return product_id
        except Exception as e:
            logger.error(f"Ошибка добавления продукта: {e}")
            return None

    def get_product(self, product_id: int) -> Optional[Dict]:
        """Получение продукта по ID"""
        for product in self.products:
            if product.get('id') == product_id:
                return product.copy()
        return None

    def get_all_products(self, include_inactive: bool = False) -> List[Dict]:
        """Получение всех продуктов"""
        if include_inactive:
            return [p.copy() for p in self.products]
        return [p.copy() for p in self.products if p.get('is_active', True)]

    def get_subsystems(self) -> List[str]:
        """Получение списка подсистем"""
        subsystems = set()
        for product in self.products:
            if product.get('is_active', True):
                subs = product.get('subsystem', 'Не определена')
                if subs and subs != 'Не определена':
                    subsystems.add(subs)
        has_undefined = any(p.get('subsystem') == 'Не определена' for p in self.products if p.get('is_active', True))
        if has_undefined:
            subsystems.add('Не определена')
        return sorted(list(subsystems))

    def get_stats(self) -> Dict:
        """Получение статистики"""
        return self.stats.copy()

    def validate_database(self) -> Dict[str, Any]:
        """Проверка целостности базы данных"""
        issues = []
        stats = {
            'total_products': len(self.products),
            'total_checks': len(self.checks),
            'valid_products': 0,
            'valid_checks': 0,
            'invalid_products': 0,
            'invalid_checks': 0,
            'duplicate_product_ids': [],
            'duplicate_check_ids': [],
            'duplicate_check_names': [],
            'missing_product_ids': [],
            'missing_check_ids': [],
            'validation_errors': []
        }

        seen_product_ids = set()
        for i, product in enumerate(self.products):
            pid = product.get('id')
            if pid is None:
                stats['missing_product_ids'].append(i)
                issues.append(f"Продукт на позиции {i} не имеет ID")
                continue
            if pid in seen_product_ids:
                stats['duplicate_product_ids'].append(pid)
                issues.append(f"Дублирующийся ID продукта: {pid}")
            else:
                seen_product_ids.add(pid)
            if self._validate_product(product):
                stats['valid_products'] += 1
            else:
                stats['invalid_products'] += 1
                issues.append(f"Невалидная структура у продукта ID {pid}")

        seen_check_ids = set()
        seen_check_names = set()
        for i, check in enumerate(self.checks):
            cid = check.get('id')
            if cid is None:
                stats['missing_check_ids'].append(i)
                issues.append(f"Проверка на позиции {i} не имеет ID")
                continue
            if cid in seen_check_ids:
                stats['duplicate_check_ids'].append(cid)
                issues.append(f"Дублирующийся ID проверки: {cid}")
            else:
                seen_check_ids.add(cid)
            name = check.get('name', '')
            group = check.get('group', '')
            name_key = f"{name.lower()}|{group.lower()}"
            if name_key in seen_check_names and not check.get('is_deleted', False):
                stats['duplicate_check_names'].append(name_key)
                issues.append(f"Дублирующееся название проверки: '{name}' в группе '{group}'")
            else:
                seen_check_names.add(name_key)
            if self._validate_check(check):
                stats['valid_checks'] += 1
            else:
                stats['invalid_checks'] += 1
                issues.append(f"Невалидная структура у проверки ID {cid}")

        stats['issues'] = issues
        stats['is_valid'] = len(issues) == 0
        return stats

    def repair_database(self) -> bool:
        """Попытка восстановления/репарации базы данных"""
        logger.info("Запуск процедуры восстановления базы данных")
        try:
            backup_path = self.create_backup('before_repair')
            validation = self.validate_database()
            if validation['is_valid']:
                logger.info("База данных не требует восстановления")
                return True
            if validation['duplicate_product_ids']:
                self._fix_duplicate_product_ids()
            if validation['duplicate_check_ids']:
                self._fix_duplicate_check_ids()
            for idx in validation['missing_product_ids']:
                if idx < len(self.products):
                    self.products[idx]['id'] = self._generate_product_id()
                    logger.info(f"Добавлен ID для продукта на позиции {idx}")
            for idx in validation['missing_check_ids']:
                if idx < len(self.checks):
                    self.checks[idx]['id'] = self._generate_check_id()
                    logger.info(f"Добавлен ID для проверки на позиции {idx}")
            cleaned_products = []
            for product in self.products:
                cleaned = self._clean_product(product)
                if cleaned:
                    cleaned_products.append(cleaned)
                else:
                    logger.warning(f"Продукт {product.get('id', 'unknown')} удален из-за невалидной структуры")
            self.products = cleaned_products
            self._save_all()
            logger.info("База данных восстановлена")
            return True
        except Exception as e:
            logger.error(f"Ошибка восстановления базы данных: {e}")
            return False

    def get_database_info(self) -> Dict[str, Any]:
        """Получение подробной информации о базе данных"""
        info = {
            'path': self.base_dir,
            'files': {
                'products': os.path.exists(self.products_file),
                'checks': os.path.exists(self.checks_file),
                'stats': os.path.exists(self.stats_file),
                'schema': os.path.exists(self.schema_file)
            },
            'sizes': {},
            'stats': self.stats,
            'history_count': len(os.listdir(self.history_dir)) if os.path.exists(self.history_dir) else 0,
            'backup_count': len(self.get_backups_list()),
            'validation': self.validate_database()
        }
        for name, path in [('products', self.products_file),
                           ('checks', self.checks_file),
                           ('stats', self.stats_file),
                           ('schema', self.schema_file)]:
            if os.path.exists(path):
                info['sizes'][name] = os.path.getsize(path)
        return info