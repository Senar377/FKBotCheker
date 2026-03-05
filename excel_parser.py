# excel_parser.py
"""
Модуль для парсинга Excel файлов с составом ПО
Извлекает данные из структурированных таблиц по листам
"""

import logging
import json
import os
import re
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
import pandas as pd
import numpy as np
import openpyxl
from openpyxl import load_workbook


# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('excel_parser.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger('ExcelParser')


class ExcelParserError(Exception):
    """Базовое исключение для ошибок парсинга Excel"""
    pass


class ExcelFileNotFoundError(ExcelParserError):
    """Файл Excel не найден"""
    pass


class ExcelSheetNotFoundError(ExcelParserError):
    """Лист Excel не найден"""
    pass


class ExcelParseError(ExcelParserError):
    """Ошибка при парсинге Excel"""
    pass


class ExcelParser:
    """
    Класс для парсинга Excel файлов с составом ПО
    Извлекает данные из структурированных таблиц по листам
    """

    # Ожидаемые листы в файле
    TARGET_SHEETS = {
        'ГМП': 'ГМП',
        'ГАСУ': 'ГАСУ',
        'ЭБ ЗК': 'ЭБ ЗК (256ГК)',
        'ПОИ': 'ПОИ',
        'ПУДС': 'ПУДС',
        'ПУиО': 'ПУиО',
        'ПИАО': 'ПИАО',
        'ЕПБС': 'ЕПБС',
        'ПУР': 'ПУР',
        'НСИ': 'НСИ',
        'Лицензии': 'Лицензии'
    }

    # Специфические колонки для разных листов
    SHEET_COLUMNS = {
        'ГМП': {
            'bpo': {'name_col': 1, 'version_col': 2, 'desc_col': 4, 'gk_col': 5},
            'spo': {'name_col': 9, 'version_col': 10, 'desc_col': 13, 'gk_col': 14}
        },
        'ГАСУ': {
            'bpo': {'name_col': 1, 'version_col': 2, 'desc_col': 4, 'gk_col': 5},
            'spo': {'name_col': 9, 'version_col': 10, 'desc_col': 13, 'gk_col': 14},
            'mp': {'prefix': '+МП', 'name_col': 9, 'version_col': 10, 'gk_col': 14},
            'eso': {'prefix': 'ЕСО', 'name_col': 9, 'version_col': 10, 'gk_col': 14}
        },
        'ЭБ ЗК': {
            'items': {'name_col': 2, 'version_col': 3, 'desc_col': 4, 'type_col': 1}
        },
        'ПОИ': {
            'bpo': {'name_col': 1, 'version_col': 2, 'desc_col': 4, 'gk_col': 5},
            'spo': {'name_col': 9, 'version_col': 10, 'desc_col': 13, 'gk_col': 14}
        },
        'ПУДС': {
            'items': {'name_col': 1, 'version_col': 2, 'desc_col': 4, 'gk_col': 5, 'module_col': 0}
        },
        'ПУиО': {
            'bpo': {'name_col': 1, 'version_col': 2, 'desc_col': 4, 'gk_col': 5},
            'spo': {'name_col': 9, 'version_col': 10, 'desc_col': 13, 'gk_col': 14}
        },
        'ПИАО': {
            'bpo': {'name_col': 1, 'version_col': 2, 'desc_col': 4, 'gk_col': 5},
            'spo': {'name_col': 9, 'version_col': 10, 'desc_col': 13, 'gk_col': 14}
        },
        'ЕПБС': {
            'bpo': {'name_col': 1, 'version_col': 2, 'desc_col': 4, 'gk_col': 5},
            'spo': {'name_col': 8, 'version_col': 9, 'desc_col': 12, 'gk_col': 13}
        },
        'ПУР': {
            'bpo': {'name_col': 1, 'version_col': 2, 'desc_col': 4, 'gk_col': 5},
            'spo': {'name_col': 8, 'version_col': 9, 'desc_col': 12, 'gk_col': 13, 'comment_col': 15}
        },
        'НСИ': {
            'bpo': {'name_col': 1, 'version_col': 2, 'desc_col': 4, 'gk_col': 5},
            'spo': {'name_col': 8, 'version_col': 9, 'desc_col': 12, 'gk_col': 13, 'comment_col': 15}
        },
        'Лицензии': {
            'items': {'name_col': 0, 'correct_name_col': 1, 'owner_col': 2, 'desc_col': 3}
        }
    }

    def __init__(self, json_db=None):
        """
        Инициализация парсера

        Args:
            json_db: экземпляр JSONDatabase для сохранения результатов
        """
        self.json_db = json_db
        self.current_file_path = None
        self.parsed_data = {}
        self.stats = {
            'total_products': 0,
            'by_subsystem': {},
            'by_sheet': {},
            'last_parse': None,
            'errors': []
        }

        logger.info("ExcelParser инициализирован")

    def parse_excel_file(self, file_path: str) -> Dict[str, List[Dict]]:
        """
        Парсинг Excel файла

        Args:
            file_path: путь к Excel файлу

        Returns:
            Dict: словарь с данными по листам

        Raises:
            ExcelFileNotFoundError: если файл не найден
            ExcelParseError: если ошибка при парсинге
        """
        logger.info(f"Начало парсинга файла: {file_path}")

        if not os.path.exists(file_path):
            error_msg = f"Файл не найден: {file_path}"
            logger.error(error_msg)
            self.stats['errors'].append(error_msg)
            raise ExcelFileNotFoundError(error_msg)

        self.current_file_path = file_path
        result = {}

        try:
            # Читаем все листы
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names

            logger.info(f"Найдены листы: {sheet_names}")

            # Парсим целевые листы
            total_products = 0

            for subsystem, sheet_pattern in self.TARGET_SHEETS.items():
                # Находим соответствующий лист
                target_sheet = None
                for sheet_name in sheet_names:
                    if sheet_pattern.lower() in sheet_name.lower():
                        target_sheet = sheet_name
                        break

                if target_sheet:
                    try:
                        sheet_data = self._parse_sheet(excel_file, target_sheet, subsystem)
                        if sheet_data:
                            result[subsystem] = sheet_data
                            count = len(sheet_data)
                            total_products += count
                            self.stats['by_sheet'][subsystem] = count
                            self.stats['by_subsystem'][subsystem] = count
                            logger.info(f"Лист {target_sheet} -> {subsystem}: обработано {count} продуктов")
                    except Exception as e:
                        error_msg = f"Ошибка парсинга листа {target_sheet}: {e}"
                        logger.error(error_msg, exc_info=True)
                        self.stats['errors'].append(error_msg)
                else:
                    logger.warning(f"Лист для подсистемы {subsystem} не найден")
                    self.stats['errors'].append(f"Не найден лист для {subsystem}")

            # Парсим сводный лист отдельно (если нужно)
            summary_sheet = None
            for sheet_name in sheet_names:
                if 'сводка' in sheet_name.lower() or 'письмам' in sheet_name.lower():
                    summary_sheet = sheet_name
                    break

            if summary_sheet:
                try:
                    summary_data = self._parse_summary_sheet(excel_file, summary_sheet)
                    if summary_data:
                        result['Сводка'] = summary_data
                        logger.info(f"Лист {summary_sheet}: обработано {len(summary_data)} записей")
                except Exception as e:
                    logger.error(f"Ошибка парсинга сводного листа: {e}")

            self.stats['total_products'] = total_products
            self.stats['last_parse'] = datetime.now().isoformat()

            # Сохраняем в JSON базу данных если она подключена
            if self.json_db and result:
                self._save_to_json_db(result, file_path)

            # Обновляем внутренний кэш
            self.parsed_data = result

            logger.info(f"Парсинг завершен. Всего продуктов: {total_products}")

            return result

        except Exception as e:
            error_msg = f"Критическая ошибка при парсинге Excel: {e}"
            logger.error(error_msg, exc_info=True)
            self.stats['errors'].append(error_msg)
            raise ExcelParseError(error_msg)

    def _parse_sheet(self, excel_file: pd.ExcelFile, sheet_name: str, subsystem: str) -> List[Dict]:
        """
        Парсинг конкретного листа Excel

        Args:
            excel_file: объект ExcelFile
            sheet_name: имя листа
            subsystem: название подсистемы

        Returns:
            List[Dict]: список продуктов с листа
        """
        logger.debug(f"Парсинг листа {sheet_name} для подсистемы {subsystem}")

        try:
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

            if df.empty:
                logger.warning(f"Лист {sheet_name} пуст")
                return []

            # Очистка данных - заменяем NaN на None
            df = df.replace({np.nan: None})

            products = []
            sheet_config = self.SHEET_COLUMNS.get(subsystem, {})

            if subsystem == 'ГМП':
                products.extend(self._parse_gmp_sheet(df, subsystem))
            elif subsystem == 'ГАСУ':
                products.extend(self._parse_gasu_sheet(df, subsystem))
            elif subsystem == 'ЭБ ЗК':
                products.extend(self._parse_ebzk_sheet(df, subsystem))
            elif subsystem == 'ПОИ':
                products.extend(self._parse_poi_sheet(df, subsystem))
            elif subsystem == 'ПУДС':
                products.extend(self._parse_puds_sheet(df, subsystem))
            elif subsystem == 'ПУиО':
                products.extend(self._parse_puio_sheet(df, subsystem))
            elif subsystem == 'ПИАО':
                products.extend(self._parse_piao_sheet(df, subsystem))
            elif subsystem == 'ЕПБС':
                products.extend(self._parse_epbs_sheet(df, subsystem))
            elif subsystem == 'ПУР':
                products.extend(self._parse_pur_sheet(df, subsystem))
            elif subsystem == 'НСИ':
                products.extend(self._parse_nsi_sheet(df, subsystem))
            elif subsystem == 'Лицензии':
                products.extend(self._parse_licenses_sheet(df, subsystem))

            return products

        except Exception as e:
            logger.error(f"Ошибка парсинга листа {sheet_name}: {e}", exc_info=True)
            return []

    def _parse_gmp_sheet(self, df: pd.DataFrame, subsystem: str) -> List[Dict]:
        """Парсинг листа ГМП"""
        products = []

        # Парсим БПО (колонки B-G)
        for idx in range(1, len(df)):
            name = self._get_value(df, idx, 1)  # колонка B
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 2)  # колонка C
                gk_text = self._get_value(df, idx, 5)  # колонка F
                description = self._get_value(df, idx, 4)  # колонка E

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'subsystem': subsystem,
                    'certificate': self._get_certificate_info(df, idx, 3),  # колонка D
                    'source': 'БПО'
                }
                products.append(product)

        # Парсим СПО (колонки J-O)
        for idx in range(1, len(df)):
            name = self._get_value(df, idx, 9)  # колонка J
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 10)  # колонка K
                gk_text = self._get_value(df, idx, 13)  # колонка N
                description = self._get_value(df, idx, 12)  # колонка M

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'subsystem': subsystem,
                    'certificate': self._get_certificate_info(df, idx, 11),  # колонка L
                    'source': 'СПО'
                }
                products.append(product)

        return products

    def _parse_gasu_sheet(self, df: pd.DataFrame, subsystem: str) -> List[Dict]:
        """Парсинг листа ГАСУ"""
        products = []

        # Парсим БПО
        for idx in range(1, len(df)):
            name = self._get_value(df, idx, 1)
            if name and not self._is_header(name) and not str(name).startswith('+') and not str(name).startswith('ЕСО'):
                version = self._get_value(df, idx, 2)
                gk_text = self._get_value(df, idx, 5)
                description = self._get_value(df, idx, 4)

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'subsystem': subsystem,
                    'certificate': self._get_certificate_info(df, idx, 3),
                    'source': 'БПО'
                }
                products.append(product)

        # Парсим СПО
        for idx in range(1, len(df)):
            name = self._get_value(df, idx, 9)
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 10)
                gk_text = self._get_value(df, idx, 14)
                description = self._get_value(df, idx, 13)

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'subsystem': subsystem,
                    'certificate': self._get_certificate_info(df, idx, 11),
                    'source': 'СПО'
                }
                products.append(product)

        # Парсим специальные标记 (+МП, ЕСО)
        for idx in range(1, len(df)):
            marker = self._get_value(df, idx, 0)
            if marker and (str(marker).startswith('+МП') or str(marker).startswith('ЕСО')):
                name = self._get_value(df, idx, 9)
                if name:
                    version = self._get_value(df, idx, 10)
                    gk_text = self._get_value(df, idx, 14)

                    product = {
                        'name': name,
                        'version': version,
                        'gk': self._extract_gk(gk_text) if gk_text else [],
                        'description': f"{marker}",
                        'subsystem': subsystem,
                        'certificate': self._get_certificate_info(df, idx, 11),
                        'source': str(marker)
                    }
                    products.append(product)

        return products

    def _parse_ebzk_sheet(self, df: pd.DataFrame, subsystem: str) -> List[Dict]:
        """Парсинг листа ЭБ ЗК (256ГК)"""
        products = []

        for idx in range(1, len(df)):
            name = self._get_value(df, idx, 2)  # колонка C - Наименование ПО
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 3)  # колонка D - Версия
                description = self._get_value(df, idx, 4)  # колонка E - Комментарий
                po_type = self._get_value(df, idx, 1)  # колонка B - Вид ПО

                product = {
                    'name': name,
                    'version': version,
                    'gk': [],
                    'description': description,
                    'subsystem': subsystem,
                    'po_type': po_type,
                    'source': 'ЭБ ЗК'
                }
                products.append(product)

        return products

    def _parse_poi_sheet(self, df: pd.DataFrame, subsystem: str) -> List[Dict]:
        """Парсинг листа ПОИ"""
        products = []

        # Парсим БПО
        for idx in range(1, len(df)):
            name = self._get_value(df, idx, 1)
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 2)
                gk_text = self._get_value(df, idx, 5)
                description = self._get_value(df, idx, 4)

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'subsystem': subsystem,
                    'certificate': self._get_certificate_info(df, idx, 3),
                    'source': 'БПО'
                }
                products.append(product)

        # Парсим СПО
        for idx in range(1, len(df)):
            name = self._get_value(df, idx, 9)
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 10)
                gk_text = self._get_value(df, idx, 14)
                description = self._get_value(df, idx, 13)

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'subsystem': subsystem,
                    'certificate': self._get_certificate_info(df, idx, 11),
                    'source': 'СПО'
                }
                products.append(product)

        return products

    def _parse_puds_sheet(self, df: pd.DataFrame, subsystem: str) -> List[Dict]:
        """Парсинг листа ПУДС"""
        products = []

        for idx in range(1, len(df)):
            module = self._get_value(df, idx, 0)  # колонка A - модуль
            name = self._get_value(df, idx, 1)  # колонка B - наименование
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 2)  # колонка C - версия
                gk_text = self._get_value(df, idx, 5)  # колонка F - ГК
                description = self._get_value(df, idx, 4)  # колонка E - описание

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'subsystem': subsystem,
                    'module': module,
                    'certificate': self._get_certificate_info(df, idx, 3),  # колонка D
                    'source': 'ПУДС'
                }
                products.append(product)

        return products

    def _parse_puio_sheet(self, df: pd.DataFrame, subsystem: str) -> List[Dict]:
        """Парсинг листа ПУиО"""
        products = []

        # Парсим БПО
        for idx in range(1, len(df)):
            name = self._get_value(df, idx, 1)
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 2)
                gk_text = self._get_value(df, idx, 5)
                description = self._get_value(df, idx, 4)

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'subsystem': subsystem,
                    'certificate': self._get_certificate_info(df, idx, 3),
                    'source': 'БПО'
                }
                products.append(product)

        # Парсим СПО
        for idx in range(1, len(df)):
            name = self._get_value(df, idx, 9)
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 10)
                gk_text = self._get_value(df, idx, 14)
                description = self._get_value(df, idx, 13)
                comment = self._get_value(df, idx, 15)

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'comment': comment,
                    'subsystem': subsystem,
                    'certificate': self._get_certificate_info(df, idx, 11),
                    'source': 'СПО'
                }
                products.append(product)

        return products

    def _parse_piao_sheet(self, df: pd.DataFrame, subsystem: str) -> List[Dict]:
        """Парсинг листа ПИАО"""
        products = []

        # Парсим БПО
        for idx in range(1, len(df)):
            module = self._get_value(df, idx, 0)  # колонка A - модуль
            name = self._get_value(df, idx, 1)
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 2)
                gk_text = self._get_value(df, idx, 5)
                description = self._get_value(df, idx, 4)

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'subsystem': subsystem,
                    'module': module,
                    'certificate': self._get_certificate_info(df, idx, 3),
                    'source': 'БПО'
                }
                products.append(product)

        # Парсим СПО
        for idx in range(1, len(df)):
            module = self._get_value(df, idx, 8)  # колонка I - модуль
            name = self._get_value(df, idx, 9)
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 10)
                gk_text = self._get_value(df, idx, 14)
                description = self._get_value(df, idx, 13)

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'subsystem': subsystem,
                    'module': module,
                    'certificate': self._get_certificate_info(df, idx, 11),
                    'source': 'СПО'
                }
                products.append(product)

        return products

    def _parse_epbs_sheet(self, df: pd.DataFrame, subsystem: str) -> List[Dict]:
        """Парсинг листа ЕПБС"""
        products = []

        # Парсим БПО
        for idx in range(1, len(df)):
            name = self._get_value(df, idx, 1)
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 2)
                gk_text = self._get_value(df, idx, 5)
                description = self._get_value(df, idx, 4)

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'subsystem': subsystem,
                    'certificate': self._get_certificate_info(df, idx, 3),
                    'source': 'БПО'
                }
                products.append(product)

        # Парсим СПО
        for idx in range(1, len(df)):
            module = self._get_value(df, idx, 7)  # колонка H - модуль
            name = self._get_value(df, idx, 8)  # колонка I - наименование
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 9)  # колонка J - версия
                gk_text = self._get_value(df, idx, 13)  # колонка N - ГК
                description = self._get_value(df, idx, 12)  # колонка M - описание

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'subsystem': subsystem,
                    'module': module,
                    'certificate': self._get_certificate_info(df, idx, 10),  # колонка K
                    'source': 'СПО'
                }
                products.append(product)

        return products

    def _parse_pur_sheet(self, df: pd.DataFrame, subsystem: str) -> List[Dict]:
        """Парсинг листа ПУР"""
        products = []

        # Парсим БПО (колонки B-G)
        for idx in range(1, len(df)):
            module = self._get_value(df, idx, 0)  # колонка A - модуль
            name = self._get_value(df, idx, 1)  # колонка B - наименование
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 2)  # колонка C - версия
                gk_text = self._get_value(df, idx, 5)  # колонка F - ГК
                description = self._get_value(df, idx, 4)  # колонка E - описание

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'subsystem': subsystem,
                    'module': module,
                    'certificate': self._get_certificate_info(df, idx, 3),  # колонка D
                    'source': 'БПО'
                }
                products.append(product)

        # Парсим СПО (колонки I-O)
        for idx in range(1, len(df)):
            module = self._get_value(df, idx, 7)  # колонка H - модуль
            name = self._get_value(df, idx, 8)  # колонка I - наименование
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 9)  # колонка J - версия
                gk_text = self._get_value(df, idx, 13)  # колонка N - ГК
                description = self._get_value(df, idx, 12)  # колонка M - описание
                comment = self._get_value(df, idx, 14)  # колонка O - комментарий

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'comment': comment,
                    'subsystem': subsystem,
                    'module': module,
                    'certificate': self._get_certificate_info(df, idx, 10),  # колонка K
                    'source': 'СПО'
                }
                products.append(product)

        return products

    def _parse_nsi_sheet(self, df: pd.DataFrame, subsystem: str) -> List[Dict]:
        """Парсинг листа НСИ"""
        products = []

        # Парсим БПО
        for idx in range(1, len(df)):
            name = self._get_value(df, idx, 1)
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 2)
                gk_text = self._get_value(df, idx, 5)
                description = self._get_value(df, idx, 4)

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'subsystem': subsystem,
                    'certificate': self._get_certificate_info(df, idx, 3),
                    'source': 'БПО'
                }
                products.append(product)

        # Парсим СПО
        for idx in range(1, len(df)):
            name = self._get_value(df, idx, 8)
            if name and not self._is_header(name):
                version = self._get_value(df, idx, 9)
                gk_text = self._get_value(df, idx, 13)
                description = self._get_value(df, idx, 12)
                comment = self._get_value(df, idx, 14)

                product = {
                    'name': name,
                    'version': version,
                    'gk': self._extract_gk(gk_text) if gk_text else [],
                    'description': description,
                    'comment': comment,
                    'subsystem': subsystem,
                    'certificate': self._get_certificate_info(df, idx, 10),
                    'source': 'СПО'
                }
                products.append(product)

        return products

    def _parse_licenses_sheet(self, df: pd.DataFrame, subsystem: str) -> List[Dict]:
        """Парсинг листа Лицензии"""
        products = []

        for idx in range(1, len(df)):
            name = self._get_value(df, idx, 0)  # колонка A - наименование
            if name and not self._is_header(name):
                correct_name = self._get_value(df, idx, 1)  # колонка B - правильное название
                owner = self._get_value(df, idx, 2)  # колонка C - владелец
                description = self._get_value(df, idx, 3)  # колонка D - предназначение

                product = {
                    'name': correct_name if correct_name else name,
                    'version': None,
                    'gk': [],
                    'description': description,
                    'subsystem': subsystem,
                    'owner': owner,
                    'source': 'Лицензии'
                }
                products.append(product)

        return products

    def _parse_summary_sheet(self, excel_file: pd.ExcelFile, sheet_name: str) -> List[Dict]:
        """Парсинг сводного листа с письмами"""
        try:
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            products = []

            for idx, row in df.iterrows():
                gk = row.iloc[0] if len(row) > 0 else None  # Номер ГК
                systems = row.iloc[1] if len(row) > 1 else None  # Системы/подсистемы

                if pd.notna(gk) and pd.notna(systems):
                    gk_list = self._extract_gk(str(gk))
                    if gk_list:
                        product = {
                            'name': f"ГК {gk_list[0]}",
                            'version': None,
                            'gk': gk_list,
                            'description': str(systems) if pd.notna(systems) else None,
                            'subsystem': 'Сводка',
                            'source': 'Сводка по письмам'
                        }
                        products.append(product)

            return products
        except Exception as e:
            logger.error(f"Ошибка парсинга сводного листа: {e}")
            return []

    def _get_value(self, df: pd.DataFrame, row: int, col: int) -> Optional[str]:
        """Безопасное получение значения из ячейки"""
        try:
            if col < len(df.columns) and row < len(df):
                val = df.iloc[row, col]
                if pd.notna(val):
                    return str(val).strip()
            return None
        except:
            return None

    def _is_header(self, text: str) -> bool:
        """Проверка, является ли текст заголовком"""
        if not text:
            return False
        text_lower = str(text).lower()
        headers = ['модуль', 'наименование', 'версия', 'сертификат', 'предназначение',
                   'бпо', 'спо', 'языки', 'платформы', 'комментарий', '№']
        return any(h in text_lower for h in headers)

    def _get_certificate_info(self, df: pd.DataFrame, row: int, col: int) -> Optional[str]:
        """Получение информации о сертификате"""
        val = self._get_value(df, row, col)
        if val:
            val_lower = val.lower()
            if 'да' in val_lower or 'есть' in val_lower or 'сертифицирован' in val_lower:
                return 'да'
            elif 'нет' in val_lower:
                return 'нет'
        return val

    def _extract_gk(self, text: str) -> List[str]:
        """Извлечение ГК из текста"""
        if not text:
            return []

        # Паттерн для поиска ГК
        pattern = r'ФКУ\d{3,4}(?:[/-]\d{2,4})?(?:/\w+)?'
        matches = re.findall(pattern, text, re.IGNORECASE)
        return [m.upper() for m in matches]

    def _save_to_json_db(self, data: Dict[str, List[Dict]], source_file: str):
        """
        Сохранение данных в JSON базу данных

        Args:
            data: словарь с данными
            source_file: исходный файл
        """
        if not self.json_db:
            return

        logger.info(f"Сохранение данных в JSON базу данных: {len(data)} подсистем")

        added_count = 0
        updated_count = 0

        for subsystem, products in data.items():
            if subsystem == 'Сводка':
                continue  # Пропускаем сводный лист

            for product in products:
                # Проверяем существование продукта
                existing = self.json_db.find_products(
                    subsystem=subsystem,
                    search=product['name']
                )

                if existing:
                    # Обновляем существующий продукт
                    existing_product = existing[0]
                    product_id = existing_product['id']

                    updates = {
                        'version': product.get('version'),
                        'gk': product.get('gk', []),
                        'description': product.get('description'),
                        'certificate': product.get('certificate'),
                        'module': product.get('module')
                    }

                    if self.json_db.update_product(product_id, updates):
                        updated_count += 1
                else:
                    # Добавляем новый продукт
                    product['source_file'] = source_file
                    if self.json_db.add_product(product):
                        added_count += 1

        logger.info(f"JSON база данных обновлена: добавлено {added_count}, обновлено {updated_count}")

    def get_stats(self) -> Dict:
        """Получение статистики"""
        return self.stats.copy()

    def get_errors(self) -> List[str]:
        """Получение списка ошибок"""
        return self.stats['errors'].copy()

    def clear_errors(self):
        """Очистка списка ошибок"""
        self.stats['errors'] = []

    def print_summary(self):
        """Вывод сводки по парсингу"""
        print("\n" + "=" * 60)
        print("СВОДКА ПО ПАРСИНГУ EXCEL")
        print("=" * 60)

        print(f"Файл: {self.current_file_path or 'Не указан'}")
        print(f"Дата парсинга: {self.stats['last_parse'] or 'Не выполнялся'}")
        print(f"Всего продуктов: {self.stats['total_products']}")

        if self.stats['by_sheet']:
            print("\nПо подсистемам:")
            for subs, count in self.stats['by_sheet'].items():
                print(f"  {subs}: {count}")

        if self.stats['errors']:
            print(f"\nОшибки ({len(self.stats['errors'])}):")
            for error in self.stats['errors'][:5]:
                print(f"  • {error}")
            if len(self.stats['errors']) > 5:
                print(f"  ... и еще {len(self.stats['errors']) - 5}")

        print("=" * 60)


# Пример использования
if __name__ == "__main__":
    # Настройка логирования для примера
    logging.basicConfig(level=logging.INFO)

    # Создаем парсер
    parser = ExcelParser()

    # Путь к файлу Excel
    excel_file = "Состав ПО ИС ФК (согласованный по письмам).xlsx"

    try:
        # Парсим файл
        data = parser.parse_excel_file(excel_file)

        # Выводим сводку
        parser.print_summary()

        # Показываем пример данных для ГМП
        if 'ГМП' in data:
            print(f"\nПример данных для ГМП:")
            for i, product in enumerate(data['ГМП'][:3]):
                print(f"\nПродукт {i + 1}:")
                print(f"  Наименование: {product.get('name')}")
                print(f"  Версия: {product.get('version')}")
                print(f"  ГК: {product.get('gk')}")
                print(f"  Описание: {product.get('description')}")

    except ExcelFileNotFoundError as e:
        print(f"Ошибка: {e}")
        print("Укажите правильный путь к файлу Excel")
    except Exception as e:
        print(f"Непредвиденная ошибка: {e}")