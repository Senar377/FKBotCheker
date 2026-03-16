# versions_dialog.py
"""
Диалог для отображения и редактирования версий ПО из документа
С интеграцией с JSON базой данных и нечетким поиском
"""

import logging
import re
import json
import os
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
from difflib import SequenceMatcher

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QLineEdit,
    QComboBox, QMessageBox, QGroupBox, QTextEdit, QSplitter,
    QMenu, QApplication, QWidget, QFileDialog, QTabWidget,
    QInputDialog, QDialogButtonBox, QFormLayout, QCheckBox,
    QGridLayout, QFrame, QToolTip, QSpinBox, QDateEdit,
    QProgressBar
)
from PyQt6.QtCore import Qt, pyqtSignal, QTimer, QSettings, QDate, QThread
from PyQt6.QtGui import QFont, QColor, QAction, QTextCursor, QTextCharFormat, QIcon

# Импортируем наши модули
from excel_parser import ExcelParser
from json_database import JSONDatabase, JSONDatabaseError

logger = logging.getLogger(__name__)


class FuzzySearchThread(QThread):
    """Поток для выполнения нечеткого поиска наименований ПО из БД в тексте документа"""

    progress = pyqtSignal(int)
    result_ready = pyqtSignal(list)
    finished = pyqtSignal()

    def __init__(self, document_text, db_products, threshold=0.8):
        super().__init__()
        self.document_text = document_text
        self.db_products = db_products
        self.threshold = threshold
        self.lines = document_text.split('\n') if document_text else []

    def run(self):
        results = []
        total = len(self.db_products)

        # Предобработка текста - разбиваем на строки и слова
        lines_with_words = []
        for line_num, line in enumerate(self.lines, 1):
            words = re.findall(r'\b\w+\b', line.lower())
            lines_with_words.append({
                'num': line_num,
                'text': line,
                'words': words,
                'words_set': set(words)
            })

        for i, product in enumerate(self.db_products):
            if i % 10 == 0:
                self.progress.emit(int(i * 100 / total))

            name = product.get('name', '')
            if not name:
                continue

            # Ищем продукт в тексте
            match_info = self.fuzzy_search_in_text(name, lines_with_words)

            result = {
                'product': product,
                'found': match_info is not None,
                'match_info': match_info,
                'similarity': match_info[0] if match_info else 0,
                'line_number': match_info[1] if match_info else None,
                'context': match_info[2] if match_info else None,
                'match_type': match_info[3] if match_info else None
            }
            results.append(result)

        self.progress.emit(100)
        self.result_ready.emit(results)
        self.finished.emit()

    def fuzzy_search_in_text(self, product_name, lines_with_words):
        """
        Нечеткий поиск названия продукта в тексте с порогом 80%
        """
        if not product_name or not lines_with_words:
            return None

        product_lower = product_name.lower()
        best_match = None
        best_ratio = 0

        # Разбиваем название на слова (игнорируем короткие слова)
        product_words = [w for w in re.findall(r'\b\w+\b', product_lower) if len(w) > 2]

        if not product_words:
            return None

        for line_data in lines_with_words:
            line_num = line_data['num']
            line_text = line_data['text']
            line_words = line_data['words']
            line_words_set = line_data['words_set']

            # 1. Сначала проверяем точное вхождение всей строки
            if product_lower in line_text.lower():
                ratio = 1.0
                context = self.get_context(line_num)
                return (ratio, line_num, context, 'exact')

            # 2. Проверяем вхождение всех слов продукта
            words_found = 0
            for word in product_words:
                if word in line_words_set:
                    words_found += 1

            if words_found == len(product_words):
                ratio = 1.0
                context = self.get_context(line_num)
                return (ratio, line_num, context, 'all_words')

            # 3. Проверяем вхождение большинства слов (больше 80%)
            if words_found > 0:
                word_ratio = words_found / len(product_words)
                if word_ratio >= 0.8 and word_ratio > best_ratio:
                    best_ratio = word_ratio
                    best_match = (word_ratio, line_num, self.get_context(line_num), 'most_words')

            # 4. Нечеткое сравнение для каждого слова
            for word in line_words:
                if len(word) > 3 and len(word) < 30:
                    # Сравниваем слово с названием продукта
                    ratio = SequenceMatcher(None, product_lower, word).ratio()
                    if ratio >= self.threshold and ratio > best_ratio:
                        best_ratio = ratio
                        best_match = (ratio, line_num, self.get_context(line_num), 'fuzzy_word')

                    # Сравниваем слово с каждым словом продукта
                    for p_word in product_words:
                        if len(p_word) > 3:
                            word_ratio = SequenceMatcher(None, p_word, word).ratio()
                            if word_ratio >= self.threshold and word_ratio > best_ratio:
                                best_ratio = word_ratio
                                best_match = (word_ratio, line_num, self.get_context(line_num), 'fuzzy')

            # 5. Проверяем, содержит ли строка значительную часть слов продукта
            if len(line_words) > 0:
                # Считаем количество совпадающих слов с учетом нечеткого сравнения
                fuzzy_matches = 0
                for p_word in product_words:
                    for line_word in line_words:
                        if len(line_word) > 3:
                            if SequenceMatcher(None, p_word, line_word).ratio() >= 0.8:
                                fuzzy_matches += 1
                                break

                if fuzzy_matches > 0:
                    fuzzy_ratio = fuzzy_matches / len(product_words)
                    if fuzzy_ratio >= 0.8 and fuzzy_ratio > best_ratio:
                        best_ratio = fuzzy_ratio
                        best_match = (fuzzy_ratio, line_num, self.get_context(line_num), 'fuzzy_multiple')

        return best_match

    def get_context(self, line_num, context_lines=1):
        """Получение контекста вокруг строки"""
        start = max(0, line_num - context_lines - 1)
        end = min(len(self.lines), line_num + context_lines)

        context = []
        for i in range(start, end):
            if i == line_num - 1:
                prefix = "→ "
            else:
                prefix = "  "
            context.append(f"{prefix}{self.lines[i][:150]}")

        return '\n'.join(context)


class VersionCleaner:
    """Класс для очистки и нормализации версий"""

    # Паттерны для мусора, который нужно исключить
    GARBAGE_PATTERNS = [
        r'\d{2}\.\d{2}\.\d{4}',  # даты DD.MM.YYYY
        r'\d{2}\.\d{2}\.\d{2}',  # даты DD.MM.YY
        r'\d{4}-\d{2}-\d{2}',  # даты YYYY-MM-DD
        r'№\s*\d+',  # номера документов
        r'приказ\s*\d+',  # приказы
        r'постановление\s*\d+',  # постановления
        r'распоряжение\s*\d+',  # распоряжения
        r'письмо\s*\d+',  # письма
        r'от\s*\d{2}\.\d{2}\.\d{4}',  # от ДД.ММ.ГГГГ
        r'ред\.?\s*\d+\.\d+',  # редакции
        r'версия\s*\d+\.\d+',  # слово "версия" с цифрами
        r'https?://[^\s]+',  # ссылки на сайты
        r'www\.[^\s]+',  # ссылки на сайты
        r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}',  # email адреса
    ]

    @staticmethod
    def clean_version(version_str):
        """Очистка версии от мусора"""
        if not version_str:
            return ""

        version = str(version_str).strip()

        # Проверяем, не является ли строка мусором
        for pattern in VersionCleaner.GARBAGE_PATTERNS:
            if re.match(pattern, version, re.IGNORECASE):
                return ""

        # Убираем лишние пробелы и специальные символы в начале и конце
        version = re.sub(r'^[^\d]+', '', version)
        version = re.sub(r'[^\d\.]+$', '', version)

        # Паттерны для поиска версии в тексте (только после наименования ПО)
        patterns = [
            r'(\d+\.\d+\.\d+\.\d+)',  # 1.2.3.4
            r'(\d+\.\d+\.\d+)',  # 1.2.3
            r'(\d+\.\d+)',  # 1.2
            r'(\d+)',  # 1
        ]

        for pattern in patterns:
            match = re.search(pattern, version)
            if match:
                return match.group(1)

        # Если ничего не нашли, возвращаем пустую строку
        return ""

    @staticmethod
    def is_valid_version(version):
        """Проверка, является ли строка валидной версией"""
        if not version:
            return False

        # Проверяем, что версия содержит цифры и точки
        if not re.search(r'\d', version):
            return False

        # Проверяем, что версия не слишком длинная
        if len(version) > 20:
            return False

        # Проверяем, что версия не содержит лишних символов
        if re.search(r'[^\d\.]', version):
            return False

        # Проверяем, что версия имеет правильный формат
        parts = version.split('.')
        if len(parts) > 4:  # Не больше 4 частей
            return False

        for part in parts:
            if not part.isdigit():
                return False
            if len(part) > 5:  # Каждая часть не больше 5 цифр
                return False

        return True

    @staticmethod
    def is_garbage(text):
        """Проверка, является ли текст мусором"""
        if not text:
            return True

        text_lower = text.lower()

        # Проверяем на наличие ключевых слов мусора
        garbage_keywords = [
            'приказ', 'постановление', 'распоряжение', 'письмо',
            'от', '№', 'дата', 'год', 'г.', 'редакция', 'ред.',
            'утвержд', 'внесен', 'изменен', 'дополнен',
            'http', 'https', 'www', '.ru', '.com', '.org', 'email'
        ]

        for keyword in garbage_keywords:
            if keyword in text_lower:
                return True

        # Проверяем на паттерны дат
        date_patterns = [
            r'\d{2}\.\d{2}\.\d{4}',
            r'\d{2}\.\d{2}\.\d{2}',
            r'\d{4}-\d{2}-\d{2}',
        ]

        for pattern in date_patterns:
            if re.search(pattern, text):
                return True

        # Проверяем на наличие ссылок
        url_patterns = [
            r'https?://[^\s]+',
            r'www\.[^\s]+',
        ]

        for pattern in url_patterns:
            if re.search(pattern, text):
                return True

        return False


class GKManager:
    """Класс для управления ГК"""

    def __init__(self):
        self.settings = QSettings("ФедеральноеКазначейство", "Settings")
        self.document_gk = []  # ГК найденные в документе
        self.first_pages_gk = []  # ГК с первых страниц
        self.gk_to_subsystem = {}  # Словарь соответствия ГК и подсистем

    def extract_gk_from_text(self, text, page_info=None, max_pages=3):
        """Извлечение ГК из текста"""
        if not text:
            return []

        # Получаем формат ГК из настроек
        gk_pattern = self.settings.value(
            "gk_format",
            r"ФКУ\d{3,4}(?:[/-]\d{2,4})?(?:/\w+)?"
        )

        # Ищем все ГК в тексте
        matches = re.findall(gk_pattern, text, re.IGNORECASE)
        all_gk = [m.upper() for m in matches]
        self.document_gk = sorted(list(set(all_gk)))

        # Определяем подсистемы для каждого ГК
        self.map_gk_to_subsystems()

        # Если есть информация о страницах, ищем ГК на первых страницах
        if page_info:
            pages_text = []
            for i, (start, end) in enumerate(page_info):
                if i < max_pages:
                    page_text = text[start:end]
                    pages_text.append(page_text)

            first_pages_matches = []
            for page_text in pages_text:
                matches = re.findall(gk_pattern, page_text, re.IGNORECASE)
                first_pages_matches.extend(matches)

            self.first_pages_gk = sorted(list(set([m.upper() for m in first_pages_matches])))
        else:
            self.first_pages_gk = self.document_gk[:]

        return self.document_gk

    def map_gk_to_subsystems(self):
        """Сопоставление ГК с подсистемами"""
        # Словарь соответствия ГК и подсистем
        gk_mapping = {
            'ФКУ0375/2025/РИС': 'ГМП',
            'ФКУ0241/2025/РИС': 'ГАСУ',
            'ФКУ0240/2025/РИС': 'ГАСУ',
            'ФКУ0215/2025/РИС': 'ПОИ',
            'ФКУ0237/2025/РИС': 'ПУДС',
            'ФКУ0233/2025/РИС': 'ПУДС',
            'ФКУ0246/2025/РИС': 'ПУДС',
            'ФКУ0257/2025/РИС': 'ПУиО',
            'ФКУ0231/2025/РИС': 'ПУиО',
            'ФКУ0247/2025/РИС': 'ПУиО',
            'ФКУ0261/2025/РИС': 'ПУиО',
            'ФКУ0173/2025/РИС': 'ПУиО',
            'ФКУ0358/2025/РИС': 'ПУиО',
            'ФКУ0346/2025/РИС': 'ПИАО',
            'ФКУ0404/2025/РИС': 'ПИАО',
            'ФКУ0336/2025/РИС': 'ЕПБС',
            'ФКУ0289/2025/РИС': 'ЕПБС',
            'ФКУ0232/2025/РИС': 'ПУР',
        }

        self.gk_to_subsystem = {}
        for gk in self.document_gk:
            # Проверяем точное соответствие
            if gk in gk_mapping:
                self.gk_to_subsystem[gk] = gk_mapping[gk]
            else:
                # Проверяем частичное соответствие
                for key, subsystem in gk_mapping.items():
                    if key in gk or gk in key:
                        self.gk_to_subsystem[gk] = subsystem
                        break
                else:
                    self.gk_to_subsystem[gk] = 'Не определена'

    def get_subsystem_for_gk(self, gk):
        """Получение подсистемы для ГК"""
        return self.gk_to_subsystem.get(gk, 'Не определена')

    def get_gk_priority(self, gk_list):
        """Получение приоритета для списка ГК"""
        if not gk_list:
            return 0

        priority = 0
        priority_from_doc = self.settings.value("priority_from_doc", True, type=bool)
        priority_from_first_pages = self.settings.value("priority_from_first_pages", True, type=bool)

        for gk in gk_list:
            if priority_from_doc and gk in self.document_gk:
                priority += 10
                if priority_from_first_pages and gk in self.first_pages_gk:
                    priority += 20
            elif gk:
                priority += 1

        return priority

    def has_matching_gk(self, gk_list):
        """Проверка наличия совпадающих ГК с документом"""
        if not gk_list or not self.document_gk:
            return False

        for gk in gk_list:
            if gk in self.document_gk:
                return True

        return False


class TableDetector:
    """Класс для определения и извлечения таблиц из документа"""

    def __init__(self):
        self.settings = QSettings("ФедеральноеКазначейство", "Settings")

    def detect_tables(self, text):
        """Определение таблиц в тексте"""
        if not self.settings.value("detect_tables", True, type=bool):
            return []

        lines = text.split('\n')
        tables = []
        current_table = []
        in_table = False
        min_cols = self.settings.value("min_table_cols", 3, type=int)
        min_rows = self.settings.value("min_table_rows", 3, type=int)

        # Получаем разделители
        separators_str = self.settings.value("table_separators", "|, \\t,  ,")
        separators = [s.strip() for s in separators_str.split(',')]

        for i, line in enumerate(lines):
            # Проверяем, является ли строка частью таблицы
            is_table_line = False
            cols = []
            used_sep = None

            # Проверяем наличие разделителей
            for sep in separators:
                if sep and sep in line:
                    # Считаем количество колонок
                    if sep == '  ':
                        cols = [col.strip() for col in re.split(r'\s{2,}', line) if col.strip()]
                        used_sep = '  '
                    elif sep == '\\t':
                        cols = [col.strip() for col in line.split('\t') if col.strip()]
                        used_sep = '\t'
                    else:
                        cols = [col.strip() for col in line.split(sep) if col.strip()]
                        used_sep = sep

                    if len(cols) >= min_cols:
                        is_table_line = True
                        break

            if is_table_line:
                if not in_table:
                    in_table = True
                    current_table = []
                current_table.append({
                    'line_num': i + 1,
                    'text': line,
                    'columns': cols,
                    'raw_columns': line.split(used_sep) if used_sep else [],
                    'separator': used_sep
                })
            else:
                if in_table and len(current_table) >= min_rows:
                    # Сохраняем таблицу
                    tables.append({
                        'start_line': current_table[0]['line_num'],
                        'end_line': current_table[-1]['line_num'],
                        'rows': current_table,
                        'headers': self._identify_headers(current_table)
                    })
                in_table = False
                current_table = []

        # Проверяем последнюю таблицу
        if in_table and len(current_table) >= min_rows:
            tables.append({
                'start_line': current_table[0]['line_num'],
                'end_line': current_table[-1]['line_num'],
                'rows': current_table,
                'headers': self._identify_headers(current_table)
            })

        return tables

    def _identify_headers(self, table_rows):
        """Определение строки заголовков таблицы"""
        if not table_rows:
            return None

        # Признаки заголовков
        headers_str = self.settings.value("table_headers", "наименование, продукт, версия, гк, описание")
        header_keywords = [h.strip().lower() for h in headers_str.split(',')]

        # Проверяем первую строку на наличие ключевых слов
        first_row = table_rows[0]
        first_row_text = ' '.join(first_row['columns']).lower()

        for keyword in header_keywords:
            if keyword in first_row_text:
                return first_row

        # Проверяем вторую строку
        if len(table_rows) > 1:
            second_row = table_rows[1]
            second_row_text = ' '.join(second_row['columns']).lower()
            for keyword in header_keywords:
                if keyword in second_row_text:
                    return second_row

        return None

    def extract_versions_from_tables(self, tables):
        """Извлечение версий и ГК из таблиц"""
        versions = []

        for table in tables:
            headers = table['headers']
            start_idx = 0

            # Если есть заголовки, пропускаем их
            if headers:
                for i, row in enumerate(table['rows']):
                    if row == headers:
                        start_idx = i + 1
                        break

            # Определяем индексы колонок
            name_col_idx = -1
            version_col_idx = -1
            gk_col_idx = -1

            if headers:
                for i, col in enumerate(headers['columns']):
                    col_lower = col.lower()
                    if any(kw in col_lower for kw in ['наименование', 'продукт', 'по', 'name']):
                        name_col_idx = i
                    elif any(kw in col_lower for kw in ['версия', 'вер', 'version']):
                        version_col_idx = i
                    elif any(kw in col_lower for kw in ['гк', 'контракт', 'gk']):
                        gk_col_idx = i

            # Извлекаем данные из строк таблицы
            for row in table['rows'][start_idx:]:
                row_data = {
                    'name': '',
                    'version': '',
                    'gk': [],
                    'line_num': row['line_num'],
                    'source': 'table',
                    'table_start': table['start_line'],
                    'table_end': table['end_line']
                }

                if name_col_idx >= 0 and name_col_idx < len(row['columns']):
                    name = row['columns'][name_col_idx]
                    # Проверяем, не является ли имя мусором
                    if not VersionCleaner.is_garbage(name):
                        row_data['name'] = name

                if version_col_idx >= 0 and version_col_idx < len(row['columns']):
                    version_text = row['columns'][version_col_idx]
                    # Проверяем, не является ли версия мусором
                    if not VersionCleaner.is_garbage(version_text):
                        cleaned = VersionCleaner.clean_version(version_text)
                        row_data['version'] = cleaned if cleaned else ''

                if gk_col_idx >= 0 and gk_col_idx < len(row['columns']):
                    gk_text = row['columns'][gk_col_idx]
                    # Извлекаем ГК из текста
                    gk_pattern = self.settings.value(
                        "gk_format",
                        r"ФКУ\d{3,4}(?:[/-]\d{2,4})?(?:/\w+)?"
                    )
                    gk_matches = re.findall(gk_pattern, gk_text, re.IGNORECASE)
                    row_data['gk'] = [gk.upper() for gk in gk_matches if not VersionCleaner.is_garbage(gk)]

                if row_data['name'] and (row_data['version'] or row_data['gk']):
                    versions.append(row_data)

        return versions


class ProductEditDialog(QDialog):
    """Диалог для редактирования/добавления продукта"""

    def __init__(self, parent=None, product_data=None, subsystems=None):
        super().__init__(parent)
        self.product_data = product_data or {}
        self.subsystems = subsystems or []
        self.result_data = None

        self.setWindowTitle("Редактирование продукта" if product_data else "Добавление продукта")
        self.resize(600, 500)

        self.init_ui()
        self.load_data()

    def init_ui(self):
        """Инициализация интерфейса"""
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Форма
        form_layout = QFormLayout()
        form_layout.setSpacing(10)

        # Наименование
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("Введите наименование ПО")
        form_layout.addRow("Наименование ПО:*", self.name_edit)

        # Версия
        self.version_edit = QLineEdit()
        self.version_edit.setPlaceholderText("Введите версию")
        form_layout.addRow("Версия:", self.version_edit)

        # Подсистема
        self.subsystem_combo = QComboBox()
        self.subsystem_combo.addItem("Не определена")
        for subs in sorted(self.subsystems):
            if subs != "Не определена":
                self.subsystem_combo.addItem(subs)
        form_layout.addRow("Подсистема:", self.subsystem_combo)

        # ГК
        self.gk_edit = QLineEdit()
        self.gk_edit.setPlaceholderText("Номера ГК через запятую")
        form_layout.addRow("ГК:", self.gk_edit)

        # Описание
        self.description_edit = QTextEdit()
        self.description_edit.setMaximumHeight(80)
        form_layout.addRow("Описание:", self.description_edit)

        # Сертификат
        self.certificate_edit = QLineEdit()
        self.certificate_edit.setPlaceholderText("Наличие сертификата ФСТЭК")
        form_layout.addRow("Сертификат:", self.certificate_edit)

        # Платформа
        self.platform_edit = QLineEdit()
        self.platform_edit.setPlaceholderText("Используемые платформы")
        form_layout.addRow("Платформа:", self.platform_edit)

        # Язык
        self.language_edit = QLineEdit()
        self.language_edit.setPlaceholderText("Языки программирования")
        form_layout.addRow("Язык:", self.language_edit)

        # Владелец
        self.owner_edit = QLineEdit()
        self.owner_edit.setPlaceholderText("Владелец/Правообладатель")
        form_layout.addRow("Владелец:", self.owner_edit)

        # Модуль
        self.module_edit = QLineEdit()
        self.module_edit.setPlaceholderText("Модуль/компонент")
        form_layout.addRow("Модуль:", self.module_edit)

        # Источник
        self.source_edit = QLineEdit()
        self.source_edit.setPlaceholderText("Источник данных")
        form_layout.addRow("Источник:", self.source_edit)

        layout.addLayout(form_layout)

        # Информация о последнем обновлении
        if self.product_data and self.product_data.get('last_updated'):
            info_label = QLabel(f"Последнее обновление: {self.product_data['last_updated']}")
            info_label.setStyleSheet("color: gray; font-size: 10px;")
            layout.addWidget(info_label)

        # Кнопки
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok |
                                      QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.validate_and_accept)
        button_box.rejected.connect(self.reject)

        layout.addWidget(button_box)

        self.setLayout(layout)

    def load_data(self):
        """Загрузка данных в форму"""
        if self.product_data:
            self.name_edit.setText(self.product_data.get('name', ''))
            self.version_edit.setText(self.product_data.get('version', ''))

            # Подсистема
            subsystem = self.product_data.get('subsystem', 'Не определена')
            index = self.subsystem_combo.findText(subsystem)
            if index >= 0:
                self.subsystem_combo.setCurrentIndex(index)

            # ГК
            gk_list = self.product_data.get('gk', [])
            self.gk_edit.setText(', '.join(gk_list))

            self.description_edit.setText(self.product_data.get('description', ''))
            self.certificate_edit.setText(self.product_data.get('certificate', ''))
            self.platform_edit.setText(self.product_data.get('platform', ''))
            self.language_edit.setText(self.product_data.get('language', ''))
            self.owner_edit.setText(self.product_data.get('owner', ''))
            self.module_edit.setText(self.product_data.get('module', ''))
            self.source_edit.setText(self.product_data.get('source_file', ''))

    def validate_and_accept(self):
        """Валидация и принятие данных"""
        # Проверка обязательных полей
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Ошибка", "Наименование ПО обязательно для заполнения")
            return

        # Собираем данные
        self.result_data = {
            'name': name,
            'version': self.version_edit.text().strip() or None,
            'subsystem': self.subsystem_combo.currentText(),
            'gk': [gk.strip() for gk in self.gk_edit.text().split(',') if gk.strip()],
            'description': self.description_edit.toPlainText().strip() or None,
            'certificate': self.certificate_edit.text().strip() or None,
            'platform': self.platform_edit.text().strip() or None,
            'language': self.language_edit.text().strip() or None,
            'owner': self.owner_edit.text().strip() or None,
            'module': self.module_edit.text().strip() or None,
            'source_file': self.source_edit.text().strip() or None
        }

        self.accept()

    def get_data(self):
        """Получение введенных данных"""
        return self.result_data


class VersionsDialog(QDialog):
    """Диалог для отображения и редактирования версий ПО из документа"""

    def __init__(self, parent=None, document_text="", page_info=None):
        super().__init__(parent)
        self.document_text = document_text
        self.page_info = page_info or []
        self.versions_data = []
        self.table_versions = []
        self.comparison_results = []
        self.fuzzy_search_results = []
        self.software_patterns = self.get_software_patterns()
        self.category_keywords = self.build_category_keywords()
        self.settings = QSettings("ФедеральноеКазначейство", "Settings")

        # Инициализируем JSON базу данных
        db_dir = self.settings.value("json_db_dir", "json_database")
        self.json_db = JSONDatabase(db_dir)

        # Инициализируем парсер Excel
        self.excel_parser = ExcelParser(self.json_db)

        # Инициализируем менеджеры
        self.gk_manager = GKManager()
        self.table_detector = TableDetector()

        # Загружаем настройки
        self.load_settings()

        self.setWindowTitle("Версии БПО/СПО в документе")
        self.resize(1500, 950)

        if parent and hasattr(parent, 'styleSheet'):
            self.setStyleSheet(parent.styleSheet())

        # Извлекаем ГК из документа
        self.extract_gk_from_document()

        # Определяем таблицы в документе
        self.detect_tables_in_document()

        # Сканируем документ на версии
        self.scan_document_for_versions()

        self.init_ui()
        self.display_versions()

    def load_settings(self):
        """Загрузка настроек"""
        self.auto_clean_versions = self.settings.value("auto_clean_versions", True, type=bool)
        self.min_version_length = self.settings.value("min_version_length", 3, type=int)
        self.enable_gk_sort = self.settings.value("enable_gk_sort", True, type=bool)
        self.gk_sort_type = self.settings.value("gk_sort_type", "match")
        self.extract_from_tables = self.settings.value("extract_from_tables", True, type=bool)
        self.table_priority = self.settings.value("table_priority", True, type=bool)
        self.hide_invalid_versions = self.settings.value("hide_invalid_versions", False, type=bool)
        self.show_only_matched = self.settings.value("show_only_matched", False, type=bool)
        self.sort_by_subsystem = self.settings.value("sort_by_subsystem", True, type=bool)

        # Загружаем порядок подсистем
        subsystem_order = self.settings.value(
            "subsystem_order",
            "ГМП, ГАСУ, ПОИ, ПУДС, ПУиО, ПИАО, ЕПБС, ПУР, НСИ"
        )
        self.subsystem_priority = {
            s.strip(): i for i, s in enumerate(subsystem_order.split(','))
        }

    def extract_gk_from_document(self):
        """Извлечение ГК из документа"""
        gk_pages = self.settings.value("gk_pages", 3, type=int)
        self.gk_manager.extract_gk_from_text(self.document_text, self.page_info, gk_pages)

    def detect_tables_in_document(self):
        """Определение таблиц в документе"""
        if self.extract_from_tables:
            tables = self.table_detector.detect_tables(self.document_text)
            self.table_versions = self.table_detector.extract_versions_from_tables(tables)
            logger.info(f"Найдено {len(self.table_versions)} записей в таблицах")

    def init_ui(self):
        """Инициализация интерфейса"""
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        self.tab_widget = QTabWidget()

        # Вкладка 1: Версии из документа
        self.document_tab = QWidget()
        self.init_document_tab()
        self.tab_widget.addTab(self.document_tab, "📄 Версии из документа")

        # Вкладка 2: Сравнение с базой данных
        self.comparison_tab = QWidget()
        self.init_comparison_tab()
        self.tab_widget.addTab(self.comparison_tab, "🔍 Сравнение с БД")

        # Вкладка 3: Управление базой данных
        self.database_tab = QWidget()
        self.init_database_tab()
        self.tab_widget.addTab(self.database_tab, "💾 База данных")

        # Вкладка 4: СПО/БПО (нечеткий поиск) - создаем, но не добавляем
        self.fuzzy_tab = QWidget()
        self.init_fuzzy_tab()
        self.fuzzy_tab_index = -1  # Индекс вкладки, -1 означает что не добавлена

        # Применяем начальную видимость вкладки
        self.update_fuzzy_tab_visibility()

        main_layout.addWidget(self.tab_widget)

        # Кнопка закрытия
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        close_btn.setMinimumHeight(35)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                font-weight: bold;
                border: none;
                border-radius: 5px;
                padding: 8px 20px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        main_layout.addLayout(btn_layout)

        self.setLayout(main_layout)

    def init_fuzzy_tab(self):
        """Инициализация вкладки нечеткого поиска СПО/БПО с порогом 80%"""
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Верхняя панель с информацией
        info_panel = QGroupBox("Информация о поиске")
        info_layout = QHBoxLayout()

        self.fuzzy_stats_label = QLabel("📊 Всего продуктов в БД: 0 | Найдено: 0 | Не найдено: 0")
        self.fuzzy_stats_label.setStyleSheet("font-weight: bold; color: #2c3e50;")

        info_layout.addWidget(self.fuzzy_stats_label)
        info_layout.addStretch()

        info_panel.setLayout(info_layout)
        layout.addWidget(info_panel)

        # Панель управления
        control_panel = QGroupBox("Параметры поиска (порог схожести фиксирован 80%)")
        control_layout = QHBoxLayout()

        # Кнопка запуска поиска
        self.fuzzy_search_btn = QPushButton("🔍 Найти СПО/БПО в документе")
        self.fuzzy_search_btn.clicked.connect(self.start_fuzzy_search)
        self.fuzzy_search_btn.setMinimumHeight(40)
        self.fuzzy_search_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                font-weight: bold;
                border-radius: 5px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #229954;
            }
            QPushButton:disabled {
                background-color: #95a5a6;
            }
        """)

        # Прогресс бар
        self.fuzzy_progress = QProgressBar()
        self.fuzzy_progress.setVisible(False)

        # Фильтры
        control_layout.addWidget(QLabel("Подсистема:"))
        self.fuzzy_subsystem_filter = QComboBox()
        self.fuzzy_subsystem_filter.addItem("Все подсистемы")
        self.fuzzy_subsystem_filter.currentTextChanged.connect(self.filter_fuzzy_table)
        control_layout.addWidget(self.fuzzy_subsystem_filter)

        control_layout.addWidget(QLabel("Статус:"))
        self.fuzzy_status_filter = QComboBox()
        self.fuzzy_status_filter.addItems(["Все", "✅ Найдено", "❌ Не найдено"])
        self.fuzzy_status_filter.currentTextChanged.connect(self.filter_fuzzy_table)
        control_layout.addWidget(self.fuzzy_status_filter)

        control_layout.addWidget(QLabel("Тип совпадения:"))
        self.fuzzy_match_filter = QComboBox()
        self.fuzzy_match_filter.addItems(["Все", "Точное", "Все слова", "Большинство слов", "Нечеткое"])
        self.fuzzy_match_filter.currentTextChanged.connect(self.filter_fuzzy_table)
        control_layout.addWidget(self.fuzzy_match_filter)

        control_layout.addStretch()

        control_panel.setLayout(control_layout)
        layout.addWidget(control_panel)

        # Панель поиска
        search_panel = QWidget()
        search_layout = QHBoxLayout()
        search_layout.setContentsMargins(0, 0, 0, 0)

        search_layout.addWidget(QLabel("🔍 Фильтр по названию:"))
        self.fuzzy_search_input = QLineEdit()
        self.fuzzy_search_input.setPlaceholderText("Введите часть названия для фильтрации...")
        self.fuzzy_search_input.textChanged.connect(self.filter_fuzzy_table)
        search_layout.addWidget(self.fuzzy_search_input)

        search_layout.addWidget(QLabel("Мин. схожесть:"))
        self.fuzzy_min_similarity = QComboBox()
        self.fuzzy_min_similarity.addItems(["80%", "85%", "90%", "95%", "100%"])
        self.fuzzy_min_similarity.setCurrentText("80%")
        self.fuzzy_min_similarity.currentTextChanged.connect(self.filter_fuzzy_table)
        search_layout.addWidget(self.fuzzy_min_similarity)

        search_panel.setLayout(search_layout)
        layout.addWidget(search_panel)

        layout.addWidget(self.fuzzy_search_btn)
        layout.addWidget(self.fuzzy_progress)

        # Таблица результатов
        self.fuzzy_table = QTableWidget()
        self.fuzzy_table.setColumnCount(11)
        self.fuzzy_table.setHorizontalHeaderLabels([
            "✅", "Наименование ПО (из БД)", "Версия", "Подсистема",
            "ГК", "Схожесть", "Строка", "Тип совпадения",
            "Контекст", "ID", "Действия"
        ])

        # Настройка размеров
        header = self.fuzzy_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)

        self.fuzzy_table.setColumnWidth(0, 30)  # ✅
        self.fuzzy_table.setColumnWidth(1, 350)  # Наименование ПО
        self.fuzzy_table.setColumnWidth(2, 100)  # Версия
        self.fuzzy_table.setColumnWidth(3, 120)  # Подсистема
        self.fuzzy_table.setColumnWidth(4, 150)  # ГК
        self.fuzzy_table.setColumnWidth(5, 80)  # Схожесть
        self.fuzzy_table.setColumnWidth(6, 60)  # Строка
        self.fuzzy_table.setColumnWidth(7, 120)  # Тип совпадения
        self.fuzzy_table.setColumnWidth(8, 300)  # Контекст
        self.fuzzy_table.setColumnWidth(9, 50)  # ID
        self.fuzzy_table.setColumnWidth(10, 100)  # Действия

        self.fuzzy_table.setAlternatingRowColors(True)
        self.fuzzy_table.setSortingEnabled(True)
        self.fuzzy_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        layout.addWidget(self.fuzzy_table)

        # Кнопки экспорта
        export_layout = QHBoxLayout()

        export_csv_btn = QPushButton("📊 Экспорт в CSV")
        export_csv_btn.clicked.connect(self.export_fuzzy_to_csv)

        export_report_btn = QPushButton("📄 Сформировать отчет")
        export_report_btn.clicked.connect(self.generate_fuzzy_report)

        export_layout.addWidget(export_csv_btn)
        export_layout.addWidget(export_report_btn)
        export_layout.addStretch()

        layout.addLayout(export_layout)

        self.fuzzy_tab.setLayout(layout)

    def init_database_tab(self):
        """Инициализация вкладки управления базой данных"""
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Верхняя панель
        top_panel = QWidget()
        top_layout = QHBoxLayout()
        top_layout.setContentsMargins(0, 0, 0, 0)

        # Статистика БД
        stats = self.json_db.get_stats()
        stats_text = f"📊 Всего записей: {stats['total_products']} | "
        stats_text += f"✅ С сертификатом: {stats['with_certificate']} | "
        stats_text += f"❌ Без сертификата: {stats['without_certificate']}"

        self.db_stats_label = QLabel(stats_text)
        self.db_stats_label.setStyleSheet("font-weight: bold; color: #2c3e50;")

        # Кнопки управления
        refresh_btn = QPushButton("🔄 Обновить")
        refresh_btn.clicked.connect(self.refresh_database_tab)
        refresh_btn.setMinimumHeight(30)

        backup_btn = QPushButton("💾 Создать бэкап")
        backup_btn.clicked.connect(self.create_backup)
        backup_btn.setMinimumHeight(30)

        import_excel_btn = QPushButton("📥 Импорт из Excel")
        import_excel_btn.clicked.connect(self.import_from_excel)
        import_excel_btn.setMinimumHeight(30)

        top_layout.addWidget(self.db_stats_label)
        top_layout.addStretch()
        top_layout.addWidget(refresh_btn)
        top_layout.addWidget(backup_btn)
        top_layout.addWidget(import_excel_btn)

        top_panel.setLayout(top_layout)
        layout.addWidget(top_panel)

        # Панель фильтров
        filter_panel = QGroupBox("Фильтры")
        filter_layout = QGridLayout()

        # Поиск
        filter_layout.addWidget(QLabel("Поиск:"), 0, 0)
        self.db_search_input = QLineEdit()
        self.db_search_input.setPlaceholderText("Введите название продукта...")
        self.db_search_input.textChanged.connect(self.filter_database_table)
        filter_layout.addWidget(self.db_search_input, 0, 1, 1, 2)

        # Подсистема
        filter_layout.addWidget(QLabel("Подсистема:"), 1, 0)
        self.db_subsystem_filter = QComboBox()
        self.db_subsystem_filter.addItem("Все подсистемы")
        subsystems = self.json_db.get_subsystems()
        if subsystems:
            self.db_subsystem_filter.addItems(sorted(subsystems))
        self.db_subsystem_filter.currentTextChanged.connect(self.filter_database_table)
        filter_layout.addWidget(self.db_subsystem_filter, 1, 1)

        # Сертификат
        filter_layout.addWidget(QLabel("Сертификат:"), 1, 2)
        self.db_cert_filter = QComboBox()
        self.db_cert_filter.addItems(["Все", "С сертификатом", "Без сертификата"])
        self.db_cert_filter.currentTextChanged.connect(self.filter_database_table)
        filter_layout.addWidget(self.db_cert_filter, 1, 3)

        filter_panel.setLayout(filter_layout)
        layout.addWidget(filter_panel)

        # Таблица базы данных
        self.db_table = QTableWidget()
        self.db_table.setColumnCount(8)
        self.db_table.setHorizontalHeaderLabels([
            "ID", "Наименование", "Версия", "Подсистема", "ГК",
            "Сертификат", "Источник", "Действия"
        ])

        # Настройка размеров
        header = self.db_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)

        self.db_table.setColumnWidth(0, 50)  # ID
        self.db_table.setColumnWidth(1, 300)  # Наименование
        self.db_table.setColumnWidth(2, 120)  # Версия
        self.db_table.setColumnWidth(3, 120)  # Подсистема
        self.db_table.setColumnWidth(4, 150)  # ГК
        self.db_table.setColumnWidth(5, 100)  # Сертификат
        self.db_table.setColumnWidth(6, 150)  # Источник
        self.db_table.setColumnWidth(7, 100)  # Действия

        self.db_table.setAlternatingRowColors(True)
        self.db_table.setSortingEnabled(True)
        self.db_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        # Контекстное меню
        self.db_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.db_table.customContextMenuRequested.connect(self.show_db_context_menu)

        layout.addWidget(self.db_table)

        # Панель с бэкапами
        backup_panel = QGroupBox("Резервные копии")
        backup_layout = QVBoxLayout()

        self.backup_combo = QComboBox()
        self.backup_combo.setMinimumHeight(30)
        self.refresh_backups_list()

        backup_buttons = QHBoxLayout()
        restore_btn = QPushButton("↩ Восстановить из выбранной копии")
        restore_btn.clicked.connect(self.restore_from_backup)
        delete_backup_btn = QPushButton("🗑 Удалить выбранную копию")
        delete_backup_btn.clicked.connect(self.delete_backup)

        backup_buttons.addWidget(restore_btn)
        backup_buttons.addWidget(delete_backup_btn)
        backup_buttons.addStretch()

        backup_layout.addWidget(self.backup_combo)
        backup_layout.addLayout(backup_buttons)

        backup_panel.setLayout(backup_layout)
        layout.addWidget(backup_panel)

        # Устанавливаем layout для database_tab
        self.database_tab.setLayout(layout)

        # Загружаем данные
        self.load_database_table()

    def init_document_tab(self):
        """Инициализация вкладки с версиями из документа"""
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Верхняя панель с информацией
        top_panel = self.create_info_panel()
        layout.addWidget(top_panel)

        # Панель инструментов
        toolbar_layout = QHBoxLayout()

        export_csv_btn = QPushButton("📊 Экспорт в CSV")
        export_csv_btn.clicked.connect(self.export_to_csv)

        copy_all_btn = QPushButton("📋 Копировать все")
        copy_all_btn.clicked.connect(self.copy_all_versions)

        refresh_btn = QPushButton("🔄 Обновить")
        refresh_btn.clicked.connect(self.refresh_versions)

        self.clean_versions_btn = QPushButton("🧹 Очистить версии")
        self.clean_versions_btn.clicked.connect(self.clean_versions)
        self.clean_versions_btn.setStyleSheet("""
            QPushButton {
                background-color: #e67e22;
                color: white;
                border-radius: 3px;
                padding: 5px 10px;
            }
            QPushButton:hover {
                background-color: #d35400;
            }
        """)

        toolbar_layout.addWidget(export_csv_btn)
        toolbar_layout.addWidget(copy_all_btn)
        toolbar_layout.addWidget(self.clean_versions_btn)
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(refresh_btn)

        layout.addLayout(toolbar_layout)

        # Панель фильтров
        filter_layout = QHBoxLayout()

        self.doc_search_input = QLineEdit()
        self.doc_search_input.setPlaceholderText("🔍 Поиск по продукту или версии...")
        self.doc_search_input.textChanged.connect(self.filter_document_table)

        self.doc_category_filter = QComboBox()
        self.doc_category_filter.addItem("Все категории")
        self.doc_category_filter.currentTextChanged.connect(self.filter_document_table)

        self.doc_subsystem_filter = QComboBox()
        self.doc_subsystem_filter.addItem("Все подсистемы")
        self.doc_subsystem_filter.currentTextChanged.connect(self.filter_document_table)

        filter_layout.addWidget(QLabel("Поиск:"))
        filter_layout.addWidget(self.doc_search_input)
        filter_layout.addWidget(QLabel("Категория:"))
        filter_layout.addWidget(self.doc_category_filter)
        filter_layout.addWidget(QLabel("Подсистема:"))
        filter_layout.addWidget(self.doc_subsystem_filter)
        filter_layout.addStretch()

        layout.addLayout(filter_layout)

        # Таблица версий из документа
        self.doc_table = QTableWidget()
        self.doc_table.setColumnCount(9)  # Добавляем колонку для действий
        self.doc_table.setHorizontalHeaderLabels([
            "Категория", "Подсистема", "Продукт", "Версия", "Страница",
            "Строка", "Источник", "Контекст", "Действия"
        ])

        # Настройка размеров столбцов
        header = self.doc_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(True)

        self.doc_table.setColumnWidth(0, 120)
        self.doc_table.setColumnWidth(1, 100)
        self.doc_table.setColumnWidth(2, 250)
        self.doc_table.setColumnWidth(3, 100)
        self.doc_table.setColumnWidth(4, 80)
        self.doc_table.setColumnWidth(5, 80)
        self.doc_table.setColumnWidth(6, 80)
        self.doc_table.setColumnWidth(7, 300)
        self.doc_table.setColumnWidth(8, 100)

        self.doc_table.setAlternatingRowColors(True)
        self.doc_table.setSortingEnabled(True)
        self.doc_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        layout.addWidget(self.doc_table)

        # Статистика
        stats_layout = QHBoxLayout()
        self.doc_total_label = QLabel("Всего найдено версий: 0")
        self.doc_categories_label = QLabel("Категорий: 0")
        self.doc_unique_label = QLabel("Уникальных продуктов: 0")
        self.doc_valid_versions_label = QLabel("✅ Валидных версий: 0")
        self.doc_table_versions_label = QLabel("📊 Из таблиц: 0")

        stats_layout.addWidget(self.doc_total_label)
        stats_layout.addWidget(self.doc_categories_label)
        stats_layout.addWidget(self.doc_unique_label)
        stats_layout.addWidget(self.doc_valid_versions_label)
        stats_layout.addWidget(self.doc_table_versions_label)
        stats_layout.addStretch()

        layout.addLayout(stats_layout)

        self.document_tab.setLayout(layout)

    def init_comparison_tab(self):
        """Инициализация вкладки сравнения с базой данных"""
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Верхняя панель с информацией
        top_panel = self.create_info_panel()
        layout.addWidget(top_panel)

        # Панель управления
        control_panel = QGroupBox("Параметры сравнения")
        control_layout = QHBoxLayout()

        self.compare_btn = QPushButton("🔄 Сравнить с БД")
        self.compare_btn.clicked.connect(self.compare_with_database)
        self.compare_btn.setMinimumHeight(35)
        self.compare_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                font-weight: bold;
                border-radius: 5px;
                padding: 8px 15px;
            }
            QPushButton:hover {
                background-color: #229954;
            }
        """)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("🔍 Поиск по продукту...")
        self.search_input.textChanged.connect(self.filter_comparison_table)

        self.status_filter = QComboBox()
        self.status_filter.addItems(["Все статусы", "✅ Найдено", "❌ Не найдено", "⚠ Версия ниже требуемой"])
        self.status_filter.currentTextChanged.connect(self.filter_comparison_table)

        # Фильтр по ГК
        self.gk_filter = QComboBox()
        self.gk_filter.addItem("Все ГК")
        if self.gk_manager.first_pages_gk:
            self.gk_filter.addItems(self.gk_manager.first_pages_gk)
        self.gk_filter.currentTextChanged.connect(self.filter_comparison_table)

        # Фильтр по подсистеме
        self.subsystem_filter = QComboBox()
        self.subsystem_filter.addItem("Все подсистемы")
        subsystems = self.json_db.get_subsystems()
        self.subsystem_filter.addItems(sorted(subsystems))
        self.subsystem_filter.currentTextChanged.connect(self.filter_comparison_table)

        control_layout.addWidget(self.compare_btn)
        control_layout.addWidget(self.search_input)
        control_layout.addWidget(QLabel("Статус:"))
        control_layout.addWidget(self.status_filter)
        control_layout.addWidget(QLabel("ГК:"))
        control_layout.addWidget(self.gk_filter)
        control_layout.addWidget(QLabel("Подсистема:"))
        control_layout.addWidget(self.subsystem_filter)
        control_layout.addStretch()

        control_panel.setLayout(control_layout)
        layout.addWidget(control_panel)

        # Таблица результатов
        self.comparison_table = QTableWidget()
        self.comparison_table.setColumnCount(11)  # Добавляем колонку для действий
        self.comparison_table.setHorizontalHeaderLabels([
            "Продукт (из БД)", "Требуемая версия", "ГК (из БД)", "Подсистема",
            "Найдено в документе", "Версия в документе", "Источник",
            "Сравнение", "Статус", "В БД", "Действие"
        ])

        # Настройка размеров столбцов
        header = self.comparison_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)
        header.setStretchLastSection(False)

        self.comparison_table.setColumnWidth(0, 300)
        self.comparison_table.setColumnWidth(1, 150)
        self.comparison_table.setColumnWidth(2, 150)
        self.comparison_table.setColumnWidth(3, 100)
        self.comparison_table.setColumnWidth(4, 120)
        self.comparison_table.setColumnWidth(5, 120)
        self.comparison_table.setColumnWidth(6, 80)
        self.comparison_table.setColumnWidth(7, 150)
        self.comparison_table.setColumnWidth(8, 200)
        self.comparison_table.setColumnWidth(9, 50)
        self.comparison_table.setColumnWidth(10, 100)

        self.comparison_table.setAlternatingRowColors(True)
        self.comparison_table.setSortingEnabled(True)
        self.comparison_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        self.comparison_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.comparison_table.customContextMenuRequested.connect(self.show_comparison_context_menu)

        layout.addWidget(self.comparison_table)

        # Статистика
        stats_group = QGroupBox("Статистика")
        stats_layout = QHBoxLayout()

        self.total_products_label = QLabel("Всего продуктов в БД: 0")
        self.found_products_label = QLabel("✅ Найдено в документе: 0")
        self.not_found_label = QLabel("❌ Не найдено: 0")
        self.compliant_label = QLabel("👍 Версии соответствуют: 0")
        self.non_compliant_label = QLabel("👎 Версии ниже: 0")

        stats_layout.addWidget(self.total_products_label)
        stats_layout.addWidget(self.found_products_label)
        stats_layout.addWidget(self.not_found_label)
        stats_layout.addWidget(self.compliant_label)
        stats_layout.addWidget(self.non_compliant_label)
        stats_layout.addStretch()

        stats_group.setLayout(stats_layout)
        layout.addWidget(stats_group)

        self.comparison_tab.setLayout(layout)

    def create_info_panel(self):
        """Создание информационной панели"""
        panel = QWidget()
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        # Левая часть с основной информацией
        info_layout = QHBoxLayout()

        doc_info = QLabel(f"📄 Документ: {len(self.document_text)} символов")

        gk_text = f"🔑 Найдено ГК: {len(self.gk_manager.document_gk)}"
        if self.gk_manager.first_pages_gk:
            gk_text += f" (на первых стр: {len(self.gk_manager.first_pages_gk)})"

        gk_label = QLabel(gk_text)
        gk_label.setStyleSheet("color: #e67e22; font-weight: bold;")

        tables_text = f"📊 Найдено таблиц: {len(self.table_versions)}"
        tables_label = QLabel(tables_text)
        tables_label.setStyleSheet("color: #27ae60;")

        info_layout.addWidget(doc_info)
        info_layout.addWidget(gk_label)
        info_layout.addWidget(tables_label)
        info_layout.addStretch()

        # Правая часть - номер ГК
        gk_number_layout = QHBoxLayout()
        gk_number_layout.setSpacing(2)

        self.gk_number_label = QLabel("№")
        self.gk_number_label.setStyleSheet("""
            QLabel {
                color: #e67e22;
                font-weight: bold;
                font-size: 14px;
                background-color: #fff3e0;
                padding: 4px 8px;
                border: 1px solid #e67e22;
                border-radius: 4px;
            }
        """)

        self.gk_value_label = QLabel("")
        self.gk_value_label.setStyleSheet("""
            QLabel {
                color: #e67e22;
                font-weight: bold;
                font-size: 14px;
                background-color: #fff3e0;
                padding: 4px 8px;
                border: 1px solid #e67e22;
                border-radius: 4px;
                min-width: 150px;
            }
        """)

        self.update_gk_display()

        gk_number_layout.addWidget(self.gk_number_label)
        gk_number_layout.addWidget(self.gk_value_label)

        layout.addLayout(info_layout)
        layout.addStretch()
        layout.addLayout(gk_number_layout)

        panel.setLayout(layout)
        return panel

    def update_gk_display(self):
        """Обновление отображения ГК"""
        if self.gk_manager.document_gk:
            if self.gk_manager.first_pages_gk:
                primary_gk = self.gk_manager.first_pages_gk[0]
            else:
                primary_gk = self.gk_manager.document_gk[0]

            subsystem = self.gk_manager.get_subsystem_for_gk(primary_gk)
            self.gk_value_label.setText(f"{primary_gk} [{subsystem}]")

            if len(self.gk_manager.document_gk) > 1:
                tooltip_text = "Найденные ГК:\n"
                for gk in self.gk_manager.document_gk:
                    subs = self.gk_manager.get_subsystem_for_gk(gk)
                    tooltip_text += f"• {gk} [{subs}]\n"
                self.gk_value_label.setToolTip(tooltip_text)
        else:
            self.gk_value_label.setText("ГК не найдены")

    def scan_document_for_versions(self):
        """Поиск версий в документе"""
        self.versions_data = []
        lines = self.document_text.split('\n')

        # Список известных продуктов для контекстного поиска
        known_products = [
            'postgres', 'oracle', 'mysql', 'kafka', 'nginx', 'docker',
            'java', 'python', 'windows', 'linux', 'astra', 'ред ос',
            'ufos', 'уфос', 'svip', 'svbo', 'apache', 'tomcat',
            'elasticsearch', 'redis', 'mongodb', 'cassandra', 'clickhouse',
            'hadoop', 'spark', 'zookeeper', 'consul', 'deckhouse',
            'криптопро', 'universe', 'smart vista', 'arenadata', 'postgres pro',
            'red os', 'альт', 'debian', 'ubuntu', 'centos', 'freebsd',
            'wildfly', 'haproxy', 'etcd', 'prometheus', 'grafana',
            'filebeat', 'logstash', 'kibana', 'opensearch', 'zabbix'
        ]

        # Паттерны для поиска версий
        version_patterns = [
            r'(?:версия|ver\.|v\.|редакция)\s*[:\s]*([\d\.]+(?:[a-z]?\d*)?)',
            r'(\d+\.\d+\.\d+\.\d+(?:[.-]\d+)?)',
            r'(\d+\.\d+\.\d+(?:[.-]\d+)?)',
            r'(\d+\.\d+(?:[.-]\d+)?)',
            r'(\d+(?:\.\d+)?)',
        ]

        categories = {
            'windows': 'Операционные системы',
            'linux': 'Операционные системы',
            'astra': 'Операционные системы',
            'ред ос': 'Операционные системы',
            'red os': 'Операционные системы',
            'альт': 'Операционные системы',
            'debian': 'Операционные системы',
            'ubuntu': 'Операционные системы',
            'centos': 'Операционные системы',
            'postgres': 'Базы данных',
            'oracle': 'Базы данных',
            'mysql': 'Базы данных',
            'mariadb': 'Базы данных',
            'mongodb': 'Базы данных',
            'redis': 'Базы данных',
            'cassandra': 'Базы данных',
            'clickhouse': 'Базы данных',
            'kafka': 'Брокеры сообщений',
            'activemq': 'Брокеры сообщений',
            'rabbitmq': 'Брокеры сообщений',
            'docker': 'Контейнеризация',
            'kubernetes': 'Контейнеризация',
            'deckhouse': 'Контейнеризация',
            'podman': 'Контейнеризация',
            'nginx': 'Веб-серверы',
            'apache': 'Веб-серверы',
            'tomcat': 'Веб-серверы',
            'jetty': 'Веб-серверы',
            'haproxy': 'Балансировщики',
            'java': 'Языки программирования',
            'python': 'Языки программирования',
            'php': 'Языки программирования',
            'javascript': 'Языки программирования',
            'ufos': 'Прикладное ПО',
            'уфос': 'Прикладное ПО',
            'svip': 'Прикладное ПО',
            'svbo': 'Прикладное ПО',
            'криптопро': 'Криптография',
            'wildfly': 'Серверы приложений',
            'etcd': 'Распределенные системы',
            'consul': 'Распределенные системы',
            'zookeeper': 'Распределенные системы',
            'prometheus': 'Мониторинг',
            'grafana': 'Мониторинг',
            'zabbix': 'Мониторинг',
            'opensearch': 'Поисковые системы',
            'elasticsearch': 'Поисковые системы',
            'hadoop': 'Big Data',
            'spark': 'Big Data',
            'arenadata': 'Big Data'
        }

        def count_digits(version):
            return sum(c.isdigit() for c in version)

        for line_num, line in enumerate(lines):
            line_lower = line.lower()

            if VersionCleaner.is_garbage(line):
                continue

            has_product = False
            found_product = ""
            product_position = -1

            for product in known_products:
                pos = line_lower.find(product)
                if pos != -1:
                    has_product = True
                    found_product = product
                    product_position = pos
                    break

            if not has_product:
                continue

            category = 'Прочее ПО'
            for key, cat in categories.items():
                if key in line_lower:
                    category = cat
                    break

            line_after_product = line[product_position + len(found_product):]

            found_versions = []

            for pattern in version_patterns:
                matches = re.finditer(pattern, line_after_product, re.IGNORECASE)
                for match in matches:
                    if match.groups():
                        version = match.group(1).strip()
                        cleaned_version = VersionCleaner.clean_version(version)
                        if not cleaned_version or len(cleaned_version) < self.min_version_length:
                            continue
                        if VersionCleaner.is_garbage(cleaned_version):
                            continue
                        digit_count = count_digits(cleaned_version)
                        found_versions.append((cleaned_version, digit_count, match.start()))

            if found_versions:
                found_versions.sort(key=lambda x: (-x[1], x[2]))
                best_version = found_versions[0][0]

                version_pos_in_line = product_position + len(found_product) + found_versions[0][2]
                text_before_version = line[:version_pos_in_line].strip()
                words = text_before_version.split()
                if len(words) > 4:
                    product = ' '.join(words[-4:])
                else:
                    product = text_before_version

                if len(product) > 80:
                    product = product[:80] + "..."

                gk_pattern = self.settings.value(
                    "gk_format",
                    r"ФКУ\d{3,4}(?:[/-]\d{2,4})?(?:/\w+)?"
                )
                gk_matches = re.findall(gk_pattern, line, re.IGNORECASE)
                gk_list = [gk.upper() for gk in gk_matches if not VersionCleaner.is_garbage(gk)]

                subsystem = 'Не определена'
                if gk_list:
                    subsystem = self.gk_manager.get_subsystem_for_gk(gk_list[0])

                char_position = sum(len(l) + 1 for l in lines[:line_num]) + product_position

                self.versions_data.append({
                    'category': category,
                    'subsystem': subsystem,
                    'product': product,
                    'version': best_version,
                    'original_version': best_version,
                    'page': self.find_page_for_position(char_position),
                    'line': line_num + 1,
                    'context': line.strip()[:200],
                    'gk': gk_list,
                    'source': 'text',
                    'is_valid': VersionCleaner.is_valid_version(best_version),
                    'digit_count': count_digits(best_version)
                })

        # Добавляем версии из таблиц
        if self.extract_from_tables and self.table_versions:
            if self.table_priority:
                for tv in self.table_versions:
                    tv['category'] = self.determine_category(tv['name'])
                    subsystem = 'Не определена'
                    if tv.get('gk'):
                        subsystem = self.gk_manager.get_subsystem_for_gk(tv['gk'][0])
                    tv['subsystem'] = subsystem
                    tv['source'] = 'table'
                    tv['is_valid'] = VersionCleaner.is_valid_version(tv['version'])
                    tv['digit_count'] = count_digits(tv['version'])
                    self.versions_data.insert(0, tv)
            else:
                for tv in self.table_versions:
                    tv['category'] = self.determine_category(tv['name'])
                    subsystem = 'Не определена'
                    if tv.get('gk'):
                        subsystem = self.gk_manager.get_subsystem_for_gk(tv['gk'][0])
                    tv['subsystem'] = subsystem
                    tv['source'] = 'table'
                    tv['is_valid'] = VersionCleaner.is_valid_version(tv['version'])
                    tv['digit_count'] = count_digits(tv['version'])
                    self.versions_data.append(tv)

        # Убираем дубликаты
        self.versions_data = self.remove_duplicates_with_max_version(self.versions_data)
        self.versions_data.sort(key=lambda x: (-x.get('digit_count', 0), x.get('line', 0)))

    def remove_duplicates_with_max_version(self, data):
        """Удаление дубликатов, оставляя версию с максимальным количеством цифр"""
        product_dict = {}

        for item in data:
            if item.get('source') == 'table':
                key = (item.get('name', ''), 'table')
            else:
                key = (item.get('product', ''), 'text')

            if not key[0]:
                continue

            digit_count = sum(c.isdigit() for c in item.get('version', ''))

            if key not in product_dict or digit_count > product_dict[key].get('digit_count', 0):
                item['digit_count'] = digit_count
                product_dict[key] = item

        return list(product_dict.values())

    def determine_category(self, name):
        """Определение категории по названию"""
        if not name:
            return 'Прочее ПО'

        name_lower = name.lower()

        categories = {
            'Операционные системы': ['windows', 'linux', 'ubuntu', 'debian', 'centos', 'astra', 'ред ос', 'alt',
                                     'rosa'],
            'Базы данных': ['postgres', 'oracle', 'mysql', 'mariadb', 'mongodb', 'redis', 'cassandra', 'clickhouse'],
            'Виртуализация': ['vmware', 'virtualbox', 'docker', 'kubernetes', 'k8s', 'kvm'],
            'Веб-серверы': ['apache', 'nginx', 'iis', 'tomcat', 'jetty', 'haproxy'],
            'Языки программирования': ['python', 'java', 'php', 'node.js', 'ruby', 'go', '.net', 'javascript'],
            'Российское ПО': ['мойофис', 'р7', 'onlyoffice', '1с', 'vk', 'астра', 'ред ос', 'deckhouse'],
        }

        for category, keywords in categories.items():
            for keyword in keywords:
                if keyword in name_lower:
                    return category

        return 'Прочее ПО'

    def find_page_for_position(self, position):
        """Определение страницы по позиции"""
        if not self.page_info:
            return 1

        for page_num, (start, end) in enumerate(self.page_info, 1):
            if start <= position < end:
                return page_num

        return 1

    def display_versions(self):
        """Отображение версий из документа"""
        self.doc_table.setRowCount(len(self.versions_data))

        categories = set()
        subsystems = set()
        valid_count = 0
        table_count = 0

        for i, ver in enumerate(self.versions_data):
            # Категория
            self.doc_table.setItem(i, 0, QTableWidgetItem(ver['category']))
            categories.add(ver['category'])

            # Подсистема
            subsystem_item = QTableWidgetItem(ver.get('subsystem', 'Не определена'))
            if ver.get('subsystem') != 'Не определена':
                subsystem_item.setForeground(QColor(0, 100, 200))
            self.doc_table.setItem(i, 1, subsystem_item)
            subsystems.add(ver.get('subsystem', 'Не определена'))

            # Продукт
            product_name = ver.get('name', ver.get('product', ''))
            self.doc_table.setItem(i, 2, QTableWidgetItem(product_name))

            # Версия
            version_item = QTableWidgetItem(ver['version'])
            if ver.get('is_valid', False):
                version_item.setForeground(QColor(0, 150, 0))
                valid_count += 1
            else:
                version_item.setForeground(QColor(255, 0, 0))
            self.doc_table.setItem(i, 3, version_item)

            # Страница
            self.doc_table.setItem(i, 4, QTableWidgetItem(str(ver.get('page', 1))))

            # Строка
            line_num = ver.get('line', ver.get('line_num', 1))
            self.doc_table.setItem(i, 5, QTableWidgetItem(str(line_num)))

            # Источник
            source = ver.get('source', 'text')
            source_item = QTableWidgetItem("📊 Таблица" if source == 'table' else "📄 Текст")
            if source == 'table':
                source_item.setForeground(QColor(0, 100, 200))
                table_count += 1
            self.doc_table.setItem(i, 6, source_item)

            # Контекст
            context = ver.get('context', '')
            if ver.get('gk'):
                context += f" [ГК: {', '.join(ver['gk'])}]"
            self.doc_table.setItem(i, 7, QTableWidgetItem(context[:200]))

            # Кнопка добавления в БД
            add_btn = QPushButton("➕ В БД")
            add_btn.setProperty('row', i)
            add_btn.clicked.connect(lambda checked, r=i: self.add_version_to_db(r))
            add_btn.setMaximumWidth(60)
            self.doc_table.setCellWidget(i, 8, add_btn)

        # Обновляем фильтры
        self.doc_category_filter.clear()
        self.doc_category_filter.addItem("Все категории")
        self.doc_category_filter.addItems(sorted(categories))

        self.doc_subsystem_filter.clear()
        self.doc_subsystem_filter.addItem("Все подсистемы")
        self.doc_subsystem_filter.addItems(sorted([s for s in subsystems if s != 'Не определена'] + ['Не определена']))

        self.update_document_stats(valid_count, table_count)
        self.update_gk_display()

    def update_document_stats(self, valid_count=None, table_count=None):
        """Обновление статистики документа"""
        visible_rows = 0
        categories = set()
        products = set()
        valid = 0
        tables = 0

        for row in range(self.doc_table.rowCount()):
            if not self.doc_table.isRowHidden(row):
                visible_rows += 1
                cat_item = self.doc_table.item(row, 0)
                prod_item = self.doc_table.item(row, 2)
                source_item = self.doc_table.item(row, 6)
                version_item = self.doc_table.item(row, 3)

                if cat_item:
                    categories.add(cat_item.text())
                if prod_item:
                    products.add(prod_item.text())
                if source_item and "Таблица" in source_item.text():
                    tables += 1
                if version_item and version_item.foreground().color() == QColor(0, 150, 0):
                    valid += 1

        self.doc_total_label.setText(f"Всего найдено версий: {visible_rows} (из {len(self.versions_data)})")
        self.doc_categories_label.setText(f"Категорий: {len(categories)}")
        self.doc_unique_label.setText(f"Уникальных продуктов: {len(products)}")
        self.doc_valid_versions_label.setText(f"✅ Валидных версий: {valid}")
        self.doc_table_versions_label.setText(f"📊 Из таблиц: {tables}")

    def clean_versions(self):
        """Очистка версий от мусора"""
        cleaned_count = 0
        removed_count = 0

        for row in range(self.doc_table.rowCount()):
            version_item = self.doc_table.item(row, 3)
            if version_item:
                original_version = version_item.text()
                cleaned_version = VersionCleaner.clean_version(original_version)

                if cleaned_version and cleaned_version != original_version:
                    version_item.setText(cleaned_version)
                    version_item.setForeground(QColor(0, 150, 0))
                    cleaned_count += 1

                if not VersionCleaner.is_valid_version(cleaned_version if cleaned_version else original_version):
                    version_item.setForeground(QColor(255, 0, 0))
                    removed_count += 1

        self.update_document_stats()

        QMessageBox.information(
            self, "Очистка версий",
            f"Очищено версий: {cleaned_count}\n"
            f"Невалидных версий: {removed_count}"
        )

    def add_version_to_db(self, row):
        """Добавление версии из документа в базу данных"""
        if row >= len(self.versions_data):
            return

        version_data = self.versions_data[row]

        # Подготавливаем данные для диалога
        product_data = {
            'name': version_data.get('name', version_data.get('product', '')),
            'version': version_data.get('version', ''),
            'gk': version_data.get('gk', []),
            'subsystem': version_data.get('subsystem', 'Не определена'),
            'description': version_data.get('context', '')[:200],
            'source_file': f"Документ (стр. {version_data.get('page', 1)}, строка {version_data.get('line', 1)})"
        }

        # Открываем диалог редактирования
        subsystems = self.json_db.get_subsystems()
        dialog = ProductEditDialog(self, product_data, subsystems)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_product = dialog.get_data()
            if new_product:
                # Добавляем в БД
                product_id = self.json_db.add_product(new_product)
                if product_id:
                    QMessageBox.information(
                        self, "Успешно",
                        f"Продукт добавлен в базу данных (ID: {product_id})"
                    )
                    self.refresh_database_tab()
                else:
                    QMessageBox.warning(
                        self, "Ошибка",
                        "Не удалось добавить продукт в базу данных"
                    )

    def compare_with_database(self):
        """Сравнение документа с данными из базы данных"""
        db_products = self.json_db.get_all_products()

        if not db_products:
            QMessageBox.warning(self, "Ошибка", "Нет данных в базе данных для сравнения")
            return

        self.comparison_results = []
        self.comparison_table.setRowCount(0)

        for product in db_products:
            result = self.find_product_in_document(product)
            self.comparison_results.append(result)

        self.display_comparison_results()
        self.update_comparison_stats()

    def find_product_in_document(self, product):
        """Поиск продукта в документе"""
        result = {
            'product': product,
            'found': False,
            'matches': [],
            'best_version': None,
            'best_version_parsed': None,
            'comparison': None,
            'is_compliant': None,
            'matched_gk': [],
            'source': None,
            'subsystem': product.get('subsystem', 'Не определена'),
            'in_db': True
        }

        name = product['name'].lower()

        for ver in self.versions_data:
            ver_name = ver.get('name', ver.get('product', '')).lower()

            if name in ver_name or ver_name in name:
                result['found'] = True
                result['matches'].append({
                    'version': ver['version'],
                    'gk': ver.get('gk', []),
                    'source': ver.get('source', 'text'),
                    'line': ver.get('line', ver.get('line_num', 1))
                })

        if result['matches']:
            result['matches'].sort(key=lambda x: 0 if x['source'] == 'table' else 1)
            best = result['matches'][0]
            result['best_version'] = best['version']
            result['best_version_parsed'] = self.parse_version(best['version'])
            result['matched_gk'] = best.get('gk', [])
            result['source'] = best['source']

            required_version = product['version']
            if required_version:
                required_parsed = self.parse_version(required_version)
                comparison = self.compare_versions_parsed(
                    result['best_version_parsed'],
                    required_parsed
                )
                result['comparison'] = comparison
                result['is_compliant'] = comparison >= 0

        return result

    def parse_version(self, version_str):
        """Парсинг версии для сравнения"""
        if not version_str:
            return [0]

        version_str = str(version_str).lower().strip()
        version_str = re.sub(r'и выше.*$', '', version_str)
        version_str = re.sub(r'и новее.*$', '', version_str)
        version_str = re.sub(r'\([^)]*\)', '', version_str)
        version_str = re.sub(r'на базе.*$', '', version_str)

        numbers = re.findall(r'\d+', version_str)
        parts = []
        for num in numbers:
            try:
                parts.append(int(num))
            except ValueError:
                pass

        return parts if parts else [0]

    def compare_versions_parsed(self, v1_parts, v2_parts):
        """Сравнение версий"""
        if not v1_parts or not v2_parts:
            return 0

        max_len = max(len(v1_parts), len(v2_parts))
        v1 = v1_parts + [0] * (max_len - len(v1_parts))
        v2 = v2_parts + [0] * (max_len - len(v2_parts))

        for a, b in zip(v1, v2):
            if a > b:
                return 1
            elif a < b:
                return -1

        return 0

    def display_comparison_results(self):
        """Отображение результатов сравнения"""
        self.comparison_table.setRowCount(len(self.comparison_results))

        for i, result in enumerate(self.comparison_results):
            product = result['product']

            # Продукт
            name_item = QTableWidgetItem(product['name'])
            name_item.setData(Qt.ItemDataRole.UserRole, result)
            self.comparison_table.setItem(i, 0, name_item)

            # Требуемая версия
            self.comparison_table.setItem(i, 1, QTableWidgetItem(product['version'] or ''))

            # ГК из БД
            gk_text = ', '.join(product.get('gk', []))
            self.comparison_table.setItem(i, 2, QTableWidgetItem(gk_text))

            # Подсистема
            subsystem_item = QTableWidgetItem(result.get('subsystem', 'Не определена'))
            if result.get('subsystem') != 'Не определена':
                subsystem_item.setForeground(QColor(0, 100, 200))
            self.comparison_table.setItem(i, 3, subsystem_item)

            # Найдено в документе
            if result['found']:
                found_text = f"✅ Найдено ({len(result['matches'])})"
                found_color = QColor(0, 150, 0)

                if result['matched_gk']:
                    found_text += f"\n🔑 ГК: {', '.join(result['matched_gk'])}"
            else:
                found_text = "❌ Не найдено"
                found_color = QColor(255, 0, 0)

            found_item = QTableWidgetItem(found_text)
            found_item.setForeground(found_color)
            self.comparison_table.setItem(i, 4, found_item)

            # Версия в документе
            doc_version = result['best_version'] if result['best_version'] else ''
            doc_version_item = QTableWidgetItem(doc_version)
            if doc_version:
                doc_version_item.setForeground(QColor(0, 0, 150))
            self.comparison_table.setItem(i, 5, doc_version_item)

            # Источник
            source_text = "📊 Таблица" if result.get('source') == 'table' else "📄 Текст"
            source_item = QTableWidgetItem(source_text)
            if result.get('source') == 'table':
                source_item.setForeground(QColor(0, 100, 200))
            self.comparison_table.setItem(i, 6, source_item)

            # Сравнение
            if result['found'] and result['best_version']:
                if result['is_compliant'] is not None:
                    if result['is_compliant']:
                        comp_text = f"✅ {doc_version} >= {product['version']}"
                        comp_color = QColor(0, 150, 0)
                    else:
                        comp_text = f"❌ {doc_version} < {product['version']}"
                        comp_color = QColor(255, 0, 0)
                else:
                    comp_text = "⚠ Не удалось сравнить"
                    comp_color = QColor(255, 140, 0)
            else:
                comp_text = "—"
                comp_color = QColor(128, 128, 128)

            comp_item = QTableWidgetItem(comp_text)
            comp_item.setForeground(comp_color)
            self.comparison_table.setItem(i, 7, comp_item)

            # Статус
            if not result['found']:
                status = "❌ Не найдено в документе"
                status_color = QColor(255, 0, 0)
            elif result['is_compliant']:
                status = "✅ Версия соответствует требованиям"
                status_color = QColor(0, 150, 0)
            elif result['is_compliant'] is False:
                status = "⚠ Версия ниже требуемой"
                status_color = QColor(255, 140, 0)
            else:
                status = "🔍 Найдено (версия не определена)"
                status_color = QColor(0, 0, 255)

            status_item = QTableWidgetItem(status)
            status_item.setForeground(status_color)
            self.comparison_table.setItem(i, 8, status_item)

            # В БД (галочка)
            in_db_item = QTableWidgetItem("✅")
            in_db_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            in_db_item.setForeground(QColor(0, 150, 0))
            self.comparison_table.setItem(i, 9, in_db_item)

            # Кнопка редактирования
            edit_btn = QPushButton("✏️ Ред.")
            edit_btn.setProperty('row', i)
            edit_btn.clicked.connect(lambda checked, r=i: self.edit_product_from_comparison(r))
            edit_btn.setMaximumWidth(60)
            self.comparison_table.setCellWidget(i, 10, edit_btn)

        self.comparison_table.resizeRowsToContents()

    def edit_product_from_comparison(self, row):
        """Редактирование продукта из результатов сравнения"""
        if row >= len(self.comparison_results):
            return

        result = self.comparison_results[row]
        product = result['product']

        # Открываем диалог редактирования
        subsystems = self.json_db.get_subsystems()
        dialog = ProductEditDialog(self, product, subsystems)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            updated_product = dialog.get_data()
            if updated_product:
                # Обновляем в БД
                if self.json_db.update_product(product['id'], updated_product):
                    QMessageBox.information(
                        self, "Успешно",
                        f"Продукт обновлен"
                    )
                    # Обновляем результаты сравнения
                    self.comparison_results[row]['product'] = updated_product
                    self.display_comparison_results()
                    self.refresh_database_tab()
                else:
                    QMessageBox.warning(
                        self, "Ошибка",
                        "Не удалось обновить продукт"
                    )

    def filter_comparison_table(self):
        """Фильтрация таблицы сравнения"""
        search_text = self.search_input.text().lower()
        status_filter = self.status_filter.currentText()
        gk_filter = self.gk_filter.currentText()
        subsystem_filter = self.subsystem_filter.currentText()

        for row in range(self.comparison_table.rowCount()):
            show_row = True

            if search_text:
                name_item = self.comparison_table.item(row, 0)
                if name_item and search_text not in name_item.text().lower():
                    show_row = False

            if show_row and status_filter != "Все статусы":
                status_item = self.comparison_table.item(row, 8)
                if status_item:
                    if status_filter == "✅ Найдено" and "❌" in status_item.text():
                        show_row = False
                    elif status_filter == "❌ Не найдено" and "✅" in status_item.text():
                        show_row = False
                    elif status_filter == "⚠ Версия ниже требуемой" and "⚠" not in status_item.text():
                        show_row = False

            if show_row and gk_filter != "Все ГК":
                result = self.comparison_results[row]
                if gk_filter not in result.get('matched_gk', []):
                    show_row = False

            if show_row and subsystem_filter != "Все подсистемы":
                result = self.comparison_results[row]
                if result.get('subsystem') != subsystem_filter:
                    show_row = False

            self.comparison_table.setRowHidden(row, not show_row)

        self.update_comparison_stats()

    def update_comparison_stats(self):
        """Обновление статистики сравнения"""
        total = len(self.comparison_results)

        found = 0
        not_found = 0
        compliant = 0
        non_compliant = 0

        for row in range(self.comparison_table.rowCount()):
            if not self.comparison_table.isRowHidden(row):
                status_item = self.comparison_table.item(row, 8)
                if status_item:
                    status = status_item.text()
                    if "❌ Не найдено" in status:
                        not_found += 1
                    else:
                        found += 1
                        if "✅ Версия соответствует" in status:
                            compliant += 1
                        elif "⚠ Версия ниже" in status:
                            non_compliant += 1

        self.total_products_label.setText(f"Всего продуктов в БД: {total}")
        self.found_products_label.setText(f"✅ Найдено в документе: {found}")
        self.not_found_label.setText(f"❌ Не найдено: {not_found}")
        self.compliant_label.setText(f"👍 Версии соответствуют: {compliant}")
        self.non_compliant_label.setText(f"👎 Версии ниже: {non_compliant}")

    def load_database_table(self):
        """Загрузка данных в таблицу базы данных"""
        products = self.json_db.get_all_products()
        self.db_table.setRowCount(len(products))

        for i, product in enumerate(products):
            # ID
            id_item = QTableWidgetItem(str(product.get('id', '')))
            id_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.db_table.setItem(i, 0, id_item)

            # Наименование
            name_item = QTableWidgetItem(product.get('name', ''))
            self.db_table.setItem(i, 1, name_item)

            # Версия
            version_item = QTableWidgetItem(product.get('version', ''))
            self.db_table.setItem(i, 2, version_item)

            # Подсистема
            subsystem_item = QTableWidgetItem(product.get('subsystem', 'Не определена'))
            if product.get('subsystem') != 'Не определена':
                subsystem_item.setForeground(QColor(0, 100, 200))
            self.db_table.setItem(i, 3, subsystem_item)

            # ГК
            gk_list = product.get('gk', [])
            gk_text = ', '.join(gk_list) if gk_list else ''
            gk_item = QTableWidgetItem(gk_text)
            if gk_list:
                gk_item.setForeground(QColor(200, 100, 0))
            self.db_table.setItem(i, 4, gk_item)

            # Сертификат
            cert = product.get('certificate', '')
            cert_item = QTableWidgetItem(cert if cert else '—')
            if cert and 'да' in cert.lower():
                cert_item.setForeground(QColor(0, 150, 0))
            self.db_table.setItem(i, 5, cert_item)

            # Источник
            source_item = QTableWidgetItem(product.get('source_file', '')[:50])
            self.db_table.setItem(i, 6, source_item)

            # Кнопки действий
            actions_widget = QWidget()
            actions_layout = QHBoxLayout()
            actions_layout.setContentsMargins(2, 2, 2, 2)
            actions_layout.setSpacing(2)

            edit_btn = QPushButton("✏️")
            edit_btn.setMaximumWidth(30)
            edit_btn.setProperty('row', i)
            edit_btn.clicked.connect(lambda checked, r=i: self.edit_db_product(r))

            delete_btn = QPushButton("🗑️")
            delete_btn.setMaximumWidth(30)
            delete_btn.setProperty('row', i)
            delete_btn.clicked.connect(lambda checked, r=i: self.delete_db_product(r))

            actions_layout.addWidget(edit_btn)
            actions_layout.addWidget(delete_btn)

            actions_widget.setLayout(actions_layout)
            self.db_table.setCellWidget(i, 7, actions_widget)

    def edit_db_product(self, row):
        """Редактирование продукта в базе данных"""
        product_id = int(self.db_table.item(row, 0).text())
        product = self.json_db.get_product(product_id)

        if not product:
            QMessageBox.warning(self, "Ошибка", "Продукт не найден")
            return

        subsystems = self.json_db.get_subsystems()
        dialog = ProductEditDialog(self, product, subsystems)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            updated_product = dialog.get_data()
            if updated_product:
                if self.json_db.update_product(product_id, updated_product):
                    QMessageBox.information(self, "Успешно", "Продукт обновлен")
                    self.load_database_table()
                    self.update_db_stats()
                else:
                    QMessageBox.warning(self, "Ошибка", "Не удалось обновить продукт")

    def delete_db_product(self, row):
        """Удаление продукта из базы данных"""
        product_id = int(self.db_table.item(row, 0).text())
        product_name = self.db_table.item(row, 1).text()

        reply = QMessageBox.question(
            self, "Подтверждение",
            f"Удалить продукт '{product_name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            if self.json_db.delete_product(product_id, hard_delete=False):
                QMessageBox.information(self, "Успешно", "Продукт удален")
                self.load_database_table()
                self.update_db_stats()
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось удалить продукт")

    def filter_database_table(self):
        """Фильтрация таблицы базы данных"""
        search_text = self.db_search_input.text().lower()
        subsystem = self.db_subsystem_filter.currentText()
        cert_filter = self.db_cert_filter.currentText()

        for row in range(self.db_table.rowCount()):
            show_row = True

            if search_text:
                name_item = self.db_table.item(row, 1)
                if name_item and search_text not in name_item.text().lower():
                    show_row = False

            if show_row and subsystem != "Все подсистемы":
                subs_item = self.db_table.item(row, 3)
                if subs_item and subs_item.text() != subsystem:
                    show_row = False

            if show_row and cert_filter != "Все":
                cert_item = self.db_table.item(row, 5)
                has_cert = cert_item and cert_item.text() != '—'

                if cert_filter == "С сертификатом" and not has_cert:
                    show_row = False
                elif cert_filter == "Без сертификата" and has_cert:
                    show_row = False

            self.db_table.setRowHidden(row, not show_row)

    def update_db_stats(self):
        """Обновление статистики базы данных"""
        stats = self.json_db.get_stats()
        stats_text = f"📊 Всего записей: {stats['total_products']} | "
        stats_text += f"✅ С сертификатом: {stats['with_certificate']} | "
        stats_text += f"❌ Без сертификата: {stats['without_certificate']}"
        self.db_stats_label.setText(stats_text)

    def refresh_database_tab(self):
        """Обновление вкладки базы данных"""
        self.load_database_table()
        self.update_db_stats()
        self.refresh_backups_list()
        # Обновляем фильтр подсистем на вкладке нечеткого поиска
        self.update_fuzzy_subsystem_filter()

    def refresh_backups_list(self):
        """Обновление списка резервных копий"""
        self.backup_combo.clear()
        backups = self.json_db.get_backups_list()

        if backups:
            for backup in backups:
                created = backup.get('created_at', 'unknown')
                reason = backup.get('reason', 'manual')
                products = backup.get('total_products', '?')
                self.backup_combo.addItem(
                    f"{created} - {reason} ({products} записей)",
                    backup.get('path')
                )
        else:
            self.backup_combo.addItem("Нет доступных резервных копий")

    def create_backup(self):
        """Создание резервной копии"""
        backup_path = self.json_db.create_backup('manual')
        if backup_path:
            QMessageBox.information(
                self, "Успешно",
                f"Резервная копия создана:\n{backup_path}"
            )
            self.refresh_backups_list()
        else:
            QMessageBox.warning(self, "Ошибка", "Не удалось создать резервную копию")

    def restore_from_backup(self):
        """Восстановление из резервной копии"""
        if self.backup_combo.count() == 0 or self.backup_combo.currentText() == "Нет доступных резервных копий":
            QMessageBox.warning(self, "Ошибка", "Нет доступных резервных копий")
            return

        backup_path = self.backup_combo.currentData()
        if not backup_path:
            return

        reply = QMessageBox.question(
            self, "Подтверждение",
            "Восстановление из резервной копии приведет к потере текущих данных. Продолжить?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            if self.json_db.restore_from_backup(backup_path):
                QMessageBox.information(self, "Успешно", "Данные восстановлены из резервной копии")
                self.refresh_database_tab()
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось восстановить данные")

    def delete_backup(self):
        """Удаление резервной копии"""
        if self.backup_combo.count() == 0 or self.backup_combo.currentText() == "Нет доступных резервных копий":
            return

        backup_path = self.backup_combo.currentData()
        if not backup_path:
            return

        reply = QMessageBox.question(
            self, "Подтверждение",
            f"Удалить резервную копию?\n{backup_path}",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            import shutil
            try:
                shutil.rmtree(backup_path)
                QMessageBox.information(self, "Успешно", "Резервная копия удалена")
                self.refresh_backups_list()
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось удалить резервную копию: {e}")

    def import_from_excel(self):
        """Импорт данных из Excel файла"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите файл Excel", "",
            "Excel files (*.xlsx *.xls)"
        )

        if not file_path:
            return

        try:
            # Парсим Excel файл
            data = self.excel_parser.parse_excel_file(file_path)

            if data:
                self.excel_parser.print_summary()

                reply = QMessageBox.question(
                    self, "Импорт завершен",
                    f"Импортировано {self.excel_parser.stats['total_products']} продуктов.\n"
                    f"Ошибок: {len(self.excel_parser.get_errors())}\n\n"
                    "Обновить отображение базы данных?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )

                if reply == QMessageBox.StandardButton.Yes:
                    self.refresh_database_tab()

        except Exception as e:
            QMessageBox.warning(self, "Ошибка импорта", str(e))

    def show_db_context_menu(self, position):
        """Контекстное меню для таблицы базы данных"""
        menu = QMenu()

        add_action = QAction("➕ Добавить продукт", self)
        add_action.triggered.connect(self.add_new_product)

        refresh_action = QAction("🔄 Обновить", self)
        refresh_action.triggered.connect(self.refresh_database_tab)

        export_action = QAction("📤 Экспорт в JSON", self)
        export_action.triggered.connect(self.export_db_to_json)

        menu.addAction(add_action)
        menu.addSeparator()
        menu.addAction(refresh_action)
        menu.addAction(export_action)

        menu.exec(self.db_table.viewport().mapToGlobal(position))

    def add_new_product(self):
        """Добавление нового продукта в БД"""
        subsystems = self.json_db.get_subsystems()
        dialog = ProductEditDialog(self, None, subsystems)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            new_product = dialog.get_data()
            if new_product:
                product_id = self.json_db.add_product(new_product)
                if product_id:
                    QMessageBox.information(
                        self, "Успешно",
                        f"Продукт добавлен в базу данных (ID: {product_id})"
                    )
                    self.refresh_database_tab()
                else:
                    QMessageBox.warning(
                        self, "Ошибка",
                        "Не удалось добавить продукт в базу данных"
                    )

    def export_db_to_json(self):
        """Экспорт базы данных в JSON файл"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить как JSON", "products_export.json",
            "JSON files (*.json)"
        )

        if file_path:
            if self.json_db.export_to_file(file_path, 'json'):
                QMessageBox.information(self, "Успешно", f"Данные экспортированы в {file_path}")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось экспортировать данные")

    def show_comparison_context_menu(self, position):
        """Контекстное меню для таблицы сравнения"""
        menu = QMenu()

        copy_name_action = QAction("📋 Копировать название", self)
        copy_name_action.triggered.connect(self.copy_comparison_name)

        show_details_action = QAction("🔍 Показать детали", self)
        show_details_action.triggered.connect(self.show_details_for_selected)

        menu.addAction(copy_name_action)
        menu.addSeparator()
        menu.addAction(show_details_action)

        menu.exec(self.comparison_table.viewport().mapToGlobal(position))

    def copy_comparison_name(self):
        """Копирование названия продукта"""
        current_row = self.comparison_table.currentRow()
        if current_row >= 0:
            name_item = self.comparison_table.item(current_row, 0)
            if name_item:
                QApplication.clipboard().setText(name_item.text())
                QToolTip.showText(self.mapToGlobal(self.rect().center()), "Название скопировано")

    def show_details_for_selected(self):
        """Показ деталей для выбранной строки"""
        current_row = self.comparison_table.currentRow()
        if current_row >= 0:
            self.show_product_details(current_row)

    def show_product_details(self, row):
        """Показ деталей продукта"""
        result = self.comparison_results[row]
        product = result['product']

        details = f"<h3>Детали продукта</h3>"
        details += f"<p><b>Наименование:</b> {product['name']}</p>"
        details += f"<p><b>Требуемая версия:</b> {product['version']}</p>"
        details += f"<p><b>Подсистема:</b> {result.get('subsystem', 'Не определена')}</p>"

        if product.get('description'):
            details += f"<p><b>Описание:</b> {product['description']}</p>"

        if product.get('gk'):
            details += f"<p><b>ГК в БД:</b> {', '.join(product['gk'])}</p>"

        if result['matched_gk']:
            details += f"<p><b>Найденные ГК:</b> {', '.join(result['matched_gk'])}</p>"

        if result['found']:
            details += f"<h4>Найдено в документе:</h4>"
            details += f"<p><b>Версия:</b> {result['best_version']}</p>"
            details += f"<p><b>Источник:</b> {'Таблица' if result.get('source') == 'table' else 'Текст'}</p>"

            if result['is_compliant'] is not None:
                if result['is_compliant']:
                    details += f"<p><b style='color:green'>✅ Версия соответствует требованиям (>= {product['version']})</b></p>"
                else:
                    details += f"<p><b style='color:red'>❌ Версия ниже требуемой ({result['best_version']} < {product['version']})</b></p>"
            else:
                details += f"<p><b style='color:orange'>⚠ Не удалось сравнить версии</b></p>"

            details += "<h4>Все совпадения:</h4><ul>"
            for match in result['matches'][:5]:
                details += f"<li>Строка {match.get('line', '?')}: версия {match['version']}"
                if match.get('gk'):
                    details += f" [ГК: {', '.join(match['gk'])}]"
                if match.get('source') == 'table':
                    details += " (из таблицы)"
                details += "</li>"
            if len(result['matches']) > 5:
                details += f"<li>... и еще {len(result['matches']) - 5} совпадений</li>"
            details += "</ul>"
        else:
            details += "<p><b style='color:red'>❌ Продукт не найден в документе</b></p>"

        msg = QMessageBox(self)
        msg.setWindowTitle(f"Детали: {product['name']}")
        msg.setText(details)
        msg.setTextFormat(Qt.TextFormat.RichText)
        msg.setMinimumWidth(600)
        msg.exec()

    def copy_all_versions(self):
        """Копирование всех версий в буфер обмена"""
        text = "Категория\tПодсистема\tПродукт\tВерсия\tСтраница\tСтрока\tИсточник\tКонтекст\n"

        for row in range(self.doc_table.rowCount()):
            if not self.doc_table.isRowHidden(row):
                row_text = []
                for col in range(7):  # Без колонки действий
                    item = self.doc_table.item(row, col)
                    if item:
                        row_text.append(item.text())
                    else:
                        row_text.append("")
                text += '\t'.join(row_text) + '\n'

        QApplication.clipboard().setText(text)
        QMessageBox.information(self, "Копирование", f"Скопировано {self.get_visible_doc_rows()} строк")

    def get_visible_doc_rows(self):
        """Получение количества видимых строк"""
        count = 0
        for row in range(self.doc_table.rowCount()):
            if not self.doc_table.isRowHidden(row):
                count += 1
        return count

    def filter_document_table(self):
        """Фильтрация таблицы документа"""
        search_text = self.doc_search_input.text().lower()
        category = self.doc_category_filter.currentText()
        subsystem = self.doc_subsystem_filter.currentText()

        for row in range(self.doc_table.rowCount()):
            show_row = True

            if search_text:
                row_text = ""
                for col in [0, 1, 2, 3, 7]:  # Категория, Подсистема, Продукт, Версия, Контекст
                    item = self.doc_table.item(row, col)
                    if item:
                        row_text += item.text().lower() + " "

                if search_text not in row_text:
                    show_row = False

            if show_row and category != "Все категории":
                cat_item = self.doc_table.item(row, 0)
                if cat_item and cat_item.text() != category:
                    show_row = False

            if show_row and subsystem != "Все подсистемы":
                subsystem_item = self.doc_table.item(row, 1)
                if subsystem_item and subsystem_item.text() != subsystem:
                    show_row = False

            self.doc_table.setRowHidden(row, not show_row)

        self.update_document_stats()

    def export_to_csv(self):
        """Экспорт в CSV"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить как CSV", f"versions_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "CSV files (*.csv)"
        )

        if file_path:
            try:
                import csv
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(
                        ["Категория", "Подсистема", "Продукт", "Версия", "Страница", "Строка", "Источник", "Контекст",
                         "ГК"]
                    )

                    for row in range(self.doc_table.rowCount()):
                        if not self.doc_table.isRowHidden(row):
                            row_data = []
                            for col in range(7):
                                item = self.doc_table.item(row, col)
                                if item:
                                    row_data.append(item.text())
                                else:
                                    row_data.append("")

                            version_data = self.versions_data[row] if row < len(self.versions_data) else {}
                            gk_text = ', '.join(version_data.get('gk', []))
                            row_data.append(gk_text)

                            writer.writerow(row_data)

                QMessageBox.information(self, "Успех", f"Данные экспортированы в {file_path}")
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось экспортировать: {str(e)}")

    def refresh_versions(self):
        """Обновление списка версий"""
        self.scan_document_for_versions()
        self.display_versions()
        self.update_gk_display()
        QToolTip.showText(self.mapToGlobal(self.rect().center()), f"Найдено {len(self.versions_data)} версий")

    def build_category_keywords(self):
        """Построение ключевых слов для категорий"""
        return {}

    def get_software_patterns(self):
        """Получение паттернов для поиска ПО"""
        return {}

    # === Методы для вкладки нечеткого поиска ===

    def start_fuzzy_search(self):
        """Запуск нечеткого поиска с порогом 80%"""
        db_products = self.json_db.get_all_products()

        if not db_products:
            QMessageBox.warning(self, "Ошибка", "Нет данных в базе данных для поиска")
            return

        # Блокируем кнопку
        self.fuzzy_search_btn.setEnabled(False)
        self.fuzzy_search_btn.setText("⏳ Поиск...")

        # Показываем прогресс
        self.fuzzy_progress.setVisible(True)
        self.fuzzy_progress.setValue(0)

        # Запускаем поток с фиксированным порогом 80%
        self.fuzzy_thread = FuzzySearchThread(
            self.document_text,
            db_products,
            0.8  # фиксированный порог 80%
        )
        self.fuzzy_thread.progress.connect(self.fuzzy_progress.setValue)
        self.fuzzy_thread.result_ready.connect(self.display_fuzzy_results)
        self.fuzzy_thread.finished.connect(self.on_fuzzy_search_finished)
        self.fuzzy_thread.start()

    def on_fuzzy_search_finished(self):
        """Обработка завершения поиска"""
        self.fuzzy_search_btn.setEnabled(True)
        self.fuzzy_search_btn.setText("🔍 Найти СПО/БПО в документе")
        self.fuzzy_progress.setVisible(False)

    def display_fuzzy_results(self, results):
        """Отображение результатов нечеткого поиска"""
        self.fuzzy_search_results = results
        self.fuzzy_table.setRowCount(len(results))

        found_count = 0

        for i, result in enumerate(results):
            product = result['product']
            found = result['found']
            match_info = result.get('match_info')

            # Статус
            status_item = QTableWidgetItem("✅" if found else "❌")
            status_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            if found:
                status_item.setForeground(QColor(0, 150, 0))
                found_count += 1
            else:
                status_item.setForeground(QColor(255, 0, 0))
            self.fuzzy_table.setItem(i, 0, status_item)

            # Наименование ПО
            self.fuzzy_table.setItem(i, 1, QTableWidgetItem(product['name']))

            # Версия
            self.fuzzy_table.setItem(i, 2, QTableWidgetItem(product.get('version', '')))

            # Подсистема (из базы данных)
            subsystem_item = QTableWidgetItem(product.get('subsystem', 'Не определена'))
            if product.get('subsystem') != 'Не определена':
                subsystem_item.setForeground(QColor(0, 100, 200))
            self.fuzzy_table.setItem(i, 3, subsystem_item)

            # ГК
            gk_text = ', '.join(product.get('gk', []))
            gk_item = QTableWidgetItem(gk_text)
            if gk_text:
                gk_item.setForeground(QColor(200, 100, 0))
            self.fuzzy_table.setItem(i, 4, gk_item)

            # Схожесть
            if found and match_info:
                similarity = int(match_info[0] * 100)
                similarity_item = QTableWidgetItem(f"{similarity}%")
                if similarity >= 95:
                    similarity_item.setForeground(QColor(0, 150, 0))
                elif similarity >= 80:
                    similarity_item.setForeground(QColor(255, 140, 0))
                else:
                    similarity_item.setForeground(QColor(255, 0, 0))
                self.fuzzy_table.setItem(i, 5, similarity_item)
            else:
                self.fuzzy_table.setItem(i, 5, QTableWidgetItem("—"))

            # Строка
            if found and match_info:
                self.fuzzy_table.setItem(i, 6, QTableWidgetItem(str(match_info[1])))
            else:
                self.fuzzy_table.setItem(i, 6, QTableWidgetItem("—"))

            # Тип совпадения
            type_display = {
                'exact': 'Точное',
                'all_words': 'Все слова',
                'most_words': 'Большинство слов',
                'fuzzy_word': 'Нечеткое (слово)',
                'fuzzy': 'Нечеткое',
                'fuzzy_multiple': 'Нечеткое (несколько)'
            }

            if found and match_info and len(match_info) > 3:
                match_type = match_info[3]
                type_text = type_display.get(match_type, match_type)
                type_item = QTableWidgetItem(type_text)

                if match_type == 'exact' or match_type == 'all_words':
                    type_item.setForeground(QColor(0, 150, 0))
                elif match_type == 'most_words':
                    type_item.setForeground(QColor(0, 100, 200))
                else:
                    type_item.setForeground(QColor(255, 140, 0))

                self.fuzzy_table.setItem(i, 7, type_item)
            else:
                self.fuzzy_table.setItem(i, 7, QTableWidgetItem("—"))

            # Контекст
            if found and match_info and len(match_info) > 2:
                context_item = QTableWidgetItem(match_info[2])
                self.fuzzy_table.setItem(i, 8, context_item)
            else:
                self.fuzzy_table.setItem(i, 8, QTableWidgetItem(""))

            # ID
            id_item = QTableWidgetItem(str(product.get('id', '')))
            id_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.fuzzy_table.setItem(i, 9, id_item)

            # Кнопки действий
            actions_widget = QWidget()
            actions_layout = QHBoxLayout()
            actions_layout.setContentsMargins(2, 2, 2, 2)
            actions_layout.setSpacing(2)

            edit_btn = QPushButton("✏️")
            edit_btn.setMaximumWidth(30)
            edit_btn.setProperty('row', i)
            edit_btn.clicked.connect(lambda checked, r=i: self.edit_fuzzy_product(r))

            goto_btn = QPushButton("🔍")
            goto_btn.setMaximumWidth(30)
            goto_btn.setProperty('row', i)
            goto_btn.clicked.connect(lambda checked, r=i: self.goto_fuzzy_location(r))

            actions_layout.addWidget(edit_btn)
            actions_layout.addWidget(goto_btn)

            actions_widget.setLayout(actions_layout)
            self.fuzzy_table.setCellWidget(i, 10, actions_widget)

        self.fuzzy_table.resizeRowsToContents()

        # Обновляем статистику
        total = len(results)
        not_found = total - found_count
        self.fuzzy_stats_label.setText(
            f"📊 Всего продуктов в БД: {total} | "
            f"✅ Найдено в документе: {found_count} | "
            f"❌ Не найдено: {not_found}"
        )

        # Обновляем фильтры
        self.update_fuzzy_subsystem_filter()

    def update_fuzzy_subsystem_filter(self):
        """Обновление фильтра подсистем на вкладке нечеткого поиска"""
        current_text = self.fuzzy_subsystem_filter.currentText()
        self.fuzzy_subsystem_filter.clear()
        self.fuzzy_subsystem_filter.addItem("Все подсистемы")

        subsystems = set()
        for result in self.fuzzy_search_results:
            subs = result['product'].get('subsystem', 'Не определена')
            subsystems.add(subs)

        for subs in sorted(subsystems):
            self.fuzzy_subsystem_filter.addItem(subs)

        # Восстанавливаем выбранное значение
        index = self.fuzzy_subsystem_filter.findText(current_text)
        if index >= 0:
            self.fuzzy_subsystem_filter.setCurrentIndex(index)

    def filter_fuzzy_table(self):
        """Фильтрация таблицы нечеткого поиска"""
        search_text = self.fuzzy_search_input.text().lower()
        subsystem = self.fuzzy_subsystem_filter.currentText()
        status = self.fuzzy_status_filter.currentText()
        match_type = self.fuzzy_match_filter.currentText()

        # Получаем минимальный процент схожести
        min_sim_text = self.fuzzy_min_similarity.currentText().replace('%', '')
        min_similarity = int(min_sim_text) if min_sim_text else 80

        for row in range(self.fuzzy_table.rowCount()):
            show_row = True

            # Фильтр по названию
            if search_text:
                product_item = self.fuzzy_table.item(row, 1)
                if product_item and search_text not in product_item.text().lower():
                    show_row = False

            # Фильтр по подсистеме
            if show_row and subsystem != "Все подсистемы":
                subs_item = self.fuzzy_table.item(row, 3)
                if subs_item and subs_item.text() != subsystem:
                    show_row = False

            # Фильтр по статусу
            if show_row and status != "Все":
                status_item = self.fuzzy_table.item(row, 0)
                if status_item:
                    if status == "✅ Найдено" and status_item.text() != "✅":
                        show_row = False
                    elif status == "❌ Не найдено" and status_item.text() != "❌":
                        show_row = False

            # Фильтр по типу совпадения
            if show_row and match_type != "Все":
                type_item = self.fuzzy_table.item(row, 7)
                if type_item:
                    type_text = type_item.text()
                    if match_type == "Точное" and type_text not in ["Точное", "Все слова"]:
                        show_row = False
                    elif match_type == "Все слова" and type_text != "Все слова":
                        show_row = False
                    elif match_type == "Большинство слов" and type_text != "Большинство слов":
                        show_row = False
                    elif match_type == "Нечеткое" and "Нечеткое" not in type_text:
                        show_row = False

            # Фильтр по минимальной схожести
            if show_row and min_similarity > 0:
                sim_item = self.fuzzy_table.item(row, 5)
                if sim_item and sim_item.text() != "—":
                    sim_value = int(sim_item.text().replace('%', ''))
                    if sim_value < min_similarity:
                        show_row = False

            self.fuzzy_table.setRowHidden(row, not show_row)

    def edit_fuzzy_product(self, row):
        """Редактирование продукта из результатов нечеткого поиска"""
        if row >= len(self.fuzzy_search_results):
            return

        result = self.fuzzy_search_results[row]
        product = result['product']

        # Открываем диалог редактирования
        subsystems = self.json_db.get_subsystems()
        dialog = ProductEditDialog(self, product, subsystems)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            updated_product = dialog.get_data()
            if updated_product:
                # Обновляем в БД
                if self.json_db.update_product(product['id'], updated_product):
                    QMessageBox.information(
                        self, "Успешно",
                        f"Продукт обновлен"
                    )
                    # Обновляем результаты поиска
                    self.fuzzy_search_results[row]['product'] = updated_product
                    self.display_fuzzy_results(self.fuzzy_search_results)
                    self.refresh_database_tab()
                else:
                    QMessageBox.warning(
                        self, "Ошибка",
                        "Не удалось обновить продукт"
                    )

    def goto_fuzzy_location(self, row):
        """Переход к месту в документе, где найдено совпадение"""
        if row >= len(self.fuzzy_search_results):
            return

        result = self.fuzzy_search_results[row]
        if not result['found']:
            QMessageBox.information(self, "Информация", "Продукт не найден в документе")
            return

        match_info = result.get('match_info')
        if not match_info or len(match_info) < 2:
            return

        line_num = match_info[1]

        # Переключаемся на вкладку с документом
        self.tab_widget.setCurrentIndex(0)  # Вкладка документа

        # Ищем строку в таблице документа
        for doc_row in range(self.doc_table.rowCount()):
            line_item = self.doc_table.item(doc_row, 5)  # Колонка "Строка"
            if line_item and line_item.text() == str(line_num):
                self.doc_table.selectRow(doc_row)
                self.doc_table.scrollToItem(line_item)

                # Подсвечиваем строку
                for col in range(self.doc_table.columnCount()):
                    item = self.doc_table.item(doc_row, col)
                    if item:
                        item.setBackground(QColor(255, 255, 0, 100))

                # Сбрасываем подсветку через 2 секунды
                QTimer.singleShot(2000, lambda: self.clear_row_highlight(doc_row))
                break

    def clear_row_highlight(self, row):
        """Сброс подсветки строки"""
        for col in range(self.doc_table.columnCount()):
            item = self.doc_table.item(row, col)
            if item:
                item.setBackground(QColor(255, 255, 255))

    def export_fuzzy_to_csv(self):
        """Экспорт результатов нечеткого поиска в CSV"""
        if not self.fuzzy_search_results:
            QMessageBox.warning(self, "Ошибка", "Нет данных для экспорта")
            return

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить как CSV", f"fuzzy_search_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "CSV files (*.csv)"
        )

        if file_path:
            try:
                import csv
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow([
                        "Статус", "Наименование ПО", "Версия", "Подсистема", "ГК",
                        "Схожесть", "Строка", "Тип совпадения", "Контекст", "ID"
                    ])

                    for result in self.fuzzy_search_results:
                        product = result['product']
                        found = result['found']
                        match_info = result.get('match_info')

                        type_display = {
                            'exact': 'Точное',
                            'all_words': 'Все слова',
                            'most_words': 'Большинство слов',
                            'fuzzy_word': 'Нечеткое (слово)',
                            'fuzzy': 'Нечеткое',
                            'fuzzy_multiple': 'Нечеткое (несколько)'
                        }

                        row_data = [
                            "Найдено" if found else "Не найдено",
                            product['name'],
                            product.get('version', ''),
                            product.get('subsystem', 'Не определена'),
                            ', '.join(product.get('gk', []))
                        ]

                        if found and match_info:
                            similarity = int(match_info[0] * 100)
                            match_type = match_info[3] if len(match_info) > 3 else 'unknown'
                            type_text = type_display.get(match_type, match_type)

                            row_data.extend([
                                f"{similarity}%",
                                str(match_info[1]),
                                type_text,
                                match_info[2] if len(match_info) > 2 else '',
                                str(product.get('id', ''))
                            ])
                        else:
                            row_data.extend(['—', '—', '—', '', str(product.get('id', ''))])

                        writer.writerow(row_data)

                QMessageBox.information(self, "Успех", f"Данные экспортированы в {file_path}")
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось экспортировать: {str(e)}")

    def generate_fuzzy_report(self):
        """Генерация отчета по результатам нечеткого поиска"""
        if not self.fuzzy_search_results:
            QMessageBox.warning(self, "Ошибка", "Нет данных для формирования отчета")
            return

        total = len(self.fuzzy_search_results)
        found = sum(1 for r in self.fuzzy_search_results if r['found'])
        not_found = total - found

        # Группировка по подсистемам
        by_subsystem = {}
        for result in self.fuzzy_search_results:
            subs = result['product'].get('subsystem', 'Не определена')
            if subs not in by_subsystem:
                by_subsystem[subs] = {'total': 0, 'found': 0, 'names': []}
            by_subsystem[subs]['total'] += 1
            if result['found']:
                by_subsystem[subs]['found'] += 1
            else:
                by_subsystem[subs]['names'].append(result['product']['name'])

        # Группировка по типу совпадений
        match_types = {}
        for result in self.fuzzy_search_results:
            if result['found']:
                match_info = result.get('match_info')
                if match_info and len(match_info) > 3:
                    match_type = match_info[3]
                    match_types[match_type] = match_types.get(match_type, 0) + 1

        type_display = {
            'exact': 'Точные совпадения',
            'all_words': 'Все слова',
            'most_words': 'Большинство слов',
            'fuzzy_word': 'Нечеткие (по слову)',
            'fuzzy': 'Нечеткие',
            'fuzzy_multiple': 'Нечеткие (несколько слов)'
        }

        # Формируем отчет
        report = f"""
        <h1>Отчет по результатам нечеткого поиска СПО/БПО</h1>
        <p><b>Дата:</b> {datetime.now().strftime('%d.%m.%Y %H:%M')}</p>
        <p><b>Порог схожести:</b> 80% (фиксированный)</p>

        <h2>Общая статистика</h2>
        <ul>
            <li>Всего продуктов в БД: <b>{total}</b></li>
            <li><b style='color:green'>Найдено в документе: {found}</b></li>
            <li><b style='color:red'>Не найдено: {not_found}</b></li>
            <li>Процент покрытия: <b>{found / total * 100:.1f}%</b></li>
        </ul>

        <h2>Типы совпадений</h2>
        <ul>
        """

        for match_type, count in match_types.items():
            display_name = type_display.get(match_type, match_type)
            report += f"<li><b>{display_name}:</b> {count}</li>"

        report += "</ul>"

        report += """
        <h2>Статистика по подсистемам</h2>
        <table border='1' cellpadding='5' style='border-collapse: collapse;'>
            <tr>
                <th>Подсистема</th>
                <th>Всего</th>
                <th>Найдено</th>
                <th>Процент</th>
            </tr>
        """

        for subs, stats in sorted(by_subsystem.items()):
            percent = stats['found'] / stats['total'] * 100 if stats['total'] > 0 else 0
            report += f"""
            <tr>
                <td>{subs}</td>
                <td>{stats['total']}</td>
                <td>{stats['found']}</td>
                <td>{percent:.1f}%</td>
            </tr>
            """

        report += "</table>"

        # Детальный список не найденных продуктов по подсистемам
        report += "<h2>Не найденные продукты по подсистемам</h2>"

        for subs, stats in sorted(by_subsystem.items()):
            if stats['names']:
                report += f"<h3>{subs}</h3><ul>"
                for name in sorted(stats['names'])[:20]:  # Ограничиваем до 20 на подсистему
                    report += f"<li>{name}</li>"
                if len(stats['names']) > 20:
                    report += f"<li>... и еще {len(stats['names']) - 20} продуктов</li>"
                report += "</ul>"

        # Показываем отчет
        msg = QMessageBox(self)
        msg.setWindowTitle("Отчет по нечеткому поиску (порог 80%)")
        msg.setText(report)
        msg.setTextFormat(Qt.TextFormat.RichText)
        msg.setMinimumWidth(800)
        msg.setMinimumHeight(600)
        msg.exec()

    # ============= НОВЫЙ МЕТОД ДЛЯ ОБНОВЛЕНИЯ ВИДИМОСТИ ВКЛАДКИ =============
    def update_fuzzy_tab_visibility(self):
        """Обновление видимости вкладки нечеткого поиска на основе настроек"""
        settings = QSettings("ФедеральноеКазначейство", "Settings")
        show_fuzzy_tab = settings.value("show_fuzzy_tab", False, type=bool)

        # Находим текущий индекс вкладки нечеткого поиска
        current_index = -1
        for i in range(self.tab_widget.count()):
            if "СПО/БПО" in self.tab_widget.tabText(i):
                current_index = i
                break

        # Если вкладка должна быть показана, но её нет - добавляем
        if show_fuzzy_tab and current_index == -1:
            # Добавляем вкладку перед вкладкой базы данных
            self.tab_widget.insertTab(
                self.tab_widget.count() - 1,  # перед последней вкладкой
                self.fuzzy_tab,
                "🎯 СПО/БПО (нечеткий поиск 80%)"
            )
            print("Вкладка нечеткого поиска добавлена")  # Для отладки

        # Если вкладка не должна быть показана, но она есть - удаляем
        elif not show_fuzzy_tab and current_index != -1:
            self.tab_widget.removeTab(current_index)
            print("Вкладка нечеткого поиска удалена")  # Для отладки

    # ============= ОПЦИОНАЛЬНО: МЕТОД ДЛЯ ОБРАБОТКИ СОБЫТИЯ ПОКАЗА =============
    def showEvent(self, event):
        """Обработчик события показа диалога"""
        super().showEvent(event)
        # Обновляем видимость вкладки при каждом показе диалога
        self.update_fuzzy_tab_visibility()