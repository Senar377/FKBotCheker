from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QTableWidget, QTableWidgetItem, QHeaderView, QLineEdit,
    QComboBox, QMessageBox, QGroupBox, QTextEdit, QSplitter,
    QMenu, QApplication, QWidget, QFileDialog, QTabWidget, QProgressBar,
    QInputDialog, QDialogButtonBox, QFormLayout, QCheckBox,
    QGridLayout, QFrame, QToolTip
)
from PyQt6.QtCore import Qt, pyqtSignal, QTimer, QThread, QSettings, QObject
from PyQt6.QtGui import QFont, QColor, QAction, QTextCursor, QTextCharFormat
import re
import logging
import pandas as pd
import json
import os
import pickle
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime

logger = logging.getLogger(__name__)


class ExcelProductLoader:
    """Загрузчик продуктов из Excel с правильной структурой"""

    @staticmethod
    def load_products_from_excel(file_path):
        """
        Загружает продукты из Excel файла со структурой:
        - Модуль/компонент (при наличии)
        - Наименование ПО
        - Версия
        - Сертификат ФСТЭК
        - Предназначение/описание
        - ГК (при необходимости)
        """
        products = []

        try:
            excel_file = pd.ExcelFile(file_path)

            for sheet_name in excel_file.sheet_names:
                # Пропускаем листы с metadata
                if sheet_name.startswith('>'):
                    continue

                df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

                # Ищем заголовки
                header_row = None
                for i in range(min(10, len(df))):
                    row = df.iloc[i].astype(str).tolist()
                    row_text = ' '.join(row).lower()

                    # Проверяем наличие ключевых заголовков
                    if ('наименование по' in row_text or
                            'продукт' in row_text or
                            'версия' in row_text):
                        header_row = i
                        break

                if header_row is None:
                    continue

                # Устанавливаем заголовки
                df.columns = df.iloc[header_row].astype(str)
                df = df.iloc[header_row + 1:].reset_index(drop=True)

                # Определяем колонки
                name_col = None
                version_col = None
                desc_col = None
                gk_col = None
                module_col = None
                cert_col = None

                for col in df.columns:
                    col_lower = str(col).lower()

                    if 'наименование по' in col_lower:
                        name_col = col
                    elif 'версия' in col_lower:
                        version_col = col
                    elif 'предназначение' in col_lower or 'описание' in col_lower:
                        desc_col = col
                    elif 'гк' in col_lower or 'контракт' in col_lower:
                        gk_col = col
                    elif 'модуль' in col_lower or 'компонент' in col_lower:
                        module_col = col
                    elif 'сертификат' in col_lower:
                        cert_col = col

                # Обрабатываем строки
                for idx, row in df.iterrows():
                    # Пропускаем пустые строки
                    if all(pd.isna(val) for val in row):
                        continue

                    name = str(row[name_col]) if name_col and pd.notna(row[name_col]) else None
                    if not name or name.lower() in ['nan', 'none', '', 'бпо', 'спо']:
                        continue

                    version = str(row[version_col]) if version_col and pd.notna(row[version_col]) else ''
                    description = str(row[desc_col]) if desc_col and pd.notna(row[desc_col]) else ''
                    gk = str(row[gk_col]) if gk_col and pd.notna(row[gk_col]) else ''
                    module = str(row[module_col]) if module_col and pd.notna(row[module_col]) else ''
                    certificate = str(row[cert_col]) if cert_col and pd.notna(row[cert_col]) else ''

                    # Извлекаем ГК из строки
                    gk_list = ExcelProductLoader._extract_gk(gk)

                    # Дополнительно ищем ГК в других колонках, если не нашли
                    if not gk_list and gk_col:
                        # Проверяем другие колонки на наличие ГК
                        for col in df.columns:
                            if col != name_col and col != version_col:
                                val = str(row[col]) if pd.notna(row[col]) else ''
                                extracted = ExcelProductLoader._extract_gk(val)
                                gk_list.extend(extracted)

                    products.append({
                        'name': name.strip(),
                        'module': module.strip() if module != 'nan' else '',
                        'version': version.strip() if version != 'nan' else '',
                        'description': description.strip() if description != 'nan' else '',
                        'certificate': certificate.strip() if certificate != 'nan' else '',
                        'gk': list(set(gk_list)),  # Убираем дубликаты
                        'sheet': sheet_name,
                        'row': idx + header_row + 2,
                        'full_version': version.strip(),
                        'parsed_version': ExcelProductLoader._parse_version(version)
                    })

            # Убираем дубликаты по названию
            unique_products = {}
            for p in products:
                key = (p['name'].lower(), p['version'])
                if key not in unique_products:
                    unique_products[key] = p

            return list(unique_products.values())

        except Exception as e:
            logger.error(f"Ошибка загрузки Excel: {e}")
            return []

    @staticmethod
    def _extract_gk(text):
        """Извлекает ГК из текста"""
        if not text or pd.isna(text):
            return []

        text = str(text)
        # Паттерн для поиска ГК вида ФКУ000/2025
        pattern = r'[А-ЯA-Z]{3,4}\d{3,4}(?:[/-]\d{2,4})?'
        matches = re.findall(pattern, text, re.IGNORECASE)

        result = []
        for match in matches:
            clean = match.strip().replace(' ', '').upper()
            # Проверяем, что это действительно ГК (буквы + цифры)
            if re.match(r'^[А-ЯA-Z]{3,4}\d{3,4}(?:[/-]\d{2,4})?$', clean):
                result.append(clean)

        return list(set(result))

    @staticmethod
    def _parse_version(version_str):
        """Парсинг версии для сравнения"""
        if pd.isna(version_str) or not version_str:
            return [0]

        version_str = str(version_str).lower().strip()

        # Убираем текст "и выше", "и выше" и т.п.
        version_str = re.sub(r'и выше.*$', '', version_str)
        version_str = re.sub(r'и новее.*$', '', version_str)
        version_str = re.sub(r'\(.*\)', '', version_str)

        parts = []
        numbers = re.findall(r'\d+', version_str)
        for num in numbers:
            try:
                parts.append(int(num))
            except ValueError:
                pass

        # Обработка специальных обозначений
        if 'c' in version_str or 'g' in version_str:
            letter_match = re.search(r'(\d+)([cg])', version_str)
            if letter_match and parts:
                letter_value = 100 if letter_match.group(2) == 'c' else 50
                parts[-1] = parts[-1] * 1000 + letter_value

        if 'sp' in version_str or 'service pack' in version_str:
            sp_match = re.search(r'sp[\s]*(\d+)', version_str, re.IGNORECASE)
            if sp_match and parts:
                sp_num = int(sp_match.group(1))
                parts.append(sp_num * 100)

        return parts if parts else [0]


class ExcelLoaderWorker(QObject):
    """Рабочий поток для загрузки Excel"""

    finished = pyqtSignal(list)
    progress = pyqtSignal(int, str)
    error = pyqtSignal(str)

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path

    def run(self):
        try:
            self.progress.emit(10, "Чтение Excel файла...")
            products = ExcelProductLoader.load_products_from_excel(self.file_path)
            self.progress.emit(100, f"Загружено {len(products)} продуктов")
            self.finished.emit(products)
        except Exception as e:
            self.error.emit(str(e))


class ExcelProductSearchDialog(QDialog):
    """Диалог для поиска продуктов из Excel в документе"""

    def __init__(self, parent=None, document_text="", excel_file_path=None):
        super().__init__(parent)
        self.document_text = document_text
        self.excel_file_path = excel_file_path
        self.products = []  # Продукты из Excel
        self.search_results = []  # Результаты поиска в документе
        self.loader_thread = None
        self.worker = None

        self.setWindowTitle("Поиск продуктов из Excel в документе")
        self.resize(1400, 900)

        if parent and hasattr(parent, 'styleSheet'):
            self.setStyleSheet(parent.styleSheet())

        self.init_ui()
        self.load_excel_products()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        # Заголовок
        title = QLabel("🔍 Поиск продуктов из Excel в документе")
        title.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color: #3498db; padding: 10px;")
        layout.addWidget(title)

        # Информационная панель
        info_layout = QHBoxLayout()
        self.doc_info = QLabel(f"📄 Документ: {len(self.document_text)} символов")
        self.excel_info = QLabel("📊 Загрузка данных из Excel...")
        info_layout.addWidget(self.doc_info)
        info_layout.addWidget(self.excel_info)
        info_layout.addStretch()
        layout.addLayout(info_layout)

        # Прогресс-бар
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        # Панель фильтров
        filter_group = QGroupBox("Фильтры")
        filter_layout = QHBoxLayout()

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Поиск по названию продукта...")
        self.search_input.textChanged.connect(self.filter_results)

        self.gk_filter = QComboBox()
        self.gk_filter.addItem("Все ГК")
        self.gk_filter.currentTextChanged.connect(self.filter_results)

        self.status_filter = QComboBox()
        self.status_filter.addItems(["Все статусы", "✅ Найдено в документе", "❌ Не найдено", "⚠ Версия ниже требуемой"])
        self.status_filter.currentTextChanged.connect(self.filter_results)

        self.sheet_filter = QComboBox()
        self.sheet_filter.addItem("Все листы")
        self.sheet_filter.currentTextChanged.connect(self.filter_results)

        filter_layout.addWidget(QLabel("🔍 Поиск:"))
        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(QLabel("📋 ГК:"))
        filter_layout.addWidget(self.gk_filter)
        filter_layout.addWidget(QLabel("📊 Статус:"))
        filter_layout.addWidget(self.status_filter)
        filter_layout.addWidget(QLabel("📑 Лист:"))
        filter_layout.addWidget(self.sheet_filter)
        filter_layout.addStretch()

        filter_group.setLayout(filter_layout)
        layout.addWidget(filter_group)

        # Таблица результатов
        self.table = QTableWidget()
        self.table.setColumnCount(11)
        self.table.setHorizontalHeaderLabels([
            "Модуль/Компонент", "Наименование ПО", "Версия (Excel)",
            "Предназначение/Описание", "ГК", "Сертификат", "Лист Excel",
            "Найдено в документе", "Версия в документе", "Сравнение версий", "Действие"
        ])

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(7, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(8, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(9, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(10, QHeaderView.ResizeMode.Fixed)

        self.table.setColumnWidth(10, 100)
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table.setSortingEnabled(True)

        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        layout.addWidget(self.table)

        # Статистика
        stats_group = QGroupBox("Статистика")
        stats_layout = QHBoxLayout()

        self.total_products = QLabel("Всего продуктов: 0")
        self.found_products = QLabel("✅ Найдено в документе: 0")
        self.not_found_products = QLabel("❌ Не найдено: 0")
        self.compliant_versions = QLabel("👍 Версии соответствуют: 0")
        self.non_compliant_versions = QLabel("👎 Версии ниже требуемой: 0")

        stats_layout.addWidget(self.total_products)
        stats_layout.addWidget(self.found_products)
        stats_layout.addWidget(self.not_found_products)
        stats_layout.addWidget(self.compliant_versions)
        stats_layout.addWidget(self.non_compliant_versions)
        stats_layout.addStretch()

        stats_group.setLayout(stats_layout)
        layout.addWidget(stats_group)

        # Кнопки
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        export_btn = QPushButton("📤 Экспорт результатов")
        export_btn.clicked.connect(self.export_results)
        export_btn.setMinimumHeight(35)

        copy_all_btn = QPushButton("📋 Копировать все")
        copy_all_btn.clicked.connect(self.copy_all_to_clipboard)
        copy_all_btn.setMinimumHeight(35)

        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        close_btn.setMinimumHeight(35)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                font-weight: bold;
                border-radius: 5px;
                padding: 8px 15px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)

        btn_layout.addWidget(export_btn)
        btn_layout.addWidget(copy_all_btn)
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def load_excel_products(self):
        """Загрузка продуктов из Excel"""
        if not self.excel_file_path:
            QMessageBox.warning(self, "Ошибка", "Не указан файл Excel")
            return

        self.progress.setVisible(True)
        self.progress.setValue(10)
        self.excel_info.setText("⏳ Загрузка данных из Excel...")

        # Загружаем продукты в отдельном потоке
        self.loader_thread = QThread()
        self.worker = ExcelLoaderWorker(self.excel_file_path)
        self.worker.moveToThread(self.loader_thread)

        self.loader_thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.on_products_loaded)
        self.worker.progress.connect(self.update_progress)
        self.worker.error.connect(self.on_load_error)
        self.worker.finished.connect(self.loader_thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.loader_thread.finished.connect(self.loader_thread.deleteLater)

        self.loader_thread.start()

    def on_products_loaded(self, products):
        """Обработка загруженных продуктов"""
        self.products = products
        self.progress.setValue(100)
        self.progress.setVisible(False)

        self.excel_info.setText(f"📊 Загружено продуктов: {len(products)}")

        # Обновляем фильтр ГК
        all_gk = set()
        all_sheets = set()
        for p in products:
            for gk in p.get('gk', []):
                all_gk.add(gk)
            all_sheets.add(p.get('sheet', ''))

        self.gk_filter.clear()
        self.gk_filter.addItem("Все ГК")
        self.gk_filter.addItems(sorted(all_gk))

        self.sheet_filter.clear()
        self.sheet_filter.addItem("Все листы")
        self.sheet_filter.addItems(sorted(all_sheets))

        # Ищем продукты в документе
        self.search_products_in_document()

    def update_progress(self, value, message):
        """Обновление прогресса"""
        self.progress.setValue(value)
        self.excel_info.setText(message)

    def on_load_error(self, error_msg):
        """Обработка ошибки загрузки"""
        self.progress.setVisible(False)
        QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить Excel: {error_msg}")
        self.excel_info.setText("❌ Ошибка загрузки")

    def search_products_in_document(self):
        """Поиск продуктов из Excel в документе"""
        self.search_results = []

        lines = self.document_text.split('\n')
        full_text_lower = self.document_text.lower()

        for product in self.products:
            result = {
                'product': product,
                'found': False,
                'matches': [],
                'best_version': None,
                'best_version_parsed': None,
                'comparison': None,
                'is_compliant': None
            }

            # Ищем название продукта в документе
            name_lower = product['name'].lower()

            # Разбиваем название на ключевые слова (убираем короткие слова)
            words = re.findall(r'[a-zа-я]+', name_lower)
            keywords = [w for w in words if len(w) > 2]

            # Добавляем специальные ключевые слова для известных продуктов
            special_keywords = {
                'postgres': ['postgres', 'pg', 'postgresql'],
                'oracle': ['oracle'],
                'mysql': ['mysql'],
                'mssql': ['mssql', 'sql server'],
                'windows': ['windows', 'win'],
                'linux': ['linux'],
                'red os': ['ред ос', 'red os', 'murom'],
                'astra': ['astra', 'астра'],
                'java': ['java', 'jdk', 'jre'],
                'kafka': ['kafka'],
                'nginx': ['nginx'],
                'docker': ['docker'],
                'kubernetes': ['k8s', 'kubernetes'],
                'haproxy': ['haproxy', 'ha proxy'],
                'openshift': ['openshift'],
                'deckhouse': ['deckhouse'],
                'consul': ['consul'],
                'zookeeper': ['zookeeper', 'zk'],
                'elasticsearch': ['elastic', 'elasticsearch'],
                'redis': ['redis'],
                'mongodb': ['mongo', 'mongodb'],
                'cassandra': ['cassandra'],
                'clickhouse': ['clickhouse'],
                'hadoop': ['hadoop', 'hdfs'],
                'spark': ['spark', 'apache spark'],
                'ufos': ['уфос', 'ufos'],
                'svbo': ['svbo'],
                'svip': ['svip'],
                '1c': ['1с', '1c', '1с:предприятие'],
                'мойофис': ['мойофис', 'myoffice'],
                'р7': ['р7', 'r7']
            }

            # Добавляем специальные ключевые слова
            for key, values in special_keywords.items():
                if key in name_lower:
                    keywords.extend(values)

            # Убираем дубликаты
            keywords = list(set(keywords))

            # Поиск по строкам
            for line_num, line in enumerate(lines):
                line_lower = line.lower()

                # Проверяем, содержит ли строка ключевые слова
                match_score = 0
                matched_keywords = []

                for kw in keywords:
                    if kw in line_lower:
                        match_score += 10
                        matched_keywords.append(kw)

                # Если нашли совпадение, ищем версию
                if match_score > 0:
                    # Ищем версию в этой строке
                    version_match = self._extract_version_from_line(line)

                    if version_match:
                        result['found'] = True
                        result['matches'].append({
                            'line': line_num + 1,
                            'text': line.strip(),
                            'version': version_match,
                            'score': match_score,
                            'matched_keywords': matched_keywords
                        })

            # Выбираем лучшее совпадение
            if result['matches']:
                result['matches'].sort(key=lambda x: (x['score'], len(x['version']) if x['version'] else 0),
                                       reverse=True)
                best_match = result['matches'][0]
                result['best_version'] = best_match['version']
                result['best_version_parsed'] = self._parse_version(best_match['version'])

                # Сравниваем версии
                if product.get('parsed_version') and product['parsed_version'] != [0]:
                    comparison = self._compare_versions(
                        result['best_version_parsed'],
                        product['parsed_version']
                    )
                    result['comparison'] = comparison
                    result['is_compliant'] = comparison >= 0

            self.search_results.append(result)

        self.display_results()
        self.update_statistics()

    def _extract_version_from_line(self, line):
        """Извлечение версии из строки"""
        # Паттерны для поиска версий
        patterns = [
            r'версия\s*[:\s]*([\d\.]+(?:[a-z]?\d*)?)',
            r'v\.?\s*([\d\.]+(?:[a-z]?\d*)?)',
            r'ver\.?\s*([\d\.]+(?:[a-z]?\d*)?)',
            r'(\d+\.\d+(?:\.\d+)?(?:[\.-][a-z]?\d+)?)',
            r'(\d+[\.]\d+[\.]\d+(?:[\.]\d+)?)',
            r'(\d+\.\d+)',
            r'(\d+)(?=\s*(?:\.|,|\s|$))'
        ]

        for pattern in patterns:
            match = re.search(pattern, line, re.IGNORECASE)
            if match:
                version = match.group(1).strip()
                # Проверяем, что это похоже на версию (содержит цифры)
                if re.search(r'\d', version):
                    return version

        return None

    def _parse_version(self, version_str):
        """Парсинг версии для сравнения"""
        if not version_str:
            return [0]

        version_str = str(version_str).lower().strip()

        # Убираем текст в скобках
        version_str = re.sub(r'\([^)]*\)', '', version_str)

        parts = []
        numbers = re.findall(r'\d+', version_str)
        for num in numbers:
            try:
                parts.append(int(num))
            except ValueError:
                pass

        return parts if parts else [0]

    def _compare_versions(self, v1_parts, v2_parts):
        """Сравнение версий: возвращает 1 если v1 > v2, 0 если равны, -1 если v1 < v2"""
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

    def display_results(self):
        """Отображение результатов"""
        self.table.setRowCount(len(self.search_results))

        for i, result in enumerate(self.search_results):
            product = result['product']

            # Модуль/компонент
            module_item = QTableWidgetItem(product.get('module', ''))
            module_item.setData(Qt.ItemDataRole.UserRole, result)
            self.table.setItem(i, 0, module_item)

            # Наименование ПО
            name_item = QTableWidgetItem(product['name'])
            name_item.setData(Qt.ItemDataRole.UserRole + 1, product)
            self.table.setItem(i, 1, name_item)

            # Версия (Excel)
            version_item = QTableWidgetItem(product.get('version', ''))
            self.table.setItem(i, 2, version_item)

            # Предназначение/Описание
            desc = product.get('description', '')
            if len(desc) > 100:
                desc = desc[:100] + "..."
            desc_item = QTableWidgetItem(desc)
            desc_item.setToolTip(product.get('description', ''))
            self.table.setItem(i, 3, desc_item)

            # ГК
            gk_text = ', '.join(product.get('gk', []))
            self.table.setItem(i, 4, QTableWidgetItem(gk_text))

            # Сертификат
            cert_item = QTableWidgetItem(product.get('certificate', ''))
            self.table.setItem(i, 5, cert_item)

            # Лист Excel
            sheet_item = QTableWidgetItem(product.get('sheet', ''))
            self.table.setItem(i, 6, sheet_item)

            # Найдено в документе
            if result['found']:
                found_text = f"✅ Найдено ({len(result['matches'])})"
                found_color = QColor(0, 150, 0)
            else:
                found_text = "❌ Не найдено"
                found_color = QColor(255, 0, 0)

            found_item = QTableWidgetItem(found_text)
            found_item.setForeground(found_color)
            self.table.setItem(i, 7, found_item)

            # Версия в документе
            doc_version = result['best_version'] if result['best_version'] else ''
            doc_version_item = QTableWidgetItem(doc_version)
            if doc_version:
                doc_version_item.setForeground(QColor(0, 0, 150))
            self.table.setItem(i, 8, doc_version_item)

            # Сравнение версий
            if result['found'] and result['best_version']:
                if result['is_compliant'] is not None:
                    if result['is_compliant']:
                        comp_text = f"✅ {doc_version} >= {product.get('version', '')}"
                        comp_color = QColor(0, 150, 0)
                    else:
                        comp_text = f"❌ {doc_version} < {product.get('version', '')}"
                        comp_color = QColor(255, 0, 0)
                else:
                    comp_text = "⚠ Не удалось сравнить"
                    comp_color = QColor(255, 140, 0)
            else:
                comp_text = "—"
                comp_color = QColor(128, 128, 128)

            comp_item = QTableWidgetItem(comp_text)
            comp_item.setForeground(comp_color)
            self.table.setItem(i, 9, comp_item)

            # Кнопка деталей
            details_btn = QPushButton("🔍 Детали")
            details_btn.setProperty('row', i)
            details_btn.clicked.connect(lambda checked, r=i: self.show_product_details(r))
            details_btn.setMaximumWidth(80)
            self.table.setCellWidget(i, 10, details_btn)

        self.table.resizeRowsToContents()

    def show_product_details(self, row):
        """Показ деталей продукта"""
        result = self.search_results[row]
        product = result['product']

        details = f"<h3>Детали продукта</h3>"
        details += f"<p><b>Наименование:</b> {product['name']}</p>"

        if product.get('module'):
            details += f"<p><b>Модуль/компонент:</b> {product['module']}</p>"

        details += f"<p><b>Требуемая версия:</b> {product.get('version', '')}</p>"

        if product.get('description'):
            details += f"<p><b>Описание:</b> {product['description']}</p>"

        if product.get('certificate'):
            details += f"<p><b>Сертификат ФСТЭК:</b> {product['certificate']}</p>"

        if product.get('gk'):
            details += f"<p><b>ГК:</b> {', '.join(product['gk'])}</p>"

        details += f"<p><b>Лист Excel:</b> {product.get('sheet', '')}</p>"

        if result['found']:
            details += f"<h4>Найдено в документе:</h4>"
            details += f"<p><b>Версия в документе:</b> {result['best_version']}</p>"

            if result['is_compliant'] is not None:
                if result['is_compliant']:
                    details += f"<p><b style='color:green'>✅ Версия соответствует требованиям (>= {product.get('version', '')})</b></p>"
                else:
                    details += f"<p><b style='color:red'>❌ Версия ниже требуемой ({result['best_version']} < {product.get('version', '')})</b></p>"
            else:
                details += f"<p><b style='color:orange'>⚠ Не удалось сравнить версии</b></p>"

            details += "<h4>Все совпадения:</h4><ul>"
            for match in result['matches'][:5]:
                details += f"<li>Строка {match['line']}: {match['text'][:100]}..."
                if match.get('version'):
                    details += f" <b>(версия: {match['version']})</b>"
                if match.get('matched_keywords'):
                    details += f" <i>[ключевые слова: {', '.join(match['matched_keywords'][:3])}]</i>"
                details += "</li>"
            if len(result['matches']) > 5:
                details += f"<li>... и еще {len(result['matches']) - 5} совпадений</li>"
            details += "</ul>"
        else:
            details += "<p><b style='color:red'>❌ Продукт не найден в документе</b></p>"
            details += "<p>Возможные причины:</p><ul>"
            details += "<li>Продукт упоминается под другим названием</li>"
            details += "<li>Продукт отсутствует в документе</li>"
            details += "<li>Требуется более точный поиск</li>"
            details += "</ul>"

        msg = QMessageBox(self)
        msg.setWindowTitle(f"Детали: {product['name']}")
        msg.setText(details)
        msg.setTextFormat(Qt.TextFormat.RichText)
        msg.setMinimumWidth(600)
        msg.exec()

    def filter_results(self):
        """Фильтрация результатов"""
        search_text = self.search_input.text().lower()
        gk_filter = self.gk_filter.currentText()
        status_filter = self.status_filter.currentText()
        sheet_filter = self.sheet_filter.currentText()

        for row in range(self.table.rowCount()):
            show_row = True

            # Поиск по тексту
            if search_text:
                row_text = ""
                for col in [0, 1, 2, 3]:
                    item = self.table.item(row, col)
                    if item:
                        row_text += item.text().lower() + " "

                if search_text not in row_text:
                    show_row = False

            # Фильтр по ГК
            if show_row and gk_filter != "Все ГК":
                gk_item = self.table.item(row, 4)
                if gk_item and gk_filter not in gk_item.text():
                    show_row = False

            # Фильтр по листу
            if show_row and sheet_filter != "Все листы":
                sheet_item = self.table.item(row, 6)
                if sheet_item and sheet_item.text() != sheet_filter:
                    show_row = False

            # Фильтр по статусу
            if show_row and status_filter != "Все статусы":
                status_item = self.table.item(row, 7)
                comp_item = self.table.item(row, 9)
                if status_item and comp_item:
                    if status_filter == "✅ Найдено в документе" and "✅" not in status_item.text():
                        show_row = False
                    elif status_filter == "❌ Не найдено" and "✅" in status_item.text():
                        show_row = False
                    elif status_filter == "⚠ Версия ниже требуемой" and "❌" not in comp_item.text():
                        show_row = False

            self.table.setRowHidden(row, not show_row)

        self.update_statistics()

    def update_statistics(self):
        """Обновление статистики"""
        total = len(self.search_results)

        found = 0
        not_found = 0
        compliant = 0
        non_compliant = 0

        for row in range(self.table.rowCount()):
            if not self.table.isRowHidden(row):
                status_item = self.table.item(row, 7)
                comp_item = self.table.item(row, 9)
                if status_item and comp_item:
                    if "✅" in status_item.text():
                        found += 1
                        if "✅" in comp_item.text():
                            compliant += 1
                        elif "❌" in comp_item.text():
                            non_compliant += 1
                    else:
                        not_found += 1

        self.total_products.setText(f"Всего продуктов: {total}")
        self.found_products.setText(f"✅ Найдено в документе: {found}")
        self.not_found_products.setText(f"❌ Не найдено: {not_found}")
        self.compliant_versions.setText(f"👍 Версии соответствуют: {compliant}")
        self.non_compliant_versions.setText(f"👎 Версии ниже требуемой: {non_compliant}")

    def show_context_menu(self, position):
        """Контекстное меню для таблицы"""
        menu = QMenu()

        copy_row_action = QAction("📋 Копировать строку", self)
        copy_row_action.triggered.connect(self.copy_selected_row)

        copy_name_action = QAction("📋 Копировать название", self)
        copy_name_action.triggered.connect(self.copy_selected_name)

        copy_version_action = QAction("📋 Копировать версию", self)
        copy_version_action.triggered.connect(self.copy_selected_version)

        show_details_action = QAction("🔍 Показать детали", self)
        show_details_action.triggered.connect(self.show_details_for_selected)

        menu.addAction(copy_row_action)
        menu.addAction(copy_name_action)
        menu.addAction(copy_version_action)
        menu.addSeparator()
        menu.addAction(show_details_action)

        menu.exec(self.table.viewport().mapToGlobal(position))

    def copy_selected_row(self):
        """Копирование выбранной строки"""
        current_row = self.table.currentRow()
        if current_row >= 0:
            text_parts = []
            for col in range(self.table.columnCount() - 1):
                item = self.table.item(current_row, col)
                if item:
                    text_parts.append(item.text())
                else:
                    text_parts.append("")

            QApplication.clipboard().setText('\t'.join(text_parts))
            QToolTip.showText(self.mapToGlobal(self.rect().center()), "Строка скопирована")

    def copy_selected_name(self):
        """Копирование названия продукта"""
        current_row = self.table.currentRow()
        if current_row >= 0:
            name_item = self.table.item(current_row, 1)
            if name_item:
                QApplication.clipboard().setText(name_item.text())
                QToolTip.showText(self.mapToGlobal(self.rect().center()), "Название скопировано")

    def copy_selected_version(self):
        """Копирование версии"""
        current_row = self.table.currentRow()
        if current_row >= 0:
            version_item = self.table.item(current_row, 2)
            if version_item:
                QApplication.clipboard().setText(version_item.text())
                QToolTip.showText(self.mapToGlobal(self.rect().center()), "Версия скопирована")

    def show_details_for_selected(self):
        """Показ деталей для выбранной строки"""
        current_row = self.table.currentRow()
        if current_row >= 0:
            self.show_product_details(current_row)

    def copy_all_to_clipboard(self):
        """Копирование всех видимых строк в буфер обмена"""
        text = "Модуль/Компонент\tНаименование ПО\tВерсия (Excel)\tПредназначение/Описание\tГК\tСертификат\tЛист Excel\tНайдено в документе\tВерсия в документе\tСравнение версий\n"

        for row in range(self.table.rowCount()):
            if not self.table.isRowHidden(row):
                row_text = []
                for col in range(self.table.columnCount() - 1):
                    item = self.table.item(row, col)
                    if item:
                        row_text.append(item.text())
                    else:
                        row_text.append("")
                text += '\t'.join(row_text) + '\n'

        QApplication.clipboard().setText(text)

        visible_count = sum(1 for row in range(self.table.rowCount())
                            if not self.table.isRowHidden(row))
        QMessageBox.information(self, "Копирование", f"Скопировано {visible_count} строк")

    def export_results(self):
        """Экспорт результатов в CSV"""
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить результаты",
            f"excel_products_search_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "CSV files (*.csv)"
        )

        if file_path:
            try:
                import csv
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow([
                        "Модуль/Компонент", "Наименование ПО", "Версия (Excel)",
                        "Предназначение/Описание", "ГК", "Сертификат", "Лист Excel",
                        "Найдено в документе", "Версия в документе", "Сравнение версий"
                    ])

                    for row in range(self.table.rowCount()):
                        if not self.table.isRowHidden(row):
                            row_data = []
                            for col in range(self.table.columnCount() - 1):
                                item = self.table.item(row, col)
                                if item:
                                    row_data.append(item.text())
                                else:
                                    row_data.append("")
                            writer.writerow(row_data)

                QMessageBox.information(self, "Успех", f"Результаты экспортированы в {file_path}")
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось экспортировать: {str(e)}")


class GKVersionComparisonDialog(QDialog):
    def __init__(self, parent=None, document_text="", excel_data=None, document_versions=None):
        super().__init__(parent)
        self.document_text = document_text
        self.excel_data = excel_data or {}
        self.document_versions = document_versions or []
        self.found_gk = []
        self.comparison_results = []
        self.setWindowTitle("Сравнение версий по ГК")
        self.resize(1200, 800)

        if parent and hasattr(parent, 'styleSheet'):
            self.setStyleSheet(parent.styleSheet())

        self.init_ui()
        self.search_gk_in_document()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        title = QLabel("🔄 Сравнение версий по ГК")
        title.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color: #3498db; padding: 10px;")
        layout.addWidget(title)

        info_layout = QHBoxLayout()
        self.doc_info_label = QLabel(f"📄 Документ: {len(self.document_text)} символов")
        self.excel_info_label = QLabel(f"📊 Данные Excel: {len(self.excel_data)} записей")
        self.versions_info_label = QLabel(f"🔍 Версий в документе: {len(self.document_versions)}")
        info_layout.addWidget(self.doc_info_label)
        info_layout.addWidget(self.excel_info_label)
        info_layout.addWidget(self.versions_info_label)
        info_layout.addStretch()
        layout.addLayout(info_layout)

        control_panel = QGroupBox("Параметры сравнения")
        control_layout = QHBoxLayout()

        self.compare_all_btn = QPushButton("🔄 Сравнить все ГК")
        self.compare_all_btn.clicked.connect(self.compare_all_gk)
        self.compare_all_btn.setMinimumHeight(35)
        self.compare_all_btn.setStyleSheet("""
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

        self.gk_filter_combo = QComboBox()
        self.gk_filter_combo.addItem("Все ГК")
        self.gk_filter_combo.currentTextChanged.connect(self.filter_by_gk)

        self.status_filter_combo = QComboBox()
        self.status_filter_combo.addItems(
            ["Все статусы", "✅ Соответствует", "❌ Не соответствует", "⚠ Требует проверки", "❓ ГК не найден"])
        self.status_filter_combo.currentTextChanged.connect(self.filter_results)

        control_layout.addWidget(self.compare_all_btn)
        control_layout.addWidget(QLabel("Фильтр по ГК:"))
        control_layout.addWidget(self.gk_filter_combo)
        control_layout.addWidget(QLabel("Статус:"))
        control_layout.addWidget(self.status_filter_combo)
        control_layout.addStretch()

        control_panel.setLayout(control_layout)
        layout.addWidget(control_panel)

        splitter = QSplitter(Qt.Orientation.Vertical)

        gk_widget = QWidget()
        gk_layout = QVBoxLayout(gk_widget)
        gk_layout.setContentsMargins(0, 0, 0, 0)

        gk_label = QLabel("📋 Найденные ГК в документе:")
        gk_label.setStyleSheet("font-weight: bold; font-size: 12px;")
        gk_layout.addWidget(gk_label)

        self.gk_table = QTableWidget()
        self.gk_table.setColumnCount(5)
        self.gk_table.setHorizontalHeaderLabels(["№", "ГК", "Статус в Excel", "Версий в Excel", "Версий в документе"])
        self.gk_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.gk_table.setAlternatingRowColors(True)
        self.gk_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.gk_table.itemSelectionChanged.connect(self.on_gk_selected)

        gk_layout.addWidget(self.gk_table)

        results_widget = QWidget()
        results_layout = QVBoxLayout(results_widget)
        results_layout.setContentsMargins(0, 0, 0, 0)

        results_label = QLabel("📊 Результаты сравнения версий:")
        results_label.setStyleSheet("font-weight: bold; font-size: 12px;")
        results_layout.addWidget(results_label)

        filter_panel = QHBoxLayout()
        filter_panel.addWidget(QLabel("Поиск:"))
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Поиск по названию ПО или версии...")
        self.search_input.textChanged.connect(self.filter_results)
        filter_panel.addWidget(self.search_input)

        self.category_filter = QComboBox()
        self.category_filter.addItem("Все категории")
        self.category_filter.currentTextChanged.connect(self.filter_results)
        filter_panel.addWidget(QLabel("Категория:"))
        filter_panel.addWidget(self.category_filter)

        filter_panel.addStretch()
        results_layout.addLayout(filter_panel)

        self.results_table = QTableWidget()
        self.results_table.setColumnCount(9)
        self.results_table.setHorizontalHeaderLabels([
            "ГК", "Наименование ПО", "Версия в документе",
            "Версия в Excel", "Категория", "Сравнение", "Статус", "Лист Excel", "Действие"
        ])

        header = self.results_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(7, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(8, QHeaderView.ResizeMode.Fixed)

        self.results_table.setColumnWidth(8, 100)
        self.results_table.setAlternatingRowColors(True)
        self.results_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        self.results_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.results_table.customContextMenuRequested.connect(self.show_results_context_menu)

        results_layout.addWidget(self.results_table)

        splitter.addWidget(gk_widget)
        splitter.addWidget(results_widget)
        splitter.setSizes([200, 400])

        layout.addWidget(splitter)

        stats_group = QGroupBox("Статистика сравнения")
        stats_layout = QHBoxLayout()

        self.total_gk_label = QLabel("Всего ГК: 0")
        self.total_versions_label = QLabel("Всего версий: 0")
        self.compliant_label = QLabel("✅ Соответствует: 0")
        self.not_compliant_label = QLabel("❌ Не соответствует: 0")
        self.needs_check_label = QLabel("⚠ Требует проверки: 0")
        self.gk_not_found_label = QLabel("❓ ГК не найден: 0")

        stats_layout.addWidget(self.total_gk_label)
        stats_layout.addWidget(self.total_versions_label)
        stats_layout.addWidget(self.compliant_label)
        stats_layout.addWidget(self.not_compliant_label)
        stats_layout.addWidget(self.needs_check_label)
        stats_layout.addWidget(self.gk_not_found_label)
        stats_layout.addStretch()

        stats_group.setLayout(stats_layout)
        layout.addWidget(stats_group)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        export_btn = QPushButton("📤 Экспорт результатов")
        export_btn.clicked.connect(self.export_results)
        export_btn.setMinimumHeight(35)

        copy_all_btn = QPushButton("📋 Копировать все")
        copy_all_btn.clicked.connect(self.copy_all_to_clipboard)
        copy_all_btn.setMinimumHeight(35)

        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        close_btn.setMinimumHeight(35)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                font-weight: bold;
                border-radius: 5px;
                padding: 8px 15px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)

        btn_layout.addWidget(export_btn)
        btn_layout.addWidget(copy_all_btn)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def extract_gk_from_text(self, text):
        if not text:
            return []

        pattern = r'[А-ЯA-Z]{2,5}[-\s]?\d{2,5}(?:[/-]\d{2,4})?(?:\s*/\s*\d+)?'
        matches = re.findall(pattern, text, re.IGNORECASE)

        result = []
        for match in matches:
            clean_match = match.strip().replace(' ', '')
            clean_match = re.sub(r'[^\w/-]', '', clean_match)
            if len(clean_match) >= 5:
                result.append(clean_match.upper())

        return list(set(result))

    def normalize_product_name(self, name):
        if not name:
            return ""

        name = str(name).lower().strip()
        name = re.sub(r'[^\w\s-]', ' ', name)
        name = re.sub(r'\s+', ' ', name)

        replacements = {
            'postgres pro': 'postgrespro',
            'postgresql': 'postgres',
            'postgres pro enterprise': 'postgresproenterprise',
            'postgres pro certified': 'postgresprocertified',
            'microsoft sql server': 'mssql',
            'sql server': 'mssql',
            'windows server': 'windowsserver',
            'red hat': 'redhat',
            'oracle database': 'oracle',
            'mysql': 'mysql',
            'mariadb': 'mariadb',
            'mongodb': 'mongodb',
            'redis': 'redis',
            'apache': 'apache',
            'nginx': 'nginx',
            'docker': 'docker',
            'kubernetes': 'kubernetes',
            'vmware': 'vmware',
            'virtualbox': 'virtualbox',
            'python': 'python',
            'java': 'java',
            'node.js': 'nodejs',
            'php': 'php',
            '1с:предприятие': '1c',
            '1с:enterprise': '1c',
            'мойофис': 'myoffice',
            'р7-офис': 'r7office',
            'onlyoffice': 'onlyoffice',
            'astra linux': 'astra',
            'ред ос': 'redos',
            'alt linux': 'altlinux'
        }

        for old, new in replacements.items():
            name = name.replace(old, new)

        common_words = ['software', 'программа', 'система', 'подсистема',
                        'пакет', 'комплекс', 'application', 'server', 'версия',
                        'version', 'platform', 'платформа', 'среда', 'редакция']
        for word in common_words:
            name = name.replace(word, '')

        return name.strip()

    def calculate_name_match_improved(self, doc_name, excel_name):
        if not doc_name or not excel_name:
            return 0

        doc_norm = self.normalize_product_name(doc_name)
        excel_norm = self.normalize_product_name(excel_name)

        if doc_norm == excel_norm:
            return 100

        if doc_norm in excel_norm or excel_norm in doc_norm:
            return 95

        doc_words = set(doc_norm.split())
        excel_words = set(excel_norm.split())

        if not doc_words or not excel_words:
            return 0

        common = doc_words.intersection(excel_words)

        score = 0
        for word in common:
            if len(word) > 2:
                if word in ['postgres', 'oracle', 'mysql', 'windows', 'linux', 'mssql', '1c']:
                    score += 30
                else:
                    score += 15

        for doc_word in doc_words:
            for excel_word in excel_words:
                if len(doc_word) > 3 and len(excel_word) > 3:
                    if doc_word in excel_word or excel_word in doc_word:
                        score += 10

        return min(score, 100)

    def find_matching_excel_entries(self, gk, doc_product):
        matches = []

        for norm_name, data in self.excel_data.items():
            excel_gk_list = data.get('gk', [])
            if gk in excel_gk_list:
                match_score = self.calculate_name_match_improved(doc_product, data['name'])
                matches.append({
                    'data': data,
                    'score': match_score + 50,
                    'match_type': 'gk'
                })

        if not matches:
            for norm_name, data in self.excel_data.items():
                match_score = self.calculate_name_match_improved(doc_product, data['name'])
                if match_score > 60:
                    matches.append({
                        'data': data,
                        'score': match_score,
                        'match_type': 'name'
                    })

        matches.sort(key=lambda x: x['score'], reverse=True)
        return matches

    def search_gk_in_document(self):
        self.found_gk = self.extract_gk_from_text(self.document_text)

        self.gk_table.setRowCount(len(self.found_gk))

        gk_list = ["Все ГК"]
        for i, gk in enumerate(self.found_gk):
            self.gk_table.setItem(i, 0, QTableWidgetItem(str(i + 1)))

            gk_item = QTableWidgetItem(gk)
            gk_item.setData(Qt.ItemDataRole.UserRole, gk)
            self.gk_table.setItem(i, 1, gk_item)
            gk_list.append(gk)

            excel_data = self.get_excel_data_for_gk(gk)
            if excel_data:
                status_item = QTableWidgetItem("✅ Есть в Excel")
                status_item.setForeground(QColor(0, 150, 0))
                self.gk_table.setItem(i, 2, status_item)

                excel_versions = self.count_excel_versions_for_gk(gk)
                self.gk_table.setItem(i, 3, QTableWidgetItem(str(excel_versions)))
            else:
                status_item = QTableWidgetItem("❌ Нет в Excel")
                status_item.setForeground(QColor(255, 0, 0))
                self.gk_table.setItem(i, 2, status_item)
                self.gk_table.setItem(i, 3, QTableWidgetItem("0"))

            doc_versions = self.count_document_versions_for_gk(gk)
            self.gk_table.setItem(i, 4, QTableWidgetItem(str(doc_versions)))

        self.gk_table.resizeRowsToContents()

        self.gk_filter_combo.clear()
        self.gk_filter_combo.addItems(gk_list)

        self.compare_all_gk()

    def get_excel_data_for_gk(self, gk):
        result = []
        for norm_name, data in self.excel_data.items():
            if gk in data.get('gk', []):
                result.append(data)
        return result

    def count_excel_versions_for_gk(self, gk):
        count = 0
        for norm_name, data in self.excel_data.items():
            if gk in data.get('gk', []):
                if data.get('parsed_versions'):
                    count += len(data['parsed_versions'])
                else:
                    count += len(data.get('versions', []))
        return count

    def count_document_versions_for_gk(self, gk):
        excel_data = self.get_excel_data_for_gk(gk)
        product_names = [d['name'].lower() for d in excel_data]

        count = 0
        for version in self.document_versions:
            product_lower = version.get('product', '').lower()
            for excel_name in product_names:
                if excel_name in product_lower or product_lower in excel_name:
                    count += 1
                    break

        return count

    def compare_all_gk(self):
        self.comparison_results = []
        categories = set()

        for gk in self.found_gk:
            gk_matches = []

            for doc_version in self.document_versions:
                doc_product = doc_version.get('product', '')
                doc_version_str = doc_version.get('version', '')
                doc_category = doc_version.get('category', 'Прочее')

                excel_matches = self.find_matching_excel_entries(gk, doc_product)

                if excel_matches:
                    best_match = excel_matches[0]['data']
                    match_score = excel_matches[0]['score']

                    excel_versions = []
                    if best_match.get('parsed_versions'):
                        excel_versions = [v['original'] for v in best_match['parsed_versions']]
                    else:
                        excel_versions = best_match.get('versions', [])

                    excel_versions.sort(key=lambda x: self.parse_version(x), reverse=True)

                    comparison_result = self.compare_doc_with_excel_versions(
                        doc_version_str, excel_versions
                    )

                    gk_matches.append({
                        'gk': gk,
                        'doc_name': doc_product,
                        'doc_version': doc_version_str,
                        'excel_name': best_match['name'],
                        'excel_version': comparison_result['excel_version'],
                        'category': doc_category,
                        'comparison': comparison_result['comparison_text'],
                        'status': comparison_result['status'],
                        'status_color': comparison_result['color'],
                        'excel_sheets': ', '.join(best_match.get('sheets', [])),
                        'match_score': match_score,
                        'all_excel_versions': excel_versions,
                        'match_type': excel_matches[0]['match_type']
                    })
                    categories.add(doc_category)
                else:
                    gk_matches.append({
                        'gk': gk,
                        'doc_name': doc_product,
                        'doc_version': doc_version_str,
                        'excel_name': 'НЕ НАЙДЕНО',
                        'excel_version': '',
                        'category': doc_category,
                        'comparison': '—',
                        'status': '❌ Нет соответствия в Excel',
                        'status_color': QColor(255, 0, 0),
                        'excel_sheets': '',
                        'match_score': 0,
                        'all_excel_versions': [],
                        'match_type': 'none'
                    })
                    categories.add(doc_category)

            self.comparison_results.extend(gk_matches)

        self.display_results(categories)

    def compare_doc_with_excel_versions(self, doc_version, excel_versions):
        if not excel_versions:
            return {
                'excel_version': '',
                'comparison_text': '—',
                'status': '⚠ Нет версий в Excel',
                'color': QColor(255, 140, 0)
            }

        doc_parts = self.parse_version(doc_version)

        parsed_excel_versions = []
        for v in excel_versions:
            parsed = self.parse_version(v)
            parsed_excel_versions.append((v, parsed))

        parsed_excel_versions.sort(key=lambda x: x[1], reverse=True)

        best_match = None
        best_comparison = None
        best_status = None
        best_color = None

        for excel_version, excel_parts in parsed_excel_versions:
            comparison = self.compare_version_parts(doc_parts, excel_parts)

            if comparison == 1:
                status = f"✅ {doc_version} > {excel_version}"
                color = QColor(0, 150, 0)
            elif comparison == 0:
                status = f"✅ {doc_version} = {excel_version}"
                color = QColor(0, 150, 0)
            elif comparison == -1:
                status = f"❌ {doc_version} < {excel_version}"
                color = QColor(255, 0, 0)
            else:
                continue

            if best_match is None:
                best_match = excel_version
                best_comparison = comparison
                best_status = status
                best_color = color
            elif comparison == 1 and best_comparison != 1:
                best_match = excel_version
                best_comparison = comparison
                best_status = status
                best_color = color
            elif comparison == 0 and best_comparison == -1:
                best_match = excel_version
                best_comparison = comparison
                best_status = status
                best_color = color

        if best_match:
            return {
                'excel_version': best_match,
                'comparison_text': best_status.split(' ')[1] if ' ' in best_status else best_status,
                'status': best_status,
                'color': best_color
            }
        else:
            return {
                'excel_version': excel_versions[0],
                'comparison_text': '—',
                'status': '⚠ Требует проверки',
                'color': QColor(255, 140, 0)
            }

    def parse_version(self, version_str):
        if not version_str:
            return [0]

        version_str = str(version_str).lower().strip()

        parts = []
        numbers = re.findall(r'\d+', version_str)
        for num in numbers:
            try:
                parts.append(int(num))
            except ValueError:
                pass

        if 'c' in version_str or 'g' in version_str:
            letter_match = re.search(r'(\d+)([cg])', version_str)
            if letter_match and parts:
                letter_value = 100 if letter_match.group(2) == 'c' else 50
                parts[-1] = parts[-1] * 1000 + letter_value

        if 'sp' in version_str or 'service pack' in version_str:
            sp_match = re.search(r'sp[\s]*(\d+)', version_str, re.IGNORECASE)
            if sp_match and parts:
                sp_num = int(sp_match.group(1))
                parts.append(sp_num * 100)

        if 'r' in version_str and not 'red' in version_str:
            r_match = re.search(r'r[\s]*(\d+)', version_str, re.IGNORECASE)
            if r_match and parts:
                r_num = int(r_match.group(1))
                parts.append(r_num)

        if 'update' in version_str:
            u_match = re.search(r'update[\s]*(\d+)', version_str, re.IGNORECASE)
            if u_match and parts:
                u_num = int(u_match.group(1))
                parts.append(u_num)

        if 'build' in version_str:
            b_match = re.search(r'build[\s]*(\d+)', version_str, re.IGNORECASE)
            if b_match and parts:
                b_num = int(b_match.group(1))
                parts.append(b_num)

        return parts if parts else [0]

    def compare_version_parts(self, parts1, parts2):
        if not parts1 or not parts2:
            return 0

        max_len = max(len(parts1), len(parts2))
        p1 = parts1 + [0] * (max_len - len(parts1))
        p2 = parts2 + [0] * (max_len - len(parts2))

        for a, b in zip(p1, p2):
            if a > b:
                return 1
            elif a < b:
                return -1

        return 0

    def display_results(self, categories):
        self.results_table.setRowCount(len(self.comparison_results))

        for i, result in enumerate(self.comparison_results):
            self.results_table.setItem(i, 0, QTableWidgetItem(result['gk']))

            name_item = QTableWidgetItem(result['doc_name'])
            name_item.setData(Qt.ItemDataRole.UserRole, result)
            self.results_table.setItem(i, 1, name_item)

            self.results_table.setItem(i, 2, QTableWidgetItem(result['doc_version']))

            excel_version_item = QTableWidgetItem(result['excel_version'])
            if result['excel_name'] == 'НЕ НАЙДЕНО':
                excel_version_item.setForeground(QColor(128, 128, 128))
            self.results_table.setItem(i, 3, excel_version_item)

            self.results_table.setItem(i, 4, QTableWidgetItem(result['category']))

            comparison_item = QTableWidgetItem(result['comparison'])
            self.results_table.setItem(i, 5, comparison_item)

            status_text = result['status']
            if result.get('match_type') == 'name' and '✅' in status_text:
                status_text = "🔍 " + status_text + " (по названию)"
            elif result.get('match_type') == 'gk' and '✅' in status_text:
                status_text = "🔑 " + status_text + " (по ГК)"

            status_item = QTableWidgetItem(status_text)
            status_item.setForeground(result['status_color'])
            self.results_table.setItem(i, 6, status_item)

            sheets_item = QTableWidgetItem(result.get('excel_sheets', ''))
            self.results_table.setItem(i, 7, sheets_item)

            details_btn = QPushButton("🔍 Детали")
            details_btn.setMaximumWidth(80)
            details_btn.setProperty('row', i)
            details_btn.clicked.connect(lambda checked, r=i: self.show_version_details(r))
            self.results_table.setCellWidget(i, 8, details_btn)

        self.results_table.resizeRowsToContents()
        self.update_category_filter(categories)
        self.update_statistics()

    def update_category_filter(self, categories):
        current = self.category_filter.currentText()
        self.category_filter.clear()
        self.category_filter.addItem("Все категории")
        self.category_filter.addItems(sorted(categories))

        index = self.category_filter.findText(current)
        if index >= 0:
            self.category_filter.setCurrentIndex(index)

    def filter_by_gk(self, gk_text):
        if gk_text == "Все ГК":
            for row in range(self.results_table.rowCount()):
                self.results_table.setRowHidden(row, False)
        else:
            for row in range(self.results_table.rowCount()):
                gk_item = self.results_table.item(row, 0)
                if gk_item:
                    self.results_table.setRowHidden(row, gk_item.text() != gk_text)
        self.filter_results()

    def filter_results(self):
        search_text = self.search_input.text().lower()
        category = self.category_filter.currentText()
        status_filter = self.status_filter_combo.currentText()
        gk_filter = self.gk_filter_combo.currentText()

        for row in range(self.results_table.rowCount()):
            if self.results_table.isRowHidden(row) and gk_filter != "Все ГК":
                continue

            show_row = True

            if search_text:
                row_text = ""
                for col in [0, 1, 2, 3]:
                    item = self.results_table.item(row, col)
                    if item:
                        row_text += item.text().lower() + " "

                if search_text not in row_text:
                    show_row = False

            if category != "Все категории":
                cat_item = self.results_table.item(row, 4)
                if cat_item and cat_item.text() != category:
                    show_row = False

            if status_filter != "Все статусы":
                status_item = self.results_table.item(row, 6)
                if status_item:
                    status = status_item.text()
                    if status_filter == "✅ Соответствует" and "✅" not in status:
                        show_row = False
                    elif status_filter == "❌ Не соответствует" and "❌" not in status:
                        show_row = False
                    elif status_filter == "⚠ Требует проверки" and "⚠" not in status:
                        show_row = False
                    elif status_filter == "❓ ГК не найден" and "❓" not in status:
                        show_row = False

            self.results_table.setRowHidden(row, not show_row)

        self.update_statistics()

    def on_gk_selected(self):
        selected_rows = self.gk_table.selectionModel().selectedRows()
        if not selected_rows:
            return

        row = selected_rows[0].row()
        gk_item = self.gk_table.item(row, 1)
        if not gk_item:
            return

        selected_gk = gk_item.text()

        index = self.gk_filter_combo.findText(selected_gk)
        if index >= 0:
            self.gk_filter_combo.setCurrentIndex(index)

    def show_version_details(self, row):
        result = self.comparison_results[row]

        details = f"<h3>Детали сравнения</h3>\n\n"
        details += f"<p><b>ГК:</b> {result['gk']}</p>\n"
        details += f"<p><b>Продукт в документе:</b> {result['doc_name']}</p>\n"
        details += f"<p><b>Версия в документе:</b> {result['doc_version']}</p>\n"
        details += f"<p><b>Продукт в Excel:</b> {result['excel_name']}</p>\n"
        details += f"<p><b>Версия в Excel:</b> {result['excel_version']}</p>\n"
        details += f"<p><b>Категория:</b> {result['category']}</p>\n"
        details += f"<p><b>Статус:</b> {result['status']}</p>\n"
        details += f"<p><b>Тип совпадения:</b> {result.get('match_type', 'неизвестно')}</p>\n"
        details += f"<p><b>Листы Excel:</b> {result.get('excel_sheets', '')}</p>\n"

        if result.get('all_excel_versions'):
            details += "<p><b>Все версии в Excel:</b></p><ul>"
            for v in result['all_excel_versions'][:5]:
                details += f"<li>{v}</li>"
            if len(result['all_excel_versions']) > 5:
                details += f"<li>... и еще {len(result['all_excel_versions']) - 5}</li>"
            details += "</ul>"

        msg = QMessageBox(self)
        msg.setWindowTitle("Детали версии")
        msg.setText(details)
        msg.setTextFormat(Qt.TextFormat.RichText)
        msg.exec()

    def update_statistics(self):
        total_gk = len(self.found_gk)

        visible_rows = 0
        for row in range(self.results_table.rowCount()):
            if not self.results_table.isRowHidden(row):
                visible_rows += 1

        compliant = 0
        not_compliant = 0
        needs_check = 0
        gk_not_found = 0

        for row in range(self.results_table.rowCount()):
            if not self.results_table.isRowHidden(row):
                status_item = self.results_table.item(row, 6)
                if status_item:
                    status = status_item.text()
                    if "✅" in status:
                        compliant += 1
                    elif "❌" in status and "❓" not in status:
                        not_compliant += 1
                    elif "⚠" in status:
                        needs_check += 1
                    elif "❓" in status:
                        gk_not_found += 1

        self.total_gk_label.setText(f"Всего ГК: {total_gk}")
        self.total_versions_label.setText(f"Всего версий: {visible_rows}")
        self.compliant_label.setText(f"✅ Соответствует: {compliant}")
        self.not_compliant_label.setText(f"❌ Не соответствует: {not_compliant}")
        self.needs_check_label.setText(f"⚠ Требует проверки: {needs_check}")
        self.gk_not_found_label.setText(f"❓ ГК не найден: {gk_not_found}")

    def show_results_context_menu(self, position):
        menu = QMenu()

        copy_row_action = QAction("📋 Копировать строку", self)
        copy_row_action.triggered.connect(self.copy_selected_row)

        copy_gk_action = QAction("📋 Копировать ГК", self)
        copy_gk_action.triggered.connect(self.copy_selected_gk)

        show_details_action = QAction("🔍 Показать детали", self)
        show_details_action.triggered.connect(self.show_details_for_selected)

        menu.addAction(copy_row_action)
        menu.addAction(copy_gk_action)
        menu.addSeparator()
        menu.addAction(show_details_action)

        menu.exec(self.results_table.viewport().mapToGlobal(position))

    def copy_selected_row(self):
        current_row = self.results_table.currentRow()
        if current_row >= 0:
            text_parts = []
            for col in range(self.results_table.columnCount() - 1):
                item = self.results_table.item(current_row, col)
                if item:
                    text_parts.append(item.text())
                else:
                    text_parts.append("")

            QApplication.clipboard().setText('\t'.join(text_parts))

    def copy_selected_gk(self):
        current_row = self.results_table.currentRow()
        if current_row >= 0:
            gk_item = self.results_table.item(current_row, 0)
            if gk_item:
                QApplication.clipboard().setText(gk_item.text())

    def show_details_for_selected(self):
        current_row = self.results_table.currentRow()
        if current_row >= 0:
            self.show_version_details(current_row)

    def copy_all_to_clipboard(self):
        text = "ГК\tНаименование ПО\tВерсия в документе\tВерсия в Excel\tКатегория\tСравнение\tСтатус\tЛист Excel\n"

        for row in range(self.results_table.rowCount()):
            if not self.results_table.isRowHidden(row):
                row_text = []
                for col in range(self.results_table.columnCount() - 1):
                    item = self.results_table.item(row, col)
                    if item:
                        row_text.append(item.text())
                    else:
                        row_text.append("")
                text += '\t'.join(row_text) + '\n'

        QApplication.clipboard().setText(text)

        visible_count = sum(1 for row in range(self.results_table.rowCount())
                            if not self.results_table.isRowHidden(row))
        QMessageBox.information(self, "Копирование", f"Скопировано {visible_count} строк")

    def export_results(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить результаты", f"gk_comparison_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "CSV files (*.csv)"
        )

        if file_path:
            try:
                import csv
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow([
                        "ГК", "Наименование ПО", "Версия в документе",
                        "Версия в Excel", "Категория", "Сравнение", "Статус", "Лист Excel"
                    ])

                    for row in range(self.results_table.rowCount()):
                        if not self.results_table.isRowHidden(row):
                            row_data = []
                            for col in range(self.results_table.columnCount() - 1):
                                item = self.results_table.item(row, col)
                                if item:
                                    row_data.append(item.text())
                                else:
                                    row_data.append("")
                            writer.writerow(row_data)

                QMessageBox.information(self, "Успех", f"Результаты экспортированы в {file_path}")
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось экспортировать: {str(e)}")


class GKSearchDialog(QDialog):
    def __init__(self, parent=None, document_text="", excel_data=None):
        super().__init__(parent)
        self.document_text = document_text
        self.excel_data = excel_data or {}
        self.found_gk = []
        self.matching_versions = []
        self.setWindowTitle("Поиск ГК и версий ПО")
        self.resize(1000, 700)

        if parent and hasattr(parent, 'styleSheet'):
            self.setStyleSheet(parent.styleSheet())

        self.init_ui()
        self.search_gk_in_document()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        title = QLabel("🔍 Поиск ГК и соответствующих версий ПО")
        title.setFont(QFont("Arial", 16, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color: #3498db; padding: 10px;")
        layout.addWidget(title)

        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        line.setStyleSheet("background-color: #3498db; max-height: 2px;")
        layout.addWidget(line)

        info_layout = QHBoxLayout()
        self.doc_info_label = QLabel(f"📄 Документ: {len(self.document_text)} символов")
        self.excel_info_label = QLabel(f"📊 Данные Excel: {len(self.excel_data)} записей")
        info_layout.addWidget(self.doc_info_label)
        info_layout.addWidget(self.excel_info_label)
        info_layout.addStretch()
        layout.addLayout(info_layout)

        gk_group = QGroupBox("📋 Найденные ГК в документе")
        gk_group.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        gk_layout = QVBoxLayout()

        self.gk_list = QTableWidget()
        self.gk_list.setColumnCount(4)
        self.gk_list.setHorizontalHeaderLabels(["№", "ГК", "Статус в Excel", "Кол-во версий"])
        self.gk_list.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.gk_list.setAlternatingRowColors(True)
        self.gk_list.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        self.gk_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.gk_list.customContextMenuRequested.connect(self.show_gk_context_menu)

        gk_layout.addWidget(self.gk_list)
        gk_group.setLayout(gk_layout)
        layout.addWidget(gk_group)

        compare_btn = QPushButton("🔄 Перейти к сравнению версий по ГК")
        compare_btn.clicked.connect(self.go_to_comparison)
        compare_btn.setMinimumHeight(35)
        compare_btn.setStyleSheet("""
            QPushButton {
                background-color: #f39c12;
                color: white;
                font-weight: bold;
                border-radius: 5px;
                padding: 8px 15px;
            }
            QPushButton:hover {
                background-color: #e67e22;
            }
        """)
        layout.addWidget(compare_btn)

        versions_group = QGroupBox("📦 Версии ПО по найденным ГК")
        versions_group.setFont(QFont("Arial", 10, QFont.Weight.Bold))
        versions_layout = QVBoxLayout()

        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("🔍 Фильтр:"))
        self.version_search = QLineEdit()
        self.version_search.setPlaceholderText("Поиск по названию ПО или версии...")
        self.version_search.textChanged.connect(self.filter_versions_table)
        filter_layout.addWidget(self.version_search)

        self.category_filter = QComboBox()
        self.category_filter.addItem("Все категории")
        self.category_filter.currentTextChanged.connect(self.filter_versions_table)
        filter_layout.addWidget(QLabel("Категория:"))
        filter_layout.addWidget(self.category_filter)

        self.status_filter = QComboBox()
        self.status_filter.addItems(["Все статусы", "✅ Найден в Excel", "❌ Не найден в Excel"])
        self.status_filter.currentTextChanged.connect(self.filter_versions_table)
        filter_layout.addWidget(QLabel("Статус:"))
        filter_layout.addWidget(self.status_filter)

        filter_layout.addStretch()
        versions_layout.addLayout(filter_layout)

        self.versions_table = QTableWidget()
        self.versions_table.setColumnCount(7)
        self.versions_table.setHorizontalHeaderLabels([
            "ГК", "Наименование ПО", "Версия", "Категория",
            "Лист Excel", "Статус", "Действие"
        ])

        header = self.versions_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.Fixed)

        self.versions_table.setColumnWidth(6, 100)
        self.versions_table.setAlternatingRowColors(True)
        self.versions_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        self.versions_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.versions_table.customContextMenuRequested.connect(self.show_versions_context_menu)

        versions_layout.addWidget(self.versions_table)
        versions_group.setLayout(versions_layout)
        layout.addWidget(versions_group)

        stats_group = QGroupBox("📊 Статистика")
        stats_layout = QHBoxLayout()

        self.gk_count_label = QLabel("Найдено ГК: 0")
        self.gk_found_label = QLabel("Найдено в Excel: 0")
        self.gk_not_found_label = QLabel("Не найдено в Excel: 0")
        self.versions_count_label = QLabel("Всего версий: 0")

        stats_layout.addWidget(self.gk_count_label)
        stats_layout.addWidget(self.gk_found_label)
        stats_layout.addWidget(self.gk_not_found_label)
        stats_layout.addWidget(self.versions_count_label)
        stats_layout.addStretch()

        stats_group.setLayout(stats_layout)
        layout.addWidget(stats_group)

        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        export_btn = QPushButton("📤 Экспорт результатов")
        export_btn.clicked.connect(self.export_results)
        export_btn.setMinimumHeight(35)
        export_btn.setStyleSheet("""
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

        copy_all_btn = QPushButton("📋 Копировать все")
        copy_all_btn.clicked.connect(self.copy_all_to_clipboard)
        copy_all_btn.setMinimumHeight(35)

        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        close_btn.setMinimumHeight(35)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
                font-weight: bold;
                border-radius: 5px;
                padding: 8px 15px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)

        btn_layout.addWidget(copy_all_btn)
        btn_layout.addWidget(export_btn)
        btn_layout.addWidget(close_btn)

        layout.addLayout(btn_layout)

        self.setLayout(layout)

    def go_to_comparison(self):
        document_versions = []
        if hasattr(self.parent(), 'versions_data'):
            document_versions = self.parent().versions_data

        dialog = GKVersionComparisonDialog(
            self.parent(),
            self.document_text,
            self.excel_data,
            document_versions
        )
        dialog.exec()

    def extract_gk_from_text(self, text):
        if not text:
            return []

        pattern = r'[А-ЯA-Z]{2,5}[-\s]?\d{2,5}(?:[/-]\d{2,4})?(?:\s*/\s*\d+)?'
        matches = re.findall(pattern, text, re.IGNORECASE)

        result = []
        for match in matches:
            clean_match = match.strip().replace(' ', '')
            clean_match = re.sub(r'[^\w/-]', '', clean_match)
            if len(clean_match) >= 5:
                result.append(clean_match.upper())

        return list(set(result))

    def search_gk_in_document(self):
        self.found_gk = self.extract_gk_from_text(self.document_text)

        self.gk_list.setRowCount(len(self.found_gk))

        for i, gk in enumerate(self.found_gk):
            self.gk_list.setItem(i, 0, QTableWidgetItem(str(i + 1)))

            gk_item = QTableWidgetItem(gk)
            gk_item.setData(Qt.ItemDataRole.UserRole, gk)
            self.gk_list.setItem(i, 1, gk_item)

            found_in_excel, versions_count = self.find_gk_in_excel_with_count(gk)

            status_item = QTableWidgetItem()
            if found_in_excel:
                status_item.setText("✅ Найден в Excel")
                status_item.setForeground(QColor(0, 150, 0))
                status_item.setData(Qt.ItemDataRole.UserRole, "found")
            else:
                status_item.setText("❌ Не найден в Excel")
                status_item.setForeground(QColor(255, 0, 0))
                status_item.setData(Qt.ItemDataRole.UserRole, "not_found")

            self.gk_list.setItem(i, 2, status_item)

            self.gk_list.setItem(i, 3, QTableWidgetItem(str(versions_count)))

        self.gk_list.resizeRowsToContents()

        self.collect_versions_by_gk()

        self.update_statistics()

        self.update_category_filter()

    def find_gk_in_excel_with_count(self, gk):
        if not self.excel_data:
            return False, 0

        versions_count = 0
        found = False

        for norm_name, data in self.excel_data.items():
            if gk in data.get('gk', []):
                found = True
                if data.get('parsed_versions'):
                    versions_count += len(data['parsed_versions'])
                else:
                    versions_count += len(data.get('versions', []))

        return found, versions_count

    def get_versions_for_gk(self, gk):
        versions = []

        if not self.excel_data:
            return versions

        for norm_name, data in self.excel_data.items():
            if gk in data.get('gk', []):
                if data.get('parsed_versions'):
                    for ver_info in data['parsed_versions']:
                        versions.append({
                            'gk': gk,
                            'name': data['name'],
                            'version': ver_info['original'],
                            'category': self.determine_category(data['name']),
                            'sheets': ', '.join(data.get('sheets', [])),
                            'status': '✅ Найден в Excel',
                            'status_color': QColor(0, 150, 0)
                        })
                else:
                    for version in data.get('versions', []):
                        versions.append({
                            'gk': gk,
                            'name': data['name'],
                            'version': version,
                            'category': self.determine_category(data['name']),
                            'sheets': ', '.join(data.get('sheets', [])),
                            'status': '✅ Найден в Excel',
                            'status_color': QColor(0, 150, 0)
                        })

        return versions

    def determine_category(self, name):
        name_lower = name.lower()

        categories = {
            'Операционные системы': ['windows', 'linux', 'ubuntu', 'debian', 'centos', 'astra', 'ред ос', 'alt', 'rosa',
                                     'macos'],
            'Базы данных': ['oracle', 'mysql', 'postgresql', 'mariadb', 'mongodb', 'redis', 'sql server', 'database'],
            'Виртуализация': ['vmware', 'virtualbox', 'docker', 'kubernetes', 'hyper-v', 'kvm', 'proxmox'],
            'Веб-серверы': ['apache', 'nginx', 'iis', 'tomcat', 'jetty'],
            'Языки программирования': ['python', 'java', 'php', 'node.js', 'ruby', 'go', '.net', 'javascript'],
            'Российское ПО': ['мойофис', 'р7', 'onlyoffice', '1с', 'vk workspace', 'астра', 'ред ос'],
            'Сетевое оборудование': ['cisco', 'juniper', 'mikrotik', 'd-link', 'huawei']
        }

        for category, keywords in categories.items():
            for keyword in keywords:
                if keyword in name_lower:
                    return category

        return 'Прочее ПО'

    def collect_versions_by_gk(self):
        self.matching_versions = []
        categories = set()

        if not self.found_gk:
            self.show_all_versions_with_note("В документе не найдено ГК")
            return

        found_any = False
        for gk in self.found_gk:
            versions = self.get_versions_for_gk(gk)
            if versions:
                self.matching_versions.extend(versions)
                found_any = True
                for v in versions:
                    categories.add(v['category'])

        if not found_any:
            self.show_all_versions_with_note("ГК из документа не найдены в Excel")
        else:
            self.display_versions(categories)

    def show_all_versions_with_note(self, note):
        self.matching_versions = []
        categories = set()

        for norm_name, data in self.excel_data.items():
            status = f"❓ {note}"

            if data.get('parsed_versions'):
                for ver_info in data['parsed_versions']:
                    category = self.determine_category(data['name'])
                    self.matching_versions.append({
                        'gk': '—',
                        'name': data['name'],
                        'version': ver_info['original'],
                        'category': category,
                        'sheets': ', '.join(data.get('sheets', [])),
                        'status': status,
                        'status_color': QColor(255, 140, 0)
                    })
                    categories.add(category)
            else:
                for version in data.get('versions', []):
                    category = self.determine_category(data['name'])
                    self.matching_versions.append({
                        'gk': '—',
                        'name': data['name'],
                        'version': version,
                        'category': category,
                        'sheets': ', '.join(data.get('sheets', [])),
                        'status': status,
                        'status_color': QColor(255, 140, 0)
                    })
                    categories.add(category)

        self.display_versions(categories)

    def display_versions(self, categories=None):
        self.versions_table.setRowCount(len(self.matching_versions))

        if categories is None:
            categories = set()

        for i, ver in enumerate(self.matching_versions):
            gk_item = QTableWidgetItem(ver['gk'])
            if ver['gk'] == '—':
                gk_item.setForeground(QColor(128, 128, 128))
            self.versions_table.setItem(i, 0, gk_item)

            name_item = QTableWidgetItem(ver['name'])
            name_item.setData(Qt.ItemDataRole.UserRole, ver)
            self.versions_table.setItem(i, 1, name_item)

            self.versions_table.setItem(i, 2, QTableWidgetItem(ver['version']))

            category_item = QTableWidgetItem(ver['category'])
            categories.add(ver['category'])
            self.versions_table.setItem(i, 3, category_item)

            self.versions_table.setItem(i, 4, QTableWidgetItem(ver['sheets']))

            status_item = QTableWidgetItem(ver['status'])
            status_item.setForeground(ver['status_color'])
            self.versions_table.setItem(i, 5, status_item)

            copy_btn = QPushButton("📋 Копировать")
            copy_btn.setMaximumWidth(80)
            copy_btn.setProperty('row', i)
            copy_btn.clicked.connect(lambda checked, r=i: self.copy_version_row(r))
            self.versions_table.setCellWidget(i, 6, copy_btn)

        self.versions_table.resizeRowsToContents()

        self.update_category_filter(categories)

    def update_category_filter(self, categories=None):
        current_text = self.category_filter.currentText()
        self.category_filter.clear()
        self.category_filter.addItem("Все категории")

        if categories:
            self.category_filter.addItems(sorted(categories))

        index = self.category_filter.findText(current_text)
        if index >= 0:
            self.category_filter.setCurrentIndex(index)

    def filter_versions_table(self):
        search_text = self.version_search.text().lower()
        category = self.category_filter.currentText()
        status = self.status_filter.currentText()

        for row in range(self.versions_table.rowCount()):
            show_row = True

            if search_text:
                row_text = ""
                for col in [0, 1, 2]:
                    item = self.versions_table.item(row, col)
                    if item:
                        row_text += item.text().lower() + " "

                if search_text not in row_text:
                    show_row = False

            if category != "Все категории":
                cat_item = self.versions_table.item(row, 3)
                if cat_item and cat_item.text() != category:
                    show_row = False

            if status != "Все статусы":
                status_item = self.versions_table.item(row, 5)
                if status_item:
                    if status == "✅ Найден в Excel" and "✅" not in status_item.text():
                        show_row = False
                    elif status == "❌ Не найден в Excel" and "❌" not in status_item.text():
                        show_row = False

            self.versions_table.setRowHidden(row, not show_row)

    def update_statistics(self):
        total_gk = len(self.found_gk)
        gk_found = 0
        gk_not_found = 0

        for row in range(self.gk_list.rowCount()):
            status_item = self.gk_list.item(row, 2)
            if status_item:
                if "✅" in status_item.text():
                    gk_found += 1
                else:
                    gk_not_found += 1

        self.gk_count_label.setText(f"Найдено ГК: {total_gk}")
        self.gk_found_label.setText(f"Найдено в Excel: {gk_found}")
        self.gk_not_found_label.setText(f"Не найдено в Excel: {gk_not_found}")
        self.versions_count_label.setText(f"Всего версий: {len(self.matching_versions)}")

    def show_gk_context_menu(self, position):
        item = self.gk_list.itemAt(position)
        if not item:
            return

        row = self.gk_list.rowAt(position.y())
        if row < 0:
            return

        gk_item = self.gk_list.item(row, 1)
        if not gk_item:
            return

        gk = gk_item.text()

        menu = QMenu()

        copy_action = QAction("📋 Копировать ГК", self)
        copy_action.triggered.connect(lambda: QApplication.clipboard().setText(gk))

        filter_action = QAction("🔍 Показать только этот ГК", self)
        filter_action.triggered.connect(lambda: self.filter_by_gk(gk))

        menu.addAction(copy_action)
        menu.addAction(filter_action)

        menu.exec(self.gk_list.viewport().mapToGlobal(position))

    def filter_by_gk(self, gk):
        self.version_search.clear()
        self.category_filter.setCurrentIndex(0)
        self.status_filter.setCurrentIndex(0)

        for row in range(self.versions_table.rowCount()):
            gk_item = self.versions_table.item(row, 0)
            if gk_item:
                show_row = (gk_item.text() == gk)
                self.versions_table.setRowHidden(row, not show_row)

    def show_versions_context_menu(self, position):
        item = self.versions_table.itemAt(position)
        if not item:
            return

        row = self.versions_table.rowAt(position.y())
        if row < 0:
            return

        gk_item = self.versions_table.item(row, 0)
        name_item = self.versions_table.item(row, 1)
        version_item = self.versions_table.item(row, 2)

        if not name_item:
            return

        menu = QMenu()

        copy_name_action = QAction("📋 Копировать название", self)
        copy_name_action.triggered.connect(lambda: QApplication.clipboard().setText(name_item.text()))

        copy_version_action = QAction("📋 Копировать версию", self)
        copy_version_action.triggered.connect(lambda: QApplication.clipboard().setText(version_item.text()))

        copy_gk_action = QAction("📋 Копировать ГК", self)
        copy_gk_action.triggered.connect(lambda: QApplication.clipboard().setText(gk_item.text() if gk_item else ""))

        copy_row_action = QAction("📋 Копировать строку", self)
        copy_row_action.triggered.connect(lambda: self.copy_version_row(row))

        menu.addAction(copy_name_action)
        menu.addAction(copy_version_action)
        menu.addAction(copy_gk_action)
        menu.addSeparator()
        menu.addAction(copy_row_action)

        menu.exec(self.versions_table.viewport().mapToGlobal(position))

    def copy_version_row(self, row):
        text_parts = []
        for col in range(self.versions_table.columnCount() - 1):
            item = self.versions_table.item(row, col)
            if item:
                text_parts.append(item.text())
            else:
                text_parts.append("")

        QApplication.clipboard().setText('\t'.join(text_parts))

        QToolTip.showText(self.mapToGlobal(self.rect().center()), "Строка скопирована")

    def copy_all_to_clipboard(self):
        text = "ГК\tНаименование ПО\tВерсия\tКатегория\tЛист Excel\tСтатус\n"

        for row in range(self.versions_table.rowCount()):
            if not self.versions_table.isRowHidden(row):
                row_text = []
                for col in range(self.versions_table.columnCount() - 1):
                    item = self.versions_table.item(row, col)
                    if item:
                        row_text.append(item.text())
                    else:
                        row_text.append("")
                text += '\t'.join(row_text) + '\n'

        QApplication.clipboard().setText(text)

        visible_count = sum(1 for row in range(self.versions_table.rowCount())
                            if not self.versions_table.isRowHidden(row))
        QMessageBox.information(self, "Копирование", f"Скопировано {visible_count} строк")

    def export_results(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить результаты", f"gk_search_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "CSV files (*.csv)"
        )

        if file_path:
            try:
                import csv
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(["ГК", "Наименование ПО", "Версия", "Категория", "Лист Excel", "Статус"])

                    for row in range(self.versions_table.rowCount()):
                        if not self.versions_table.isRowHidden(row):
                            row_data = []
                            for col in range(self.versions_table.columnCount() - 1):
                                item = self.versions_table.item(row, col)
                                if item:
                                    row_data.append(item.text())
                                else:
                                    row_data.append("")
                            writer.writerow(row_data)

                QMessageBox.information(self, "Успех", f"Результаты экспортированы в {file_path}")
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось экспортировать: {str(e)}")


class ExcelDataManager:
    def __init__(self):
        self.settings = QSettings("ФедеральноеКазначейство", "ExcelDataManager")
        self.data_file = "excel_versions_data.pkl"
        self.metadata_file = "excel_metadata.json"

    def save_data(self, data, file_path):
        try:
            with open(self.data_file, 'wb') as f:
                pickle.dump(data, f)

            metadata = {
                'file_path': file_path,
                'file_name': os.path.basename(file_path),
                'save_date': datetime.now().isoformat(),
                'record_count': len(data)
            }

            with open(self.metadata_file, 'w', encoding='utf-8') as f:
                json.dump(metadata, f, ensure_ascii=False, indent=2)

            self.settings.setValue("last_excel_file", file_path)
            self.settings.setValue("last_excel_save_date", datetime.now().isoformat())

            return True
        except Exception as e:
            logger.error(f"Ошибка сохранения данных: {e}")
            return False

    def load_data(self):
        try:
            if os.path.exists(self.data_file):
                with open(self.data_file, 'rb') as f:
                    data = pickle.load(f)

                metadata = {}
                if os.path.exists(self.metadata_file):
                    with open(self.metadata_file, 'r', encoding='utf-8') as f:
                        metadata = json.load(f)

                return data, metadata
            return None, None
        except Exception as e:
            logger.error(f"Ошибка загрузки данных: {e}")
            return None, None

    def clear_data(self):
        try:
            if os.path.exists(self.data_file):
                os.remove(self.data_file)
            if os.path.exists(self.metadata_file):
                os.remove(self.metadata_file)
            self.settings.remove("last_excel_file")
            self.settings.remove("last_excel_save_date")
            return True
        except Exception as e:
            logger.error(f"Ошибка очистки данных: {e}")
            return False

    def get_last_file(self):
        return self.settings.value("last_excel_file", None)


class EditExcelDataDialog(QDialog):
    def __init__(self, parent=None, data=None, norm_name=None):
        super().__init__(parent)
        self.data = data.copy() if data else {}
        self.norm_name = norm_name
        self.setWindowTitle(f"Редактирование: {data.get('name', '')}")
        self.resize(600, 500)

        if parent and hasattr(parent, 'styleSheet'):
            self.setStyleSheet(parent.styleSheet())

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        form_layout = QFormLayout()
        form_layout.setSpacing(10)

        self.name_edit = QLineEdit(self.data.get('name', ''))
        form_layout.addRow("Наименование:", self.name_edit)

        versions_layout = QVBoxLayout()
        self.versions_list = QTableWidget()
        self.versions_list.setColumnCount(2)
        self.versions_list.setHorizontalHeaderLabels(["Версия", "Действие"])

        versions = self.data.get('versions', [])

        self.versions_list.setRowCount(len(versions))
        for i, version in enumerate(versions):
            version_item = QTableWidgetItem(version)
            self.versions_list.setItem(i, 0, version_item)

            btn_remove = QPushButton("❌")
            btn_remove.setMaximumWidth(30)
            btn_remove.clicked.connect(lambda checked, r=i: self.remove_version(r))
            self.versions_list.setCellWidget(i, 1, btn_remove)

        self.versions_list.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.versions_list.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        self.versions_list.setColumnWidth(1, 40)

        versions_layout.addWidget(self.versions_list)

        add_version_layout = QHBoxLayout()
        self.new_version_input = QLineEdit()
        self.new_version_input.setPlaceholderText("Новая версия...")
        self.add_version_btn = QPushButton("➕ Добавить версию")
        self.add_version_btn.clicked.connect(self.add_version)

        add_version_layout.addWidget(self.new_version_input)
        add_version_layout.addWidget(self.add_version_btn)

        versions_layout.addLayout(add_version_layout)

        form_layout.addRow("Версии:", versions_layout)

        gk_layout = QVBoxLayout()
        self.gk_list = QTableWidget()
        self.gk_list.setColumnCount(2)
        self.gk_list.setHorizontalHeaderLabels(["ГК", "Действие"])

        gk_items = self.data.get('gk', [])
        self.gk_list.setRowCount(len(gk_items))
        for i, gk in enumerate(gk_items):
            gk_item = QTableWidgetItem(gk)
            self.gk_list.setItem(i, 0, gk_item)

            btn_remove = QPushButton("❌")
            btn_remove.setMaximumWidth(30)
            btn_remove.clicked.connect(lambda checked, r=i: self.remove_gk(r))
            self.gk_list.setCellWidget(i, 1, btn_remove)

        self.gk_list.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.gk_list.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        self.gk_list.setColumnWidth(1, 40)

        gk_layout.addWidget(self.gk_list)

        add_gk_layout = QHBoxLayout()
        self.new_gk_input = QLineEdit()
        self.new_gk_input.setPlaceholderText("Новый ГК...")
        self.add_gk_btn = QPushButton("➕ Добавить ГК")
        self.add_gk_btn.clicked.connect(self.add_gk)

        add_gk_layout.addWidget(self.new_gk_input)
        add_gk_layout.addWidget(self.add_gk_btn)

        gk_layout.addLayout(add_gk_layout)

        form_layout.addRow("ГК:", gk_layout)

        sheets_layout = QVBoxLayout()
        self.sheets_list = QTableWidget()
        self.sheets_list.setColumnCount(2)
        self.sheets_list.setHorizontalHeaderLabels(["Лист", "Действие"])

        sheets = self.data.get('sheets', [])
        self.sheets_list.setRowCount(len(sheets))
        for i, sheet in enumerate(sheets):
            sheet_item = QTableWidgetItem(sheet)
            self.sheets_list.setItem(i, 0, sheet_item)

            btn_remove = QPushButton("❌")
            btn_remove.setMaximumWidth(30)
            btn_remove.clicked.connect(lambda checked, r=i: self.remove_sheet(r))
            self.sheets_list.setCellWidget(i, 1, btn_remove)

        self.sheets_list.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.sheets_list.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        self.sheets_list.setColumnWidth(1, 40)

        sheets_layout.addWidget(self.sheets_list)

        add_sheet_layout = QHBoxLayout()
        self.new_sheet_input = QLineEdit()
        self.new_sheet_input.setPlaceholderText("Новый лист...")
        self.add_sheet_btn = QPushButton("➕ Добавить лист")
        self.add_sheet_btn.clicked.connect(self.add_sheet)

        add_sheet_layout.addWidget(self.new_sheet_input)
        add_sheet_layout.addWidget(self.add_sheet_btn)

        sheets_layout.addLayout(add_sheet_layout)

        form_layout.addRow("Листы:", sheets_layout)

        layout.addLayout(form_layout)

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        layout.addWidget(button_box)

        self.setLayout(layout)

    def add_version(self):
        new_version = self.new_version_input.text().strip()
        if new_version:
            row = self.versions_list.rowCount()
            self.versions_list.insertRow(row)

            version_item = QTableWidgetItem(new_version)
            self.versions_list.setItem(row, 0, version_item)

            btn_remove = QPushButton("❌")
            btn_remove.setMaximumWidth(30)
            btn_remove.clicked.connect(lambda checked, r=row: self.remove_version(r))
            self.versions_list.setCellWidget(row, 1, btn_remove)

            self.new_version_input.clear()

    def remove_version(self, row):
        self.versions_list.removeRow(row)

    def add_gk(self):
        new_gk = self.new_gk_input.text().strip()
        if new_gk:
            row = self.gk_list.rowCount()
            self.gk_list.insertRow(row)

            gk_item = QTableWidgetItem(new_gk)
            self.gk_list.setItem(row, 0, gk_item)

            btn_remove = QPushButton("❌")
            btn_remove.setMaximumWidth(30)
            btn_remove.clicked.connect(lambda checked, r=row: self.remove_gk(r))
            self.gk_list.setCellWidget(row, 1, btn_remove)

            self.new_gk_input.clear()

    def remove_gk(self, row):
        self.gk_list.removeRow(row)

    def add_sheet(self):
        new_sheet = self.new_sheet_input.text().strip()
        if new_sheet:
            row = self.sheets_list.rowCount()
            self.sheets_list.insertRow(row)

            sheet_item = QTableWidgetItem(new_sheet)
            self.sheets_list.setItem(row, 0, sheet_item)

            btn_remove = QPushButton("❌")
            btn_remove.setMaximumWidth(30)
            btn_remove.clicked.connect(lambda checked, r=row: self.remove_sheet(r))
            self.sheets_list.setCellWidget(row, 1, btn_remove)

            self.new_sheet_input.clear()

    def remove_sheet(self, row):
        self.sheets_list.removeRow(row)

    def get_updated_data(self):
        versions = []
        for row in range(self.versions_list.rowCount()):
            item = self.versions_list.item(row, 0)
            if item and item.text():
                versions.append(item.text())

        gk = []
        for row in range(self.gk_list.rowCount()):
            item = self.gk_list.item(row, 0)
            if item and item.text():
                gk.append(item.text())

        sheets = []
        for row in range(self.sheets_list.rowCount()):
            item = self.sheets_list.item(row, 0)
            if item and item.text():
                sheets.append(item.text())

        return {
            'name': self.name_edit.text(),
            'versions': versions,
            'gk': gk,
            'sheets': sheets
        }


class ExcelLoaderThread(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, file_path):
        super().__init__()
        self.file_path = file_path

    def find_columns(self, columns):
        """Поиск нужных колонок (из parsingexel.py)"""
        name_cols = []
        version_cols = []
        contract_cols = []

        # Ключевые слова для поиска
        name_keywords = ['наименование', 'название', 'по', 'name', 'product', 'software', 'программа', 'п/о',
                         'подсистема', 'система', 'бпо', 'спо', 'наименование по', 'программное обеспечение']
        version_keywords = ['версия', 'version', 'вер', 'v.', 'var', 'редакция', 'ver', 'требуемая версия', 'required',
                            'версия по', 'версия программы']
        contract_keywords = ['гк', 'контракт', 'contract', 'договор', '№', 'номер', 'gk', 'код', 'ид', 'код гк']

        for idx, col in enumerate(columns):
            col_lower = str(col).lower() if pd.notna(col) else ""

            # Проверяем каждую категорию
            if any(kw in col_lower for kw in name_keywords):
                name_cols.append(idx)
            elif any(kw in col_lower for kw in version_keywords):
                version_cols.append(idx)
            elif any(kw in col_lower for kw in contract_keywords):
                contract_cols.append(idx)

        # Если колонки не найдены, используем первые колонки как запасной вариант
        if not name_cols and len(columns) > 0:
            name_cols = [0]  # Первая колонка как наименование
        if not version_cols and len(columns) > 1:
            version_cols = [1]  # Вторая колонка как версия

        return name_cols, version_cols, contract_cols

    def extract_gk_pattern(self, text):
        """Извлекает только ГК вида ФКУ000/2025 (из parsingexel.py)"""
        if pd.isna(text) or not text:
            return []

        text = str(text).strip()

        # Паттерн для поиска ГК вида ФКУ000/2025
        # ФКУ - буквы, 000 - цифры, /2025 - необязательная часть с годом
        pattern = r'[А-ЯA-Z]{3,4}\d{3,4}(?:[/-]\d{2,4})?'

        # Ищем все совпадения
        matches = re.findall(pattern, text, re.IGNORECASE)

        # Очищаем и фильтруем
        result = []
        for match in matches:
            clean_match = match.strip()
            # Проверяем что это действительно наш формат
            if re.match(r'^[А-ЯA-Z]{3,4}\d{3,4}(?:[/-]\d{2,4})?$', clean_match, re.IGNORECASE):
                result.append(clean_match.upper())  # Приводим к верхнему регистру

        return result

    def normalize_name(self, name):
        """Нормализация имени для поиска"""
        if pd.isna(name) or not name:
            return ""
        name = str(name).lower().strip()
        name = re.sub(r'[^\w\s-]', ' ', name)
        name = re.sub(r'\s+', ' ', name)
        return name

    def parse_version(self, version_str):
        """Парсинг версии в числовое представление"""
        if pd.isna(version_str) or not version_str:
            return [0]

        version_str = str(version_str).lower().strip()

        parts = []
        numbers = re.findall(r'\d+', version_str)
        for num in numbers:
            try:
                parts.append(int(num))
            except ValueError:
                pass

        if not parts:
            return [0]

        # Обработка специальных обозначений
        if 'c' in version_str or 'g' in version_str:
            letter_match = re.search(r'(\d+)([cg])', version_str)
            if letter_match and parts:
                letter_value = 100 if letter_match.group(2) == 'c' else 50
                parts[-1] = parts[-1] * 1000 + letter_value

        if 'sp' in version_str or 'service pack' in version_str:
            sp_match = re.search(r'sp[\s]*(\d+)', version_str, re.IGNORECASE)
            if sp_match and parts:
                sp_num = int(sp_match.group(1))
                parts.append(sp_num * 100)

        return parts

    def run(self):
        try:
            self.progress.emit(5, "Чтение Excel файла...")

            # Открываем файл
            excel_file = pd.ExcelFile(self.file_path)
            sheets = excel_file.sheet_names
            total_sheets = len(sheets)

            self.progress.emit(10, f"Найдено листов: {total_sheets}")

            versions_data = {}
            total_records = 0

            # Обрабатываем каждый лист
            for sheet_idx, sheet_name in enumerate(sheets):
                progress_base = 10 + int(80 * sheet_idx / total_sheets)

                self.progress.emit(
                    progress_base,
                    f"Обработка листа {sheet_idx + 1}/{total_sheets}: {sheet_name}"
                )

                try:
                    # Пробуем разные строки для заголовков
                    df = None
                    header_row = 1  # По умолчанию вторая строка (индекс 1)

                    for try_header in range(5):  # Пробуем первые 5 строк
                        try:
                            df = pd.read_excel(
                                self.file_path,
                                sheet_name=sheet_name,
                                header=try_header,
                                dtype=str
                            )

                            if df.empty:
                                continue

                            # Проверяем, есть ли в этой строке осмысленные заголовки
                            name_cols, version_cols, contract_cols = self.find_columns(df.columns)

                            if name_cols or version_cols:  # Нашли хотя бы одну колонку
                                header_row = try_header
                                break
                        except:
                            continue

                    if df is None or df.empty:
                        continue

                    # Находим колонки
                    name_cols, version_cols, contract_cols = self.find_columns(df.columns)

                    # Обрабатываем каждую строку
                    for row_idx, row in df.iterrows():
                        # Собираем наименования
                        names = []
                        for col_idx in name_cols:
                            if col_idx < len(row):
                                val = str(row.iloc[col_idx]).strip() if pd.notna(row.iloc[col_idx]) else ""
                                if val and val.lower() not in ['nan', 'none', '', 'null'] and len(val) > 2:
                                    names.append(val)

                        if not names:
                            continue

                        # Собираем версии
                        versions = []
                        for col_idx in version_cols:
                            if col_idx < len(row):
                                val = str(row.iloc[col_idx]).strip() if pd.notna(row.iloc[col_idx]) else ""
                                if val and val.lower() not in ['nan', 'none', '', 'null'] and len(val) > 0:
                                    versions.append(val)

                        # Собираем ГК (только нужного формата)
                        all_gk = []
                        for col_idx in contract_cols:
                            if col_idx < len(row):
                                val = str(row.iloc[col_idx]).strip() if pd.notna(row.iloc[col_idx]) else ""
                                if val and val.lower() not in ['nan', 'none', '', 'null']:
                                    # Извлекаем только ГК вида ФКУ000/2025
                                    gk_list = self.extract_gk_pattern(val)
                                    all_gk.extend(gk_list)

                        # Убираем дубликаты ГК
                        unique_gk = []
                        for gk in all_gk:
                            if gk and gk not in unique_gk:
                                unique_gk.append(gk)

                        # Для каждого наименования создаем запись
                        for name in names:
                            if not name:
                                continue

                            norm_name = self.normalize_name(name)

                            if norm_name not in versions_data:
                                versions_data[norm_name] = {
                                    'name': name,
                                    'versions': [],
                                    'gk': [],
                                    'sheets': set(),
                                    'parsed_versions': []
                                }

                            # Добавляем версии
                            for v in versions:
                                if v and v not in versions_data[norm_name]['versions']:
                                    versions_data[norm_name]['versions'].append(v)
                                    parsed = self.parse_version(v)
                                    versions_data[norm_name]['parsed_versions'].append({
                                        'original': v,
                                        'parsed': parsed
                                    })

                            # Добавляем ГК
                            for gk in unique_gk:
                                if gk and gk not in versions_data[norm_name]['gk']:
                                    versions_data[norm_name]['gk'].append(gk)

                            # Добавляем лист
                            versions_data[norm_name]['sheets'].add(sheet_name)

                        total_records += 1

                except Exception as e:
                    logger.error(f"Ошибка обработки листа {sheet_name}: {e}")
                    continue

            # Постобработка данных
            for norm_name, data in versions_data.items():
                # Преобразуем set в list
                data['sheets'] = list(data['sheets'])

                # Убираем дубликаты версий
                if data['parsed_versions']:
                    unique_versions = {}
                    for v in data['parsed_versions']:
                        key = v['original']
                        if key not in unique_versions:
                            unique_versions[key] = v
                    data['parsed_versions'] = list(unique_versions.values())
                    # Сортируем версии по убыванию
                    data['parsed_versions'].sort(key=lambda x: x['parsed'], reverse=True)

                # Убираем дубликаты версий в простом списке
                data['versions'] = list(set(data['versions']))

            self.progress.emit(100, f"Загружено {len(versions_data)} записей")

            result = {
                'data': versions_data,
                'total_records': total_records,
                'unique_names': len(versions_data),
                'file_name': os.path.basename(self.file_path)
            }

            self.finished.emit(result)

        except Exception as e:
            self.error.emit(str(e))


class VersionsDialog(QDialog):
    def __init__(self, parent=None, document_text="", page_info=None):
        super().__init__(parent)
        self.document_text = document_text
        self.page_info = page_info or []
        self.versions_data = []
        self.filtered_data = []
        self.excel_versions_data = {}
        self.comparison_results = []
        self.software_patterns = self.get_software_patterns()
        self.category_keywords = self.build_category_keywords()
        self.excel_file_path = None
        self.excel_loader = None
        self.data_manager = ExcelDataManager()
        self.auto_load_enabled = True

        self.setWindowTitle("Версии БПО/СПО в документе")
        self.resize(1400, 950)

        if parent and hasattr(parent, 'styleSheet'):
            self.setStyleSheet(parent.styleSheet())

        self.init_ui()
        self.scan_document_for_versions()
        self.display_versions()
        self.load_saved_excel_data()

    def build_category_keywords(self):
        keywords = {}

        for category, patterns in self.software_patterns.items():
            category_keywords = []
            for pattern, product_name in patterns:
                words = re.findall(r'[А-Яа-яA-Za-z]+', product_name)
                category_keywords.extend(words)

            if 'Операционные системы' in category:
                category_keywords.extend(['ос', 'операционная', 'windows', 'linux', 'macos', 'ubuntu', 'debian'])
            elif 'Базы данных' in category:
                category_keywords.extend(['бд', 'база', 'database', 'sql', 'oracle', 'mysql', 'postgresql'])
            elif 'Виртуализация' in category:
                category_keywords.extend(['виртуализация', 'vmware', 'virtualbox', 'docker', 'kubernetes'])
            elif 'Веб-серверы' in category:
                category_keywords.extend(['веб', 'web', 'сервер', 'apache', 'nginx', 'iis'])
            elif 'Языки программирования' in category:
                category_keywords.extend(['язык', 'программирование', 'java', 'python', 'php', 'javascript'])
            elif 'Сетевые устройства' in category:
                category_keywords.extend(['сеть', 'сетевой', 'маршрутизатор', 'коммутатор', 'cisco', 'juniper'])
            elif 'Российское ПО' in category:
                category_keywords.extend(['российский', 'отечественный', 'мойофис', 'астра', 'ред ос'])

            keywords[category] = list(set(category_keywords))

        return keywords

    def determine_category_from_excel(self, excel_name):
        excel_name_lower = excel_name.lower()

        for category, keywords in self.category_keywords.items():
            for keyword in keywords:
                if keyword.lower() in excel_name_lower:
                    return category

        if any(word in excel_name_lower for word in ['windows', 'linux', 'macos', 'ос']):
            return 'Операционные системы'
        elif any(word in excel_name_lower for word in ['база', 'database', 'sql']):
            return 'Базы данных'
        elif any(word in excel_name_lower for word in ['виртуал', 'vmware', 'docker']):
            return 'Виртуализация и контейнеры'
        elif any(word in excel_name_lower for word in ['веб', 'web', 'сервер']):
            return 'Веб-серверы и прокси'
        elif any(word in excel_name_lower for word in ['java', 'python', 'php', 'язык']):
            return 'Языки программирования и среды'
        elif any(word in excel_name_lower for word in ['сеть', 'маршрутизатор', 'коммутатор']):
            return 'Сетевые устройства'
        elif any(word in excel_name_lower for word in ['российский', 'отечественный']):
            return 'Российское ПО'

        return 'Прочее ПО'

    def get_software_patterns(self):
        return {
            'Операционные системы': [
                (r'Windows\s+(?:Server\s+)?(\d+(?:\.\d+)?(?:\s*(?:R2|SP\d+))?)', 'Windows'),
                (r'Windows\s+(\d+(?:\.\d+)?)\s*(?:Pro|Enterprise|Home|Education)?', 'Windows'),
                (r'Windows\s+Server\s+(\d+(?:\.\d+)?(?:\s*R2)?)', 'Windows Server'),
                (r'Ubuntu\s+(\d+(?:\.\d+)?(?:\s*LTS)?)', 'Ubuntu'),
                (r'Debian\s+(\d+(?:\.\d+)?)', 'Debian'),
                (r'CentOS\s+(\d+(?:\.\d+)?)', 'CentOS'),
                (r'Red\s*Hat\s+(?:Enterprise\s+Linux\s+)?(\d+(?:\.\d+)?)', 'Red Hat'),
                (r'Astra\s+Linux\s+(?:\w+\s+)?(\d+(?:\.\d+)?(?:\s*\.\d+)?)', 'Astra Linux'),
                (r'РЕД\s+ОС\s+(\d+(?:\.\d+)?)', 'РЕД ОС'),
                (r'ALT\s+Linux\s+(\d+(?:\.\d+)?)', 'ALT Linux'),
                (r'ROSA\s+Linux\s+(\d+(?:\.\d+)?)', 'ROSA Linux'),
                (
                    r'macOS\s+(?:Sequoia|Sonoma|Ventura|Monterey|Big Sur|Catalina|Mojave|High Sierra|Sierra|El Capitan|Yosemite|Mavericks|Mountain Lion|Lion|Snow Leopard|Leopard|Tiger|Panther|Jaguar|Puma|Cheetah)\s*(?:(\d+(?:\.\d+)?))?',
                    'macOS'),
            ],
            'Базы данных': [
                (r'Oracle\s+(?:Database\s+)?(\d+(?:[cg]\d+)?(?:\.\d+)?)', 'Oracle'),
                (r'Oracle\s+(\d+(?:\.\d+)?[a-z]?)', 'Oracle'),
                (r'MySQL\s+(\d+(?:\.\d+)?)', 'MySQL'),
                (r'PostgreSQL\s+(\d+(?:\.\d+)?)', 'PostgreSQL'),
                (r'MongoDB\s+(\d+(?:\.\d+)?)', 'MongoDB'),
                (r'Redis\s+(\d+(?:\.\d+)?)', 'Redis'),
                (r'Microsoft\s+SQL\s+Server\s+(\d+(?:\.\d+)?)', 'MS SQL Server'),
                (r'SQLite\s+(\d+(?:\.\d+)?)', 'SQLite'),
                (r'MariaDB\s+(\d+(?:\.\d+)?)', 'MariaDB'),
                (r'Cassandra\s+(\d+(?:\.\d+)?)', 'Cassandra'),
            ],
            'Виртуализация и контейнеры': [
                (r'VMware\s+(?:vSphere|ESXi|Workstation|Fusion)?\s*(\d+(?:\.\d+)?)', 'VMware'),
                (r'VirtualBox\s+(\d+(?:\.\d+)?)', 'VirtualBox'),
                (r'Docker\s+(?:version\s+)?(\d+(?:\.\d+)?)', 'Docker'),
                (r'Kubernetes\s+(\d+(?:\.\d+)?)', 'Kubernetes'),
                (r'OpenStack\s+(\d+(?:\.\d+)?)', 'OpenStack'),
                (r'Hyper-V\s+(\d+(?:\.\d+)?)', 'Hyper-V'),
                (r'KVM\s+(\d+(?:\.\d+)?)', 'KVM'),
                (r'Proxmox\s+(\d+(?:\.\d+)?)', 'Proxmox'),
            ],
            'Веб-серверы и прокси': [
                (r'Apache\s+(?:HTTP\s+Server\s+)?(\d+(?:\.\d+)?)', 'Apache'),
                (r'Nginx\s+(\d+(?:\.\d+)?)', 'Nginx'),
                (r'IIS\s+(\d+(?:\.\d+)?)', 'IIS'),
                (r'Tomcat\s+(\d+(?:\.\d+)?)', 'Tomcat'),
                (r'Jetty\s+(\d+(?:\.\d+)?)', 'Jetty'),
                (r'HAProxy\s+(\d+(?:\.\d+)?)', 'HAProxy'),
            ],
            'Языки программирования и среды': [
                (r'Java\s+(?:version\s+)?(\d+(?:\.\d+)?(?:_\d+)?)', 'Java'),
                (r'Python\s+(\d+(?:\.\d+)?)', 'Python'),
                (r'Node\.js\s+(\d+(?:\.\d+)?)', 'Node.js'),
                (r'PHP\s+(\d+(?:\.\d+)?)', 'PHP'),
                (r'Ruby\s+(\d+(?:\.\d+)?)', 'Ruby'),
                (r'Go\s+(?:version\s+)?(\d+(?:\.\d+)?)', 'Go'),
                (r'.NET\s+(?:Core\s+)?(\d+(?:\.\d+)?)', '.NET'),
                (r'Django\s+(\d+(?:\.\d+)?)', 'Django'),
            ],
            'Сетевые устройства': [
                (r'Cisco\s+(?:IOS|NX-OS|IOS-XE)\s+(\d+(?:\.\d+)?(?:\([\d.]+\))?)', 'Cisco IOS'),
                (r'Juniper\s+(?:Junos|JunOS)\s+(\d+(?:\.\d+)?(?:R\d)?)', 'Juniper Junos'),
                (r'Palo\s+Alto\s+(?:PAN-OS\s+)?(\d+(?:\.\d+)?)', 'Palo Alto PAN-OS'),
                (r'Check\s+Point\s+(?:R|Version)?\s*(\d+(?:\.\d+)?(?:\.\d+)?)', 'Check Point'),
                (r'MikroTik\s+RouterOS\s+(\d+(?:\.\d+)?)', 'MikroTik RouterOS'),
                (r'D-Link\s+(\w+)?\s*(?:firmware\s+)?(\d+(?:\.\d+)?)', 'D-Link'),
            ],
            'Российское ПО': [
                (r'МойОфис\s+(\d+(?:\.\d+)?)', 'МойОфис'),
                (r'Р7-Офис\s+(\d+(?:\.\d+)?)', 'Р7-Офис'),
                (r'P7-Office\s+(\d+(?:\.\d+)?)', 'P7-Office'),
                (r'OnlyOffice\s+(\d+(?:\.\d+)?)', 'OnlyOffice'),
                (r'1С:Предприятие\s+(\d+(?:\.\d+)?)', '1С:Предприятие'),
                (r'1С:Enterprise\s+(\d+(?:\.\d+)?)', '1С:Enterprise'),
                (r'VK\s+WorkSpace\s+(\d+(?:\.\d+)?)', 'VK WorkSpace'),
            ],
            'Прочее ПО': [
                (r'Elasticsearch\s+(\d+(?:\.\d+)?)', 'Elasticsearch'),
                (r'Kibana\s+(\d+(?:\.\d+)?)', 'Kibana'),
                (r'Logstash\s+(\d+(?:\.\d+)?)', 'Logstash'),
                (r'Grafana\s+(\d+(?:\.\d+)?)', 'Grafana'),
                (r'Prometheus\s+(\d+(?:\.\d+)?)', 'Prometheus'),
                (r'Zabbix\s+(\d+(?:\.\d+)?)', 'Zabbix'),
                (r'Ansible\s+(\d+(?:\.\d+)?)', 'Ansible'),
                (r'Terraform\s+(\d+(?:\.\d+)?)', 'Terraform'),
            ]
        }

    def init_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        self.tab_widget = QTabWidget()

        self.document_tab = QWidget()
        self.init_document_tab()
        self.tab_widget.addTab(self.document_tab, "📄 Версии из документа")

        self.comparison_tab = QWidget()
        self.init_comparison_tab()
        self.tab_widget.addTab(self.comparison_tab, "🔍 Сравнение с Excel (>=)")

        self.excel_tab = QWidget()
        self.init_excel_tab()
        self.tab_widget.addTab(self.excel_tab, "📊 Данные из Excel")

        self.gk_search_tab = QWidget()
        self.init_gk_search_tab()
        self.tab_widget.addTab(self.gk_search_tab, "🔎 Поиск по ГК")

        # Добавляем новую вкладку для поиска продуктов из Excel в документе
        self.excel_products_tab = QWidget()
        self.init_excel_products_tab()
        self.tab_widget.addTab(self.excel_products_tab, "📋 Поиск продуктов из Excel")

        main_layout.addWidget(self.tab_widget)

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

    def init_excel_products_tab(self):
        """Инициализация вкладки поиска продуктов из Excel"""
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        info_layout = QHBoxLayout()

        self.excel_products_info = QLabel("📊 Для поиска загрузите файл Excel с составом ПО")
        info_layout.addWidget(self.excel_products_info)
        info_layout.addStretch()

        self.load_excel_products_btn = QPushButton("📂 Загрузить Excel файл с составом ПО")
        self.load_excel_products_btn.clicked.connect(self.search_excel_products)
        self.load_excel_products_btn.setMinimumHeight(35)
        self.load_excel_products_btn.setStyleSheet("""
            QPushButton {
                background-color: #9b59b6;
                color: white;
                font-weight: bold;
                border-radius: 5px;
                padding: 8px 15px;
            }
            QPushButton:hover {
                background-color: #8e44ad;
            }
        """)
        info_layout.addWidget(self.load_excel_products_btn)

        layout.addLayout(info_layout)

        # Добавляем описание
        desc_label = QLabel(
            "Этот инструмент ищет в документе продукты, указанные в Excel-файле "
            "(по наименованию, модулю/компоненту) и сравнивает версии."
        )
        desc_label.setWordWrap(True)
        desc_label.setStyleSheet("color: #7f8c8d; padding: 5px;")
        layout.addWidget(desc_label)

        layout.addStretch()

        self.excel_products_tab.setLayout(layout)

    def init_gk_search_tab(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        info_layout = QHBoxLayout()

        self.gk_doc_info = QLabel(f"📄 Документ: {len(self.document_text)} символов")
        self.gk_excel_info = QLabel(f"📊 Данные Excel: {len(self.excel_versions_data)} записей")

        info_layout.addWidget(self.gk_doc_info)
        info_layout.addWidget(self.gk_excel_info)
        info_layout.addStretch()

        self.search_gk_btn = QPushButton("🔍 Найти ГК в документе")
        self.search_gk_btn.clicked.connect(self.search_gk_in_current_tab)
        self.search_gk_btn.setMinimumHeight(35)
        self.search_gk_btn.setStyleSheet("""
            QPushButton {
                background-color: #e67e22;
                color: white;
                font-weight: bold;
                border-radius: 5px;
                padding: 8px 15px;
            }
            QPushButton:hover {
                background-color: #d35400;
            }
        """)
        info_layout.addWidget(self.search_gk_btn)

        layout.addLayout(info_layout)

        splitter = QSplitter(Qt.Orientation.Vertical)

        gk_widget = QWidget()
        gk_layout = QVBoxLayout(gk_widget)
        gk_layout.setContentsMargins(0, 0, 0, 0)

        gk_label = QLabel("Найденные ГК в документе:")
        gk_label.setStyleSheet("font-weight: bold; font-size: 12px;")
        gk_layout.addWidget(gk_label)

        self.gk_tab_table = QTableWidget()
        self.gk_tab_table.setColumnCount(4)
        self.gk_tab_table.setHorizontalHeaderLabels(["№", "ГК", "Статус в Excel", "Кол-во версий"])
        self.gk_tab_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.gk_tab_table.setAlternatingRowColors(True)

        gk_layout.addWidget(self.gk_tab_table)

        versions_widget = QWidget()
        versions_layout = QVBoxLayout(versions_widget)
        versions_layout.setContentsMargins(0, 0, 0, 0)

        versions_label = QLabel("Версии ПО:")
        versions_label.setStyleSheet("font-weight: bold; font-size: 12px;")
        versions_layout.addWidget(versions_label)

        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Поиск:"))
        self.gk_version_search = QLineEdit()
        self.gk_version_search.setPlaceholderText("Поиск по названию или версии...")
        self.gk_version_search.textChanged.connect(self.filter_gk_versions)
        filter_layout.addWidget(self.gk_version_search)

        self.gk_category_filter = QComboBox()
        self.gk_category_filter.addItem("Все категории")
        self.gk_category_filter.currentTextChanged.connect(self.filter_gk_versions)
        filter_layout.addWidget(QLabel("Категория:"))
        filter_layout.addWidget(self.gk_category_filter)

        filter_layout.addStretch()
        versions_layout.addLayout(filter_layout)

        self.gk_versions_table = QTableWidget()
        self.gk_versions_table.setColumnCount(7)
        self.gk_versions_table.setHorizontalHeaderLabels([
            "ГК", "Наименование ПО", "Версия", "Категория",
            "Лист Excel", "Статус", "Действие"
        ])

        header = self.gk_versions_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.Fixed)

        self.gk_versions_table.setColumnWidth(6, 100)
        self.gk_versions_table.setAlternatingRowColors(True)

        versions_layout.addWidget(self.gk_versions_table)

        splitter.addWidget(gk_widget)
        splitter.addWidget(versions_widget)
        splitter.setSizes([200, 400])

        layout.addWidget(splitter)

        stats_layout = QHBoxLayout()
        self.gk_tab_stats = QLabel("Нажмите кнопку 'Найти ГК в документе' для поиска")
        self.gk_tab_stats.setStyleSheet("font-weight: bold; padding: 5px;")
        stats_layout.addWidget(self.gk_tab_stats)
        stats_layout.addStretch()

        export_gk_btn = QPushButton("📤 Экспорт результатов")
        export_gk_btn.clicked.connect(self.export_gk_results)
        stats_layout.addWidget(export_gk_btn)

        layout.addLayout(stats_layout)

        self.gk_search_tab.setLayout(layout)

    def search_excel_products(self):
        """Поиск продуктов из Excel в документе"""
        if not self.document_text:
            QMessageBox.warning(self, "Ошибка", "Нет загруженного документа")
            return

        # Выбираем файл Excel
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл с составом ПО", "",
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )

        if not file_path:
            return

        # Создаем и показываем диалог поиска
        dialog = ExcelProductSearchDialog(self, self.document_text, file_path)
        dialog.exec()

    def search_gk_in_current_tab(self):
        if not self.document_text:
            QMessageBox.warning(self, "Ошибка", "Нет загруженного документа")
            return

        found_gk = self.extract_gk_from_text(self.document_text)

        self.gk_tab_table.setRowCount(len(found_gk))

        categories = set()
        all_versions = []

        for i, gk in enumerate(found_gk):
            self.gk_tab_table.setItem(i, 0, QTableWidgetItem(str(i + 1)))

            gk_item = QTableWidgetItem(gk)
            gk_item.setData(Qt.ItemDataRole.UserRole, gk)
            self.gk_tab_table.setItem(i, 1, gk_item)

            found_in_excel, versions_count = self.find_gk_in_excel_with_count(gk)

            status_item = QTableWidgetItem()
            if found_in_excel:
                status_item.setText("✅ Найден в Excel")
                status_item.setForeground(QColor(0, 150, 0))

                versions = self.get_versions_for_gk(gk)
                all_versions.extend(versions)
                for v in versions:
                    categories.add(v['category'])
            else:
                status_item.setText("❌ Не найден в Excel")
                status_item.setForeground(QColor(255, 0, 0))

            self.gk_tab_table.setItem(i, 2, status_item)

            self.gk_tab_table.setItem(i, 3, QTableWidgetItem(str(versions_count)))

        self.gk_tab_table.resizeRowsToContents()

        if not all_versions and self.excel_versions_data:
            for norm_name, data in self.excel_versions_data.items():
                status = "❓ ГК не найден в документе"
                if data.get('parsed_versions'):
                    for ver_info in data['parsed_versions']:
                        all_versions.append({
                            'gk': '—',
                            'name': data['name'],
                            'version': ver_info['original'],
                            'category': self.determine_category_from_excel(data['name']),
                            'sheets': ', '.join(data.get('sheets', [])),
                            'status': status,
                            'status_color': QColor(255, 140, 0)
                        })
                        categories.add(self.determine_category_from_excel(data['name']))
                else:
                    for version in data.get('versions', []):
                        all_versions.append({
                            'gk': '—',
                            'name': data['name'],
                            'version': version,
                            'category': self.determine_category_from_excel(data['name']),
                            'sheets': ', '.join(data.get('sheets', [])),
                            'status': status,
                            'status_color': QColor(255, 140, 0)
                        })
                        categories.add(self.determine_category_from_excel(data['name']))

        self.display_gk_versions(all_versions, categories)

        gk_found = sum(1 for row in range(self.gk_tab_table.rowCount())
                       if "✅" in self.gk_tab_table.item(row, 2).text())
        gk_not_found = len(found_gk) - gk_found

        stats = f"Найдено ГК: {len(found_gk)} | В Excel: {gk_found} | Не в Excel: {gk_not_found} | Версий: {len(all_versions)}"
        self.gk_tab_stats.setText(stats)

    def display_gk_versions(self, versions, categories):
        self.gk_versions_table.setRowCount(len(versions))

        for i, ver in enumerate(versions):
            gk_item = QTableWidgetItem(ver['gk'])
            if ver['gk'] == '—':
                gk_item.setForeground(QColor(128, 128, 128))
            self.gk_versions_table.setItem(i, 0, gk_item)

            self.gk_versions_table.setItem(i, 1, QTableWidgetItem(ver['name']))

            self.gk_versions_table.setItem(i, 2, QTableWidgetItem(ver['version']))

            category_item = QTableWidgetItem(ver['category'])
            self.gk_versions_table.setItem(i, 3, category_item)

            self.gk_versions_table.setItem(i, 4, QTableWidgetItem(ver['sheets']))

            status_item = QTableWidgetItem(ver['status'])
            status_item.setForeground(ver.get('status_color', QColor(0, 0, 0)))
            self.gk_versions_table.setItem(i, 5, status_item)

            copy_btn = QPushButton("📋 Копировать")
            copy_btn.setMaximumWidth(80)
            copy_btn.setProperty('row', i)
            copy_btn.clicked.connect(lambda checked, r=i: self.copy_gk_version_row(r))
            self.gk_versions_table.setCellWidget(i, 6, copy_btn)

        self.gk_versions_table.resizeRowsToContents()

        self.gk_category_filter.clear()
        self.gk_category_filter.addItem("Все категории")
        self.gk_category_filter.addItems(sorted(categories))

    def filter_gk_versions(self):
        search_text = self.gk_version_search.text().lower()
        category = self.gk_category_filter.currentText()

        for row in range(self.gk_versions_table.rowCount()):
            show_row = True

            if search_text:
                row_text = ""
                for col in [0, 1, 2]:
                    item = self.gk_versions_table.item(row, col)
                    if item:
                        row_text += item.text().lower() + " "

                if search_text not in row_text:
                    show_row = False

            if category != "Все категории":
                cat_item = self.gk_versions_table.item(row, 3)
                if cat_item and cat_item.text() != category:
                    show_row = False

            self.gk_versions_table.setRowHidden(row, not show_row)

    def copy_gk_version_row(self, row):
        text_parts = []
        for col in range(self.gk_versions_table.columnCount() - 1):
            item = self.gk_versions_table.item(row, col)
            if item:
                text_parts.append(item.text())
            else:
                text_parts.append("")

        QApplication.clipboard().setText('\t'.join(text_parts))

        QToolTip.showText(self.mapToGlobal(self.rect().center()), "Строка скопирована")

    def export_gk_results(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить результаты ГК", f"gk_search_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "CSV files (*.csv)"
        )

        if file_path:
            try:
                import csv
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(["ГК", "Наименование ПО", "Версия", "Категория", "Лист Excel", "Статус"])

                    for row in range(self.gk_versions_table.rowCount()):
                        if not self.gk_versions_table.isRowHidden(row):
                            row_data = []
                            for col in range(self.gk_versions_table.columnCount() - 1):
                                item = self.gk_versions_table.item(row, col)
                                if item:
                                    row_data.append(item.text())
                                else:
                                    row_data.append("")
                            writer.writerow(row_data)

                QMessageBox.information(self, "Успех", f"Результаты экспортированы в {file_path}")
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось экспортировать: {str(e)}")

    def extract_gk_from_text(self, text):
        if not text:
            return []

        pattern = r'[А-ЯA-Z]{2,5}[-\s]?\d{2,5}(?:[/-]\d{2,4})?(?:\s*/\s*\d+)?'
        matches = re.findall(pattern, text, re.IGNORECASE)

        result = []
        for match in matches:
            clean_match = match.strip().replace(' ', '')
            clean_match = re.sub(r'[^\w/-]', '', clean_match)
            if len(clean_match) >= 5:
                result.append(clean_match.upper())

        return list(set(result))

    def find_gk_in_excel_with_count(self, gk):
        if not self.excel_versions_data:
            return False, 0

        versions_count = 0
        found = False

        for norm_name, data in self.excel_versions_data.items():
            if gk in data.get('gk', []):
                found = True
                if data.get('parsed_versions'):
                    versions_count += len(data['parsed_versions'])
                else:
                    versions_count += len(data.get('versions', []))

        return found, versions_count

    def get_versions_for_gk(self, gk):
        versions = []

        if not self.excel_versions_data:
            return versions

        for norm_name, data in self.excel_versions_data.items():
            if gk in data.get('gk', []):
                if data.get('parsed_versions'):
                    for ver_info in data['parsed_versions']:
                        versions.append({
                            'gk': gk,
                            'name': data['name'],
                            'version': ver_info['original'],
                            'category': self.determine_category_from_excel(data['name']),
                            'sheets': ', '.join(data.get('sheets', [])),
                            'status': '✅ Найден в Excel',
                            'status_color': QColor(0, 150, 0)
                        })
                else:
                    for version in data.get('versions', []):
                        versions.append({
                            'gk': gk,
                            'name': data['name'],
                            'version': version,
                            'category': self.determine_category_from_excel(data['name']),
                            'sheets': ', '.join(data.get('sheets', [])),
                            'status': '✅ Найден в Excel',
                            'status_color': QColor(0, 150, 0)
                        })

        return versions

    def init_document_tab(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        toolbar_widget = self.create_document_toolbar()
        layout.addWidget(toolbar_widget)

        filter_widget = self.create_document_filter_panel()
        layout.addWidget(filter_widget)

        splitter = QSplitter(Qt.Orientation.Vertical)

        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "Категория", "Продукт", "Версия", "Страница",
            "Строка", "Контекст", "Действие"
        ])

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.Fixed)

        self.table.setColumnWidth(0, 150)
        self.table.setColumnWidth(1, 200)
        self.table.setColumnWidth(2, 100)
        self.table.setColumnWidth(3, 80)
        self.table.setColumnWidth(4, 80)
        self.table.setColumnWidth(6, 100)

        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)
        self.table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)

        self.table.doubleClicked.connect(self.go_to_version_position)

        splitter.addWidget(self.table)

        context_panel = QGroupBox("Контекст найденной версии")
        context_layout = QVBoxLayout(context_panel)

        self.context_text = QTextEdit()
        self.context_text.setReadOnly(True)
        self.context_text.setMaximumHeight(150)
        self.context_text.setFont(QFont("Consolas", 10))

        context_layout.addWidget(self.context_text)

        splitter.addWidget(context_panel)
        splitter.setSizes([600, 200])

        layout.addWidget(splitter)

        stats_layout = QHBoxLayout()
        self.total_label = QLabel("Всего найдено версий: 0")
        self.categories_label = QLabel("Категорий: 0")
        self.unique_label = QLabel("Уникальных продуктов: 0")

        stats_layout.addWidget(self.total_label)
        stats_layout.addWidget(self.categories_label)
        stats_layout.addWidget(self.unique_label)
        stats_layout.addStretch()

        layout.addLayout(stats_layout)

        self.document_tab.setLayout(layout)

    def init_comparison_tab(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        excel_panel = QGroupBox("Загрузка данных из Excel")
        excel_layout = QHBoxLayout(excel_panel)

        self.load_excel_btn = QPushButton("📂 Загрузить Excel файл с версиями")
        self.load_excel_btn.clicked.connect(self.load_excel_file)
        self.load_excel_btn.setMinimumHeight(35)
        self.load_excel_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                font-weight: bold;
                border: none;
                border-radius: 5px;
                padding: 8px 15px;
            }
            QPushButton:hover {
                background-color: #229954;
            }
        """)

        self.excel_file_label = QLabel("Файл не загружен")
        self.excel_file_label.setStyleSheet("color: gray;")

        self.excel_progress = QProgressBar()
        self.excel_progress.setVisible(False)
        self.excel_progress.setMaximumHeight(20)

        excel_layout.addWidget(self.load_excel_btn)
        excel_layout.addWidget(self.excel_file_label, 1)
        excel_layout.addWidget(self.excel_progress)

        layout.addWidget(excel_panel)

        data_panel = QGroupBox("Управление данными")
        data_layout = QHBoxLayout(data_panel)

        self.save_data_btn = QPushButton("💾 Сохранить данные")
        self.save_data_btn.clicked.connect(self.save_current_data)
        self.save_data_btn.setEnabled(False)
        self.save_data_btn.setMinimumHeight(30)

        self.clear_data_btn = QPushButton("🗑 Очистить сохраненные данные")
        self.clear_data_btn.clicked.connect(self.clear_saved_data)
        self.clear_data_btn.setMinimumHeight(30)
        self.clear_data_btn.setStyleSheet("""
            QPushButton {
                background-color: #e74c3c;
                color: white;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
        """)

        self.auto_load_check = QCheckBox("Автозагрузка при запуске")
        self.auto_load_check.setChecked(self.auto_load_enabled)
        self.auto_load_check.stateChanged.connect(self.toggle_auto_load)

        data_layout.addWidget(self.save_data_btn)
        data_layout.addWidget(self.clear_data_btn)
        data_layout.addStretch()
        data_layout.addWidget(self.auto_load_check)

        layout.addWidget(data_panel)

        compare_panel = QGroupBox("Параметры сравнения")
        compare_layout = QHBoxLayout(compare_panel)

        self.compare_btn = QPushButton("🔄 Сравнить версии (документ >= Excel)")
        self.compare_btn.clicked.connect(self.compare_versions_with_excel)
        self.compare_btn.setEnabled(False)
        self.compare_btn.setMinimumHeight(35)
        self.compare_btn.setStyleSheet("""
            QPushButton {
                background-color: #f39c12;
                color: white;
                font-weight: bold;
                border: none;
                border-radius: 5px;
                padding: 8px 15px;
            }
            QPushButton:hover {
                background-color: #e67e22;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
        """)

        self.compare_filter = QComboBox()
        self.compare_filter.addItems(["Все", "✅ Соответствует (>=)", "❌ Не соответствует", "❓ Не найдено в Excel"])
        self.compare_filter.currentTextChanged.connect(self.filter_comparison_table)

        compare_layout.addWidget(self.compare_btn)
        compare_layout.addWidget(QLabel("Фильтр:"))
        compare_layout.addWidget(self.compare_filter)
        compare_layout.addStretch()

        layout.addWidget(compare_panel)

        self.comparison_table = QTableWidget()
        self.comparison_table.setColumnCount(10)
        self.comparison_table.setHorizontalHeaderLabels([
            "Продукт (документ)", "Категория", "Версия (документ)",
            "Продукт (Excel)", "Требуемая версия", "ГК",
            "Сравнение", "Статус", "Детали", "Действие"
        ])

        header = self.comparison_table.horizontalHeader()
        for i in range(10):
            header.setSectionResizeMode(i, QHeaderView.ResizeMode.Interactive)

        self.comparison_table.setColumnWidth(0, 180)
        self.comparison_table.setColumnWidth(1, 130)
        self.comparison_table.setColumnWidth(2, 100)
        self.comparison_table.setColumnWidth(3, 180)
        self.comparison_table.setColumnWidth(4, 100)
        self.comparison_table.setColumnWidth(5, 100)
        self.comparison_table.setColumnWidth(6, 100)
        self.comparison_table.setColumnWidth(7, 150)
        self.comparison_table.setColumnWidth(8, 200)
        self.comparison_table.setColumnWidth(9, 80)

        self.comparison_table.setAlternatingRowColors(True)
        self.comparison_table.setSortingEnabled(True)

        self.comparison_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.comparison_table.customContextMenuRequested.connect(self.show_comparison_context_menu)

        layout.addWidget(self.comparison_table)

        self.comparison_stats = QLabel("Загрузите Excel файл для сравнения")
        self.comparison_stats.setStyleSheet("font-weight: bold; padding: 5px;")
        layout.addWidget(self.comparison_stats)

        self.comparison_tab.setLayout(layout)

    def init_excel_tab(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        top_panel = QHBoxLayout()

        search_layout = QHBoxLayout()
        self.excel_search_input = QLineEdit()
        self.excel_search_input.setPlaceholderText("Поиск по названию...")
        self.excel_search_input.textChanged.connect(self.filter_excel_table)

        search_layout.addWidget(QLabel("🔍 Поиск:"))
        search_layout.addWidget(self.excel_search_input)

        top_panel.addLayout(search_layout)
        top_panel.addStretch()

        self.add_record_btn = QPushButton("➕ Добавить запись")
        self.add_record_btn.clicked.connect(self.add_excel_record)
        self.add_record_btn.setEnabled(False)

        self.export_excel_btn = QPushButton("📤 Экспорт данных")
        self.export_excel_btn.clicked.connect(self.export_excel_data)
        self.export_excel_btn.setEnabled(False)

        top_panel.addWidget(self.add_record_btn)
        top_panel.addWidget(self.export_excel_btn)

        layout.addLayout(top_panel)

        self.excel_table = QTableWidget()
        self.excel_table.setColumnCount(6)
        self.excel_table.setHorizontalHeaderLabels([
            "Наименование", "Версии", "ГК", "Листы", "Действие", "Редактировать"
        ])

        header = self.excel_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Fixed)

        self.excel_table.setColumnWidth(4, 100)
        self.excel_table.setColumnWidth(5, 100)
        self.excel_table.setAlternatingRowColors(True)
        self.excel_table.setSortingEnabled(True)

        self.excel_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.excel_table.customContextMenuRequested.connect(self.show_excel_context_menu)

        layout.addWidget(self.excel_table)

        self.excel_stats = QLabel("Загрузите Excel файл для просмотра данных")
        self.excel_stats.setStyleSheet("padding: 5px;")
        layout.addWidget(self.excel_stats)

        self.excel_tab.setLayout(layout)

    def create_document_toolbar(self):
        toolbar_widget = QWidget()
        toolbar_layout = QHBoxLayout(toolbar_widget)
        toolbar_layout.setContentsMargins(0, 0, 0, 0)
        toolbar_layout.setSpacing(10)

        export_csv_btn = QPushButton("📊 Экспорт в CSV")
        export_csv_btn.clicked.connect(self.export_to_csv)

        export_html_btn = QPushButton("🌐 Экспорт в HTML")
        export_html_btn.clicked.connect(self.export_to_html)

        copy_all_btn = QPushButton("📋 Копировать все")
        copy_all_btn.clicked.connect(self.copy_all_to_clipboard)

        refresh_btn = QPushButton("🔄 Обновить")
        refresh_btn.clicked.connect(self.refresh_versions)

        toolbar_layout.addWidget(export_csv_btn)
        toolbar_layout.addWidget(export_html_btn)
        toolbar_layout.addWidget(copy_all_btn)
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(refresh_btn)

        return toolbar_widget

    def create_document_filter_panel(self):
        filter_widget = QWidget()
        filter_layout = QHBoxLayout(filter_widget)
        filter_layout.setContentsMargins(0, 0, 0, 0)
        filter_layout.setSpacing(10)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Поиск по продукту или версии...")
        self.search_input.textChanged.connect(self.filter_table)

        self.category_filter = QComboBox()
        self.category_filter.addItem("Все категории")
        self.category_filter.currentTextChanged.connect(self.filter_table)

        self.version_filter = QComboBox()
        self.version_filter.addItem("Все версии")
        self.version_filter.currentTextChanged.connect(self.filter_table)

        filter_layout.addWidget(QLabel("🔍 Поиск:"))
        filter_layout.addWidget(self.search_input)
        filter_layout.addWidget(QLabel("📁 Категория:"))
        filter_layout.addWidget(self.category_filter)
        filter_layout.addWidget(QLabel("🔢 Версия:"))
        filter_layout.addWidget(self.version_filter)
        filter_layout.addStretch()

        return filter_widget

    def parse_version(self, version_str):
        if not version_str:
            return [0]

        version_str = str(version_str).lower().strip()

        parts = []
        numbers = re.findall(r'\d+', version_str)
        for num in numbers:
            try:
                parts.append(int(num))
            except ValueError:
                pass

        if 'c' in version_str or 'g' in version_str:
            letter_match = re.search(r'(\d+)([cg])', version_str)
            if letter_match and parts:
                letter_value = 100 if letter_match.group(2) == 'c' else 50
                parts[-1] = parts[-1] * 1000 + letter_value

        if 'sp' in version_str or 'service pack' in version_str:
            sp_match = re.search(r'sp[\s]*(\d+)', version_str, re.IGNORECASE)
            if sp_match and parts:
                sp_num = int(sp_match.group(1))
                parts.append(sp_num * 100)

        if 'r' in version_str and not 'red' in version_str:
            r_match = re.search(r'r[\s]*(\d+)', version_str, re.IGNORECASE)
            if r_match and parts:
                r_num = int(r_match.group(1))
                parts.append(r_num)

        if 'update' in version_str:
            u_match = re.search(r'update[\s]*(\d+)', version_str, re.IGNORECASE)
            if u_match and parts:
                u_num = int(u_match.group(1))
                parts.append(u_num)

        if 'build' in version_str:
            b_match = re.search(r'build[\s]*(\d+)', version_str, re.IGNORECASE)
            if b_match and parts:
                b_num = int(b_match.group(1))
                parts.append(b_num)

        return parts if parts else [0]

    def compare_versions(self, v1, v2):
        v1_parts = self.parse_version(v1)
        v2_parts = self.parse_version(v2)

        max_len = max(len(v1_parts), len(v2_parts))
        v1_parts += [0] * (max_len - len(v1_parts))
        v2_parts += [0] * (max_len - len(v2_parts))

        for p1, p2 in zip(v1_parts, v2_parts):
            if p1 > p2:
                return 1
            elif p1 < p2:
                return -1

        return 0

    def normalize_name_for_comparison(self, name):
        if not name:
            return ""

        name = str(name).lower().strip()

        replacements = {
            'postgres pro': 'postgrespro',
            'postgresql': 'postgres',
            'postgres pro enterprise': 'postgresproenterprise',
            'postgres pro certified': 'postgresprocertified',
            'microsoft sql server': 'mssql',
            'sql server': 'mssql',
            'windows server': 'windowsserver',
            'red hat': 'redhat',
            'oracle database': 'oracle',
            'mysql': 'mysql',
            'mariadb': 'mariadb',
            'mongodb': 'mongodb',
            'redis': 'redis',
            'apache': 'apache',
            'nginx': 'nginx',
            'docker': 'docker',
            'kubernetes': 'kubernetes',
            'vmware': 'vmware',
            'virtualbox': 'virtualbox',
            'python': 'python',
            'java': 'java',
            'node.js': 'nodejs',
            'php': 'php',
            '1с:предприятие': '1c',
            '1с:enterprise': '1c',
            'мойофис': 'myoffice',
            'р7-офис': 'r7office',
            'onlyoffice': 'onlyoffice',
            'astra linux': 'astra',
            'ред ос': 'redos',
            'alt linux': 'altlinux'
        }

        for old, new in replacements.items():
            name = name.replace(old, new)

        name = re.sub(r'[^\w\s-]', ' ', name)
        name = re.sub(r'\s+', ' ', name)

        return name.strip()

    def find_best_excel_match(self, product_name, gk_list=None):
        if not self.excel_versions_data:
            return None

        best_match = None
        best_score = 0

        product_norm = self.normalize_name_for_comparison(product_name)

        for norm_name, data in self.excel_versions_data.items():
            excel_name_norm = self.normalize_name_for_comparison(data['name'])
            score = 0

            if gk_list:
                for gk in gk_list:
                    if gk in data.get('gk', []):
                        score += 60

            if product_norm == excel_name_norm:
                score += 100
            elif product_norm in excel_name_norm or excel_name_norm in product_norm:
                score += 80
            else:
                product_words = set(product_norm.split())
                excel_words = set(excel_name_norm.split())
                common = product_words.intersection(excel_words)

                for word in common:
                    if len(word) > 2:
                        if word in ['postgres', 'oracle', 'mysql', 'windows', 'linux', 'mssql', '1c']:
                            score += 30
                        else:
                            score += 15

            if score > best_score:
                best_score = score
                best_match = data

        return best_match if best_score > 50 else None

    def scan_document_for_versions(self):
        self.versions_data = []
        lines = self.document_text.split('\n')

        for line_num, line in enumerate(lines):
            for category, patterns in self.software_patterns.items():
                for pattern, product_name in patterns:
                    try:
                        matches = re.finditer(pattern, line, re.IGNORECASE)
                        for match in matches:
                            if match.groups():
                                version = match.group(1).strip()
                                full_match = match.group(0).strip()

                                position = sum(len(l) + 1 for l in lines[:line_num]) + len(line[:match.start()])
                                page = self.find_page_for_position(position)

                                version = re.sub(r'[^\d\.]', '', version)

                                if version:
                                    self.versions_data.append({
                                        'category': category,
                                        'product': product_name,
                                        'version': version,
                                        'full_text': full_match,
                                        'line_number': line_num + 1,
                                        'page': page,
                                        'context': line.strip(),
                                        'start_pos': match.start(),
                                        'end_pos': match.end(),
                                        'line': line,
                                        'position': position,
                                        'parsed_version': self.parse_version(version)
                                    })
                    except Exception as e:
                        logger.error(f"Ошибка при поиске по паттерну {pattern}: {e}")

        self.versions_data = self.remove_duplicates(self.versions_data)
        self.versions_data.sort(key=lambda x: (x['category'], x['product'], x['version']))

    def remove_duplicates(self, data):
        seen = set()
        unique = []

        for item in data:
            key = (item['category'], item['product'], item['version'], item['line_number'])
            if key not in seen:
                seen.add(key)
                unique.append(item)

        return unique

    def find_page_for_position(self, position):
        if not self.page_info:
            return 1

        for page_num, (start, end) in enumerate(self.page_info, 1):
            if start <= position < end:
                return page_num

        return 1

    def display_versions(self):
        self.table.setRowCount(0)

        categories = set()
        versions = set()

        for i, version in enumerate(self.versions_data):
            self.table.insertRow(i)

            category_item = QTableWidgetItem(version['category'])
            category_item.setData(Qt.ItemDataRole.UserRole, version)
            self.table.setItem(i, 0, category_item)
            categories.add(version['category'])

            product_item = QTableWidgetItem(version['product'])
            self.table.setItem(i, 1, product_item)

            version_item = QTableWidgetItem(version['version'])
            version_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            versions.add(version['version'])
            self.table.setItem(i, 2, version_item)

            page_item = QTableWidgetItem(str(version['page']))
            page_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(i, 3, page_item)

            line_item = QTableWidgetItem(str(version['line_number']))
            line_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(i, 4, line_item)

            context = version['context'][:100] + "..." if len(version['context']) > 100 else version['context']
            context_item = QTableWidgetItem(context)
            context_item.setToolTip(version['context'])
            self.table.setItem(i, 5, context_item)

            btn = QPushButton("🔍 Перейти")
            btn.setProperty('row', i)
            btn.clicked.connect(lambda checked, r=i: self.go_to_version_position_by_row(r))
            btn.setMaximumWidth(80)
            self.table.setCellWidget(i, 6, btn)

        self.update_filters(categories, versions)
        self.update_stats()

    def update_filters(self, categories, versions):
        self.category_filter.clear()
        self.category_filter.addItem("Все категории")
        self.category_filter.addItems(sorted(categories))

        self.version_filter.clear()
        self.version_filter.addItem("Все версии")
        self.version_filter.addItems(sorted(versions))

    def filter_table(self):
        search_text = self.search_input.text().lower()
        category = self.category_filter.currentText()
        version_filter = self.version_filter.currentText()

        for row in range(self.table.rowCount()):
            show_row = True

            if search_text:
                row_text = ""
                for col in [0, 1, 2, 5]:
                    item = self.table.item(row, col)
                    if item:
                        row_text += item.text().lower() + " "

                if search_text not in row_text:
                    show_row = False

            if category != "Все категории":
                cat_item = self.table.item(row, 0)
                if cat_item and cat_item.text() != category:
                    show_row = False

            if version_filter != "Все версии":
                ver_item = self.table.item(row, 2)
                if ver_item and ver_item.text() != version_filter:
                    show_row = False

            self.table.setRowHidden(row, not show_row)

        self.update_stats()

    def update_stats(self):
        visible_rows = 0
        categories = set()
        products = set()

        for row in range(self.table.rowCount()):
            if not self.table.isRowHidden(row):
                visible_rows += 1
                cat_item = self.table.item(row, 0)
                prod_item = self.table.item(row, 1)

                if cat_item:
                    categories.add(cat_item.text())
                if prod_item:
                    products.add(prod_item.text())

        self.total_label.setText(f"Всего найдено версий: {visible_rows} (из {len(self.versions_data)})")
        self.categories_label.setText(f"Категорий: {len(categories)}")
        self.unique_label.setText(f"Уникальных продуктов: {len(products)}")

    def go_to_version_position(self, index):
        self.go_to_version_position_by_row(index.row())

    def go_to_version_position_by_row(self, row):
        item = self.table.item(row, 0)
        if not item:
            return

        version_data = item.data(Qt.ItemDataRole.UserRole)
        if not version_data:
            return

        self.show_context(version_data)

        parent = self.parent()
        while parent:
            if hasattr(parent, 'document_viewer') and parent.document_viewer:
                viewer = parent.document_viewer
                if version_data.get('page'):
                    viewer.go_to_page(version_data['page'])

                QTimer.singleShot(300, lambda: self.highlight_in_viewer(viewer, version_data))
                break
            parent = parent.parent()

    def highlight_in_viewer(self, viewer, version_data):
        if not viewer or not hasattr(viewer, 'text_edit'):
            return

        cursor = viewer.text_edit.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.Start)

        found = viewer.text_edit.find(version_data['full_text'])
        if found:
            format = QTextCharFormat()
            format.setBackground(QColor(255, 255, 0, 150))
            format.setFontWeight(QFont.Weight.Bold)
            cursor.mergeCharFormat(format)

    def show_context(self, version_data):
        context = f"""=== Найденная версия ===
Категория: {version_data['category']}
Продукт: {version_data['product']}
Версия: {version_data['version']}
Страница: {version_data['page']}
Строка: {version_data['line_number']}

Контекст:
{version_data['context']}

Полный текст:
{version_data['full_text']}
"""
        self.context_text.setPlainText(context)

    def load_saved_excel_data(self):
        if not self.auto_load_enabled:
            return

        data, metadata = self.data_manager.load_data()
        if data and metadata:
            self.excel_versions_data = data
            self.excel_file_path = metadata.get('file_path', '')
            self.excel_file_label.setText(f"📄 Загружено из сохранения: {metadata.get('file_name', '')}")
            self.excel_file_label.setStyleSheet("color: green;")

            self.display_excel_data()
            self.compare_btn.setEnabled(True)
            self.save_data_btn.setEnabled(True)
            self.add_record_btn.setEnabled(True)
            self.export_excel_btn.setEnabled(True)

            self.gk_excel_info.setText(f"📊 Данные Excel: {len(self.excel_versions_data)} записей")

            save_date = metadata.get('save_date', '')
            if save_date:
                try:
                    date = datetime.fromisoformat(save_date)
                    self.excel_stats.setText(
                        f"Загружено из сохранения от {date.strftime('%d.%m.%Y %H:%M')}. "
                        f"Записей: {metadata.get('record_count', 0)}"
                    )
                except:
                    pass

    def save_current_data(self):
        if not self.excel_versions_data:
            QMessageBox.warning(self, "Ошибка", "Нет данных для сохранения")
            return

        if self.data_manager.save_data(self.excel_versions_data, self.excel_file_path):
            QMessageBox.information(self, "Успех", "Данные успешно сохранены")
            self.excel_file_label.setText(f"📄 {os.path.basename(self.excel_file_path)} (сохранено)")
        else:
            QMessageBox.critical(self, "Ошибка", "Не удалось сохранить данные")

    def clear_saved_data(self):
        reply = QMessageBox.question(
            self, "Подтверждение",
            "Вы уверены, что хотите очистить все сохраненные данные?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            if self.data_manager.clear_data():
                self.excel_versions_data = {}
                self.excel_file_path = None
                self.excel_file_label.setText("Файл не загружен")
                self.excel_file_label.setStyleSheet("color: gray;")
                self.display_excel_data()
                self.compare_btn.setEnabled(False)
                self.save_data_btn.setEnabled(False)
                self.add_record_btn.setEnabled(False)
                self.export_excel_btn.setEnabled(False)
                self.comparison_table.setRowCount(0)
                self.comparison_stats.setText("Данные очищены")
                self.excel_stats.setText("Данные очищены")
                self.gk_excel_info.setText("📊 Данные Excel: 0 записей")
                QMessageBox.information(self, "Успех", "Сохраненные данные очищены")
            else:
                QMessageBox.critical(self, "Ошибка", "Не удалось очистить данные")

    def toggle_auto_load(self, state):
        self.auto_load_enabled = state == Qt.CheckState.Checked.value

    def load_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл с версиями", "",
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )

        if not file_path:
            return

        self.excel_file_path = file_path
        self.excel_file_label.setText(f"📄 {os.path.basename(file_path)}")
        self.excel_progress.setVisible(True)
        self.excel_progress.setValue(0)
        self.load_excel_btn.setEnabled(False)
        self.compare_btn.setEnabled(False)

        self.excel_loader = ExcelLoaderThread(file_path)
        self.excel_loader.progress.connect(self.update_excel_progress)
        self.excel_loader.finished.connect(self.on_excel_loaded)
        self.excel_loader.error.connect(self.on_excel_error)
        self.excel_loader.start()

    def update_excel_progress(self, value, message):
        self.excel_progress.setValue(value)
        self.excel_file_label.setText(message)

    def on_excel_loaded(self, result):
        self.excel_versions_data = result['data']
        self.excel_progress.setVisible(False)
        self.load_excel_btn.setEnabled(True)
        self.compare_btn.setEnabled(True)
        self.save_data_btn.setEnabled(True)
        self.add_record_btn.setEnabled(True)
        self.export_excel_btn.setEnabled(True)

        self.gk_excel_info.setText(f"📊 Данные Excel: {len(self.excel_versions_data)} записей")

        self.display_excel_data()

        stats_text = f"Загружено: {result['unique_names']} уникальных записей из {result['file_name']}"
        self.excel_stats.setText(stats_text)

        reply = QMessageBox.question(
            self, "Сохранение",
            "Сохранить загруженные данные для использования в следующий раз?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.save_current_data()

    def on_excel_error(self, error_msg):
        self.excel_progress.setVisible(False)
        self.load_excel_btn.setEnabled(True)
        self.excel_file_label.setText("Ошибка загрузки")
        QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить Excel файл:\n{error_msg}")

    def display_excel_data(self):
        self.excel_table.setRowCount(0)

        if not self.excel_versions_data:
            return

        items = list(self.excel_versions_data.items())
        items.sort(key=lambda x: x[1]['name'])

        for i, (norm_name, data) in enumerate(items):
            self.excel_table.insertRow(i)

            name_item = QTableWidgetItem(data['name'])
            name_item.setData(Qt.ItemDataRole.UserRole, norm_name)
            name_item.setData(Qt.ItemDataRole.UserRole + 1, data)
            self.excel_table.setItem(i, 0, name_item)

            if data.get('parsed_versions'):
                versions_text = ', '.join([v['original'] for v in data['parsed_versions'][:3]])
                if len(data['parsed_versions']) > 3:
                    versions_text += f"... (+{len(data['parsed_versions']) - 3})"
            else:
                versions_text = ', '.join(data['versions'][:3])
                if len(data['versions']) > 3:
                    versions_text += f"... (+{len(data['versions']) - 3})"

            self.excel_table.setItem(i, 1, QTableWidgetItem(versions_text))

            gk_text = ', '.join(data['gk'][:2])
            if len(data['gk']) > 2:
                gk_text += f"... (+{len(data['gk']) - 2})"
            self.excel_table.setItem(i, 2, QTableWidgetItem(gk_text))

            sheets_text = ', '.join(data['sheets'][:2])
            if len(data['sheets']) > 2:
                sheets_text += f"... (+{len(data['sheets']) - 2})"
            self.excel_table.setItem(i, 3, QTableWidgetItem(sheets_text))

            view_btn = QPushButton("👁 Просмотр")
            view_btn.setProperty('data', data)
            view_btn.clicked.connect(lambda checked, d=data: self.show_excel_item_details(d))
            self.excel_table.setCellWidget(i, 4, view_btn)

            edit_btn = QPushButton("✏ Ред.")
            edit_btn.setProperty('norm_name', norm_name)
            edit_btn.setProperty('data', data)
            edit_btn.clicked.connect(lambda checked, n=norm_name, d=data: self.edit_excel_record(n, d))
            self.excel_table.setCellWidget(i, 5, edit_btn)

        self.excel_table.resizeRowsToContents()

    def show_excel_item_details(self, data):
        from PyQt6.QtWidgets import QDialog, QVBoxLayout, QTextEdit

        dialog = QDialog(self)
        dialog.setWindowTitle(f"Детали: {data['name']}")
        dialog.resize(600, 400)

        layout = QVBoxLayout()

        text_edit = QTextEdit()
        text_edit.setReadOnly(True)

        details = f"<h3>{data['name']}</h3>"

        if data.get('parsed_versions'):
            details += "<p><b>Версии (от новых к старым):</b></p><ol>"
            for v in data['parsed_versions']:
                details += f"<li>{v['original']}</li>"
            details += "</ol>"
        else:
            details += "<p><b>Версии:</b></p><ul>"
            for v in data['versions']:
                details += f"<li>{v}</li>"
            details += "</ul>"

        if data['gk']:
            details += "<p><b>ГК:</b></p><ul>"
            for gk in data['gk']:
                details += f"<li>{gk}</li>"
            details += "</ul>"

        details += "<p><b>Листы:</b></p><ul>"
        for sheet in data['sheets']:
            details += f"<li>{sheet}</li>"
        details += "</ul>"

        text_edit.setHtml(details)
        layout.addWidget(text_edit)

        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(dialog.close)
        layout.addWidget(close_btn)

        dialog.setLayout(layout)
        dialog.exec()

    def edit_excel_record(self, norm_name, data):
        dialog = EditExcelDataDialog(self, data, norm_name)
        if dialog.exec():
            updated_data = dialog.get_updated_data()

            if norm_name in self.excel_versions_data:
                parsed_versions = []
                for v in updated_data['versions']:
                    parsed = self.parse_version(v)
                    parsed_versions.append({
                        'original': v,
                        'parsed': parsed
                    })

                self.excel_versions_data[norm_name] = {
                    'name': updated_data['name'],
                    'versions': updated_data['versions'],
                    'parsed_versions': parsed_versions,
                    'gk': updated_data['gk'],
                    'sheets': updated_data['sheets']
                }

                self.display_excel_data()

                reply = QMessageBox.question(
                    self, "Сохранение",
                    "Сохранить изменения?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )

                if reply == QMessageBox.StandardButton.Yes:
                    self.save_current_data()

    def add_excel_record(self):
        name, ok = QInputDialog.getText(self, "Новая запись", "Введите наименование:")
        if ok and name:
            norm_name = self.normalize_name_for_comparison(name)

            if norm_name in self.excel_versions_data:
                QMessageBox.warning(self, "Ошибка", "Запись с таким названием уже существует")
                return

            new_data = {
                'name': name,
                'versions': [],
                'parsed_versions': [],
                'gk': [],
                'sheets': ['Ручное добавление']
            }

            self.excel_versions_data[norm_name] = new_data
            self.display_excel_data()
            self.edit_excel_record(norm_name, new_data)

    def delete_excel_record(self, norm_name):
        reply = QMessageBox.question(
            self, "Подтверждение",
            f"Удалить запись '{self.excel_versions_data[norm_name]['name']}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            del self.excel_versions_data[norm_name]
            self.display_excel_data()

            reply = QMessageBox.question(
                self, "Сохранение",
                "Сохранить изменения?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )

            if reply == QMessageBox.StandardButton.Yes:
                self.save_current_data()

    def export_excel_data(self):
        from PyQt6.QtWidgets import QFileDialog
        import csv

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить данные Excel", "excel_data.csv", "CSV files (*.csv)"
        )

        if file_path:
            try:
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(["Наименование", "Версии", "ГК", "Листы"])

                    for norm_name, data in self.excel_versions_data.items():
                        versions = ', '.join(data['versions'])
                        gk = ', '.join(data['gk'])
                        sheets = ', '.join(data['sheets'])
                        writer.writerow([data['name'], versions, gk, sheets])

                QMessageBox.information(self, "Успех", f"Данные экспортированы в {file_path}")
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось экспортировать: {str(e)}")

    def show_excel_context_menu(self, position):
        item = self.excel_table.itemAt(position)
        if not item:
            return

        row = self.excel_table.rowAt(position.y())
        if row < 0:
            return

        name_item = self.excel_table.item(row, 0)
        if not name_item:
            return

        norm_name = name_item.data(Qt.ItemDataRole.UserRole)
        data = name_item.data(Qt.ItemDataRole.UserRole + 1)

        menu = QMenu()

        edit_action = QAction("✏ Редактировать", self)
        edit_action.triggered.connect(lambda: self.edit_excel_record(norm_name, data))

        delete_action = QAction("❌ Удалить", self)
        delete_action.triggered.connect(lambda: self.delete_excel_record(norm_name))

        copy_action = QAction("📋 Копировать название", self)
        copy_action.triggered.connect(lambda: QApplication.clipboard().setText(data['name']))

        menu.addAction(edit_action)
        menu.addAction(delete_action)
        menu.addSeparator()
        menu.addAction(copy_action)

        menu.exec(self.excel_table.viewport().mapToGlobal(position))

    def filter_excel_table(self):
        search_text = self.excel_search_input.text().lower()

        for row in range(self.excel_table.rowCount()):
            show_row = True

            if search_text:
                name_item = self.excel_table.item(row, 0)
                if name_item:
                    if search_text not in name_item.text().lower():
                        show_row = False

            self.excel_table.setRowHidden(row, not show_row)

    def compare_versions_with_excel(self):
        if not self.excel_versions_data:
            QMessageBox.warning(self, "Ошибка", "Сначала загрузите Excel файл")
            return

        self.comparison_results = []
        self.comparison_table.setRowCount(0)

        doc_gk_list = self.extract_gk_from_text(self.document_text)

        for doc_version in self.versions_data:
            doc_product = doc_version['product']
            doc_version_str = doc_version['version']
            doc_category = doc_version['category']
            doc_parsed = doc_version['parsed_version']

            relevant_gk = []
            for gk in doc_gk_list:
                if gk in doc_version.get('context', ''):
                    relevant_gk.append(gk)

            best_match = self.find_best_excel_match(doc_product, relevant_gk)

            if best_match:
                if best_match.get('parsed_versions'):
                    max_version = max(best_match['parsed_versions'],
                                      key=lambda x: x['parsed'])
                    excel_version = max_version['original']
                    excel_parsed = max_version['parsed']
                else:
                    excel_versions = best_match.get('versions', [])
                    if excel_versions:
                        excel_versions.sort(key=self.parse_version, reverse=True)
                        excel_version = excel_versions[0]
                        excel_parsed = self.parse_version(excel_version)
                    else:
                        excel_version = ''
                        excel_parsed = [0]

                comparison = self.compare_versions(doc_version_str, excel_version)
                is_compliant = comparison >= 0

                self.comparison_results.append({
                    'doc_product': doc_product,
                    'doc_category': doc_category,
                    'doc_version': doc_version_str,
                    'doc_parsed': doc_parsed,
                    'excel_name': best_match['name'],
                    'excel_version': excel_version,
                    'excel_parsed': excel_parsed,
                    'excel_versions_full': [v['original'] if isinstance(v, dict) else v
                                            for v in best_match.get('parsed_versions', best_match.get('versions', []))],
                    'excel_gk': best_match.get('gk', []),
                    'comparison': comparison,
                    'is_compliant': is_compliant,
                    'match_score': 100,
                    'source': 'excel',
                    'matched_gk': relevant_gk
                })
            else:
                self.comparison_results.append({
                    'doc_product': doc_product,
                    'doc_category': doc_category,
                    'doc_version': doc_version_str,
                    'doc_parsed': doc_parsed,
                    'excel_name': 'НЕ НАЙДЕНО',
                    'excel_version': '',
                    'excel_parsed': [0],
                    'excel_versions_full': [],
                    'excel_gk': [],
                    'comparison': None,
                    'is_compliant': False,
                    'match_score': 0,
                    'source': 'none',
                    'matched_gk': relevant_gk
                })

        self.display_comparison_results()

    def display_comparison_results(self):
        self.comparison_table.setRowCount(len(self.comparison_results))

        for i, result in enumerate(self.comparison_results):
            self.comparison_table.setItem(i, 0, QTableWidgetItem(result['doc_product']))

            self.comparison_table.setItem(i, 1, QTableWidgetItem(result['doc_category']))

            self.comparison_table.setItem(i, 2, QTableWidgetItem(result['doc_version']))

            if result['source'] == 'internal':
                excel_name_item = QTableWidgetItem("📚 " + result['excel_name'])
                excel_name_item.setForeground(QColor(0, 100, 200))
            else:
                excel_name_item = QTableWidgetItem(result['excel_name'])

            if result['excel_name'] == 'НЕ НАЙДЕНО':
                excel_name_item.setForeground(QColor(255, 0, 0))
            self.comparison_table.setItem(i, 3, excel_name_item)

            excel_version_item = QTableWidgetItem(result['excel_version'])
            if result['source'] == 'internal':
                excel_version_item.setForeground(QColor(0, 100, 200))
            elif result['excel_name'] != 'НЕ НАЙДЕНО' and result['source'] == 'excel':
                excel_version_item.setForeground(QColor(0, 0, 255))
            self.comparison_table.setItem(i, 4, excel_version_item)

            gk_text = ', '.join(result['excel_gk'][:2]) if result['excel_gk'] else ''
            if result.get('matched_gk'):
                gk_text += f" (найдено: {', '.join(result['matched_gk'][:1])})"
            self.comparison_table.setItem(i, 5, QTableWidgetItem(gk_text))

            if result['excel_name'] == 'НЕ НАЙДЕНО' or result['source'] == 'internal':
                comparison_text = "—"
            else:
                comp = result.get('comparison')
                if comp == 1:
                    comparison_text = f"{result['doc_version']} > {result['excel_version']}"
                elif comp == 0:
                    comparison_text = f"{result['doc_version']} = {result['excel_version']}"
                elif comp == -1:
                    comparison_text = f"{result['doc_version']} < {result['excel_version']}"
                else:
                    comparison_text = "—"

            comparison_item = QTableWidgetItem(comparison_text)
            self.comparison_table.setItem(i, 6, comparison_item)

            if result['excel_name'] == 'НЕ НАЙДЕНО':
                status = "❓ Не найдено"
                color = QColor(128, 128, 128)
            elif result['source'] == 'internal':
                status = "📚 Внутренние данные"
                color = QColor(0, 100, 200)
            elif result.get('is_compliant', False):
                status = "✅ Соответствует (>=)"
                color = QColor(0, 150, 0)
            else:
                status = f"❌ Не соответствует (<)"
                color = QColor(255, 0, 0)

            status_item = QTableWidgetItem(status)
            status_item.setForeground(color)
            self.comparison_table.setItem(i, 7, status_item)

            if result['excel_name'] == 'НЕ НАЙДЕНО':
                details = "Продукт не найден в Excel"
                if result.get('matched_gk'):
                    details += f"\nНайден ГК: {', '.join(result['matched_gk'])}"
            elif result['source'] == 'internal':
                details = f"Найдено во внутренних данных. Категория: {result['doc_category']}"
            else:
                if result.get('is_compliant'):
                    details = f"Версия {result['doc_version']} >= {result['excel_version']}"
                else:
                    details = f"Требуется версия >= {result['excel_version']}, найдена {result['doc_version']}"

                if len(result['excel_versions_full']) > 1:
                    details += f"\nДоступные версии в Excel: {', '.join(result['excel_versions_full'][:3])}"
                    if len(result['excel_versions_full']) > 3:
                        details += f" (+{len(result['excel_versions_full']) - 3})"

            details_item = QTableWidgetItem(details)
            self.comparison_table.setItem(i, 8, details_item)

            if result['source'] == 'excel' and result['excel_name'] != 'НЕ НАЙДЕНО':
                edit_btn = QPushButton("✏")
                edit_btn.setMaximumWidth(30)
                edit_btn.setProperty('excel_name', result['excel_name'])
                edit_btn.clicked.connect(lambda checked, name=result['excel_name']: self.edit_from_comparison(name))
                self.comparison_table.setCellWidget(i, 9, edit_btn)

        self.comparison_table.resizeRowsToContents()

        total = len(self.comparison_results)
        compliant = sum(
            1 for r in self.comparison_results if r.get('is_compliant', False) and r['source'] != 'internal')
        not_compliant = sum(
            1 for r in self.comparison_results if r['source'] == 'excel' and not r.get('is_compliant', False))
        not_found = sum(1 for r in self.comparison_results if r['source'] == 'none')

        stats_text = f"Всего: {total} | ✅ Соответствует: {compliant} | ❌ Не соответствует: {not_compliant} | ❓ Не найдено: {not_found}"
        self.comparison_stats.setText(stats_text)

    def edit_from_comparison(self, excel_name):
        for norm_name, data in self.excel_versions_data.items():
            if data['name'] == excel_name:
                self.edit_excel_record(norm_name, data)
                self.tab_widget.setCurrentIndex(2)
                break

    def filter_comparison_table(self):
        filter_text = self.compare_filter.currentText()

        for row in range(self.comparison_table.rowCount()):
            show_row = True
            status_item = self.comparison_table.item(row, 7)

            if status_item:
                status = status_item.text()

                if filter_text == "✅ Соответствует (>=)" and "✅" not in status:
                    show_row = False
                elif filter_text == "❌ Не соответствует" and ("❌" not in status or "✅" in status):
                    show_row = False
                elif filter_text == "❓ Не найдено в Excel" and "❓" not in status:
                    show_row = False

            self.comparison_table.setRowHidden(row, not show_row)

    def show_comparison_context_menu(self, position):
        menu = QMenu()

        copy_row_action = QAction("Копировать строку", self)
        copy_row_action.triggered.connect(lambda: self.copy_comparison_row())

        export_action = QAction("Экспорт результатов", self)
        export_action.triggered.connect(self.export_comparison_results)

        show_details_action = QAction("Показать все версии из Excel", self)
        show_details_action.triggered.connect(self.show_selected_excel_details)

        edit_action = QAction("✏ Редактировать запись в Excel", self)
        edit_action.triggered.connect(self.edit_selected_from_comparison)

        menu.addAction(copy_row_action)
        menu.addAction(show_details_action)
        menu.addAction(edit_action)
        menu.addSeparator()
        menu.addAction(export_action)

        menu.exec(self.comparison_table.viewport().mapToGlobal(position))

    def copy_comparison_row(self):
        current_row = self.comparison_table.currentRow()
        if current_row >= 0:
            text_parts = []
            for col in range(self.comparison_table.columnCount() - 1):
                item = self.comparison_table.item(current_row, col)
                if item:
                    text_parts.append(item.text())
                else:
                    text_parts.append("")

            QApplication.clipboard().setText('\t'.join(text_parts))
            self.show_temporary_message("Строка скопирована")

    def show_selected_excel_details(self):
        current_row = self.comparison_table.currentRow()
        if current_row < 0:
            return

        result = self.comparison_results[current_row]
        if result['excel_name'] in ['НЕ НАЙДЕНО', 'ВНУТРЕННИЕ ДАННЫЕ']:
            QMessageBox.information(self, "Информация", "Запись не найдена в Excel")
            return

        for norm_name, data in self.excel_versions_data.items():
            if data['name'] == result['excel_name']:
                self.show_excel_item_details(data)
                break

    def edit_selected_from_comparison(self):
        current_row = self.comparison_table.currentRow()
        if current_row < 0:
            return

        result = self.comparison_results[current_row]
        if result['excel_name'] in ['НЕ НАЙДЕНО', 'ВНУТРЕННИЕ ДАННЫЕ']:
            QMessageBox.information(self, "Информация", "Запись не найдена в Excel")
            return

        self.edit_from_comparison(result['excel_name'])

    def export_comparison_results(self):
        from PyQt6.QtWidgets import QFileDialog
        import csv

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить результаты сравнения", "comparison_results.csv", "CSV files (*.csv)"
        )

        if file_path:
            try:
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow([
                        "Продукт (документ)", "Категория", "Версия (документ)",
                        "Продукт (Excel)", "Требуемая версия", "ГК",
                        "Сравнение", "Статус", "Детали", "Источник"
                    ])

                    for row in range(self.comparison_table.rowCount()):
                        if not self.comparison_table.isRowHidden(row):
                            row_data = []
                            for col in range(self.comparison_table.columnCount() - 1):
                                item = self.comparison_table.item(row, col)
                                if item:
                                    row_data.append(item.text())
                                else:
                                    row_data.append("")
                            result = self.comparison_results[row]
                            row_data.append(result.get('source', 'unknown'))
                            writer.writerow(row_data)

                QMessageBox.information(self, "Успех", f"Результаты экспортированы в {file_path}")
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось экспортировать: {str(e)}")

    def show_context_menu(self, position):
        menu = QMenu()

        copy_version_action = QAction("📋 Копировать версию", self)
        copy_version_action.triggered.connect(self.copy_selected_version)

        copy_product_action = QAction("📋 Копировать продукт", self)
        copy_product_action.triggered.connect(self.copy_selected_product)

        copy_all_action = QAction("📋 Копировать строку", self)
        copy_all_action.triggered.connect(self.copy_selected_row)

        menu.addAction(copy_version_action)
        menu.addAction(copy_product_action)
        menu.addSeparator()
        menu.addAction(copy_all_action)

        menu.exec(self.table.viewport().mapToGlobal(position))

    def copy_selected_version(self):
        current_row = self.table.currentRow()
        if current_row >= 0:
            version_item = self.table.item(current_row, 2)
            if version_item:
                QApplication.clipboard().setText(version_item.text())
                self.show_temporary_message(f"Скопировано: {version_item.text()}")

    def copy_selected_product(self):
        current_row = self.table.currentRow()
        if current_row >= 0:
            product_item = self.table.item(current_row, 1)
            if product_item:
                QApplication.clipboard().setText(product_item.text())
                self.show_temporary_message(f"Скопировано: {product_item.text()}")

    def copy_selected_row(self):
        current_row = self.table.currentRow()
        if current_row >= 0:
            text_parts = []
            for col in range(self.table.columnCount() - 1):
                item = self.table.item(current_row, col)
                if item:
                    text_parts.append(item.text())

            QApplication.clipboard().setText('\t'.join(text_parts))
            self.show_temporary_message("Строка скопирована")

    def copy_all_to_clipboard(self):
        text = "Категория\tПродукт\tВерсия\tСтраница\tСтрока\tКонтекст\n"

        for row in range(self.table.rowCount()):
            if not self.table.isRowHidden(row):
                row_text = []
                for col in range(6):
                    item = self.table.item(row, col)
                    if item:
                        row_text.append(item.text())
                    else:
                        row_text.append("")
                text += '\t'.join(row_text) + '\n'

        QApplication.clipboard().setText(text)
        self.show_temporary_message(f"Скопировано {self.get_visible_row_count()} строк")

    def get_visible_row_count(self):
        count = 0
        for row in range(self.table.rowCount()):
            if not self.table.isRowHidden(row):
                count += 1
        return count

    def show_temporary_message(self, message):
        QToolTip.showText(self.mapToGlobal(self.rect().center()), message)

    def export_to_csv(self):
        from PyQt6.QtWidgets import QFileDialog
        import csv

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить как CSV", "versions.csv", "CSV files (*.csv)"
        )

        if file_path:
            try:
                with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(["Категория", "Продукт", "Версия", "Страница", "Строка", "Контекст"])

                    for row in range(self.table.rowCount()):
                        if not self.table.isRowHidden(row):
                            row_data = []
                            for col in range(6):
                                item = self.table.item(row, col)
                                if item:
                                    row_data.append(item.text())
                                else:
                                    row_data.append("")
                            writer.writerow(row_data)

                QMessageBox.information(self, "Успех", f"Данные экспортированы в {file_path}")
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось экспортировать: {str(e)}")

    def export_to_html(self):
        from PyQt6.QtWidgets import QFileDialog

        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить как HTML", "versions.html", "HTML files (*.html)"
        )

        if file_path:
            try:
                html = """
                <html>
                <head>
                    <title>Версии БПО/СПО</title>
                    <style>
                        body { font-family: Arial, sans-serif; margin: 20px; }
                        h1 { color: #2c3e50; }
                        table { border-collapse: collapse; width: 100%%; }
                        th { background-color: #3498db; color: white; padding: 10px; }
                        td { border: 1px solid #ddd; padding: 8px; }
                        tr:nth-child(even) { background-color: #f2f2f2; }
                        tr:hover { background-color: #e8f4f8; }
                    </style>
                </head>
                <body>
                    <h1>Версии БПО/СПО в документе</h1>
                    <table>
                        <tr>
                            <th>Категория</th>
                            <th>Продукт</th>
                            <th>Версия</th>
                            <th>Страница</th>
                            <th>Строка</th>
                            <th>Контекст</th>
                        </tr>
                """

                for row in range(self.table.rowCount()):
                    if not self.table.isRowHidden(row):
                        html += "<tr>"
                        for col in range(6):
                            item = self.table.item(row, col)
                            if item:
                                html += f"<td>{item.text()}</td>"
                            else:
                                html += "<td></td>"
                        html += "</tr>"

                html += """
                    </table>
                </body>
                </html>
                """

                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(html)

                QMessageBox.information(self, "Успех", f"Данные экспортированы в {file_path}")
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось экспортировать: {str(e)}")

    def refresh_versions(self):
        self.scan_document_for_versions()
        self.display_versions()
        self.show_temporary_message(f"Найдено {len(self.versions_data)} версий")