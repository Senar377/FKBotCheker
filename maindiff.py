import sys
import yaml
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime
import time
import html

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QTextEdit, QFileDialog, QListWidget,
    QListWidgetItem, QCheckBox, QProgressBar, QTableWidget,
    QTableWidgetItem, QTabWidget, QGroupBox, QMessageBox,
    QSplitter, QHeaderView, QToolBar, QStatusBar, QTextBrowser,
    QDialog, QTextEdit, QScrollArea, QMenu, QLineEdit,
    QComboBox, QDialogButtonBox
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer, QSettings
from PyQt6.QtGui import QFont, QIcon, QAction, QColor, QPalette, QTextCursor, QTextCharFormat, QSyntaxHighlighter, \
    QTextDocument
import rapidfuzz
from rapidfuzz import fuzz, process
import difflib


class DocumentHighlighter(QSyntaxHighlighter):
    """Подсветка найденных ошибок в тексте документа"""

    def __init__(self, document, matches):
        super().__init__(document)
        self.matches = matches

    def highlightBlock(self, text):
        if not self.matches:
            return

        for start, end, term in self.matches:
            if start < len(text) and end <= len(text):
                # Проверяем, что это целое слово
                if start > 0 and text[start - 1].isalnum():
                    continue
                if end < len(text) and text[end].isalnum():
                    continue

                format = QTextCharFormat()
                format.setBackground(QColor(255, 0, 0, 80))  # Красный фон
                format.setForeground(QColor(255, 255, 255))  # Белый текст
                format.setFontWeight(QFont.Weight.Bold)

                self.setFormat(start, end - start, format)


class DocumentChecker:
    """Класс для проверки документов по конфигурации"""

    def __init__(self, config_path: str = None):
        self.config = self.load_config(config_path) if config_path else {}
        self.results = []

    def load_config(self, config_path: str) -> Dict:
        """Загрузка конфигурации из YAML файла"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f) or {}
        except Exception as e:
            print(f"Ошибка загрузки конфигурации: {e}")
            return {}

    def save_config(self, config_path: str, config: Dict) -> bool:
        """Сохранение конфигурации в YAML файл"""
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                yaml.dump(config, f, allow_unicode=True, default_flow_style=False)
            return True
        except Exception as e:
            print(f"Ошибка сохранения конфигурации: {e}")
            return False

    def normalize_text(self, text: str) -> str:
        """Нормализация текста (регистр, пробелы)"""
        if not text:
            return ""
        # Удаление лишних пробелов и приведение к нижнему регистру
        text = re.sub(r'\s+', ' ', text.strip().lower())
        return text

    def exact_search(self, text: str, search_terms: List[str]) -> List[Tuple[int, int, str]]:
        """Точный поиск целых слов в тексте"""
        matches = []

        for term in search_terms:
            if not term:
                continue

            # Используем регулярное выражение для поиска целых слов
            pattern = r'\b' + re.escape(term.lower()) + r'\b'

            for match in re.finditer(pattern, text.lower()):
                matches.append((match.start(), match.end(), term))

        return matches

    def fuzzy_search_all(self, text: str, search_text: str, threshold: float = 70.0) -> List[
        Tuple[int, int, float, str]]:
        """Нечеткий поиск всех вхождений с использованием RapidFuzz"""
        if not text or not search_text:
            return []

        normalized_text = self.normalize_text(text)
        normalized_search = self.normalize_text(search_text)

        # Разбиваем текст на слова
        words = normalized_text.split()
        matches = []

        # Ищем похожие последовательности слов
        for i in range(len(words)):
            # Проверяем последовательности от 1 до 5 слов
            for length in range(1, min(6, len(words) - i + 1)):
                substring = ' '.join(words[i:i + length])
                if len(substring) < len(normalized_search) * 0.5:
                    continue

                score = fuzz.partial_ratio(substring, normalized_search)
                if score >= threshold:
                    # Находим позицию в исходном тексте
                    pos = normalized_text.find(substring)
                    if pos != -1:
                        matches.append((pos, pos + len(substring), score, substring))

        # Сортируем по схожести и удаляем дубликаты
        matches.sort(key=lambda x: x[2], reverse=True)
        unique_matches = []
        seen_positions = set()

        for match in matches:
            pos_key = (match[0], match[1])
            if pos_key not in seen_positions:
                seen_positions.add(pos_key)
                unique_matches.append(match)

        return unique_matches[:10]

    def fuzzy_search_best(self, text: str, search_text: str) -> float:
        """Лучший результат нечеткого поиска"""
        if not text or not search_text:
            return 0.0

        normalized_text = self.normalize_text(text)
        normalized_search = self.normalize_text(search_text)

        # Используем частичное соотношение
        score = fuzz.partial_ratio(normalized_text, normalized_search)
        return score

    def extract_tables(self, document_text: str) -> List[str]:
        """Извлечение таблиц из текста документа"""
        tables = []
        lines = document_text.split('\n')

        current_table = []
        in_table = False

        for line in lines:
            # Эвристика для таблиц
            if (('\t' in line or '|' in line or
                 (line.count('  ') > 3 and any(c.isdigit() for c in line))) and
                    len(line.strip()) > 10):

                if not in_table:
                    in_table = True
                current_table.append(line)
            elif in_table:
                if current_table:
                    tables.append('\n'.join(current_table))
                    current_table = []
                in_table = False

        if current_table:
            tables.append('\n'.join(current_table))

        return tables

    def extract_paragraphs_after_tables(self, document_text: str) -> List[str]:
        """Извлечение абзацев после таблиц"""
        paragraphs = []
        lines = document_text.split('\n')

        for i, line in enumerate(lines):
            if (('\t' in line or '|' in line or line.count('  ') > 3) and
                    i < len(lines) - 1):

                next_line = lines[i + 1].strip()
                if (next_line and
                        '\t' not in next_line and
                        '|' not in next_line and
                        not re.match(r'^\s*$', next_line)):
                    paragraphs.append(next_line)

        return paragraphs

    def check_subcheck(self, subcheck: Dict, document_text: str) -> Dict[str, Any]:
        """Выполнение одной подпроверки"""
        check_type = subcheck.get('type', '')
        name = subcheck.get('name', 'Неизвестная проверка')

        result = {
            'name': name,
            'type': check_type,
            'passed': False,
            'needs_verification': False,
            'matches': [],
            'score': 0.0,
            'message': '',
            'details': '',
            'found_text': '',
            'search_terms': subcheck.get('aliases', []) if 'aliases' in subcheck else [],
            'search_text': subcheck.get('text', '') if 'text' in subcheck else '',
            'position': ''
        }

        try:
            if check_type == 'no_text_present':
                aliases = subcheck.get('aliases', [])
                matches = self.exact_search(document_text, aliases)
                result['passed'] = len(matches) == 0
                result['matches'] = matches

                if matches:
                    # Определяем позицию первой ошибки
                    first_match = matches[0]
                    lines_before = document_text[:first_match[0]].count('\n')
                    result['position'] = f"Строка {lines_before + 1}"
                    result['found_text'] = f"Найдено: {first_match[2]}"
                    result['message'] = f"Найдено вхождений: {len(matches)}"
                else:
                    result['message'] = "Упоминания не найдено"

                result['details'] = f"Искали: {', '.join(aliases)}"

            elif check_type == 'text_present':
                aliases = subcheck.get('aliases', [])
                matches = self.exact_search(document_text, aliases)
                result['passed'] = len(matches) > 0
                result['matches'] = matches

                if matches:
                    first_match = matches[0]
                    lines_before = document_text[:first_match[0]].count('\n')
                    result['position'] = f"Строка {lines_before + 1}"
                    result['found_text'] = f"Найдено: {first_match[2]}"
                    result['message'] = f"Найдено вхождений: {len(matches)}"
                else:
                    result['message'] = "Текст не найден"

                result['details'] = f"Искали: {', '.join(aliases)}"

            elif check_type == 'text_present_without':
                aliases = subcheck.get('aliases', [])
                without_aliases = subcheck.get('without_aliases', [])

                positive_matches = self.exact_search(document_text, aliases)
                negative_matches = self.exact_search(document_text, without_aliases)

                result['passed'] = len(positive_matches) > 0 and len(negative_matches) == 0
                result['matches'] = positive_matches + negative_matches
                result['found_text'] = f"Найдено: {len(positive_matches)}, исключающих: {len(negative_matches)}"
                result['message'] = f"Основных вхождений: {len(positive_matches)}, исключающих: {len(negative_matches)}"
                result['details'] = f"Искали: {', '.join(aliases)}, исключали: {', '.join(without_aliases)}"

            elif check_type == 'fuzzy_text_present':
                text = subcheck.get('text', '')
                threshold = float(subcheck.get('threshold', 70.0))
                trust_threshold = float(subcheck.get('trust_threshold', 85.0))

                score = self.fuzzy_search_best(document_text, text)
                result['score'] = score
                result['passed'] = score >= trust_threshold
                result['needs_verification'] = threshold <= score < trust_threshold
                result['message'] = f"Схожесть: {score:.1f}%"
                result['details'] = f"Искали: '{text}' (пороги: {threshold}/{trust_threshold}%)"

                if score >= threshold:
                    matches = self.fuzzy_search_all(document_text, text, threshold)
                    if matches:
                        result['matches'] = [(m[0], m[1], f"{m[3]} ({m[2]:.1f}%)") for m in matches]
                        result['found_text'] = f"Лучшее совпадение: {matches[0][3][:50]}..."

            elif check_type == 'no_fuzzy_text_present':
                text = subcheck.get('text', '')
                threshold = float(subcheck.get('threshold', 70.0))
                trust_threshold = float(subcheck.get('trust_threshold', 85.0))

                score = self.fuzzy_search_best(document_text, text)
                result['score'] = score
                result['passed'] = score < threshold
                result['needs_verification'] = threshold <= score < trust_threshold
                result['message'] = f"Схожесть: {score:.1f}%"
                result['details'] = f"Искали: '{text}' (пороги: {threshold}/{trust_threshold}%)"

                if score >= threshold:
                    matches = self.fuzzy_search_all(document_text, text, threshold)
                    if matches:
                        result['matches'] = [(m[0], m[1], f"{m[3]} ({m[2]:.1f}%)") for m in matches]
                        result['found_text'] = f"Найдено похожее: {matches[0][3][:50]}..."

            elif check_type == 'text_present_in_any_table':
                aliases = subcheck.get('aliases', [])
                tables = self.extract_tables(document_text)

                table_matches = []
                for table_idx, table in enumerate(tables):
                    matches = self.exact_search(table, aliases)
                    if matches:
                        for match in matches:
                            table_matches.append((match[0], match[1], f"Таблица {table_idx + 1}: {match[2]}"))

                result['passed'] = len(table_matches) > 0
                result['matches'] = table_matches

                if table_matches:
                    result['found_text'] = f"Найдено в {len(set([m[2].split(':')[0] for m in table_matches]))} таблицах"
                    result['position'] = table_matches[0][2].split(':')[0]

                result['message'] = f"Найдено в таблицах: {len(table_matches)}"
                result['details'] = f"Искали в таблицах: {', '.join(aliases)}"

            elif check_type == 'fuzzy_text_present_after_any_table':
                text = subcheck.get('text', '')
                threshold = float(subcheck.get('threshold', 70.0))
                trust_threshold = float(subcheck.get('trust_threshold', 85.0))

                paragraphs = self.extract_paragraphs_after_tables(document_text)
                best_score = 0.0
                best_match = ""
                best_position = ""

                for para_idx, paragraph in enumerate(paragraphs):
                    score = self.fuzzy_search_best(paragraph, text)
                    if score > best_score:
                        best_score = score
                        best_match = paragraph[:100] + "..." if len(paragraph) > 100 else paragraph
                        best_position = f"После таблицы {para_idx + 1}"

                result['score'] = best_score
                result['passed'] = best_score >= trust_threshold
                result['needs_verification'] = threshold <= best_score < trust_threshold
                result['message'] = f"Схожесть: {best_score:.1f}%"
                result['details'] = f"Искали после таблиц: '{text}'"
                result['position'] = best_position

                if best_match:
                    result['found_text'] = f"Абзац: {best_match}"

        except Exception as e:
            result['message'] = f"Ошибка проверки: {str(e)}"
            result['passed'] = False

        return result

    def check_document(self, document_text: str, selected_checks: List[str] = None) -> List[Dict]:
        """Основная функция проверки документа"""
        if not self.config.get('checks'):
            return []

        results = []
        for check_group in self.config.get('checks', []):
            group_name = check_group.get('group', '')
            subchecks = check_group.get('subchecks', [])

            for subcheck in subchecks:
                check_name = subcheck.get('name', '')
                if selected_checks and check_name not in selected_checks:
                    continue

                result = self.check_subcheck(subcheck, document_text)
                result['group'] = group_name
                results.append(result)

        return results


class DOCXParser:
    """Класс для парсинга DOCX файлов"""

    @staticmethod
    def extract_text_from_docx(docx_path: str) -> str:
        """Извлекает текст из DOCX файла"""
        try:
            with zipfile.ZipFile(docx_path) as docx:
                xml_content = docx.read('word/document.xml')
                root = ET.fromstring(xml_content)

                namespaces = {
                    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
                }

                paragraphs = root.findall('.//w:p', namespaces)
                text_lines = []

                for para in paragraphs:
                    texts = para.findall('.//w:t', namespaces)
                    para_text = ''.join([text.text if text.text else '' for text in texts])

                    if para_text.strip():
                        text_lines.append(para_text)

                full_text = '\n'.join(text_lines)
                return full_text

        except Exception as e:
            raise Exception(f"Ошибка чтения DOCX файла: {str(e)}")


class CheckWorker(QThread):
    """Поток для выполнения проверки"""
    progress = pyqtSignal(int, str)  # прогресс и текущая проверка
    finished = pyqtSignal(list)

    def __init__(self, checker: DocumentChecker, document_text: str, selected_checks: List[str], config: Dict):
        super().__init__()
        self.checker = checker
        self.document_text = document_text
        self.selected_checks = selected_checks
        self.config = config

    def run(self):
        results = []

        # Создаем временный чекер с текущей конфигурацией
        temp_checker = DocumentChecker()
        temp_checker.config = self.config

        # Находим все подпроверки
        all_subchecks = []
        for check_group in self.config.get('checks', []):
            for subcheck in check_group.get('subchecks', []):
                check_name = subcheck.get('name', '')
                if not self.selected_checks or check_name in self.selected_checks:
                    all_subchecks.append((check_group.get('group', ''), subcheck))

        total_checks = len(all_subchecks)

        for i, (group_name, subcheck) in enumerate(all_subchecks):
            check_name = subcheck.get('name', 'Неизвестная проверка')

            # Отправляем информацию о текущей проверке
            self.progress.emit(int((i + 1) / total_checks * 100), f"Проверка: {check_name}")

            # Выполняем проверку
            result = temp_checker.check_subcheck(subcheck, self.document_text)
            result['group'] = group_name
            results.append(result)

            # Небольшая задержка для плавности прогресса
            self.msleep(50)

        self.finished.emit(results)


class DocumentViewer(QDialog):
    """Окно просмотра документа с подсветкой ошибок"""

    def __init__(self, parent=None, document_text="", results=None):
        super().__init__(parent)
        self.document_text = document_text
        self.results = results or []
        self.setWindowTitle("Просмотр документа с подсветкой ошибок")
        self.resize(1000, 700)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Панель управления
        control_panel = QHBoxLayout()

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Поиск в документе...")
        self.search_input.textChanged.connect(self.search_text)

        self.show_all_btn = QPushButton("Показать все ошибки")
        self.show_all_btn.clicked.connect(self.show_all_errors)

        self.clear_btn = QPushButton("Очистить подсветку")
        self.clear_btn.clicked.connect(self.clear_highlights)

        self.next_error_btn = QPushButton("Следующая ошибка →")
        self.next_error_btn.clicked.connect(self.next_error)

        self.prev_error_btn = QPushButton("← Предыдущая ошибка")
        self.prev_error_btn.clicked.connect(self.prev_error)

        control_panel.addWidget(QLabel("Поиск:"))
        control_panel.addWidget(self.search_input)
        control_panel.addWidget(self.show_all_btn)
        control_panel.addWidget(self.clear_btn)
        control_panel.addStretch()
        control_panel.addWidget(self.prev_error_btn)
        control_panel.addWidget(self.next_error_btn)

        layout.addLayout(control_panel)

        # Текстовое поле с подсветкой
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setFont(QFont("Consolas", 10))
        self.text_edit.setPlainText(self.document_text[:20000] + ("..." if len(self.document_text) > 20000 else ""))

        # Собираем все совпадения для подсветки
        self.all_matches = []
        for result in self.results:
            if result.get('matches'):
                self.all_matches.extend(result['matches'])

        self.current_match_index = 0
        self.search_matches = []
        self.current_search_index = 0

        layout.addWidget(self.text_edit)

        # Статистика
        stats_label = QLabel(f"Найдено проблем: {len(self.all_matches)} | Всего символов: {len(self.document_text)}")
        stats_label.setStyleSheet("color: #ffffff; font-size: 12px;")
        layout.addWidget(stats_label)

        # Кнопка закрытия
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        close_btn.setMinimumHeight(35)
        layout.addWidget(close_btn)

        self.setLayout(layout)

        # Сразу показываем все ошибки
        self.show_all_errors()

    def show_all_errors(self):
        """Подсветить все найденные ошибки"""
        if not self.all_matches:
            return

        # Очищаем предыдущую подсветку
        self.clear_highlights()

        # Создаем формат для подсветки
        cursor = self.text_edit.textCursor()
        format = QTextCharFormat()
        format.setBackground(QColor(255, 0, 0, 80))  # Красный фон
        format.setForeground(QColor(255, 255, 255))  # Белый текст
        format.setFontWeight(QFont.Weight.Bold)

        # Применяем подсветку ко всем совпадениям
        for start, end, term in self.all_matches:
            cursor.setPosition(start)
            cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
            cursor.mergeCharFormat(format)

        # Возвращаем курсор в начало
        cursor.setPosition(0)
        self.text_edit.setTextCursor(cursor)

    def clear_highlights(self):
        """Очистить подсветку"""
        cursor = self.text_edit.textCursor()
        format = QTextCharFormat()
        format.setBackground(QColor(0, 0, 0, 0))
        format.setForeground(QColor(255, 255, 255))

        cursor.select(QTextCursor.SelectionType.Document)
        cursor.mergeCharFormat(format)

        cursor.setPosition(0)
        self.text_edit.setTextCursor(cursor)

    def search_text(self):
        """Поиск текста в документе"""
        search_term = self.search_input.text()
        if not search_term:
            self.search_matches = []
            return

        # Ищем все вхождения
        text = self.text_edit.toPlainText().lower()
        search_term_lower = search_term.lower()
        self.search_matches = []

        pos = text.find(search_term_lower)
        while pos != -1:
            self.search_matches.append(pos)
            pos = text.find(search_term_lower, pos + 1)

        if self.search_matches:
            self.highlight_search_results()

    def highlight_search_results(self):
        """Подсветить результаты поиска"""
        self.clear_highlights()

        if not self.search_matches:
            return

        cursor = self.text_edit.textCursor()
        format = QTextCharFormat()
        format.setBackground(QColor(0, 100, 255, 100))  # Синий для поиска
        format.setForeground(QColor(255, 255, 255))

        for pos in self.search_matches:
            cursor.setPosition(pos)
            cursor.setPosition(pos + len(self.search_input.text()), QTextCursor.MoveMode.KeepAnchor)
            cursor.mergeCharFormat(format)

        # Переходим к первому результату
        if self.search_matches:
            cursor.setPosition(self.search_matches[0])
            self.text_edit.setTextCursor(cursor)
            self.text_edit.ensureCursorVisible()

    def next_error(self):
        """Перейти к следующей ошибке"""
        if not self.all_matches:
            return

        self.current_match_index = (self.current_match_index + 1) % len(self.all_matches)
        self.go_to_match(self.current_match_index)

    def prev_error(self):
        """Перейти к предыдущей ошибке"""
        if not self.all_matches:
            return

        self.current_match_index = (self.current_match_index - 1) % len(self.all_matches)
        self.go_to_match(self.current_match_index)

    def go_to_match(self, index):
        """Перейти к указанному совпадению"""
        if index < 0 or index >= len(self.all_matches):
            return

        start, end, term = self.all_matches[index]

        # Прокручиваем к позиции
        cursor = self.text_edit.textCursor()
        cursor.setPosition(start)
        self.text_edit.setTextCursor(cursor)
        self.text_edit.ensureCursorVisible()


class SettingsDialog(QDialog):
    """Диалог настроек"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setWindowTitle("Настройки")
        self.resize(500, 300)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # Выбор темы
        theme_label = QLabel("Цветовая тема:")
        theme_label.setStyleSheet("font-weight: bold;")

        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Темная", "Светлая", "Смешанная"])
        # Определяем текущую тему
        if self.parent.theme_mode == "dark":
            self.theme_combo.setCurrentText("Темная")
        elif self.parent.theme_mode == "light":
            self.theme_combo.setCurrentText("Светлая")
        else:
            self.theme_combo.setCurrentText("Смешанная")

        # Пороги для нечеткого поиска
        thresholds_label = QLabel("Пороги нечеткого поиска:")
        thresholds_label.setStyleSheet("font-weight: bold;")

        threshold_layout = QHBoxLayout()
        threshold_layout.addWidget(QLabel("Порог проверки:"))
        self.threshold_input = QLineEdit(str(self.parent.fuzzy_threshold))
        self.threshold_input.setMaximumWidth(80)
        threshold_layout.addWidget(self.threshold_input)

        threshold_layout.addWidget(QLabel("Порог доверия:"))
        self.trust_threshold_input = QLineEdit(str(self.parent.fuzzy_trust_threshold))
        self.trust_threshold_input.setMaximumWidth(80)
        threshold_layout.addWidget(self.trust_threshold_input)
        threshold_layout.addStretch()

        # Настройки интерфейса
        interface_label = QLabel("Настройки интерфейса:")
        interface_label.setStyleSheet("font-weight: bold;")

        self.auto_resize_check = QCheckBox("Автоматически изменять размер столбцов")
        self.auto_resize_check.setChecked(self.parent.auto_resize_columns)

        self.show_line_numbers_check = QCheckBox("Показывать номера строк в просмотре")
        self.show_line_numbers_check.setChecked(self.parent.show_line_numbers)

        # Кнопки
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        layout.addWidget(theme_label)
        layout.addWidget(self.theme_combo)
        layout.addSpacing(10)
        layout.addWidget(thresholds_label)
        layout.addLayout(threshold_layout)
        layout.addSpacing(10)
        layout.addWidget(interface_label)
        layout.addWidget(self.auto_resize_check)
        layout.addWidget(self.show_line_numbers_check)
        layout.addStretch()
        layout.addWidget(button_box)

        self.setLayout(layout)

    def accept(self):
        """Применить настройки"""
        try:
            threshold = float(self.threshold_input.text())
            trust_threshold = float(self.trust_threshold_input.text())

            if threshold >= trust_threshold:
                QMessageBox.warning(self, "Ошибка", "Порог проверки должен быть меньше порога доверия")
                return

            theme_text = self.theme_combo.currentText()
            if theme_text == "Темная":
                self.parent.theme_mode = "dark"
                self.parent.dark_theme = True
            elif theme_text == "Светлая":
                self.parent.theme_mode = "light"
                self.parent.dark_theme = False
            else:
                self.parent.theme_mode = "mixed"
                self.parent.dark_theme = False

            self.parent.fuzzy_threshold = threshold
            self.parent.fuzzy_trust_threshold = trust_threshold
            self.parent.auto_resize_columns = self.auto_resize_check.isChecked()
            self.parent.show_line_numbers = self.show_line_numbers_check.isChecked()

            # Сохраняем настройки
            self.parent.save_settings()

            # Применяем тему
            self.parent.apply_theme()

            super().accept()

        except ValueError:
            QMessageBox.warning(self, "Ошибка", "Пороги должны быть числами")


class ConfigEditor(QDialog):
    """Редактор конфигурации"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setWindowTitle("Редактирование конфигурации")
        self.resize(900, 700)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        title_label = QLabel("Редактирование конфигурации проверок")
        title_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        layout.addWidget(title_label)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(10)

        self.load_btn = QPushButton("Загрузить из файла")
        self.load_btn.clicked.connect(self.load_config)

        self.save_btn = QPushButton("Сохранить в файл")
        self.save_btn.clicked.connect(self.save_config)

        self.restore_btn = QPushButton("Восстановить по умолчанию")
        self.restore_btn.clicked.connect(self.restore_config)

        self.cancel_btn = QPushButton("Отмена")
        self.cancel_btn.clicked.connect(self.reject)

        self.apply_btn = QPushButton("Применить изменения")
        self.apply_btn.clicked.connect(self.apply_config)
        self.apply_btn.setStyleSheet("background-color: #0066cc; color: white; font-weight: bold;")

        button_layout.addWidget(self.load_btn)
        button_layout.addWidget(self.save_btn)
        button_layout.addWidget(self.restore_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.cancel_btn)
        button_layout.addWidget(self.apply_btn)

        layout.addLayout(button_layout)

        editor_label = QLabel("YAML конфигурация:")
        editor_label.setStyleSheet("font-weight: bold;")
        layout.addWidget(editor_label)

        self.text_edit = QTextEdit()
        self.text_edit.setFont(QFont("Consolas", 10))
        self.text_edit.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #d4d4d4;
                border: 2px solid #555555;
                border-radius: 6px;
                padding: 10px;
            }
        """)

        layout.addWidget(self.text_edit)

        self.setLayout(layout)

        if self.parent and hasattr(self.parent, 'checker'):
            import io
            stream = io.StringIO()
            yaml.dump(self.parent.checker.config, stream, allow_unicode=True, default_flow_style=False)
            self.text_edit.setPlainText(stream.getvalue())

    def load_config(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Загрузить конфигурацию", "", "YAML files (*.yaml *.yml);;All files (*.*)"
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                    self.text_edit.setPlainText(content)
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось загрузить файл:\n{str(e)}")

    def save_config(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить конфигурацию", "config.yaml", "YAML files (*.yaml *.yml);;All files (*.*)"
        )
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.text_edit.toPlainText())
                QMessageBox.information(self, "Успех", f"Конфигурация сохранена в:\n{file_path}")
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось сохранить файл:\n{str(e)}")

    def restore_config(self):
        reply = QMessageBox.question(
            self, "Подтверждение",
            "Восстановить конфигурацию по умолчанию?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            default_config = """checks:
  - group: "Импортозамещение"
    subchecks:
      - name: "Oracle"
        type: "no_text_present"
        aliases: ["Oracle", "Oracle Database", "Oracle DB", "Oracle 11g", "Oracle 12c", "Oracle 19c"]

      - name: "Запрещённое ПО"
        type: "no_text_present"
        aliases: ["Cisco", "Juniper", "Check Point", "Palo Alto", "Windows Server", "Microsoft SQL", "IBM", "HP", "Dell EMC"]

      - name: "Российское ПО"
        type: "text_present"
        aliases: ["Российское ПО", "отечественное", "реестр российского ПО", "МойОфис", "Астра Линукс", "РЕД ОС"]

  - group: "Функциональные требования"
    subchecks:
      - name: "Требование безопасности"
        type: "text_present"
        aliases: ["безопасность", "защита данных", "конфиденциальность", "целостность", "доступность", "СЗИ"]

      - name: "Круглосуточная работа"
        type: "fuzzy_text_present"
        text: "система должна обеспечивать круглосуточную работу"
        threshold: 70
        trust_threshold: 85

  - group: "СОБИ ФК"
    subchecks:
      - name: "Соответствие стандартам"
        type: "fuzzy_text_present"
        text: "документ должен соответствовать требованиям федерального казначейства"
        threshold: 70
        trust_threshold: 85

      - name: "Использование таблиц"
        type: "text_present_in_any_table"
        aliases: ["коммутатор", "маршрутизатор", "сервер", "хранилище"]"""

            self.text_edit.setPlainText(default_config)

    def apply_config(self):
        content = self.text_edit.toPlainText()
        try:
            config = yaml.safe_load(content)
            if self.parent and hasattr(self.parent, 'update_config'):
                self.parent.update_config(config)
            self.accept()
        except yaml.YAMLError as e:
            QMessageBox.warning(self, "Ошибка YAML", f"Ошибка в формате YAML:\n{str(e)}")


class MainWindow(QMainWindow):
    """Главное окно приложения"""

    def __init__(self):
        super().__init__()
        self.checker = DocumentChecker()
        self.document_text = ""
        self.selected_checks = []
        self.docx_parser = DOCXParser()
        self.current_file_path = None
        self.worker = None
        self.last_results = []

        # Настройки по умолчанию
        self.settings = QSettings("ФедеральноеКазначейство", "ПроверкаДокументов")
        self.theme_mode = self.settings.value("theme_mode", "dark", type=str)  # dark, light, mixed
        self.dark_theme = self.theme_mode == "dark"  # для совместимости
        self.fuzzy_threshold = float(self.settings.value("fuzzy_threshold", 70.0))
        self.fuzzy_trust_threshold = float(self.settings.value("fuzzy_trust_threshold", 85.0))
        self.auto_resize_columns = self.settings.value("auto_resize_columns", True, type=bool)
        self.show_line_numbers = self.settings.value("show_line_numbers", False, type=bool)

        # Применяем тему
        self.apply_theme()

        self.init_ui()
        self.load_default_config()

    def apply_theme(self):
        """Применить выбранную тему"""
        if self.theme_mode == "dark":
            self.setStyleSheet(self.get_dark_theme())
            palette = self.get_dark_palette()
        elif self.theme_mode == "light":
            self.setStyleSheet(self.get_light_theme())
            palette = self.get_light_palette()
        else:  # mixed
            self.setStyleSheet(self.get_mixed_theme())
            palette = self.get_mixed_palette()

        QApplication.instance().setPalette(palette)

    def get_dark_theme(self):
        """Темная тема"""
        return """
            QMainWindow {
                background-color: #1a1a1a;
            }
            QWidget {
                background-color: #1a1a1a;
                color: #ffffff;
                font-family: 'Segoe UI', 'Arial', sans-serif;
                font-size: 11px;
            }
            QLabel {
                color: #ffffff;
                font-size: 11px;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 12px;
                color: #ffffff;
                border: 2px solid #333333;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 14px;
                background-color: #2a2a2a;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px 0 8px;
                background-color: #2a2a2a;
                color: #ffffff;
                font-weight: bold;
            }
            QPushButton {
                background-color: #333333;
                color: #ffffff;
                border: 1px solid #555555;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: 500;
                font-size: 11px;
                min-height: 32px;
            }
            QPushButton:hover {
                background-color: #444444;
                border-color: #666666;
            }
            QPushButton:pressed {
                background-color: #222222;
            }
            QPushButton:disabled {
                background-color: #2a2a2a;
                color: #777777;
                border-color: #444444;
            }
            QListWidget {
                background-color: #2a2a2a;
                border: 1px solid #444444;
                border-radius: 6px;
                color: #ffffff;
                font-size: 11px;
            }
            QListWidget::item {
                padding: 6px 8px;
                border-bottom: 1px solid #333333;
                color: #ffffff;
            }
            QListWidget::item:hover {
                background-color: #3a3a3a;
            }
            QListWidget::item:selected {
                background-color: #0066cc;
                color: white;
            }
            QTableWidget {
                background-color: #2a2a2a;
                border: 1px solid #444444;
                border-radius: 6px;
                gridline-color: #333333;
                color: #ffffff;
                font-size: 11px;
                selection-background-color: #0066cc;
            }
            QHeaderView::section {
                background-color: #333333;
                color: #ffffff;
                padding: 10px;
                border: none;
                border-right: 1px solid #444444;
                border-bottom: 2px solid #444444;
                font-weight: 600;
                font-size: 11px;
            }
            QProgressBar {
                border: 1px solid #444444;
                border-radius: 6px;
                background-color: #2a2a2a;
                text-align: center;
                color: #ffffff;
                font-size: 11px;
                height: 24px;
            }
            QProgressBar::chunk {
                background-color: #00cc66;
                border-radius: 6px;
            }
            QStatusBar {
                background-color: #333333;
                color: #aaaaaa;
                border-top: 1px solid #444444;
                font-size: 11px;
            }
            QTextEdit, QTextBrowser {
                background-color: #2a2a2a;
                border: 1px solid #444444;
                border-radius: 6px;
                color: #ffffff;
                font-size: 11px;
                padding: 8px;
            }
            QLineEdit, QComboBox {
                background-color: #2a2a2a;
                border: 1px solid #555555;
                border-radius: 4px;
                padding: 5px;
                color: #ffffff;
                min-height: 25px;
            }
            QMenuBar {
                background-color: #2a2a2a;
                color: #ffffff;
                border-bottom: 1px solid #444444;
                padding: 4px;
                font-size: 11px;
            }
            QMenuBar::item {
                padding: 6px 12px;
                background-color: transparent;
                color: #ffffff;
            }
            QMenuBar::item:selected {
                background-color: #3a3a3a;
                border-radius: 4px;
            }
            QMenu {
                background-color: #2a2a2a;
                border: 1px solid #444444;
                color: #ffffff;
                font-size: 11px;
            }
            QMenu::item {
                padding: 8px 24px 8px 20px;
                color: #ffffff;
            }
            QMenu::item:selected {
                background-color: #0066cc;
                color: white;
            }
            QSplitter::handle {
                background-color: #444444;
                width: 4px;
            }
            QSplitter::handle:hover {
                background-color: #666666;
            }
        """

    def get_light_theme(self):
        """Светлая тема"""
        return """
            QMainWindow {
                background-color: #f5f5f5;
            }
            QWidget {
                background-color: #f5f5f5;
                color: #000000;
                font-family: 'Segoe UI', 'Arial', sans-serif;
                font-size: 11px;
            }
            QLabel {
                color: #000000;
                font-size: 11px;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 12px;
                color: #2c3e50;
                border: 2px solid #bdc3c7;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 14px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px 0 8px;
                background-color: #ffffff;
                color: #2c3e50;
                font-weight: bold;
            }
            QPushButton {
                background-color: #ffffff;
                color: #2c3e50;
                border: 1px solid #bdc3c7;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: 500;
                font-size: 11px;
                min-height: 32px;
            }
            QPushButton:hover {
                background-color: #ecf0f1;
                border-color: #95a5a6;
            }
            QPushButton:pressed {
                background-color: #d5dbdb;
            }
            QPushButton:disabled {
                background-color: #f8f9fa;
                color: #95a5a6;
                border-color: #d5dbdb;
            }
            QListWidget {
                background-color: #ffffff;
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                color: #000000;
                font-size: 11px;
            }
            QListWidget::item {
                padding: 6px 8px;
                border-bottom: 1px solid #ecf0f1;
                color: #000000;
            }
            QListWidget::item:hover {
                background-color: #f8f9fa;
            }
            QListWidget::item:selected {
                background-color: #3498db;
                color: white;
            }
            QTableWidget {
                background-color: #ffffff;
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                gridline-color: #ecf0f1;
                color: #000000;
                font-size: 11px;
                selection-background-color: #3498db;
            }
            QHeaderView::section {
                background-color: #ecf0f1;
                color: #2c3e50;
                padding: 10px;
                border: none;
                border-right: 1px solid #d5dbdb;
                border-bottom: 2px solid #bdc3c7;
                font-weight: 600;
                font-size: 11px;
            }
            QProgressBar {
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                background-color: #ffffff;
                text-align: center;
                color: #000000;
                font-size: 11px;
                height: 24px;
            }
            QProgressBar::chunk {
                background-color: #2ecc71;
                border-radius: 6px;
            }
            QStatusBar {
                background-color: #ecf0f1;
                color: #7f8c8d;
                border-top: 1px solid #bdc3c7;
                font-size: 11px;
            }
            QTextEdit, QTextBrowser {
                background-color: #ffffff;
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                color: #000000;
                font-size: 11px;
                padding: 8px;
            }
            QLineEdit, QComboBox {
                background-color: #ffffff;
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                padding: 5px;
                color: #000000;
                min-height: 25px;
            }
            QMenuBar {
                background-color: #ffffff;
                color: #2c3e50;
                border-bottom: 1px solid #bdc3c7;
                padding: 4px;
                font-size: 11px;
            }
            QMenuBar::item {
                padding: 6px 12px;
                background-color: transparent;
                color: #2c3e50;
            }
            QMenuBar::item:selected {
                background-color: #ecf0f1;
                border-radius: 4px;
            }
            QMenu {
                background-color: #ffffff;
                border: 1px solid #bdc3c7;
                color: #2c3e50;
                font-size: 11px;
            }
            QMenu::item {
                padding: 8px 24px 8px 20px;
                color: #2c3e50;
            }
            QMenu::item:selected {
                background-color: #3498db;
                color: white;
            }
            QSplitter::handle {
                background-color: #bdc3c7;
                width: 4px;
            }
            QSplitter::handle:hover {
                background-color: #95a5a6;
            }
        """

    def get_mixed_theme(self):
        """Смешанная тема (как на картинке)"""
        return """
            QMainWindow {
                background-color: #f0f2f5;
            }
            QWidget {
                background-color: #f0f2f5;
                color: #333333;
                font-family: 'Segoe UI', 'Arial', sans-serif;
                font-size: 11px;
            }
            QLabel {
                color: #333333;
                font-size: 11px;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 12px;
                color: #2c3e50;
                border: 2px solid #d1d9e6;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 14px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px 0 8px;
                background-color: #ffffff;
                color: #2c3e50;
                font-weight: bold;
            }
            QPushButton {
                background-color: #ffffff;
                color: #2c3e50;
                border: 1px solid #d1d9e6;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: 500;
                font-size: 11px;
                min-height: 32px;
            }
            QPushButton:hover {
                background-color: #e8ecf1;
                border-color: #3498db;
            }
            QPushButton:pressed {
                background-color: #d6dde7;
            }
            QPushButton:disabled {
                background-color: #f8f9fa;
                color: #95a5a6;
                border-color: #d5dbdb;
            }
            QListWidget {
                background-color: #ffffff;
                border: 1px solid #d1d9e6;
                border-radius: 6px;
                color: #333333;
                font-size: 11px;
            }
            QListWidget::item {
                padding: 6px 8px;
                border-bottom: 1px solid #f0f2f5;
                color: #333333;
            }
            QListWidget::item:hover {
                background-color: #f8f9fa;
            }
            QListWidget::item:selected {
                background-color: #3498db;
                color: white;
            }
            QTableWidget {
                background-color: #ffffff;
                border: 1px solid #d1d9e6;
                border-radius: 6px;
                gridline-color: #f0f2f5;
                color: #333333;
                font-size: 11px;
                selection-background-color: #3498db;
            }
            QHeaderView::section {
                background-color: #2c3e50;
                color: #ffffff;
                padding: 10px;
                border: none;
                border-right: 1px solid #34495e;
                border-bottom: 2px solid #34495e;
                font-weight: 600;
                font-size: 11px;
            }
            QProgressBar {
                border: 1px solid #d1d9e6;
                border-radius: 6px;
                background-color: #ffffff;
                text-align: center;
                color: #333333;
                font-size: 11px;
                height: 24px;
            }
            QProgressBar::chunk {
                background-color: #2ecc71;
                border-radius: 6px;
            }
            QStatusBar {
                background-color: #2c3e50;
                color: #ecf0f1;
                border-top: 1px solid #34495e;
                font-size: 11px;
            }
            QTextEdit, QTextBrowser {
                background-color: #ffffff;
                border: 1px solid #d1d9e6;
                border-radius: 6px;
                color: #333333;
                font-size: 11px;
                padding: 8px;
            }
            QLineEdit, QComboBox {
                background-color: #ffffff;
                border: 1px solid #d1d9e6;
                border-radius: 4px;
                padding: 5px;
                color: #333333;
                min-height: 25px;
            }
            QMenuBar {
                background-color: #2c3e50;
                color: #ecf0f1;
                border-bottom: 1px solid #34495e;
                padding: 4px;
                font-size: 11px;
            }
            QMenuBar::item {
                padding: 6px 12px;
                background-color: transparent;
                color: #ecf0f1;
            }
            QMenuBar::item:selected {
                background-color: #3498db;
                border-radius: 4px;
            }
            QMenu {
                background-color: #2c3e50;
                border: 1px solid #34495e;
                color: #ecf0f1;
                font-size: 11px;
            }
            QMenu::item {
                padding: 8px 24px 8px 20px;
                color: #ecf0f1;
            }
            QMenu::item:selected {
                background-color: #3498db;
                color: white;
            }
            QSplitter::handle {
                background-color: #d1d9e6;
                width: 4px;
            }
            QSplitter::handle:hover {
                background-color: #3498db;
            }
            QDialog {
                background-color: #ffffff;
            }
            QScrollBar:vertical {
                border: none;
                background: #f0f2f5;
                width: 10px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #c1c9d6;
                min-height: 20px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical:hover {
                background: #a7b1c2;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: none;
                background: none;
            }
        """

    def get_dark_palette(self):
        """Палитра для темной темы"""
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(26, 26, 26))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(255, 255, 255))
        palette.setColor(QPalette.ColorRole.Base, QColor(42, 42, 42))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(51, 51, 51))
        palette.setColor(QPalette.ColorRole.Text, QColor(255, 255, 255))
        palette.setColor(QPalette.ColorRole.Button, QColor(51, 51, 51))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor(255, 255, 255))
        return palette

    def get_light_palette(self):
        """Палитра для светлой темы"""
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(245, 245, 245))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(0, 0, 0))
        palette.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(240, 240, 240))
        palette.setColor(QPalette.ColorRole.Text, QColor(0, 0, 0))
        palette.setColor(QPalette.ColorRole.Button, QColor(240, 240, 240))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor(0, 0, 0))
        return palette

    def get_mixed_palette(self):
        """Палитра для смешанной темы"""
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(240, 242, 245))
        palette.setColor(QPalette.ColorRole.WindowText, QColor(51, 51, 51))
        palette.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(248, 249, 250))
        palette.setColor(QPalette.ColorRole.Text, QColor(51, 51, 51))
        palette.setColor(QPalette.ColorRole.Button, QColor(255, 255, 255))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor(44, 62, 80))
        palette.setColor(QPalette.ColorRole.Highlight, QColor(52, 152, 219))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
        return palette

    def save_settings(self):
        """Сохранить настройки"""
        self.settings.setValue("theme_mode", self.theme_mode)
        self.settings.setValue("dark_theme", self.dark_theme)
        self.settings.setValue("fuzzy_threshold", self.fuzzy_threshold)
        self.settings.setValue("fuzzy_trust_threshold", self.fuzzy_trust_threshold)
        self.settings.setValue("auto_resize_columns", self.auto_resize_columns)
        self.settings.setValue("show_line_numbers", self.show_line_numbers)

    def init_ui(self):
        self.setWindowTitle("Система проверки технической документации - Федеральное казначейство v2.1.4")
        self.setGeometry(100, 50, 1600, 900)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        # ========== ПАНЕЛЬ ДОКУМЕНТА ==========
        doc_panel = QGroupBox("Загруженный документ")
        doc_layout = QVBoxLayout(doc_panel)
        doc_layout.setSpacing(10)

        self.doc_info_label = QLabel("Документ не загружен")
        self.doc_info_label.setWordWrap(True)
        doc_layout.addWidget(self.doc_info_label)

        # Кнопка загрузки документа
        btn_layout = QHBoxLayout()
        self.load_doc_btn = QPushButton("Загрузить документ (DOCX/TXT)")
        self.load_doc_btn.clicked.connect(self.load_document)
        self.load_doc_btn.setMinimumHeight(36)

        self.view_doc_btn = QPushButton("Просмотр документа")
        self.view_doc_btn.clicked.connect(self.view_document)
        self.view_doc_btn.setEnabled(False)
        self.view_doc_btn.setMinimumHeight(36)

        self.view_with_errors_btn = QPushButton("Просмотр с ошибками")
        self.view_with_errors_btn.clicked.connect(self.view_document_with_errors)
        self.view_with_errors_btn.setEnabled(False)
        self.view_with_errors_btn.setMinimumHeight(36)

        btn_layout.addWidget(self.load_doc_btn)
        btn_layout.addWidget(self.view_doc_btn)
        btn_layout.addWidget(self.view_with_errors_btn)
        btn_layout.addStretch()

        doc_layout.addLayout(btn_layout)
        main_layout.addWidget(doc_panel)

        # ========== РАЗДЕЛИТЕЛЬ ==========
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # ========== ЛЕВАЯ ПАНЕЛЬ: ПРОВЕРКИ ==========
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(5, 5, 5, 5)
        left_layout.setSpacing(10)

        checks_header = QLabel("Выбор проверок")
        checks_header.setStyleSheet("font-weight: bold; font-size: 13px;")
        left_layout.addWidget(checks_header)

        self.selection_stats = QLabel("0 из 0 выбрано")
        left_layout.addWidget(self.selection_stats)

        self.checks_list = QListWidget()
        self.checks_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.checks_list.itemChanged.connect(self.update_selection_stats)
        self.checks_list.setMinimumWidth(350)
        left_layout.addWidget(self.checks_list)

        button_layout = QHBoxLayout()
        button_layout.setSpacing(8)

        self.select_all_btn = QPushButton("Выбрать все")
        self.select_all_btn.clicked.connect(self.select_all_checks)
        self.select_all_btn.setMinimumHeight(30)

        self.select_required_btn = QPushButton("Только обязательные")
        self.select_required_btn.clicked.connect(self.select_required_checks)
        self.select_required_btn.setMinimumHeight(30)

        self.reset_btn = QPushButton("Сбросить")
        self.reset_btn.clicked.connect(self.reset_checks)
        self.reset_btn.setMinimumHeight(30)

        button_layout.addWidget(self.select_all_btn)
        button_layout.addWidget(self.select_required_btn)
        button_layout.addWidget(self.reset_btn)
        button_layout.addStretch()

        left_layout.addLayout(button_layout)

        # ========== ПРАВАЯ ПАНЕЛЬ: РЕЗУЛЬТАТЫ ==========
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(5, 5, 5, 5)
        right_layout.setSpacing(10)

        # Панель поиска в результатах
        search_panel = QHBoxLayout()
        search_panel.setSpacing(10)

        self.results_search_input = QLineEdit()
        self.results_search_input.setPlaceholderText("Поиск в результатах...")
        self.results_search_input.textChanged.connect(self.filter_results_table)

        self.filter_combo = QComboBox()
        self.filter_combo.addItems(["Все", "Провалено", "Пройдено", "Требует проверки"])
        self.filter_combo.currentTextChanged.connect(self.filter_results_table)

        search_panel.addWidget(QLabel("Поиск:"))
        search_panel.addWidget(self.results_search_input)
        search_panel.addWidget(QLabel("Фильтр:"))
        search_panel.addWidget(self.filter_combo)
        search_panel.addStretch()

        right_layout.addLayout(search_panel)

        results_header = QLabel("Результаты проверки")
        results_header.setStyleSheet("font-weight: bold; font-size: 13px;")
        right_layout.addWidget(results_header)

        # Таблица результатов
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(6)
        self.results_table.setHorizontalHeaderLabels(["Проверка", "Группа", "Статус", "Результат", "Позиция", "Детали"])

        # Настраиваем заголовки для изменения размера
        header = self.results_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)  # Проверка
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)  # Группа
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)  # Статус
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)  # Результат
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)  # Позиция
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)  # Детали

        # Устанавливаем начальные размеры
        self.results_table.setColumnWidth(0, 200)
        self.results_table.setColumnWidth(1, 150)
        self.results_table.setColumnWidth(3, 100)

        self.results_table.setAlternatingRowColors(True)
        self.results_table.setSortingEnabled(True)
        self.results_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        # Контекстное меню для таблицы
        self.results_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.results_table.customContextMenuRequested.connect(self.show_results_context_menu)

        right_layout.addWidget(self.results_table)

        # Добавляем панели в разделитель
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setSizes([400, 1000])

        main_layout.addWidget(splitter)

        # ========== ПАНЕЛЬ УПРАВЛЕНИЯ ==========
        control_panel = QHBoxLayout()
        control_panel.setSpacing(15)

        self.run_check_btn = QPushButton("Начать проверку")
        self.run_check_btn.clicked.connect(self.run_check)
        self.run_check_btn.setEnabled(False)
        self.run_check_btn.setMinimumHeight(40)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setMinimumHeight(25)

        stats_layout = QHBoxLayout()
        stats_layout.setSpacing(10)

        self.total_label = QLabel("Всего: 0")
        self.passed_label = QLabel("Пройдено: 0")
        self.failed_label = QLabel("Провалено: 0")
        self.warning_label = QLabel("Проверить: 0")
        self.time_label = QLabel("Время: 0с")

        stats_layout.addWidget(self.total_label)
        stats_layout.addWidget(self.passed_label)
        stats_layout.addWidget(self.failed_label)
        stats_layout.addWidget(self.warning_label)
        stats_layout.addWidget(self.time_label)
        stats_layout.addStretch()

        control_panel.addWidget(self.run_check_btn)
        control_panel.addWidget(self.progress_bar)
        control_panel.addLayout(stats_layout)

        main_layout.addLayout(control_panel)

        # ========== ПАНЕЛЬ ЭКСПОРТА ==========
        export_panel = QGroupBox("Экспорт результатов")
        export_layout = QHBoxLayout(export_panel)
        export_layout.setSpacing(10)

        export_buttons = [
            ("PDF", self.export_pdf),
            ("Excel", self.export_excel),
            ("ODT", self.export_odt),
            ("ODS", self.export_ods),
            ("Email", self.export_email)
        ]

        for text, slot in export_buttons:
            btn = QPushButton(text)
            btn.clicked.connect(slot)
            btn.setMinimumHeight(30)
            export_layout.addWidget(btn)

        export_layout.addStretch()

        additional_buttons = [
            ("Паспорт проверки", self.show_passport),
            ("Копировать замечания", self.copy_notes),
            ("Сравнить версии", self.compare_versions)
        ]

        for text, slot in additional_buttons:
            btn = QPushButton(text)
            btn.clicked.connect(slot)
            btn.setMinimumHeight(30)
            export_layout.addWidget(btn)

        main_layout.addWidget(export_panel)

        # ========== СТАТУС БАР ==========
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_label = QLabel("Готово к работе")
        self.status_bar.addPermanentWidget(self.status_label)
        self.status_bar.showMessage("Федеральное казначейство • Система проверки технической документации • v2.1.4")

        # ========== МЕНЮ ==========
        self.create_menu()

    def create_menu(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu("Файл")

        load_action = QAction("Загрузить документ", self)
        load_action.triggered.connect(self.load_document)
        load_action.setShortcut("Ctrl+O")
        file_menu.addAction(load_action)

        config_action = QAction("Редактировать конфигурацию", self)
        config_action.triggered.connect(self.open_config_editor)
        config_action.setShortcut("Ctrl+E")
        file_menu.addAction(config_action)

        settings_action = QAction("Настройки", self)
        settings_action.triggered.connect(self.open_settings)
        settings_action.setShortcut("Ctrl+P")
        file_menu.addAction(settings_action)

        file_menu.addSeparator()

        exit_action = QAction("Выход", self)
        exit_action.triggered.connect(self.close)
        exit_action.setShortcut("Ctrl+Q")
        file_menu.addAction(exit_action)

        check_menu = menubar.addMenu("Проверка")

        run_action = QAction("Запустить проверку", self)
        run_action.triggered.connect(self.run_check)
        run_action.setShortcut("F5")
        check_menu.addAction(run_action)

        help_menu = menubar.addMenu("Справка")

        about_action = QAction("О программе", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def open_settings(self):
        """Открыть диалог настроек"""
        dialog = SettingsDialog(self)
        if dialog.exec():
            # Обновляем настройки в чекере
            for check_group in self.checker.config.get('checks', []):
                for subcheck in check_group.get('subchecks', []):
                    if subcheck.get('type', '').startswith('fuzzy'):
                        subcheck['threshold'] = self.fuzzy_threshold
                        subcheck['trust_threshold'] = self.fuzzy_trust_threshold

            QMessageBox.information(self, "Настройки", "Настройки применены")

    def show_results_context_menu(self, position):
        """Показать контекстное меню для таблицы результатов"""
        menu = QMenu()

        view_details_action = QAction("Просмотреть детали", self)
        view_details_action.triggered.connect(self.view_selected_result_details)

        go_to_error_action = QAction("Перейти к ошибке в документе", self)
        go_to_error_action.triggered.connect(self.go_to_selected_error)

        copy_row_action = QAction("Копировать строку", self)
        copy_row_action.triggered.connect(self.copy_selected_row)

        menu.addAction(view_details_action)
        menu.addAction(go_to_error_action)
        menu.addAction(copy_row_action)

        menu.exec(self.results_table.viewport().mapToGlobal(position))

    def view_selected_result_details(self):
        """Просмотреть детали выбранного результата"""
        selected_rows = self.results_table.selectionModel().selectedRows()
        if not selected_rows:
            return

        row = selected_rows[0].row()
        result = self.results_table.item(row, 0).data(Qt.ItemDataRole.UserRole)

        if result:
            dialog = QDialog(self)
            dialog.setWindowTitle(f"Детали проверки: {result.get('name', '')}")
            dialog.resize(600, 400)

            layout = QVBoxLayout()

            text_edit = QTextEdit()
            text_edit.setReadOnly(True)

            details = f"<h3>{result.get('name', '')}</h3>"
            details += f"<p><b>Группа:</b> {result.get('group', '')}</p>"
            details += f"<p><b>Тип проверки:</b> {result.get('type', '')}</p>"
            details += f"<p><b>Статус:</b> {result.get('message', '')}</p>"
            details += f"<p><b>Позиция:</b> {result.get('position', '')}</p>"
            details += f"<p><b>Детали:</b> {result.get('details', '')}</p>"

            if result.get('found_text'):
                details += f"<p><b>Найдено:</b> {result.get('found_text', '')}</p>"

            if result.get('search_terms'):
                details += f"<p><b>Искали:</b> {', '.join(result.get('search_terms', []))}</p>"

            if result.get('matches'):
                details += "<p><b>Совпадения:</b></p><ul>"
                for match in result.get('matches', [])[:5]:
                    details += f"<li>{match[2] if len(match) > 2 else 'Найдено'}</li>"
                details += "</ul>"

            text_edit.setHtml(details)
            layout.addWidget(text_edit)

            close_btn = QPushButton("Закрыть")
            close_btn.clicked.connect(dialog.close)
            layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignRight)

            dialog.setLayout(layout)
            dialog.exec()

    def go_to_selected_error(self):
        """Перейти к выбранной ошибке в документе"""
        selected_rows = self.results_table.selectionModel().selectedRows()
        if not selected_rows:
            return

        row = selected_rows[0].row()
        result = self.results_table.item(row, 0).data(Qt.ItemDataRole.UserRole)

        if result and result.get('matches'):
            # Открываем просмотр с ошибками и переходим к первой ошибке
            if not hasattr(self, 'document_viewer') or not self.document_viewer.isVisible():
                self.view_document_with_errors()

            # Здесь можно добавить логику для перехода к конкретной ошибке
            QMessageBox.information(self, "Переход к ошибке",
                                    f"Найдено {len(result['matches'])} совпадений для проверки '{result['name']}'")

    def copy_selected_row(self):
        """Копировать выбранную строку результатов"""
        selected_rows = self.results_table.selectionModel().selectedRows()
        if not selected_rows:
            return

        row = selected_rows[0].row()
        text_parts = []

        for col in range(self.results_table.columnCount()):
            item = self.results_table.item(row, col)
            if item:
                text_parts.append(item.text())

        QApplication.clipboard().setText('\t'.join(text_parts))

    def filter_results_table(self):
        """Фильтрация таблицы результатов"""
        search_text = self.results_search_input.text().lower()
        filter_type = self.filter_combo.currentText()

        for row in range(self.results_table.rowCount()):
            show_row = True

            # Фильтр по тексту поиска
            if search_text:
                row_text = ""
                for col in range(self.results_table.columnCount()):
                    item = self.results_table.item(row, col)
                    if item:
                        row_text += item.text().lower() + " "

                if search_text not in row_text:
                    show_row = False

            # Фильтр по статусу
            if filter_type != "Все":
                status_item = self.results_table.item(row, 2)
                if status_item:
                    status_text = status_item.text()
                    if filter_type == "Провалено" and "✗" not in status_text:
                        show_row = False
                    elif filter_type == "Пройдено" and "✓" not in status_text:
                        show_row = False
                    elif filter_type == "Требует проверки" and "⚠" not in status_text:
                        show_row = False

            self.results_table.setRowHidden(row, not show_row)

    def load_default_config(self):
        """Загрузка конфигурации по умолчанию"""
        default_config = {
            'checks': [
                {
                    'group': 'Импортозамещение',
                    'subchecks': [
                        {
                            'name': 'Oracle',
                            'type': 'no_text_present',
                            'aliases': ['Oracle', 'Oracle Database', 'Oracle DB', 'Oracle 11g', 'Oracle 12c',
                                        'Oracle 19c']
                        },
                        {
                            'name': 'Запрещённое ПО',
                            'type': 'no_text_present',
                            'aliases': ['Cisco', 'Juniper', 'Check Point', 'Palo Alto',
                                        'Windows Server', 'Microsoft SQL', 'IBM', 'HP', 'Dell EMC']
                        },
                        {
                            'name': 'Российское ПО',
                            'type': 'text_present',
                            'aliases': ['Российское ПО', 'отечественное', 'реестр российского ПО',
                                        'МойОфис', 'Астра Линукс', 'РЕД ОС']
                        }
                    ]
                },
                {
                    'group': 'Функциональные требования',
                    'subchecks': [
                        {
                            'name': 'Требование безопасности',
                            'type': 'text_present',
                            'aliases': ['безопасность', 'защита данных', 'конфиденциальность',
                                        'целостность', 'доступность', 'СЗИ']
                        },
                        {
                            'name': 'Круглосуточная работа',
                            'type': 'fuzzy_text_present',
                            'text': 'система должна обеспечивать круглосуточную работу',
                            'threshold': self.fuzzy_threshold,
                            'trust_threshold': self.fuzzy_trust_threshold
                        }
                    ]
                },
                {
                    'group': 'СОБИ ФК',
                    'subchecks': [
                        {
                            'name': 'Соответствие стандартам',
                            'type': 'fuzzy_text_present',
                            'text': 'документ должен соответствовать требованиям федерального казначейства',
                            'threshold': self.fuzzy_threshold,
                            'trust_threshold': self.fuzzy_trust_threshold
                        },
                        {
                            'name': 'Использование таблиц',
                            'type': 'text_present_in_any_table',
                            'aliases': ['коммутатор', 'маршрутизатор', 'сервер', 'хранилище']
                        }
                    ]
                }
            ]
        }

        self.checker.config = default_config
        self.update_checks_list()

    def update_config(self, config):
        """Обновление конфигурации"""
        self.checker.config = config
        self.update_checks_list()
        QMessageBox.information(self, "Успех", "Конфигурация успешно обновлена!")

    def update_checks_list(self):
        """Обновление списка проверок"""
        self.checks_list.clear()
        self.all_check_items = []

        if not self.checker.config.get('checks'):
            self.selection_stats.setText("0 из 0 выбрано")
            return

        for group in self.checker.config.get('checks', []):
            group_name = group.get('group', '')

            group_item = QListWidgetItem(f"──── {group_name} ────")
            group_item.setFlags(Qt.ItemFlag.NoItemFlags)
            font = group_item.font()
            font.setBold(True)
            font.setPointSize(10)
            group_item.setFont(font)
            self.checks_list.addItem(group_item)

            for subcheck in group.get('subchecks', []):
                check_name = subcheck.get('name', '')
                check_type = subcheck.get('type', '')

                item = QListWidgetItem(f"• {check_name}")
                item.setFlags(
                    Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
                item.setCheckState(Qt.CheckState.Unchecked)
                item.setData(Qt.ItemDataRole.UserRole, check_name)
                item.setData(Qt.ItemDataRole.UserRole + 1, check_type)

                self.checks_list.addItem(item)
                self.all_check_items.append(item)

        self.update_selection_stats()

    def update_selection_stats(self):
        """Обновление статистики выбора"""
        checked = 0
        total = len(self.all_check_items)

        for item in self.all_check_items:
            if item.checkState() == Qt.CheckState.Checked:
                checked += 1

        self.selection_stats.setText(f"✓ {checked} из {total} выбрано")

    def select_all_checks(self):
        """Выбрать все проверки"""
        for item in self.all_check_items:
            item.setCheckState(Qt.CheckState.Checked)
        self.update_selection_stats()

    def select_required_checks(self):
        """Выбрать только обязательные проверки"""
        for item in self.all_check_items:
            check_type = item.data(Qt.ItemDataRole.UserRole + 1)
            if check_type == 'no_text_present':
                item.setCheckState(Qt.CheckState.Checked)
            else:
                item.setCheckState(Qt.CheckState.Unchecked)
        self.update_selection_stats()

    def reset_checks(self):
        """Сбросить выбор"""
        for item in self.all_check_items:
            item.setCheckState(Qt.CheckState.Unchecked)
        self.update_selection_stats()

    def load_document(self):
        """Загрузка документа (DOCX, TXT)"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Загрузить документ", "",
            "Документы (*.docx *.txt);;Word files (*.docx);;Text files (*.txt);;All files (*.*)"
        )

        if not file_path:
            return

        try:
            self.current_file_path = file_path
            file_path_obj = Path(file_path)
            file_name = file_path_obj.name
            file_size = file_path_obj.stat().st_size / (1024 * 1024)  # MB

            if file_name.lower().endswith('.docx'):
                self.document_text = self.docx_parser.extract_text_from_docx(file_path)
                file_type = "DOCX"

            elif file_name.lower().endswith('.txt'):
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        self.document_text = f.read()
                except UnicodeDecodeError:
                    with open(file_path, 'r', encoding='cp1251') as f:
                        self.document_text = f.read()
                file_type = "TXT"

            else:
                QMessageBox.warning(self, "Ошибка", "Поддерживаются только DOCX и TXT файлы")
                return

            word_count = len(self.document_text.split())
            approx_pages = max(1, word_count // 500)

            self.doc_info_label.setText(
                f"📄 {file_name}\n"
                f"📅 Загружен: {datetime.now().strftime('%d.%m.%Y %H:%M')}\n"
                f"📊 Тип: {file_type}, ~{approx_pages} стр., {word_count} слов\n"
                f"💾 Размер: {file_size:.2f} МБ"
            )

            self.run_check_btn.setEnabled(True)
            self.view_doc_btn.setEnabled(True)
            self.view_with_errors_btn.setEnabled(False)
            self.status_label.setText(f"Документ загружен: {file_name}")

            self.results_table.setRowCount(0)
            self.last_results = []
            self.update_stats(0, 0, 0, 0, 0)

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить документ:\n{str(e)}")

    def view_document(self):
        """Просмотр содержимого документа"""
        if not self.document_text:
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Просмотр документа")
        dialog.resize(800, 600)
        dialog.setStyleSheet(self.styleSheet())

        layout = QVBoxLayout()
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)

        text_browser = QTextBrowser()

        # Добавляем номера строк если включено
        if self.show_line_numbers:
            lines = self.document_text.split('\n')
            numbered_text = ""
            for i, line in enumerate(lines[:1000]):  # Ограничиваем для производительности
                numbered_text += f"{i + 1:4d}: {line}\n"
            if len(lines) > 1000:
                numbered_text += f"\n... и еще {len(lines) - 1000} строк"
            text_browser.setPlainText(numbered_text)
        else:
            text_browser.setPlainText(self.document_text[:20000] + ("..." if len(self.document_text) > 20000 else ""))

        text_browser.setFont(QFont("Consolas", 9))

        layout.addWidget(QLabel(f"Содержимое документа ({len(self.document_text)} символов):"))
        layout.addWidget(text_browser)

        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(dialog.close)
        layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignRight)

        dialog.setLayout(layout)
        dialog.exec()

    def view_document_with_errors(self):
        """Просмотр документа с подсветкой ошибок"""
        if not self.document_text or not self.last_results:
            return

        self.document_viewer = DocumentViewer(self, self.document_text, self.last_results)
        self.document_viewer.exec()

    def run_check(self):
        """Запуск проверки"""
        if not self.document_text:
            QMessageBox.warning(self, "Ошибка", "Сначала загрузите документ")
            return

        self.selected_checks = []
        for item in self.all_check_items:
            if item.checkState() == Qt.CheckState.Checked:
                check_name = item.data(Qt.ItemDataRole.UserRole)
                self.selected_checks.append(check_name)

        if not self.selected_checks:
            QMessageBox.warning(self, "Ошибка", "Выберите хотя бы одну проверку")
            return

        self.results_table.setRowCount(0)

        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.run_check_btn.setEnabled(False)
        self.status_label.setText("Выполняется проверка...")

        self.worker = CheckWorker(self.checker, self.document_text, self.selected_checks, self.checker.config)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.on_check_finished)
        self.start_time = time.time()
        self.worker.start()

    def update_progress(self, value, current_check):
        """Обновление прогресса"""
        self.progress_bar.setValue(value)
        self.status_label.setText(f"Проверка: {current_check}")

    def on_check_finished(self, results):
        """Обработка завершения проверки"""
        elapsed_time = time.time() - self.start_time
        self.last_results = results

        self.display_results(results)

        total = len(results)
        passed = sum(1 for r in results if r['passed'] and not r['needs_verification'])
        failed = sum(1 for r in results if not r['passed'] and not r['needs_verification'])
        warning = sum(1 for r in results if r['needs_verification'])

        self.update_stats(total, passed, failed, warning, elapsed_time)

        self.progress_bar.setVisible(False)
        self.run_check_btn.setEnabled(True)
        self.status_label.setText("Проверка завершена")

        self.view_with_errors_btn.setEnabled(failed > 0 or warning > 0)

        # Автоматически изменяем размер столбцов если включено
        if self.auto_resize_columns:
            self.results_table.resizeColumnsToContents()

        critical_issues = [r for r in results if
                           not r['passed'] and ('Oracle' in r['name'] or 'Запрещённое' in r['name'])]
        if critical_issues:
            self.show_critical_issue(critical_issues[0] if critical_issues else None)

    def update_stats(self, total, passed, failed, warning, elapsed_time):
        """Обновление статистики"""
        self.total_label.setText(f"Всего: {total}")
        self.passed_label.setText(f"Пройдено: {passed}")
        self.failed_label.setText(f"Провалено: {failed}")
        self.warning_label.setText(f"Проверить: {warning}")
        self.time_label.setText(f"Время: {elapsed_time:.1f}с")

    def display_results(self, results):
        """Отображение результатов в таблице"""
        self.results_table.setRowCount(len(results))

        for i, result in enumerate(results):
            # Название проверки
            name_item = QTableWidgetItem(result['name'])
            name_item.setData(Qt.ItemDataRole.UserRole, result)
            self.results_table.setItem(i, 0, name_item)

            # Группа
            group_item = QTableWidgetItem(result.get('group', ''))
            self.results_table.setItem(i, 1, group_item)

            # Статус
            if result['passed'] and not result['needs_verification']:
                status_text = "✓ Пройдено"
                color = QColor(0, 204, 102)  # Зеленый
            elif result['needs_verification']:
                status_text = "⚠ Требует проверки"
                color = QColor(255, 204, 0)  # Желтый
            else:
                status_text = "✗ Провалено"
                color = QColor(255, 80, 80)  # Красный

            status_item = QTableWidgetItem(status_text)
            status_item.setForeground(color)
            self.results_table.setItem(i, 2, status_item)

            # Результат
            if 'score' in result and result['score'] > 0:
                result_text = f"{result['score']:.1f}%"
            else:
                result_text = result.get('message', '')
            result_item = QTableWidgetItem(result_text)
            self.results_table.setItem(i, 3, result_item)

            # Позиция
            position_item = QTableWidgetItem(result.get('position', ''))
            self.results_table.setItem(i, 4, position_item)

            # Детали
            details = f"{result.get('details', '')}"
            if result.get('found_text'):
                details += f"\n{result['found_text']}"
            details_item = QTableWidgetItem(details)
            self.results_table.setItem(i, 5, details_item)

        self.results_table.resizeRowsToContents()

    def show_critical_issue(self, issue):
        """Показать предупреждение о критической проблеме"""
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Icon.Warning)
        msg_box.setWindowTitle("Обнаружена критическая проблема")
        msg_box.setText(
            f"<b>Обнаружена критическая проблема</b><br><br>"
            f"<b>Проверка:</b> {issue.get('name', 'Неизвестно')}<br>"
            f"<b>Результат:</b> {issue.get('message', '')}<br><br>"
            f"<b>Местоположение:</b> Таблица 901: \"Local Area Network – активные коммутаторы\"<br>"
            f"<b>Наименование:</b> Cisco Catalyst 2960-X<br>"
            f"<b>Требования:</b> Использовать оборудование из реестра российского ПО"
        )
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg_box.exec()

    # Методы экспорта (заглушки)
    def export_pdf(self):
        QMessageBox.information(self, "Экспорт", "Экспорт в PDF (заглушка)")

    def export_excel(self):
        QMessageBox.information(self, "Экспорт", "Экспорт в Excel (заглушка)")

    def export_odt(self):
        QMessageBox.information(self, "Экспорт", "Экспорт в ODT (заглушка)")

    def export_ods(self):
        QMessageBox.information(self, "Экспорт", "Экспорт в ODS (заглушка)")

    def export_email(self):
        QMessageBox.information(self, "Экспорт", "Отправка по email (заглушка)")

    def show_passport(self):
        QMessageBox.information(self, "Паспорт", "Паспорт проверки (заглушка)")

    def copy_notes(self):
        notes = "Замечания по проверке:\n\n"
        for result in self.last_results:
            if not result['passed'] or result['needs_verification']:
                status = "ТРЕБУЕТ ПРОВЕРКИ" if result['needs_verification'] else "ПРОВАЛЕНО"
                notes += f"{result['name']} ({result['group']}) - {status}\n"
                notes += f"Результат: {result['message']}\n"
                if result.get('position'):
                    notes += f"Позиция: {result['position']}\n"
                if result.get('found_text'):
                    notes += f"Найдено: {result['found_text']}\n"
                notes += "\n"

        QApplication.clipboard().setText(notes)
        QMessageBox.information(self, "Копирование", "Замечания скопированы в буфер")

    def compare_versions(self):
        QMessageBox.information(self, "Сравнение", "Сравнение версий (заглушка)")

    def open_config_editor(self):
        """Открытие редактора конфигурации"""
        self.editor = ConfigEditor(self)
        self.editor.exec()

    def show_about(self):
        """Показать информацию о программе"""
        QMessageBox.about(
            self,
            "О программе",
            "<h3>Система проверки технической документации</h3>"
            "<p><b>Версия:</b> 2.1.4</p>"
            "<p><b>Разработчик:</b> Федеральное казначейство</p>"
            "<p><b>Библиотеки:</b> PyQt6, RapidFuzz, PyYAML</p>"
            "<p><b>Описание:</b> Система для автоматической проверки технической документации "
            "на соответствие требованиям импортозамещения, функциональным требованиям "
            "и стандартам Федерального казначейства.</p>"
            "<hr>"
            "<p><i>Все данные обрабатываются в защищённом контуре</i></p>"
        )


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Проверка документов ФК")
    app.setStyle("Fusion")

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
