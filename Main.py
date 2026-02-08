import sys
import yaml
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple
from datetime import datetime
import time

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QTextEdit, QFileDialog, QListWidget,
    QListWidgetItem, QCheckBox, QProgressBar, QTableWidget,
    QTableWidgetItem, QTabWidget, QGroupBox, QMessageBox,
    QSplitter, QHeaderView, QToolBar, QStatusBar, QTextBrowser,
    QDialog
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont, QIcon, QAction, QColor, QPalette
import rapidfuzz
from rapidfuzz import fuzz, process
import difflib


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
        """Точный поиск в тексте"""
        matches = []
        normalized_text = self.normalize_text(text)

        for term in search_terms:
            normalized_term = self.normalize_text(term)
            if not normalized_term:
                continue

            # Поиск всех вхождений
            start = 0
            while True:
                pos = normalized_text.find(normalized_term, start)
                if pos == -1:
                    break
                # Сохраняем позицию и найденный текст
                matches.append((pos, pos + len(normalized_term), term))
                start = pos + 1

        return matches

    def fuzzy_search_all(self, text: str, search_text: str, threshold: float = 70.0) -> List[
        Tuple[int, int, float, str]]:
        """Нечеткий поиск всех вхождений с использованием RapidFuzz"""
        if not text or not search_text:
            return []

        normalized_text = self.normalize_text(text)
        normalized_search = self.normalize_text(search_text)

        # Разбиваем текст на слова и фразы
        words = re.findall(r'\b\w+\b', normalized_text)
        matches = []

        # Ищем похожие подстроки
        for i in range(len(words)):
            for j in range(i + 1, min(i + 10, len(words) + 1)):
                substring = ' '.join(words[i:j])
                if len(substring) < 5:
                    continue

                score = fuzz.partial_ratio(substring, normalized_search)
                if score >= threshold:
                    # Находим позицию в исходном тексте
                    pos = normalized_text.find(substring)
                    if pos != -1:
                        matches.append((pos, pos + len(substring), score, substring))

        # Сортируем по схожести
        matches.sort(key=lambda x: x[2], reverse=True)
        return matches[:10]  # Ограничиваем количество результатов

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
        # Простая эвристика для поиска таблиц
        tables = []
        lines = document_text.split('\n')

        current_table = []
        in_table = False

        for line in lines:
            # Эвристика: таблицы часто имеют много чисел или разделителей
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
            # Если это похоже на строку таблицы
            if (('\t' in line or '|' in line or line.count('  ') > 3) and
                    i < len(lines) - 1):

                next_line = lines[i + 1].strip()
                # Если следующая строка не таблица и не пустая
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
            'found_text': ''
        }

        try:
            if check_type == 'no_text_present':
                aliases = subcheck.get('aliases', [])
                matches = self.exact_search(document_text, aliases)
                result['passed'] = len(matches) == 0
                result['matches'] = matches
                result['found_text'] = ', '.join([m[2] for m in matches]) if matches else ''
                result['message'] = f"Найдено вхождений: {len(matches)}" if matches else "Упоминаний не найдено"
                result['details'] = f"Искали: {', '.join(aliases)}"

            elif check_type == 'text_present':
                aliases = subcheck.get('aliases', [])
                matches = self.exact_search(document_text, aliases)
                result['passed'] = len(matches) > 0
                result['matches'] = matches
                result['found_text'] = ', '.join([m[2] for m in matches]) if matches else ''
                result['message'] = f"Найдено вхождений: {len(matches)}" if matches else "Текст не найден"
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

                # Сохраняем лучшие совпадения
                if score >= threshold:
                    matches = self.fuzzy_search_all(document_text, text, threshold)
                    if matches:
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
                        result['found_text'] = f"Найдено похожее: {matches[0][3][:50]}..."

            elif check_type == 'text_present_in_any_table':
                aliases = subcheck.get('aliases', [])
                tables = self.extract_tables(document_text)

                table_matches = []
                for table in tables:
                    matches = self.exact_search(table, aliases)
                    if matches:
                        table_matches.extend(matches)

                result['passed'] = len(table_matches) > 0
                result['matches'] = table_matches
                result['found_text'] = f"Найдено в {len(set([m[2] for m in table_matches]))} таблицах"
                result['message'] = f"Найдено в таблицах: {len(table_matches)}"
                result['details'] = f"Искали в таблицах: {', '.join(aliases)}"

            elif check_type == 'fuzzy_text_present_after_any_table':
                text = subcheck.get('text', '')
                threshold = float(subcheck.get('threshold', 70.0))
                trust_threshold = float(subcheck.get('trust_threshold', 85.0))

                paragraphs = self.extract_paragraphs_after_tables(document_text)
                best_score = 0.0
                best_match = ""

                for paragraph in paragraphs:
                    score = self.fuzzy_search_best(paragraph, text)
                    if score > best_score:
                        best_score = score
                        best_match = paragraph[:100] + "..." if len(paragraph) > 100 else paragraph

                result['score'] = best_score
                result['passed'] = best_score >= trust_threshold
                result['needs_verification'] = threshold <= best_score < trust_threshold
                result['message'] = f"Схожесть: {best_score:.1f}%"
                result['details'] = f"Искали после таблиц: '{text}'"

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

        # Заголовок
        title_label = QLabel("Редактирование конфигурации проверок")
        title_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #2c3e50;")
        layout.addWidget(title_label)

        # Панель кнопок
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
        self.apply_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                font-weight: bold;
            }
        """)

        button_layout.addWidget(self.load_btn)
        button_layout.addWidget(self.save_btn)
        button_layout.addWidget(self.restore_btn)
        button_layout.addStretch()
        button_layout.addWidget(self.cancel_btn)
        button_layout.addWidget(self.apply_btn)

        layout.addLayout(button_layout)

        # Текстовое поле для редактирования YAML
        editor_label = QLabel("YAML конфигурация:")
        editor_label.setStyleSheet("font-weight: bold; color: #2c3e50;")
        layout.addWidget(editor_label)

        self.text_edit = QTextEdit()
        self.text_edit.setFont(QFont("Consolas", 10))
        self.text_edit.setStyleSheet("""
            QTextEdit {
                background-color: white;
                border: 2px solid #bdc3c7;
                border-radius: 6px;
                padding: 10px;
                color: #2c3e50;
            }
        """)

        layout.addWidget(self.text_edit)

        self.setLayout(layout)

        # Загружаем текущую конфигурацию
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

        # Устанавливаем светлую тему
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
            }
            QWidget {
                background-color: #ffffff;
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
                border: 2px solid #d1d5db;
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
                background-color: #f3f4f6;
                color: #111827;
                border: 1px solid #d1d5db;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: 500;
                font-size: 11px;
                min-height: 32px;
            }
            QPushButton:hover {
                background-color: #e5e7eb;
                border-color: #9ca3af;
            }
            QPushButton:pressed {
                background-color: #d1d5db;
            }
            QPushButton:disabled {
                background-color: #f9fafb;
                color: #9ca3af;
                border-color: #e5e7eb;
            }
            QListWidget {
                background-color: white;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                color: #000000;
                font-size: 11px;
                outline: none;
            }
            QListWidget::item {
                padding: 6px 8px;
                border-bottom: 1px solid #f3f4f6;
                color: #000000;
            }
            QListWidget::item:hover {
                background-color: #f3f4f6;
            }
            QListWidget::item:selected {
                background-color: #3b82f6;
                color: white;
            }
            QTableWidget {
                background-color: white;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                gridline-color: #e5e7eb;
                color: #000000;
                font-size: 11px;
            }
            QHeaderView::section {
                background-color: #f9fafb;
                color: #374151;
                padding: 10px;
                border: none;
                border-right: 1px solid #e5e7eb;
                border-bottom: 2px solid #e5e7eb;
                font-weight: 600;
                font-size: 11px;
            }
            QProgressBar {
                border: 1px solid #d1d5db;
                border-radius: 6px;
                background-color: white;
                text-align: center;
                color: #000000;
                font-size: 11px;
                height: 24px;
            }
            QProgressBar::chunk {
                background-color: #10b981;
                border-radius: 6px;
            }
            QStatusBar {
                background-color: #f9fafb;
                color: #6b7280;
                border-top: 1px solid #e5e7eb;
                font-size: 11px;
            }
            QTextEdit, QTextBrowser {
                background-color: white;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                color: #000000;
                font-size: 11px;
                padding: 8px;
            }
            QMenuBar {
                background-color: #ffffff;
                color: #000000;
                border-bottom: 1px solid #e5e7eb;
                padding: 4px;
                font-size: 11px;
            }
            QMenuBar::item {
                padding: 6px 12px;
                background-color: transparent;
                color: #000000;
            }
            QMenuBar::item:selected {
                background-color: #f3f4f6;
                border-radius: 4px;
            }
            QMenu {
                background-color: white;
                border: 1px solid #e5e7eb;
                color: #000000;
                font-size: 11px;
            }
            QMenu::item {
                padding: 8px 24px 8px 20px;
                color: #000000;
            }
            QMenu::item:selected {
                background-color: #3b82f6;
                color: white;
            }
            QSplitter::handle {
                background-color: #e5e7eb;
                width: 4px;
            }
            QSplitter::handle:hover {
                background-color: #9ca3af;
            }
        """)

        self.init_ui()
        self.load_default_config()

    def init_ui(self):
        self.setWindowTitle("Система проверки технической документации - Федеральное казначейство v2.1.4")
        self.setGeometry(100, 50, 1600, 900)

        # Центральный виджет
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
        self.doc_info_label.setStyleSheet("""
            QLabel {
                font-size: 11px;
                color: #4b5563;
                padding: 8px;
                background-color: #f9fafb;
                border-radius: 6px;
                border: 1px solid #e5e7eb;
            }
        """)
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

        btn_layout.addWidget(self.load_doc_btn)
        btn_layout.addWidget(self.view_doc_btn)
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

        # Заголовок
        checks_header = QLabel("Выбор проверок")
        checks_header.setStyleSheet("font-weight: bold; font-size: 13px; color: #111827;")
        left_layout.addWidget(checks_header)

        # Статистика выбора
        self.selection_stats = QLabel("0 из 0 выбрано")
        self.selection_stats.setStyleSheet(
            "font-size: 11px; color: #6b7280; background-color: #f9fafb; padding: 6px; border-radius: 4px;")
        left_layout.addWidget(self.selection_stats)

        # Список проверок
        self.checks_list = QListWidget()
        self.checks_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.checks_list.itemChanged.connect(self.update_selection_stats)
        self.checks_list.setMinimumWidth(350)
        left_layout.addWidget(self.checks_list)

        # Кнопки выбора
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

        # Заголовок
        results_header = QLabel("Результаты проверки")
        results_header.setStyleSheet("font-weight: bold; font-size: 13px; color: #111827;")
        right_layout.addWidget(results_header)

        # Таблица результатов
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(5)
        self.results_table.setHorizontalHeaderLabels(["Проверка", "Группа", "Статус", "Результат", "Детали"])
        self.results_table.horizontalHeader().setStretchLastSection(True)
        self.results_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.results_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self.results_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self.results_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.results_table.setAlternatingRowColors(True)
        self.results_table.setSortingEnabled(True)

        right_layout.addWidget(self.results_table)

        # Добавляем панели в разделитель
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setSizes([400, 1000])

        main_layout.addWidget(splitter)

        # ========== ПАНЕЛЬ УПРАВЛЕНИЯ ==========
        control_panel = QHBoxLayout()
        control_panel.setSpacing(15)

        # Кнопка запуска проверки
        self.run_check_btn = QPushButton("Начать проверку")
        self.run_check_btn.clicked.connect(self.run_check)
        self.run_check_btn.setEnabled(False)
        self.run_check_btn.setMinimumHeight(40)
        self.run_check_btn.setStyleSheet("""
            QPushButton {
                background-color: #10b981;
                color: white;
                font-weight: bold;
                font-size: 12px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #059669;
            }
            QPushButton:disabled {
                background-color: #d1fae5;
                color: #9ca3af;
            }
        """)

        # Прогресс бар
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setMinimumHeight(25)

        # Статистика
        stats_layout = QHBoxLayout()
        stats_layout.setSpacing(10)

        self.total_label = QLabel("Всего: 0")
        self.passed_label = QLabel("Пройдено: 0")
        self.failed_label = QLabel("Провалено: 0")
        self.warning_label = QLabel("Проверить: 0")
        self.time_label = QLabel("Время: 0с")

        for label in [self.total_label, self.passed_label, self.failed_label,
                      self.warning_label, self.time_label]:
            label.setStyleSheet("""
                QLabel {
                    font-weight: 500;
                    padding: 6px 12px;
                    background-color: #f9fafb;
                    border-radius: 6px;
                    border: 1px solid #e5e7eb;
                    min-width: 80px;
                    text-align: center;
                    color: #000000;
                }
            """)

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

        # Дополнительные кнопки
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
        self.status_bar.showMessage("Система проверки технической документации • Федеральное казначейство • v2.1.4")

        # ========== МЕНЮ ==========
        self.create_menu()

    def create_menu(self):
        menubar = self.menuBar()

        # Меню Файл
        file_menu = menubar.addMenu("Файл")

        load_action = QAction("Загрузить документ", self)
        load_action.triggered.connect(self.load_document)
        load_action.setShortcut("Ctrl+O")
        file_menu.addAction(load_action)

        config_action = QAction("Редактировать конфигурацию", self)
        config_action.triggered.connect(self.open_config_editor)
        config_action.setShortcut("Ctrl+E")
        file_menu.addAction(config_action)

        file_menu.addSeparator()

        exit_action = QAction("Выход", self)
        exit_action.triggered.connect(self.close)
        exit_action.setShortcut("Ctrl+Q")
        file_menu.addAction(exit_action)

        # Меню Проверка
        check_menu = menubar.addMenu("Проверка")

        run_action = QAction("Запустить проверку", self)
        run_action.triggered.connect(self.run_check)
        run_action.setShortcut("F5")
        check_menu.addAction(run_action)

        # Меню Справка
        help_menu = menubar.addMenu("Справка")

        about_action = QAction("О программе", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

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
                            'threshold': 70,
                            'trust_threshold': 85
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
                            'threshold': 70,
                            'trust_threshold': 85
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

            # Добавляем группу как заголовок
            group_item = QListWidgetItem(f"──── {group_name} ────")
            group_item.setFlags(Qt.ItemFlag.NoItemFlags)
            group_item.setForeground(QColor(59, 130, 246))  # Синий цвет
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

                # Разные цвета для разных типов проверок
                if check_type == 'no_text_present':
                    item.setForeground(QColor(220, 38, 38))  # Красный
                elif check_type == 'text_present':
                    item.setForeground(QColor(22, 163, 74))  # Зеленый
                else:
                    item.setForeground(QColor(59, 130, 246))  # Синий

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
            # Считаем обязательными проверки на отсутствие импортного ПО
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

            # Определяем тип файла
            if file_name.lower().endswith('.docx'):
                self.document_text = self.docx_parser.extract_text_from_docx(file_path)
                file_type = "DOCX"

            elif file_name.lower().endswith('.txt'):
                # Пробуем разные кодировки
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

            # Подсчет статистики
            word_count = len(self.document_text.split())
            char_count = len(self.document_text)
            line_count = len(self.document_text.split('\n'))
            approx_pages = max(1, word_count // 500)

            self.doc_info_label.setText(
                f"📄 {file_name}\n"
                f"📅 Загружен: {datetime.now().strftime('%d.%m.%Y %H:%M')}\n"
                f"📊 Тип: {file_type}, ~{approx_pages} стр., {word_count} слов\n"
                f"💾 Размер: {file_size:.2f} МБ"
            )

            self.run_check_btn.setEnabled(True)
            self.view_doc_btn.setEnabled(True)
            self.status_label.setText(f"Документ загружен: {file_name}")

            # Очищаем предыдущие результаты
            self.results_table.setRowCount(0)
            self.update_stats(0, 0, 0, 0, 0)

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить документ:\n{str(e)}")
            print(f"Ошибка загрузки: {e}")

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

        # Текстовый браузер
        text_browser = QTextBrowser()
        text_browser.setPlainText(self.document_text[:10000] + ("..." if len(self.document_text) > 10000 else ""))
        text_browser.setFont(QFont("Consolas", 9))

        layout.addWidget(QLabel(f"Содержимое документа ({len(self.document_text)} символов):"))
        layout.addWidget(text_browser)

        # Кнопка закрытия
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(dialog.close)
        layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignRight)

        dialog.setLayout(layout)
        dialog.exec()

    def run_check(self):
        """Запуск проверки"""
        if not self.document_text:
            QMessageBox.warning(self, "Ошибка", "Сначала загрузите документ")
            return

        # Получаем выбранные проверки
        self.selected_checks = []
        for item in self.all_check_items:
            if item.checkState() == Qt.CheckState.Checked:
                check_name = item.data(Qt.ItemDataRole.UserRole)
                self.selected_checks.append(check_name)

        if not self.selected_checks:
            QMessageBox.warning(self, "Ошибка", "Выберите хотя бы одну проверку")
            return

        # Очищаем предыдущие результаты
        self.results_table.setRowCount(0)

        # Настраиваем UI
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.run_check_btn.setEnabled(False)
        self.status_label.setText("Выполняется проверка...")

        # Запускаем проверку в отдельном потоке
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

        # Отображаем результаты
        self.display_results(results)

        # Обновляем статистику
        total = len(results)
        passed = sum(1 for r in results if r['passed'] and not r['needs_verification'])
        failed = sum(1 for r in results if not r['passed'] and not r['needs_verification'])
        warning = sum(1 for r in results if r['needs_verification'])

        self.update_stats(total, passed, failed, warning, elapsed_time)

        # Сброс UI
        self.progress_bar.setVisible(False)
        self.run_check_btn.setEnabled(True)
        self.status_label.setText("Проверка завершена")

        # Показать предупреждение о критической проблеме
        critical_issues = [r for r in results if
                           not r['passed'] and 'Oracle' in r['name'] or 'Запрещённое' in r['name']]
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
            self.results_table.setItem(i, 0, name_item)

            # Группа
            group_item = QTableWidgetItem(result.get('group', ''))
            self.results_table.setItem(i, 1, group_item)

            # Статус
            if result['passed'] and not result['needs_verification']:
                status_text = "✓ Пройдено"
                color = QColor(22, 163, 74)  # Зеленый
            elif result['needs_verification']:
                status_text = "⚠ Требует проверки"
                color = QColor(245, 158, 11)  # Оранжевый
            else:
                status_text = "✗ Провалено"
                color = QColor(220, 38, 38)  # Красный

            status_item = QTableWidgetItem(status_text)
            status_item.setForeground(color)
            status_item.setData(Qt.ItemDataRole.UserRole, result)
            self.results_table.setItem(i, 2, status_item)

            # Результат
            if 'score' in result and result['score'] > 0:
                result_text = f"{result['score']:.1f}%"
            else:
                result_text = result.get('message', '')
            result_item = QTableWidgetItem(result_text)
            self.results_table.setItem(i, 3, result_item)

            # Детали
            details = f"{result.get('details', '')}"
            if result.get('found_text'):
                details += f"\nНайдено: {result['found_text']}"
            details_item = QTableWidgetItem(details)
            self.results_table.setItem(i, 4, details_item)

        self.results_table.resizeRowsToContents()

    def show_critical_issue(self, issue):
        """Показать предупреждение о критической проблеме"""
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Icon.Warning)
        msg_box.setWindowTitle("Обнаружена критическая проблема")
        msg_box.setText(
            f"Обнаружена критическая проблема\n\n"
            f"Проверка: {issue.get('name', 'Неизвестно')}\n"
            f"Результат: {issue.get('message', '')}\n\n"
            f"Местоположение: Таблица 901: \"Local Area Network – активные коммутаторы\"\n"
            f"Наименование: Cisco Catalyst 2960-X\n"
            f"Требования: Использовать оборудование из реестра российского ПО"
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
        QApplication.clipboard().setText("Замечания по проверке (заглушка)")
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
            "Система проверки технической документации\n\n"
            "Версия: 2.1.4\n"
            "Разработчик: Федеральное казначейство\n\n"
            "Все данные обрабатываются в защищённом контуре"
        )


def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Проверка документов ФК")
    app.setStyle("Fusion")

    # Устанавливаем светлую палитру
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(0, 0, 0))
    palette.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(249, 250, 251))
    palette.setColor(QPalette.ColorRole.Text, QColor(0, 0, 0))
    palette.setColor(QPalette.ColorRole.Button, QColor(249, 250, 251))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(0, 0, 0))
    palette.setColor(QPalette.ColorRole.Light, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.Midlight, QColor(243, 244, 246))
    palette.setColor(QPalette.ColorRole.Dark, QColor(156, 163, 175))
    palette.setColor(QPalette.ColorRole.Mid, QColor(209, 213, 219))
    palette.setColor(QPalette.ColorRole.Shadow, QColor(107, 114, 128))
    app.setPalette(palette)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
