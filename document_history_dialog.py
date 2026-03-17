# document_history_dialog.py
"""
Диалог для просмотра истории версий документа и сравнения результатов
С поддержкой комментариев, тегов и отображением информации о ГК
Группирует версии по базовому имени файла
"""

import os
import logging
from datetime import datetime
from typing import List, Dict, Any, Optional

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QHeaderView, QSplitter,
    QWidget, QGroupBox, QMessageBox, QComboBox, QTextEdit,
    QTabWidget, QListWidget, QListWidgetItem, QTreeWidget,
    QTreeWidgetItem, QFrame, QScrollArea, QGridLayout,
    QDialogButtonBox, QMenu, QApplication, QInputDialog,
    QLineEdit, QFormLayout, QCompleter, QToolButton
)
from PyQt6.QtCore import Qt, pyqtSignal, QSize, QStringListModel
from PyQt6.QtGui import QFont, QColor, QIcon, QAction, QPixmap

from document_history import DocumentHistory, DocumentVersion, extract_base_filename

logger = logging.getLogger(__name__)


class CommentDialog(QDialog):
    """Диалог для добавления/просмотра комментариев"""

    def __init__(self, history: DocumentHistory, version_id: str, parent=None):
        super().__init__(parent)
        self.history = history
        self.version_id = version_id
        self.version = history.get_version(version_id)

        self.setWindowTitle(f"Комментарии к версии")
        self.resize(600, 500)

        if parent and hasattr(parent, 'styleSheet'):
            self.setStyleSheet(parent.styleSheet())

        self.init_ui()
        self.load_comments()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Информация о версии
        info_frame = QFrame()
        info_frame.setFrameStyle(QFrame.Shape.Box)
        info_frame.setStyleSheet("background-color: #f0f2f5; border-radius: 5px; padding: 5px;")
        info_layout = QVBoxLayout(info_frame)

        if self.version:
            # Название файла с версией
            if self.version.file_version:
                name_text = f"{self.version.base_filename} (версия {self.version.file_version})"
            else:
                name_text = self.version.full_filename
            info_layout.addWidget(QLabel(f"📄 {name_text}"))

            # ГК
            primary_gk = self.version.primary_gk
            if primary_gk:
                gk_text = f"🔑 {primary_gk.get('number', '')} [{primary_gk.get('subsystem', 'Не определена')}]"
                if primary_gk.get('date'):
                    gk_text += f" от {primary_gk.get('date')}"
                info_layout.addWidget(QLabel(gk_text))

            # Дата создания
            info_layout.addWidget(QLabel(f"📅 {self.version.created_at[:16]}"))

        layout.addWidget(info_frame)

        # Список комментариев
        self.comments_list = QListWidget()
        self.comments_list.setMinimumHeight(300)
        layout.addWidget(QLabel("Существующие комментарии:"))
        layout.addWidget(self.comments_list)

        # Поле для нового комментария
        form_layout = QFormLayout()
        self.comment_input = QTextEdit()
        self.comment_input.setMaximumHeight(80)
        self.comment_input.setPlaceholderText("Введите комментарий...")
        form_layout.addRow("Новый комментарий:", self.comment_input)

        # Кнопки
        button_layout = QHBoxLayout()

        add_btn = QPushButton("➕ Добавить комментарий")
        add_btn.clicked.connect(self.add_comment)
        add_btn.setMinimumHeight(35)

        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        close_btn.setMinimumHeight(35)

        button_layout.addWidget(add_btn)
        button_layout.addStretch()
        button_layout.addWidget(close_btn)

        layout.addLayout(form_layout)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def load_comments(self):
        """Загрузка комментариев"""
        self.comments_list.clear()

        if self.version and self.version.comments:
            for comment in self.version.comments:
                created = comment.get('created_at', '')[:16] if comment.get('created_at') else ''
                author = comment.get('author', 'user')
                text = comment.get('text', '')

                item_text = f"[{created}] {author}:\n{text}"
                item = QListWidgetItem(item_text)
                item.setData(Qt.ItemDataRole.UserRole, comment.get('id'))

                # Разные цвета для разных авторов
                if author == 'system':
                    item.setForeground(QColor(150, 150, 150))
                elif author == 'user':
                    item.setForeground(QColor(0, 100, 200))

                self.comments_list.addItem(item)
        else:
            self.comments_list.addItem("Нет комментариев")

    def add_comment(self):
        """Добавление нового комментария"""
        comment_text = self.comment_input.toPlainText().strip()
        if not comment_text:
            QMessageBox.warning(self, "Ошибка", "Введите текст комментария")
            return

        if self.history.add_comment_to_version(self.version_id, comment_text, "user"):
            self.comment_input.clear()
            # Перезагружаем версию
            self.version = self.history.get_version(self.version_id)
            self.load_comments()
            QMessageBox.information(self, "Успех", "Комментарий добавлен")
        else:
            QMessageBox.warning(self, "Ошибка", "Не удалось добавить комментарий")


class TagManagerDialog(QDialog):
    """Диалог для управления тегами версии"""

    def __init__(self, history: DocumentHistory, version_id: str, parent=None):
        super().__init__(parent)
        self.history = history
        self.version_id = version_id
        self.version = history.get_version(version_id)

        self.setWindowTitle(f"Управление тегами")
        self.resize(400, 300)

        if parent and hasattr(parent, 'styleSheet'):
            self.setStyleSheet(parent.styleSheet())

        self.init_ui()
        self.load_tags()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Информация о версии
        info_label = QLabel(f"Версия: {self.version_id[:20]}...")
        info_label.setStyleSheet("font-weight: bold; color: #3498db;")
        layout.addWidget(info_label)

        if self.version and self.version.primary_gk:
            gk = self.version.primary_gk
            gk_text = f"🔑 {gk.get('number')} [{gk.get('subsystem')}]"
            if gk.get('date'):
                gk_text += f" от {gk.get('date')}"
            layout.addWidget(QLabel(gk_text))

        # Список тегов
        self.tags_list = QListWidget()
        layout.addWidget(QLabel("Текущие теги:"))
        layout.addWidget(self.tags_list)

        # Добавление нового тега
        add_layout = QHBoxLayout()
        self.tag_input = QLineEdit()
        self.tag_input.setPlaceholderText("Введите новый тег...")

        # Автодополнение из существующих тегов
        all_tags = self.history.get_all_tags()
        completer = QCompleter(all_tags)
        completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.tag_input.setCompleter(completer)

        add_btn = QPushButton("➕ Добавить")
        add_btn.clicked.connect(self.add_tag)

        add_layout.addWidget(self.tag_input)
        add_layout.addWidget(add_btn)

        layout.addLayout(add_layout)

        # Кнопки
        button_layout = QHBoxLayout()

        remove_btn = QPushButton("🗑️ Удалить выбранный тег")
        remove_btn.clicked.connect(self.remove_tag)

        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)

        button_layout.addWidget(remove_btn)
        button_layout.addStretch()
        button_layout.addWidget(close_btn)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def load_tags(self):
        """Загрузка тегов"""
        self.tags_list.clear()
        if self.version and self.version.tags:
            for tag in self.version.tags:
                item = QListWidgetItem(f"🏷️ {tag}")
                item.setData(Qt.ItemDataRole.UserRole, tag)
                self.tags_list.addItem(item)
        else:
            self.tags_list.addItem("Нет тегов")

    def add_tag(self):
        """Добавление тега"""
        tag = self.tag_input.text().strip()
        if not tag:
            QMessageBox.warning(self, "Ошибка", "Введите название тега")
            return

        if self.history.add_tag_to_version(self.version_id, tag):
            self.tag_input.clear()
            self.version = self.history.get_version(self.version_id)
            self.load_tags()
            QMessageBox.information(self, "Успех", f"Тег '{tag}' добавлен")
        else:
            QMessageBox.warning(self, "Ошибка", "Не удалось добавить тег")

    def remove_tag(self):
        """Удаление выбранного тега"""
        current = self.tags_list.currentItem()
        if not current:
            QMessageBox.warning(self, "Ошибка", "Выберите тег для удаления")
            return

        tag = current.data(Qt.ItemDataRole.UserRole)
        reply = QMessageBox.question(
            self, "Подтверждение",
            f"Удалить тег '{tag}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            if self.history.remove_tag_from_version(self.version_id, tag):
                self.version = self.history.get_version(self.version_id)
                self.load_tags()
                QMessageBox.information(self, "Успех", f"Тег '{tag}' удален")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось удалить тег")


class VersionComparisonDialog(QDialog):
    """Диалог для сравнения двух версий документа"""

    def __init__(self, history: DocumentHistory, version_id1: str, version_id2: str, parent=None):
        super().__init__(parent)
        self.history = history
        self.version_id1 = version_id1
        self.version_id2 = version_id2
        self.comparison = history.compare_versions(version_id1, version_id2)

        self.setWindowTitle("Сравнение версий документа")
        self.resize(1200, 800)

        if parent and hasattr(parent, 'styleSheet'):
            self.setStyleSheet(parent.styleSheet())

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Верхняя панель с информацией о версиях
        top_panel = self.create_info_panel()
        layout.addWidget(top_panel)

        # Основной контент с вкладками
        tab_widget = QTabWidget()

        # Вкладка со статистикой изменений
        stats_tab = self.create_stats_tab()
        tab_widget.addTab(stats_tab, "📊 Статистика изменений")

        # Вкладка с изменениями в ошибках
        errors_tab = self.create_errors_tab()
        tab_widget.addTab(errors_tab, "❌ Изменения в ошибках")

        # Вкладка с изменениями в тексте
        text_tab = self.create_text_diff_tab()
        tab_widget.addTab(text_tab, "📄 Изменения в тексте")

        layout.addWidget(tab_widget)

        # Кнопка закрытия
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        close_btn.setMinimumHeight(35)
        layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignRight)

        self.setLayout(layout)

    def create_info_panel(self) -> QWidget:
        panel = QWidget()
        panel.setStyleSheet("""
            QWidget {
                background-color: #f0f2f5;
                border-radius: 8px;
                padding: 10px;
            }
        """)

        layout = QHBoxLayout()

        # Информация о первой версии
        v1_box = QGroupBox("Версия 1")
        v1_layout = QVBoxLayout()

        v1_info = self.history.get_version(self.version_id1)
        if v1_info:
            # Название файла с версией
            if v1_info.file_version:
                name_text = f"{v1_info.base_filename}\n(версия {v1_info.file_version})"
            else:
                name_text = v1_info.full_filename
            v1_layout.addWidget(QLabel(f"📄 {name_text}"))

            # ГК для первой версии
            if v1_info.primary_gk:
                gk = v1_info.primary_gk
                gk_text = f"🔑 {gk.get('number')} [{gk.get('subsystem')}]"
                if gk.get('date'):
                    gk_text += f"\n📅 {gk.get('date')}"
                v1_layout.addWidget(QLabel(gk_text))

            v1_layout.addWidget(QLabel(f"📅 {v1_info.created_at[:10]} {v1_info.created_at[11:19]}"))
            v1_layout.addWidget(QLabel(f"📊 Всего проверок: {v1_info.stats['total']}"))
            v1_layout.addWidget(QLabel(f"✅ Пройдено: {v1_info.stats['passed']}"))
            v1_layout.addWidget(QLabel(f"❌ Провалено: {v1_info.stats['failed']}"))
            if v1_info.tags:
                v1_layout.addWidget(QLabel(f"🏷️ Теги: {', '.join(v1_info.tags)}"))

        v1_box.setLayout(v1_layout)

        # Информация о второй версии
        v2_box = QGroupBox("Версия 2")
        v2_layout = QVBoxLayout()

        v2_info = self.history.get_version(self.version_id2)
        if v2_info:
            # Название файла с версией
            if v2_info.file_version:
                name_text = f"{v2_info.base_filename}\n(версия {v2_info.file_version})"
            else:
                name_text = v2_info.full_filename
            v2_layout.addWidget(QLabel(f"📄 {name_text}"))

            # ГК для второй версии
            if v2_info.primary_gk:
                gk = v2_info.primary_gk
                gk_text = f"🔑 {gk.get('number')} [{gk.get('subsystem')}]"
                if gk.get('date'):
                    gk_text += f"\n📅 {gk.get('date')}"
                v2_layout.addWidget(QLabel(gk_text))

            v2_layout.addWidget(QLabel(f"📅 {v2_info.created_at[:10]} {v2_info.created_at[11:19]}"))
            v2_layout.addWidget(QLabel(f"📊 Всего проверок: {v2_info.stats['total']}"))
            v2_layout.addWidget(QLabel(f"✅ Пройдено: {v2_info.stats['passed']}"))
            v2_layout.addWidget(QLabel(f"❌ Провалено: {v2_info.stats['failed']}"))
            if v2_info.tags:
                v2_layout.addWidget(QLabel(f"🏷️ Теги: {', '.join(v2_info.tags)}"))

        v2_box.setLayout(v2_layout)

        layout.addWidget(v1_box)
        layout.addWidget(v2_box)

        panel.setLayout(layout)
        return panel

    def create_stats_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout()

        # Сводка изменений
        summary_group = QGroupBox("Сводка изменений")
        summary_layout = QGridLayout()

        stats_diff = self.comparison.get('stats_diff', {})

        row = 0
        for key, label in [('passed', '✅ Пройдено'),
                           ('failed', '❌ Провалено'),
                           ('needs_verification', '⚠ Требует проверки')]:
            data = stats_diff.get(key, {})
            diff = data.get('diff', 0)

            summary_layout.addWidget(QLabel(f"{label}:"), row, 0)
            summary_layout.addWidget(QLabel(f"{data.get('old', 0)} → {data.get('new', 0)}"), row, 1)

            if diff > 0:
                change_label = QLabel(f"+{diff}")
                change_label.setStyleSheet("color: green; font-weight: bold;")
            elif diff < 0:
                change_label = QLabel(f"{diff}")
                change_label.setStyleSheet("color: red; font-weight: bold;")
            else:
                change_label = QLabel("0")

            summary_layout.addWidget(change_label, row, 2)
            row += 1

        # Изменения в тексте
        text_diff = self.comparison.get('text_diff', {})
        summary_layout.addWidget(QLabel("📄 Изменения в тексте:"), row, 0)
        summary_layout.addWidget(QLabel(f"{text_diff.get('added_lines', 0)} добавлено, "
                                        f"{text_diff.get('removed_lines', 0)} удалено"), row, 1)

        summary_group.setLayout(summary_layout)
        layout.addWidget(summary_group)

        # Таблица изменений
        changes_group = QGroupBox("Детальные изменения")
        changes_layout = QVBoxLayout()

        changes_table = QTableWidget()
        changes_table.setColumnCount(4)
        changes_table.setHorizontalHeaderLabels(["Проверка", "Старый статус", "Новый статус", "Изменение"])

        results_diff = self.comparison.get('results_diff', {})
        changes_table.setRowCount(len(results_diff.get('changed', [])))

        for i, change in enumerate(results_diff.get('changed', [])):
            changes_table.setItem(i, 0, QTableWidgetItem(change.get('name', '')))
            changes_table.setItem(i, 1, QTableWidgetItem(change.get('old_status', '')))
            changes_table.setItem(i, 2, QTableWidgetItem(change.get('new_status', '')))

            change_text = "✅ Исправлено" if "Пройдено" in change.get('new_status', '') else "❌ Появилась"
            change_item = QTableWidgetItem(change_text)
            if "Исправлено" in change_text:
                change_item.setForeground(QColor(0, 150, 0))
            else:
                change_item.setForeground(QColor(255, 0, 0))
            changes_table.setItem(i, 3, change_item)

        changes_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        changes_layout.addWidget(changes_table)

        changes_group.setLayout(changes_layout)
        layout.addWidget(changes_group)

        tab.setLayout(layout)
        return tab

    def create_errors_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout()

        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Исправленные ошибки
        fixed_widget = QWidget()
        fixed_layout = QVBoxLayout(fixed_widget)

        fixed_label = QLabel(f"✅ Исправленные ошибки ({len(self.comparison.get('fixed_errors', []))})")
        fixed_label.setStyleSheet("font-weight: bold; color: green;")
        fixed_layout.addWidget(fixed_label)

        fixed_list = QListWidget()
        for error in self.comparison.get('fixed_errors', []):
            item = QListWidgetItem(f"✓ {error.get('name', '')}")
            item.setForeground(QColor(0, 150, 0))
            fixed_list.addItem(item)

        fixed_layout.addWidget(fixed_list)

        # Новые ошибки
        new_widget = QWidget()
        new_layout = QVBoxLayout(new_widget)

        new_label = QLabel(f"❌ Новые ошибки ({len(self.comparison.get('new_errors', []))})")
        new_label.setStyleSheet("font-weight: bold; color: red;")
        new_layout.addWidget(new_label)

        new_list = QListWidget()
        for error in self.comparison.get('new_errors', []):
            item = QListWidgetItem(f"✗ {error.get('name', '')}")
            item.setForeground(QColor(255, 0, 0))
            new_list.addItem(item)

        new_layout.addWidget(new_list)

        # Неизменные ошибки
        unchanged_widget = QWidget()
        unchanged_layout = QVBoxLayout(unchanged_widget)

        unchanged_label = QLabel(f"⚠ Неизменные ошибки ({len(self.comparison.get('unchanged_errors', []))})")
        unchanged_label.setStyleSheet("font-weight: bold; color: orange;")
        unchanged_layout.addWidget(unchanged_label)

        unchanged_list = QListWidget()
        for error in self.comparison.get('unchanged_errors', []):
            item = QListWidgetItem(f"• {error.get('name', '')}")
            item.setForeground(QColor(255, 140, 0))
            unchanged_list.addItem(item)

        unchanged_layout.addWidget(unchanged_list)

        splitter.addWidget(fixed_widget)
        splitter.addWidget(new_widget)
        splitter.addWidget(unchanged_widget)
        splitter.setSizes([400, 400, 400])

        layout.addWidget(splitter)

        tab.setLayout(layout)
        return tab

    def create_text_diff_tab(self) -> QWidget:
        tab = QWidget()
        layout = QVBoxLayout()

        text_diff = self.comparison.get('text_diff', {})

        info_label = QLabel(f"Изменено строк: {text_diff.get('total_diff_lines', 0)} "
                            f"(+{text_diff.get('added_lines', 0)} / -{text_diff.get('removed_lines', 0)})")
        layout.addWidget(info_label)

        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setFont(QFont("Consolas", 10))

        diff_text = "\n".join(text_diff.get('diff_lines', []))
        text_edit.setPlainText(diff_text)

        layout.addWidget(text_edit)

        tab.setLayout(layout)
        return tab


class VersionDetailsDialog(QDialog):
    """Диалог для просмотра деталей версии и переключения между версиями"""

    version_changed = pyqtSignal(str)  # Сигнал при смене версии

    def __init__(self, history: DocumentHistory, group_key: str, initial_version_id: str = None, parent=None):
        super().__init__(parent)
        self.history = history
        self.group_key = group_key
        self.group_info = history.get_group(group_key)
        self.versions = history.get_group_versions(group_key)
        self.current_version_id = initial_version_id or (self.versions[0]['version_id'] if self.versions else None)

        self.setWindowTitle(f"Детали документа")
        self.resize(1100, 700)

        if parent and hasattr(parent, 'styleSheet'):
            self.setStyleSheet(parent.styleSheet())

        self.init_ui()
        if self.current_version_id:
            self.load_version(self.current_version_id)

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Верхняя панель с информацией о документе
        top_panel = self.create_document_info_panel()
        layout.addWidget(top_panel)

        # Панель выбора версии
        version_panel = self.create_version_selector()
        layout.addWidget(version_panel)

        # Основной контент
        content_splitter = QSplitter(Qt.Orientation.Horizontal)

        # Левая панель - информация о версии
        left_panel = self.create_version_info_panel()
        content_splitter.addWidget(left_panel)

        # Правая панель - результаты проверки
        right_panel = self.create_results_panel()
        content_splitter.addWidget(right_panel)

        content_splitter.setSizes([400, 700])
        layout.addWidget(content_splitter)

        # Кнопки
        button_layout = QHBoxLayout()

        load_btn = QPushButton("📂 Загрузить эту версию")
        load_btn.clicked.connect(self.load_current_version)
        load_btn.setMinimumHeight(35)

        compare_btn = QPushButton("🔄 Сравнить с другой версией")
        compare_btn.clicked.connect(self.compare_with_other)
        compare_btn.setMinimumHeight(35)

        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        close_btn.setMinimumHeight(35)

        button_layout.addWidget(load_btn)
        button_layout.addWidget(compare_btn)
        button_layout.addStretch()
        button_layout.addWidget(close_btn)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def create_document_info_panel(self) -> QWidget:
        panel = QWidget()
        panel.setStyleSheet("""
            QWidget {
                background-color: #f0f2f5;
                border-radius: 8px;
                padding: 10px;
            }
        """)

        layout = QHBoxLayout()

        if self.group_info:
            # Название документа
            name_label = QLabel(f"📄 {self.group_info.get('base_filename', 'Неизвестный документ')}")
            name_label.setStyleSheet("font-weight: bold; font-size: 14px;")
            layout.addWidget(name_label)

            layout.addStretch()

            # Статистика
            versions_count = self.group_info.get('version_count', 0)
            stats_label = QLabel(f"📊 Всего версий: {versions_count}")
            stats_label.setStyleSheet("color: #7f8c8d;")
            layout.addWidget(stats_label)

            # Основной ГК
            primary_gk = self.group_info.get('current_primary_gk', {})
            if primary_gk:
                gk_text = f"🔑 {primary_gk.get('number', '')} [{primary_gk.get('subsystem', 'Не определена')}]"
                if primary_gk.get('date'):
                    gk_text += f" от {primary_gk.get('date')}"
                gk_label = QLabel(gk_text)
                gk_label.setStyleSheet("color: #e67e22;")
                layout.addWidget(gk_label)

        panel.setLayout(layout)
        return panel

    def create_version_selector(self) -> QWidget:
        panel = QWidget()
        layout = QHBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        layout.addWidget(QLabel("Версия:"))

        self.version_combo = QComboBox()
        self.version_combo.setMinimumWidth(300)

        for version in self.versions:
            created = version.get('created_at', '')[:16] if version.get('created_at') else 'Неизвестно'
            file_version = version.get('file_version', '')
            stats = version.get('stats', {})

            if file_version:
                display_text = f"{created} - Версия {file_version} (✅{stats.get('passed', 0)} ❌{stats.get('failed', 0)})"
            else:
                display_text = f"{created} (✅{stats.get('passed', 0)} ❌{stats.get('failed', 0)})"

            self.version_combo.addItem(display_text, version.get('version_id'))

        # Устанавливаем текущую версию
        if self.current_version_id:
            index = self.version_combo.findData(self.current_version_id)
            if index >= 0:
                self.version_combo.setCurrentIndex(index)

        self.version_combo.currentIndexChanged.connect(self.on_version_changed)

        layout.addWidget(self.version_combo)
        layout.addStretch()

        # Кнопки управления
        comments_btn = QPushButton("💬 Комментарии")
        comments_btn.clicked.connect(self.show_comments)
        layout.addWidget(comments_btn)

        tags_btn = QPushButton("🏷️ Теги")
        tags_btn.clicked.connect(self.manage_tags)
        layout.addWidget(tags_btn)

        panel.setLayout(layout)
        return panel

    def create_version_info_panel(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)

        # Информация о версии
        info_group = QGroupBox("Информация о версии")
        info_layout = QVBoxLayout()

        self.version_info_text = QTextEdit()
        self.version_info_text.setReadOnly(True)
        self.version_info_text.setMaximumHeight(200)
        info_layout.addWidget(self.version_info_text)

        info_group.setLayout(info_layout)
        layout.addWidget(info_group)

        # Комментарии
        comments_group = QGroupBox("Последние комментарии")
        comments_layout = QVBoxLayout()

        self.comments_list = QListWidget()
        self.comments_list.setMaximumHeight(150)
        comments_layout.addWidget(self.comments_list)

        comments_group.setLayout(comments_layout)
        layout.addWidget(comments_group)

        # Теги
        tags_group = QGroupBox("Теги")
        tags_layout = QVBoxLayout()

        self.tags_list = QListWidget()
        self.tags_list.setMaximumHeight(100)
        tags_layout.addWidget(self.tags_list)

        tags_group.setLayout(tags_layout)
        layout.addWidget(tags_group)

        layout.addStretch()

        return panel

    def create_results_panel(self) -> QWidget:
        panel = QWidget()
        layout = QVBoxLayout(panel)

        # Результаты проверки
        results_group = QGroupBox("Результаты проверки")
        results_layout = QVBoxLayout()

        self.results_table = QTableWidget()
        self.results_table.setColumnCount(4)
        self.results_table.setHorizontalHeaderLabels(["Проверка", "Статус", "Результат", "Страница"])

        header = self.results_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)

        results_layout.addWidget(self.results_table)

        results_group.setLayout(results_layout)
        layout.addWidget(results_group)

        return panel

    def on_version_changed(self, index: int):
        """Обработка смены версии"""
        if index >= 0:
            version_id = self.version_combo.itemData(index)
            self.current_version_id = version_id
            self.load_version(version_id)
            self.version_changed.emit(version_id)

    def load_version(self, version_id: str):
        """Загрузка и отображение версии"""
        version = self.history.get_version(version_id)

        if not version:
            return

        # Информация о версии
        info_text = f"<h3>Версия: {version_id[:20]}...</h3>"

        if version.file_version:
            info_text += f"<p><b>Версия файла:</b> {version.file_version}</p>"

        info_text += f"<p><b>Дата создания:</b> {version.created_at}</p>"
        info_text += f"<p><b>Файл:</b> {version.full_filename}</p>"
        info_text += f"<p><b>Родительская версия:</b> {version.parent_version_id or '—'}</p>"
        info_text += f"<p><b>Статистика:</b> ✅ {version.stats['passed']} | ❌ {version.stats['failed']} | ⚠ {version.stats['needs_verification']}</p>"

        if version.primary_gk:
            gk = version.primary_gk
            info_text += f"<p><b>Основной ГК:</b> {gk.get('number')} [{gk.get('subsystem')}]"
            if gk.get('date'):
                info_text += f" от {gk.get('date')}"
            info_text += "</p>"

        self.version_info_text.setHtml(info_text)

        # Комментарии
        self.comments_list.clear()
        if version.comments:
            for comment in version.comments[-5:]:  # Последние 5 комментариев
                created = comment.get('created_at', '')[:16]
                text = comment.get('text', '')[:50]
                item = QListWidgetItem(f"[{created}] {text}...")
                self.comments_list.addItem(item)
        else:
            self.comments_list.addItem("Нет комментариев")

        # Теги
        self.tags_list.clear()
        if version.tags:
            for tag in version.tags:
                item = QListWidgetItem(f"🏷️ {tag}")
                self.tags_list.addItem(item)
        else:
            self.tags_list.addItem("Нет тегов")

        # Результаты проверки
        self.results_table.setRowCount(len(version.check_results))

        for i, result in enumerate(version.check_results):
            self.results_table.setItem(i, 0, QTableWidgetItem(result.get('name', '')))

            if result.get('passed', False) and not result.get('needs_verification', False):
                status_item = QTableWidgetItem("✅ Пройдено")
                status_item.setForeground(QColor(0, 150, 0))
            elif result.get('needs_verification', False):
                status_item = QTableWidgetItem("⚠ Требует проверки")
                status_item.setForeground(QColor(255, 140, 0))
            else:
                status_item = QTableWidgetItem("❌ Провалено")
                status_item.setForeground(QColor(255, 0, 0))

            self.results_table.setItem(i, 1, status_item)
            self.results_table.setItem(i, 2, QTableWidgetItem(result.get('message', '')))
            self.results_table.setItem(i, 3, QTableWidgetItem(str(result.get('page', ''))))

        self.results_table.resizeRowsToContents()

    def load_current_version(self):
        """Загрузка текущей версии в основное приложение"""
        if self.current_version_id:
            self.version_changed.emit(self.current_version_id)
            self.accept()

    def compare_with_other(self):
        """Сравнение с другой версией"""
        if len(self.versions) < 2:
            QMessageBox.warning(self, "Ошибка", "Недостаточно версий для сравнения")
            return

        # Диалог выбора версии для сравнения
        items = []
        for version in self.versions:
            if version['version_id'] != self.current_version_id:
                created = version.get('created_at', '')[:16]
                file_version = version.get('file_version', '')
                if file_version:
                    items.append(f"{created} - Версия {file_version}")
                else:
                    items.append(created)

        item, ok = QInputDialog.getItem(
            self, "Выбор версии",
            "Выберите версию для сравнения:",
            items, 0, False
        )

        if ok and item:
            # Находим ID выбранной версии
            idx = items.index(item)
            other_version = [v for v in self.versions if v['version_id'] != self.current_version_id][idx]

            dialog = VersionComparisonDialog(
                self.history,
                self.current_version_id,
                other_version['version_id'],
                self
            )
            dialog.exec()

    def show_comments(self):
        """Показать диалог комментариев"""
        if self.current_version_id:
            dialog = CommentDialog(self.history, self.current_version_id, self)
            dialog.exec()
            # Обновляем отображение
            self.load_version(self.current_version_id)

    def manage_tags(self):
        """Управление тегами"""
        if self.current_version_id:
            dialog = TagManagerDialog(self.history, self.current_version_id, self)
            dialog.exec()
            # Обновляем отображение
            self.load_version(self.current_version_id)


class DocumentGroupWidget(QWidget):
    """Виджет группы документов с выпадающим списком версий"""

    group_selected = pyqtSignal(str, str)  # group_key, version_id
    view_details = pyqtSignal(str, str)  # group_key, version_id

    def __init__(self, history: DocumentHistory, group_info: Dict, parent=None):
        super().__init__(parent)
        self.history = history
        self.group_info = group_info
        self.group_key = group_info.get('group_key')
        self.versions = []

        self.setup_ui()
        self.load_versions()

    def setup_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(5)

        # Заголовок с информацией о документе
        title_layout = QHBoxLayout()

        # Иконка раскрытия
        self.expand_btn = QToolButton()
        self.expand_btn.setArrowType(Qt.ArrowType.RightArrow)
        self.expand_btn.setCheckable(True)
        self.expand_btn.setChecked(False)
        self.expand_btn.toggled.connect(self.toggle_versions)
        self.expand_btn.setStyleSheet("QToolButton { border: none; }")
        title_layout.addWidget(self.expand_btn)

        # Информация о документе
        info_layout = QVBoxLayout()

        # Название файла
        base_filename = self.group_info.get('base_filename', 'Неизвестный документ')
        name_label = QLabel(f"📄 {base_filename}")
        name_label.setStyleSheet("font-weight: bold; font-size: 12px;")
        info_layout.addWidget(name_label)

        # Информация о ГК
        primary_gk = self.group_info.get('current_primary_gk', {})
        if primary_gk:
            gk_text = f"🔑 {primary_gk.get('number', '')} [{primary_gk.get('subsystem', 'Не определена')}]"
            if primary_gk.get('date'):
                gk_text += f" от {primary_gk.get('date')}"

            gk_label = QLabel(gk_text)
            gk_label.setStyleSheet("color: #e67e22; font-size: 11px;")
            info_layout.addWidget(gk_label)

        # Статистика
        versions_count = self.group_info.get('version_count', 0)
        stats_label = QLabel(
            f"📊 Версий: {versions_count} | Последний доступ: {self.group_info.get('last_accessed', '')[:10]}")
        stats_label.setStyleSheet("color: #7f8c8d; font-size: 10px;")
        info_layout.addWidget(stats_label)

        title_layout.addLayout(info_layout)
        title_layout.addStretch()

        # Кнопка просмотра деталей
        view_btn = QPushButton("👁️ Просмотр")
        view_btn.setMaximumWidth(80)
        view_btn.clicked.connect(self.show_details)
        title_layout.addWidget(view_btn)

        layout.addLayout(title_layout)

        # Разделитель
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(separator)

        # Список версий (изначально скрыт)
        self.versions_widget = QWidget()
        self.versions_layout = QVBoxLayout(self.versions_widget)
        self.versions_layout.setContentsMargins(20, 5, 5, 5)
        self.versions_layout.setSpacing(2)
        self.versions_widget.setVisible(False)

        layout.addWidget(self.versions_widget)

        self.setLayout(layout)

        # Стили
        self.setStyleSheet("""
            DocumentGroupWidget {
                border: 1px solid #d1d9e6;
                border-radius: 8px;
                background-color: white;
                margin: 2px;
            }
            DocumentGroupWidget:hover {
                background-color: #f8f9fa;
                border-color: #3498db;
            }
        """)

    def load_versions(self):
        """Загрузка версий группы"""
        self.versions = self.history.get_group_versions(self.group_key)

        # Очищаем список версий
        for i in reversed(range(self.versions_layout.count())):
            widget = self.versions_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)

        if not self.versions:
            label = QLabel("  Нет сохраненных версий")
            label.setStyleSheet("color: #95a5a6; font-style: italic;")
            self.versions_layout.addWidget(label)
            return

        # Добавляем версии
        for version in self.versions:
            version_widget = self.create_version_widget(version)
            self.versions_layout.addWidget(version_widget)

    def create_version_widget(self, version_info: Dict) -> QWidget:
        """Создание виджета для версии"""
        widget = QWidget()
        layout = QHBoxLayout()
        layout.setContentsMargins(5, 5, 5, 5)

        # Информация о версии
        info_layout = QVBoxLayout()

        # Дата и версия
        created = version_info.get('created_at', '')
        date_str = created[:16] if created else 'Неизвестно'
        file_version = version_info.get('file_version', '')

        if file_version:
            title_text = f"📅 {date_str} - Версия {file_version}"
        else:
            title_text = f"📅 {date_str}"

        title_label = QLabel(title_text)
        title_label.setStyleSheet("font-weight: bold; font-size: 11px;")
        info_layout.addWidget(title_label)

        # ГК для этой версии
        primary_gk = version_info.get('primary_gk')
        if primary_gk:
            gk_text = f"🔑 {primary_gk.get('number', '')} [{primary_gk.get('subsystem', 'Не определена')}]"
            if primary_gk.get('date'):
                gk_text += f" от {primary_gk.get('date')}"
            gk_label = QLabel(gk_text)
            gk_label.setStyleSheet("color: #e67e22; font-size: 10px;")
            info_layout.addWidget(gk_label)

        # Статистика
        stats = version_info.get('stats', {})
        stats_text = f"✅ {stats.get('passed', 0)} | ❌ {stats.get('failed', 0)} | ⚠ {stats.get('needs_verification', 0)}"
        stats_label = QLabel(stats_text)
        stats_label.setStyleSheet("font-size: 10px;")
        info_layout.addWidget(stats_label)

        layout.addLayout(info_layout)
        layout.addStretch()

        # Кнопки действий
        btn_layout = QHBoxLayout()

        select_btn = QPushButton("📂 Выбрать")
        select_btn.setMaximumWidth(70)
        select_btn.clicked.connect(lambda: self.group_selected.emit(self.group_key, version_info['version_id']))
        btn_layout.addWidget(select_btn)

        details_btn = QPushButton("👁️")
        details_btn.setMaximumWidth(30)
        details_btn.setToolTip("Просмотр деталей")
        details_btn.clicked.connect(lambda: self.view_details.emit(self.group_key, version_info['version_id']))
        btn_layout.addWidget(details_btn)

        layout.addLayout(btn_layout)

        widget.setLayout(layout)

        # Стили для версии
        widget.setStyleSheet("""
            QWidget {
                border: 1px solid #ecf0f1;
                border-radius: 4px;
                background-color: #f8f9fa;
            }
            QWidget:hover {
                background-color: #e8ecf1;
                border-color: #3498db;
            }
        """)

        return widget

    def toggle_versions(self, checked):
        """Показать/скрыть список версий"""
        self.versions_widget.setVisible(checked)
        self.expand_btn.setArrowType(Qt.ArrowType.DownArrow if checked else Qt.ArrowType.RightArrow)

    def show_details(self):
        """Показать детали документа"""
        if self.versions:
            self.view_details.emit(self.group_key, self.versions[0]['version_id'])


class DocumentHistoryDialog(QDialog):
    """Диалог для просмотра истории документов"""

    version_selected = pyqtSignal(str)  # Сигнал при выборе версии для загрузки

    def __init__(self, history: DocumentHistory, file_path: str = None, parent=None):
        super().__init__(parent)
        self.history = history
        self.current_file_path = file_path
        self.current_groups = []

        self.setWindowTitle("История документов")
        self.resize(1100, 750)

        if parent and hasattr(parent, 'styleSheet'):
            self.setStyleSheet(parent.styleSheet())

        self.init_ui()
        self.load_groups()
        self.load_tags()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Верхняя панель с фильтрами
        top_panel = QWidget()
        top_layout = QHBoxLayout(top_panel)
        top_layout.setContentsMargins(0, 0, 0, 0)

        self.file_label = QLabel("Все документы")
        self.file_label.setStyleSheet("font-weight: bold;")

        # Поиск по тегам
        self.tag_search = QComboBox()
        self.tag_search.addItem("Все теги")
        self.tag_search.setMinimumWidth(150)
        self.tag_search.setPlaceholderText("Фильтр по тегу...")
        self.tag_search.currentTextChanged.connect(self.filter_by_tag)

        # Поиск по ГК
        self.gk_search = QLineEdit()
        self.gk_search.setPlaceholderText("🔍 Поиск по номеру ГК...")
        self.gk_search.setMinimumWidth(200)
        self.gk_search.textChanged.connect(self.filter_by_gk)

        refresh_btn = QPushButton("🔄 Обновить")
        refresh_btn.clicked.connect(self.load_groups)

        top_layout.addWidget(self.file_label)
        top_layout.addStretch()
        top_layout.addWidget(QLabel("Тег:"))
        top_layout.addWidget(self.tag_search)
        top_layout.addWidget(QLabel("ГК:"))
        top_layout.addWidget(self.gk_search)
        top_layout.addWidget(refresh_btn)

        layout.addWidget(top_panel)

        # Область прокрутки для групп документов
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)

        scroll_widget = QWidget()
        self.groups_layout = QVBoxLayout(scroll_widget)
        self.groups_layout.setContentsMargins(5, 5, 5, 5)
        self.groups_layout.setSpacing(10)
        self.groups_layout.addStretch()

        scroll_area.setWidget(scroll_widget)
        layout.addWidget(scroll_area)

        # Кнопка закрытия
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

        self.setLayout(layout)

    def load_tags(self):
        """Загрузка всех тегов для фильтра"""
        self.tag_search.clear()
        self.tag_search.addItem("Все теги")
        all_tags = self.history.get_all_tags()
        self.tag_search.addItems(all_tags)

    def load_groups(self):
        """Загрузка списка групп документов"""
        # Очищаем текущие группы (кроме stretch в конце)
        while self.groups_layout.count() > 1:
            item = self.groups_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        groups = self.history.get_all_groups()
        self.current_groups = groups

        if not groups:
            label = QLabel("  Нет сохраненных документов")
            label.setStyleSheet("color: #95a5a6; font-style: italic; padding: 20px;")
            self.groups_layout.insertWidget(0, label)
            return

        # Создаем виджеты для каждой группы
        for group_info in groups:
            group_widget = DocumentGroupWidget(self.history, group_info, self)
            group_widget.group_selected.connect(self.on_version_selected)
            group_widget.view_details.connect(self.show_group_details)
            self.groups_layout.insertWidget(self.groups_layout.count() - 1, group_widget)

    def filter_by_tag(self, tag: str):
        """Фильтрация групп по тегу"""
        if tag == "Все теги" or not tag:
            # Показываем все группы
            for i in range(self.groups_layout.count() - 1):
                widget = self.groups_layout.itemAt(i).widget()
                if isinstance(widget, DocumentGroupWidget):
                    widget.setVisible(True)
            return

        # Фильтруем по тегу
        for i in range(self.groups_layout.count() - 1):
            widget = self.groups_layout.itemAt(i).widget()
            if isinstance(widget, DocumentGroupWidget):
                group_tags = widget.group_info.get('tags', [])
                widget.setVisible(tag in group_tags)

    def filter_by_gk(self, gk_text: str):
        """Фильтрация групп по номеру ГК"""
        if not gk_text:
            # Показываем все группы
            for i in range(self.groups_layout.count() - 1):
                widget = self.groups_layout.itemAt(i).widget()
                if isinstance(widget, DocumentGroupWidget):
                    widget.setVisible(True)
            return

        gk_upper = gk_text.upper()

        for i in range(self.groups_layout.count() - 1):
            widget = self.groups_layout.itemAt(i).widget()
            if isinstance(widget, DocumentGroupWidget):
                group_info = widget.group_info
                primary_gk = group_info.get('current_primary_gk', {})
                gk_number = primary_gk.get('number', '')

                if gk_upper in gk_number.upper():
                    widget.setVisible(True)
                else:
                    widget.setVisible(False)

    def show_group_details(self, group_key: str, version_id: str):
        """Показать детали группы"""
        dialog = VersionDetailsDialog(self.history, group_key, version_id, self)
        dialog.version_changed.connect(self.on_version_selected)
        dialog.exec()

    def on_version_selected(self, group_key: str, version_id: str):
        """Обработка выбора версии"""
        self.version_selected.emit(version_id)
        self.accept()