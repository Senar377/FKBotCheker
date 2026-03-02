# app_window.py
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QListWidget, QListWidgetItem,
    QProgressBar, QTableWidget, QTableWidgetItem,
    QGroupBox, QMessageBox, QSplitter, QLineEdit,
    QComboBox, QMenu, QFileDialog, QHeaderView, QStatusBar, QSizePolicy
)
from PyQt6.QtCore import Qt, QSettings
from PyQt6.QtGui import QFont, QAction, QColor
import time
import yaml
from datetime import datetime

from document_checker import DocumentChecker
from docx_parser import DOCXParser
from add_check_dialog import AddCheckDialog
from settings_dialog import SettingsDialog
from config_editor import ConfigEditor
from document_viewer import DocumentViewer
from check_worker import CheckWorker
from manage_checks_dialog import ManageChecksDialog
from versions_dialog import VersionsDialog
import logging

logger = logging.getLogger(__name__)


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
        self.page_info = []  # Информация о страницах документа
        self.document_viewer = None

        # Настройки по умолчанию
        self.settings = QSettings("ФедеральноеКазначейство", "ПроверкаДокументов")
        self.theme_mode = self.settings.value("theme_mode", "dark", type=str)
        self.dark_theme = self.theme_mode == "dark"
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
        elif self.theme_mode == "light":
            self.setStyleSheet(self.get_light_theme())
        else:  # mixed
            self.setStyleSheet(self.get_mixed_theme())

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
                color: #34495e;
                min-height: 25px;
            }
            QSpinBox {
                background-color: #2a2a2a;
                border: 1px solid #555555;
                border-radius: 4px;
                color: #ffffff;
                min-height: 25px;
                padding: 2px;
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
            QSpinBox {
                background-color: #ffffff;
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                color: #000000;
                min-height: 25px;
                padding: 2px;
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
        """Смешанная тема"""
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
            QSpinBox {
                background-color: #ffffff;
                border: 1px solid #d1d9e6;
                border-radius: 4px;
                color: #333333;
                min-height: 25px;
                padding: 2px;
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

    def save_settings(self):
        """Сохранить настройки"""
        self.settings.setValue("theme_mode", self.theme_mode)
        self.settings.setValue("dark_theme", self.dark_theme)
        self.settings.setValue("fuzzy_threshold", self.fuzzy_threshold)
        self.settings.setValue("fuzzy_trust_threshold", self.fuzzy_trust_threshold)
        self.settings.setValue("auto_resize_columns", self.auto_resize_columns)
        self.settings.setValue("show_line_numbers", self.show_line_numbers)

    def init_ui(self):
        self.setWindowTitle("Система проверки технической документации - Федеральное казначейство v2.2.0")
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

        self.versions_btn = QPushButton("Версии БПО/СПО")
        self.versions_btn.clicked.connect(self.show_versions_dialog)
        self.versions_btn.setEnabled(False)  # Изначально disabled, пока не загружен документ
        self.versions_btn.setMinimumHeight(36)
        self.versions_btn.setStyleSheet("""
            QPushButton {
                background-color: #9b59b6;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #8e44ad;
            }
            QPushButton:disabled {
                background-color: #7f8c8d;
            }
        """)


        btn_layout.addWidget(self.load_doc_btn)
        btn_layout.addWidget(self.view_doc_btn)
        btn_layout.addWidget(self.view_with_errors_btn)
        btn_layout.addWidget(self.versions_btn)
        btn_layout.addStretch()

        doc_layout.addLayout(btn_layout)
        main_layout.addWidget(doc_panel)

        # ========== РАЗДЕЛИТЕЛЬ ==========
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # Устанавливаем политику размера для разделителя
        splitter.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        # ========== ЛЕВАЯ ПАНЕЛЬ: ПРОВЕРКИ ==========
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(5, 5, 5, 5)
        left_layout.setSpacing(10)

        # Панель управления проверками
        controls_layout = QHBoxLayout()
        controls_layout.setSpacing(10)

        checks_header = QLabel("Выбор проверок")
        checks_header.setStyleSheet("font-weight: bold; font-size: 13px;")
        controls_layout.addWidget(checks_header)

        controls_layout.addStretch()

        # Кнопка добавления новой проверки
        self.add_check_btn = QPushButton("+ Добавить проверку")
        self.add_check_btn.clicked.connect(self.add_new_check)
        self.add_check_btn.setMinimumHeight(30)
        self.add_check_btn.setStyleSheet("""
            QPushButton {
                background-color: #0066cc;
                color: white;
                font-weight: bold;
                border-radius: 5px;
                padding: 5px 10px;
            }
            QPushButton:hover {
                background-color: #0055aa;
            }
        """)
        controls_layout.addWidget(self.add_check_btn)

        left_layout.addLayout(controls_layout)

        self.selection_stats = QLabel("0 из 0 выбрано")
        left_layout.addWidget(self.selection_stats)

        self.checks_list = QListWidget()
        self.checks_list.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.checks_list.itemChanged.connect(self.update_selection_stats)
        self.checks_list.setMinimumWidth(350)

        # Устанавливаем политику размера для списка проверок
        self.checks_list.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        # Контекстное меню для списка проверок
        self.checks_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.checks_list.customContextMenuRequested.connect(self.show_checks_context_menu)

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
        self.results_table.setColumnCount(8)
        self.results_table.setHorizontalHeaderLabels(
            ["Проверка", "Группа", "Статус", "Результат", "Страница", "Позиция", "Детали", "Действие"])

        # Настраиваем заголовки для масштабирования
        header = self.results_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)  # Столбец "Детали" будет растягиваться
        header.setSectionResizeMode(7, QHeaderView.ResizeMode.ResizeToContents)

        # Устанавливаем политику размера для таблицы
        self.results_table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        # Устанавливаем начальные размеры
        self.results_table.setColumnWidth(0, 200)
        self.results_table.setColumnWidth(1, 150)
        self.results_table.setColumnWidth(3, 100)
        self.results_table.setColumnWidth(4, 80)
        self.results_table.setColumnWidth(7, 90)

        self.results_table.setAlternatingRowColors(True)
        self.results_table.setSortingEnabled(True)
        self.results_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        # Контекстное меню для таблицы
        self.results_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.results_table.customContextMenuRequested.connect(self.show_results_context_menu)

        # Двойной клик по строке
        self.results_table.doubleClicked.connect(self.go_to_error_from_table)

        right_layout.addWidget(self.results_table)

        # Добавляем панели в разделитель
        splitter.addWidget(left_panel)
        splitter.addWidget(right_panel)
        splitter.setSizes([400, 1000])

        # Устанавливаем коэффициент растяжения для разделителя
        splitter.setStretchFactor(0, 1)  # Левая панель - минимальное растяжение
        splitter.setStretchFactor(1, 3)  # Правая панель - большее растяжение

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

        # Прогресс-бар также должен растягиваться
        self.progress_bar.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

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
        self.status_bar.showMessage("Федеральное казначейство • Система проверки технической документации • v2.2.0")

        # ========== МЕНЮ ==========
        self.create_menu()


        #============================
        self.manage_checks_btn = QPushButton("📋 Управление проверками")
        self.manage_checks_btn.clicked.connect(self.open_manage_checks)
        self.manage_checks_btn.setMinimumHeight(30)
        controls_layout.addWidget(self.manage_checks_btn)


    def create_menu(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu("Файл")

        load_action = QAction("Загрузить документ", self)
        load_action.triggered.connect(self.load_document)
        load_action.setShortcut("Ctrl+O")
        file_menu.addAction(load_action)

        new_check_action = QAction("Добавить новую проверку", self)
        new_check_action.triggered.connect(self.add_new_check)
        new_check_action.setShortcut("Ctrl+N")
        file_menu.addAction(new_check_action)

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

        manage_checks_action = QAction("Управление проверками", self)
        manage_checks_action.triggered.connect(self.open_manage_checks)
        file_menu.addAction(manage_checks_action)

        manage_checks_action = QAction("📋 Управление проверками", self)
        manage_checks_action.triggered.connect(self.open_manage_checks)
        manage_checks_action.setShortcut("Ctrl+M")
        file_menu.addAction(manage_checks_action)

    def open_manage_checks(self):
        """Открыть диалог управления проверками"""
        dialog = ManageChecksDialog(self, self.checker.config)
        dialog.config_changed.connect(self.update_config)
        dialog.exec()

    def show_checks_context_menu(self, position):
        """Показать контекстное меню для списка проверок"""
        item = self.checks_list.itemAt(position)
        if not item or item.flags() & Qt.ItemFlag.ItemIsUserCheckable == 0:
            return

        menu = QMenu()

        edit_action = QAction("Редактировать проверку", self)
        edit_action.triggered.connect(lambda: self.edit_check(item))

        delete_action = QAction("Удалить проверку", self)
        delete_action.triggered.connect(lambda: self.delete_check(item))

        duplicate_action = QAction("Дублировать проверку", self)
        duplicate_action.triggered.connect(lambda: self.duplicate_check(item))

        menu.addAction(edit_action)
        menu.addAction(duplicate_action)
        menu.addAction(delete_action)

        menu.exec(self.checks_list.viewport().mapToGlobal(position))

    def open_manage_checks(self):
        """Открыть диалог управления проверками"""
        dialog = ManageChecksDialog(self, self.checker.config)
        dialog.config_changed.connect(self.update_config)
        dialog.exec()

    def add_new_check(self):
        """Добавить новую проверку"""
        try:
            dialog = AddCheckDialog(self)
            if dialog.exec():
                check_data = dialog.get_check_data()
                self.add_check_to_config(check_data)
                self.update_checks_list()
                QMessageBox.information(self, "Успех", "Новая проверка добавлена!")
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            QMessageBox.critical(self, "Ошибка",
                                 f"Ошибка при добавлении проверки:\n{str(e)}\n\nПодробности:\n{error_details}")

    def edit_check(self, item):
        """Редактировать существующую проверку"""
        check_name = item.data(Qt.ItemDataRole.UserRole)
        group_name = None

        # Находим группу проверки
        for group in self.checker.config.get('checks', []):
            for subcheck in group.get('subchecks', []):
                if subcheck.get('name') == check_name:
                    group_name = group.get('group', '')
                    edit_data = subcheck.copy()
                    edit_data['group'] = group_name

                    # Открываем диалог редактирования
                    dialog = AddCheckDialog(self, edit_data)
                    if dialog.exec():
                        if hasattr(dialog, 'deleted') and dialog.deleted:
                            # Удаление проверки
                            self.remove_check_from_config(check_name, group_name)
                        else:
                            # Обновление проверки
                            updated_data = dialog.get_check_data()
                            self.update_check_in_config(check_name, group_name, updated_data)

                        self.update_checks_list()
                    return

    def delete_check(self, item):
        """Удалить проверку"""
        check_name = item.data(Qt.ItemDataRole.UserRole)

        reply = QMessageBox.question(
            self, "Удаление",
            f"Вы уверены, что хотите удалить проверку '{check_name}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            # Находим и удаляем проверку
            for group in self.checker.config.get('checks', []):
                subchecks = group.get('subchecks', [])
                for i, subcheck in enumerate(subchecks):
                    if subcheck.get('name') == check_name:
                        subchecks.pop(i)
                        self.update_checks_list()
                        QMessageBox.information(self, "Успех", f"Проверка '{check_name}' удалена")
                        return

    def duplicate_check(self, item):
        """Дублировать проверку"""
        check_name = item.data(Qt.ItemDataRole.UserRole)

        # Находим оригинальную проверку
        for group in self.checker.config.get('checks', []):
            for subcheck in group.get('subchecks', []):
                if subcheck.get('name') == check_name:
                    # Создаем копию с новым именем
                    new_check = subcheck.copy()
                    new_check['name'] = f"{check_name} (копия)"

                    # Добавляем в ту же группу
                    group['subchecks'].append(new_check)

                    self.update_checks_list()
                    QMessageBox.information(self, "Успех", f"Проверка '{check_name}' продублирована")
                    return

    def add_check_to_config(self, check_data):
        """Добавить проверку в конфигурацию"""
        group_name = check_data.pop('group')

        # Находим группу или создаем новую
        target_group = None
        for group in self.checker.config.get('checks', []):
            if group.get('group') == group_name:
                target_group = group
                break

        if not target_group:
            # Создаем новую группу
            target_group = {
                'group': group_name,
                'subchecks': []
            }
            self.checker.config.setdefault('checks', []).append(target_group)

        # Добавляем проверку
        target_group['subchecks'].append(check_data)

    def remove_check_from_config(self, check_name, group_name):
        """Удалить проверку из конфигурации"""
        for group in self.checker.config.get('checks', []):
            if group.get('group') == group_name:
                subchecks = group.get('subchecks', [])
                for i, subcheck in enumerate(subchecks):
                    if subcheck.get('name') == check_name:
                        subchecks.pop(i)

                        # Если группа пустая, удаляем ее
                        if not subchecks:
                            self.checker.config['checks'].remove(group)
                        return

    def update_check_in_config(self, old_name, old_group, new_data):
        """Обновить проверку в конфигурации"""
        # Удаляем старую версию
        self.remove_check_from_config(old_name, old_group)

        # Добавляем обновленную версию
        self.add_check_to_config(new_data)

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
                description = subcheck.get('description', '')

                # Формируем текст элемента
                item_text = f"• {check_name}"
                if description:
                    item_text += f"\n   {description[:50]}..." if len(description) > 50 else f"\n   {description}"

                item = QListWidgetItem(item_text)
                item.setFlags(
                    Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled | Qt.ItemFlag.ItemIsSelectable)
                item.setCheckState(Qt.CheckState.Unchecked)
                item.setData(Qt.ItemDataRole.UserRole, check_name)
                item.setData(Qt.ItemDataRole.UserRole + 1, check_type)

                # Добавляем подсказку
                if description:
                    item.setToolTip(description)

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
            from pathlib import Path
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

            # Разбиваем документ на страницы
            self.calculate_page_info()

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
            self.versions_btn.setEnabled(True)
            self.status_label.setText(f"Документ загружен: {file_name}")

            self.results_table.setRowCount(0)
            self.last_results = []
            self.update_stats(0, 0, 0, 0, 0)

            logger.info(f"Документ загружен: {file_name}, размер: {file_size:.2f} МБ, страниц: {len(self.page_info)}")

        except Exception as e:
            logger.error(f"Ошибка загрузки документа: {str(e)}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить документ:\n{str(e)}")

    def calculate_page_info(self):
        """Вычисление информации о страницах документа"""
        self.page_info = []

        if not self.document_text:
            return

        lines = self.document_text.split('\n')
        current_page_text = []
        current_chars = 0
        current_start = 0

        # Оптимальные настройки для страницы
        chars_per_page = 1800
        max_lines_per_page = 50

        for line in lines:
            line_length = len(line)
            line_end = current_start + line_length

            if (current_chars + line_length > chars_per_page and current_chars > 0) or \
                    (len(current_page_text) >= max_lines_per_page) or \
                    (line.strip() == '' and len(current_page_text) > 30):

                page_text = '\n'.join(current_page_text)
                page_end = current_start + len(page_text)
                self.page_info.append((current_start, page_end))

                current_page_text = [line]
                current_chars = line_length
                current_start = page_end + 1
            else:
                current_page_text.append(line)
                current_chars += line_length + 1

        if current_page_text:
            page_text = '\n'.join(current_page_text)
            page_end = current_start + len(page_text)
            self.page_info.append((current_start, page_end))

        logger.info(f"Документ разбит на {len(self.page_info)} страниц")

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

        from PyQt6.QtWidgets import QTextBrowser
        text_browser = QTextBrowser()

        # Добавляем номера строк если включено
        if self.show_line_numbers:
            lines = self.document_text.split('\n')
            numbered_text = ""
            for i, line in enumerate(lines[:1000]):
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
        """Просмотр документа с подсветкой ТОЛЬКО ошибок"""
        if not self.document_text:
            QMessageBox.warning(self, "Ошибка", "Нет загруженного документа")
            return

        if not self.last_results:
            QMessageBox.warning(self, "Ошибка", "Нет результатов проверки для отображения")
            return

        # Фильтруем результаты: только ошибки и те, что требуют проверки
        error_results = []
        for result in self.last_results:
            # Проверяем различные условия для определения ошибок
            is_error = result.get('is_error', False)
            needs_check = result.get('needs_verification', False)
            passed = result.get('passed', False)

            # Ошибка если:
            # 1. Явно помечено как ошибка
            # 2. Проверка провалена И не требует дополнительной проверки
            # 3. Требует проверки (неопределенный статус)
            if is_error or (not passed and not needs_check) or needs_check:
                error_results.append(result)

        if not error_results:
            QMessageBox.information(self, "Нет ошибок",
                                    "В документе не найдено ошибок или проверок, требующих внимания.")
            return

        self.document_viewer = DocumentViewer(self, self.document_text, error_results)
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

        logger.info(f"Запуск проверки: {len(self.selected_checks)} проверок выбрано")

        self.worker = CheckWorker(self.checker, self.document_text, self.selected_checks, self.checker.config,
                                  self.page_info)
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
            self.resize_table_columns()

        critical_issues = [r for r in results if
                           not r['passed'] and ('Oracle' in r['name'] or 'Запрещённое' in r['name'])]
        if critical_issues:
            self.show_critical_issue(critical_issues[0] if critical_issues else None)

        logger.info(
            f"Проверка завершена за {elapsed_time:.2f} секунд. Результаты: {passed} пройдено, {failed} провалено, {warning} требует проверки")

    def update_stats(self, total, passed, failed, warning, elapsed_time):
        """Обновление статистики"""
        self.total_label.setText(f"Всего: {total}")
        self.passed_label.setText(f"Пройдено: {passed}")
        self.failed_label.setText(f"Провалено: {failed}")
        self.warning_label.setText(f"Проверить: {warning}")
        self.time_label.setText(f"Время: {elapsed_time:.1f}с")

    def display_results(self, results):
        """Отображение результатов в таблице"""
        self.last_results = results
        logger.info(f"Отображается {len(results)} результатов")

        for i, result in enumerate(results):
            logger.debug(f"Результат {i}: {result.get('name')}, passed={result.get('passed')}, "
                         f"is_error={result.get('is_error')}, needs_verification={result.get('needs_verification')}")

        self.results_table.setRowCount(len(results))

        for i, result in enumerate(results):
            # Название проверки
            name_item = QTableWidgetItem(result['name'])
            name_item.setData(Qt.ItemDataRole.UserRole, result)
            self.results_table.setItem(i, 0, name_item)

            # Группа
            group_item = QTableWidgetItem(result.get('group', ''))
            self.results_table.setItem(i, 1, group_item)

            # Статус - УЛУЧШЕННЫЙ
            if result['passed'] and not result.get('needs_verification', False):
                status_text = "✓ Пройдено"
                color = "#00cc66"  # Зеленый
            elif result.get('needs_verification', False):
                status_text = "⚠ Требует проверки"
                color = "#ffcc00"  # Желтый
            elif result.get('is_error', False):
                status_text = "✗ Ошибка"
                color = "#ff5050"  # Красный
            else:
                status_text = "✗ Провалено"
                color = "#ff5050"  # Красный

            status_item = QTableWidgetItem(status_text)
            status_item.setForeground(QColor(color))
            self.results_table.setItem(i, 2, status_item)

            # Результат
            if 'score' in result and result['score'] > 0:
                result_text = f"{result['score']:.1f}%"
            else:
                result_text = result.get('message', '')
            result_item = QTableWidgetItem(result_text)
            self.results_table.setItem(i, 3, result_item)

            # Страница - УЛУЧШЕННОЕ ОТОБРАЖЕНИЕ
            page_text = str(result.get('page', '')) if result.get('page') else ""
            page_item = QTableWidgetItem(page_text)
            self.results_table.setItem(i, 4, page_item)

            # Позиция - УЛУЧШЕННОЕ ОТОБРАЖЕНИЕ
            position_text = result.get('position', '')
            if not position_text and result.get('line_number'):
                position_text = f"Строка {result['line_number']}"
            position_item = QTableWidgetItem(position_text)
            self.results_table.setItem(i, 5, position_item)

            # Детали - УЛУЧШЕННЫЕ С УЧЕТОМ VERSION_COMPARISON
            if result['type'] == 'version_comparison':
                details = result.get('details', '')
                if result.get('section_results'):
                    details += "\n\nРезультаты по показателям:\n"
                    for sr in result.get('section_results', [])[:5]:  # Показываем первые 5
                        details += f"  {sr.get('result', '')}\n"
                    if len(result.get('section_results', [])) > 5:
                        details += f"  ... и еще {len(result.get('section_results', [])) - 5} показателей"
            else:
                details = f"{result.get('details', '')}"
                if result.get('found_text'):
                    details += f"\nНайдено: {result['found_text']}"
                if result.get('context'):
                    details += f"\nКонтекст: {result['context'][:100]}..."

            details_item = QTableWidgetItem(details)
            self.results_table.setItem(i, 6, details_item)

            # Кнопка для перехода к ошибке - ТОЛЬКО ДЛЯ ОШИБОК И ПРОВЕРОК
            is_error_or_verification = result.get('is_error', False) or result.get('needs_verification', False)
            if result.get('matches') and is_error_or_verification:
                btn = QPushButton("Перейти")
                btn.setProperty('row', i)
                btn.setProperty('result', result)
                btn.clicked.connect(lambda checked, r=i, res=result: self.go_to_error_in_viewer(r, res))
                btn.setMaximumWidth(80)
                btn.setMinimumHeight(25)
                self.results_table.setCellWidget(i, 7, btn)
            else:
                self.results_table.setItem(i, 7, QTableWidgetItem(""))

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

    def go_to_error_in_viewer(self, row, result):
        """Перейти к ошибке в просмотре документа"""
        if not self.document_text or not self.last_results:
            QMessageBox.warning(self, "Ошибка", "Нет документа или результатов проверки")
            return

        # Проверяем, открыт ли уже просмотр с ошибками
        if not self.document_viewer or not self.document_viewer.isVisible():
            self.document_viewer = DocumentViewer(self, self.document_text, self.last_results)

        self.document_viewer.show()
        self.document_viewer.raise_()
        self.document_viewer.activateWindow()

        if result.get('page'):
            self.document_viewer.go_to_page(result['page'])

        self.document_viewer.show_all_errors()

    def show_versions_dialog(self):
        """Показать диалог с версиями БПО/СПО"""
        if not self.document_text:
            QMessageBox.warning(self, "Ошибка", "Сначала загрузите документ")
            return

        from versions_dialog import VersionsDialog
        dialog = VersionsDialog(self, self.document_text, self.page_info)
        dialog.exec()

    def resize_table_columns(self):
        """Автоматически изменить размер столбцов таблицы"""
        self.results_table.resizeColumnsToContents()
        self.results_table.setColumnWidth(7, 90)

    def filter_results_table(self):
        """Фильтрация таблицы результатов"""
        search_text = self.results_search_input.text().lower()
        filter_type = self.filter_combo.currentText()

        for row in range(self.results_table.rowCount()):
            show_row = True

            if search_text:
                row_text = ""
                for col in range(self.results_table.columnCount()):
                    item = self.results_table.item(row, col)
                    if item:
                        row_text += item.text().lower() + " "

                if search_text not in row_text:
                    show_row = False

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

    def show_results_context_menu(self, position):
        """Показать контекстное меню для таблицы результатов"""
        menu = QMenu()

        view_details_action = QAction("Просмотреть детали", self)
        view_details_action.triggered.connect(self.view_selected_result_details)

        go_to_error_action = QAction("Перейти к ошибке в документе", self)
        go_to_error_action.triggered.connect(self.go_to_selected_error)

        copy_row_action = QAction("Копировать строку", self)
        copy_row_action.triggered.connect(self.copy_selected_row)

        resize_columns_action = QAction("Автоматически изменить размер столбцов", self)
        resize_columns_action.triggered.connect(self.resize_table_columns)

        menu.addAction(view_details_action)
        menu.addAction(go_to_error_action)
        menu.addAction(copy_row_action)
        menu.addSeparator()
        menu.addAction(resize_columns_action)

        menu.exec(self.results_table.viewport().mapToGlobal(position))

    def view_selected_result_details(self):
        """Просмотреть детали выбранного результата"""
        selected_rows = self.results_table.selectionModel().selectedRows()
        if not selected_rows:
            return

        row = selected_rows[0].row()
        result_item = self.results_table.item(row, 0)
        if not result_item:
            return

        result = result_item.data(Qt.ItemDataRole.UserRole)

        if result:
            from PyQt6.QtWidgets import QTextEdit
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
            details += f"<p><b>Страница:</b> {result.get('page', '')}</p>"
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
        result_item = self.results_table.item(row, 0)
        if not result_item:
            return

        result = result_item.data(Qt.ItemDataRole.UserRole)
        if result:
            self.go_to_error_in_viewer(row, result)

    def go_to_error_from_table(self, index):
        """Перейти к ошибке по двойному клику в таблице"""
        row = index.row()
        result_item = self.results_table.item(row, 0)
        if not result_item:
            return

        result = result_item.data(Qt.ItemDataRole.UserRole)
        if result:
            self.go_to_error_in_viewer(row, result)

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

        from PyQt6.QtWidgets import QApplication
        QApplication.clipboard().setText('\t'.join(text_parts))

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
        from PyQt6.QtWidgets import QApplication
        notes = "Замечания по проверке:\n\n"
        for result in self.last_results:
            if not result['passed'] or result['needs_verification']:
                status = "ТРЕБУЕТ ПРОВЕРКИ" if result['needs_verification'] else "ПРОВАЛЕНО"
                notes += f"{result['name']} ({result['group']}) - {status}\n"
                notes += f"Результат: {result['message']}\n"
                if result.get('page'):
                    notes += f"Страница: {result['page']}\n"
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

    def open_settings(self):
        """Открыть диалог настроек"""
        dialog = SettingsDialog(self)
        if dialog.exec():
            for check_group in self.checker.config.get('checks', []):
                for subcheck in check_group.get('subchecks', []):
                    if subcheck.get('type', '').startswith('fuzzy'):
                        subcheck['threshold'] = self.fuzzy_threshold
                        subcheck['trust_threshold'] = self.fuzzy_trust_threshold

            QMessageBox.information(self, "Настройки", "Настройки применены")

    def show_about(self):
        """Показать информацию о программе"""
        QMessageBox.about(
            self,
            "О программе",
            "<h3>Система проверки технической документации</h3>"
            "<p><b>Версия:</b> 2.0.7</p>"
            "<p><b>Разработчик:</b> Федеральное казначейство, Кашапов Арсен УИИ</p>"
            "<p><b>Библиотеки:</b> PyQt6, RapidFuzz, PyYAML</p>"
            "<p><b>Описание:</b> Система для автоматической проверки технической документации "
            "на соответствие требованиям импортозамещения, функциональным требованиям "
            "и стандартам Федерального казначейства.</p>"
            "<hr>"
            "<p><i>Все данные обрабатываются в защищённом контуре</i></p>"
        )