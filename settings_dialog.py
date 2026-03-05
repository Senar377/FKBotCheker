# settings_dialog.py
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QComboBox,
    QLineEdit, QCheckBox, QDialogButtonBox,
    QHBoxLayout, QMessageBox, QGroupBox,
    QSpinBox, QTabWidget, QWidget, QFileDialog,
    QPushButton, QRadioButton, QButtonGroup,
    QTextEdit, QDateEdit
)
from PyQt6.QtCore import Qt, QSettings, QDate
import os
import re
from datetime import datetime


class SettingsDialog(QDialog):
    """Диалог настроек"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent = parent
        self.setWindowTitle("Настройки")
        self.resize(750, 700)

        # Загружаем настройки
        self.settings = QSettings("ФедеральноеКазначейство", "Settings")

        self.init_ui()
        self.load_settings()

    def init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(15)

        # Создаем вкладки
        tab_widget = QTabWidget()

        # Вкладка "Общие"
        general_tab = self.create_general_tab()
        tab_widget.addTab(general_tab, "📋 Общие")

        # Вкладка "Поиск и сравнение"
        search_tab = self.create_search_tab()
        tab_widget.addTab(search_tab, "🔍 Поиск и сравнение")

        # Вкладка "ГК"
        gk_tab = self.create_gk_tab()
        tab_widget.addTab(gk_tab, "🔑 ГК")

        # Вкладка "Таблицы"
        tables_tab = self.create_tables_tab()
        tab_widget.addTab(tables_tab, "📊 Таблицы")

        # Вкладка "Интерфейс"
        interface_tab = self.create_interface_tab()
        tab_widget.addTab(interface_tab, "🎨 Интерфейс")

        # Вкладка "Excel"
        excel_tab = self.create_excel_tab()
        tab_widget.addTab(excel_tab, "📊 Excel")

        # Вкладка "Проверка файлов"
        file_check_tab = self.create_file_check_tab()
        tab_widget.addTab(file_check_tab, "📁 Проверка файлов")

        layout.addWidget(tab_widget)

        # Кнопки
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)

        layout.addWidget(button_box)

        self.setLayout(layout)

    def create_general_tab(self):
        """Создание вкладки общих настроек"""
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)

        # Группа "Сортировка"
        sort_group = QGroupBox("Сортировка")
        sort_layout = QVBoxLayout()

        self.sort_by_subsystem_check = QCheckBox("Автоматическая сортировка по подсистемам")
        self.sort_by_subsystem_check.setToolTip(
            "При включении результаты будут автоматически группироваться по подсистемам")
        sort_layout.addWidget(self.sort_by_subsystem_check)

        subsystem_order_layout = QHBoxLayout()
        subsystem_order_layout.addWidget(QLabel("Приоритет подсистем:"))
        self.subsystem_order_input = QLineEdit()
        self.subsystem_order_input.setPlaceholderText("ГМП, ГАСУ, ПОИ, ПУДС, ПУиО, ПИАО, ЕПБС, ПУР, НСИ")
        subsystem_order_layout.addWidget(self.subsystem_order_input)
        sort_layout.addLayout(subsystem_order_layout)

        sort_group.setLayout(sort_layout)
        layout.addWidget(sort_group)

        # Группа "Фильтрация"
        filter_group = QGroupBox("Фильтрация")
        filter_layout = QVBoxLayout()

        self.hide_invalid_versions_check = QCheckBox("Скрывать невалидные версии")
        self.hide_invalid_versions_check.setToolTip("Не показывать версии, которые не удалось распарсить")
        filter_layout.addWidget(self.hide_invalid_versions_check)

        self.show_only_matched_check = QCheckBox("Показывать только найденные продукты")
        self.show_only_matched_check.setToolTip("Скрывать продукты, не найденные в документе")
        filter_layout.addWidget(self.show_only_matched_check)

        filter_group.setLayout(filter_layout)
        layout.addWidget(filter_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def create_search_tab(self):
        """Создание вкладки настроек поиска"""
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)

        # Группа "Нечеткий поиск"
        fuzzy_group = QGroupBox("Нечеткий поиск")
        fuzzy_layout = QVBoxLayout()

        threshold_layout = QHBoxLayout()
        threshold_layout.addWidget(QLabel("Порог проверки (0-100):"))
        self.threshold_input = QLineEdit()
        self.threshold_input.setMaximumWidth(80)
        threshold_layout.addWidget(self.threshold_input)
        threshold_layout.addWidget(QLabel("%"))
        threshold_layout.addStretch()
        fuzzy_layout.addLayout(threshold_layout)

        trust_layout = QHBoxLayout()
        trust_layout.addWidget(QLabel("Порог доверия (0-100):"))
        self.trust_threshold_input = QLineEdit()
        self.trust_threshold_input.setMaximumWidth(80)
        trust_layout.addWidget(self.trust_threshold_input)
        trust_layout.addWidget(QLabel("%"))
        trust_layout.addStretch()
        fuzzy_layout.addLayout(trust_layout)

        self.enable_fuzzy_check = QCheckBox("Включить нечеткий поиск")
        self.enable_fuzzy_check.setToolTip("Искать похожие названия, а не только точные совпадения")
        fuzzy_layout.addWidget(self.enable_fuzzy_check)

        fuzzy_group.setLayout(fuzzy_layout)
        layout.addWidget(fuzzy_group)

        # Группа "Версии"
        version_group = QGroupBox("Обработка версий")
        version_layout = QVBoxLayout()

        self.auto_clean_versions_check = QCheckBox("Автоматически очищать версии от мусора")
        version_layout.addWidget(self.auto_clean_versions_check)

        min_version_layout = QHBoxLayout()
        min_version_layout.addWidget(QLabel("Минимальная длина версии:"))
        self.min_version_spin = QSpinBox()
        self.min_version_spin.setMinimum(1)
        self.min_version_spin.setMaximum(10)
        self.min_version_spin.setValue(3)
        min_version_layout.addWidget(self.min_version_spin)
        min_version_layout.addStretch()
        version_layout.addLayout(min_version_layout)

        version_group.setLayout(version_layout)
        layout.addWidget(version_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def create_gk_tab(self):
        """Создание вкладки настроек ГК"""
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)

        # Группа "Поиск ГК"
        gk_search_group = QGroupBox("Поиск ГК")
        gk_search_layout = QVBoxLayout()

        pages_layout = QHBoxLayout()
        pages_layout.addWidget(QLabel("Искать ГК на первых N страницах:"))
        self.gk_pages_spin = QSpinBox()
        self.gk_pages_spin.setMinimum(1)
        self.gk_pages_spin.setMaximum(100)
        self.gk_pages_spin.setValue(3)
        pages_layout.addWidget(self.gk_pages_spin)
        pages_layout.addStretch()
        gk_search_layout.addLayout(pages_layout)

        self.auto_extract_gk_check = QCheckBox("Автоматически извлекать ГК при загрузке документа")
        gk_search_layout.addWidget(self.auto_extract_gk_check)

        # Формат ГК
        gk_format_layout = QHBoxLayout()
        gk_format_layout.addWidget(QLabel("Формат ГК:"))
        self.gk_format_input = QLineEdit()
        self.gk_format_input.setText("ФКУ\\d{3,4}(?:[/-]\\d{2,4})?(?:/\\w+)?")
        self.gk_format_input.setToolTip("Регулярное выражение для поиска ГК")
        gk_format_layout.addWidget(self.gk_format_input)
        gk_search_layout.addLayout(gk_format_layout)

        gk_search_group.setLayout(gk_search_layout)
        layout.addWidget(gk_search_group)

        # Группа "Сортировка по ГК"
        gk_sort_group = QGroupBox("Сортировка по ГК")
        gk_sort_layout = QVBoxLayout()

        self.enable_gk_sort_check = QCheckBox("Включить функциональную сортировку по ГК")
        self.enable_gk_sort_check.setToolTip("Сортировать результаты по соответствию ГК из документа")
        gk_sort_layout.addWidget(self.enable_gk_sort_check)

        # Тип сортировки
        sort_type_group = QButtonGroup(self)

        self.sort_by_match_radio = QRadioButton("Сначала те, у которых ГК совпадает с документом")
        self.sort_by_match_radio.setChecked(True)
        sort_type_group.addButton(self.sort_by_match_radio)
        gk_sort_layout.addWidget(self.sort_by_match_radio)

        self.sort_by_presence_radio = QRadioButton("Сначала те, у которых есть ГК (без учета совпадения)")
        sort_type_group.addButton(self.sort_by_presence_radio)
        gk_sort_layout.addWidget(self.sort_by_presence_radio)

        gk_sort_group.setLayout(gk_sort_layout)
        layout.addWidget(gk_sort_group)

        # Группа "Приоритет ГК"
        gk_priority_group = QGroupBox("Приоритет ГК")
        gk_priority_layout = QVBoxLayout()

        self.priority_from_doc_check = QCheckBox("Приоритет у ГК, найденных в документе")
        self.priority_from_doc_check.setChecked(True)
        gk_priority_layout.addWidget(self.priority_from_doc_check)

        self.priority_from_first_pages_check = QCheckBox("Повышенный приоритет у ГК с первых страниц")
        self.priority_from_first_pages_check.setChecked(True)
        gk_priority_layout.addWidget(self.priority_from_first_pages_check)

        gk_priority_group.setLayout(gk_priority_layout)
        layout.addWidget(gk_priority_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def create_tables_tab(self):
        """Создание вкладки настроек таблиц"""
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)

        # Группа "Определение таблиц"
        table_detect_group = QGroupBox("Определение таблиц в документе")
        table_detect_layout = QVBoxLayout()

        self.detect_tables_check = QCheckBox("Автоматически определять таблицы в документе")
        self.detect_tables_check.setChecked(True)
        table_detect_layout.addWidget(self.detect_tables_check)

        # Минимальное количество столбцов
        min_cols_layout = QHBoxLayout()
        min_cols_layout.addWidget(QLabel("Минимальное количество столбцов в таблице:"))
        self.min_table_cols_spin = QSpinBox()
        self.min_table_cols_spin.setMinimum(2)
        self.min_table_cols_spin.setMaximum(20)
        self.min_table_cols_spin.setValue(3)
        min_cols_layout.addWidget(self.min_table_cols_spin)
        min_cols_layout.addStretch()
        table_detect_layout.addLayout(min_cols_layout)

        # Минимальное количество строк
        min_rows_layout = QHBoxLayout()
        min_rows_layout.addWidget(QLabel("Минимальное количество строк в таблице:"))
        self.min_table_rows_spin = QSpinBox()
        self.min_table_rows_spin.setMinimum(2)
        self.min_table_rows_spin.setMaximum(50)
        self.min_table_rows_spin.setValue(3)
        min_rows_layout.addWidget(self.min_table_rows_spin)
        min_rows_layout.addStretch()
        table_detect_layout.addLayout(min_rows_layout)

        table_detect_group.setLayout(table_detect_layout)
        layout.addWidget(table_detect_group)

        # Группа "Извлечение данных из таблиц"
        table_extract_group = QGroupBox("Извлечение данных из таблиц")
        table_extract_layout = QVBoxLayout()

        self.extract_from_tables_check = QCheckBox("Извлекать версии и ГК из таблиц")
        self.extract_from_tables_check.setChecked(True)
        table_extract_layout.addWidget(self.extract_from_tables_check)

        # Приоритет данных из таблиц
        self.table_priority_check = QCheckBox("Данные из таблиц имеют приоритет над текстом")
        self.table_priority_check.setChecked(True)
        table_extract_layout.addWidget(self.table_priority_check)

        table_extract_group.setLayout(table_extract_layout)
        layout.addWidget(table_extract_group)

        # Группа "Форматы таблиц"
        table_formats_group = QGroupBox("Форматы таблиц")
        table_formats_layout = QVBoxLayout()

        # Разделители столбцов
        separators_layout = QHBoxLayout()
        separators_layout.addWidget(QLabel("Разделители столбцов:"))
        self.table_separators_input = QLineEdit()
        self.table_separators_input.setText("|, \\t,  ,")
        separators_layout.addWidget(self.table_separators_input)
        table_formats_layout.addLayout(separators_layout)

        # Признаки заголовков
        headers_layout = QHBoxLayout()
        headers_layout.addWidget(QLabel("Признаки заголовков:"))
        self.table_headers_input = QLineEdit()
        self.table_headers_input.setText("наименование, продукт, версия, гк, описание")
        headers_layout.addWidget(self.table_headers_input)
        table_formats_layout.addLayout(headers_layout)

        table_formats_group.setLayout(table_formats_layout)
        layout.addWidget(table_formats_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def create_interface_tab(self):
        """Создание вкладки настроек интерфейса"""
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)

        # Группа "Тема"
        theme_group = QGroupBox("Тема оформления")
        theme_layout = QVBoxLayout()

        theme_layout.addWidget(QLabel("Выберите цветовую тему:"))

        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Темная", "Светлая", "Смешанная", "Системная"])
        theme_layout.addWidget(self.theme_combo)

        theme_group.setLayout(theme_layout)
        layout.addWidget(theme_group)

        # Группа "Таблицы"
        table_group = QGroupBox("Таблицы")
        table_layout = QVBoxLayout()

        self.auto_resize_check = QCheckBox("Автоматически изменять размер столбцов")
        self.auto_resize_check.setToolTip("Автоматически подгонять ширину столбцов под содержимое")
        table_layout.addWidget(self.auto_resize_check)

        self.alternating_rows_check = QCheckBox("Чередовать цвет строк")
        self.alternating_rows_check.setToolTip("Использовать разные цвета для четных и нечетных строк")
        table_layout.addWidget(self.alternating_rows_check)

        self.show_grid_check = QCheckBox("Показывать сетку таблицы")
        table_layout.addWidget(self.show_grid_check)

        table_group.setLayout(table_layout)
        layout.addWidget(table_group)

        # Группа "Текстовый просмотр"
        text_group = QGroupBox("Текстовый просмотр")
        text_layout = QVBoxLayout()

        self.show_line_numbers_check = QCheckBox("Показывать номера строк")
        text_layout.addWidget(self.show_line_numbers_check)

        self.word_wrap_check = QCheckBox("Переносить длинные строки")
        text_layout.addWidget(self.word_wrap_check)

        font_size_layout = QHBoxLayout()
        font_size_layout.addWidget(QLabel("Размер шрифта:"))
        self.font_size_spin = QSpinBox()
        self.font_size_spin.setMinimum(8)
        self.font_size_spin.setMaximum(24)
        self.font_size_spin.setValue(10)
        font_size_layout.addWidget(self.font_size_spin)
        font_size_layout.addStretch()
        text_layout.addLayout(font_size_layout)

        text_group.setLayout(text_layout)
        layout.addWidget(text_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def create_excel_tab(self):
        """Создание вкладки настроек Excel"""
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)

        # Группа "Путь к Excel файлу"
        file_group = QGroupBox("Файл с составом ПО")
        file_layout = QVBoxLayout()

        path_layout = QHBoxLayout()
        path_layout.addWidget(QLabel("Путь к Excel файлу:"))

        self.excel_path_input = QLineEdit()
        self.excel_path_input.setReadOnly(True)
        self.excel_path_input.setPlaceholderText("Файл не выбран")
        path_layout.addWidget(self.excel_path_input)

        self.browse_btn = QPushButton("📁 Обзор")
        self.browse_btn.clicked.connect(self.browse_excel_file)
        path_layout.addWidget(self.browse_btn)

        file_layout.addLayout(path_layout)

        self.file_info_label = QLabel("Файл не выбран")
        self.file_info_label.setStyleSheet("color: #7f8c8d; font-style: italic;")
        file_layout.addWidget(self.file_info_label)

        file_group.setLayout(file_layout)
        layout.addWidget(file_group)

        # Группа "Автозагрузка"
        auto_group = QGroupBox("Автозагрузка")
        auto_layout = QVBoxLayout()

        self.auto_load_excel_check = QCheckBox("Автоматически загружать Excel файл при запуске")
        auto_layout.addWidget(self.auto_load_excel_check)

        self.auto_reload_check = QCheckBox("Автоматически перезагружать при изменении файла")
        auto_layout.addWidget(self.auto_reload_check)

        auto_group.setLayout(auto_layout)
        layout.addWidget(auto_group)

        # Группа "Сохранение"
        save_group = QGroupBox("Сохранение")
        save_layout = QVBoxLayout()

        self.save_history_check = QCheckBox("Сохранять историю загрузок")
        save_layout.addWidget(self.save_history_check)

        max_history_layout = QHBoxLayout()
        max_history_layout.addWidget(QLabel("Максимум записей в истории:"))
        self.max_history_spin = QSpinBox()
        self.max_history_spin.setMinimum(5)
        self.max_history_spin.setMaximum(50)
        self.max_history_spin.setValue(20)
        max_history_layout.addWidget(self.max_history_spin)
        max_history_layout.addStretch()
        save_layout.addLayout(max_history_layout)

        save_group.setLayout(save_layout)
        layout.addWidget(save_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def create_file_check_tab(self):
        """Создание вкладки проверки файлов"""
        tab = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(15)

        # Группа "Проверка имени файла"
        filename_group = QGroupBox("Проверка имени файла")
        filename_layout = QVBoxLayout()

        self.check_filename_check = QCheckBox("Автоматически проверять имя файла при загрузке")
        self.check_filename_check.setChecked(True)
        filename_layout.addWidget(self.check_filename_check)

        filename_pattern_layout = QHBoxLayout()
        filename_pattern_layout.addWidget(QLabel("Шаблон имени файла:"))
        self.filename_pattern_input = QLineEdit()
        self.filename_pattern_input.setText(
            r"\d+\.\d+\.\d+,\d+\(\d+,\d+;\d+,\d+;\d+\.\d+\)\.ПД\.\d+-\d+\.\d+ \d+\(\d+,\d+\)")
        self.filename_pattern_input.setToolTip("Регулярное выражение для проверки имени файла")
        filename_pattern_layout.addWidget(self.filename_pattern_input)
        filename_layout.addLayout(filename_pattern_layout)

        self.filename_example_label = QLabel("Пример: 30275697.23.01,00(02,00;03,00;04.00).ПД.023-01.01 1(4,8)")
        self.filename_example_label.setStyleSheet("color: #7f8c8d; font-style: italic;")
        filename_layout.addWidget(self.filename_example_label)

        filename_group.setLayout(filename_layout)
        layout.addWidget(filename_group)

        # Группа "Результаты проверки"
        results_group = QGroupBox("Отображение результатов")
        results_layout = QVBoxLayout()

        self.show_filename_warning_check = QCheckBox("Показывать предупреждение при несоответствии имени файла")
        self.show_filename_warning_check.setChecked(True)
        results_layout.addWidget(self.show_filename_warning_check)

        results_group.setLayout(results_layout)
        layout.addWidget(results_group)

        layout.addStretch()
        tab.setLayout(layout)
        return tab

    def browse_excel_file(self):
        """Выбор Excel файла"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл с составом ПО", "",
            "Excel files (*.xlsx *.xls);;All files (*.*)"
        )

        if file_path:
            self.excel_path_input.setText(file_path)
            self.update_file_info(file_path)

    def update_file_info(self, file_path):
        """Обновление информации о файле"""
        if os.path.exists(file_path):
            file_size = os.path.getsize(file_path) / 1024
            mod_time = os.path.getmtime(file_path)
            mod_date = datetime.fromtimestamp(mod_time).strftime("%d.%m.%Y %H:%M")

            self.file_info_label.setText(
                f"Размер: {file_size:.1f} KB | Изменен: {mod_date}"
            )
            self.file_info_label.setStyleSheet("color: green;")
        else:
            self.file_info_label.setText("Файл не найден")
            self.file_info_label.setStyleSheet("color: red;")

    def load_settings(self):
        """Загрузка сохраненных настроек"""
        # Общие настройки
        self.sort_by_subsystem_check.setChecked(
            self.settings.value("sort_by_subsystem", True, type=bool)
        )
        self.subsystem_order_input.setText(
            self.settings.value("subsystem_order", "ГМП, ГАСУ, ПОИ, ПУДС, ПУиО, ПИАО, ЕПБС, ПУР, НСИ")
        )
        self.hide_invalid_versions_check.setChecked(
            self.settings.value("hide_invalid_versions", False, type=bool)
        )
        self.show_only_matched_check.setChecked(
            self.settings.value("show_only_matched", False, type=bool)
        )

        # Настройки поиска
        self.threshold_input.setText(
            self.settings.value("fuzzy_threshold", "60")
        )
        self.trust_threshold_input.setText(
            self.settings.value("fuzzy_trust_threshold", "80")
        )
        self.enable_fuzzy_check.setChecked(
            self.settings.value("enable_fuzzy", True, type=bool)
        )
        self.auto_clean_versions_check.setChecked(
            self.settings.value("auto_clean_versions", True, type=bool)
        )
        self.min_version_spin.setValue(
            self.settings.value("min_version_length", 3, type=int)
        )

        # Настройки ГК
        self.gk_pages_spin.setValue(
            self.settings.value("gk_pages", 3, type=int)
        )
        self.auto_extract_gk_check.setChecked(
            self.settings.value("auto_extract_gk", True, type=bool)
        )
        self.gk_format_input.setText(
            self.settings.value("gk_format", "ФКУ\\d{3,4}(?:[/-]\\d{2,4})?(?:/\\w+)?")
        )
        self.enable_gk_sort_check.setChecked(
            self.settings.value("enable_gk_sort", True, type=bool)
        )

        sort_type = self.settings.value("gk_sort_type", "match")
        if sort_type == "match":
            self.sort_by_match_radio.setChecked(True)
        else:
            self.sort_by_presence_radio.setChecked(True)

        self.priority_from_doc_check.setChecked(
            self.settings.value("priority_from_doc", True, type=bool)
        )
        self.priority_from_first_pages_check.setChecked(
            self.settings.value("priority_from_first_pages", True, type=bool)
        )

        # Настройки таблиц
        self.detect_tables_check.setChecked(
            self.settings.value("detect_tables", True, type=bool)
        )
        self.min_table_cols_spin.setValue(
            self.settings.value("min_table_cols", 3, type=int)
        )
        self.min_table_rows_spin.setValue(
            self.settings.value("min_table_rows", 3, type=int)
        )
        self.extract_from_tables_check.setChecked(
            self.settings.value("extract_from_tables", True, type=bool)
        )
        self.table_priority_check.setChecked(
            self.settings.value("table_priority", True, type=bool)
        )
        self.table_separators_input.setText(
            self.settings.value("table_separators", "|, \\t,  ,")
        )
        self.table_headers_input.setText(
            self.settings.value("table_headers", "наименование, продукт, версия, гк, описание")
        )

        # Настройки интерфейса
        theme = self.settings.value("theme", "Темная")
        self.theme_combo.setCurrentText(theme)

        self.auto_resize_check.setChecked(
            self.settings.value("auto_resize_columns", True, type=bool)
        )
        self.alternating_rows_check.setChecked(
            self.settings.value("alternating_rows", True, type=bool)
        )
        self.show_grid_check.setChecked(
            self.settings.value("show_grid", True, type=bool)
        )
        self.show_line_numbers_check.setChecked(
            self.settings.value("show_line_numbers", True, type=bool)
        )
        self.word_wrap_check.setChecked(
            self.settings.value("word_wrap", True, type=bool)
        )
        self.font_size_spin.setValue(
            self.settings.value("font_size", 10, type=int)
        )

        # Настройки Excel
        excel_path = self.settings.value("excel_file_path", "")
        if excel_path:
            self.excel_path_input.setText(excel_path)
            self.update_file_info(excel_path)

        self.auto_load_excel_check.setChecked(
            self.settings.value("auto_load_excel", True, type=bool)
        )
        self.auto_reload_check.setChecked(
            self.settings.value("auto_reload_excel", False, type=bool)
        )
        self.save_history_check.setChecked(
            self.settings.value("save_history", True, type=bool)
        )
        self.max_history_spin.setValue(
            self.settings.value("max_history", 20, type=int)
        )

        # Настройки проверки файлов
        self.check_filename_check.setChecked(
            self.settings.value("check_filename", True, type=bool)
        )
        self.filename_pattern_input.setText(
            self.settings.value("filename_pattern",
                                r"\d+\.\d+\.\d+,\d+\(\d+,\d+;\d+,\d+;\d+\.\d+\)\.ПД\.\d+-\d+\.\d+ \d+\(\d+,\d+\)")
        )
        self.show_filename_warning_check.setChecked(
            self.settings.value("show_filename_warning", True, type=bool)
        )

    def save_settings(self):
        """Сохранение настроек"""
        # Общие настройки
        self.settings.setValue("sort_by_subsystem", self.sort_by_subsystem_check.isChecked())
        self.settings.setValue("subsystem_order", self.subsystem_order_input.text())
        self.settings.setValue("hide_invalid_versions", self.hide_invalid_versions_check.isChecked())
        self.settings.setValue("show_only_matched", self.show_only_matched_check.isChecked())

        # Настройки поиска
        self.settings.setValue("fuzzy_threshold", self.threshold_input.text())
        self.settings.setValue("fuzzy_trust_threshold", self.trust_threshold_input.text())
        self.settings.setValue("enable_fuzzy", self.enable_fuzzy_check.isChecked())
        self.settings.setValue("auto_clean_versions", self.auto_clean_versions_check.isChecked())
        self.settings.setValue("min_version_length", self.min_version_spin.value())

        # Настройки ГК
        self.settings.setValue("gk_pages", self.gk_pages_spin.value())
        self.settings.setValue("auto_extract_gk", self.auto_extract_gk_check.isChecked())
        self.settings.setValue("gk_format", self.gk_format_input.text())
        self.settings.setValue("enable_gk_sort", self.enable_gk_sort_check.isChecked())

        sort_type = "match" if self.sort_by_match_radio.isChecked() else "presence"
        self.settings.setValue("gk_sort_type", sort_type)

        self.settings.setValue("priority_from_doc", self.priority_from_doc_check.isChecked())
        self.settings.setValue("priority_from_first_pages", self.priority_from_first_pages_check.isChecked())

        # Настройки таблиц
        self.settings.setValue("detect_tables", self.detect_tables_check.isChecked())
        self.settings.setValue("min_table_cols", self.min_table_cols_spin.value())
        self.settings.setValue("min_table_rows", self.min_table_rows_spin.value())
        self.settings.setValue("extract_from_tables", self.extract_from_tables_check.isChecked())
        self.settings.setValue("table_priority", self.table_priority_check.isChecked())
        self.settings.setValue("table_separators", self.table_separators_input.text())
        self.settings.setValue("table_headers", self.table_headers_input.text())

        # Настройки интерфейса
        self.settings.setValue("theme", self.theme_combo.currentText())
        self.settings.setValue("auto_resize_columns", self.auto_resize_check.isChecked())
        self.settings.setValue("alternating_rows", self.alternating_rows_check.isChecked())
        self.settings.setValue("show_grid", self.show_grid_check.isChecked())
        self.settings.setValue("show_line_numbers", self.show_line_numbers_check.isChecked())
        self.settings.setValue("word_wrap", self.word_wrap_check.isChecked())
        self.settings.setValue("font_size", self.font_size_spin.value())

        # Настройки Excel
        self.settings.setValue("excel_file_path", self.excel_path_input.text())
        self.settings.setValue("auto_load_excel", self.auto_load_excel_check.isChecked())
        self.settings.setValue("auto_reload_excel", self.auto_reload_check.isChecked())
        self.settings.setValue("save_history", self.save_history_check.isChecked())
        self.settings.setValue("max_history", self.max_history_spin.value())

        # Настройки проверки файлов
        self.settings.setValue("check_filename", self.check_filename_check.isChecked())
        self.settings.setValue("filename_pattern", self.filename_pattern_input.text())
        self.settings.setValue("show_filename_warning", self.show_filename_warning_check.isChecked())

    def accept(self):
        """Применить настройки"""
        try:
            # Проверка порогов
            threshold = float(self.threshold_input.text())
            trust_threshold = float(self.trust_threshold_input.text())

            if threshold >= trust_threshold:
                QMessageBox.warning(
                    self, "Ошибка",
                    "Порог проверки должен быть меньше порога доверия"
                )
                return

            if threshold < 0 or threshold > 100 or trust_threshold < 0 or trust_threshold > 100:
                QMessageBox.warning(
                    self, "Ошибка",
                    "Пороги должны быть в диапазоне от 0 до 100"
                )
                return

            # Проверка регулярного выражения для ГК
            try:
                if self.auto_extract_gk_check.isChecked():
                    re.compile(self.gk_format_input.text())
            except re.error as e:
                QMessageBox.warning(
                    self, "Ошибка",
                    f"Некорректное регулярное выражение для ГК: {e}"
                )
                return

            # Проверка регулярного выражения для имени файла
            try:
                if self.check_filename_check.isChecked():
                    re.compile(self.filename_pattern_input.text())
            except re.error as e:
                QMessageBox.warning(
                    self, "Ошибка",
                    f"Некорректное регулярное выражение для имени файла: {e}"
                )
                return

            # Сохраняем настройки
            self.save_settings()

            # Применяем к родительскому окну
            if self.parent:
                # Тема
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

                # Пороги
                self.parent.fuzzy_threshold = threshold
                self.parent.fuzzy_trust_threshold = trust_threshold

                # Настройки интерфейса
                self.parent.auto_resize_columns = self.auto_resize_check.isChecked()
                self.parent.show_line_numbers = self.show_line_numbers_check.isChecked()

                # Путь к Excel файлу
                self.parent.excel_file_path = self.excel_path_input.text()

                # Применяем тему
                self.parent.apply_theme()

            super().accept()

        except ValueError:
            QMessageBox.warning(self, "Ошибка", "Пороги должны быть числами")


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

            # Проверяем наличие разделителей
            for sep in separators:
                if sep and sep in line:
                    # Считаем количество колонок
                    if sep == '  ':
                        cols = [col.strip() for col in re.split(r'\s{2,}', line) if col.strip()]
                    elif sep == '\\t':
                        cols = [col.strip() for col in line.split('\t') if col.strip()]
                    else:
                        cols = [col.strip() for col in line.split(sep) if col.strip()]

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
                    'raw_columns': line.split(sep) if sep else []
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
        first_row_text = first_row['text'].lower()

        for keyword in header_keywords:
            if keyword in first_row_text:
                return first_row

        # Проверяем вторую строку
        if len(table_rows) > 1:
            second_row = table_rows[1]
            second_row_text = second_row['text'].lower()
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
                start_idx = table['rows'].index(headers) + 1 if headers in table['rows'] else 0

            # Определяем индексы колонок
            name_col_idx = -1
            version_col_idx = -1
            gk_col_idx = -1

            if headers:
                for i, col in enumerate(headers['columns']):
                    col_lower = col.lower()
                    if 'наименование' in col_lower or 'продукт' in col_lower or 'по' in col_lower:
                        name_col_idx = i
                    elif 'версия' in col_lower or 'вер' in col_lower:
                        version_col_idx = i
                    elif 'гк' in col_lower or 'контракт' in col_lower:
                        gk_col_idx = i

            # Извлекаем данные из строк таблицы
            for row in table['rows'][start_idx:]:
                row_data = {
                    'name': '',
                    'version': '',
                    'gk': [],
                    'line_num': row['line_num'],
                    'source': 'table'
                }

                if name_col_idx >= 0 and name_col_idx < len(row['columns']):
                    row_data['name'] = row['columns'][name_col_idx]

                if version_col_idx >= 0 and version_col_idx < len(row['columns']):
                    row_data['version'] = row['columns'][version_col_idx]

                if gk_col_idx >= 0 and gk_col_idx < len(row['columns']):
                    gk_text = row['columns'][gk_col_idx]
                    # Извлекаем ГК из текста
                    gk_pattern = self.settings.value("gk_format", "ФКУ\\d{3,4}(?:[/-]\\d{2,4})?(?:/\\w+)?")
                    gk_matches = re.findall(gk_pattern, gk_text, re.IGNORECASE)
                    row_data['gk'] = [gk.upper() for gk in gk_matches]

                if row_data['name'] or row_data['version'] or row_data['gk']:
                    versions.append(row_data)

        return versions


class FileChecker:
    """Класс для проверки файлов"""

    def __init__(self):
        self.settings = QSettings("ФедеральноеКазначейство", "Settings")

    def check_filename(self, filename):
        """Проверка имени файла по шаблону"""
        if not self.settings.value("check_filename", True, type=bool):
            return True, "Проверка отключена"

        pattern = self.settings.value(
            "filename_pattern",
            r"\d+\.\d+\.\d+,\d+\(\d+,\d+;\d+,\d+;\d+\.\d+\)\.ПД\.\d+-\d+\.\d+ \d+\(\d+,\d+\)"
        )

        try:
            if re.match(pattern, filename):
                return True, "Имя файла соответствует шаблону"
            else:
                return False, "Имя файла не соответствует шаблону"
        except Exception as e:
            return False, f"Ошибка проверки: {e}"

    def get_filename_info(self, filename):
        """Получение информации из имени файла"""
        info = {}

        # Извлекаем номер документа
        doc_num_match = re.search(r'(\d+)\.\d+\.\d+,\d+', filename)
        if doc_num_match:
            info['document_number'] = doc_num_match.group(1)

        # Извлекаем версию
        version_match = re.search(r'(\d+\.\d+,\d+)\((\d+,\d+;\d+,\d+;\d+\.\d+)\)', filename)
        if version_match:
            info['main_version'] = version_match.group(1)
            info['sub_versions'] = version_match.group(2)

        # Извлекаем код ПД
        pd_match = re.search(r'ПД\.(\d+-\d+\.\d+)', filename)
        if pd_match:
            info['pd_code'] = pd_match.group(1)

        # Извлекаем номер и часть
        num_part_match = re.search(r'(\d+)\((\d+,\d+)\)$', filename)
        if num_part_match:
            info['document_part'] = num_part_match.group(1)
            info['section'] = num_part_match.group(2)

        return info

    def validate_file(self, file_path):
        """Полная проверка файла"""
        results = {
            'filename_valid': False,
            'filename_message': '',
            'filename_info': {}
        }

        # Проверка имени файла
        filename = os.path.basename(file_path)
        is_valid, message = self.check_filename(filename)
        results['filename_valid'] = is_valid
        results['filename_message'] = message
        results['filename_info'] = self.get_filename_info(filename)

        return results