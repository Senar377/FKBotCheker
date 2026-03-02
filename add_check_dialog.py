# add_check_dialog.py - ИСПРАВЛЕННЫЙ КОД С ПРОКРУТКОЙ И ЕДИНОЙ ЦВЕТОВОЙ СХЕМОЙ
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QLineEdit, QComboBox,
    QTextEdit, QDialogButtonBox, QPushButton, QFormLayout,
    QWidget, QHBoxLayout, QMessageBox, QScrollArea, QSpinBox,
    QCheckBox, QGroupBox, QFrame, QSizePolicy
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QColor, QFont, QPalette
import re


class AddCheckDialog(QDialog):
    """Диалог для добавления новой проверки"""

    def __init__(self, parent=None, edit_data=None):
        super().__init__(parent)
        self.edit_data = edit_data
        self.is_edit_mode = edit_data is not None
        self.deleted = False
        self.version_sections_widgets = []  # Список виджетов секций версий
        self.combined_sections_widgets = []  # Список виджетов условий комбинированных проверок

        # Словарь для хранения виджетов и их меток
        self.field_widgets = {}
        self.label_widgets = {}

        if self.is_edit_mode:
            self.setWindowTitle("Редактировать проверку")
        else:
            self.setWindowTitle("Добавить новую проверку")

        self.resize(900, 700)
        self.setMinimumSize(800, 600)

        # Установка единой цветовой схемы
        self.setup_theme()

        self.init_ui()

    def setup_theme(self):
        """Настройка единой цветовой схемы"""
        # Темная тема для лучшего контраста
        dark_theme = """
            QDialog {
                background-color: #2c3e50;
            }
            QLabel {
                color: #34495e;
                font-size: 13px;
            }
            QLabel[color="red"] {
                color: #e74c3c;
            }
            QLabel[color="green"] {
                color: #2ecc71;
            }
            QLineEdit, QTextEdit, QSpinBox {
                background-color: #34495e;
                color: #ecf0f1;
                border: 2px solid #7f8c8d;
                border-radius: 6px;
                padding: 8px;
                font-size: 13px;
                selection-background-color: #3498db;
            }
            QComboBox {
                background-color: #34495e;
                color: #ffffff;  
                border: 2px solid #7f8c8d;
                border-radius: 6px;
                padding: 8px;
                font-size: 13px;
                selection-background-color: #3498db;
            }
            QComboBox QAbstractItemView {
                background-color: #34495e;
                color: #ffffff;  
                border: 2px solid #7f8c8d;
                selection-background-color: #3498db;
                selection-color: #ffffff;
            }
            QComboBox:focus {
                border-color: #3498db;
            }
            QComboBox:disabled {
                background-color: #2c3e50;
                color: #95a5a6;
            }
            QLineEdit:focus, QTextEdit:focus, QSpinBox:focus {
                border-color: #3498db;
            }
            QLineEdit:disabled, QTextEdit:disabled, QSpinBox:disabled {
                background-color: #2c3e50;
                color: #95a5a6;
            }
            QPushButton {
                background-color: #3498db;
                color: white;
                font-weight: bold;
                border: none;
                border-radius: 6px;
                padding: 10px 20px;
                font-size: 13px;
                min-height: 35px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #21618c;
            }
            QPushButton:disabled {
                background-color: #7f8c8d;
                color: #bdc3c7;
            }
            QPushButton#delete_button {
                background-color: #e74c3c;
            }
            QPushButton#delete_button:hover {
                background-color: #c0392b;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 14px;
                color: #ecf0f1;
                border: 2px solid #3498db;
                border-radius: 10px;
                margin-top: 15px;
                padding-top: 20px;
                background-color: #34495e;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 0 10px 0 10px;
                background-color: #3498db;
                color: white;
                border-radius: 5px;
                font-weight: bold;
            }
            QGroupBox#combined_group {
                border-color: #9b59b6;
                background-color: #34495e;
            }
            QGroupBox#combined_group::title {
                background-color: #9b59b6;
            }
            QCheckBox {
                color: #ecf0f1;
                font-size: 13px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
            QScrollArea {
                border: 2px solid #7f8c8d;
                border-radius: 8px;
                background-color: #2c3e50;
            }
            QScrollBar:vertical {
                border: none;
                background: #34495e;
                width: 12px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background: #7f8c8d;
                border-radius: 6px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background: #95a5a6;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: none;
                background: none;
            }
        """

        self.setStyleSheet(dark_theme)

    def init_ui(self):
        # Основной контейнер с прокруткой
        main_scroll = QScrollArea()
        main_scroll.setWidgetResizable(True)
        main_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        main_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        # Основной виджет для содержимого
        content_widget = QWidget()
        content_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        main_layout = QVBoxLayout(content_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(15)

        # Заголовок
        title_label = QLabel("Добавление новой проверки" if not self.is_edit_mode else "Редактирование проверки")
        title_label.setStyleSheet("""
            font-size: 18px; 
            font-weight: bold; 
            color: #ecf0f1;
            padding: 15px;
            background-color: #3498db;
            border-radius: 8px;
            qproperty-alignment: 'AlignCenter';
        """)
        main_layout.addWidget(title_label)

        # Форма с полями
        form_widget = QWidget()
        form_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.MinimumExpanding)
        form_layout = QVBoxLayout(form_widget)
        form_layout.setSpacing(12)
        form_layout.setContentsMargins(10, 10, 10, 10)

        # === ОСНОВНЫЕ ПОЛЯ ===
        self.create_basic_fields(form_layout)

        # === УСЛОВНО ВИДИМЫЕ ПОЛЯ ===
        self.create_conditional_fields(form_layout)

        # === СЕКЦИИ ВЕРСИЙ ===
        self.create_version_sections(form_layout)

        # === СЕКЦИИ КОМБИНИРОВАННЫХ ПРОВЕРОК ===
        self.create_combined_sections(form_layout)

        # === ОПИСАНИЕ ===
        self.create_description_field(form_layout)

        main_layout.addWidget(form_widget)

        # Подсказка
        self.create_hint_label(main_layout)

        # Кнопки
        self.create_buttons(main_layout)

        # Устанавливаем виджет в скролл
        main_scroll.setWidget(content_widget)

        # Основной layout диалога
        dialog_layout = QVBoxLayout(self)
        dialog_layout.setContentsMargins(0, 0, 0, 0)
        dialog_layout.addWidget(main_scroll)

        # Заполняем данные если режим редактирования
        if self.is_edit_mode:
            self.fill_form_data()

        # Обновляем видимость полей
        self.on_type_changed(self.type_combo.currentText())

    def create_basic_fields(self, form_layout):
        """Создание основных полей формы"""
        # Название проверки
        name_layout = QHBoxLayout()
        name_label = QLabel("Название проверки:")
        name_label.setMinimumWidth(200)
        self.name_input = QLineEdit()
        self.name_input.setPlaceholderText("Например: Проверка на Oracle")
        name_layout.addWidget(name_label)
        name_layout.addWidget(self.name_input)
        form_layout.addLayout(name_layout)

        self.field_widgets['name'] = self.name_input
        self.label_widgets['name'] = name_label

        # Группа проверок
        group_layout = QHBoxLayout()
        group_label = QLabel("Группа проверок:")
        group_label.setMinimumWidth(200)
        self.group_input = QComboBox()
        self.group_input.setEditable(True)
        self.group_input.setPlaceholderText("Например: Импортозамещение")
        self.group_input.addItems([
            "Импортозамещение",
            "Функциональные требования",
            "СОБИ ФК",
            "Безопасность",
            "Производительность",
            "Документация",
            "Показатели назначения",
            "Комбинированные проверки"
        ])
        group_layout.addWidget(group_label)
        group_layout.addWidget(self.group_input)
        form_layout.addLayout(group_layout)

        self.field_widgets['group'] = self.group_input
        self.label_widgets['group'] = group_label

        # Тип проверки
        type_layout = QHBoxLayout()
        type_label = QLabel("Тип проверки:")
        type_label.setMinimumWidth(200)
        self.type_combo = QComboBox()
        self.type_combo.addItems([
            "no_text_present - Текст не должен присутствовать",
            "text_present - Текст должен присутствовать",
            "text_present_without - Текст должен быть, а другой - нет",
            "fuzzy_text_present - Нечеткое соответствие текста",
            "no_fuzzy_text_present - Нечеткое несоответствие текста",
            "text_present_in_any_table - Текст в таблицах",
            "fuzzy_text_present_after_any_table - Текст после таблиц",
            "version_comparison - Проверка показателей назначения (версии)",
            "combined_check - Комбинированная проверка (логическое И/ИЛИ)"
        ])
        self.type_combo.currentTextChanged.connect(self.on_type_changed)
        type_layout.addWidget(type_label)
        type_layout.addWidget(self.type_combo)
        form_layout.addLayout(type_layout)

        self.field_widgets['type'] = self.type_combo
        self.label_widgets['type'] = type_label

    def create_conditional_fields(self, form_layout):
        """Создание условно видимых полей"""
        # Основной текст для поиска
        self.text_container = QWidget()
        text_layout = QHBoxLayout(self.text_container)
        text_layout.setContentsMargins(0, 0, 0, 0)
        self.text_label = QLabel("Текст для поиска:")
        self.text_label.setMinimumWidth(200)
        self.text_input = QTextEdit()
        self.text_input.setMaximumHeight(100)
        self.text_input.setPlaceholderText("Текст для поиска (для нечеткого поиска)")
        text_layout.addWidget(self.text_label)
        text_layout.addWidget(self.text_input)
        form_layout.addWidget(self.text_container)
        self.text_container.setVisible(False)

        self.field_widgets['text'] = self.text_input
        self.label_widgets['text'] = self.text_label

        # Алиасы
        self.aliases_container = QWidget()
        aliases_layout = QHBoxLayout(self.aliases_container)
        aliases_layout.setContentsMargins(0, 0, 0, 0)
        self.aliases_label = QLabel("Алиасы (через запятую):")
        self.aliases_label.setMinimumWidth(200)
        self.aliases_input = QTextEdit()
        self.aliases_input.setMaximumHeight(100)
        self.aliases_input.setPlaceholderText(
            "Варианты написания через запятую\nНапример: Oracle, Oracle DB, Oracle Database")
        aliases_layout.addWidget(self.aliases_label)
        aliases_layout.addWidget(self.aliases_input)
        form_layout.addWidget(self.aliases_container)
        self.aliases_container.setVisible(False)

        self.field_widgets['aliases'] = self.aliases_input
        self.label_widgets['aliases'] = self.aliases_label

        # Исключающие алиасы
        self.without_container = QWidget()
        without_layout = QHBoxLayout(self.without_container)
        without_layout.setContentsMargins(0, 0, 0, 0)
        self.without_label = QLabel("Исключающие алиасы:")
        self.without_label.setMinimumWidth(200)
        self.without_aliases_input = QTextEdit()
        self.without_aliases_input.setMaximumHeight(100)
        self.without_aliases_input.setPlaceholderText(
            "Текст, который НЕ должен присутствовать\nНапример: PostgreSQL, MySQL")
        without_layout.addWidget(self.without_label)
        without_layout.addWidget(self.without_aliases_input)
        form_layout.addWidget(self.without_container)
        self.without_container.setVisible(False)

        self.field_widgets['without_aliases'] = self.without_aliases_input
        self.label_widgets['without_aliases'] = self.without_label

        # Пороги для нечеткого поиска
        self.thresholds_container = QWidget()
        thresholds_layout = QHBoxLayout(self.thresholds_container)
        thresholds_layout.setContentsMargins(0, 0, 0, 0)
        self.thresholds_label = QLabel("Пороги для нечеткого поиска:")
        self.thresholds_label.setMinimumWidth(200)

        threshold_widget = QWidget()
        threshold_inner_layout = QHBoxLayout(threshold_widget)
        threshold_inner_layout.setContentsMargins(0, 0, 0, 0)
        threshold_inner_layout.setSpacing(10)

        self.threshold_input = QLineEdit("70")
        self.threshold_input.setMaximumWidth(80)
        self.trust_threshold_input = QLineEdit("85")
        self.trust_threshold_input.setMaximumWidth(80)

        threshold_inner_layout.addWidget(QLabel("Порог:"))
        threshold_inner_layout.addWidget(self.threshold_input)
        threshold_inner_layout.addWidget(QLabel("Порог доверия:"))
        threshold_inner_layout.addWidget(self.trust_threshold_input)
        threshold_inner_layout.addStretch()

        thresholds_layout.addWidget(self.thresholds_label)
        thresholds_layout.addWidget(threshold_widget)
        form_layout.addWidget(self.thresholds_container)
        self.thresholds_container.setVisible(False)

        self.field_widgets['thresholds'] = threshold_widget
        self.label_widgets['thresholds'] = self.thresholds_label

    def create_version_sections(self, form_layout):
        """Создание секций для проверки показателей назначения"""
        self.version_sections_group = QGroupBox("Показатели назначения (проверка версий)")
        self.version_sections_group.setVisible(False)

        version_layout = QVBoxLayout(self.version_sections_group)
        version_layout.setSpacing(10)

        # Информация о типе проверки
        version_info_label = QLabel("""
        <div style='color: #bdc3c7; font-size: 12px; padding: 5px;'>
        Проверка версий ПО и оборудования с поддержкой различных операторов сравнения.
        Можно использовать операторы: =, !=, >, <, >=, <=
        </div>
        """)
        version_info_label.setWordWrap(True)
        version_layout.addWidget(version_info_label)

        # Кнопка добавления секции
        self.add_section_btn = QPushButton("+ Добавить показатель")
        self.add_section_btn.clicked.connect(self.add_version_section)
        self.add_section_btn.setVisible(False)
        version_layout.addWidget(self.add_section_btn)

        # Контейнер для секций с прокруткой
        self.sections_scroll = QScrollArea()
        self.sections_scroll.setWidgetResizable(True)
        self.sections_scroll.setMinimumHeight(200)
        self.sections_container = QWidget()
        self.sections_container_layout = QVBoxLayout(self.sections_container)
        self.sections_container_layout.setSpacing(8)
        self.sections_scroll.setWidget(self.sections_container)
        version_layout.addWidget(self.sections_scroll)

        # Настройки проверки
        settings_widget = QWidget()
        settings_layout = QVBoxLayout(settings_widget)
        settings_layout.setSpacing(8)

        # Режим проверки
        mode_layout = QHBoxLayout()
        self.strict_mode_check = QCheckBox("Строгий режим (все показатели должны быть выполнены)")
        self.strict_mode_check.setChecked(True)
        self.strict_mode_check.setVisible(False)
        mode_layout.addWidget(self.strict_mode_check)
        mode_layout.addStretch()
        settings_layout.addLayout(mode_layout)

        # Требуемое количество
        required_layout = QHBoxLayout()
        required_label = QLabel("Требуется показателей:")
        self.required_total_spin = QSpinBox()
        self.required_total_spin.setMinimum(1)
        self.required_total_spin.setMaximum(100)
        self.required_total_spin.setValue(3)
        self.required_total_spin.setVisible(False)
        required_layout.addWidget(required_label)
        required_layout.addWidget(self.required_total_spin)
        required_layout.addStretch()
        settings_layout.addLayout(required_layout)

        version_layout.addWidget(settings_widget)
        form_layout.addWidget(self.version_sections_group)

    def create_combined_sections(self, form_layout):
        """Создание секций для комбинированных проверок"""
        self.combined_sections_group = QGroupBox("Комбинированные проверки (логические условия)")
        self.combined_sections_group.setObjectName("combined_group")
        self.combined_sections_group.setVisible(False)

        combined_layout = QVBoxLayout(self.combined_sections_group)
        combined_layout.setSpacing(10)

        # Информация о типе проверки
        combined_info_label = QLabel("""
        <div style='color: #bdc3c7; font-size: 12px; padding: 5px;'>
        Объединение нескольких условий с логическими операторами И/ИЛИ. 
        Каждое условие может быть любого типа проверки.
        </div>
        """)
        combined_info_label.setWordWrap(True)
        combined_layout.addWidget(combined_info_label)

        # Кнопка добавления условия
        self.add_combined_section_btn = QPushButton("+ Добавить условие")
        self.add_combined_section_btn.clicked.connect(self.add_combined_section)
        self.add_combined_section_btn.setVisible(False)
        combined_layout.addWidget(self.add_combined_section_btn)

        # Контейнер для условий с прокруткой
        self.combined_sections_scroll = QScrollArea()
        self.combined_sections_scroll.setWidgetResizable(True)
        self.combined_sections_scroll.setMinimumHeight(200)
        self.combined_sections_container = QWidget()
        self.combined_sections_container_layout = QVBoxLayout(self.combined_sections_container)
        self.combined_sections_container_layout.setSpacing(8)
        self.combined_sections_scroll.setWidget(self.combined_sections_container)
        combined_layout.addWidget(self.combined_sections_scroll)

        # Настройки логических операторов
        logic_widget = QWidget()
        logic_layout = QVBoxLayout(logic_widget)
        logic_layout.setSpacing(8)

        # Логический оператор
        operator_layout = QHBoxLayout()
        operator_label = QLabel("Логический оператор:")
        self.logic_operator_combo = QComboBox()
        self.logic_operator_combo.addItems(["И (AND) - все условия должны быть выполнены",
                                            "ИЛИ (OR) - достаточно выполнения указанного количества условий"])
        self.logic_operator_combo.setVisible(False)
        operator_layout.addWidget(operator_label)
        operator_layout.addWidget(self.logic_operator_combo)
        operator_layout.addStretch()
        logic_layout.addLayout(operator_layout)

        # Требуемое количество
        required_combined_layout = QHBoxLayout()
        required_combined_label = QLabel("Требуется выполнить условий (для ИЛИ):")
        self.required_passed_spin = QSpinBox()
        self.required_passed_spin.setMinimum(1)
        self.required_passed_spin.setMaximum(100)
        self.required_passed_spin.setValue(1)
        self.required_passed_spin.setVisible(False)
        required_combined_layout.addWidget(required_combined_label)
        required_combined_layout.addWidget(self.required_passed_spin)
        required_combined_layout.addStretch()
        logic_layout.addLayout(required_combined_layout)

        combined_layout.addWidget(logic_widget)
        form_layout.addWidget(self.combined_sections_group)

    def create_description_field(self, form_layout):
        """Создание поля описания"""
        desc_container = QWidget()
        desc_layout = QHBoxLayout(desc_container)
        desc_layout.setContentsMargins(0, 0, 0, 0)
        desc_label = QLabel("Описание:")
        desc_label.setMinimumWidth(200)
        self.description_input = QTextEdit()
        self.description_input.setMaximumHeight(120)
        self.description_input.setPlaceholderText("Описание проверки (необязательно)")
        desc_layout.addWidget(desc_label)
        desc_layout.addWidget(self.description_input)
        form_layout.addWidget(desc_container)

        self.field_widgets['description'] = self.description_input
        self.label_widgets['description'] = desc_label

    def create_hint_label(self, main_layout):
        """Создание подсказки"""
        hint_label = QLabel("""
        <div style='background-color: #34495e; padding: 12px; border-radius: 8px; border: 1px solid #3498db;'>
        <span style='color: #3498db; font-weight: bold;'>Подсказка:</span>
        <ul style='color: #bdc3c7; margin: 5px 0; padding-left: 20px;'>
        <li>Для <b>no_text_present</b> - укажите запрещенные слова в поле "Алиасы"</li>
        <li>Для <b>text_present</b> - укажите обязательные слова в поле "Алиасы"</li>
        <li>Для <b>version_comparison</b> - проверка показателей назначения (версии ПО/оборудования) с операторами сравнения</li>
        <li>Для <b>combined_check</b> - объединение нескольких условий с логическими операторами</li>
        </ul>
        </div>
        """)
        hint_label.setWordWrap(True)
        hint_label.setStyleSheet("font-size: 12px;")
        main_layout.addWidget(hint_label)

    def create_buttons(self, main_layout):
        """Создание кнопок"""
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.validate_and_accept)
        button_box.rejected.connect(self.reject)

        if self.is_edit_mode:
            delete_btn = QPushButton("Удалить проверку")
            delete_btn.setObjectName("delete_button")
            delete_btn.clicked.connect(self.delete_check)
            button_box.addButton(delete_btn, QDialogButtonBox.ButtonRole.ActionRole)

        main_layout.addWidget(button_box)

    def on_type_changed(self, type_text):
        """Обновляет видимость полей в зависимости от типа проверки"""
        type_code = type_text.split(" - ")[0] if " - " in type_text else type_text

        # Определяем какие поля нужны для каждого типа
        needs_text = type_code in ['fuzzy_text_present', 'no_fuzzy_text_present',
                                   'fuzzy_text_present_after_any_table']
        needs_aliases = type_code in ['no_text_present', 'text_present',
                                      'text_present_without', 'text_present_in_any_table']
        needs_without = type_code == 'text_present_without'
        needs_thresholds = type_code in ['fuzzy_text_present', 'no_fuzzy_text_present',
                                         'fuzzy_text_present_after_any_table']
        needs_version_sections = type_code == 'version_comparison'
        needs_combined_sections = type_code == 'combined_check'

        # Устанавливаем видимость контейнеров полей
        self.text_container.setVisible(needs_text)
        self.aliases_container.setVisible(needs_aliases)
        self.without_container.setVisible(needs_without)
        self.thresholds_container.setVisible(needs_thresholds)

        # Секции версий (показатели назначения)
        self.version_sections_group.setVisible(needs_version_sections)
        if hasattr(self, 'add_section_btn'):
            self.add_section_btn.setVisible(needs_version_sections)
        if hasattr(self, 'required_total_spin'):
            self.required_total_spin.setVisible(needs_version_sections)
        if hasattr(self, 'strict_mode_check'):
            self.strict_mode_check.setVisible(needs_version_sections)

        # Секции комбинированных проверок
        self.combined_sections_group.setVisible(needs_combined_sections)
        if hasattr(self, 'add_combined_section_btn'):
            self.add_combined_section_btn.setVisible(needs_combined_sections)
        if hasattr(self, 'logic_operator_combo'):
            self.logic_operator_combo.setVisible(needs_combined_sections)
        if hasattr(self, 'required_passed_spin'):
            self.required_passed_spin.setVisible(needs_combined_sections)

        # Если это проверка версий и нет секций - добавляем одну по умолчанию
        if needs_version_sections and not self.version_sections_widgets:
            self.add_version_section()

        # Если это комбинированная проверка и нет секций - добавляем одно условие по умолчанию
        if needs_combined_sections and not self.combined_sections_widgets:
            self.add_combined_section()

    def add_version_section(self, section_data=None):
        """Добавляет виджет секции версии"""
        section_widget = QWidget()
        section_widget.setStyleSheet("""
            QWidget {
                background-color: #34495e;
                border: 1px solid #3498db;
                border-radius: 6px;
                padding: 10px;
                margin: 5px 0;
            }
            QLabel {
                color: #ecf0f1;
                font-size: 13px;
            }
            QLineEdit {
                background-color: #2c3e50;
                color: #ffffff;
                border: 1px solid #7f8c8d;
                border-radius: 4px;
                padding: 6px;
                selection-background-color: #3498db;
            }
            QLineEdit:focus {
                border-color: #3498db;
            }
            QComboBox {
                background-color: #2c3e50;
                color: #ffffff;
                border: 1px solid #7f8c8d;
                border-radius: 4px;
                padding: 6px;
                min-width: 80px;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #ffffff;
                margin-right: 5px;
            }
            QComboBox QAbstractItemView {
                background-color: #2c3e50;
                color: #ffffff;
                border: 1px solid #3498db;
                selection-background-color: #3498db;
                selection-color: #ffffff;
            }
            QTextEdit {
                background-color: #2c3e50;
                color: #ffffff;
                border: 1px solid #7f8c8d;
                border-radius: 4px;
                padding: 6px;
                selection-background-color: #3498db;
            }
            QTextEdit:focus {
                border-color: #3498db;
            }
            QPushButton#delete_button {
                background-color: #e74c3c;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 6px 12px;
                font-weight: bold;
                min-width: 80px;
            }
            QPushButton#delete_button:hover {
                background-color: #c0392b;
            }
            QPushButton#delete_button:pressed {
                background-color: #962d22;
            }
        """)

        section_layout = QVBoxLayout(section_widget)
        section_layout.setSpacing(10)

        # Верхняя строка: название, версия и оператор
        top_layout = QHBoxLayout()
        top_layout.setSpacing(10)

        # Название показателя
        name_label = QLabel("Название:")
        name_label.setFixedWidth(70)
        name_input = QLineEdit()
        name_input.setPlaceholderText("Например: ОС")
        name_input.setMinimumWidth(150)
        if section_data:
            name_input.setText(section_data.get('name', ''))

        # Требуемая версия
        version_label = QLabel("Версия:")
        version_label.setFixedWidth(50)
        version_input = QLineEdit()
        version_input.setPlaceholderText("10.0")
        version_input.setMaximumWidth(100)
        if section_data:
            version_input.setText(section_data.get('required_version', ''))

        # Оператор сравнения
        operator_label = QLabel("Оператор:")
        operator_label.setFixedWidth(70)
        operator_combo = QComboBox()
        operator_combo.setMaximumWidth(100)
        operator_combo.addItems([
            ">=", "<=", ">", "<", "=", "!="
        ])
        operator = section_data.get('operator', '>=') if section_data else '>='
        index = operator_combo.findText(operator)
        if index >= 0:
            operator_combo.setCurrentIndex(index)

        top_layout.addWidget(name_label)
        top_layout.addWidget(name_input)
        top_layout.addWidget(version_label)
        top_layout.addWidget(version_input)
        top_layout.addWidget(operator_label)
        top_layout.addWidget(operator_combo)
        top_layout.addStretch()

        # Кнопка удаления
        delete_btn = QPushButton("Удалить")
        delete_btn.setObjectName("delete_button")
        delete_btn.setMaximumWidth(100)
        top_layout.addWidget(delete_btn)

        section_layout.addLayout(top_layout)

        # Паттерны поиска
        patterns_label = QLabel("Регулярные выражения для поиска версий (по одному на строку):")
        patterns_label.setStyleSheet("color: #ecf0f1; font-weight: bold; margin-top: 5px;")
        patterns_input = QTextEdit()
        patterns_input.setMaximumHeight(100)
        patterns_input.setPlaceholderText(
            "Примеры:\nWindows\\s+([\\d\\.]+)\nLinux\\s+kernel\\s+([\\d\\.]+)\nВерсия ОС\\s+([\\d\\.]+)")
        if section_data:
            patterns_input.setPlainText('\n'.join(section_data.get('version_patterns', [])))

        section_layout.addWidget(patterns_label)
        section_layout.addWidget(patterns_input)

        # Сохраняем ссылки на виджеты
        section_widget.name_input = name_input
        section_widget.version_input = version_input
        section_widget.operator_combo = operator_combo
        section_widget.patterns_input = patterns_input

        # Обработчик удаления
        delete_btn.clicked.connect(lambda: self.remove_version_section(section_widget))

        # Добавляем в контейнер
        self.sections_container_layout.addWidget(section_widget)
        self.version_sections_widgets.append(section_widget)

        return section_widget

    def remove_version_section(self, section_widget):
        """Удаляет секцию версии"""
        if section_widget in self.version_sections_widgets:
            self.version_sections_widgets.remove(section_widget)
            self.sections_container_layout.removeWidget(section_widget)
            section_widget.setParent(None)
            section_widget.deleteLater()

    def add_combined_section(self, section_data=None):
        """Добавляет виджет условия комбинированной проверки"""
        section_widget = QWidget()
        section_widget.setStyleSheet("""
            QWidget {
                background-color: #34495e;
                border: 1px solid #9b59b6;
                border-radius: 6px;
                padding: 10px;
                margin: 5px 0;
            }
            QLabel {
                color: #ecf0f1;
                font-size: 13px;
            }
            QLineEdit {
                background-color: #2c3e50;
                color: #ffffff;
                border: 1px solid #7f8c8d;
                border-radius: 4px;
                padding: 6px;
                selection-background-color: #9b59b6;
            }
            QLineEdit:focus {
                border-color: #9b59b6;
            }
            QComboBox {
                background-color: #2c3e50;
                color: #ffffff;
                border: 1px solid #7f8c8d;
                border-radius: 4px;
                padding: 6px;
                min-width: 200px;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 5px solid #ffffff;
                margin-right: 5px;
            }
            QComboBox QAbstractItemView {
                background-color: #2c3e50;
                color: #ffffff;
                border: 1px solid #9b59b6;
                selection-background-color: #9b59b6;
                selection-color: #ffffff;
            }
            QTextEdit {
                background-color: #2c3e50;
                color: #ffffff;
                border: 1px solid #7f8c8d;
                border-radius: 4px;
                padding: 6px;
                selection-background-color: #9b59b6;
            }
            QTextEdit:focus {
                border-color: #9b59b6;
            }
            QPushButton#delete_button {
                background-color: #e74c3c;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 6px 12px;
                font-weight: bold;
                min-width: 80px;
            }
            QPushButton#delete_button:hover {
                background-color: #c0392b;
            }
            QPushButton#delete_button:pressed {
                background-color: #962d22;
            }
        """)

        section_layout = QVBoxLayout(section_widget)
        section_layout.setSpacing(10)

        # Верхняя строка: тип условия и название
        top_layout = QHBoxLayout()
        top_layout.setSpacing(10)

        # Тип условия
        type_label = QLabel("Тип условия:")
        type_label.setFixedWidth(90)
        type_combo = QComboBox()
        type_combo.setMinimumWidth(250)
        type_combo.addItems([
            "no_text_present - Текст не должен присутствовать",
            "text_present - Текст должен присутствовать",
            "text_present_without - Текст должен быть, а другой - нет",
            "fuzzy_text_present - Нечеткое соответствие текста",
            "no_fuzzy_text_present - Нечеткое несоответствие текста"
        ])

        if section_data:
            for i in range(type_combo.count()):
                if type_combo.itemText(i).startswith(section_data.get('type', '')):
                    type_combo.setCurrentIndex(i)
                    break

        # Название условия
        name_label = QLabel("Название:")
        name_label.setFixedWidth(60)
        name_input = QLineEdit()
        name_input.setPlaceholderText("Например: Проверка Oracle")
        name_input.setMinimumWidth(200)
        if section_data:
            name_input.setText(section_data.get('name', ''))

        top_layout.addWidget(type_label)
        top_layout.addWidget(type_combo)
        top_layout.addWidget(name_label)
        top_layout.addWidget(name_input)
        top_layout.addStretch()

        # Кнопка удаления
        delete_btn = QPushButton("Удалить")
        delete_btn.setObjectName("delete_button")
        delete_btn.setMaximumWidth(100)
        top_layout.addWidget(delete_btn)

        section_layout.addLayout(top_layout)

        # Контейнер для полей условий
        condition_fields_widget = QWidget()
        condition_layout = QVBoxLayout(condition_fields_widget)
        condition_layout.setSpacing(8)

        # Поля ввода
        aliases_label = QLabel("Алиасы (через запятую):")
        aliases_label.setStyleSheet("font-weight: bold;")
        aliases_input = QTextEdit()
        aliases_input.setMaximumHeight(60)
        aliases_input.setPlaceholderText("Слова для поиска через запятую")

        text_label = QLabel("Текст для нечеткого поиска:")
        text_label.setStyleSheet("font-weight: bold;")
        text_input = QTextEdit()
        text_input.setMaximumHeight(60)
        text_input.setPlaceholderText("Текст для нечеткого поиска")

        without_label = QLabel("Исключающие алиасы:")
        without_label.setStyleSheet("font-weight: bold;")
        without_input = QTextEdit()
        without_input.setMaximumHeight(60)
        without_input.setPlaceholderText("Текст, который НЕ должен присутствовать")

        # Контейнер для порогов
        thresholds_container = QWidget()
        thresholds_layout = QHBoxLayout(thresholds_container)
        thresholds_layout.setContentsMargins(0, 0, 0, 0)
        thresholds_layout.setSpacing(10)

        thresholds_label = QLabel("Пороги:")
        thresholds_label.setStyleSheet("font-weight: bold;")
        threshold_input = QLineEdit("70")
        threshold_input.setMaximumWidth(80)
        trust_threshold_input = QLineEdit("85")
        trust_threshold_input.setMaximumWidth(80)

        thresholds_layout.addWidget(thresholds_label)
        thresholds_layout.addWidget(QLabel("Порог:"))
        thresholds_layout.addWidget(threshold_input)
        thresholds_layout.addWidget(QLabel("Доверие:"))
        thresholds_layout.addWidget(trust_threshold_input)
        thresholds_layout.addStretch()

        # Заполняем данные если есть
        if section_data:
            if 'aliases' in section_data:
                aliases_input.setPlainText(', '.join(section_data['aliases']))
            if 'text' in section_data:
                text_input.setPlainText(section_data['text'])
            if 'without_aliases' in section_data:
                without_input.setPlainText(', '.join(section_data['without_aliases']))
            if 'threshold' in section_data:
                threshold_input.setText(str(section_data['threshold']))
            if 'trust_threshold' in section_data:
                trust_threshold_input.setText(str(section_data['trust_threshold']))

        condition_layout.addWidget(aliases_label)
        condition_layout.addWidget(aliases_input)
        condition_layout.addWidget(text_label)
        condition_layout.addWidget(text_input)
        condition_layout.addWidget(without_label)
        condition_layout.addWidget(without_input)
        condition_layout.addWidget(thresholds_container)

        section_layout.addWidget(condition_fields_widget)

        # Обработчик изменения типа условия
        def on_type_changed_in_section(type_text):
            type_code = type_text.split(" - ")[0] if " - " in type_text else type_text

            needs_text = type_code in ['fuzzy_text_present', 'no_fuzzy_text_present']
            needs_without = type_code == 'text_present_without'
            needs_thresholds = type_code in ['fuzzy_text_present', 'no_fuzzy_text_present']

            # Показываем/скрываем соответствующие поля
            text_label.setVisible(needs_text)
            text_input.setVisible(needs_text)
            without_label.setVisible(needs_without)
            without_input.setVisible(needs_without)
            thresholds_container.setVisible(needs_thresholds)

            # Показываем/скрываем поле алиасов
            aliases_label.setVisible(not needs_text)
            aliases_input.setVisible(not needs_text)

        type_combo.currentTextChanged.connect(on_type_changed_in_section)
        on_type_changed_in_section(type_combo.currentText())

        # Сохраняем ссылки на виджеты
        section_widget.type_combo = type_combo
        section_widget.name_input = name_input
        section_widget.aliases_input = aliases_input
        section_widget.text_input = text_input
        section_widget.without_input = without_input
        section_widget.threshold_input = threshold_input
        section_widget.trust_threshold_input = trust_threshold_input
        section_widget.on_type_changed = on_type_changed_in_section

        # Обработчик удаления
        delete_btn.clicked.connect(lambda: self.remove_combined_section(section_widget))

        # Добавляем в контейнер
        self.combined_sections_container_layout.addWidget(section_widget)
        self.combined_sections_widgets.append(section_widget)

        return section_widget

    def remove_combined_section(self, section_widget):
        """Удаляет условие комбинированной проверки"""
        if section_widget in self.combined_sections_widgets:
            self.combined_sections_widgets.remove(section_widget)
            self.combined_sections_container_layout.removeWidget(section_widget)
            section_widget.setParent(None)
            section_widget.deleteLater()

    def fill_form_data(self):
        """Заполняет форму данными для редактирования"""
        if not self.edit_data:
            return

        # Базовые поля
        self.name_input.setText(self.edit_data.get('name', ''))

        # Группа (если есть, добавляем в комбобокс)
        group = self.edit_data.get('group', '')
        if group and self.group_input.findText(group) == -1:
            self.group_input.addItem(group)
        self.group_input.setCurrentText(group)

        # Тип проверки
        check_type = self.edit_data.get('type', '')
        for i in range(self.type_combo.count()):
            if self.type_combo.itemText(i).startswith(check_type):
                self.type_combo.setCurrentIndex(i)
                break

        # Текст
        self.text_input.setPlainText(self.edit_data.get('text', ''))

        # Алиасы
        aliases = self.edit_data.get('aliases', [])
        self.aliases_input.setPlainText(', '.join(aliases))

        # Исключающие алиасы
        without_aliases = self.edit_data.get('without_aliases', [])
        self.without_aliases_input.setPlainText(', '.join(without_aliases))

        # Пороги
        self.threshold_input.setText(str(self.edit_data.get('threshold', '70')))
        self.trust_threshold_input.setText(str(self.edit_data.get('trust_threshold', '85')))

        # Заполняем секции версий если тип подходящий
        if check_type == 'version_comparison':
            version_sections = self.edit_data.get('version_sections', [])
            for section in version_sections:
                self.add_version_section(section)

            self.required_total_spin.setValue(self.edit_data.get('required_total_indicators', 3))
            self.strict_mode_check.setChecked(self.edit_data.get('strict_mode', True))

        # Заполняем условия комбинированной проверки если тип подходящий
        elif check_type == 'combined_check':
            conditions = self.edit_data.get('conditions', [])
            for condition in conditions:
                self.add_combined_section(condition)

            logic_operator = self.edit_data.get('logic_operator', 'AND')
            operator_text = "И (AND) - все условия должны быть выполнены" if logic_operator == 'AND' else "ИЛИ (OR) - достаточно выполнения указанного количества условий"
            self.logic_operator_combo.setCurrentText(operator_text)
            self.required_passed_spin.setValue(self.edit_data.get('required_passed', 1))

        # Описание
        self.description_input.setPlainText(self.edit_data.get('description', ''))

    def validate_and_accept(self):
        """Проверяет данные и принимает форму"""
        # Получаем данные
        name = self.name_input.text().strip()
        group = self.group_input.currentText().strip()
        type_full = self.type_combo.currentText()
        check_type = type_full.split(" - ")[0] if " - " in type_full else type_full

        if not name:
            QMessageBox.warning(self, "Ошибка", "Введите название проверки")
            return

        if not group:
            QMessageBox.warning(self, "Ошибка", "Введите или выберите группу")
            return

        # Валидация в зависимости от типа
        if check_type in ['fuzzy_text_present', 'no_fuzzy_text_present',
                          'fuzzy_text_present_after_any_table']:
            text = self.text_input.toPlainText().strip()
            if not text:
                QMessageBox.warning(self, "Ошибка",
                                    f"Для типа '{check_type}' необходимо указать текст для поиска")
                return

            # Проверяем пороги
            try:
                threshold = float(self.threshold_input.text())
                trust_threshold = float(self.trust_threshold_input.text())
                if threshold >= trust_threshold:
                    QMessageBox.warning(self, "Ошибка",
                                        "Порог проверки должен быть меньше порога доверия")
                    return
            except ValueError:
                QMessageBox.warning(self, "Ошибка", "Пороги должны быть числами")
                return

        elif check_type == 'version_comparison':
            # Проверяем, что есть хотя бы одна секция
            if not self.version_sections_widgets:
                QMessageBox.warning(self, "Ошибка",
                                    "Для типа 'version_comparison' необходимо добавить хотя бы один показатель версии")
                return

            # Проверяем каждую секцию
            for section_widget in self.version_sections_widgets:
                section_name = section_widget.name_input.text().strip()
                version = section_widget.version_input.text().strip()
                patterns = section_widget.patterns_input.toPlainText().strip()

                if not section_name:
                    QMessageBox.warning(self, "Ошибка",
                                        "Не указано название показателя")
                    return

                if not version:
                    QMessageBox.warning(self, "Ошибка",
                                        f"Не указана требуемая версия в показателе '{section_name}'")
                    return

                if not patterns:
                    QMessageBox.warning(self, "Ошибка",
                                        f"Не указаны паттерны поиска в показателе '{section_name}'")
                    return

                # Проверяем валидность паттернов
                pattern_lines = [p.strip() for p in patterns.split('\n') if p.strip()]
                for pattern in pattern_lines:
                    try:
                        re.compile(pattern)
                    except re.error as e:
                        QMessageBox.warning(self, "Ошибка",
                                            f"Некорректное регулярное выражение в показателе '{section_name}':\n{pattern}\n\nОшибка: {str(e)}")
                        return

        elif check_type == 'combined_check':
            # Проверяем, что есть хотя бы одно условие
            if not self.combined_sections_widgets:
                QMessageBox.warning(self, "Ошибка",
                                    "Для комбинированной проверки необходимо добавить хотя бы одно условие")
                return

            # Проверяем каждое условие
            for section_widget in self.combined_sections_widgets:
                section_name = section_widget.name_input.text().strip()
                type_full = section_widget.type_combo.currentText()
                section_type = type_full.split(" - ")[0] if " - " in type_full else type_full

                if not section_name:
                    QMessageBox.warning(self, "Ошибка",
                                        "Не указано название условия")
                    return

                # Вызываем обработчик изменения типа для обновления видимости полей
                if hasattr(section_widget, 'on_type_changed'):
                    section_widget.on_type_changed(type_full)

                if section_type in ['fuzzy_text_present', 'no_fuzzy_text_present']:
                    text = section_widget.text_input.toPlainText().strip()
                    if not text:
                        QMessageBox.warning(self, "Ошибка",
                                            f"Для типа '{section_type}' необходимо указать текст для поиска")
                        return

                    # Проверяем пороги
                    try:
                        threshold = float(section_widget.threshold_input.text())
                        trust_threshold = float(section_widget.trust_threshold_input.text())
                        if threshold >= trust_threshold:
                            QMessageBox.warning(self, "Ошибка",
                                                f"В условии '{section_name}': порог проверки должен быть меньше порога доверия")
                            return
                    except ValueError:
                        QMessageBox.warning(self, "Ошибка",
                                            f"В условии '{section_name}': пороги должны быть числами")
                        return
                else:
                    aliases_text = section_widget.aliases_input.toPlainText().strip()
                    if not aliases_text:
                        QMessageBox.warning(self, "Ошибка",
                                            f"Для типа '{section_type}' необходимо указать алиасы")
                        return

        else:
            aliases_text = self.aliases_input.toPlainText().strip()
            if not aliases_text:
                QMessageBox.warning(self, "Ошибка",
                                    f"Для типа '{check_type}' необходимо указать алиасы")
                return

        self.accept()

    def get_check_data(self):
        """Возвращает данные проверки в формате для конфига"""
        name = self.name_input.text().strip()
        group = self.group_input.currentText().strip()
        type_full = self.type_combo.currentText()
        check_type = type_full.split(" - ")[0] if " - " in type_full else type_full

        check_data = {
            'name': name,
            'type': check_type,
            'group': group
        }

        # Добавляем поля в зависимости от типа
        if check_type in ['fuzzy_text_present', 'no_fuzzy_text_present',
                          'fuzzy_text_present_after_any_table']:
            check_data['text'] = self.text_input.toPlainText().strip()
            check_data['threshold'] = float(self.threshold_input.text())
            check_data['trust_threshold'] = float(self.trust_threshold_input.text())

        elif check_type == 'version_comparison':
            # Собираем секции версий
            version_sections = []
            for section_widget in self.version_sections_widgets:
                section_data = {
                    'name': section_widget.name_input.text().strip(),
                    'required_version': section_widget.version_input.text().strip(),
                    'operator': section_widget.operator_combo.currentText().strip(),
                    'version_patterns': [p.strip() for p in
                                         section_widget.patterns_input.toPlainText().strip().split('\n')
                                         if p.strip()]
                }
                version_sections.append(section_data)

            check_data['version_sections'] = version_sections
            check_data['required_total_indicators'] = self.required_total_spin.value()
            check_data['strict_mode'] = self.strict_mode_check.isChecked()

        elif check_type == 'combined_check':
            # Собираем условия комбинированной проверки
            combined_conditions = []
            for section_widget in self.combined_sections_widgets:
                type_full = section_widget.type_combo.currentText()
                condition_type = type_full.split(" - ")[0] if " - " in type_full else type_full

                condition_data = {
                    'name': section_widget.name_input.text().strip(),
                    'type': condition_type
                }

                # Добавляем поля в зависимости от типа
                if condition_type in ['fuzzy_text_present', 'no_fuzzy_text_present']:
                    condition_data['text'] = section_widget.text_input.toPlainText().strip()
                    condition_data['threshold'] = float(section_widget.threshold_input.text())
                    condition_data['trust_threshold'] = float(section_widget.trust_threshold_input.text())
                else:
                    aliases_text = section_widget.aliases_input.toPlainText().strip()
                    if aliases_text:
                        aliases = [alias.strip() for alias in aliases_text.split(',') if alias.strip()]
                        condition_data['aliases'] = aliases

                    if condition_type == 'text_present_without':
                        without_text = section_widget.without_input.toPlainText().strip()
                        if without_text:
                            without_aliases = [alias.strip() for alias in without_text.split(',') if alias.strip()]
                            condition_data['without_aliases'] = without_aliases

                combined_conditions.append(condition_data)

            check_data['conditions'] = combined_conditions
            check_data['logic_operator'] = 'AND' if "И (AND)" in self.logic_operator_combo.currentText() else 'OR'
            check_data['required_passed'] = self.required_passed_spin.value()

        else:
            aliases_text = self.aliases_input.toPlainText().strip()
            if aliases_text:
                aliases = [alias.strip() for alias in aliases_text.split(',') if alias.strip()]
                check_data['aliases'] = aliases

        # Исключающие алиасы
        without_text = self.without_aliases_input.toPlainText().strip()
        if without_text and check_type == 'text_present_without':
            without_aliases = [alias.strip() for alias in without_text.split(',') if alias.strip()]
            check_data['without_aliases'] = without_aliases

        # Описание
        description = self.description_input.toPlainText().strip()
        if description:
            check_data['description'] = description

        return check_data

    def delete_check(self):
        """Удаление проверки"""
        reply = QMessageBox.question(
            self, "Удаление",
            "Вы уверены, что хотите удалить эту проверку?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.deleted = True
            self.reject()