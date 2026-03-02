# settings_dialog.py
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QComboBox,
    QLineEdit, QCheckBox, QDialogButtonBox,
    QHBoxLayout, QMessageBox
)
from PyQt6.QtCore import Qt


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