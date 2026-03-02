# config_editor.py
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QLabel, QPushButton,
    QTextEdit, QHBoxLayout, QFileDialog, QMessageBox,
    QDialogButtonBox
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont
import yaml
import io


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
        aliases: ["коммутатор", "маршрутизатор", "сервер", "хранилище"]

  - group: "Показатели назначения"
    subchecks:
      - name: "Минимальные версии ПО"
        type: "version_comparison"
        description: "Проверка минимальных версий программного обеспечения"
        version_sections:
          - name: "Операционная система"
            required_version: "10.0"
            version_patterns:
              - "Windows\\s+([\\d\\.]+)"
              - "Linux\\s+kernel\\s+([\\d\\.]+)"
              - "ОС\\s+([\\d\\.]+)"

          - name: "СУБД"
            required_version: "12.0"
            version_patterns:
              - "Oracle\\s+Database\\s+([\\d\\.]+)"
              - "PostgreSQL\\s+([\\d\\.]+)"
              - "MySQL\\s+([\\d\\.]+)"

          - name: "Веб-сервер"
            required_version: "2.4"
            version_patterns:
              - "Apache\\s+([\\d\\.]+)"
              - "nginx\\s+([\\d\\.]+)"

        required_total_indicators: 2
        strict_mode: false

      - name: "Требования к аппаратному обеспечению"
        type: "version_comparison"
        description: "Проверка версий аппаратных компонентов"
        version_sections:
          - name: "Процессор"
            required_version: "8.0"
            version_patterns:
              - "CPU\\s+([\\d\\.]+)"
              - "процессор\\s+([\\d\\.]+)"
              - "ядер\\s+([\\d\\.]+)"

          - name: "Оперативная память"
            required_version: "16.0"
            version_patterns:
              - "RAM\\s+([\\d\\.]+)"
              - "память\\s+([\\d\\.]+)"

          - name: "Хранилище"
            required_version: "1.0"
            version_patterns:
              - "SSD\\s+([\\d\\.]+)"
              - "HDD\\s+([\\d\\.]+)"
              - "диск\\s+([\\d\\.]+)"

        required_total_indicators: 2
        strict_mode: false

  - group: "Комбинированные проверки"
    subchecks:
      - name: "Импортозамещение И безопасность"
        type: "combined_check"
        description: "Комбинированная проверка на отсутствие импортного ПО и наличие требований безопасности"
        logic_operator: "AND"
        required_passed: 2
        conditions:
          - name: "Отсутствие Oracle"
            type: "no_text_present"
            aliases: ["Oracle", "Oracle Database"]

          - name: "Требования безопасности"
            type: "text_present"
            aliases: ["безопасность", "защита данных", "конфиденциальность"]

      - name: "Российское ПО ИЛИ импортозамещение"
        type: "combined_check"
        description: "Должно быть либо российское ПО, либо отсутствовать импортное"
        logic_operator: "OR"
        required_passed: 1
        conditions:
          - name: "Российское ПО присутствует"
            type: "text_present"
            aliases: ["Российское ПО", "отечественное", "МойОфис"]

          - name: "Импортное ПО отсутствует"
            type: "no_text_present"
            aliases: ["Cisco", "Juniper", "Check Point"]"""

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