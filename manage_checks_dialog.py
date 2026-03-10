# manage_checks_dialog.py
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView,
    QDialogButtonBox, QAbstractItemView, QMenu, QInputDialog,
    QToolBar, QWidget, QSplitter, QTextEdit, QLineEdit,
    QCheckBox, QGroupBox, QComboBox, QTabWidget, QApplication
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QAction, QColor, QFont
from add_check_dialog import AddCheckDialog
from json_database import JSONDatabase
import yaml

class ManageChecksDialog(QDialog):
    """Диалог управления всеми проверками"""

    config_changed = pyqtSignal(dict)

    def __init__(self, parent=None, db: JSONDatabase = None):
        super().__init__(parent)
        self.parent = parent
        self.db = db or JSONDatabase()
        self.filtered_rows = []

        self.setWindowTitle("Управление проверками")
        self.resize(1400, 900)

        self.setup_ui()
        self.load_checks_to_table()
        self.update_filters()
        self.update_stats()

        if parent and hasattr(parent, 'styleSheet'):
            self.setStyleSheet(parent.styleSheet())

    def setup_ui(self):
        """Настройка интерфейса"""
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        toolbar = self.create_toolbar()
        main_layout.addWidget(toolbar)

        splitter = QSplitter(Qt.Orientation.Vertical)

        top_panel = QWidget()
        top_layout = QVBoxLayout(top_panel)
        top_layout.setContentsMargins(0, 0, 0, 0)

        filter_panel = QHBoxLayout()

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Поиск по названию, группе или описанию...")
        self.search_input.textChanged.connect(self.filter_table)

        self.group_filter = QComboBox()
        self.group_filter.addItem("Все группы")
        self.group_filter.currentTextChanged.connect(self.filter_table)

        self.type_filter = QComboBox()
        self.type_filter.addItem("Все типы")
        self.type_filter.currentTextChanged.connect(self.filter_table)

        self.status_filter = QComboBox()
        self.status_filter.addItems(["Все", "Включенные", "Выключенные"])
        self.status_filter.currentTextChanged.connect(self.filter_table)

        filter_panel.addWidget(QLabel("Поиск:"))
        filter_panel.addWidget(self.search_input, 2)
        filter_panel.addWidget(QLabel("Группа:"))
        filter_panel.addWidget(self.group_filter)
        filter_panel.addWidget(QLabel("Тип:"))
        filter_panel.addWidget(self.type_filter)
        filter_panel.addWidget(QLabel("Статус:"))
        filter_panel.addWidget(self.status_filter)
        filter_panel.addStretch()

        top_layout.addLayout(filter_panel)

        self.table = QTableWidget()
        self.table.setColumnCount(8)
        self.table.setHorizontalHeaderLabels([
            "ID", "Вкл", "Название", "Группа", "Тип", "Описание", "Создано", "Действия"
        ])

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(7, QHeaderView.ResizeMode.ResizeToContents)

        self.table.setColumnWidth(3, 150)
        self.table.setColumnWidth(4, 180)
        self.table.setColumnWidth(6, 150)
        self.table.setColumnWidth(7, 200)

        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)

        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_table_context_menu)

        top_layout.addWidget(self.table)

        bottom_panel = QTabWidget()

        preview_tab = QWidget()
        preview_layout = QVBoxLayout(preview_tab)
        preview_label = QLabel("Предпросмотр конфигурации:")
        preview_layout.addWidget(preview_label)

        self.yaml_preview = QTextEdit()
        self.yaml_preview.setReadOnly(True)
        self.yaml_preview.setFont(QFont("Consolas", 9))
        preview_layout.addWidget(self.yaml_preview)

        stats_tab = QWidget()
        stats_layout = QVBoxLayout(stats_tab)

        self.stats_group = QGroupBox("Статистика")
        stats_inner_layout = QVBoxLayout(self.stats_group)

        self.total_label = QLabel("Всего проверок: 0")
        self.enabled_label = QLabel("Включено: 0")
        self.disabled_label = QLabel("Выключено: 0")
        self.groups_label = QLabel("Групп: 0")
        self.types_label = QLabel("Типов проверок: 0")

        stats_inner_layout.addWidget(self.total_label)
        stats_inner_layout.addWidget(self.enabled_label)
        stats_inner_layout.addWidget(self.disabled_label)
        stats_inner_layout.addWidget(self.groups_label)
        stats_inner_layout.addWidget(self.types_label)

        stats_layout.addWidget(self.stats_group)
        stats_layout.addStretch()

        bottom_panel.addTab(preview_tab, "YAML предпросмотр")
        bottom_panel.addTab(stats_tab, "Статистика")

        splitter.addWidget(top_panel)
        splitter.addWidget(bottom_panel)
        splitter.setSizes([600, 300])

        main_layout.addWidget(splitter)

        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok |
            QDialogButtonBox.StandardButton.Cancel |
            QDialogButtonBox.StandardButton.Apply
        )
        button_box.accepted.connect(self.accept_changes)
        button_box.rejected.connect(self.reject)

        apply_btn = button_box.button(QDialogButtonBox.StandardButton.Apply)
        apply_btn.setText("Применить")
        apply_btn.clicked.connect(self.apply_changes)

        main_layout.addWidget(button_box)

        self.setLayout(main_layout)

    def create_toolbar(self):
        """Создание панели инструментов"""
        toolbar = QToolBar()

        actions = [
            ("➕ Добавить", self.add_check),
            ("✏️ Редактировать", self.edit_selected),
            ("🗑️ Удалить", self.delete_selected),
            ("📋 Дублировать", self.duplicate_selected),
            (None, None),
            ("✅ Включить", self.enable_selected),
            ("❌ Выключить", self.disable_selected),
            (None, None),
            ("🗂️ Группировать", self.group_selected),
            ("📤 Экспорт YAML", self.export_yaml),
            ("📥 Импорт YAML", self.import_yaml),
            (None, None),
            ("🔄 Обновить", self.load_checks_to_table)
        ]

        for text, slot in actions:
            if text is None:
                toolbar.addSeparator()
            else:
                btn = QPushButton(text)
                btn.clicked.connect(slot)
                btn.setMaximumHeight(30)
                toolbar.addWidget(btn)

        return toolbar

    def load_checks_to_table(self):
        """Загрузка проверок в таблицу"""
        self.table.setRowCount(0)
        all_checks = self.db.get_all_checks(include_disabled=True, include_deleted=False)

        for row, check in enumerate(all_checks):
            self.table.insertRow(row)

            id_item = QTableWidgetItem(str(check.get('id', '')))
            id_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row, 0, id_item)

            enabled_item = QTableWidgetItem()
            enabled_item.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
            enabled_item.setCheckState(Qt.CheckState.Checked if check.get('is_enabled', True) else Qt.CheckState.Unchecked)
            enabled_item.setData(Qt.ItemDataRole.UserRole, check.get('id'))
            self.table.setItem(row, 1, enabled_item)

            name_item = QTableWidgetItem(check.get('name', ''))
            name_item.setData(Qt.ItemDataRole.UserRole, check)
            self.table.setItem(row, 2, name_item)

            group_item = QTableWidgetItem(check.get('group', 'Без группы'))
            self.table.setItem(row, 3, group_item)

            type_item = QTableWidgetItem(self.get_type_display(check.get('type', '')))
            self.table.setItem(row, 4, type_item)

            desc = check.get('description', '')
            desc_item = QTableWidgetItem(desc[:100] + '...' if len(desc) > 100 else desc)
            desc_item.setToolTip(desc)
            self.table.setItem(row, 5, desc_item)

            created = check.get('created_at', '')[:10] if check.get('created_at') else ''
            created_item = QTableWidgetItem(created)
            created_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.table.setItem(row, 6, created_item)

            action_widget = QWidget()
            action_layout = QHBoxLayout(action_widget)
            action_layout.setContentsMargins(0, 0, 0, 0)
            action_layout.setSpacing(5)

            edit_btn = QPushButton("✏️")
            edit_btn.setMaximumWidth(30)
            edit_btn.clicked.connect(lambda checked, r=row: self.edit_check_by_row(r))

            delete_btn = QPushButton("🗑️")
            delete_btn.setMaximumWidth(30)
            delete_btn.clicked.connect(lambda checked, r=row: self.delete_check_by_row(r))

            duplicate_btn = QPushButton("📋")
            duplicate_btn.setMaximumWidth(30)
            duplicate_btn.clicked.connect(lambda checked, r=row: self.duplicate_check_by_row(r))

            action_layout.addWidget(edit_btn)
            action_layout.addWidget(delete_btn)
            action_layout.addWidget(duplicate_btn)
            action_layout.addStretch()

            self.table.setCellWidget(row, 7, action_widget)

        self.update_yaml_preview()
        self.filter_table()

    def get_type_display(self, check_type):
        """Получить описательное название типа проверки"""
        type_map = {
            'no_text_present': 'Текст не должен присутствовать',
            'text_present': 'Текст должен присутствовать',
            'text_present_without': 'Текст должен быть, а другой - нет',
            'fuzzy_text_present': 'Нечеткое соответствие',
            'no_fuzzy_text_present': 'Нечеткое несоответствие',
            'text_present_in_any_table': 'Текст в таблицах',
            'fuzzy_text_present_after_any_table': 'Текст после таблиц',
            'version_comparison': 'Проверка версий',
            'combined_check': 'Комбинированная проверка'
        }
        return type_map.get(check_type, check_type)

    def update_filters(self):
        """Обновление фильтров"""
        current_group = self.group_filter.currentText()
        current_type = self.type_filter.currentText()

        self.group_filter.clear()
        self.group_filter.addItem("Все группы")
        self.group_filter.addItems(self.db.get_groups())

        self.type_filter.clear()
        self.type_filter.addItem("Все типы")
        self.type_filter.addItems(self.db.get_check_types())

        if current_group in [self.group_filter.itemText(i) for i in range(self.group_filter.count())]:
            self.group_filter.setCurrentText(current_group)

        if current_type in [self.type_filter.itemText(i) for i in range(self.type_filter.count())]:
            self.type_filter.setCurrentText(current_type)

    def filter_table(self):
        """Фильтрация таблицы"""
        search_text = self.search_input.text().lower()
        group = self.group_filter.currentText()
        check_type = self.type_filter.currentText()
        status = self.status_filter.currentText()

        for row in range(self.table.rowCount()):
            show_row = True

            name_item = self.table.item(row, 2)
            group_item = self.table.item(row, 3)
            type_item = self.table.item(row, 4)
            desc_item = self.table.item(row, 5)
            enabled_item = self.table.item(row, 1)

            if search_text:
                row_text = ""
                if name_item:
                    row_text += name_item.text().lower() + " "
                if group_item:
                    row_text += group_item.text().lower() + " "
                if desc_item:
                    row_text += desc_item.text().lower() + " "
                if search_text not in row_text:
                    show_row = False

            if group != "Все группы" and group_item and group_item.text() != group:
                show_row = False

            if check_type != "Все типы" and type_item and type_item.text() != self.get_type_display(check_type):
                show_row = False

            if status != "Все":
                is_enabled = enabled_item and enabled_item.checkState() == Qt.CheckState.Checked
                if status == "Включенные" and not is_enabled:
                    show_row = False
                elif status == "Выключенные" and is_enabled:
                    show_row = False

            self.table.setRowHidden(row, not show_row)

    def update_yaml_preview(self):
        """Обновление предпросмотра YAML"""
        checks = self.db.get_all_checks(include_disabled=True, include_deleted=False)
        yaml_text = yaml.dump(checks, allow_unicode=True, default_flow_style=False, indent=2)
        self.yaml_preview.setPlainText(yaml_text)

    def update_stats(self):
        """Обновление статистики"""
        all_checks = self.db.get_all_checks(include_disabled=True, include_deleted=False)
        enabled = sum(1 for c in all_checks if c.get('is_enabled', True))
        disabled = len(all_checks) - enabled
        groups = len(self.db.get_groups())
        types = len(self.db.get_check_types())

        self.total_label.setText(f"Всего проверок: {len(all_checks)}")
        self.enabled_label.setText(f"Включено: {enabled}")
        self.disabled_label.setText(f"Выключено: {disabled}")
        self.groups_label.setText(f"Групп: {groups}")
        self.types_label.setText(f"Типов проверок: {types}")

    def show_table_context_menu(self, position):
        """Показать контекстное меню таблицы"""
        menu = QMenu()
        selected_rows = self.table.selectionModel().selectedRows()

        if selected_rows:
            edit_action = QAction("✏️ Редактировать", self)
            edit_action.triggered.connect(self.edit_selected)
            menu.addAction(edit_action)

            delete_action = QAction("🗑️ Удалить", self)
            delete_action.triggered.connect(self.delete_selected)
            menu.addAction(delete_action)

            duplicate_action = QAction("📋 Дублировать", self)
            duplicate_action.triggered.connect(self.duplicate_selected)
            menu.addAction(duplicate_action)

            menu.addSeparator()

            enable_action = QAction("✅ Включить", self)
            enable_action.triggered.connect(self.enable_selected)
            menu.addAction(enable_action)

            disable_action = QAction("❌ Выключить", self)
            disable_action.triggered.connect(self.disable_selected)
            menu.addAction(disable_action)

            menu.addSeparator()

        add_action = QAction("➕ Добавить проверку", self)
        add_action.triggered.connect(self.add_check)
        menu.addAction(add_action)

        refresh_action = QAction("🔄 Обновить", self)
        refresh_action.triggered.connect(self.load_checks_to_table)
        menu.addAction(refresh_action)

        menu.exec(self.table.viewport().mapToGlobal(position))

    def add_check(self):
        """Добавить новую проверку"""
        dialog = AddCheckDialog(self)
        if dialog.exec():
            check_data = dialog.get_check_data()
            check_id = self.db.add_check(check_data)
            if check_id:
                self.load_checks_to_table()
                self.update_filters()
                self.update_stats()
                QMessageBox.information(self, "Успех", f"Проверка добавлена с ID: {check_id}")

    def edit_selected(self):
        """Редактировать выбранную проверку"""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "Ошибка", "Выберите проверку для редактирования")
            return
        row = selected_rows[0].row()
        self.edit_check_by_row(row)

    def edit_check_by_row(self, row):
        """Редактировать проверку по строке"""
        name_item = self.table.item(row, 2)
        if not name_item:
            return

        check_data = name_item.data(Qt.ItemDataRole.UserRole)
        if not check_data:
            return

        dialog = AddCheckDialog(self, check_data)
        if dialog.exec():
            if hasattr(dialog, 'deleted') and dialog.deleted:
                self.db.delete_check(check_data['id'], hard_delete=True)
            else:
                updated_data = dialog.get_check_data()
                self.db.update_check(check_data['id'], updated_data)

            self.load_checks_to_table()
            self.update_filters()
            self.update_stats()

    def delete_selected(self):
        """Удалить выбранные проверки"""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "Ошибка", "Выберите проверки для удаления")
            return

        reply = QMessageBox.question(
            self, "Подтверждение",
            f"Удалить {len(selected_rows)} проверок?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            for index in sorted(selected_rows, key=lambda x: x.row(), reverse=True):
                self.delete_check_by_row(index.row())
            self.load_checks_to_table()
            self.update_filters()
            self.update_stats()

    def delete_check_by_row(self, row):
        """Удалить проверку по строке"""
        name_item = self.table.item(row, 2)
        if not name_item:
            return
        check_data = name_item.data(Qt.ItemDataRole.UserRole)
        if check_data:
            self.db.delete_check(check_data['id'], hard_delete=True)

    def duplicate_selected(self):
        """Дублировать выбранные проверки"""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "Ошибка", "Выберите проверки для дублирования")
            return

        for index in selected_rows:
            self.duplicate_check_by_row(index.row())

        self.load_checks_to_table()
        self.update_filters()
        self.update_stats()

    def duplicate_check_by_row(self, row):
        """Дублировать проверку по строке"""
        name_item = self.table.item(row, 2)
        if not name_item:
            return

        check_data = name_item.data(Qt.ItemDataRole.UserRole)
        if check_data:
            new_check = check_data.copy()
            new_check.pop('id', None)
            new_check.pop('created_at', None)
            new_check['name'] = f"{new_check.get('name', '')} (копия)"
            self.db.add_check(new_check)

    def enable_selected(self):
        """Включить выбранные проверки"""
        selected_rows = self.table.selectionModel().selectedRows()
        for index in selected_rows:
            item = self.table.item(index.row(), 1)
            if item:
                check_id = item.data(Qt.ItemDataRole.UserRole)
                if check_id:
                    self.db.enable_check(check_id, True)
        self.load_checks_to_table()
        self.update_stats()

    def disable_selected(self):
        """Выключить выбранные проверки"""
        selected_rows = self.table.selectionModel().selectedRows()
        for index in selected_rows:
            item = self.table.item(index.row(), 1)
            if item:
                check_id = item.data(Qt.ItemDataRole.UserRole)
                if check_id:
                    self.db.enable_check(check_id, False)
        self.load_checks_to_table()
        self.update_stats()

    def group_selected(self):
        """Группировать выбранные проверки"""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "Ошибка", "Выберите проверки для группировки")
            return

        group_name, ok = QInputDialog.getText(
            self, "Название группы", "Введите название новой группы:"
        )

        if ok and group_name:
            for index in selected_rows:
                item = self.table.item(index.row(), 2)
                if item:
                    check_data = item.data(Qt.ItemDataRole.UserRole)
                    if check_data:
                        self.db.update_check(check_data['id'], {'group': group_name})
            self.load_checks_to_table()
            self.update_filters()
            self.update_stats()

    def export_yaml(self):
        """Экспорт проверок в YAML"""
        from PyQt6.QtWidgets import QFileDialog
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Экспорт проверок", "checks_export.yaml", "YAML files (*.yaml *.yml);;All files (*.*)"
        )

        if file_path:
            try:
                checks = self.db.get_all_checks(include_disabled=True, include_deleted=False)
                with open(file_path, 'w', encoding='utf-8') as f:
                    yaml.dump(checks, f, allow_unicode=True, default_flow_style=False, indent=2)
                QMessageBox.information(self, "Успех", f"Проверки экспортированы в:\n{file_path}")
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось экспортировать:\n{str(e)}")

    def import_yaml(self):
        """Импорт проверок из YAML с умным объединением"""
        from PyQt6.QtWidgets import QFileDialog, QProgressDialog

        file_path, _ = QFileDialog.getOpenFileName(
            self, "Импорт проверок", "", "YAML files (*.yaml *.yml);;All files (*.*)"
        )

        if file_path:
            try:
                # Показываем прогресс
                progress = QProgressDialog("Импорт проверок...", "Отмена", 0, 100, self)
                progress.setWindowModality(Qt.WindowModality.WindowModal)
                progress.setValue(10)

                # Спрашиваем режим импорта
                from PyQt6.QtWidgets import QCheckBox
                msg_box = QMessageBox(self)
                msg_box.setWindowTitle("Режим импорта")
                msg_box.setText("Выберите режим импорта:")

                auto_merge_cb = QCheckBox("Автоматически объединять с существующими проверками")
                auto_merge_cb.setChecked(True)

                msg_box.setCheckBox(auto_merge_cb)
                msg_box.setStandardButtons(QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel)

                if msg_box.exec() != QMessageBox.StandardButton.Ok:
                    return

                progress.setValue(30)

                # Выполняем импорт
                added, updated, skipped, details = self.db.import_checks_from_yaml(
                    file_path,
                    auto_merge=auto_merge_cb.isChecked()
                )

                progress.setValue(70)

                # Обновляем интерфейс
                self.load_checks_to_table()
                self.update_filters()
                self.update_stats()

                progress.setValue(100)

                # Показываем результат
                result_text = f"Импорт завершен:\n\n"
                result_text += f"✅ Добавлено новых проверок: {added}\n"
                result_text += f"🔄 Обновлено существующих: {updated}\n"
                result_text += f"⏭️ Пропущено: {skipped}\n\n"

                if details:
                    result_text += "Детали:\n"
                    for detail in details[:10]:  # Показываем первые 10
                        check_name = detail.get('check', {}).get('name', 'Unknown')
                        action = detail.get('action', 'unknown')
                        messages = detail.get('messages', [])
                        if messages:
                            result_text += f"  • {check_name}: {action} - {messages[0]}\n"

                    if len(details) > 10:
                        result_text += f"  ... и еще {len(details) - 10} операций\n"

                QMessageBox.information(self, "Результат импорта", result_text)

            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось импортировать:\n{str(e)}")
                import traceback
                traceback.print_exc()

    def apply_changes(self):
        """Применить изменения"""
        for row in range(self.table.rowCount()):
            enabled_item = self.table.item(row, 1)
            if enabled_item:
                check_id = enabled_item.data(Qt.ItemDataRole.UserRole)
                is_enabled = enabled_item.checkState() == Qt.CheckState.Checked
                if check_id:
                    self.db.enable_check(check_id, is_enabled)

        self.db._save_checks()
        self.update_yaml_preview()
        self.update_stats()
        QMessageBox.information(self, "Успех", "Изменения применены!")

    def accept_changes(self):
        """Принять изменения и закрыть диалог"""
        self.apply_changes()
        config = {'checks': self.db.get_all_checks(include_disabled=True, include_deleted=False)}
        self.config_changed.emit(config)
        self.accept()