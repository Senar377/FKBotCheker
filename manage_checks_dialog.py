from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QTableWidget, QTableWidgetItem, QMessageBox, QHeaderView,
    QDialogButtonBox, QAbstractItemView, QMenu, QInputDialog,
    QToolBar, QWidget, QSplitter, QTextEdit, QLineEdit,
    QCheckBox, QGroupBox, QComboBox, QTabWidget
)
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QAction, QIcon, QColor, QFont
from add_check_dialog import AddCheckDialog
import yaml


class ManageChecksDialog(QDialog):
    """Диалог управления всеми проверками"""

    config_changed = pyqtSignal(dict)  # Сигнал об изменении конфигурации

    def __init__(self, parent=None, config=None):
        super().__init__(parent)
        self.parent = parent
        self.config = config or {}
        self.original_config = yaml.dump(config, default_flow_style=False) if config else ""

        self.setWindowTitle("Управление проверками")
        self.resize(1200, 800)

        self.setup_ui()
        self.load_config_to_table()
        self.update_stats()

        # Применяем тему родителя
        if parent and hasattr(parent, 'styleSheet'):
            self.setStyleSheet(parent.styleSheet())

    def setup_ui(self):
        """Настройка интерфейса"""
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        # Панель инструментов
        toolbar = self.create_toolbar()
        main_layout.addWidget(toolbar)

        # Разделитель
        splitter = QSplitter(Qt.Orientation.Vertical)

        # Верхняя панель: таблица проверок
        top_panel = QWidget()
        top_layout = QVBoxLayout(top_panel)

        # Панель поиска и фильтров
        filter_panel = QHBoxLayout()

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Поиск по названию, группе или типу...")
        self.search_input.textChanged.connect(self.filter_table)

        self.group_filter = QComboBox()
        self.group_filter.addItem("Все группы")
        self.group_filter.currentTextChanged.connect(self.filter_table)

        self.type_filter = QComboBox()
        self.type_filter.addItem("Все типы")
        self.type_filter.currentTextChanged.connect(self.filter_table)

        filter_panel.addWidget(QLabel("Поиск:"))
        filter_panel.addWidget(self.search_input)
        filter_panel.addWidget(QLabel("Группа:"))
        filter_panel.addWidget(self.group_filter)
        filter_panel.addWidget(QLabel("Тип:"))
        filter_panel.addWidget(self.type_filter)
        filter_panel.addStretch()

        top_layout.addLayout(filter_panel)

        # Таблица проверок
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "✓", "Название", "Группа", "Тип", "Статус", "Описание", "Действия"
        ])

        # Настройка заголовков
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.Fixed)

        self.table.setColumnWidth(0, 40)  # Чекбокс
        self.table.setColumnWidth(2, 150)  # Группа
        self.table.setColumnWidth(3, 180)  # Тип
        self.table.setColumnWidth(4, 100)  # Статус
        self.table.setColumnWidth(6, 200)  # Действия

        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setAlternatingRowColors(True)
        self.table.setSortingEnabled(True)

        # Контекстное меню
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_table_context_menu)

        top_layout.addWidget(self.table)

        # Нижняя панель: предпросмотр и статистика
        bottom_panel = QTabWidget()

        # Вкладка предпросмотра YAML
        preview_tab = QWidget()
        preview_layout = QVBoxLayout(preview_tab)

        preview_label = QLabel("Предпросмотр YAML:")
        preview_layout.addWidget(preview_label)

        self.yaml_preview = QTextEdit()
        self.yaml_preview.setReadOnly(True)
        self.yaml_preview.setFont(QFont("Consolas", 9))
        preview_layout.addWidget(self.yaml_preview)

        # Вкладка статистики
        stats_tab = QWidget()
        stats_layout = QVBoxLayout(stats_tab)

        self.stats_group = QGroupBox("Статистика")
        stats_inner_layout = QVBoxLayout(self.stats_group)

        self.total_label = QLabel("Всего проверок: 0")
        self.groups_label = QLabel("Групп: 0")
        self.types_label = QLabel("Типов проверок: 0")
        self.enabled_label = QLabel("Включено: 0")

        stats_inner_layout.addWidget(self.total_label)
        stats_inner_layout.addWidget(self.groups_label)
        stats_inner_layout.addWidget(self.types_label)
        stats_inner_layout.addWidget(self.enabled_label)

        stats_layout.addWidget(self.stats_group)
        stats_layout.addStretch()

        bottom_panel.addTab(preview_tab, "YAML предпросмотр")
        bottom_panel.addTab(stats_tab, "Статистика")

        # Добавляем панели в разделитель
        splitter.addWidget(top_panel)
        splitter.addWidget(bottom_panel)
        splitter.setSizes([500, 300])

        main_layout.addWidget(splitter)

        # Кнопки
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

        # Кнопки
        actions = [
            ("Добавить проверку", self.add_check, "➕"),
            ("Редактировать", self.edit_selected, "✏️"),
            ("Удалить", self.delete_selected, "🗑️"),
            ("Дублировать", self.duplicate_selected, "📋"),
            (None, None, None),  # Разделитель
            ("Включить все", self.enable_all, "✅"),
            ("Выключить все", self.disable_all, "❌"),
            (None, None, None),  # Разделитель
            ("Группировать", self.group_selected, "🗂️"),
            ("Экспорт", self.export_config, "📤"),
            ("Импорт", self.import_config, "📥")
        ]

        for text, slot, icon in actions:
            if text is None:
                toolbar.addSeparator()
            else:
                btn = QPushButton(f"{icon} {text}")
                btn.clicked.connect(slot)
                toolbar.addWidget(btn)

        return toolbar

    def load_config_to_table(self):
        """Загрузка конфигурации в таблицу"""
        self.table.setRowCount(0)

        groups = self.config.get('checks', [])
        all_groups = set()
        all_types = set()

        row = 0
        for group in groups:
            group_name = group.get('group', 'Без группы')
            all_groups.add(group_name)

            for check in group.get('subchecks', []):
                self.table.insertRow(row)

                # Чекбокс включения
                checkbox_item = QTableWidgetItem()
                checkbox_item.setCheckState(Qt.CheckState.Checked)
                checkbox_item.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
                self.table.setItem(row, 0, checkbox_item)

                # Название
                name_item = QTableWidgetItem(check.get('name', 'Без названия'))
                name_item.setData(Qt.ItemDataRole.UserRole, {
                    'group_name': group_name,
                    'check_data': check,
                    'group_index': groups.index(group),
                    'check_index': group['subchecks'].index(check)
                })
                self.table.setItem(row, 1, name_item)

                # Группа
                group_item = QTableWidgetItem(group_name)
                self.table.setItem(row, 2, group_item)

                # Тип
                check_type = check.get('type', 'Неизвестно')
                type_item = QTableWidgetItem(self.get_type_description(check_type))
                self.table.setItem(row, 3, type_item)
                all_types.add(check_type)

                # Статус
                status_item = QTableWidgetItem("✅ Включена")
                status_item.setForeground(QColor("#00cc66"))
                self.table.setItem(row, 4, status_item)

                # Описание
                desc = check.get('description', 'Без описания')
                desc_item = QTableWidgetItem(desc[:100] + "..." if len(desc) > 100 else desc)
                desc_item.setToolTip(desc)
                self.table.setItem(row, 5, desc_item)

                # Действия
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

                self.table.setCellWidget(row, 6, action_widget)

                row += 1

        # Обновляем фильтры
        self.group_filter.clear()
        self.group_filter.addItem("Все группы")
        self.group_filter.addItems(sorted(all_groups))

        self.type_filter.clear()
        self.type_filter.addItem("Все типы")
        self.type_filter.addItems(sorted(all_types))

        self.update_yaml_preview()

    def get_type_description(self, check_type):
        """Получить описательное название типа проверки"""
        type_map = {
            'no_text_present': 'Текст не должен присутствовать',
            'text_present': 'Текст должен присутствовать',
            'fuzzy_text_present': 'Нечеткое соответствие',
            'version_comparison': 'Проверка версий',
            'combined_check': 'Комбинированная проверка',
            'text_present_in_any_table': 'Текст в таблицах'
        }
        return type_map.get(check_type, check_type)

    def filter_table(self):
        """Фильтрация таблицы"""
        search_text = self.search_input.text().lower()
        group_filter = self.group_filter.currentText()
        type_filter = self.type_filter.currentText()

        for row in range(self.table.rowCount()):
            show_row = True

            if search_text:
                row_text = ""
                for col in [1, 2, 3, 5]:  # Название, группа, тип, описание
                    item = self.table.item(row, col)
                    if item:
                        row_text += item.text().lower() + " "

                if search_text not in row_text:
                    show_row = False

            if group_filter != "Все группы":
                group_item = self.table.item(row, 2)
                if group_item and group_item.text() != group_filter:
                    show_row = False

            if type_filter != "Все типы":
                type_item = self.table.item(row, 3)
                if type_item and self.get_type_description(type_filter) != type_item.text():
                    show_row = False

            self.table.setRowHidden(row, not show_row)

    def update_yaml_preview(self):
        """Обновление предпросмотра YAML"""
        if self.config:
            yaml_text = yaml.dump(self.config, allow_unicode=True, default_flow_style=False, indent=2)
            self.yaml_preview.setPlainText(yaml_text)

    def update_stats(self):
        """Обновление статистики"""
        groups = self.config.get('checks', [])
        total_checks = sum(len(group.get('subchecks', [])) for group in groups)

        all_groups = set()
        all_types = set()

        for group in groups:
            all_groups.add(group.get('group', ''))
            for check in group.get('subchecks', []):
                all_types.add(check.get('type', ''))

        self.total_label.setText(f"Всего проверок: {total_checks}")
        self.groups_label.setText(f"Групп: {len(all_groups)}")
        self.types_label.setText(f"Типов проверок: {len(all_types)}")
        self.enabled_label.setText(f"Включено: {total_checks}")  # TODO: считать реально включенные

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
            enable_action.triggered.connect(lambda: self.set_selected_state(Qt.CheckState.Checked))
            menu.addAction(enable_action)

            disable_action = QAction("❌ Выключить", self)
            disable_action.triggered.connect(lambda: self.set_selected_state(Qt.CheckState.Unchecked))
            menu.addAction(disable_action)

            menu.addSeparator()

        add_action = QAction("➕ Добавить проверку", self)
        add_action.triggered.connect(self.add_check)
        menu.addAction(add_action)

        menu.exec(self.table.viewport().mapToGlobal(position))

    def add_check(self):
        """Добавить новую проверку"""
        dialog = AddCheckDialog(self)
        if dialog.exec():
            check_data = dialog.get_check_data()
            self.add_check_to_config(check_data)
            self.load_config_to_table()
            self.update_stats()

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
        item = self.table.item(row, 1)
        if not item:
            return

        data = item.data(Qt.ItemDataRole.UserRole)
        check_data = data['check_data'].copy()
        check_data['group'] = data['group_name']

        dialog = AddCheckDialog(self, check_data)
        if dialog.exec():
            if hasattr(dialog, 'deleted') and dialog.deleted:
                # Удаление
                self.remove_check_from_config(data)
            else:
                # Обновление
                updated_data = dialog.get_check_data()
                self.update_check_in_config(data, updated_data)

            self.load_config_to_table()
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
            # Удаляем в обратном порядке, чтобы индексы не сбивались
            for index in sorted(selected_rows, key=lambda x: x.row(), reverse=True):
                self.delete_check_by_row(index.row())

            self.load_config_to_table()
            self.update_stats()

    def delete_check_by_row(self, row):
        """Удалить проверку по строке"""
        item = self.table.item(row, 1)
        if not item:
            return

        data = item.data(Qt.ItemDataRole.UserRole)
        self.remove_check_from_config(data)

    def duplicate_selected(self):
        """Дублировать выбранные проверки"""
        selected_rows = self.table.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "Ошибка", "Выберите проверки для дублирования")
            return

        for index in selected_rows:
            self.duplicate_check_by_row(index.row())

        self.load_config_to_table()
        self.update_stats()

    def duplicate_check_by_row(self, row):
        """Дублировать проверку по строке"""
        item = self.table.item(row, 1)
        if not item:
            return

        data = item.data(Qt.ItemDataRole.UserRole)
        check_data = data['check_data'].copy()

        # Создаем новое имя
        check_name = check_data.get('name', '')
        check_data['name'] = f"{check_name} (копия)"
        check_data['group'] = data['group_name']

        self.add_check_to_config(check_data)

    def set_selected_state(self, state):
        """Установить состояние выбранных проверок"""
        selected_rows = self.table.selectionModel().selectedRows()
        for index in selected_rows:
            checkbox = self.table.item(index.row(), 0)
            if checkbox:
                checkbox.setCheckState(state)

    def enable_all(self):
        """Включить все проверки"""
        for row in range(self.table.rowCount()):
            checkbox = self.table.item(row, 0)
            if checkbox:
                checkbox.setCheckState(Qt.CheckState.Checked)

    def disable_all(self):
        """Выключить все проверки"""
        for row in range(self.table.rowCount()):
            checkbox = self.table.item(row, 0)
            if checkbox:
                checkbox.setCheckState(Qt.CheckState.Unchecked)

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
                item = self.table.item(index.row(), 1)
                if item:
                    data = item.data(Qt.ItemDataRole.UserRole)
                    # Обновляем группу в данных
                    data['group_name'] = group_name
                    item.setData(Qt.ItemDataRole.UserRole, data)
                    # Обновляем отображение
                    self.table.item(index.row(), 2).setText(group_name)

            # Перезагружаем таблицу для обновления структуры конфига
            self.rebuild_config_from_table()
            self.load_config_to_table()
            self.update_stats()

    def export_config(self):
        """Экспортировать конфигурацию"""
        from PyQt6.QtWidgets import QFileDialog
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Экспорт конфигурации", "checks_config.yaml", "YAML files (*.yaml *.yml);;All files (*.*)"
        )

        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    yaml.dump(self.config, f, allow_unicode=True, default_flow_style=False)
                QMessageBox.information(self, "Успех", f"Конфигурация экспортирована в:\n{file_path}")
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось экспортировать:\n{str(e)}")

    def import_config(self):
        """Импортировать конфигурацию"""
        from PyQt6.QtWidgets import QFileDialog
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Импорт конфигурации", "", "YAML files (*.yaml *.yml);;All files (*.*)"
        )

        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    imported_config = yaml.safe_load(f)

                # Объединяем с текущей конфигурацией
                self.merge_configs(imported_config)

                self.load_config_to_table()
                self.update_stats()
                QMessageBox.information(self, "Успех", f"Конфигурация импортирована из:\n{file_path}")

            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось импортировать:\n{str(e)}")

    def merge_configs(self, imported_config):
        """Объединить импортированную конфигурацию с текущей"""
        imported_checks = imported_config.get('checks', [])
        current_checks = self.config.get('checks', [])

        # Создаем словарь текущих групп для быстрого доступа
        current_groups = {group['group']: group for group in current_checks}

        for imported_group in imported_checks:
            group_name = imported_group.get('group', '')

            if group_name in current_groups:
                # Группа существует, добавляем проверки
                existing_group = current_groups[group_name]
                existing_subchecks = existing_group.get('subchecks', [])

                # Добавляем только новые проверки
                existing_check_names = {check['name'] for check in existing_subchecks}
                for check in imported_group.get('subchecks', []):
                    if check['name'] not in existing_check_names:
                        existing_subchecks.append(check)

                existing_group['subchecks'] = existing_subchecks
            else:
                # Новая группа, добавляем полностью
                current_checks.append(imported_group)
                current_groups[group_name] = imported_group

        self.config['checks'] = current_checks

    def add_check_to_config(self, check_data):
        """Добавить проверку в конфигурацию"""
        group_name = check_data.pop('group')

        # Находим группу или создаем новую
        target_group = None
        for group in self.config.get('checks', []):
            if group.get('group') == group_name:
                target_group = group
                break

        if not target_group:
            target_group = {'group': group_name, 'subchecks': []}
            self.config.setdefault('checks', []).append(target_group)

        target_group['subchecks'].append(check_data)

    def remove_check_from_config(self, data):
        """Удалить проверку из конфигурации"""
        group_index = data['group_index']
        check_index = data['check_index']

        if 0 <= group_index < len(self.config.get('checks', [])):
            group = self.config['checks'][group_index]
            if 0 <= check_index < len(group.get('subchecks', [])):
                group['subchecks'].pop(check_index)

                # Если группа пустая, удаляем ее
                if not group['subchecks']:
                    self.config['checks'].pop(group_index)

    def update_check_in_config(self, old_data, new_data):
        """Обновить проверку в конфигурации"""
        # Удаляем старую
        self.remove_check_from_config(old_data)

        # Добавляем новую
        self.add_check_to_config(new_data)

    def rebuild_config_from_table(self):
        """Перестроить конфигурацию из таблицы"""
        new_checks = {}

        for row in range(self.table.rowCount()):
            if self.table.isRowHidden(row):
                continue

            item = self.table.item(row, 1)
            if not item:
                continue

            data = item.data(Qt.ItemDataRole.UserRole)
            group_name = data['group_name']
            check_data = data['check_data'].copy()

            if group_name not in new_checks:
                new_checks[group_name] = []

            new_checks[group_name].append(check_data)

        # Преобразуем в формат конфигурации
        self.config['checks'] = [
            {'group': group_name, 'subchecks': checks}
            for group_name, checks in new_checks.items()
        ]

    def accept_changes(self):
        """Принять изменения"""
        self.apply_changes()
        self.accept()

    def apply_changes(self):
        """Применить изменения"""
        self.rebuild_config_from_table()
        self.config_changed.emit(self.config)
        self.update_yaml_preview()
        QMessageBox.information(self, "Успех", "Изменения применены!")