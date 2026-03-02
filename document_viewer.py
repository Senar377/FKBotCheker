# document_viewer.py - ОБНОВЛЕННЫЙ КОД С ПРОЦЕНТАМИ СХОЖЕСТИ ДЛЯ КАЖДОЙ ОШИБКИ
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QTextEdit, QLineEdit, QSpinBox,
    QMessageBox, QToolTip, QWidget,
    QSplitter, QGroupBox, QTextBrowser
)
from PyQt6.QtCore import Qt, QTimer, QPoint
from PyQt6.QtGui import (
    QFont, QColor, QTextCursor, QTextCharFormat,
    QTextOption, QMouseEvent
)
import re
import logging

logger = logging.getLogger(__name__)


class DocumentViewer(QDialog):
    """Окно просмотра документа с подсветкой ошибок и подсказками по исправлению"""

    def __init__(self, parent=None, document_text="", results=None):
        super().__init__(parent)
        self.document_text = document_text or ""
        self.results = results or []

        if not self.document_text:
            logger.error("DocumentViewer: document_text is empty")

        logger.info(f"DocumentViewer: document length={len(self.document_text)}, "
                    f"results count={len(self.results)}")

        self.setWindowTitle("Просмотр документа с подсветкой ошибок")
        self.resize(1200, 800)

        # Словарь для связи ошибок с их описаниями
        self.error_details = {}
        self.current_error_index = 0
        self.search_terms_info = {}  # Словарь с информацией о том, что искала проверка

        self.init_ui()
        self.collect_search_terms_info()

    def collect_search_terms_info(self):
        """Собираем информацию о том, что искала каждая проверка"""
        for result_idx, result in enumerate(self.results):
            check_name = result.get('name', 'Неизвестная проверка')
            check_type = result.get('type', '')

            # Получаем информацию о том, что искала проверка
            search_info = self.extract_search_info_from_result(result)
            self.search_terms_info[check_name] = {
                'type': check_type,
                'search_info': search_info,
                'message': result.get('message', ''),
                'passed': result.get('passed', False),
                'result': result  # Сохраняем весь результат для использования
            }

            logger.debug(f"Собрана информация для проверки '{check_name}': {search_info}")

    def extract_search_info_from_result(self, result):
        """Извлекает информацию о том, что искала проверка"""
        check_type = result.get('type', '')
        search_info = []

        if check_type == 'no_text_present':
            terms = result.get('search_terms', [])
            if terms:
                search_info.append(f"Запрещенные слова: {', '.join(terms[:10])}")
                if len(terms) > 10:
                    search_info[-1] += f"... (всего {len(terms)})"

        elif check_type == 'text_present':
            terms = result.get('search_terms', [])
            if terms:
                search_info.append(f"Обязательные слова: {', '.join(terms[:10])}")
                if len(terms) > 10:
                    search_info[-1] += f"... (всего {len(terms)})"

        elif check_type in ['fuzzy_text_present', 'no_fuzzy_text_present']:
            text = result.get('search_text', '')
            if text:
                if len(text) > 100:
                    short_text = text[:100] + "..."
                else:
                    short_text = text
                search_info.append(f"Текст для поиска: \"{short_text}\"")

            # Добавляем информацию о найденных совпадениях с процентами
            detailed_matches = result.get('detailed_matches', [])
            if detailed_matches:
                search_info.append(f"Найдено совпадений: {len(detailed_matches)}")
                for i, match in enumerate(detailed_matches[:3], 1):
                    search_info.append(f"  {i}. Схожесть: {match['best_match_score']:.1f}% ({match['match_quality']})")
                if len(detailed_matches) > 3:
                    search_info.append(f"  ... и еще {len(detailed_matches) - 3}")

        elif check_type == 'text_present_without':
            positive = result.get('search_terms', [])
            negative = []
            if 'without_aliases' in result:
                negative = result['without_aliases']

            if positive:
                search_info.append(f"Должны быть: {', '.join(positive[:5])}")
                if len(positive) > 5:
                    search_info[-1] += f"... (всего {len(positive)})"

            if negative:
                search_info.append(f"Не должны быть: {', '.join(negative[:5])}")
                if len(negative) > 5:
                    search_info[-1] += f"... (всего {len(negative)})"

        elif check_type == 'version_comparison':
            section_results = result.get('section_results', [])
            for section in section_results:
                name = section.get('name', 'Показатель')
                required = section.get('required_version', '')
                found = section.get('found_version', '')
                if required or found:
                    info = f"{name}: требуется ≥ {required}"
                    if found:
                        info += f", найдено {found}"
                    search_info.append(info)

        elif check_type == 'combined_check':
            conditions = result.get('condition_results', [])
            search_info.append(f"Комбинированная проверка ({len(conditions)} условий)")
            for i, condition in enumerate(conditions, 1):
                name = condition.get('name', f'Условие {i}')
                passed = condition.get('passed', False)
                status = "✓" if passed else "✗"
                search_info.append(f"  {status} {name}")

        elif check_type == 'text_present_in_any_table':
            terms = result.get('search_terms', [])
            if terms:
                search_info.append(f"Поиск в таблицах: {', '.join(terms[:8])}")
                if len(terms) > 8:
                    search_info[-1] += f"... (всего {len(terms)})"

        elif check_type == 'fuzzy_text_present_after_any_table':
            text = result.get('search_text', '')
            if text:
                if len(text) > 80:
                    short_text = text[:80] + "..."
                else:
                    short_text = text
                search_info.append(f"Текст после таблиц: \"{short_text}\"")

            threshold = result.get('score', 0)
            search_info.append(f"Порог схожести: {threshold:.1f}%")

        # Если не собрали информацию, используем общее сообщение
        if not search_info:
            message = result.get('message', '')
            if message:
                search_info.append(f"Результат: {message}")
            else:
                search_info.append("Информация о поиске недоступна")

        return search_info

    def init_ui(self):
        main_layout = QVBoxLayout()
        main_layout.setContentsMargins(10, 10, 10, 10)
        main_layout.setSpacing(10)

        # ========== ПАНЕЛЬ УПРАВЛЕНИЯ ==========
        control_panel = QHBoxLayout()
        control_panel.setSpacing(10)

        # Навигация по страницам
        nav_layout = QHBoxLayout()
        nav_layout.setSpacing(5)

        self.prev_page_btn = QPushButton("←")
        self.prev_page_btn.clicked.connect(self.prev_page)
        self.prev_page_btn.setEnabled(False)
        self.prev_page_btn.setMaximumWidth(40)

        self.page_spin = QSpinBox()
        self.page_spin.setMinimum(1)
        self.page_spin.setMaximum(1)
        self.page_spin.valueChanged.connect(self.go_to_page)
        self.page_spin.setMinimumWidth(60)

        self.page_label = QLabel("из 1")
        self.page_label.setMinimumWidth(40)

        self.next_page_btn = QPushButton("→")
        self.next_page_btn.clicked.connect(self.next_page)
        self.next_page_btn.setEnabled(False)
        self.next_page_btn.setMaximumWidth(40)

        nav_layout.addWidget(QLabel("Страница:"))
        nav_layout.addWidget(self.prev_page_btn)
        nav_layout.addWidget(self.page_spin)
        nav_layout.addWidget(self.page_label)
        nav_layout.addWidget(self.next_page_btn)

        # Навигация по ошибкам
        error_nav_layout = QHBoxLayout()
        error_nav_layout.setSpacing(5)

        self.prev_error_btn = QPushButton("← Ошибка")
        self.prev_error_btn.clicked.connect(self.prev_error)
        self.prev_error_btn.setEnabled(False)
        self.prev_error_btn.setMinimumWidth(100)

        self.error_info_label = QLabel("Ошибка: 0/0")
        self.error_info_label.setMinimumWidth(120)

        self.next_error_btn = QPushButton("Ошибка →")
        self.next_error_btn.clicked.connect(self.next_error)
        self.next_error_btn.setEnabled(False)
        self.next_error_btn.setMinimumWidth(100)

        error_nav_layout.addWidget(self.prev_error_btn)
        error_nav_layout.addWidget(self.error_info_label)
        error_nav_layout.addWidget(self.next_error_btn)

        # Поиск
        search_layout = QHBoxLayout()
        search_layout.setSpacing(5)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Поиск в документе...")
        self.search_input.returnPressed.connect(self.search_text)

        self.search_prev_btn = QPushButton("←")
        self.search_prev_btn.clicked.connect(self.prev_search_result)
        self.search_prev_btn.setEnabled(False)
        self.search_prev_btn.setMaximumWidth(30)

        self.search_next_btn = QPushButton("→")
        self.search_next_btn.clicked.connect(self.next_search_result)
        self.search_next_btn.setEnabled(False)
        self.search_next_btn.setMaximumWidth(30)

        self.search_count_label = QLabel("Найдено: 0")
        self.search_count_label.setMinimumWidth(80)

        search_layout.addWidget(QLabel("Поиск:"))
        search_layout.addWidget(self.search_input)
        search_layout.addWidget(self.search_prev_btn)
        search_layout.addWidget(self.search_next_btn)
        search_layout.addWidget(self.search_count_label)

        # Кнопки просмотра ошибок
        error_layout = QHBoxLayout()
        error_layout.setSpacing(5)

        self.show_all_btn = QPushButton("Показать все ошибки")
        self.show_all_btn.clicked.connect(self.show_all_errors)

        self.clear_btn = QPushButton("Очистить подсветку")
        self.clear_btn.clicked.connect(self.clear_highlights)

        error_layout.addWidget(self.show_all_btn)
        error_layout.addWidget(self.clear_btn)

        # Сборка панели управления
        control_panel.addLayout(nav_layout)
        control_panel.addSpacing(10)
        control_panel.addLayout(error_nav_layout)
        control_panel.addSpacing(10)
        control_panel.addLayout(search_layout)
        control_panel.addSpacing(10)
        control_panel.addLayout(error_layout)
        control_panel.addStretch()

        main_layout.addLayout(control_panel)

        # ========== РАЗДЕЛИТЕЛЬ С ВОЗМОЖНОСТЬЮ ИЗМЕНЕНИЯ РАЗМЕРА ==========
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # ========== ЛЕВАЯ ПАНЕЛЬ: ДОКУМЕНТ ==========
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 0, 0)

        doc_label = QLabel("📄 Документ:")
        doc_label.setStyleSheet("font-weight: bold; font-size: 12px; color: #2c3e50;")
        left_layout.addWidget(doc_label)

        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setFont(QFont("Times New Roman", 11))

        # Включаем отслеживание движения мыши
        self.text_edit.viewport().setMouseTracking(True)
        self.text_edit.viewport().installEventFilter(self)

        # Сохраняем форматирование абзацев
        self.text_edit.setAcceptRichText(True)
        self.text_edit.setLineWrapMode(QTextEdit.LineWrapMode.WidgetWidth)
        self.text_edit.setWordWrapMode(QTextOption.WrapMode.WordWrap)

        left_layout.addWidget(self.text_edit)

        splitter.addWidget(left_panel)

        # ========== ПРАВАЯ ПАНЕЛЬ: ИНФОРМАЦИЯ О ПРОВЕРКЕ ==========
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(5, 5, 5, 5)
        right_layout.setSpacing(10)

        info_label = QLabel("🔍 Информация о проверке:")
        info_label.setStyleSheet("font-weight: bold; font-size: 12px; color: #2c3e50;")
        right_layout.addWidget(info_label)

        # ========== БЛОК: ЧТО ИСКАЛА ПРОВЕРКА ==========
        self.search_terms_group = QGroupBox("Что проверялось:")
        self.search_terms_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 12px;
                color: #2c3e50;
                border: 2px solid #3498db;
                border-radius: 8px;
                margin-top: 5px;
                padding-top: 15px;
                background-color: #f8f9fa;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 8px 0 8px;
                background-color: #3498db;
                color: white;
                border-radius: 4px;
            }
        """)

        search_terms_layout = QVBoxLayout(self.search_terms_group)

        self.search_terms_text = QTextBrowser()
        self.search_terms_text.setStyleSheet("""
            QTextBrowser {
                background-color: #ffffff;
                border: 1px solid #d6dbdf;
                border-radius: 5px;
                padding: 10px;
                font-size: 11px;
                color: #2c3e50;
                font-family: 'Arial', sans-serif;
                max-height: 150px;
            }
        """)
        self.search_terms_text.setReadOnly(True)
        search_terms_layout.addWidget(self.search_terms_text)

        right_layout.addWidget(self.search_terms_group)

        # ========== БЛОК: ПРИМЕР ИСПРАВЛЕНИЯ ==========
        correction_group = QGroupBox("Пример исправления:")
        correction_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 12px;
                color: #2c3e50;
                border: 2px solid #3498db;
                border-radius: 8px;
                margin-top: 5px;
                padding-top: 15px;
                background-color: #f8f9fa;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 8px 0 8px;
                background-color: #3498db;
                color: white;
                border-radius: 4px;
            }
        """)

        correction_layout = QVBoxLayout(correction_group)

        # Пример оригинальной строки (с ошибкой)
        self.original_example_label = QLabel("Обнаружено:")
        self.original_example_label.setStyleSheet("font-weight: bold; color: #d32f2f; font-size: 11px;")
        correction_layout.addWidget(self.original_example_label)

        self.original_example_text = QTextBrowser()
        self.original_example_text.setMaximumHeight(70)
        self.original_example_text.setStyleSheet("""
            QTextBrowser {
                background-color: #ffebee;
                border: 1px solid #ffcdd2;
                border-radius: 5px;
                padding: 8px;
                font-family: 'Courier New', monospace;
                font-size: 11px;
                color: #c62828;
            }
        """)
        self.original_example_text.setReadOnly(True)
        correction_layout.addWidget(self.original_example_text)

        # Стрелка преобразования
        arrow_label = QLabel("↓")
        arrow_label.setStyleSheet("""
            font-size: 20px; 
            font-weight: bold; 
            color: #3498db;
            qproperty-alignment: 'AlignCenter';
        """)
        correction_layout.addWidget(arrow_label)

        # Пример исправленной строки
        self.correct_example_label = QLabel("Исправить на:")
        self.correct_example_label.setStyleSheet("font-weight: bold; color: #2e7d32; font-size: 11px;")
        correction_layout.addWidget(self.correct_example_label)

        self.correct_example_text = QTextBrowser()
        self.correct_example_text.setMaximumHeight(70)
        self.correct_example_text.setStyleSheet("""
            QTextBrowser {
                background-color: #e8f5e8;
                border: 1px solid #c8e6c9;
                border-radius: 5px;
                padding: 8px;
                font-family: 'Courier New', monospace;
                font-size: 11px;
                color: #1b5e20;
            }
        """)
        self.correct_example_text.setReadOnly(True)
        correction_layout.addWidget(self.correct_example_text)

        right_layout.addWidget(correction_group)

        # ========== БЛОК: ОБЪЯСНЕНИЕ ==========
        explanation_group = QGroupBox("Объяснение:")
        explanation_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 12px;
                color: #2c3e50;
                border: 2px solid #3498db;
                border-radius: 8px;
                margin-top: 5px;
                padding-top: 15px;
                background-color: #f8f9fa;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 8px 0 8px;
                background-color: #3498db;
                color: white;
                border-radius: 4px;
            }
        """)

        explanation_layout = QVBoxLayout(explanation_group)

        self.explanation_text = QTextBrowser()
        self.explanation_text.setStyleSheet("""
            QTextBrowser {
                background-color: #ffffff;
                border: 1px solid #d6dbdf;
                border-radius: 5px;
                padding: 10px;
                font-size: 11px;
                color: #2c3e50;
                max-height: 120px;
            }
        """)
        self.explanation_text.setReadOnly(True)
        explanation_layout.addWidget(self.explanation_text)

        right_layout.addWidget(explanation_group)

        # Добавляем растягивающийся элемент для правильного расположения
        right_layout.addStretch()

        splitter.addWidget(right_panel)

        # Устанавливаем начальные размеры панелей
        splitter.setSizes([700, 500])

        main_layout.addWidget(splitter)

        # ========== ПАНЕЛЬ ИНФОРМАЦИИ ОБ ОШИБКЕ ==========
        self.error_detail_panel = QWidget()
        error_detail_layout = QVBoxLayout(self.error_detail_panel)
        error_detail_layout.setContentsMargins(10, 5, 10, 5)
        error_detail_layout.setSpacing(5)

        # Верхняя строка с названием ошибки
        title_row = QHBoxLayout()
        self.error_title_label = QLabel()
        self.error_title_label.setStyleSheet("font-weight: bold; color: #d32f2f; font-size: 12px;")
        self.error_title_label.setWordWrap(True)
        title_row.addWidget(self.error_title_label)
        title_row.addStretch()

        error_detail_layout.addLayout(title_row)

        # Основное описание ошибки
        self.error_desc_label = QLabel()
        self.error_desc_label.setStyleSheet("color: #555; font-size: 11px;")
        self.error_desc_label.setWordWrap(True)
        error_detail_layout.addWidget(self.error_desc_label)

        # ПАНЕЛЬ С ТЕМ, ЧТО ИСКАЛА ПРОВЕРКА
        self.search_info_panel = QWidget()
        search_info_layout = QVBoxLayout(self.search_info_panel)
        search_info_layout.setContentsMargins(8, 5, 8, 5)
        search_info_layout.setSpacing(3)

        self.search_info_label = QLabel()
        self.search_info_label.setStyleSheet("""
            font-size: 10px;
            color: #2c3e50;
            background-color: #f8f9fa;
            border-left: 3px solid #3498db;
            padding: 4px 8px;
            border-radius: 2px;
        """)
        self.search_info_label.setWordWrap(True)
        self.search_info_label.setTextFormat(Qt.TextFormat.RichText)

        search_info_layout.addWidget(self.search_info_label)
        error_detail_layout.addWidget(self.search_info_panel)

        # Позиция ошибки
        self.error_position_label = QLabel()
        self.error_position_label.setStyleSheet("color: #777; font-size: 10px;")
        self.error_position_label.setWordWrap(True)
        error_detail_layout.addWidget(self.error_position_label)

        self.error_detail_panel.setVisible(False)
        self.error_detail_panel.setStyleSheet("""
            QWidget {
                background-color: #fff3cd;
                border: 1px solid #ffeaa7;
                border-radius: 5px;
                padding: 5px;
            }
        """)

        main_layout.addWidget(self.error_detail_panel)

        # ========== СТАТИСТИКА ==========
        stats_panel = QHBoxLayout()
        self.stats_label = QLabel()
        self.stats_label.setStyleSheet("font-size: 11px; color: #666;")
        stats_panel.addWidget(self.stats_label)
        stats_panel.addStretch()

        main_layout.addLayout(stats_panel)

        # ========== КНОПКА ЗАКРЫТИЯ ==========
        close_btn = QPushButton("Закрыть")
        close_btn.clicked.connect(self.accept)
        close_btn.setMinimumHeight(35)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                font-weight: bold;
                border: none;
                border-radius: 5px;
                padding: 8px 20px;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        main_layout.addWidget(close_btn)

        self.setLayout(main_layout)

        # Инициализация переменных
        self.current_page = 1
        self.pages = []
        self.page_ranges = []  # (start_pos, end_pos) для каждой страницы
        self.all_matches = []
        self.search_matches = []
        self.current_match_index = 0
        self.current_search_index = -1
        self.error_positions = []
        self.hover_timer = QTimer()
        self.hover_timer.setSingleShot(True)
        self.hover_timer.timeout.connect(self.show_error_tooltip)
        self.hover_pos = None
        self.full_line_highlights = {}  # Словарь для подсветки полных строк

        # Установка начального текста с сохранением форматирования
        self.init_document_view()

        # Разбиваем документ на страницы
        self.split_into_pages()

        # Обновляем спинбокс
        self.page_spin.setMaximum(len(self.pages))
        self.page_label.setText(f"из {len(self.pages)}")

        # Показываем первую страницу
        self.display_current_page()

        # Собираем все совпадения ошибок (с учетом исключающих алиасов)
        self.collect_error_matches()

        # Обновляем статистику
        self.update_stats()

        # Обновляем навигацию по ошибкам
        self.update_error_navigation()

        # Устанавливаем примеры по умолчанию
        self.set_default_examples()

    def set_default_examples(self):
        """Установить примеры по умолчанию при загрузке"""
        self.search_terms_text.setPlainText("Наведите курсор на ошибку для просмотра информации о проверке...")
        self.original_example_text.setPlainText("Наведите курсор на ошибку для просмотра примера...")
        self.correct_example_text.setPlainText(
            "Здесь будет показан правильный вариант на основе того, что искала проверка")
        self.explanation_text.setPlainText("Наведите курсор на выделенную ошибку для получения объяснения")

    def eventFilter(self, obj, event):
        """Обработка событий мыши для показа подсказок"""
        if obj == self.text_edit.viewport() and event.type() == QMouseEvent.Type.MouseMove:
            self.handle_mouse_move(event)
        return super().eventFilter(obj, event)

    def handle_mouse_move(self, event):
        """Обработка движения мыши для показа подсказок об ошибках"""
        cursor = self.text_edit.cursorForPosition(event.pos())
        position = cursor.position()

        # Проверяем, находится ли курсор над ошибкой
        current_error = None
        for error_id, error_info in self.error_details.items():
            if error_info['page'] == self.current_page:
                start = error_info['start']
                end = error_info['end']
                if start <= position <= end:
                    current_error = error_info
                    break

        if current_error:
            # Показываем панель с информацией об ошибке
            self.error_detail_panel.setVisible(True)
            self.error_title_label.setText(f"Ошибка: {current_error['name']}")

            status = "ПРОВАЛЕНО" if not current_error['passed'] else "ТРЕБУЕТ ПРОВЕРКИ"
            status_color = "#d32f2f" if not current_error['passed'] else "#f57c00"

            # Добавляем информацию о проценте схожести, если есть
            similarity_info = ""
            if current_error.get('similarity_score') is not None:
                score = current_error['similarity_score']
                quality = current_error.get('match_quality', '')
                similarity_info = f"<br><b>Схожесть:</b> {score:.1f}% ({quality})"

            self.error_desc_label.setText(
                f"<span style='color: {status_color}; font-weight: bold;'>{status}</span>{similarity_info}<br>"
                f"{current_error['message']}"
            )

            position_text = f"Страница {current_error['page']}, позиция {position - start + 1}"
            self.error_position_label.setText(f"Позиция: {position_text}")

            # Обновляем информацию о том, что искала проверка
            self.update_search_terms_info(current_error['name'])

            # Обновляем информацию в панели поиска
            search_info = self.extract_search_info_from_result_by_name(current_error['name'])
            if search_info:
                search_text = "<br>".join([f"• {info}" for info in search_info])

                # Добавляем информацию о проценте схожести для этого конкретного вхождения
                if current_error.get('similarity_score') is not None:
                    search_text += f"<br><br><b>Для данного вхождения:</b> Схожесть {current_error['similarity_score']:.1f}%"
                    if current_error.get('word_scores'):
                        search_text += "<br><b>По словам:</b>"
                        for word_info in current_error['word_scores'][:5]:
                            search_text += f"<br>  • {word_info['word']}: {word_info['score']:.1f}%"

                self.search_info_label.setText(f"<b>Что искала проверка:</b><br>{search_text}")

            # Обновляем примеры и объяснение
            self.update_examples_and_explanation(current_error)

            # Останавливаем таймер подсказки
            self.hover_timer.stop()
        else:
            # Скрываем панель, если не над ошибкой
            self.error_detail_panel.setVisible(False)

            # Запускаем таймер для показа всплывающей подсказки
            self.hover_pos = event.pos()
            self.hover_timer.start(500)

    def extract_search_info_from_result_by_name(self, check_name):
        """Извлекает информацию о поиске по имени проверки"""
        if check_name in self.search_terms_info:
            return self.search_terms_info[check_name]['search_info']
        return []

    def update_search_terms_info(self, error_name):
        """Обновляет информацию о том, что искала проверка"""
        if error_name in self.search_terms_info:
            info = self.search_terms_info[error_name]
            search_info = info['search_info']

            # Форматируем информацию для отображения
            html_lines = []

            # Заголовок с типом проверки
            check_type = info['type']
            check_type_display = {
                'no_text_present': '❌ Проверка на отсутствие текста',
                'text_present': '✅ Проверка на наличие текста',
                'fuzzy_text_present': '🔍 Нечеткий поиск текста',
                'no_fuzzy_text_present': '🚫 Нечеткая проверка отсутствия',
                'text_present_without': '⚠ Проверка на наличие и отсутствие',
                'version_comparison': '📊 Проверка показателей назначения',
                'combined_check': '🔗 Комбинированная проверка',
                'text_present_in_any_table': '📋 Поиск в таблицах',
                'fuzzy_text_present_after_any_table': '🔍 Поиск после таблиц'
            }.get(check_type, f'📝 {check_type}')

            html_lines.append(f"<b>{check_type_display}</b>")

            # Добавляем информацию о поиске
            for line in search_info:
                if line.startswith("  "):  # Вложенные элементы для комбинированных проверок
                    html_lines.append(f"&nbsp;&nbsp;&nbsp;&nbsp;{line}")
                else:
                    html_lines.append(f"• {line}")

            # Добавляем результат проверки
            status = "✅ ПРОЙДЕНА" if info['passed'] else "❌ ПРОВАЛЕНА"
            status_color = "#2ecc71" if info['passed'] else "#e74c3c"

            # Показываем исключающие алиасы, если они есть
            result = info['result']
            without_aliases = result.get('without_aliases', [])
            if without_aliases:
                html_lines.append(f"<br><b>Исключающие алиасы:</b> {', '.join(without_aliases[:5])}")
                if len(without_aliases) > 5:
                    html_lines[-1] += f"... (всего {len(without_aliases)})"

            html_lines.append(f"<br><b>Результат:</b> <span style='color: {status_color};'>{status}</span>")

            # Добавляем сообщение проверки
            message = info['message']
            if message:
                html_lines.append(f"<b>Сообщение:</b> {message}")

            html_text = "<br>".join(html_lines)
            self.search_terms_text.setHtml(html_text)
        else:
            self.search_terms_text.setPlainText(f"Информация о проверке '{error_name}' не найдена")

    def update_examples_and_explanation(self, error_info):
        """Обновление примеров и объяснения для текущей ошибки"""
        error_name = error_info['name']
        found_text = error_info.get('term', '') or error_info.get('context', '')

        # Получаем актуальный контекст из документа
        start_pos = error_info.get('global_start', 0)
        end_pos = error_info.get('global_end', 0)

        if start_pos >= 0 and end_pos > start_pos:
            # Находим начало и конец строки
            line_start = self.document_text.rfind('\n', 0, start_pos) + 1
            line_end = self.document_text.find('\n', end_pos)
            if line_end == -1:
                line_end = len(self.document_text)

            # Извлекаем полную строку
            full_line = self.document_text[line_start:line_end]

            # Извлекаем найденный текст
            found_start = start_pos - line_start
            found_end = end_pos - line_start

            if 0 <= found_start < len(full_line) and found_end <= len(full_line):
                found_text = full_line[found_start:found_end]

                # Создаем пример оригинальной строки
                original_line = full_line

                # Получаем информацию о том, что искала проверка
                search_info = self.generate_correction_based_on_search_info(error_name)

                # Добавляем информацию о проценте схожести
                if error_info.get('similarity_score') is not None:
                    correction_text = f"[Схожесть: {error_info['similarity_score']:.1f}%]\n{search_info['correction']}"
                else:
                    correction_text = search_info['correction']

                # Обновляем примеры
                self.original_example_text.setPlainText(original_line.strip())
                self.correct_example_text.setPlainText(correction_text)
                self.explanation_text.setPlainText(search_info['explanation'])
                return

        # Если не удалось извлечь контекст, используем общий пример на основе типа проверки
        search_info = self.generate_correction_based_on_search_info(error_name)

        # Добавляем информацию о проценте схожести
        if error_info.get('similarity_score') is not None:
            original_display = f'Обнаружено: "{found_text[:100]}" [Схожесть: {error_info["similarity_score"]:.1f}%]'
            correction_text = f"[Схожесть: {error_info['similarity_score']:.1f}%]\n{search_info['correction']}"
        else:
            original_display = f'Обнаружено: "{found_text[:100]}"' if found_text else 'Не удалось определить контекст'
            correction_text = search_info['correction']

        self.original_example_text.setPlainText(original_display)
        self.correct_example_text.setPlainText(correction_text)
        self.explanation_text.setPlainText(search_info['explanation'])

    def generate_correction_based_on_search_info(self, error_name):
        """Генерирует пример исправления на основе того, что искала проверка"""
        if error_name in self.search_terms_info:
            info = self.search_terms_info[error_name]
            check_type = info['type']
            search_info_list = info['search_info']
            message = info['message']
            passed = info['passed']

            # Получаем исключающие алиасы из оригинального результата
            result = info['result']
            without_aliases = result.get('without_aliases', [])

            # Базовый пример исправления
            correction = ""
            explanation = message

            # Если есть исключающие алиасы и проверка text_present не пройдена
            if without_aliases and check_type == 'text_present' and not passed:
                # Проверяем, содержатся ли исключающие алиасы в документе
                found_aliases = []
                for alias in without_aliases:
                    if alias.lower() in self.document_text.lower():
                        found_aliases.append(alias)

                if found_aliases:
                    # Если найден хотя бы один исключающий алиас, проверка не должна срабатывать
                    correction = f"✅ Найдены допустимые альтернативы: {', '.join(found_aliases)}"
                    explanation = f"Обнаружены исключающие алиасы ({', '.join(found_aliases)}), поэтому проверка '{error_name}' не применяется."
                    return {
                        'correction': correction,
                        'explanation': explanation
                    }

            # Если проверка уже пройдена (passed=True), то это НЕ ошибка
            if passed:
                correction = "✅ Проверка пройдена успешно"
                explanation = "Требования проверки выполнены в полном объеме."
            else:
                # Если проверка не пройдена, генерируем рекомендации
                if check_type == 'no_text_present':
                    # Проверка на отсутствие запрещенных слов
                    if search_info_list and "Запрещенные слова:" in search_info_list[0]:
                        forbidden_words = search_info_list[0].replace("Запрещенные слова: ", "")
                        correction = f"Удалить запрещенные слова: {forbidden_words}"
                        explanation = f"Обнаружены запрещенные к использованию слова. {message}"

                elif check_type == 'text_present':
                    # Проверка на наличие обязательных слов - ДОБАВИТЬ, если не найдено
                    if search_info_list and "Обязательные слова:" in search_info_list[0]:
                        required_words = search_info_list[0].replace("Обязательные слова: ", "")
                        correction = f"Добавить обязательные слова: {required_words}"
                        explanation = f"Не найдены обязательные для документа слова. {message}"

                elif check_type == 'fuzzy_text_present':
                    # Нечеткий поиск текста - ДОБАВИТЬ или исправить
                    if search_info_list and "Текст для поиска:" in search_info_list[0]:
                        search_text = search_info_list[0].replace("Текст для поиска: ", "")
                        # Извлекаем информацию о найденных совпадениях с процентами
                        similarity_info = []
                        for line in search_info_list:
                            if "Схожесть:" in line and "%" in line:
                                similarity_info.append(line)

                        if similarity_info:
                            correction = f"Требуется текст, схожий с указанным. Найденные варианты:\n" + "\n".join(
                                similarity_info[:3])
                        else:
                            correction = f"Добавить или исправить текст: {search_text}"
                        explanation = f"Требуется текст, схожий с указанным. {message}"

                elif check_type == 'no_fuzzy_text_present':
                    # Нечеткая проверка отсутствия - УДАЛИТЬ или перефразировать
                    if search_info_list and "Текст для поиска:" in search_info_list[0]:
                        search_text = search_info_list[0].replace("Текст для поиска: ", "")
                        correction = f"Удалить или перефразировать: {search_text}"
                        explanation = f"Обнаружен текст, который не должен присутствовать. {message}"

                elif check_type == 'text_present_without':
                    # Проверка на наличие и отсутствие
                    has_required = False
                    has_forbidden = False
                    required_text = ""
                    forbidden_text = ""

                    for line in search_info_list:
                        if "Должны быть:" in line:
                            has_required = True
                            required_text = line.replace("Должны быть: ", "")
                        elif "Не должны быть:" in line:
                            has_forbidden = True
                            forbidden_text = line.replace("Не должны быть: ", "")

                    if has_required and has_forbidden:
                        correction = f"Добавить: {required_text} | Удалить: {forbidden_text}"
                    elif has_required:
                        correction = f"Добавить: {required_text}"
                    elif has_forbidden:
                        correction = f"Удалить: {forbidden_text}"

                    explanation = f"Проверка комбинации условий. {message}"

                elif check_type == 'version_comparison':
                    # Проверка версий
                    version_issues = []
                    for line in search_info_list:
                        if "требуется ≥" in line:
                            version_issues.append(line)

                    if version_issues:
                        correction = f"Обновить версии: {'; '.join(version_issues)}"
                        explanation = f"Требования к версиям не выполнены. {message}"
                    else:
                        correction = "Проверить соответствие требованиям к версиям"

                elif check_type == 'combined_check':
                    # Комбинированная проверка
                    conditions = []
                    for line in search_info_list:
                        if line.startswith("  "):
                            conditions.append(line.strip())

                    if conditions:
                        correction = f"Выполнить условия: {', '.join(conditions[:3])}"
                        if len(conditions) > 3:
                            correction += f"... (всего {len(conditions)})"
                    else:
                        correction = "Выполнить все условия проверки"

                    explanation = f"Не выполнены комбинированные условия. {message}"

                elif check_type == 'text_present_in_any_table':
                    # Поиск в таблицах
                    if search_info_list and "Поиск в таблицах:" in search_info_list[0]:
                        table_terms = search_info_list[0].replace("Поиск в таблицах: ", "")
                        correction = f"Добавить в таблицу: {table_terms}"
                        explanation = f"Требуемая информация отсутствует в таблицах. {message}"

                elif check_type == 'fuzzy_text_present_after_any_table':
                    # Поиск после таблиц
                    if search_info_list and "Текст после таблиц:" in search_info_list[0]:
                        after_table_text = search_info_list[0].replace("Текст после таблиц: ", "")
                        correction = f"Добавить после таблиц: {after_table_text}"
                        explanation = f"Требуемый текст отсутствует после таблиц. {message}"

                else:
                    # Общий случай
                    correction = "Требуется исправление на основе проверки"
                    explanation = f"{message}"

            # Если correction остался пустым, используем общий вариант
            if not correction:
                if passed:
                    correction = "✅ Проверка пройдена"
                else:
                    correction = f"Исправить согласно проверке: {error_name}"

        else:
            # Если информация о проверке не найдена
            correction = "Требуется проверка специалиста"
            explanation = f"Обнаружена проблема в проверке '{error_name}'. Необходима ручная проверка."

        return {
            'correction': correction,
            'explanation': explanation
        }

    def update_error_navigation(self):
        """Обновить состояние кнопок навигации по ошибкам"""
        error_count = len(self.error_positions)

        if error_count > 0:
            self.prev_error_btn.setEnabled(True)
            self.next_error_btn.setEnabled(True)
            self.error_info_label.setText(f"Ошибка: {self.current_error_index + 1}/{error_count}")
        else:
            self.prev_error_btn.setEnabled(False)
            self.next_error_btn.setEnabled(False)
            self.error_info_label.setText("Ошибок не найдено")

    def prev_error(self):
        """Перейти к предыдущей ошибке"""
        if self.error_positions:
            self.current_error_index = (self.current_error_index - 1) % len(self.error_positions)
            self.go_to_error(self.current_error_index)

    def next_error(self):
        """Перейти к следующей ошибке"""
        if self.error_positions:
            self.current_error_index = (self.current_error_index + 1) % len(self.error_positions)
            self.go_to_error(self.current_error_index)

    def go_to_error(self, index):
        """Перейти к указанной ошибке"""
        if 0 <= index < len(self.error_positions):
            page, pos = self.error_positions[index]

            # Если ошибка на другой странице, переключаемся на неё
            if page != self.current_page:
                self.current_page = page
                self.display_current_page()

            # Прокручиваем к ошибке
            cursor = self.text_edit.textCursor()
            cursor.setPosition(pos)
            self.text_edit.setTextCursor(cursor)
            self.text_edit.ensureCursorVisible()

            # Выделяем текущую ошибку
            self.highlight_current_error(index)

            # Обновляем информацию
            self.current_error_index = index
            self.update_error_navigation()

            # Показываем информацию об ошибке
            if index < len(self.error_details):
                error_info = list(self.error_details.values())[index]
                self.error_detail_panel.setVisible(True)
                self.error_title_label.setText(f"Ошибка: {error_info['name']}")

                status = "ПРОВАЛЕНО" if not error_info['passed'] else "ТРЕБУЕТ ПРОВЕРКИ"
                status_color = "#d32f2f" if not error_info['passed'] else "#f57c00"

                # Добавляем информацию о проценте схожести
                similarity_info = ""
                if error_info.get('similarity_score') is not None:
                    score = error_info['similarity_score']
                    quality = error_info.get('match_quality', '')
                    similarity_info = f"<br><b>Схожесть:</b> {score:.1f}% ({quality})"

                self.error_desc_label.setText(
                    f"<span style='color: {status_color}; font-weight: bold;'>{status}</span>{similarity_info}<br>"
                    f"{error_info['message']}"
                )

                self.error_position_label.setText(f"Позиция: Страница {page}, строка {self.get_line_number(pos)}")

                # Обновляем информацию о том, что искала проверка
                self.update_search_terms_info(error_info['name'])

                # Обновляем информацию в панели поиска
                search_info = self.extract_search_info_from_result_by_name(error_info['name'])
                if search_info:
                    search_text = "<br>".join([f"• {info}" for info in search_info])

                    # Добавляем информацию о проценте схожести для этого конкретного вхождения
                    if error_info.get('similarity_score') is not None:
                        search_text += f"<br><br><b>Для данного вхождения:</b> Схожесть {error_info['similarity_score']:.1f}%"
                        if error_info.get('word_scores'):
                            search_text += "<br><b>По словам:</b>"
                            for word_info in error_info['word_scores'][:5]:
                                search_text += f"<br>  • {word_info['word']}: {word_info['score']:.1f}%"

                    self.search_info_label.setText(f"<b>Что искала проверка:</b><br>{search_text}")

                # Обновляем примеры и объяснение
                self.update_examples_and_explanation(error_info)

    def highlight_current_error(self, error_index):
        """Выделить текущую ошибку специальным цветом"""
        # Сначала очищаем все выделения
        self.clear_highlights()

        # Затем подсвечиваем все ошибки с двойной подсветкой
        self.apply_highlights_to_current_page()

        # И выделяем текущую ошибку другим цветом
        if 0 <= error_index < len(self.error_positions):
            page, pos = self.error_positions[error_index]
            if page == self.current_page:
                # Находим информацию об этой ошибке
                for error_info in self.error_details.values():
                    if error_info['page'] == page and error_info['start'] <= pos <= error_info['end']:
                        # Выделяем полную строку желтым цветом
                        self.highlight_full_line(error_info)

                        # Выделяем проблемный текст красным цветом
                        cursor = self.text_edit.textCursor()
                        format = QTextCharFormat()
                        format.setBackground(QColor(255, 0, 0, 150))  # Красный
                        format.setFontWeight(QFont.Weight.Bold)
                        format.setUnderlineStyle(QTextCharFormat.UnderlineStyle.SingleUnderline)
                        format.setUnderlineColor(QColor(211, 47, 47))

                        cursor.setPosition(error_info['start'])
                        cursor.setPosition(error_info['end'], QTextCursor.MoveMode.KeepAnchor)
                        cursor.mergeCharFormat(format)

                        # Добавляем подсказку с процентом схожести
                        if error_info.get('similarity_score') is not None:
                            tooltip_format = QTextCharFormat()
                            tooltip_format.setToolTip(f"Схожесть: {error_info['similarity_score']:.1f}%")
                            cursor.mergeCharFormat(tooltip_format)
                        break

    def highlight_full_line(self, error_info):
        """Подсветить полную строку, содержащую ошибку"""
        page_idx = error_info['page'] - 1
        if 0 <= page_idx < len(self.pages):
            page_text = self.pages[page_idx]
            error_start = error_info['start']
            error_end = error_info['end']

            # Находим начало и конец строки на текущей странице
            line_start = page_text.rfind('\n', 0, error_start) + 1
            line_end = page_text.find('\n', error_end)
            if line_end == -1:
                line_end = len(page_text)

            # Подсвечиваем всю строку желтым цветом
            if line_start < line_end:
                cursor = self.text_edit.textCursor()
                format = QTextCharFormat()
                format.setBackground(QColor(255, 255, 0, 80))  # Желтый
                format.setFontWeight(QFont.Weight.Normal)

                cursor.setPosition(line_start)
                cursor.setPosition(line_end, QTextCursor.MoveMode.KeepAnchor)
                cursor.mergeCharFormat(format)

    def get_line_number(self, position):
        """Получить номер строки для позиции в текущей странице"""
        if 0 <= self.current_page - 1 < len(self.pages):
            page_text = self.pages[self.current_page - 1]
            lines_before = page_text[:position].count('\n')
            return lines_before + 1
        return 0

    def init_document_view(self):
        """Инициализация просмотра документа с сохранением форматирования"""
        # Преобразуем переносы строк в HTML-теги для сохранения форматирования
        html_text = self.document_text.replace('\n', '<br>')

        # Добавляем базовое HTML-оформление
        html_content = f"""
        <html>
        <head>
            <style>
                body {{ 
                    font-family: 'Times New Roman', serif; 
                    font-size: 11pt; 
                    line-height: 1.5;
                    margin: 20px;
                    cursor: default;
                    background-color: #ffffff;
                }}
                p {{ 
                    margin-top: 12px; 
                    margin-bottom: 12px; 
                    text-align: justify;
                }}
                .highlight-error {{
                    background-color: rgba(255, 0, 0, 0.25);
                    font-weight: bold;
                    color: #000000;
                    border-bottom: 2px dashed #d32f2f;
                    padding: 1px 2px;
                }}
                .highlight-full-line {{
                    background-color: rgba(255, 255, 0, 0.15);
                    font-weight: normal;
                    color: #000000;
                    padding: 1px 0;
                }}
                .highlight-current-error {{
                    background-color: rgba(255, 193, 7, 0.5);
                    font-weight: bold;
                    color: #000000;
                    border: 2px solid #ff9800;
                    padding: 1px 2px;
                }}
                .highlight-search {{
                    background-color: rgba(0, 100, 255, 0.3);
                    font-weight: bold;
                    color: #000000;
                    padding: 1px 2px;
                }}
                .page-number {{
                    text-align: center; 
                    font-size: 10pt; 
                    color: #666;
                    margin-top: 30px;
                    border-top: 1px solid #ddd;
                    padding-top: 10px;
                }}
            </style>
        </head>
        <body>
            {html_text}
        </body>
        </html>
        """

        self.text_edit.setHtml(html_content)

    def split_into_pages(self):
        """Улучшенное разбиение текста на страницы"""
        lines = self.document_text.split('\n')
        current_page = []
        current_length = 0
        current_start = 0
        page_length = 1800
        max_lines_per_page = 50

        for line in lines:
            line_length = len(line)
            if (current_length + line_length > page_length and current_length > 0) or \
                    (len(current_page) >= max_lines_per_page) or \
                    (line.strip() == '' and len(current_page) > 30):
                page_text = '\n'.join(current_page)
                self.pages.append(page_text)
                self.page_ranges.append((current_start, current_start + len(page_text)))
                current_page = [line]
                current_length = line_length
                current_start += len(page_text) + 1
            else:
                current_page.append(line)
                current_length += line_length + 1

        if current_page:
            page_text = '\n'.join(current_page)
            self.pages.append(page_text)
            self.page_ranges.append((current_start, current_start + len(page_text)))

        if not self.pages:
            self.pages = [self.document_text]
            self.page_ranges = [(0, len(self.document_text))]

        logger.info(f"Документ разбит на {len(self.pages)} страниц")

    def display_current_page(self):
        """Отображает текущую страницу"""
        if 0 <= self.current_page - 1 < len(self.pages):
            page_text = self.pages[self.current_page - 1]
            html_text = page_text.replace('\n', '<br>')
            html_text += f'<div class="page-number">Страница {self.current_page} из {len(self.pages)}</div>'
            self.text_edit.setHtml(html_text)

            # Обновляем метку страницы
            self.page_spin.blockSignals(True)
            self.page_spin.setValue(self.current_page)
            self.page_spin.blockSignals(False)
            self.page_label.setText(f"из {len(self.pages)}")

            # Обновляем состояние кнопок навигации
            self.prev_page_btn.setEnabled(self.current_page > 1)
            self.next_page_btn.setEnabled(self.current_page < len(self.pages))

            # Применяем подсветку для текущей страницы
            self.apply_highlights_to_current_page()

            # Выделяем текущую ошибку если она на этой странице
            if self.error_positions and self.current_error_index < len(self.error_positions):
                page, _ = self.error_positions[self.current_error_index]
                if page == self.current_page:
                    self.highlight_current_error(self.current_error_index)

    def apply_highlights_to_current_page(self):
        """Применяет подсветку к текущей странице"""
        # Очищаем предыдущую подсветку
        self.clear_highlights()

        # Подсветка ошибок с двойной подсветкой
        if self.all_matches:
            self.highlight_errors_with_full_line()

        # Подсветка результатов поиска
        if self.search_matches:
            self.highlight_search_results()

    def highlight_errors_with_full_line(self):
        """Подсветить ошибки на текущей странице с подсветкой полной строки"""
        page_idx = self.current_page - 1
        if 0 <= page_idx < len(self.pages):
            page_text = self.pages[page_idx]

            # Сначала подсвечиваем все строки с ошибками желтым цветом
            for page, start, end, term, similarity_score in self.all_matches:
                if page == self.current_page and start >= 0 and end <= len(page_text):
                    # Находим начало и конец строки
                    line_start = page_text.rfind('\n', 0, start) + 1
                    line_end = page_text.find('\n', end)
                    if line_end == -1:
                        line_end = len(page_text)

                    # Подсвечиваем всю строку желтым цветом
                    if line_start < line_end:
                        cursor = self.text_edit.textCursor()
                        format = QTextCharFormat()
                        format.setBackground(QColor(255, 255, 0, 80))  # Желтый
                        format.setFontWeight(QFont.Weight.Normal)

                        cursor.setPosition(line_start)
                        cursor.setPosition(line_end, QTextCursor.MoveMode.KeepAnchor)
                        cursor.mergeCharFormat(format)

            # Затем подсвечиваем найденный текст красным цветом
            for page, start, end, term, similarity_score in self.all_matches:
                if page == self.current_page and start >= 0 and end <= len(page_text):
                    cursor = self.text_edit.textCursor()
                    format = QTextCharFormat()
                    format.setBackground(QColor(255, 0, 0, 150))  # Красный
                    format.setFontWeight(QFont.Weight.Bold)
                    format.setUnderlineStyle(QTextCharFormat.UnderlineStyle.DashUnderline)
                    format.setUnderlineColor(QColor(211, 47, 47))

                    # Добавляем подсказку с процентом схожести
                    if similarity_score:
                        format.setToolTip(f"Схожесть: {similarity_score:.1f}%")

                    cursor.setPosition(start)
                    cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
                    cursor.mergeCharFormat(format)

    def go_to_page(self, page):
        """Перейти на указанную страницу"""
        if 1 <= page <= len(self.pages):
            self.current_page = page
            self.display_current_page()

    def prev_page(self):
        """Перейти к предыдущей странице"""
        if self.current_page > 1:
            self.current_page -= 1
            self.display_current_page()

    def next_page(self):
        """Перейти к следующей странице"""
        if self.current_page < len(self.pages):
            self.current_page += 1
            self.display_current_page()

    def collect_error_matches(self):
        """Собираем все совпадения ошибок для всего документа с процентами схожести"""
        self.all_matches = []
        self.error_positions = []
        self.error_details = {}
        error_counter = 0

        for result_idx, result in enumerate(self.results):
            if result.get('matches'):
                # Проверяем тип проверки и наличие исключающих алиасов
                check_type = result.get('type', '')
                check_name = result.get('name', '')
                check_group = result.get('group', '')

                # Получаем список исключающих алиасов
                without_aliases = result.get('without_aliases', [])

                # Если есть исключающие алиасы, проверяем, присутствуют ли они в документе
                should_exclude = False
                if without_aliases and check_type == 'text_present':
                    # Проверяем, содержатся ли исключающие алиасы в документе
                    for alias in without_aliases:
                        if alias and isinstance(alias, str) and alias.lower() in self.document_text.lower():
                            should_exclude = True
                            logger.info(f"Проверка '{check_name}' исключена, найден алиас: {alias}")
                            break

                # Если проверка должна быть исключена, пропускаем её
                if should_exclude:
                    continue

                # Получаем детальную информацию о совпадениях, если есть
                detailed_matches = result.get('detailed_matches', [])

                # Если есть детальная информация, используем её для получения процентов схожести
                if detailed_matches:
                    for match_idx, match in enumerate(detailed_matches):
                        if isinstance(match, dict):
                            start = match.get('position', 0)
                            end = match.get('end_position', 0)
                            best_score = match.get('best_match_score', 0)
                            match_quality = match.get('match_quality', '')
                            word_scores = match.get('word_matches', [])

                            # Формируем термин с информацией о схожести
                            if best_score:
                                term = f"{check_name} [Схожесть: {best_score:.1f}% ({match_quality})]"
                            else:
                                term = check_name

                            for page_idx, (page_start, page_end) in enumerate(self.page_ranges):
                                if page_start <= start < page_end:
                                    adjusted_start = start - page_start
                                    adjusted_end = min(end - page_start, page_end - page_start)

                                    self.all_matches.append(
                                        (page_idx + 1, adjusted_start, adjusted_end, term, best_score))
                                    self.error_positions.append((page_idx + 1, adjusted_start))

                                    error_id = f"error_{result_idx}_{match_idx}"
                                    context_start = max(0, start - 150)
                                    context_end = min(len(self.document_text), end + 150)
                                    context = self.document_text[context_start:context_end]

                                    if context_start > 0 and not self.document_text[context_start - 1] in [' ', '\n',
                                                                                                           '\t']:
                                        context = "..." + context
                                    if context_end < len(self.document_text) and not self.document_text[
                                                                                         context_end] in [
                                                                                         ' ', '\n', '\t', '.', ',', ';',
                                                                                         ':', '!', '?']:
                                        context = context + "..."

                                    self.error_details[error_id] = {
                                        'id': error_id,
                                        'name': check_name,
                                        'group': check_group,
                                        'message': result.get('message', ''),
                                        'page': page_idx + 1,
                                        'start': adjusted_start,
                                        'end': adjusted_end,
                                        'length': adjusted_end - adjusted_start,
                                        'passed': result.get('passed', False),
                                        'needs_verification': result.get('needs_verification', False),
                                        'term': term,
                                        'global_start': start,
                                        'global_end': end,
                                        'context': context,
                                        'type': check_type,
                                        'without_aliases': without_aliases,
                                        'similarity_score': best_score,
                                        'match_quality': match_quality,
                                        'word_scores': word_scores
                                    }
                                    error_counter += 1
                                    break
                else:
                    # Если нет детальной информации, используем обычные совпадения
                    for match_idx, match in enumerate(result['matches']):
                        if len(match) >= 2:
                            start, end = match[0], match[1]
                            term = match[2] if len(match) > 2 else check_name

                            # Проверяем тип результата - для version_comparison может быть особая обработка
                            if check_type == 'version_comparison' and 'section_results' in result:
                                # Для проверки версий используем информацию из секций
                                section_results = result.get('section_results', [])
                                for section in section_results:
                                    if section.get('position', -1) == start:
                                        found_version = section.get('found_version', 'не найдено')
                                        required = section.get('required_version', '')
                                        operator_display = section.get('operator_display', '≥')
                                        term = f"{section.get('name', 'Показатель')}: {found_version} {operator_display} {required}"
                                        break

                            for page_idx, (page_start, page_end) in enumerate(self.page_ranges):
                                if page_start <= start < page_end:
                                    adjusted_start = start - page_start
                                    adjusted_end = min(end - page_start, page_end - page_start)

                                    self.all_matches.append((page_idx + 1, adjusted_start, adjusted_end, term, None))
                                    self.error_positions.append((page_idx + 1, adjusted_start))

                                    error_id = f"error_{result_idx}_{match_idx}"
                                    context_start = max(0, start - 150)
                                    context_end = min(len(self.document_text), end + 150)
                                    context = self.document_text[context_start:context_end]

                                    if context_start > 0 and not self.document_text[context_start - 1] in [' ', '\n',
                                                                                                           '\t']:
                                        context = "..." + context
                                    if context_end < len(self.document_text) and not self.document_text[
                                                                                         context_end] in [
                                                                                         ' ', '\n', '\t', '.', ',', ';',
                                                                                         ':', '!', '?']:
                                        context = context + "..."

                                    self.error_details[error_id] = {
                                        'id': error_id,
                                        'name': check_name,
                                        'group': check_group,
                                        'message': result.get('message', ''),
                                        'page': page_idx + 1,
                                        'start': adjusted_start,
                                        'end': adjusted_end,
                                        'length': adjusted_end - adjusted_start,
                                        'passed': result.get('passed', False),
                                        'needs_verification': result.get('needs_verification', False),
                                        'term': term,
                                        'global_start': start,
                                        'global_end': end,
                                        'context': context,
                                        'type': check_type,
                                        'without_aliases': without_aliases,
                                        'similarity_score': None
                                    }
                                    error_counter += 1
                                    break

        logger.info(f"Собрано {len(self.all_matches)} совпадений ошибок с процентами схожести")

    def search_text(self):
        """Поиск текста в документе по целым словам"""
        search_term = self.search_input.text().strip()
        if not search_term:
            self.search_matches = []
            self.search_count_label.setText("Найдено: 0")
            self.search_prev_btn.setEnabled(False)
            self.search_next_btn.setEnabled(False)
            self.clear_highlights()
            self.highlight_errors_with_full_line()
            return

        self.search_matches = []
        pattern = re.compile(r'\b' + re.escape(search_term) + r'\b', re.IGNORECASE)

        for page_idx, page_text in enumerate(self.pages):
            for match in pattern.finditer(page_text):
                start, end = match.start(), match.end()
                self.search_matches.append({
                    'page': page_idx + 1,
                    'start': start,
                    'end': end,
                    'text': page_text[start:end]
                })

        self.search_count_label.setText(f"Найдено: {len(self.search_matches)}")

        if self.search_matches:
            self.search_prev_btn.setEnabled(True)
            self.search_next_btn.setEnabled(True)
            self.current_search_index = 0
            self.go_to_search_result(0)
        else:
            self.search_prev_btn.setEnabled(False)
            self.search_next_btn.setEnabled(False)
            QMessageBox.information(self, "Поиск", f"Фраза '{search_term}' не найдена")

    def highlight_search_results(self):
        """Подсветить результаты поиска на текущей странице"""
        cursor = self.text_edit.textCursor()
        format = QTextCharFormat()
        format.setBackground(QColor(0, 100, 255, 100))
        format.setFontWeight(QFont.Weight.Bold)

        for match in self.search_matches:
            if match['page'] == self.current_page:
                cursor.setPosition(match['start'])
                cursor.setPosition(match['end'], QTextCursor.MoveMode.KeepAnchor)
                cursor.mergeCharFormat(format)

    def prev_search_result(self):
        """Перейти к предыдущему результату поиска"""
        if self.search_matches:
            self.current_search_index = (self.current_search_index - 1) % len(self.search_matches)
            self.go_to_search_result(self.current_search_index)

    def next_search_result(self):
        """Перейти к следующему результату поиска"""
        if self.search_matches:
            self.current_search_index = (self.current_search_index + 1) % len(self.search_matches)
            self.go_to_search_result(self.current_search_index)

    def go_to_search_result(self, index):
        """Перейти к указанному результату поиска"""
        if 0 <= index < len(self.search_matches):
            match = self.search_matches[index]
            if match['page'] != self.current_page:
                self.current_page = match['page']
                self.display_current_page()

            cursor = self.text_edit.textCursor()
            cursor.setPosition(match['start'])
            self.text_edit.setTextCursor(cursor)
            self.text_edit.ensureCursorVisible()
            self.search_count_label.setText(f"Найдено: {len(self.search_matches)} (текущий: {index + 1})")

    def show_all_errors(self):
        """Подсветить все найденные ошибки"""
        if not self.all_matches:
            QMessageBox.information(self, "Ошибки", "Ошибок не найдено")
            return

        if self.error_positions:
            self.current_error_index = 0
            self.go_to_error(0)

        self.apply_highlights_to_current_page()
        self.update_error_navigation()

    def clear_highlights(self):
        """Очистить подсветку"""
        cursor = self.text_edit.textCursor()
        format = QTextCharFormat()
        format.setBackground(QColor(0, 0, 0, 0))
        format.setFontWeight(QFont.Weight.Normal)
        format.setUnderlineStyle(QTextCharFormat.UnderlineStyle.NoUnderline)
        format.setToolTip("")  # Очищаем подсказки

        cursor.select(QTextCursor.SelectionType.Document)
        cursor.mergeCharFormat(format)

        cursor.setPosition(0)
        self.text_edit.setTextCursor(cursor)

    def update_stats(self):
        """Обновить статистику"""
        stats = []

        if self.all_matches:
            errors_by_page = {}
            for match in self.all_matches:
                page = match[0]
                errors_by_page[page] = errors_by_page.get(page, 0) + 1

            errors_per_page = ', '.join([f'стр. {page}: {count}' for page, count in sorted(errors_by_page.items())])

            # Подсчитываем средний процент схожести
            scores = [match[4] for match in self.all_matches if match[4] is not None]
            if scores:
                avg_score = sum(scores) / len(scores)
                stats.append(f"Ошибок: {len(self.all_matches)} (сред. схожесть: {avg_score:.1f}%)")
                stats.append(f"Распределение: {errors_per_page}")
            else:
                stats.append(f"Ошибок: {len(self.all_matches)} ({errors_per_page})")

        total_chars = len(self.document_text)
        words = len(self.document_text.split())
        stats.append(f"Страниц: {len(self.pages)}")
        stats.append(f"Символов: {total_chars:,}".replace(',', ' '))
        stats.append(f"Слов: {words:,}".replace(',', ' '))

        self.stats_label.setText(" | ".join(stats))

    def show_error_tooltip(self):
        """Показать всплывающую подсказку об ошибке"""
        if not self.hover_pos:
            return

        cursor = self.text_edit.cursorForPosition(self.hover_pos)
        position = cursor.position()

        current_error = None
        for error_id, error_info in self.error_details.items():
            if error_info['page'] == self.current_page:
                start = error_info['start']
                end = error_info['end']
                if start <= position <= end:
                    current_error = error_info
                    break

        if current_error:
            tooltip_lines = [
                f"<b>{current_error['name']}</b>",
                f"{current_error['message']}",
                f"Позиция: {position - start + 1}-{position - start + current_error['length']}"
            ]

            # Добавляем информацию о проценте схожести
            if current_error.get('similarity_score') is not None:
                score = current_error['similarity_score']
                quality = current_error.get('match_quality', '')
                tooltip_lines.insert(1, f"<b>Схожесть:</b> {score:.1f}% ({quality})")

                # Добавляем информацию по словам
                if current_error.get('word_scores'):
                    tooltip_lines.append("<b>Совпадение по словам:</b>")
                    for word_info in current_error['word_scores'][:3]:
                        tooltip_lines.append(f"  • {word_info['word']}: {word_info['score']:.1f}%")

            tooltip_text = "<br>".join(tooltip_lines)

            QToolTip.showText(
                self.text_edit.mapToGlobal(self.hover_pos),
                tooltip_text,
                self.text_edit,
                self.text_edit.rect(),
                3000
            )