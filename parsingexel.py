import sys
import os
import json
import re
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QMessageBox,
    QProgressBar, QStatusBar, QTextEdit
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont


class ExcelProcessor(QThread):
    """Поток для обработки Excel с сохранением по частям"""
    progress = pyqtSignal(int, str)
    error = pyqtSignal(str)
    finished = pyqtSignal(str)

    def __init__(self, file_path, output_path):
        super().__init__()
        self.file_path = file_path
        self.output_path = output_path

    def find_columns(self, columns):
        """Поиск нужных колонок"""
        name_cols = []
        version_cols = []
        contract_cols = []

        # Ключевые слова для поиска
        name_keywords = ['наименование', 'название', 'по', 'name', 'product', 'software', 'программа', 'п/о']
        version_keywords = ['версия', 'version', 'вер', 'v.', 'var', 'редакция', 'ver']
        contract_keywords = ['гк', 'контракт', 'contract', 'договор', '№', 'номер', 'gk', 'код', 'ид']

        for col in columns:
            col_lower = str(col).lower()

            # Проверяем каждую категорию
            if any(kw in col_lower for kw in name_keywords):
                name_cols.append(col)
            elif any(kw in col_lower for kw in version_keywords):
                version_cols.append(col)
            elif any(kw in col_lower for kw in contract_keywords):
                contract_cols.append(col)

        return name_cols, version_cols, contract_cols

    def extract_gk_pattern(self, text):
        """Извлекает только ГК вида ФКУ000/2025"""
        if pd.isna(text) or not text:
            return []

        text = str(text).strip()

        # Паттерн для поиска ГК вида ФКУ000/2025
        # ФКУ - буквы, 000 - цифры, /2025 - необязательная часть с годом
        pattern = r'[А-ЯA-Z]{3,4}\d{3,4}(?:[/-]\d{2,4})?'

        # Ищем все совпадения
        matches = re.findall(pattern, text, re.IGNORECASE)

        # Очищаем и фильтруем
        result = []
        for match in matches:
            clean_match = match.strip()
            # Проверяем что это действительно наш формат
            if re.match(r'^[А-ЯA-Z]{3,4}\d{3,4}(?:[/-]\d{2,4})?$', clean_match, re.IGNORECASE):
                result.append(clean_match.upper())  # Приводим к верхнему регистру

        return result

    def run(self):
        try:
            self.progress.emit(5, "Чтение файла...")

            # Открываем файл для записи JSON
            with open(self.output_path, 'w', encoding='utf-8') as f:
                f.write('[\n')
                first_record = True
                total_records = 0

                # Получаем все листы
                excel_file = pd.ExcelFile(self.file_path)
                sheets = excel_file.sheet_names
                total_sheets = len(sheets)

                self.progress.emit(10, f"Найдено листов: {total_sheets}")

                # Обрабатываем каждый лист
                for sheet_idx, sheet_name in enumerate(sheets):
                    progress_base = 10 + int(80 * sheet_idx / total_sheets)

                    self.progress.emit(
                        progress_base,
                        f"Обработка листа {sheet_idx + 1}/{total_sheets}: {sheet_name}"
                    )

                    # Читаем лист с заголовками на второй строке
                    df = pd.read_excel(
                        self.file_path,
                        sheet_name=sheet_name,
                        header=1,
                        dtype=str
                    )

                    if df.empty:
                        continue

                    # Находим колонки
                    name_cols, version_cols, contract_cols = self.find_columns(df.columns)

                    # Обрабатываем каждую строку
                    for row_idx, row in df.iterrows():
                        # Базовые поля
                        record = {
                            'лист': sheet_name,
                            'строка': row_idx + 3
                        }

                        # Собираем наименования
                        names = []
                        for col in name_cols:
                            val = str(row[col]).strip() if pd.notna(row[col]) else ""
                            if val and val.lower() not in ['nan', 'none', '']:
                                names.append(val)

                        if names:
                            record['наименование'] = ' | '.join(names)

                        # Собираем версии
                        versions = []
                        for col in version_cols:
                            val = str(row[col]).strip() if pd.notna(row[col]) else ""
                            if val and val.lower() not in ['nan', 'none', '']:
                                versions.append(val)

                        if versions:
                            record['версия'] = ' | '.join(versions)

                        # Собираем ГК (только нужного формата)
                        all_gk = []
                        for col in contract_cols:
                            val = str(row[col]).strip() if pd.notna(row[col]) else ""
                            if val and val.lower() not in ['nan', 'none', '']:
                                # Извлекаем только ГК вида ФКУ000/2025
                                gk_list = self.extract_gk_pattern(val)
                                all_gk.extend(gk_list)

                        if all_gk:
                            # Убираем дубликаты
                            unique_gk = []
                            for gk in all_gk:
                                if gk not in unique_gk:
                                    unique_gk.append(gk)

                            if len(unique_gk) > 1:
                                # Если несколько ГК, создаем отдельные записи
                                for gk in unique_gk:
                                    gk_record = record.copy()
                                    gk_record['гк'] = gk

                                    if not first_record:
                                        f.write(',\n')
                                    json.dump(gk_record, f, ensure_ascii=False, default=str)
                                    first_record = False
                                    total_records += 1
                            else:
                                # Один ГК
                                record['гк'] = unique_gk[0]

                                if not first_record:
                                    f.write(',\n')
                                json.dump(record, f, ensure_ascii=False, default=str)
                                first_record = False
                                total_records += 1
                        else:
                            # Запись без ГК (пропускаем, если нужны только с ГК)
                            # Если хотите сохранять и записи без ГК, раскомментируйте:
                            """
                            if not first_record:
                                f.write(',\n')
                            json.dump(record, f, ensure_ascii=False, default=str)
                            first_record = False
                            total_records += 1
                            """
                            pass

                        # Обновляем прогресс каждые 100 записей
                        if total_records % 100 == 0:
                            self.progress.emit(
                                progress_base,
                                f"Обработано записей с ГК: {total_records}"
                            )

                # Закрываем JSON массив
                f.write('\n]')

            self.progress.emit(100, f"Готово! Создано записей с ГК: {total_records}")
            self.finished.emit(self.output_path)

        except Exception as e:
            self.error.emit(str(e))


class ExcelToJsonConverter(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel to JSON Converter - Фильтр ГК вида ФКУ000/2025")
        self.setGeometry(300, 300, 700, 500)

        self.file_path = ""
        self.init_ui()

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)

        # Заголовок
        title = QLabel("Excel to JSON Converter")
        title.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title.setStyleSheet("color: #2196F3;")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # Подзаголовок
        subtitle = QLabel("Фильтр: только ГК вида ФКУ000/2025")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setWordWrap(True)
        subtitle.setStyleSheet("color: #4CAF50; font-weight: bold;")
        layout.addWidget(subtitle)

        # Кнопка выбора файла
        self.btn_select = QPushButton("📂 Выбрать Excel файл")
        self.btn_select.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                padding: 8px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #45a049; }
        """)
        self.btn_select.clicked.connect(self.select_file)
        layout.addWidget(self.btn_select)

        # Информация о файле
        self.lbl_file = QLabel("Файл не выбран")
        self.lbl_file.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_file.setStyleSheet("color: gray;")
        layout.addWidget(self.lbl_file)

        # Прогресс
        self.progress = QProgressBar()
        self.progress.setVisible(False)
        layout.addWidget(self.progress)

        # Статус обработки
        self.lbl_status = QLabel("")
        self.lbl_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.lbl_status)

        # Лог
        log_label = QLabel("Лог обработки:")
        layout.addWidget(log_label)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setMaximumHeight(200)
        self.log.setFont(QFont("Courier New", 9))
        layout.addWidget(self.log)

        # Кнопка обработки
        self.btn_process = QPushButton("🔄 Обработать и сохранить JSON")
        self.btn_process.setEnabled(False)
        self.btn_process.setStyleSheet("""
            QPushButton {
                background-color: #FF9800;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover { background-color: #F57C00; }
            QPushButton:disabled { background-color: #ccc; }
        """)
        self.btn_process.clicked.connect(self.process_file)
        layout.addWidget(self.btn_process)

        # Статус
        self.status = QStatusBar()
        self.setStatusBar(self.status)
        self.status.showMessage("✅ Готов к работе")

    def log_message(self, msg):
        self.log.append(msg)
        cursor = self.log.textCursor()
        cursor.movePosition(cursor.MoveOperation.End)
        self.log.setTextCursor(cursor)

    def select_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл", "", "Excel files (*.xlsx *.xls)"
        )
        if file_path:
            self.file_path = file_path
            self.lbl_file.setText(f"📄 {os.path.basename(file_path)}")
            self.lbl_file.setStyleSheet("color: black;")
            self.btn_process.setEnabled(True)
            self.log_message(f"✅ Выбран файл: {os.path.basename(file_path)}")

    def process_file(self):
        # Выбираем место для сохранения
        base = os.path.splitext(os.path.basename(self.file_path))[0]
        default_name = f"{base}_filtered.json"

        output_path, _ = QFileDialog.getSaveFileName(
            self, "Сохранить JSON", default_name, "JSON files (*.json)"
        )

        if not output_path:
            return

        # Очищаем лог
        self.log.clear()
        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.lbl_status.setText("Обработка...")
        self.btn_process.setEnabled(False)
        self.btn_select.setEnabled(False)

        self.log_message("=" * 50)
        self.log_message("🚀 НАЧАЛО ОБРАБОТКИ")
        self.log_message(f"📁 Входной файл: {os.path.basename(self.file_path)}")
        self.log_message(f"📁 Выходной файл: {os.path.basename(output_path)}")
        self.log_message("🎯 Фильтр: только ГК вида ФКУ000/2025")

        # Запускаем обработку
        self.processor = ExcelProcessor(self.file_path, output_path)
        self.processor.progress.connect(self.on_progress)
        self.processor.error.connect(self.on_error)
        self.processor.finished.connect(self.on_finished)
        self.processor.start()

    def on_progress(self, value, message):
        self.progress.setValue(value)
        self.lbl_status.setText(message)
        self.status.showMessage(message)

        if "записей" in message.lower():
            self.log_message(f"📊 {message}")

    def on_error(self, error_msg):
        self.log_message(f"❌ ОШИБКА: {error_msg}")
        self.progress.setVisible(False)
        self.lbl_status.setText("❌ Ошибка")
        self.btn_process.setEnabled(True)
        self.btn_select.setEnabled(True)
        QMessageBox.critical(self, "Ошибка", error_msg)

    def on_finished(self, output_path):
        self.progress.setVisible(False)
        self.lbl_status.setText("✅ Готово!")
        self.btn_process.setEnabled(True)
        self.btn_select.setEnabled(True)

        size = os.path.getsize(output_path) / 1024

        self.log_message("=" * 50)
        self.log_message(f"✅ ОБРАБОТКА ЗАВЕРШЕНА")
        self.log_message(f"📁 Файл сохранен: {os.path.basename(output_path)}")
        self.log_message(f"💾 Размер: {size:.1f} KB")

        # Показываем пример
        try:
            with open(output_path, 'r', encoding='utf-8') as f:
                content = f.read()
                if len(content) > 500:
                    content = content[:500] + "..."
                self.log_message("\n📋 Пример JSON:")
                self.log_message(content)
        except:
            pass

        QMessageBox.information(
            self, "Успех",
            f"✅ JSON файл создан!\n\n"
            f"📁 {os.path.basename(output_path)}\n"
            f"💾 Размер: {size:.1f} KB\n\n"
            f"В файле только записи с ГК вида ФКУ000/2025"
        )


def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')

    window = ExcelToJsonConverter()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()