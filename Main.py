# main.py
import sys
import traceback
from PyQt6.QtWidgets import QApplication, QMessageBox
from app_window import MainWindow

def exception_hook(exctype, value, tb):
    """Обработчик необработанных исключений"""
    error_msg = ''.join(traceback.format_exception(exctype, value, tb))
    print(f"Необработанное исключение:\n{error_msg}")

    # Показать сообщение об ошибке
    msg_box = QMessageBox()
    msg_box.setIcon(QMessageBox.Icon.Critical)
    msg_box.setWindowTitle("Критическая ошибка")
    msg_box.setText(f"Произошла ошибка:\n{str(value)}")
    msg_box.setDetailedText(error_msg)
    msg_box.exec()

    # Закрыть приложение
    sys.exit(1)


def main():
    # Устанавливаем обработчик исключений
    sys.excepthook = exception_hook

    app = QApplication(sys.argv)
    app.setApplicationName("Проверка документов ФК")
    app.setStyle("Fusion")

    try:
        window = MainWindow()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        error_msg = traceback.format_exc()
        print(f"Ошибка при запуске приложения:\n{error_msg}")
        QMessageBox.critical(None, "Ошибка запуска",
                             f"Не удалось запустить приложение:\n{str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()