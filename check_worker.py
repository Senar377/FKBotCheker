# check_worker.py
from PyQt6.QtCore import QThread, pyqtSignal
import logging

logger = logging.getLogger(__name__)


class CheckWorker(QThread):
    """Поток для выполнения проверки"""
    progress = pyqtSignal(int, str)  # прогресс и текущая проверка
    finished = pyqtSignal(list)

    def __init__(self, checker, document_text: str, selected_checks: list, config: dict, page_info: list):
        super().__init__()
        self.checker = checker
        self.document_text = document_text
        self.selected_checks = selected_checks
        self.config = config
        self.page_info = page_info
        self._normalized_cache = {}

    def run(self):
        from document_checker import DocumentChecker

        results = []

        logger.info(f"Начало проверки документа. Длина текста: {len(self.document_text)} символов")
        logger.info(f"Количество страниц: {len(self.page_info)}")
        logger.info(f"Выбранные проверки: {len(self.selected_checks)}")

        # Создаем временный чекер с текущей конфигурацией
        temp_checker = DocumentChecker()
        temp_checker.config = self.config

        # Находим все подпроверки
        all_subchecks = []
        for check_group in self.config.get('checks', []):
            for subcheck in check_group.get('subchecks', []):
                check_name = subcheck.get('name', '')
                if not self.selected_checks or check_name in self.selected_checks:
                    all_subchecks.append((check_group.get('group', ''), subcheck))

        total_checks = len(all_subchecks)
        logger.info(f"Всего проверок для выполнения: {total_checks}")

        # Предварительная обработка для оптимизации
        check_types = {}
        for group_name, subcheck in all_subchecks:
            check_type = subcheck.get('type', '')
            if check_type not in check_types:
                check_types[check_type] = []
            check_types[check_type].append((group_name, subcheck))

        # Выполняем проверки группами по типу для оптимизации
        for i, (group_name, subcheck) in enumerate(all_subchecks):
            check_name = subcheck.get('name', 'Неизвестная проверка')
            check_type = subcheck.get('type', '')

            # Отправляем информацию о текущей проверке
            self.progress.emit(int((i + 1) / total_checks * 100), f"Проверка: {check_name}")

            # Выполняем проверку
            result = temp_checker.check_subcheck(subcheck, self.document_text, self.page_info)
            result['group'] = group_name
            results.append(result)

            logger.debug(
                f"Выполнена проверка: {check_name}, результат: {result.get('message', '')}, страница: {result.get('page', 0)}")

            # Адаптивная задержка для UI
            if i % 10 == 0:  # Обновляем прогресс каждые 10 проверок
                self.msleep(10)

        logger.info(f"Проверка завершена. Найдено результатов: {len(results)}")
        self.finished.emit(results)