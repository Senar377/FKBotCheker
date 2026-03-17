# app_window.py
"""
Главное окно приложения для проверки технической документации
С поддержкой истории версий документов, автосохранением и комментариями
"""

import re
import time
import yaml
import logging
import os
from datetime import datetime
from typing import Optional, Dict, Any, Tuple, List
from pathlib import Path

from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QListWidget, QListWidgetItem,
    QProgressBar, QTableWidget, QTableWidgetItem,
    QGroupBox, QMessageBox, QSplitter, QLineEdit,
    QComboBox, QMenu, QFileDialog, QHeaderView, QStatusBar, QSizePolicy,
    QFrame, QGridLayout, QTabWidget, QTextEdit, QScrollArea,
    QApplication, QInputDialog
)
from PyQt6.QtCore import Qt, QSettings, QTimer, pyqtSignal
from PyQt6.QtGui import QFont, QAction, QColor, QTextCursor

from document_checker import DocumentChecker
from docx_parser import DOCXParser
from add_check_dialog import AddCheckDialog
from settings_dialog import SettingsDialog
from config_editor import ConfigEditor
from document_viewer import DocumentViewer
from check_worker import CheckWorker
from manage_checks_dialog import ManageChecksDialog
from versions_dialog import VersionsDialog
from json_database import JSONDatabase
from document_history import DocumentHistory
from document_history_dialog import DocumentHistoryDialog

# Настройка логирования
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


class GKManager:
    """Класс для управления ГК"""

    def __init__(self):
        self.settings = QSettings("ФедеральноеКазначейство", "Settings")
        self.document_gk = []
        self.first_pages_gk = []
        self.gk_to_subsystem = {}
        self.gk_date = None

    def extract_gk_from_text(self, text, page_info=None, max_pages=3):
        """Извлечение ГК из текста"""
        if not text:
            return []

        gk_pattern = self.settings.value(
            "gk_format",
            r"ФКУ\d{3,4}(?:[/-]\d{2,4})?(?:/\w+)?"
        )

        matches = re.findall(gk_pattern, text, re.IGNORECASE)
        all_gk = [m.upper() for m in matches]
        self.document_gk = sorted(list(set(all_gk)))

        self._extract_gk_date(text)

        self.map_gk_to_subsystems()

        if page_info:
            pages_text = []
            for i, (start, end) in enumerate(page_info):
                if i < max_pages:
                    page_text = text[start:end]
                    pages_text.append(page_text)

            first_pages_matches = []
            for page_text in pages_text:
                matches = re.findall(gk_pattern, page_text, re.IGNORECASE)
                first_pages_matches.extend(matches)

            self.first_pages_gk = sorted(list(set([m.upper() for m in first_pages_matches])))
        else:
            self.first_pages_gk = self.document_gk[:]

        return self.document_gk

    def _extract_gk_date(self, text):
        """Извлечение даты создания ГК"""
        gk_pattern = self.settings.value(
            "gk_format",
            r"ФКУ\d{3,4}(?:[/-]\d{2,4})?(?:/\w+)?"
        )

        gk_match = re.search(gk_pattern, text, re.IGNORECASE)
        if gk_match:
            start_pos = max(0, gk_match.start() - 200)
            text_before_gk = text[start_pos:gk_match.start()]

            date_patterns = [
                r'от\s+(\d{2}\.\d{2}\.\d{4})',
                r'(\d{2}\.\d{2}\.\d{4})\s*г',
                r'(\d{2}\.\d{2}\.\d{4})',
                r'(\d{4}-\d{2}-\d{2})',
            ]

            for pattern in date_patterns:
                date_match = re.search(pattern, text_before_gk, re.IGNORECASE)
                if date_match:
                    self.gk_date = date_match.group(1) if date_match.groups() else date_match.group(0)
                    try:
                        for fmt in ['%d.%m.%Y', '%Y-%m-%d']:
                            try:
                                date_obj = datetime.strptime(self.gk_date, fmt)
                                self.gk_date = date_obj.strftime('%d.%m.%Y')
                                break
                            except ValueError:
                                continue
                    except:
                        pass
                    break

    def map_gk_to_subsystems(self):
        """Сопоставление ГК с подсистемами"""
        gk_mapping = {
            'ФКУ0375/2025/РИС': 'ГМП',
            'ФКУ0241/2025/РИС': 'ГАСУ',
            'ФКУ0240/2025/РИС': 'ГАСУ',
            'ФКУ0215/2025/РИС': 'ПОИ',
            'ФКУ0237/2025/РИС': 'ПУДС',
            'ФКУ0233/2025/РИС': 'ПУДС',
            'ФКУ0246/2025/РИС': 'ПУДС',
            'ФКУ0257/2025/РИС': 'ПУиО',
            'ФКУ0231/2025/РИС': 'ПУиО',
            'ФКУ0247/2025/РИС': 'ПУиО',
            'ФКУ0261/2025/РИС': 'ПУиО',
            'ФКУ0173/2025/РИС': 'ПУиО',
            'ФКУ0358/2025/РИС': 'ПУиО',
            'ФКУ0346/2025/РИС': 'ПИАО',
            'ФКУ0404/2025/РИС': 'ПИАО',
            'ФКУ0336/2025/РИС': 'ЕПБС',
            'ФКУ0289/2025/РИС': 'ЕПБС',
            'ФКУ0232/2025/РИС': 'ПУР',
        }

        self.gk_to_subsystem = {}
        for gk in self.document_gk:
            if gk in gk_mapping:
                self.gk_to_subsystem[gk] = gk_mapping[gk]
            else:
                for key, subsystem in gk_mapping.items():
                    if key in gk or gk in key:
                        self.gk_to_subsystem[gk] = subsystem
                        break
                else:
                    self.gk_to_subsystem[gk] = 'Не определена'

    def get_subsystem_for_gk(self, gk):
        """Получение подсистемы для ГК"""
        return self.gk_to_subsystem.get(gk, 'Не определена')

    def get_gk_priority(self, gk_list):
        """Получение приоритета для списка ГК"""
        if not gk_list:
            return 0

        priority = 0
        priority_from_doc = self.settings.value("priority_from_doc", True, type=bool)
        priority_from_first_pages = self.settings.value("priority_from_first_pages", True, type=bool)

        for gk in gk_list:
            if priority_from_doc and gk in self.document_gk:
                priority += 10
                if priority_from_first_pages and gk in self.first_pages_gk:
                    priority += 20
            elif gk:
                priority += 1

        return priority

    def has_matching_gk(self, gk_list):
        """Проверка наличия совпадающих ГК с документом"""
        if not gk_list or not self.document_gk:
            return False

        for gk in gk_list:
            if gk in self.document_gk:
                return True

        return False


class DocumentInfoExtractor:
    """Класс для извлечения информации из документа"""

    def __init__(self):
        self.gk_manager = GKManager()

    def extract_gk_numbers(self, text: str, page_info=None) -> List[str]:
        """Извлечение номеров ГК из текста используя GKManager"""
        if not text:
            return []

        self.gk_manager.extract_gk_from_text(text, page_info)
        return self.gk_manager.document_gk

    def get_gk_with_subsystems(self) -> Dict[str, str]:
        """Получение словаря ГК с подсистемами"""
        return self.gk_manager.gk_to_subsystem

    def get_gk_date(self) -> Optional[str]:
        """Получение даты создания ГК"""
        return self.gk_manager.gk_date

    def get_first_page_gk(self) -> Optional[str]:
        """Получение первого ГК с первой страницы"""
        if self.gk_manager.first_pages_gk:
            return self.gk_manager.first_pages_gk[0]
        return None


class FilenameParser:
    """
    Парсер и валидатор имен файлов документации согласно
    "Приложение 2 Кодирование документации v 6 1.docx"
    """

    IS_CODES = {
        '01': 'AUDPLAN (Внутренний контроль и аудит)',
        '02': 'SYD_WORK (Аналитический учет и ведение судебной работы)',
        '03': 'ЛЕКС (Управление ликвидностью)',
        '04': 'КОБРФ (Кассовое обслуживание бюджетов)',
        '08': 'ЦКС, Центр-КС (Казначейское исполнение)',
        '09': 'АС ФК (Автоматизированная система Федерального казначейства)',
        '10': 'АКСИОК.Net',
        '13': 'СПТО (Система поддержки технологического обеспечения)',
        '14': 'САВД (Система сбора, анализа и визуализации данных, КПЭ)',
        '15': 'СУЭ (Система управления эксплуатацией)',
        '18': 'ГИС ГМУ',
        '20': 'ГИИС ЭБ, ЭБ (Электронный бюджет)',
        '22': 'АСД LanDocs',
        '23': 'ГАСУ (Государственная автоматизированная система "Управление")',
        '25': 'Ведомственный портал',
        '26': 'ГИС ГМП',
        '30': 'ЕИС (Единая информационная система в сфере закупок)',
        '31': 'СОБИ (Система обеспечения безопасности информации)',
        '32': 'Официальный сайт Казначейства России',
        '33': 'СКИАО',
        '34': 'УЦ ФК (Удостоверяющий центр)',
        '35': 'ЕОИ (Единая облачная инфраструктура)',
        '36': 'АС Планирование',
        '37': 'АСУПиМ',
        '39': 'АИС УБР Роструда',
        '40': 'ВИС БУ',
        '41': 'ГИС "Независимый регистратор"',
        '42': 'ГИС ЭС',
        '43': 'Лицензионное ПО',
        '44': 'ГИС Торги',
    }

    DOCUMENT_CODES = {
        'ТЗ': 'Техническое задание',
        'ТП': 'Ведомость технического проекта',
        'П2': 'Пояснительная записка',
        'П3': 'Описание автоматизируемых функций',
        'П4': 'Описание постановки задачи',
        'П5': 'Описание информационного обеспечения',
        'П6': 'Описание организации информационной базы',
        'СА': 'Системная архитектура',
        'П9': 'Описание комплекса технических средств',
        'ПА': 'Описание программного обеспечения',
        'В4': 'Спецификация оборудования',
        'А01': 'Акт классификации ГИС',
        'А02': 'Акт оценки уровня защищенности',
        'МУ': 'Модель угроз',
        'МН': 'Модель нарушителя',
        'ТЗ1': 'ТЗ на систему защиты информации',
        'ПЗ': 'Пояснительная записка к ТП',
        'ПЛ': 'План мероприятий по защите информации',
        'ОА': 'Общая архитектура',
        'ОТ': 'Общие требования',
        'ТФ': 'Требования к форматам файлов',
        'ТТ': 'Технические требования',
        'П21': 'Технический проект',
        'П22': 'Технический проект на инфраструктуру',
        'ТВ': 'Требования к информационному взаимодействию',
        'ЭД': 'Ведомость эксплуатационной документации',
        'ПД': 'Общее описание системы',
        'ПС1': 'Паспорт',
        'ПС2': 'Паспорт ИТ-сервиса',
        'ЭП': 'Эксплуатационные показатели',
        'ИА': 'Руководство по администрированию',
        'ИМ': 'Руководство по пуско-наладке',
        'ИО': 'Инструкция по обновлению',
        'ИЭ': 'Инструкция по эксплуатации',
        'КС': 'Каталог ИТ-сервиса',
        'ИЗ': 'Руководство пользователя',
        'ТР': 'Технологический регламент',
        'ТК': 'Технологическая карта',
        'ПМ1': 'Программа предварительных испытаний',
        'ПМ2': 'Программа опытной эксплуатации',
        'ПМ3': 'Программа приемочных испытаний',
        'ПМ4': 'Программа приемо-сдаточных испытаний',
        'С6': 'Таблица соединений',
        'С7': 'План расположения оборудования',
        'ПТ': 'Порядок эксплуатации',
        'ПС3': 'ИТ-паспорт',
        'БЗ': 'База знаний',
        'ОЯ': 'Описание языка',
        'ОП': 'Описание программы',
        'РСА': 'Руководство системного администратора',
        'РПР': 'Руководство программиста',
        'РО': 'Руководство оператора',
        'ОД': 'Ведомость отчетных документов',
        'ОМ': 'Обучающие материалы',
        'ДР': 'Другое',
        'ОТЧ': 'Плановый отчет',
        'ОТ1': 'Отчет о проверочном восстановлении',
        'ОТ2': 'Отчет о результатах анализа',
        'ОТ3': 'Отчет о результатах анализа уязвимостей',
        'ОТ4': 'Отчет об обследовании',
        'ОТ5': 'Отчет об обеспечении защиты',
        'ОТ6': 'Отчет об оказании Услуг',
        'КТС': 'Схемы КТС',
        'ЧРТ': 'Чертежи',
        'А04': 'Акт о завершении пуско-наладочных работ',
        'ПР1': 'Протокол предварительных испытаний',
        'ПР2': 'Протокол комплексных испытаний',
        'ПР3': 'Протокол испытаний функциональности',
        'А05': 'Акт о приемке в опытную эксплуатацию',
        'ПР4': 'Протокол опытной эксплуатации',
        'А06': 'Акт о завершении опытной эксплуатации',
        'ПР5': 'Протокол приемочных испытаний',
        'А07': 'Акт о приемке в эксплуатацию',
        'А08': 'Акт приемки ИС',
        'АЗ': 'Аналитическая записка',
        'РТО': 'Руководство по техническому обслуживанию',
        'ФРМ': 'Формуляр',
        'ЭБ01': 'Системная архитектура ЭБ',
        'ЭБ02': 'ТЗ на систему ЭБ',
        'ЭБ03': 'ЧТЗ на подсистему',
        'ЭБ04': 'Модель угроз ЭБ',
        'ЭБ05': 'Модель нарушителя ЭБ',
        'ЭБ06': 'Акт классификации АС ЭБ',
        'ЭБ07': 'Акт классификации ИСПДн',
        'ЭБ08': 'Технические требования к инфраструктуре',
        'ЭБ09': 'Общие требования',
        'ЭБ10': 'Технический проект на инфраструктуру',
        'ЭБ11': 'Технический проект на подсистему',
        'ЭБ12': 'Требования к взаимодействию',
        'ЭБ13': 'Описание автоматизируемых функций',
        'ЭБ14': 'Пояснительная записка',
        'ЭБ15': 'Описание постановки задачи',
        'ЭБ16': 'Описание информационного обеспечения',
        'ЭБ17': 'Описание организации ИБ',
        'ЭБ18': 'Описание КТС',
        'ЭБ19': 'Описание языка',
        'ЭБ20': 'Описание программы',
        'ЭБ21': 'Технологическая инструкция',
        'ЭБ22': 'Ведомость эксплуатационных документов',
        'ЭБ23': 'Паспорт',
        'ЭБ24': 'Руководство по администрированию',
        'ЭБ25': 'Руководство по техническому обслуживанию',
        'ЭБ26': 'Руководство работников',
        'ЭБ27': 'Программа предварительных испытаний',
        'ЭБ28': 'Протокол предварительных испытаний',
        'ЭБ29': 'Акт приемки в опытную эксплуатацию',
        'ЭБ30': 'Программа опытной эксплуатации',
        'ЭБ31': 'Протокол опытной эксплуатации',
        'ЭБ32': 'Акт завершения опытной эксплуатации',
        'ЭБ33': 'Программа приемочных испытаний',
        'ЭБ34': 'Протокол приемочных испытаний',
        'ЭБ35': 'Акт приемки в эксплуатацию',
        'ЭБ99': 'Документы ЭБ',
    }

    SCOPE_CODES = {
        '1': 'ЦАФК',
        '2': 'ТОФК',
        '4': 'Исполнители',
        '5': 'МОУ',
        '6': 'Внешние',
        '7': 'Минфин',
        '8': 'ФКУ ЦОКР',
        '9': 'Все',
        '10': 'Межрегиональные',
    }

    SUBSYSTEM_CODES = {
        '20_04,00': 'ПУДС',
        '20_04,01': 'ПУДС МУЛ',
        '20_04,02': 'ПУДС КП',
        '20_05,00': 'ПУР',
        '20_09,00': 'ПУиО',
        '20_11,00': 'ПИАО',
        '20_12,00': 'ЕПБС',
        '20_13,00': 'ПОИ',
        '20_19,00': 'НСИ',
        '09_01,00': 'СУФД',
        '09_01,01': 'СУФД-Портал',
        '09_01,02': 'АРМ ОФК',
        '09_02,00': 'OEBS',
        '10_01,00': 'Аксиок.Net (децентр.)',
        '10_02,00': 'Аксиок.Net (центр.)',
        '23_01,00': 'ГАСУ-Федерация',
        '23_02,00': 'ГАСУ-Реестры',
        '23_03,00': 'ГАСУ-Аналитика',
        '23_04,00': 'ГАСУ-Портал',
        '23_05,00': 'ГАСУ-Типовое',
    }

    @classmethod
    def parse_filename(cls, filename: str) -> Tuple[bool, Optional[Dict[str, Any]], str]:
        """
        Парсит имя файла и возвращает расшифрованные компоненты.

        Args:
            filename: Имя файла для парсинга

        Returns:
            Tuple[bool, Optional[Dict], str]: (успех, словарь с данными, сообщение)
        """
        logger.info("=" * 80)
        logger.info(f"НАЧАЛО ПАРСИНГА ИМЕНИ ФАЙЛА: {filename}")
        logger.info("=" * 80)

        parts = filename.split('.')

        if len(parts) < 6:
            logger.error("Слишком мало частей в имени файла")
            return False, None, "❌ Слишком мало частей в имени файла"

        okpo = parts[0]
        is_code = parts[1]
        subsystem_code = parts[2]
        doc_code = parts[3]
        remaining = '.'.join(parts[4:])

        if '-' not in remaining:
            logger.error("В имени файла отсутствует дефис для номера документа")
            return False, None, "❌ В имени файла отсутствует дефис для номера документа"

        doc_version_parts = remaining.split('-', 1)
        doc_number_with_version = doc_version_parts[0]
        after_doc_number = doc_version_parts[1]

        doc_number_match = re.match(r'^(\d{3})', doc_number_with_version)
        if not doc_number_match:
            logger.error("Не удалось извлечь номер документа (первые 3 цифры)")
            return False, None, "❌ Не удалось извлечь номер документа (первые 3 цифры)"

        doc_number = doc_number_match.group(1)

        version_match = re.search(r'(\d{2})\.(\d{2})', after_doc_number)
        if not version_match:
            logger.error("Не удалось извлечь версию (формат XX.XX)")
            return False, None, "❌ Не удалось извлечь версию (формат XX.XX)"

        major_version = version_match.group(1)
        minor_version = version_match.group(2)
        version_end_pos = version_match.end()

        after_version = after_doc_number[version_end_pos:].strip()

        scope_match = re.match(r'^([0-9\(\);,]+)', after_version)
        if not scope_match:
            logger.error("Не удалось извлечь код области применения")
            return False, None, "❌ Не удалось извлечь код области применения"

        scope_code = scope_match.group(1)
        after_scope = after_version[len(scope_code):]

        short_name = ""
        if after_scope.startswith('_'):
            short_name = after_scope[1:]

        result = {
            'okpo': okpo,
            'is_code': is_code,
            'subsystem_code': subsystem_code,
            'doc_code': doc_code,
            'doc_number': doc_number,
            'version': f"{major_version}.{minor_version}",
            'scope_code': scope_code,
            'short_name': short_name,
            'decoded': {},
            'full_filename': filename
        }

        result['decoded']['okpo'] = f"ОКПО: {okpo}"

        is_name = cls.IS_CODES.get(is_code)
        if is_name:
            result['decoded']['is'] = f"ИС: {is_name}"
        else:
            result['decoded']['is'] = f"ИС: Неизвестный код ({is_code})"

        subsystem_codes = re.findall(r'\d{2},\d{2}', subsystem_code)

        if subsystem_codes:
            subsystem_names = []
            for code in subsystem_codes:
                subsystem_key = f"{is_code}_{code}"
                if subsystem_key in cls.SUBSYSTEM_CODES:
                    subsystem_names.append(f"{code}-{cls.SUBSYSTEM_CODES[subsystem_key]}")
                elif code == '00,00':
                    subsystem_names.append(f"{code}-без подсистем")
                elif code == '99,99':
                    subsystem_names.append(f"{code}-все подсистемы")
                else:
                    subsystem_names.append(f"{code}-?")
            result['decoded']['subsystem'] = f"Подсистемы: {'; '.join(subsystem_names)}"
        else:
            single_code = subsystem_code.strip()
            if single_code:
                subsystem_key = f"{is_code}_{single_code}"
                if subsystem_key in cls.SUBSYSTEM_CODES:
                    result['decoded']['subsystem'] = f"Подсистема: {single_code}-{cls.SUBSYSTEM_CODES[subsystem_key]}"
                elif single_code == '00,00':
                    result['decoded']['subsystem'] = f"Подсистема: без подсистем"
                elif single_code == '99,99':
                    result['decoded']['subsystem'] = f"Подсистема: все подсистемы"
                else:
                    result['decoded']['subsystem'] = f"Подсистема: {single_code}-?"

        doc_name = cls.DOCUMENT_CODES.get(doc_code)
        if doc_name:
            result['decoded']['doc'] = f"Документ: {doc_name}"
        else:
            result['decoded']['doc'] = f"Документ: Неизвестный код ({doc_code})"

        result['decoded']['number'] = f"№: {doc_number}, версия: {result['version']}"

        scope_codes_list = re.findall(r'\d+', scope_code)

        if scope_codes_list:
            scope_names = []
            for code in scope_codes_list:
                if code in cls.SCOPE_CODES:
                    scope_names.append(f"{code}-{cls.SCOPE_CODES[code]}")
                else:
                    scope_names.append(f"{code}-?")
            result['decoded']['scope'] = f"Область: {'; '.join(scope_names)}"
        else:
            result['decoded']['scope'] = "Область: не указана"

        if short_name:
            short_name_parts = short_name.split('_')
            formatted_short_name = ' / '.join(short_name_parts)
            result['decoded']['short_name'] = f"Кратко: {formatted_short_name}"

        return True, result, "✅ Имя файла соответствует стандарту ФАП"


class MainWindow(QMainWindow):
    """Главное окно приложения"""

    def __init__(self):
        super().__init__()
        self.db = JSONDatabase()
        self.checker = DocumentChecker()
        self.update_checker_config()

        self.document_text = ""
        self.selected_checks = []
        self.docx_parser = DOCXParser()
        self.current_file_path = None
        self.worker = None
        self.last_results = []
        self.page_info = []
        self.document_viewer = None

        # История документов
        self.history = DocumentHistory(db=self.db)  # Инициализация истории документов
        self.current_version_id = None    # ID текущей версии в истории

        self.document_info = {
            'filename': '',
            'file_size': 0,
            'file_type': '',
            'word_count': 0,
            'page_count': 0,
            'gk_numbers': [],
            'gk_with_subsystems': {},
            'gk_date': None,
            'first_page_gk': None,  # Добавлено поле для ГК с первой страницы
            'parsed_filename': None
        }

        self.info_extractor = DocumentInfoExtractor()

        self.settings = QSettings("ФедеральноеКазначейство", "ПроверкаДокументов")
        self.theme_mode = self.settings.value("theme_mode", "dark", type=str)
        self.dark_theme = self.theme_mode == "dark"
        self.fuzzy_threshold = float(self.settings.value("fuzzy_threshold", 70.0))
        self.fuzzy_trust_threshold = float(self.settings.value("fuzzy_trust_threshold", 85.0))
        self.auto_resize_columns = self.settings.value("auto_resize_columns", True, type=bool)
        self.show_line_numbers = self.settings.value("show_line_numbers", False, type=bool)

        self.apply_theme()
        self.init_ui()

    def update_checker_config(self):
        """Обновление конфигурации проверяльщика из базы данных"""
        enabled_checks = self.db.get_enabled_checks()
        grouped_checks = {}
        for check in enabled_checks:
            group = check.get('group', 'Без группы')
            if group not in grouped_checks:
                grouped_checks[group] = []
            grouped_checks[group].append(check)

        self.checker.config = {
            'checks': [
                {'group': group, 'subchecks': checks}
                for group, checks in grouped_checks.items()
            ]
        }

    def apply_theme(self):
        """Применить выбранную тему"""
        if self.theme_mode == "dark":
            self.setStyleSheet(self.get_dark_theme())
        elif self.theme_mode == "light":
            self.setStyleSheet(self.get_light_theme())
        else:
            self.setStyleSheet(self.get_mixed_theme())

    def get_dark_theme(self):
        """Темная тема"""
        return """
            QMainWindow {
                background-color: #1a1a1a;
            }
            QWidget {
                background-color: #1a1a1a;
                color: #ffffff;
                font-family: 'Segoe UI', 'Arial', sans-serif;
                font-size: 11px;
            }
            QLabel {
                color: #ffffff;
                font-size: 11px;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 12px;
                color: #ffffff;
                border: 2px solid #333333;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 14px;
                background-color: #2a2a2a;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px 0 8px;
                background-color: #2a2a2a;
                color: #ffffff;
                font-weight: bold;
            }
            QPushButton {
                background-color: #333333;
                color: #ffffff;
                border: 1px solid #555555;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: 500;
                font-size: 11px;
                min-height: 32px;
            }
            QPushButton:hover {
                background-color: #444444;
                border-color: #666666;
            }
            QPushButton:pressed {
                background-color: #222222;
            }
            QPushButton:disabled {
                background-color: #2a2a2a;
                color: #777777;
                border-color: #444444;
            }
            QListWidget {
                background-color: #2a2a2a;
                border: 1px solid #444444;
                border-radius: 6px;
                color: #ffffff;
                font-size: 11px;
            }
            QListWidget::item {
                padding: 6px 8px;
                border-bottom: 1px solid #333333;
                color: #ffffff;
            }
            QListWidget::item:hover {
                background-color: #3a3a3a;
            }
            QListWidget::item:selected {
                background-color: #0066cc;
                color: white;
            }
            QTableWidget {
                background-color: #2a2a2a;
                border: 1px solid #444444;
                border-radius: 6px;
                gridline-color: #333333;
                color: #ffffff;
                font-size: 11px;
                selection-background-color: #0066cc;
            }
            QHeaderView::section {
                background-color: #333333;
                color: #ffffff;
                padding: 10px;
                border: none;
                border-right: 1px solid #444444;
                border-bottom: 2px solid #444444;
                font-weight: 600;
                font-size: 11px;
            }
            QProgressBar {
                border: 1px solid #444444;
                border-radius: 6px;
                background-color: #2a2a2a;
                text-align: center;
                color: #ffffff;
                font-size: 11px;
                height: 24px;
            }
            QProgressBar::chunk {
                background-color: #00cc66;
                border-radius: 6px;
            }
            QStatusBar {
                background-color: #333333;
                color: #aaaaaa;
                border-top: 1px solid #444444;
                font-size: 11px;
            }
            QTextEdit, QTextBrowser {
                background-color: #2a2a2a;
                border: 1px solid #444444;
                border-radius: 6px;
                color: #ffffff;
                font-size: 11px;
                padding: 8px;
            }
            QLineEdit, QComboBox {
                background-color: #2a2a2a;
                border: 1px solid #555555;
                border-radius: 4px;
                padding: 5px;
                color: #ffffff;
                min-height: 25px;
            }
            QSpinBox {
                background-color: #2a2a2a;
                border: 1px solid #555555;
                border-radius: 4px;
                color: #ffffff;
                min-height: 25px;
                padding: 2px;
            }
            QMenuBar {
                background-color: #2a2a2a;
                color: #ffffff;
                border-bottom: 1px solid #444444;
                padding: 4px;
                font-size: 11px;
            }
            QMenuBar::item {
                padding: 6px 12px;
                background-color: transparent;
                color: #ffffff;
            }
            QMenuBar::item:selected {
                background-color: #3a3a3a;
                border-radius: 4px;
            }
            QMenu {
                background-color: #2a2a2a;
                border: 1px solid #444444;
                color: #ffffff;
                font-size: 11px;
            }
            QMenu::item {
                padding: 8px 24px 8px 20px;
                color: #ffffff;
            }
            QMenu::item:selected {
                background-color: #0066cc;
                color: white;
            }
            QSplitter::handle {
                background-color: #444444;
                width: 4px;
            }
            QSplitter::handle:hover {
                background-color: #666666;
            }
        """

    def get_light_theme(self):
        """Светлая тема"""
        return """
            QMainWindow {
                background-color: #f5f5f5;
            }
            QWidget {
                background-color: #f5f5f5;
                color: #000000;
                font-family: 'Segoe UI', 'Arial', sans-serif;
                font-size: 11px;
            }
            QLabel {
                color: #000000;
                font-size: 11px;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 12px;
                color: #2c3e50;
                border: 2px solid #bdc3c7;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 14px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px 0 8px;
                background-color: #ffffff;
                color: #2c3e50;
                font-weight: bold;
            }
            QPushButton {
                background-color: #ffffff;
                color: #2c3e50;
                border: 1px solid #bdc3c7;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: 500;
                font-size: 11px;
                min-height: 32px;
            }
            QPushButton:hover {
                background-color: #ecf0f1;
                border-color: #95a5a6;
            }
            QPushButton:pressed {
                background-color: #d5dbdb;
            }
            QPushButton:disabled {
                background-color: #f8f9fa;
                color: #95a5a6;
                border-color: #d5dbdb;
            }
            QListWidget {
                background-color: #ffffff;
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                color: #000000;
                font-size: 11px;
            }
            QListWidget::item {
                padding: 6px 8px;
                border-bottom: 1px solid #ecf0f1;
                color: #000000;
            }
            QListWidget::item:hover {
                background-color: #f8f9fa;
            }
            QListWidget::item:selected {
                background-color: #3498db;
                color: white;
            }
            QTableWidget {
                background-color: #ffffff;
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                gridline-color: #ecf0f1;
                color: #000000;
                font-size: 11px;
                selection-background-color: #3498db;
            }
            QHeaderView::section {
                background-color: #ecf0f1;
                color: #2c3e50;
                padding: 10px;
                border: none;
                border-right: 1px solid #d5dbdb;
                border-bottom: 2px solid #bdc3c7;
                font-weight: 600;
                font-size: 11px;
            }
            QProgressBar {
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                background-color: #ffffff;
                text-align: center;
                color: #000000;
                font-size: 11px;
                height: 24px;
            }
            QProgressBar::chunk {
                background-color: #2ecc71;
                border-radius: 6px;
            }
            QStatusBar {
                background-color: #ecf0f1;
                color: #7f8c8d;
                border-top: 1px solid #bdc3c7;
                font-size: 11px;
            }
            QTextEdit, QTextBrowser {
                background-color: #ffffff;
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                color: #000000;
                font-size: 11px;
                padding: 8px;
            }
            QLineEdit, QComboBox {
                background-color: #ffffff;
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                padding: 5px;
                color: #000000;
                min-height: 25px;
            }
            QSpinBox {
                background-color: #ffffff;
                border: 1px solid #bdc3c7;
                border-radius: 4px;
                color: #000000;
                min-height: 25px;
                padding: 2px;
            }
            QMenuBar {
                background-color: #ffffff;
                color: #2c3e50;
                border-bottom: 1px solid #bdc3c7;
                padding: 4px;
                font-size: 11px;
            }
            QMenuBar::item {
                padding: 6px 12px;
                background-color: transparent;
                color: #2c3e50;
            }
            QMenuBar::item:selected {
                background-color: #ecf0f1;
                border-radius: 4px;
            }
            QMenu {
                background-color: #ffffff;
                border: 1px solid #bdc3c7;
                color: #2c3e50;
                font-size: 11px;
            }
            QMenu::item {
                padding: 8px 24px 8px 20px;
                color: #2c3e50;
            }
            QMenu::item:selected {
                background-color: #3498db;
                color: white;
            }
            QSplitter::handle {
                background-color: #bdc3c7;
                width: 4px;
            }
            QSplitter::handle:hover {
                background-color: #95a5a6;
            }
        """

    def get_mixed_theme(self):
        """Смешанная тема"""
        return """
            QMainWindow {
                background-color: #f0f2f5;
            }
            QWidget {
                background-color: #f0f2f5;
                color: #333333;
                font-family: 'Segoe UI', 'Arial', sans-serif;
                font-size: 11px;
            }
            QLabel {
                color: #333333;
                font-size: 11px;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 12px;
                color: #2c3e50;
                border: 2px solid #d1d9e6;
                border-radius: 8px;
                margin-top: 12px;
                padding-top: 14px;
                background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 8px 0 8px;
                background-color: #ffffff;
                color: #2c3e50;
                font-weight: bold;
            }
            QPushButton {
                background-color: #ffffff;
                color: #2c3e50;
                border: 1px solid #d1d9e6;
                padding: 8px 16px;
                border-radius: 6px;
                font-weight: 500;
                font-size: 11px;
                min-height: 32px;
            }
            QPushButton:hover {
                background-color: #e8ecf1;
                border-color: #3498db;
            }
            QPushButton:pressed {
                background-color: #d6dde7;
            }
            QPushButton:disabled {
                background-color: #f8f9fa;
                color: #95a5a6;
                border-color: #d5dbdb;
            }
            QListWidget {
                background-color: #ffffff;
                border: 1px solid #d1d9e6;
                border-radius: 6px;
                color: #333333;
                font-size: 11px;
            }
            QListWidget::item {
                padding: 6px 8px;
                border-bottom: 1px solid #f0f2f5;
                color: #333333;
            }
            QListWidget::item:hover {
                background-color: #f8f9fa;
            }
            QListWidget::item:selected {
                background-color: #3498db;
                color: white;
            }
            QTableWidget {
                background-color: #ffffff;
                border: 1px solid #d1d9e6;
                border-radius: 6px;
                gridline-color: #f0f2f5;
                color: #333333;
                font-size: 11px;
                selection-background-color: #3498db;
            }
            QHeaderView::section {
                background-color: #2c3e50;
                color: #ffffff;
                padding: 10px;
                border: none;
                border-right: 1px solid #34495e;
                border-bottom: 2px solid #34495e;
                font-weight: 600;
                font-size: 11px;
            }
            QProgressBar {
                border: 1px solid #d1d9e6;
                border-radius: 6px;
                background-color: #ffffff;
                text-align: center;
                color: #333333;
                font-size: 11px;
                height: 24px;
            }
            QProgressBar::chunk {
                background-color: #2ecc71;
                border-radius: 6px;
            }
            QStatusBar {
                background-color: #2c3e50;
                color: #ecf0f1;
                border-top: 1px solid #34495e;
                font-size: 11px;
            }
            QTextEdit, QTextBrowser {
                background-color: #ffffff;
                border: 1px solid #d1d9e6;
                border-radius: 6px;
                color: #333333;
                font-size: 11px;
                padding: 8px;
            }
            QLineEdit, QComboBox {
                background-color: #ffffff;
                border: 1px solid #d1d9e6;
                border-radius: 4px;
                padding: 5px;
                color: #333333;
                min-height: 25px;
            }
            QSpinBox {
                background-color: #ffffff;
                border: 1px solid #d1d9e6;
                border-radius: 4px;
                color: #333333;
                min-height: 25px;
                padding: 2px;
            }
            QMenuBar {
                background-color: #2c3e50;
                color: #ecf0f1;
                border-bottom: 1px solid #34495e;
                padding: 4px;
                font-size: 11px;
            }
            QMenuBar::item {
                padding: 6px 12px;
                background-color: transparent;
                color: #ecf0f1;
            }
            QMenuBar::item:selected {
                background-color: #3498db;
                border-radius: 4px;
            }
            QMenu {
                background-color: #2c3e50;
                border: 1px solid #34495e;
                color: #ecf0f1;
                font-size: 11px;
            }
            QMenu::item {
                padding: 8px 24px 8px 20px;
                color: #ecf0f1;
            }
            QMenu::item:selected {
                background-color: #3498db;
                color: white;
            }
            QSplitter::handle {
                background-color: #d1d9e6;
                width: 4px;
            }
            QSplitter::handle:hover {
                background-color: #3498db;
            }
            QDialog {
                background-color: #ffffff;
            }
            QScrollBar:vertical {
                border: none;
                background: #f0f2f5;
                width: 10px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #c1c9d6;
                min-height: 20px;
                border-radius: 5px;
            }
            QScrollBar::handle:vertical:hover {
                background: #a7b1c2;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                border: none;
                background: none;
            }
        """

    def save_settings(self):
        """Сохранить настройки"""
        self.settings.setValue("theme_mode", self.theme_mode)
        self.settings.setValue("dark_theme", self.dark_theme)
        self.settings.setValue("fuzzy_threshold", self.fuzzy_threshold)
        self.settings.setValue("fuzzy_trust_threshold", self.fuzzy_trust_threshold)
        self.settings.setValue("auto_resize_columns", self.auto_resize_columns)
        self.settings.setValue("show_line_numbers", self.show_line_numbers)

    def init_ui(self):
        """Инициализация пользовательского интерфейса"""
        self.setWindowTitle("Система проверки технической документации - Федеральное казначейство v3.0.0")
        self.setGeometry(100, 50, 1700, 950)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)

        # Верхняя панель с кнопками
        top_panel = QWidget()
        top_layout = QHBoxLayout(top_panel)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(10)

        left_buttons = QHBoxLayout()

        self.load_doc_btn = QPushButton("📂 Загрузить документ")
        self.load_doc_btn.clicked.connect(self.load_document)
        self.load_doc_btn.setMinimumHeight(40)
        self.load_doc_btn.setMinimumWidth(200)
        left_buttons.addWidget(self.load_doc_btn)

        self.view_with_errors_btn = QPushButton("🔍 С ошибками")
        self.view_with_errors_btn.clicked.connect(self.view_document_with_errors)
        self.view_with_errors_btn.setEnabled(False)
        self.view_with_errors_btn.setMinimumHeight(40)
        self.view_with_errors_btn.setMinimumWidth(120)
        left_buttons.addWidget(self.view_with_errors_btn)

        self.manage_checks_btn = QPushButton("⚙️ Управление проверками")
        self.manage_checks_btn.clicked.connect(self.open_manage_checks)
        self.manage_checks_btn.setMinimumHeight(40)
        self.manage_checks_btn.setMinimumWidth(180)
        self.manage_checks_btn.setStyleSheet("""
            QPushButton {
                background-color: #9b59b6;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #8e44ad;
            }
        """)
        left_buttons.addWidget(self.manage_checks_btn)

        self.versions_btn = QPushButton("📊 Версии ПО")
        self.versions_btn.clicked.connect(self.show_versions_dialog)
        self.versions_btn.setEnabled(False)
        self.versions_btn.setMinimumHeight(40)
        self.versions_btn.setMinimumWidth(120)
        self.versions_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #7f8c8d;
            }
        """)
        left_buttons.addWidget(self.versions_btn)

        # Кнопка истории документов
        self.history_btn = QPushButton("📜 История")
        self.history_btn.clicked.connect(self.open_document_history)
        self.history_btn.setEnabled(True)
        self.history_btn.setMinimumHeight(40)
        self.history_btn.setMinimumWidth(100)
        self.history_btn.setStyleSheet("""
            QPushButton {
                background-color: #9b59b6;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #8e44ad;
            }
        """)
        left_buttons.addWidget(self.history_btn)

        # Кнопка для добавления комментария
        self.comment_btn = QPushButton("💬 Комментарий")
        self.comment_btn.clicked.connect(self.add_comment_to_current_version)
        self.comment_btn.setEnabled(True)
        self.comment_btn.setMinimumHeight(40)
        self.comment_btn.setMinimumWidth(120)
        self.comment_btn.setStyleSheet("""
            QPushButton {
                background-color: #f39c12;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e67e22;
            }
        """)
        left_buttons.addWidget(self.comment_btn)

        # Кнопка для добавления тега
        self.tag_btn = QPushButton("🏷️ Добавить тег")
        self.tag_btn.clicked.connect(self.add_tag_to_current_version)
        self.tag_btn.setEnabled(True)
        self.tag_btn.setMinimumHeight(40)
        self.tag_btn.setMinimumWidth(120)
        self.tag_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #229954;
            }
        """)
        left_buttons.addWidget(self.tag_btn)

        top_layout.addLayout(left_buttons)
        top_layout.addStretch()

        # Правая панель с кнопками экспорта
        right_buttons = QHBoxLayout()

        export_buttons = [
            ("📄 PDF", self.export_pdf),
            ("📊 Excel", self.export_excel),
            ("📝 ODT", self.export_odt),
            ("📈 ODS", self.export_ods),
        ]

        for text, slot in export_buttons:
            btn = QPushButton(text)
            btn.clicked.connect(slot)
            btn.setMinimumHeight(35)
            btn.setMinimumWidth(90)
            right_buttons.addWidget(btn)

        top_layout.addLayout(right_buttons)

        main_layout.addWidget(top_panel)

        # Панель статистики и прогресса
        stats_panel = QWidget()
        stats_layout = QHBoxLayout(stats_panel)
        stats_layout.setContentsMargins(0, 0, 0, 0)
        stats_layout.setSpacing(15)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setMinimumHeight(30)
        stats_layout.addWidget(self.progress_bar, 2)

        stats_widget = QWidget()
        stats_widget.setStyleSheet("background-color: #2c3e50; border-radius: 5px; padding: 5px;")
        inner_stats = QHBoxLayout(stats_widget)
        inner_stats.setContentsMargins(10, 5, 10, 5)
        inner_stats.setSpacing(20)

        self.total_label = QLabel("📊 Всего: 0")
        self.passed_label = QLabel("✅ Пройдено: 0")
        self.failed_label = QLabel("❌ Провалено: 0")
        self.warning_label = QLabel("⚠ Проверить: 0")
        self.time_label = QLabel("⏱ Время: 0с")

        for label in [self.total_label, self.passed_label, self.failed_label, self.warning_label, self.time_label]:
            label.setStyleSheet("color: white; font-weight: bold; font-size: 12px;")
            inner_stats.addWidget(label)

        stats_layout.addWidget(stats_widget, 3)

        self.run_check_btn = QPushButton("▶ ЗАПУСТИТЬ ПРОВЕРКУ")
        self.run_check_btn.clicked.connect(self.run_check)
        self.run_check_btn.setEnabled(False)
        self.run_check_btn.setMinimumHeight(40)
        self.run_check_btn.setMinimumWidth(200)
        self.run_check_btn.setStyleSheet("""
            QPushButton {
                background-color: #27ae60;
                color: white;
                font-weight: bold;
                font-size: 14px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #229954;
            }
            QPushButton:disabled {
                background-color: #7f8c8d;
            }
        """)
        stats_layout.addWidget(self.run_check_btn)

        main_layout.addWidget(stats_panel)

        # Основной сплиттер
        content_splitter = QSplitter(Qt.Orientation.Horizontal)
        content_splitter.setChildrenCollapsible(False)

        # Левая панель с информацией
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(5, 5, 5, 5)
        left_layout.setSpacing(10)

        main_info_group = QGroupBox("📋 Основная информация")
        main_info_layout = QVBoxLayout(main_info_group)

        self.doc_name_label = QLabel("Файл: не загружен")
        self.doc_name_label.setWordWrap(True)
        self.doc_name_label.setStyleSheet("font-weight: bold; font-size: 12px; padding: 5px;")
        main_info_layout.addWidget(self.doc_name_label)

        self.doc_stats_label = QLabel("Статистика: -")
        self.doc_stats_label.setStyleSheet("padding: 5px;")
        main_info_layout.addWidget(self.doc_stats_label)

        # Добавляем информацию о ГК и дате в основную информацию
        gk_info_widget = QWidget()
        gk_info_layout = QHBoxLayout(gk_info_widget)
        gk_info_layout.setContentsMargins(0, 0, 0, 0)
        gk_info_layout.setSpacing(10)

        gk_info_layout.addWidget(QLabel("📅 Дата ГК:"))
        self.gk_date_label = QLabel("не найдена")
        self.gk_date_label.setStyleSheet("color: #e67e22; font-weight: bold;")
        gk_info_layout.addWidget(self.gk_date_label)

        gk_info_layout.addWidget(QLabel("|"))

        gk_info_layout.addWidget(QLabel("🔑 ГК:"))
        self.gk_number_label = QLabel("не найден")
        self.gk_number_label.setStyleSheet("color: #3498db; font-weight: bold;")
        gk_info_layout.addWidget(self.gk_number_label)

        gk_info_layout.addStretch()

        main_info_layout.addWidget(gk_info_widget)

        # Информация о текущей версии
        self.version_info_label = QLabel("📜 Версия в истории: не сохранена")
        self.version_info_label.setStyleSheet("color: #3498db; font-size: 10px; padding: 2px;")
        main_info_layout.addWidget(self.version_info_label)

        left_layout.addWidget(main_info_group)

        # Группа расшифровки имени файла
        filename_group = QGroupBox("📁 Расшифровка имени файла")
        filename_layout = QVBoxLayout(filename_group)

        self.filename_info = QTextEdit()
        self.filename_info.setReadOnly(True)
        self.filename_info.setMaximumHeight(200)
        self.filename_info.setStyleSheet("font-family: monospace; font-size: 11px;")
        filename_layout.addWidget(self.filename_info)

        left_layout.addWidget(filename_group)

        left_layout.addStretch()

        # Правая панель с результатами
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(5, 5, 5, 5)
        right_layout.setSpacing(10)

        # Панель фильтров
        filter_panel = QWidget()
        filter_layout = QHBoxLayout(filter_panel)
        filter_layout.setContentsMargins(0, 0, 0, 0)
        filter_layout.setSpacing(10)

        self.results_search_input = QLineEdit()
        self.results_search_input.setPlaceholderText("🔍 Поиск в результатах...")
        self.results_search_input.textChanged.connect(self.filter_results_table)
        self.results_search_input.setMinimumHeight(30)

        self.filter_combo = QComboBox()
        self.filter_combo.addItems(["Все", "❌ Провалено", "✓ Пройдено", "⚠ Требует проверки"])
        self.filter_combo.currentTextChanged.connect(self.filter_results_table)
        self.filter_combo.setMinimumHeight(30)
        self.filter_combo.setMinimumWidth(150)

        filter_layout.addWidget(self.results_search_input, 2)
        filter_layout.addWidget(QLabel("Фильтр:"))
        filter_layout.addWidget(self.filter_combo)

        right_layout.addWidget(filter_panel)

        # Таблица результатов
        self.results_table = QTableWidget()
        self.results_table.setColumnCount(7)
        self.results_table.setHorizontalHeaderLabels(
            ["Проверка", "Группа", "Статус", "Результат", "Стр.", "Позиция", "Детали"])

        header = self.results_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.Stretch)

        self.results_table.setColumnWidth(0, 250)
        self.results_table.setColumnWidth(1, 150)
        self.results_table.setColumnWidth(3, 150)
        self.results_table.setColumnWidth(4, 50)
        self.results_table.setColumnWidth(5, 100)

        self.results_table.setAlternatingRowColors(True)
        self.results_table.setSortingEnabled(True)
        self.results_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        self.results_table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.results_table.customContextMenuRequested.connect(self.show_results_context_menu)

        self.results_table.doubleClicked.connect(self.go_to_error_from_table)

        right_layout.addWidget(self.results_table)

        content_splitter.addWidget(left_panel)
        content_splitter.addWidget(right_panel)
        content_splitter.setSizes([400, 1300])

        main_layout.addWidget(content_splitter, 1)

        # Статус бар
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_label = QLabel("✅ Готово к работе")
        self.status_bar.addPermanentWidget(self.status_label)
        self.status_bar.showMessage("Федеральное казначейство • Система проверки технической документации • v3.0.0")

        self.create_menu()

    def create_menu(self):
        """Создание меню приложения"""
        menubar = self.menuBar()

        # Меню Файл
        file_menu = menubar.addMenu("Файл")

        load_action = QAction("Загрузить документ", self)
        load_action.triggered.connect(self.load_document)
        load_action.setShortcut("Ctrl+O")
        file_menu.addAction(load_action)

        new_check_action = QAction("Добавить новую проверку", self)
        new_check_action.triggered.connect(self.add_new_check)
        new_check_action.setShortcut("Ctrl+N")
        file_menu.addAction(new_check_action)

        config_action = QAction("Редактировать конфигурацию", self)
        config_action.triggered.connect(self.open_config_editor)
        config_action.setShortcut("Ctrl+E")
        file_menu.addAction(config_action)

        settings_action = QAction("Настройки", self)
        settings_action.triggered.connect(self.open_settings)
        settings_action.setShortcut("Ctrl+P")
        file_menu.addAction(settings_action)

        file_menu.addSeparator()

        exit_action = QAction("Выход", self)
        exit_action.triggered.connect(self.close)
        exit_action.setShortcut("Ctrl+Q")
        file_menu.addAction(exit_action)

        # Меню Вид
        view_menu = menubar.addMenu("Вид")

        history_action = QAction("📜 История документа", self)
        history_action.triggered.connect(self.open_document_history)
        history_action.setShortcut("Ctrl+H")
        view_menu.addAction(history_action)

        comments_action = QAction("💬 Комментарии к текущей версии", self)
        comments_action.triggered.connect(self.add_comment_to_current_version)
        view_menu.addAction(comments_action)

        tags_action = QAction("🏷️ Добавить тег", self)
        tags_action.triggered.connect(self.add_tag_to_current_version)
        view_menu.addAction(tags_action)

        # Меню Проверка
        check_menu = menubar.addMenu("Проверка")

        run_action = QAction("Запустить проверку", self)
        run_action.triggered.connect(self.run_check)
        run_action.setShortcut("F5")
        check_menu.addAction(run_action)

        versions_action = QAction("Версии БПО/СПО", self)
        versions_action.triggered.connect(self.show_versions_dialog)
        versions_action.setShortcut("Ctrl+V")
        check_menu.addAction(versions_action)

        manage_checks_action = QAction("Управление проверками", self)
        manage_checks_action.triggered.connect(self.open_manage_checks)
        manage_checks_action.setShortcut("Ctrl+M")
        check_menu.addAction(manage_checks_action)

        # Меню Справка
        help_menu = menubar.addMenu("Справка")

        about_action = QAction("О программе", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    # ==================== МЕТОДЫ ДЛЯ РАБОТЫ С ИСТОРИЕЙ ====================

    def save_to_history(self, auto_save: bool = False):
        """
        Сохранение текущего документа и результатов в историю

        Args:
            auto_save: True если это автосохранение при закрытии
        """
        if not self.document_text:
            return None

        logger.info("=" * 50)
        logger.info("СОХРАНЕНИЕ В ИСТОРИЮ")
        logger.info(f"Auto save: {auto_save}")
        logger.info(f"Текущая версия: {self.current_version_id}")
        logger.info(f"Есть результаты: {bool(self.last_results)}")
        if self.last_results:
            logger.info(f"Количество результатов: {len(self.last_results)}")

        # Проверяем, изменился ли документ по сравнению с последней сохраненной версией
        if self.current_version_id:
            last_version = self.history.get_version(self.current_version_id)
            if last_version:
                logger.info(f"Последняя версия: {last_version.version_id}")
                logger.info(f"Текст совпадает: {last_version.document_text == self.document_text}")

                if last_version.document_text == self.document_text:
                    # Текст не изменился, просто обновляем результаты если они есть
                    if self.last_results and last_version.check_results != self.last_results:
                        logger.info("Обновление результатов существующей версии")
                        last_version.check_results = self.last_results
                        last_version.update_stats()

                        # Получаем метаданные версии из истории
                        version_info = self.history.history_data['versions'].get(self.current_version_id, {})
                        if version_info:
                            # Обновляем данные
                            version_info['check_results'] = self.last_results
                            version_info['stats'] = last_version.stats
                            version_info['last_updated'] = datetime.now().isoformat()

                            # Сохраняем
                            self.history._save_history_data()

                            # Если это автосохранение, добавляем пометку
                            if auto_save:
                                self.history.add_comment_to_version(
                                    self.current_version_id,
                                    "Автосохранение при закрытии",
                                    "system"
                                )

                            logger.info(f"Обновлены результаты для версии {self.current_version_id}")

                    return last_version

        # Создаем новую версию
        logger.info("СОЗДАНИЕ НОВОЙ ВЕРСИИ")
        metadata = {
            'theme': self.theme_mode,
            'settings': {
                'fuzzy_threshold': self.fuzzy_threshold,
                'fuzzy_trust_threshold': self.fuzzy_trust_threshold
            }
        }

        # Добавляем пометку об автосохранении
        if auto_save:
            metadata['save_type'] = 'auto_close'

        version = self.history.add_version(
            file_path=self.current_file_path or "unknown",
            document_text=self.document_text,
            check_results=self.last_results,
            document_info=self.document_info,
            metadata=metadata,
            auto_save=auto_save
        )

        # Если это автосохранение, добавляем комментарий
        if auto_save:
            self.history.add_comment_to_version(
                version.version_id,
                "Автоматически сохранено при закрытии приложения",
                "system"
            )

        self.current_version_id = version.version_id
        logger.info(f"СОЗДАНА НОВАЯ ВЕРСИЯ: {version.version_id}")

        # Выводим отладочную информацию
        self.history.debug_info()

        # Обновляем статус-бар
        save_type = "Автосохранение" if auto_save else "Сохранение"
        self.status_bar.showMessage(
            f"{save_type}: Версия {version.version_id[:16]}..."
        )

        # Обновляем информацию о версии в левой панели
        self.update_document_info()

        return version

    def open_document_history(self):
        """Открыть диалог истории документа"""
        dialog = DocumentHistoryDialog(self.history, self.current_file_path, self)
        dialog.version_selected.connect(self.load_version_from_history)
        dialog.exec()

    def load_version_from_history(self, version_id: str):
        """Загрузка версии из истории"""
        version = self.history.get_version(version_id)

        if not version:
            QMessageBox.warning(self, "Ошибка", "Не удалось загрузить версию")
            return

        # Спрашиваем, сохранять ли текущую версию перед загрузкой
        if self.document_text:
            reply = QMessageBox.question(
                self, "Сохранение",
                "Сохранить текущий документ в истории перед загрузкой другой версии?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
            )

            if reply == QMessageBox.StandardButton.Yes:
                self.save_to_history()
            elif reply == QMessageBox.StandardButton.Cancel:
                return

        # Загружаем текст документа
        self.document_text = version.document_text

        # Загружаем информацию о документе
        self.document_info = version.document_info.copy()
        self.current_file_path = version.file_path if os.path.exists(version.file_path) else None

        # Загружаем результаты проверки
        self.last_results = version.check_results

        # Восстанавливаем информацию о страницах
        self.calculate_page_info()

        # Обновляем интерфейс
        self.update_document_info()

        if self.last_results:
            self.display_results(self.last_results)

            total = len(self.last_results)
            passed = sum(1 for r in self.last_results if r.get('passed', False) and not r.get('needs_verification', False))
            failed = sum(1 for r in self.last_results if not r.get('passed', False) and not r.get('needs_verification', False))
            warning = sum(1 for r in self.last_results if r.get('needs_verification', False))

            self.update_stats(total, passed, failed, warning, 0)
            self.view_with_errors_btn.setEnabled(failed > 0 or warning > 0)

        self.run_check_btn.setEnabled(True)
        self.versions_btn.setEnabled(True)
        self.current_version_id = version_id

        # Обновляем информацию о версии
        self.update_document_info()

        # Показываем информацию о комментариях
        if version.comments:
            QMessageBox.information(
                self, "Комментарии к версии",
                f"У этой версии есть {len(version.comments)} комментариев.\n"
                f"Последний: {version.comments[-1].get('text', '')[:100]}..."
            )

        logger.info(f"Загружена версия из истории: {version_id}")

        QMessageBox.information(self, "Успех",
                               f"Версия {version_id[:16]}... загружена из истории")

    def add_comment_to_current_version(self):
        """Добавление комментария к текущей версии"""
        if not self.current_version_id:
            QMessageBox.warning(self, "Ошибка", "Нет текущей версии в истории. Сначала загрузите или сохраните документ.")
            return

        comment, ok = QInputDialog.getMultiLineText(
            self, "Добавить комментарий",
            "Введите комментарий к текущей версии:"
        )

        if ok and comment:
            if self.history.add_comment_to_version(self.current_version_id, comment, "user"):
                QMessageBox.information(self, "Успех", "Комментарий добавлен")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось добавить комментарий")

    def add_tag_to_current_version(self):
        """Добавление тега к текущей версии"""
        if not self.current_version_id:
            QMessageBox.warning(self, "Ошибка", "Нет текущей версии в истории. Сначала загрузите или сохраните документ.")
            return

        tag, ok = QInputDialog.getText(
            self, "Добавить тег",
            "Введите тег для текущей версии:"
        )

        if ok and tag:
            if self.history.add_tag_to_version(self.current_version_id, tag):
                QMessageBox.information(self, "Успех", f"Тег '{tag}' добавлен")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось добавить тег")

    # ==================== МЕТОДЫ ЗАГРУЗКИ ДОКУМЕНТА ====================

    def load_document(self):
        """Загрузка документа (DOCX, TXT)"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Загрузить документ", "",
            "Документы (*.docx *.txt);;Word files (*.docx);;Text files (*.txt);;All files (*.*)"
        )

        if not file_path:
            return

        try:
            self.current_file_path = file_path
            file_path_obj = Path(file_path)
            file_name = file_path_obj.name
            file_size = file_path_obj.stat().st_size / (1024 * 1024)

            logger.info(f"Загрузка файла: {file_name}")
            logger.info(f"Размер файла: {file_size:.2f} МБ")

            self.document_info['filename'] = file_name
            self.document_info['file_size'] = file_size

            is_valid, parsed_data, parse_message = FilenameParser.parse_filename(file_name)

            if is_valid and parsed_data:
                self.document_info['parsed_filename'] = parsed_data
                logger.info("Имя файла успешно распознано")
            else:
                logger.warning(f"Имя файла не соответствует стандарту: {file_name}")
                self.document_info['parsed_filename'] = None

                reply = QMessageBox.warning(
                    self, "Проверка имени файла",
                    f"<b>⚠️ Имя файла не соответствует стандарту ФАП!</b><br><br>"
                    f"{parse_message}<br><br>"
                    f"<b>Продолжить загрузку?</b>",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
                )
                if reply == QMessageBox.StandardButton.No:
                    logger.info("Загрузка отменена пользователем из-за несоответствия имени файла")
                    return

            if file_name.lower().endswith('.docx'):
                self.document_text = self.docx_parser.extract_text_from_docx(file_path)
                file_type = "DOCX"
                self.document_info['file_type'] = "DOCX"
                logger.info("Файл типа DOCX успешно загружен")

            elif file_name.lower().endswith('.txt'):
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        self.document_text = f.read()
                    logger.info("Файл типа TXT успешно загружен (UTF-8)")
                except UnicodeDecodeError:
                    with open(file_path, 'r', encoding='cp1251') as f:
                        self.document_text = f.read()
                    logger.info("Файл типа TXT успешно загружен (CP1251)")
                file_type = "TXT"
                self.document_info['file_type'] = "TXT"

            else:
                QMessageBox.warning(self, "Ошибка", "Поддерживаются только DOCX и TXT файлы")
                logger.warning(f"Неподдерживаемый тип файла: {file_name}")
                return

            self.calculate_page_info()

            word_count = len(self.document_text.split())

            self.document_info['word_count'] = word_count
            self.document_info['page_count'] = len(self.page_info)

            self.document_info['gk_numbers'] = self.info_extractor.extract_gk_numbers(self.document_text,
                                                                                      self.page_info)
            self.document_info['gk_date'] = self.info_extractor.get_gk_date()
            self.document_info['first_page_gk'] = self.info_extractor.get_first_page_gk()

            self.update_document_info()

            self.run_check_btn.setEnabled(True)
            self.view_with_errors_btn.setEnabled(False)
            self.versions_btn.setEnabled(True)

            self.results_table.setRowCount(0)
            self.last_results = []
            self.update_stats(0, 0, 0, 0, 0)

            # Сохраняем в историю (без результатов проверки пока)
            self.save_to_history()

            logger.info(f"Документ загружен: {file_name}, размер: {file_size:.2f} МБ, страниц: {len(self.page_info)}")
            logger.info(f"Найдено ГК: {len(self.document_info['gk_numbers'])}")
            if self.document_info['gk_date']:
                logger.info(f"Дата ГК: {self.document_info['gk_date']}")
            if self.document_info['first_page_gk']:
                logger.info(f"Первый ГК с первой страницы: {self.document_info['first_page_gk']}")

            QMessageBox.information(self, "Успех",
                                    f"Документ успешно загружен!\n\nНайдено ГК: {len(self.document_info['gk_numbers'])}")

        except Exception as e:
            logger.error(f"Ошибка загрузки документа: {str(e)}")
            logger.exception(e)
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить документ:\n{str(e)}")

    def calculate_page_info(self):
        """Вычисление информации о страницах документа"""
        self.page_info = []

        if not self.document_text:
            return

        lines = self.document_text.split('\n')
        current_page_text = []
        current_chars = 0
        current_start = 0

        chars_per_page = 1800
        max_lines_per_page = 50

        for line in lines:
            line_length = len(line)

            if (current_chars + line_length > chars_per_page and current_chars > 0) or \
                    (len(current_page_text) >= max_lines_per_page):

                page_text = '\n'.join(current_page_text)
                page_end = current_start + len(page_text)
                self.page_info.append((current_start, page_end))

                current_page_text = [line]
                current_chars = line_length
                current_start = page_end + 1
            else:
                current_page_text.append(line)
                current_chars += line_length + 1

        if current_page_text:
            page_text = '\n'.join(current_page_text)
            page_end = current_start + len(page_text)
            self.page_info.append((current_start, page_end))

        logger.info(f"Документ разбит на {len(self.page_info)} страниц")

    def update_document_info(self):
        """Обновление информации о документе в левой панели"""
        if self.current_file_path:
            file_name = self.document_info['filename']
            file_type = self.document_info['file_type']
            pages = self.document_info['page_count']
            words = self.document_info['word_count']
            size = self.document_info['file_size']

            self.doc_name_label.setText(f"📄 Файл: {file_name}")
            self.doc_stats_label.setText(f"📊 Тип: {file_type}, {pages} стр., {words} слов, {size:.2f} МБ")

            # Обновляем информацию о ГК
            if self.document_info['gk_date']:
                self.gk_date_label.setText(self.document_info['gk_date'])
                self.gk_date_label.setStyleSheet("color: #e67e22; font-weight: bold;")
            else:
                self.gk_date_label.setText("не найдена")
                self.gk_date_label.setStyleSheet("")

            # Обновляем номер ГК (берем первый с первой страницы или первый найденный)
            if self.document_info['first_page_gk']:
                gk_text = self.document_info['first_page_gk']
            elif self.document_info['gk_numbers']:
                gk_text = self.document_info['gk_numbers'][0]
            else:
                gk_text = "не найден"

            self.gk_number_label.setText(gk_text)

            # Информация о версии в истории
            if self.current_version_id:
                version = self.history.get_version(self.current_version_id)
                if version:
                    comment_count = len(version.comments)
                    tag_count = len(version.tags)
                    self.version_info_label.setText(
                        f"📜 Версия: {self.current_version_id[:16]}... "
                        f"(💬 {comment_count} | 🏷️ {tag_count})"
                    )
                    self.version_info_label.setStyleSheet("color: #27ae60; font-size: 10px; padding: 2px;")
                else:
                    self.version_info_label.setText(f"📜 Версия: {self.current_version_id[:16]}...")
                    self.version_info_label.setStyleSheet("color: #27ae60; font-size: 10px; padding: 2px;")
            else:
                self.version_info_label.setText("📜 Версия в истории: не сохранена")
                self.version_info_label.setStyleSheet("color: #3498db; font-size: 10px; padding: 2px;")

            if self.document_info['parsed_filename']:
                parsed = self.document_info['parsed_filename']
                decoded = parsed['decoded']

                text = "✅ Соответствует стандарту ФАП\n\n"
                for key, value in decoded.items():
                    text += f"• {value}\n"

                self.filename_info.setText(text)
            else:
                self.filename_info.setText("❌ Имя файла не соответствует стандарту ФАП")
        else:
            self.doc_name_label.setText("📄 Файл: не загружен")
            self.doc_stats_label.setText("📊 Статистика: -")
            self.version_info_label.setText("📜 Версия в истории: не сохранена")
            self.gk_date_label.setText("не найдена")
            self.gk_number_label.setText("не найден")
            self.filename_info.setText("Информация отсутствует")

    # ==================== МЕТОДЫ ПРОВЕРКИ ====================

    def run_check(self):
        """Запуск проверки"""
        if not self.document_text:
            QMessageBox.warning(self, "Ошибка", "Сначала загрузите документ")
            return

        self.selected_checks = []
        self.results_table.setRowCount(0)

        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.run_check_btn.setEnabled(False)
        self.status_label.setText("⏳ Выполняется проверка...")

        logger.info(f"Запуск проверки")

        self.update_checker_config()
        all_checks = []
        for group in self.checker.config.get('checks', []):
            for subcheck in group.get('subchecks', []):
                all_checks.append(subcheck.get('name', ''))

        self.worker = CheckWorker(self.checker, self.document_text, all_checks, self.checker.config,
                                  self.page_info)
        self.worker.progress.connect(self.update_progress)
        self.worker.finished.connect(self.on_check_finished)
        self.start_time = time.time()
        self.worker.start()

    def update_progress(self, value, current_check):
        """Обновление прогресса"""
        self.progress_bar.setValue(value)
        self.status_label.setText(f"⏳ Проверка: {current_check}")

    def on_check_finished(self, results):
        """Обработка завершения проверки"""
        elapsed_time = time.time() - self.start_time
        self.last_results = results

        # Сохраняем результаты в историю
        self.save_to_history()

        self.display_results(results)

        total = len(results)
        passed = sum(1 for r in results if r['passed'] and not r['needs_verification'])
        failed = sum(1 for r in results if not r['passed'] and not r['needs_verification'])
        warning = sum(1 for r in results if r['needs_verification'])

        self.update_stats(total, passed, failed, warning, elapsed_time)

        self.progress_bar.setVisible(False)
        self.run_check_btn.setEnabled(True)
        self.status_label.setText("✅ Проверка завершена")

        self.view_with_errors_btn.setEnabled(failed > 0 or warning > 0)

        if self.auto_resize_columns:
            self.resize_table_columns()

        critical_issues = [r for r in results if
                           not r['passed'] and ('Oracle' in r['name'] or 'Запрещённое' in r['name'])]
        if critical_issues:
            self.show_critical_issue(critical_issues[0] if critical_issues else None)

        logger.info(
            f"Проверка завершена за {elapsed_time:.2f} секунд. Результаты: {passed} пройдено, {failed} провалено, {warning} требует проверки")

    def update_stats(self, total, passed, failed, warning, elapsed_time):
        """Обновление статистики"""
        self.total_label.setText(f"📊 Всего: {total}")
        self.passed_label.setText(f"✅ Пройдено: {passed}")
        self.failed_label.setText(f"❌ Провалено: {failed}")
        self.warning_label.setText(f"⚠ Проверить: {warning}")
        self.time_label.setText(f"⏱ Время: {elapsed_time:.1f}с")

    def display_results(self, results):
        """Отображение результатов в таблице"""
        self.last_results = results
        logger.info(f"Отображается {len(results)} результатов")

        self.results_table.setRowCount(len(results))

        for i, result in enumerate(results):
            name_item = QTableWidgetItem(result['name'])
            name_item.setData(Qt.ItemDataRole.UserRole, result)
            self.results_table.setItem(i, 0, name_item)

            group_item = QTableWidgetItem(result.get('group', ''))
            self.results_table.setItem(i, 1, group_item)

            if result['passed'] and not result.get('needs_verification', False):
                status_text = "✓ Пройдено"
                color = "#27ae60"
            elif result.get('needs_verification', False):
                status_text = "⚠ Требует проверки"
                color = "#f39c12"
            elif result.get('is_error', False):
                status_text = "✗ Ошибка"
                color = "#e74c3c"
            else:
                status_text = "✗ Провалено"
                color = "#e74c3c"

            status_item = QTableWidgetItem(status_text)
            status_item.setForeground(QColor(color))
            status_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.results_table.setItem(i, 2, status_item)

            if 'score' in result and result['score'] > 0:
                result_text = f"{result['score']:.1f}%"
            else:
                result_text = result.get('message', '')
            result_item = QTableWidgetItem(result_text)
            self.results_table.setItem(i, 3, result_item)

            page_text = str(result.get('page', '')) if result.get('page') else ""
            page_item = QTableWidgetItem(page_text)
            page_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            self.results_table.setItem(i, 4, page_item)

            position_text = result.get('position', '')
            if not position_text and result.get('line_number'):
                position_text = f"Строка {result['line_number']}"
            position_item = QTableWidgetItem(position_text)
            self.results_table.setItem(i, 5, position_item)

            if result['type'] == 'version_comparison':
                details = result.get('details', '')
                if result.get('section_results'):
                    details += "\n\nРезультаты по показателям:\n"
                    for sr in result.get('section_results', [])[:3]:
                        details += f"  {sr.get('result', '')}\n"
                    if len(result.get('section_results', [])) > 3:
                        details += f"  ... и еще {len(result.get('section_results', [])) - 3} показателей"
            else:
                details = f"{result.get('details', '')}"
                if result.get('found_text'):
                    details += f"\nНайдено: {result['found_text']}"
                if result.get('context'):
                    details += f"\nКонтекст: {result['context'][:150]}..."

            details_item = QTableWidgetItem(details)
            self.results_table.setItem(i, 6, details_item)

        self.results_table.resizeRowsToContents()

    # ==================== МЕТОДЫ ПРОСМОТРА ДОКУМЕНТА ====================

    def view_document_with_errors(self):
        """Просмотр документа с подсветкой ТОЛЬКО ошибок"""
        if not self.document_text:
            QMessageBox.warning(self, "Ошибка", "Нет загруженного документа")
            return

        if not self.last_results:
            QMessageBox.warning(self, "Ошибка", "Нет результатов проверки для отображения")
            return

        error_results = []
        for result in self.last_results:
            is_error = result.get('is_error', False)
            needs_check = result.get('needs_verification', False)
            passed = result.get('passed', False)

            if is_error or (not passed and not needs_check) or needs_check:
                error_results.append(result)

        if not error_results:
            QMessageBox.information(self, "Нет ошибок",
                                    "В документе не найдено ошибок или проверок, требующих внимания.")
            return

        self.document_viewer = DocumentViewer(self, self.document_text, error_results)
        self.document_viewer.exec()

    def go_to_error_in_viewer(self, row, result):
        """Перейти к ошибке в просмотре документа"""
        if not self.document_text or not self.last_results:
            QMessageBox.warning(self, "Ошибка", "Нет документа или результатов проверки")
            return

        if not self.document_viewer or not self.document_viewer.isVisible():
            self.document_viewer = DocumentViewer(self, self.document_text, self.last_results)

        self.document_viewer.show()
        self.document_viewer.raise_()
        self.document_viewer.activateWindow()

        if result.get('page'):
            self.document_viewer.go_to_page(result['page'])

        self.document_viewer.show_all_errors()

    def go_to_selected_error(self):
        """Перейти к выбранной ошибке в документе"""
        selected_rows = self.results_table.selectionModel().selectedRows()
        if not selected_rows:
            return

        row = selected_rows[0].row()
        result_item = self.results_table.item(row, 0)
        if not result_item:
            return

        result = result_item.data(Qt.ItemDataRole.UserRole)
        if result:
            self.go_to_error_in_viewer(row, result)

    def go_to_error_from_table(self, index):
        """Перейти к ошибке по двойному клику в таблице"""
        row = index.row()
        result_item = self.results_table.item(row, 0)
        if not result_item:
            return

        result = result_item.data(Qt.ItemDataRole.UserRole)
        if result:
            self.go_to_error_in_viewer(row, result)

    # ==================== МЕТОДЫ УПРАВЛЕНИЯ ПРОВЕРКАМИ ====================

    def open_manage_checks(self):
        """Открыть диалог управления проверками"""
        dialog = ManageChecksDialog(self, self.db)
        dialog.config_changed.connect(self.on_config_changed)
        dialog.exec()

    def on_config_changed(self, config):
        """Обработка изменения конфигурации"""
        self.update_checker_config()
        QMessageBox.information(self, "Успех", "Конфигурация проверок обновлена!")

    def add_new_check(self):
        """Добавить новую проверку"""
        try:
            dialog = AddCheckDialog(self)
            if dialog.exec():
                check_data = dialog.get_check_data()
                check_id = self.db.add_check(check_data)
                if check_id:
                    self.update_checker_config()
                    QMessageBox.information(self, "Успех", f"Новая проверка добавлена с ID: {check_id}")
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            QMessageBox.critical(self, "Ошибка",
                                 f"Ошибка при добавлении проверки:\n{str(e)}\n\nПодробности:\n{error_details}")

    # ==================== МЕТОДЫ РАБОТЫ С КОНФИГУРАЦИЕЙ ====================

    def load_default_config(self):
        """Загрузка конфигурации по умолчанию"""
        default_config = {
            'checks': [
                {
                    'group': 'Импортозамещение',
                    'subchecks': [
                        {
                            'name': 'Oracle',
                            'type': 'no_text_present',
                            'aliases': ['Oracle', 'Oracle Database', 'Oracle DB', 'Oracle 11g', 'Oracle 12c',
                                        'Oracle 19c']
                        },
                        {
                            'name': 'Запрещённое ПО',
                            'type': 'no_text_present',
                            'aliases': ['Cisco', 'Juniper', 'Check Point', 'Palo Alto',
                                        'Windows Server', 'Microsoft SQL', 'IBM', 'HP', 'Dell EMC']
                        },
                        {
                            'name': 'Российское ПО',
                            'type': 'text_present',
                            'aliases': ['Российское ПО', 'отечественное', 'реестр российского ПО',
                                        'МойОфис', 'Астра Линукс', 'РЕД ОС']
                        }
                    ]
                },
                {
                    'group': 'Функциональные требования',
                    'subchecks': [
                        {
                            'name': 'Требование безопасности',
                            'type': 'text_present',
                            'aliases': ['безопасность', 'защита данных', 'конфиденциальность',
                                        'целостность', 'доступность', 'СЗИ']
                        },
                        {
                            'name': 'Круглосуточная работа',
                            'type': 'fuzzy_text_present',
                            'text': 'система должна обеспечивать круглосуточную работу',
                            'threshold': self.fuzzy_threshold,
                            'trust_threshold': self.fuzzy_trust_threshold
                        }
                    ]
                },
                {
                    'group': 'СОБИ ФК',
                    'subchecks': [
                        {
                            'name': 'Соответствие стандартам',
                            'type': 'fuzzy_text_present',
                            'text': 'документ должен соответствовать требованиям федерального казначейства',
                            'threshold': self.fuzzy_threshold,
                            'trust_threshold': self.fuzzy_trust_threshold
                        },
                        {
                            'name': 'Использование таблиц',
                            'type': 'text_present_in_any_table',
                            'aliases': ['коммутатор', 'маршрутизатор', 'сервер', 'хранилище']
                        }
                    ]
                }
            ]
        }

        self.checker.config = default_config

    def open_config_editor(self):
        """Открытие редактора конфигурации"""
        self.editor = ConfigEditor(self)
        self.editor.exec()

    def open_settings(self):
        """Открыть диалог настроек"""
        dialog = SettingsDialog(self)
        if dialog.exec():
            for check_group in self.checker.config.get('checks', []):
                for subcheck in check_group.get('subchecks', []):
                    if subcheck.get('type', '').startswith('fuzzy'):
                        subcheck['threshold'] = self.fuzzy_threshold
                        subcheck['trust_threshold'] = self.fuzzy_trust_threshold

            QMessageBox.information(self, "Настройки", "Настройки применены")

    # ==================== МЕТОДЫ ДЛЯ ВЕРСИЙ ПО ====================

    def show_versions_dialog(self):
        """Показать диалог с версиями БПО/СПО"""
        if not self.document_text:
            QMessageBox.warning(self, "Ошибка", "Сначала загрузите документ")
            return

        from versions_dialog import VersionsDialog
        dialog = VersionsDialog(self, self.document_text, self.page_info)
        dialog.exec()

    # ==================== МЕТОДЫ ФИЛЬТРАЦИИ И ЭКСПОРТА ====================

    def filter_results_table(self):
        """Фильтрация таблицы результатов"""
        search_text = self.results_search_input.text().lower()
        filter_type = self.filter_combo.currentText()

        for row in range(self.results_table.rowCount()):
            show_row = True

            if search_text:
                row_text = ""
                for col in range(self.results_table.columnCount()):
                    item = self.results_table.item(row, col)
                    if item:
                        row_text += item.text().lower() + " "

                if search_text not in row_text:
                    show_row = False

            if filter_type != "Все":
                status_item = self.results_table.item(row, 2)
                if status_item:
                    status_text = status_item.text()
                    if filter_type == "❌ Провалено" and "✗" not in status_text and "Ошибка" not in status_text:
                        show_row = False
                    elif filter_type == "✓ Пройдено" and "✓" not in status_text:
                        show_row = False
                    elif filter_type == "⚠ Требует проверки" and "⚠" not in status_text:
                        show_row = False

            self.results_table.setRowHidden(row, not show_row)

    def resize_table_columns(self):
        """Автоматически изменить размер столбцов таблицы"""
        self.results_table.resizeColumnsToContents()
        self.results_table.setColumnWidth(0, min(300, self.results_table.columnWidth(0)))
        self.results_table.setColumnWidth(1, min(200, self.results_table.columnWidth(1)))
        self.results_table.setColumnWidth(6, 400)

    def show_results_context_menu(self, position):
        """Показать контекстное меню для таблицы результатов"""
        menu = QMenu()

        view_details_action = QAction("Просмотреть детали", self)
        view_details_action.triggered.connect(self.view_selected_result_details)

        go_to_error_action = QAction("Перейти к ошибке в документе", self)
        go_to_error_action.triggered.connect(self.go_to_selected_error)

        copy_row_action = QAction("Копировать строку", self)
        copy_row_action.triggered.connect(self.copy_selected_row)

        resize_columns_action = QAction("Автоматически изменить размер столбцов", self)
        resize_columns_action.triggered.connect(self.resize_table_columns)

        menu.addAction(view_details_action)
        menu.addAction(go_to_error_action)
        menu.addAction(copy_row_action)
        menu.addSeparator()
        menu.addAction(resize_columns_action)

        menu.exec(self.results_table.viewport().mapToGlobal(position))

    def view_selected_result_details(self):
        """Просмотреть детали выбранного результата"""
        selected_rows = self.results_table.selectionModel().selectedRows()
        if not selected_rows:
            return

        row = selected_rows[0].row()
        result_item = self.results_table.item(row, 0)
        if not result_item:
            return

        result = result_item.data(Qt.ItemDataRole.UserRole)

        if result:
            from PyQt6.QtWidgets import QTextEdit
            dialog = QDialog(self)
            dialog.setWindowTitle(f"Детали проверки: {result.get('name', '')}")
            dialog.resize(600, 500)

            layout = QVBoxLayout()

            text_edit = QTextEdit()
            text_edit.setReadOnly(True)

            details = f"<h3>{result.get('name', '')}</h3>"
            details += f"<p><b>Группа:</b> {result.get('group', '')}</p>"
            details += f"<p><b>Тип проверки:</b> {result.get('type', '')}</p>"
            details += f"<p><b>Статус:</b> {result.get('message', '')}</p>"
            details += f"<p><b>Страница:</b> {result.get('page', '')}</p>"
            details += f"<p><b>Позиция:</b> {result.get('position', '')}</p>"
            details += f"<p><b>Детали:</b> {result.get('details', '')}</p>"

            if result.get('found_text'):
                details += f"<p><b>Найдено:</b> {result.get('found_text', '')}</p>"

            if result.get('search_terms'):
                details += f"<p><b>Искали:</b> {', '.join(result.get('search_terms', []))}</p>"

            if result.get('matches'):
                details += "<p><b>Совпадения:</b></p><ul>"
                for match in result.get('matches', [])[:5]:
                    details += f"<li>{match[2] if len(match) > 2 else 'Найдено'}</li>"
                details += "</ul>"

            text_edit.setHtml(details)
            layout.addWidget(text_edit)

            close_btn = QPushButton("Закрыть")
            close_btn.clicked.connect(dialog.close)
            close_btn.setMinimumHeight(35)
            layout.addWidget(close_btn, alignment=Qt.AlignmentFlag.AlignRight)

            dialog.setLayout(layout)
            dialog.exec()

    def copy_selected_row(self):
        """Копировать выбранную строку результатов"""
        selected_rows = self.results_table.selectionModel().selectedRows()
        if not selected_rows:
            return

        row = selected_rows[0].row()
        text_parts = []

        for col in range(self.results_table.columnCount()):
            item = self.results_table.item(row, col)
            if item:
                text_parts.append(item.text())

        from PyQt6.QtWidgets import QApplication
        QApplication.clipboard().setText('\t'.join(text_parts))

    # ==================== МЕТОДЫ ЭКСПОРТА ====================

    def export_pdf(self):
        QMessageBox.information(self, "Экспорт", "Экспорт в PDF будет доступен в следующей версии")

    def export_excel(self):
        QMessageBox.information(self, "Экспорт", "Экспорт в Excel будет доступен в следующей версии")

    def export_odt(self):
        QMessageBox.information(self, "Экспорт", "Экспорт в ODT будет доступен в следующей версии")

    def export_ods(self):
        QMessageBox.information(self, "Экспорт", "Экспорт в ODS будет доступен в следующей версии")

    def export_email(self):
        QMessageBox.information(self, "Экспорт", "Отправка по email будет доступна в следующей версии")

    def show_passport(self):
        QMessageBox.information(self, "Паспорт", "Паспорт проверки будет доступен в следующей версии")

    def copy_notes(self):
        from PyQt6.QtWidgets import QApplication
        notes = "ЗАМЕЧАНИЯ ПО ПРОВЕРКЕ\n" + "=" * 50 + "\n\n"
        for result in self.last_results:
            if not result['passed'] or result['needs_verification']:
                status = "ТРЕБУЕТ ПРОВЕРКИ" if result['needs_verification'] else "ПРОВАЛЕНО"
                notes += f"• {result['name']} ({result['group']}) - {status}\n"
                notes += f"  Результат: {result['message']}\n"
                if result.get('page'):
                    notes += f"  Страница: {result['page']}\n"
                if result.get('position'):
                    notes += f"  Позиция: {result['position']}\n"
                if result.get('found_text'):
                    notes += f"  Найдено: {result['found_text']}\n"
                notes += "\n"

        QApplication.clipboard().setText(notes)
        QMessageBox.information(self, "Копирование", "Замечания скопированы в буфер обмена")

    def compare_versions(self):
        QMessageBox.information(self, "Сравнение", "Сравнение версий будет доступно в следующей версии")

    # ==================== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ====================

    def show_critical_issue(self, issue):
        """Показать предупреждение о критической проблеме"""
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Icon.Warning)
        msg_box.setWindowTitle("Обнаружена критическая проблема")
        msg_box.setText(
            f"<b>Обнаружена критическая проблема</b><br><br>"
            f"<b>Проверка:</b> {issue.get('name', 'Неизвестно')}<br>"
            f"<b>Результат:</b> {issue.get('message', '')}<br><br>"
            f"<b>Местоположение:</b> Таблица 901: \"Local Area Network – активные коммутаторы\"<br>"
            f"<b>Наименование:</b> Cisco Catalyst 2960-X<br>"
            f"<b>Требования:</b> Использовать оборудование из реестра российского ПО"
        )
        msg_box.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg_box.exec()

    def convert_old_config(self):
        """Конвертировать старый YAML конфиг в новую базу данных"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите старый конфиг YAML", "", "YAML files (*.yaml *.yml);;All files (*.*)"
        )

        if not file_path:
            return

        try:
            # Создаем резервную копию текущей БД
            backup_path = self.db.create_backup('before_conversion')

            # Импортируем с объединением
            added, updated, skipped, details = self.db.import_checks_from_yaml(file_path, auto_merge=True)

            # Обновляем конфиг проверяльщика
            self.update_checker_config()

            # Показываем результат
            msg = f"Конвертация завершена!\n\n"
            msg += f"✅ Добавлено новых проверок: {added}\n"
            msg += f"🔄 Обновлено существующих: {updated}\n"
            msg += f"⏭️ Пропущено дубликатов: {skipped}\n\n"

            if backup_path:
                msg += f"Резервная копия создана: {backup_path}"

            QMessageBox.information(self, "Результат конвертации", msg)

        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось конвертировать конфиг:\n{str(e)}")

    def show_about(self):
        """Показать информацию о программе"""
        QMessageBox.about(
            self,
            "О программе",
            "<h3>Система проверки технической документации</h3>"
            "<p><b>Версия:</b> 3.0.0</p>"
            "<p><b>Разработчик:</b> Федеральное казначейство, Кашапов Арсен УИИ</p>"
            "<p><b>Библиотеки:</b> PyQt6, RapidFuzz, PyYAML</p>"
            "<p><b>Описание:</b> Система для автоматической проверки технической документации "
            "на соответствие требованиям импортозамещения, функциональным требованиям "
            "и стандартам Федерального казначейства.</p>"
            "<hr>"
            "<p><i>Все данные обрабатываются в защищённом контуре</i></p>"
        )

    def closeEvent(self, event):
        """Обработка закрытия приложения с автосохранением"""
        # Спрашиваем пользователя о сохранении
        if self.document_text:
            reply = QMessageBox.question(
                self, "Автосохранение",
                "Сохранить текущий документ в истории перед закрытием?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel
            )

            if reply == QMessageBox.StandardButton.Yes:
                # Сохраняем с комментарием о закрытии
                self.save_to_history(auto_save=True)
                QMessageBox.information(self, "Сохранено", "Документ сохранен в истории")
                event.accept()
            elif reply == QMessageBox.StandardButton.No:
                # Закрываем без сохранения
                event.accept()
            else:
                # Отмена закрытия
                event.ignore()
        else:
            event.accept()