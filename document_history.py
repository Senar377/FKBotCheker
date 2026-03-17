# document_history.py
"""
Модуль для управления историей версий документов и результатов их проверок
Интегрирован с JSONDatabase для хранения в единой базе данных
"""

import json
import os
import hashlib
import logging
import re
from datetime import datetime
from typing import List, Dict, Any, Optional, Tuple
import difflib

from json_database import JSONDatabase, JSONDatabaseError

logger = logging.getLogger(__name__)


def extract_base_filename(filename: str) -> str:
    """
    Извлекает базовое имя файла без версии
    Например: "document_v1.2.3.docx" -> "document.docx"
    """
    # Паттерны для поиска версии в имени файла
    version_patterns = [
        r'_v\d+[._]\d+[._]\d+',  # _v1.2.3, _v1_2_3
        r'_v\d+[._]\d+',  # _v1.2, _v1_2
        r'_v\d+',  # _v1
        r'\.v\d+[._]\d+[._]\d+',  # .v1.2.3
        r'\.v\d+[._]\d+',  # .v1.2
        r'\.v\d+',  # .v1
        r'-\d+[._]\d+[._]\d+',  # -1.2.3
        r'-\d+[._]\d+',  # -1.2
        r'-\d+',  # -1
        r'\(\d+[._]\d+[._]\d+\)',  # (1.2.3)
        r'\(\d+[._]\d+\)',  # (1.2)
        r'\(\d+\)',  # (1)
    ]

    base_name = filename
    for pattern in version_patterns:
        base_name = re.sub(pattern, '', base_name, flags=re.IGNORECASE)

    return base_name


class DocumentVersion:
    """Класс, представляющий версию документа с результатами проверки"""

    def __init__(self,
                 file_path: str,
                 document_text: str,
                 file_hash: str,
                 version_id: str = None,
                 parent_version_id: str = None,
                 check_results: List[Dict] = None,
                 document_info: Dict = None,
                 metadata: Dict = None):

        self.file_path = file_path
        self.full_filename = os.path.basename(file_path)
        self.base_filename = extract_base_filename(self.full_filename)
        self.document_text = document_text
        self.file_hash = file_hash
        self.version_id = version_id or self._generate_version_id()
        self.parent_version_id = parent_version_id
        self.check_results = check_results or []
        self.document_info = document_info or {}
        self.metadata = metadata or {}
        self.created_at = datetime.now().isoformat()
        self.word_count = len(document_text.split())
        self.char_count = len(document_text)

        # Комментарии к версии
        self.comments = []  # Список комментариев
        self.tags = []  # Теги для версии

        # Информация о ГК (из document_info или извлекаем из текста)
        self.gk_numbers = document_info.get('gk_numbers', [])
        self.gk_date = document_info.get('gk_date')
        self.gk_subsystems = document_info.get('gk_with_subsystems', {})
        self.primary_gk = self._get_primary_gk()

        # Извлекаем версию из имени файла
        self.file_version = self._extract_version_from_filename()

        # Статистика по результатам
        self.update_stats()

    def _generate_version_id(self) -> str:
        """Генерация уникального ID версии"""
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        random_part = hashlib.md5(str(timestamp).encode()).hexdigest()[:8]
        return f"v{timestamp}_{random_part}"

    def _extract_version_from_filename(self) -> Optional[str]:
        """Извлечение версии из имени файла"""
        # Паттерны для поиска версии
        patterns = [
            r'_v([\d._]+)',  # _v1.2.3
            r'\.v([\d._]+)',  # .v1.2.3
            r'-([\d._]+)',  # -1.2.3
            r'\(([\d._]+)\)',  # (1.2.3)
        ]

        for pattern in patterns:
            match = re.search(pattern, self.full_filename, re.IGNORECASE)
            if match:
                version = match.group(1)
                # Заменяем подчеркивания на точки для единообразия
                version = version.replace('_', '.')
                return version

        return None

    def _get_primary_gk(self) -> Optional[Dict]:
        """Получение основного ГК (с первой страницы)"""
        if not self.gk_numbers:
            return None

        # Берем первый ГК из списка (обычно с первой страницы)
        primary_gk = self.gk_numbers[0]
        subsystem = self.gk_subsystems.get(primary_gk, 'Не определена')

        return {
            'number': primary_gk,
            'subsystem': subsystem,
            'date': self.gk_date
        }

    def update_stats(self):
        """Обновление статистики по результатам проверки"""
        self.stats = {
            'total': len(self.check_results),
            'passed': sum(
                1 for r in self.check_results if r.get('passed', False) and not r.get('needs_verification', False)),
            'failed': sum(
                1 for r in self.check_results if not r.get('passed', False) and not r.get('needs_verification', False)),
            'needs_verification': sum(1 for r in self.check_results if r.get('needs_verification', False)),
            'errors': sum(1 for r in self.check_results if r.get('is_error', False))
        }

    def add_comment(self, comment: str, author: str = "user") -> Dict:
        """Добавление комментария к версии"""
        comment_data = {
            'id': len(self.comments) + 1,
            'text': comment,
            'author': author,
            'created_at': datetime.now().isoformat()
        }
        self.comments.append(comment_data)
        return comment_data

    def add_tag(self, tag: str):
        """Добавление тега к версии"""
        if tag and tag not in self.tags:
            self.tags.append(tag)

    def remove_tag(self, tag: str):
        """Удаление тега"""
        if tag in self.tags:
            self.tags.remove(tag)

    def to_dict(self) -> Dict:
        """Преобразование в словарь для сериализации"""
        return {
            'version_id': self.version_id,
            'parent_version_id': self.parent_version_id,
            'file_path': self.file_path,
            'full_filename': self.full_filename,
            'base_filename': self.base_filename,
            'file_version': self.file_version,
            'file_hash': self.file_hash,
            'created_at': self.created_at,
            'document_info': self.document_info,
            'stats': self.stats,
            'word_count': self.word_count,
            'char_count': self.char_count,
            'metadata': self.metadata,
            'has_results': len(self.check_results) > 0,
            'results_count': len(self.check_results),
            'comments': self.comments,
            'tags': self.tags,
            'comment_count': len(self.comments),
            'tag_count': len(self.tags),
            'gk_numbers': self.gk_numbers,
            'gk_date': self.gk_date,
            'gk_subsystems': self.gk_subsystems,
            'primary_gk': self.primary_gk
        }

    def to_dict_full(self) -> Dict:
        """Преобразование в словарь с полными данными (включая результаты)"""
        data = self.to_dict()
        data['check_results'] = self.check_results
        # Не сохраняем полный текст в метаданные, только ссылку
        data['has_full_text'] = True
        return data


class DocumentGroup:
    """Класс для группировки версий одного документа"""

    def __init__(self, base_filename: str):
        self.base_filename = base_filename
        self.versions = []  # Список ID версий
        self.created_at = datetime.now().isoformat()
        self.last_accessed = datetime.now().isoformat()
        self.tags = []  # Общие теги для всех версий

        # Текущая информация (из последней версии)
        self.current_file_path = None
        self.current_gk_numbers = []
        self.current_gk_date = None
        self.current_gk_subsystems = {}
        self.current_primary_gk = None

    def add_version(self, version_id: str, version_info: Dict):
        """Добавление версии в группу"""
        if version_id not in self.versions:
            self.versions.append(version_id)
            self.versions.sort(key=lambda v: version_info.get('created_at', ''), reverse=True)

        # Обновляем информацию из последней версии
        if self.versions:
            latest_id = self.versions[0]
            if latest_id == version_id:
                self.current_file_path = version_info.get('file_path')
                self.current_gk_numbers = version_info.get('gk_numbers', [])
                self.current_gk_date = version_info.get('gk_date')
                self.current_gk_subsystems = version_info.get('gk_subsystems', {})
                self.current_primary_gk = version_info.get('primary_gk')

        self.last_accessed = datetime.now().isoformat()

    def remove_version(self, version_id: str):
        """Удаление версии из группы"""
        if version_id in self.versions:
            self.versions.remove(version_id)

    def to_dict(self) -> Dict:
        """Преобразование в словарь для сериализации"""
        return {
            'base_filename': self.base_filename,
            'versions': self.versions,
            'created_at': self.created_at,
            'last_accessed': self.last_accessed,
            'tags': self.tags,
            'version_count': len(self.versions),
            'current_file_path': self.current_file_path,
            'current_gk_numbers': self.current_gk_numbers,
            'current_gk_date': self.current_gk_date,
            'current_gk_subsystems': self.current_gk_subsystems,
            'current_primary_gk': self.current_primary_gk
        }


class DocumentHistory:
    """
    Класс для управления историей версий документов
    Интегрирован с JSONDatabase для хранения данных
    """

    # Константы для ключей в JSON базе данных
    HISTORY_DOC_KEY = "document_history"  # Ключ для хранения истории в JSON базе

    def __init__(self, db: JSONDatabase = None, base_dir: str = None):
        """
        Инициализация хранилища истории документов

        Args:
            db: экземпляр JSONDatabase (если None, создается новый)
            base_dir: базовая директория для хранения файлов (если db не указан)
        """
        if db is not None:
            self.db = db
            self.base_dir = db.base_dir
            self.storage_type = "json_database"
        else:
            # Создаем новую JSON базу данных в указанной директории
            from json_database import JSONDatabase
            self.db = JSONDatabase(base_dir or 'document_history')
            self.base_dir = self.db.base_dir
            self.storage_type = "separate"

        # Директории для хранения текстов документов
        self.texts_dir = os.path.join(self.base_dir, 'document_texts')
        os.makedirs(self.texts_dir, exist_ok=True)

        # Загружаем или создаем индекс истории
        self.history_data = self._load_history_data()

        logger.info(f"DocumentHistory инициализирована")
        logger.info(f"Тип хранилища: {self.storage_type}")
        logger.info(f"Загружено {len(self.history_data.get('groups', {}))} групп документов")

    def _load_history_data(self) -> Dict:
        """Загрузка данных истории из JSON базы"""
        try:
            # Пытаемся загрузить историю из JSONDatabase
            # В JSONDatabase нет прямого метода для этого, поэтому используем отдельный файл
            # или расширяем схему данных JSONDatabase
            history_file = os.path.join(self.base_dir, 'document_history.json')

            if os.path.exists(history_file):
                with open(history_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if 'groups' not in data:
                        data['groups'] = {}
                    if 'versions' not in data:
                        data['versions'] = {}
                    return data
            else:
                return {'groups': {}, 'versions': {}}

        except Exception as e:
            logger.error(f"Ошибка загрузки истории: {e}")
            return {'groups': {}, 'versions': {}}

    def _save_history_data(self):
        """Сохранение данных истории в JSON базу"""
        try:
            history_file = os.path.join(self.base_dir, 'document_history.json')

            # Создаем резервную копию перед сохранением (опционально)
            if os.path.exists(history_file):
                backup_file = history_file + '.backup'
                try:
                    import shutil
                    shutil.copy2(history_file, backup_file)
                except:
                    pass

            with open(history_file, 'w', encoding='utf-8') as f:
                json.dump(self.history_data, f, ensure_ascii=False, indent=2, default=str)

            logger.info(f"История сохранена: {len(self.history_data.get('versions', {}))} версий")

            # Если используем JSONDatabase, создаем резервную копию
            if self.storage_type == "json_database":
                try:
                    self.db.create_backup('history_update')
                except Exception as e:
                    logger.warning(f"Не удалось создать бэкап в JSONDatabase: {e}")

        except Exception as e:
            logger.error(f"Ошибка сохранения истории: {e}")

    def _compute_file_hash(self, file_path: str) -> str:
        """Вычисление хеша файла для идентификации"""
        try:
            if os.path.exists(file_path):
                with open(file_path, 'rb') as f:
                    return hashlib.md5(f.read()).hexdigest()
        except Exception as e:
            logger.error(f"Ошибка вычисления хеша файла: {e}")
        return ""

    def _compute_text_hash(self, text: str) -> str:
        """Вычисление хеша текста документа"""
        return hashlib.md5(text.encode('utf-8')).hexdigest()

    def _get_group_key(self, base_filename: str) -> str:
        """Получение ключа группы на основе базового имени файла"""
        # Нормализуем имя: убираем спецсимволы, приводим к нижнему регистру
        normalized = re.sub(r'[^\w\s-]', '', base_filename.lower())
        normalized = re.sub(r'[-\s]+', '_', normalized)
        return f"group_{normalized}"

    def _save_document_text(self, version_id: str, text: str):
        """Сохранение текста документа в отдельный файл"""
        try:
            text_file = os.path.join(self.texts_dir, f"{version_id}.txt")
            with open(text_file, 'w', encoding='utf-8') as f:
                f.write(text)
            return True
        except Exception as e:
            logger.error(f"Ошибка сохранения текста документа {version_id}: {e}")
            return False

    def _load_document_text(self, version_id: str) -> Optional[str]:
        """Загрузка текста документа из файла"""
        try:
            text_file = os.path.join(self.texts_dir, f"{version_id}.txt")
            if os.path.exists(text_file):
                with open(text_file, 'r', encoding='utf-8') as f:
                    return f.read()
        except Exception as e:
            logger.error(f"Ошибка загрузки текста документа {version_id}: {e}")
        return None

    def add_version(self,
                    file_path: str,
                    document_text: str,
                    check_results: List[Dict] = None,
                    document_info: Dict = None,
                    metadata: Dict = None,
                    auto_save: bool = True) -> DocumentVersion:
        """
        Добавление новой версии документа в историю

        Args:
            file_path: путь к файлу документа
            document_text: текст документа
            check_results: результаты проверки
            document_info: информация о документе
            metadata: дополнительные метаданные
            auto_save: автоматически сохранять при закрытии

        Returns:
            DocumentVersion: созданная версия
        """
        # Вычисляем хеши
        file_hash = self._compute_file_hash(file_path) if os.path.exists(file_path) else ""
        text_hash = self._compute_text_hash(document_text)

        # Создаем версию
        version = DocumentVersion(
            file_path=file_path,
            document_text=document_text,
            file_hash=file_hash,
            check_results=check_results or [],
            document_info=document_info or {},
            metadata=metadata or {}
        )

        # Получаем ключ группы
        group_key = self._get_group_key(version.base_filename)

        # Инициализация: проверяем наличие ключей в данных истории
        if 'groups' not in self.history_data:
            self.history_data['groups'] = {}
        if 'versions' not in self.history_data:
            self.history_data['versions'] = {}

        # Ищем родительскую версию (последнюю в группе)
        parent_version_id = None
        if group_key in self.history_data['groups']:
            group_versions = self.history_data['groups'][group_key].get('versions', [])
            if group_versions:
                parent_version_id = group_versions[0]  # Последняя версия

        version.parent_version_id = parent_version_id

        # Добавляем метаданные
        version.metadata.update({
            'group_key': group_key,
            'text_hash': text_hash,
            'is_duplicate': parent_version_id is not None and
                            self.history_data['versions'].get(parent_version_id, {}).get('file_hash') == file_hash,
            'auto_save': auto_save,
            'text_file': f"{version.version_id}.txt"
        })

        # Сохраняем текст документа
        self._save_document_text(version.version_id, document_text)

        # Сохраняем check_results в version_info
        version_dict = version.to_dict()
        version_dict['check_results'] = check_results or []  # Явно сохраняем результаты

        # Сохраняем метаданные версии
        self.history_data['versions'][version.version_id] = version_dict

        # Обновляем или создаем группу
        if group_key not in self.history_data['groups']:
            self.history_data['groups'][group_key] = DocumentGroup(version.base_filename).to_dict()

        # Добавляем версию в группу
        group = DocumentGroup(version.base_filename)
        group.__dict__.update(self.history_data['groups'][group_key])
        group.add_version(version.version_id, version_dict)
        self.history_data['groups'][group_key] = group.to_dict()

        # Сохраняем данные истории
        self._save_history_data()

        # Если это дубликат (та же версия), помечаем
        if version.metadata.get('is_duplicate'):
            logger.info(f"Добавлен дубликат версии {version.version_id} для группы {group_key}")
        else:
            logger.info(f"Добавлена новая версия {version.version_id} для группы {group_key}")
            logger.info(f"Количество версий в группе: {len(self.history_data['groups'][group_key]['versions'])}")

        return version

    def get_version(self, version_id: str, load_full: bool = True) -> Optional[DocumentVersion]:
        """Получение версии по ID"""
        try:
            # Получаем метаданные версии
            version_info = self.history_data['versions'].get(version_id)
            if not version_info:
                logger.error(f"Версия {version_id} не найдена")
                return None

            # Загружаем текст документа
            document_text = self._load_document_text(version_id) if load_full else ""

            # Загружаем результаты проверки
            check_results = version_info.get('check_results', [])
            if not check_results:
                logger.warning(f"Версия {version_id} не содержит результатов проверки")

            # Создаем объект версии
            version = DocumentVersion(
                file_path=version_info.get('file_path', ''),
                document_text=document_text,
                file_hash=version_info.get('file_hash', ''),
                version_id=version_id,
                parent_version_id=version_info.get('parent_version_id'),
                check_results=check_results,
                document_info=version_info.get('document_info', {}),
                metadata=version_info.get('metadata', {})
            )
            version.created_at = version_info.get('created_at', version.created_at)
            version.comments = version_info.get('comments', [])
            version.tags = version_info.get('tags', [])
            version.stats = version_info.get('stats', version.stats)

            logger.info(f"Загружена версия {version_id} с {len(check_results)} результатами проверки")
            return version

        except Exception as e:
            logger.error(f"Ошибка загрузки версии {version_id}: {e}")
            return None

    def debug_info(self):
        """Вывод отладочной информации о состоянии истории"""
        logger.info("=" * 50)
        logger.info("ОТЛАДОЧНАЯ ИНФОРМАЦИЯ ИСТОРИИ")
        logger.info(f"Всего групп: {len(self.history_data.get('groups', {}))}")
        logger.info(f"Всего версий: {len(self.history_data.get('versions', {}))}")

        for group_key, group_info in self.history_data.get('groups', {}).items():
            logger.info(f"\nГруппа: {group_key}")
            logger.info(f"  Base filename: {group_info.get('base_filename')}")
            logger.info(f"  Версии: {group_info.get('versions', [])}")
            logger.info(f"  Количество версий: {group_info.get('version_count', 0)}")

            for version_id in group_info.get('versions', [])[:3]:  # Первые 3 версии
                version_info = self.history_data['versions'].get(version_id, {})
                logger.info(f"    Версия {version_id}:")
                logger.info(f"      Создана: {version_info.get('created_at')}")
                logger.info(f"      Результатов: {len(version_info.get('check_results', []))}")
                logger.info(f"      Статистика: {version_info.get('stats', {})}")

    def get_group(self, group_key: str) -> Optional[Dict]:
        """Получение информации о группе по ключу"""
        return self.history_data['groups'].get(group_key)

    def get_all_groups(self) -> List[Dict]:
        """Получение списка всех групп документов"""
        groups = []
        for group_key, group_info in self.history_data['groups'].items():
            group_copy = group_info.copy()
            group_copy['group_key'] = group_key
            groups.append(group_copy)

        # Сортируем по дате последнего доступа
        groups.sort(key=lambda x: x.get('last_accessed', ''), reverse=True)
        return groups

    def get_group_versions(self, group_key: str) -> List[Dict]:
        """Получение списка версий для группы"""
        group_info = self.history_data['groups'].get(group_key)
        if not group_info:
            return []

        versions = []
        for version_id in group_info.get('versions', []):
            version_info = self.history_data['versions'].get(version_id, {})
            if version_info:
                versions.append({
                    'version_id': version_id,
                    'created_at': version_info.get('created_at', ''),
                    'full_filename': version_info.get('full_filename', ''),
                    'file_version': version_info.get('file_version'),
                    'stats': version_info.get('stats', {}),
                    'comment_count': version_info.get('comment_count', 0),
                    'tag_count': version_info.get('tag_count', 0),
                    'tags': version_info.get('tags', []),
                    'gk_numbers': version_info.get('gk_numbers', []),
                    'gk_date': version_info.get('gk_date'),
                    'primary_gk': version_info.get('primary_gk'),
                    'is_current': version_id == group_info.get('versions', [])[0] if group_info.get(
                        'versions') else False
                })

        # Сортируем по дате (новые сверху)
        versions.sort(key=lambda x: x.get('created_at', ''), reverse=True)
        return versions

    def get_versions_for_file(self, file_path: str) -> List[Dict]:
        """Получение списка версий для документа по пути файла"""
        # Создаем временную версию для получения base_filename
        temp_version = DocumentVersion(
            file_path=file_path,
            document_text="",
            file_hash=""
        )
        group_key = self._get_group_key(temp_version.base_filename)

        return self.get_group_versions(group_key)

    def add_comment_to_version(self, version_id: str, comment: str, author: str = "user") -> bool:
        """Добавление комментария к версии"""
        version_info = self.history_data['versions'].get(version_id)
        if not version_info:
            return False

        # Добавляем комментарий
        comments = version_info.get('comments', [])
        comment_data = {
            'id': len(comments) + 1,
            'text': comment,
            'author': author,
            'created_at': datetime.now().isoformat()
        }
        comments.append(comment_data)
        version_info['comments'] = comments
        version_info['comment_count'] = len(comments)

        # Сохраняем
        self._save_history_data()
        logger.info(f"Добавлен комментарий к версии {version_id}")
        return True

    def add_tag_to_version(self, version_id: str, tag: str) -> bool:
        """Добавление тега к версии"""
        version_info = self.history_data['versions'].get(version_id)
        if not version_info:
            return False

        # Добавляем тег
        tags = version_info.get('tags', [])
        if tag not in tags:
            tags.append(tag)
            version_info['tags'] = tags
            version_info['tag_count'] = len(tags)

            # Обновляем теги в группе
            group_key = version_info.get('metadata', {}).get('group_key')
            if group_key and group_key in self.history_data['groups']:
                group_tags = set(self.history_data['groups'][group_key].get('tags', []))
                group_tags.add(tag)
                self.history_data['groups'][group_key]['tags'] = list(group_tags)

            # Сохраняем
            self._save_history_data()
            logger.info(f"Добавлен тег '{tag}' к версии {version_id}")
            return True

        return False

    def remove_tag_from_version(self, version_id: str, tag: str) -> bool:
        """Удаление тега из версии"""
        version_info = self.history_data['versions'].get(version_id)
        if not version_info:
            return False

        # Удаляем тег
        tags = version_info.get('tags', [])
        if tag in tags:
            tags.remove(tag)
            version_info['tags'] = tags
            version_info['tag_count'] = len(tags)

            # Сохраняем
            self._save_history_data()
            logger.info(f"Удален тег '{tag}' из версии {version_id}")
            return True

        return False

    def get_comments_for_version(self, version_id: str) -> List[Dict]:
        """Получение комментариев для версии"""
        version_info = self.history_data['versions'].get(version_id, {})
        return version_info.get('comments', [])

    def compare_versions(self, version_id1: str, version_id2: str) -> Dict:
        """
        Сравнение двух версий документа

        Returns:
            Dict с информацией о различиях:
            - text_diff: различия в тексте
            - results_diff: различия в результатах проверки
            - stats_diff: различия в статистике
            - new_errors: новые ошибки
            - fixed_errors: исправленные ошибки
            - unchanged_errors: оставшиеся ошибки
        """
        version1 = self.get_version(version_id1)
        version2 = self.get_version(version_id2)

        if not version1 or not version2:
            return {'error': 'Версии не найдены'}

        result = {
            'version1': version1.to_dict(),
            'version2': version2.to_dict(),
            'text_diff': self._compare_texts(version1.document_text, version2.document_text),
            'results_diff': self._compare_results(version1.check_results, version2.check_results),
            'stats_diff': self._compare_stats(version1.stats, version2.stats),
            'new_errors': [],
            'fixed_errors': [],
            'unchanged_errors': []
        }

        # Анализируем изменения в ошибках
        errors1 = {r['name']: r for r in version1.check_results
                   if not r.get('passed', False) or r.get('needs_verification', False)}
        errors2 = {r['name']: r for r in version2.check_results
                   if not r.get('passed', False) or r.get('needs_verification', False)}

        # Новые ошибки
        for name, error in errors2.items():
            if name not in errors1:
                result['new_errors'].append(error)

        # Исправленные ошибки
        for name, error in errors1.items():
            if name not in errors2:
                result['fixed_errors'].append(error)
            elif error.get('passed', False) != errors2[name].get('passed', False):
                result['fixed_errors'].append(error)

        # Неизменные ошибки
        for name, error in errors1.items():
            if name in errors2 and error.get('passed', False) == errors2[name].get('passed', False):
                result['unchanged_errors'].append(error)

        return result

    def _compare_texts(self, text1: str, text2: str) -> Dict:
        """Сравнение текстов"""
        lines1 = text1.splitlines()
        lines2 = text2.splitlines()

        diff = list(difflib.unified_diff(lines1, lines2, lineterm=''))

        return {
            'has_changes': len(diff) > 0,
            'diff_lines': diff[:100],  # Ограничиваем для отображения
            'total_diff_lines': len(diff),
            'added_lines': sum(1 for line in diff if line.startswith('+') and not line.startswith('+++')),
            'removed_lines': sum(1 for line in diff if line.startswith('-') and not line.startswith('---'))
        }

    def _compare_results(self, results1: List[Dict], results2: List[Dict]) -> Dict:
        """Сравнение результатов проверки"""
        # Создаем словари для быстрого доступа
        dict1 = {r['name']: r for r in results1}
        dict2 = {r['name']: r for r in results2}

        changed = []
        added = []
        removed = []

        for name, result in dict2.items():
            if name not in dict1:
                added.append(result)
            elif result.get('passed', False) != dict1[name].get('passed', False):
                changed.append({
                    'name': name,
                    'old_status': '✅ Пройдено' if dict1[name].get('passed') else '❌ Провалено',
                    'new_status': '✅ Пройдено' if result.get('passed') else '❌ Провалено',
                    'old_result': dict1[name].get('message', ''),
                    'new_result': result.get('message', '')
                })

        for name, result in dict1.items():
            if name not in dict2:
                removed.append(result)

        return {
            'total': len(results2),
            'total_previous': len(results1),
            'added': added,
            'removed': removed,
            'changed': changed,
            'added_count': len(added),
            'removed_count': len(removed),
            'changed_count': len(changed)
        }

    def _compare_stats(self, stats1: Dict, stats2: Dict) -> Dict:
        """Сравнение статистики"""
        return {
            'passed': {
                'old': stats1.get('passed', 0),
                'new': stats2.get('passed', 0),
                'diff': stats2.get('passed', 0) - stats1.get('passed', 0)
            },
            'failed': {
                'old': stats1.get('failed', 0),
                'new': stats2.get('failed', 0),
                'diff': stats2.get('failed', 0) - stats1.get('failed', 0)
            },
            'needs_verification': {
                'old': stats1.get('needs_verification', 0),
                'new': stats2.get('needs_verification', 0),
                'diff': stats2.get('needs_verification', 0) - stats1.get('needs_verification', 0)
            }
        }

    def get_group_timeline(self, group_key: str) -> List[Dict]:
        """Получение временной линии изменений группы"""
        versions = self.get_group_versions(group_key)

        timeline = []
        prev_version = None

        for version in versions:
            version_data = {
                'version': version,
                'changes': None
            }

            if prev_version:
                # Сравниваем с предыдущей версией
                comparison = self.compare_versions(prev_version['version_id'], version['version_id'])
                version_data['changes'] = {
                    'fixed_errors': len(comparison.get('fixed_errors', [])),
                    'new_errors': len(comparison.get('new_errors', [])),
                    'text_changes': comparison.get('text_diff', {}).get('total_diff_lines', 0)
                }

            timeline.append(version_data)
            prev_version = version

        return timeline

    def cleanup_old_versions(self, keep_last: int = 10):
        """Очистка старых версий, оставляя только последние N для каждой группы"""
        for group_key, group_info in self.history_data['groups'].items():
            versions = group_info.get('versions', [])
            if len(versions) > keep_last:
                versions_to_remove = versions[keep_last:]  # Удаляем старые (в конце списка)
                for version_id in versions_to_remove:
                    self._delete_version(version_id)
                group_info['versions'] = versions[:keep_last]
                group_info['version_count'] = len(group_info['versions'])

        self._save_history_data()

    def _delete_version(self, version_id: str):
        """Удаление версии"""
        try:
            # Получаем информацию о версии перед удалением
            version_info = self.history_data['versions'].get(version_id, {})
            group_key = version_info.get('metadata', {}).get('group_key')

            # Удаляем файл с текстом документа
            text_file = os.path.join(self.texts_dir, f"{version_id}.txt")
            if os.path.exists(text_file):
                os.remove(text_file)

            # Удаляем из индекса версий
            if version_id in self.history_data['versions']:
                del self.history_data['versions'][version_id]

            # Удаляем из группы
            if group_key and group_key in self.history_data['groups']:
                if version_id in self.history_data['groups'][group_key]['versions']:
                    self.history_data['groups'][group_key]['versions'].remove(version_id)
                    self.history_data['groups'][group_key]['version_count'] = len(
                        self.history_data['groups'][group_key]['versions'])

            logger.info(f"Удалена версия {version_id}")

        except Exception as e:
            logger.error(f"Ошибка удаления версии {version_id}: {e}")

    def search_groups_by_tag(self, tag: str) -> List[Dict]:
        """Поиск групп по тегу"""
        results = []
        for group_key, group_info in self.history_data['groups'].items():
            if tag in group_info.get('tags', []):
                results.append({
                    'group_key': group_key,
                    'group_info': group_info
                })
        return results

    def search_groups_by_gk(self, gk_number: str) -> List[Dict]:
        """Поиск групп по номеру ГК"""
        results = []
        gk_upper = gk_number.upper()

        for group_key, group_info in self.history_data['groups'].items():
            primary_gk = group_info.get('current_primary_gk', {})
            if primary_gk and gk_upper in primary_gk.get('number', '').upper():
                results.append({
                    'group_key': group_key,
                    'group_info': group_info
                })

        return results

    def get_all_tags(self) -> List[str]:
        """Получение списка всех используемых тегов"""
        tags = set()
        for group_info in self.history_data['groups'].values():
            tags.update(group_info.get('tags', []))
        return sorted(list(tags))

    def get_stats(self) -> Dict:
        """Получение статистики по истории"""
        stats = {
            'total_groups': len(self.history_data.get('groups', {})),
            'total_versions': len(self.history_data.get('versions', {})),
            'total_comments': 0,
            'total_tags': 0,
            'groups': []
        }

        for group_key, group_info in self.history_data.get('groups', {}).items():
            group_stats = {
                'group_key': group_key,
                'base_filename': group_info.get('base_filename'),
                'created_at': group_info.get('created_at'),
                'versions_count': len(group_info.get('versions', [])),
                'last_accessed': group_info.get('last_accessed'),
                'primary_gk': group_info.get('current_primary_gk'),
                'tags': group_info.get('tags', [])
            }
            stats['groups'].append(group_stats)

        for version_info in self.history_data.get('versions', {}).values():
            stats['total_comments'] += version_info.get('comment_count', 0)
            stats['total_tags'] += version_info.get('tag_count', 0)

        return stats