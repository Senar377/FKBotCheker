# document_checker.py
import yaml
import re
import logging
from typing import List, Dict, Any, Optional, Tuple
from rapidfuzz import fuzz, utils

logger = logging.getLogger(__name__)


class DocumentChecker:
    """Класс для проверки документов по конфигурации"""

    def __init__(self, config_path: str = None):
        self.config = self.load_config(config_path) if config_path else {}
        self.results = []
        self._normalized_cache = {}

    def load_config(self, config_path: str) -> Dict:
        """Загрузка конфигурации из YAML файла"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f) or {}
        except Exception as e:
            logger.error(f"Ошибка загрузки конфигурации: {e}")
            return {}

    def save_config(self, config_path: str, config: Dict) -> bool:
        """Сохранение конфигурации в YAML файл"""
        try:
            with open(config_path, 'w', encoding='utf-8') as f:
                yaml.dump(config, f, allow_unicode=True, default_flow_style=False)
            return True
        except Exception as e:
            logger.error(f"Ошибка сохранения конфигурации: {e}")
            return False

    def normalize_text(self, text: str) -> str:
        """Нормализация текста (регистр, пробелы) с кешированием"""
        if not text:
            return ""

        cache_key = hash(text)
        if cache_key in self._normalized_cache:
            return self._normalized_cache[cache_key]

        text = re.sub(r'\s+', ' ', text.strip().lower())
        self._normalized_cache[cache_key] = text
        return text

    def find_position_in_document(self, start_pos: int, document_text: str, page_info: List[Tuple[int, int]]) -> Dict[
        str, Any]:
        """Точное определение позиции в документе с учетом страниц"""
        result = {
            'page': 0,
            'line_number': 0,
            'position': '',
            'global_start': start_pos,
            'global_end': -1,
            'context': ''
        }

        if start_pos < 0 or start_pos >= len(document_text):
            return result

        # Определяем строку в документе
        lines_before = document_text[:start_pos].count('\n')
        result['line_number'] = lines_before + 1

        # Находим строку, на которой расположена ошибка
        lines = document_text.split('\n')
        if 0 <= lines_before < len(lines):
            current_line = lines[lines_before]
            result['position'] = f"Строка {lines_before + 1}"

            # Берем контекст из текущей строки (первые 100 символов)
            if current_line:
                context = current_line[:100]
                if len(current_line) > 100:
                    context += "..."
                result['context'] = context

        # Определяем страницу
        if page_info:
            for page_num, (page_start, page_end) in enumerate(page_info, 1):
                if page_start <= start_pos < page_end:
                    result['page'] = page_num
                    # Вычисляем строку внутри страницы
                    lines_in_page = document_text[page_start:start_pos].count('\n')
                    result['position'] = f"Страница {page_num}, строка {lines_in_page + 1}"

                    # Добавляем контекст из строки
                    line_start = max(page_start, start_pos - 100)
                    line_end = min(page_end, start_pos + 100)
                    context = document_text[line_start:line_end]
                    if context:
                        result['context'] = f"...{context}..."
                    break

        return result

    def parse_version(self, version_str: str) -> List[int]:
        """
        Парсит строку версии в список чисел.
        Пример: "10.0.1" -> [10, 0, 1]
        """
        try:
            # Извлекаем все числа из строки версии
            parts = []
            for part in re.findall(r'\d+', version_str):
                parts.append(int(part))

            # Если не нашли чисел, возвращаем [0]
            if not parts:
                logger.warning(f"Не удалось извлечь версию из строки: '{version_str}'")
                return [0]

            return parts
        except Exception as e:
            logger.error(f"Ошибка парсинга версии '{version_str}': {e}")
            return [0]

    def compare_versions_with_operator(self, found_version: str, required_version: str, operator: str) -> bool:
        """
        Сравнивает две версии с использованием указанного оператора

        Операторы:
        - 'eq', '=': равно
        - 'ne', '!=': не равно
        - 'gt', '>': больше
        - 'lt', '<': меньше
        - 'ge', '>=': больше или равно
        - 'le', '<=': меньше или равно

        Возвращает True если условие выполняется
        """
        try:
            found_parts = self.parse_version(found_version)
            required_parts = self.parse_version(required_version)

            # Дополняем нулями до одинаковой длины
            max_len = max(len(found_parts), len(required_parts))
            found_parts += [0] * (max_len - len(found_parts))
            required_parts += [0] * (max_len - len(required_parts))

            # Сравниваем по частям
            comparison_result = 0
            for found, required in zip(found_parts, required_parts):
                if found < required:
                    comparison_result = -1
                    break
                elif found > required:
                    comparison_result = 1
                    break

            # Применяем оператор
            operator = operator.lower() if operator else 'ge'  # По умолчанию '>='

            if operator in ['eq', '=', '==']:
                return comparison_result == 0
            elif operator in ['ne', '!=', '<>']:
                return comparison_result != 0
            elif operator in ['gt', '>']:
                return comparison_result == 1
            elif operator in ['lt', '<']:
                return comparison_result == -1
            elif operator in ['ge', '>=', '>=']:
                return comparison_result >= 0
            elif operator in ['le', '<=', '<=']:
                return comparison_result <= 0
            else:
                # Неизвестный оператор, по умолчанию '>='
                logger.warning(f"Неизвестный оператор '{operator}', используется '>='")
                return comparison_result >= 0

        except Exception as e:
            logger.error(
                f"Ошибка сравнения версий '{found_version}' и '{required_version}' с оператором '{operator}': {e}")
            return False

    def search_versions_with_regex(self, text: str, pattern: str) -> List[Dict[str, Any]]:
        """
        Ищет версии в тексте по регулярному выражению.
        Возвращает список найденных версий с информацией.
        """
        versions = []
        try:
            # Компилируем регулярное выражение
            regex = re.compile(pattern, re.IGNORECASE)

            # Ищем все совпадения
            for match in regex.finditer(text):
                if match.groups():
                    # Берем первую группу захвата как версию
                    found_version = match.group(1)

                    # Очищаем версию от лишних символов
                    found_version = re.sub(r'[^\d\.]', '', found_version)

                    if found_version:
                        versions.append({
                            'version': found_version,
                            'position': match.start(),
                            'end_position': match.end(),
                            'full_match': match.group(0),
                            'pattern': pattern
                        })
                        logger.debug(f"Найдена версия '{found_version}' по паттерну '{pattern}'")

        except re.error as e:
            logger.error(f"Ошибка в регулярном выражении '{pattern}': {e}")
        except Exception as e:
            logger.error(f"Ошибка при поиске версий по паттерну '{pattern}': {e}")

        return versions

    def exact_search_with_context(self, text: str, search_terms: List[str]) -> List[Tuple[int, int, str, str]]:
        """Точный поиск с возвратом контекста"""
        matches = []
        text_lower = text.lower()

        for term in search_terms:
            if not term or len(term) < 2:
                continue

            term_lower = term.lower()
            start = 0

            while True:
                pos = text_lower.find(term_lower, start)
                if pos == -1:
                    break

                # Проверяем, что это целое слово (не часть другого слова)
                is_word_boundary = True
                if pos > 0:
                    prev_char = text[pos - 1]
                    is_word_boundary = not (prev_char.isalnum() or prev_char in '_-')

                if pos + len(term) < len(text):
                    next_char = text[pos + len(term)]
                    is_word_boundary = is_word_boundary and not (next_char.isalnum() or next_char in '_-')

                if is_word_boundary:
                    # Берем контекст вокруг найденного слова
                    context_start = max(0, pos - 50)
                    context_end = min(len(text), pos + len(term) + 50)
                    context = text[context_start:context_end]

                    matches.append((pos, pos + len(term), term, context))

                start = pos + 1

        return matches

    def exact_search(self, text: str, search_terms: List[str]) -> List[Tuple[int, int, str]]:
        """Оптимизированный точный поиск целых слов в тексте"""
        matches = []
        text_lower = text.lower()

        for term in search_terms:
            if not term or len(term) < 2:
                continue

            term_lower = term.lower()
            start = 0

            while True:
                pos = text_lower.find(term_lower, start)
                if pos == -1:
                    break

                # Проверяем, что это целое слово
                if (pos == 0 or not text[pos - 1].isalnum()) and \
                        (pos + len(term) == len(text) or not text[pos + len(term)].isalnum()):
                    matches.append((pos, pos + len(term), term))

                start = pos + 1

        return matches

    def fuzzy_search_all(self, text: str, search_text: str, threshold: float = 70.0) -> List[
        Tuple[int, int, float, str, float]]:
        """
        Оптимизированный нечеткий поиск всех вхождений.
        Возвращает список кортежей (start, end, score, found_text, match_quality)
        где match_quality - это отдельный процент схожести для каждого найденного фрагмента
        """
        if not text or not search_text:
            return []

        # Оптимизация: кеширование нормализации
        normalized_text = self.normalize_text(text)
        normalized_search = self.normalize_text(search_text)

        # Если поисковый текст очень короткий - используем упрощенный поиск
        if len(normalized_search) < 10:
            matches = []
            text_words = normalized_text.split()

            for i, word in enumerate(text_words):
                if len(word) < 3:
                    continue

                # Вычисляем точную схожесть для каждого слова
                score = fuzz.ratio(word, normalized_search, processor=utils.default_process)
                if score >= threshold:
                    # Находим позицию в оригинальном тексте
                    words_before = ' '.join(text_words[:i])
                    pos = len(words_before) + (1 if words_before else 0)
                    end_pos = pos + len(word)

                    # Находим оригинальный текст с контекстом
                    context_start = max(0, pos - 50)
                    context_end = min(len(text), end_pos + 50)
                    found_text = text[context_start:context_end]

                    matches.append((pos, end_pos, score, found_text, score))

            return matches[:20]

        # Для длинных текстов используем скользящее окно с оптимизацией
        matches = []
        window_size = min(200, len(normalized_search) * 2)
        step_size = max(10, window_size // 4)

        for i in range(0, len(normalized_text) - window_size + 1, step_size):
            window = normalized_text[i:i + window_size]
            window_score = fuzz.partial_ratio(window, normalized_search)

            if window_score >= threshold:
                # Находим точную позицию в окне
                if i > 0:
                    extended_start = max(0, i - 50)
                    extended_end = min(len(normalized_text), i + window_size + 50)
                    extended_window = normalized_text[extended_start:extended_end]

                    # Ищем лучшую позицию в расширенном окне
                    best_score = 0
                    best_pos = 0
                    best_fragment = ""

                    # Оптимизированный шаг для точного поиска
                    search_step = max(1, len(normalized_search) // 2)

                    for j in range(0, len(extended_window) - len(normalized_search) + 1, search_step):
                        fragment = extended_window[j:j + len(normalized_search)]
                        fragment_score = fuzz.ratio(fragment, normalized_search, processor=utils.default_process)

                        if fragment_score > best_score:
                            best_score = fragment_score
                            best_pos = j
                            best_fragment = fragment

                    # Также проверяем с частичным совпадением для более длинных фрагментов
                    for j in range(0, len(extended_window) - len(normalized_search) * 2 + 1, search_step):
                        fragment = extended_window[j:j + len(normalized_search) * 2]
                        fragment_score = fuzz.partial_ratio(fragment, normalized_search)

                        if fragment_score > best_score:
                            best_score = fragment_score
                            best_pos = j
                            best_fragment = fragment

                    if best_score >= threshold:
                        actual_start = extended_start + best_pos
                        actual_end = actual_start + len(best_fragment)

                        # Берем текст с контекстом
                        context_start = max(0, actual_start - 50)
                        context_end = min(len(text), actual_end + 50)
                        found_text = text[context_start:context_end]

                        matches.append((actual_start, actual_end, best_score, found_text, best_score))

        # Удаляем дубликаты и перекрывающиеся результаты
        matches.sort(key=lambda x: x[2], reverse=True)
        unique_matches = []
        seen_positions = set()

        for match in matches:
            start, end, score, found_text, match_quality = match
            overlap = False
            for seen_start, seen_end in seen_positions:
                if not (end <= seen_start - 10 or start >= seen_end + 10):
                    overlap = True
                    break

            if not overlap and score >= threshold:
                unique_matches.append(match)
                seen_positions.add((start, end))

        return unique_matches[:15]

    def fuzzy_search_with_details(self, text: str, search_text: str, threshold: float = 70.0) -> List[Dict[str, Any]]:
        """
        Расширенный нечеткий поиск с детальной информацией о каждом совпадении
        """
        if not text or not search_text:
            return []

        normalized_text = self.normalize_text(text)
        normalized_search = self.normalize_text(search_text)

        results = []

        # Разбиваем поисковый текст на слова для поиска по частям
        search_words = normalized_search.split()

        # Поиск по всему тексту
        all_matches = self.fuzzy_search_all(text, search_text, threshold)

        for match in all_matches:
            start, end, overall_score, context, exact_score = match

            # Получаем точный фрагмент, который был найден
            found_fragment = normalized_text[start:end] if start < len(normalized_text) else ""

            # Вычисляем различные метрики схожести для этого фрагмента
            ratio_score = fuzz.ratio(found_fragment, normalized_search, processor=utils.default_process)
            partial_score = fuzz.partial_ratio(found_fragment, normalized_search)
            token_sort_score = fuzz.token_sort_ratio(found_fragment, normalized_search)
            token_set_score = fuzz.token_set_ratio(found_fragment, normalized_search)

            # Вычисляем схожесть по словам
            word_scores = []
            for word in search_words:
                if word in found_fragment:
                    word_pos = found_fragment.find(word)
                    if word_pos != -1:
                        word_scores.append({
                            'word': word,
                            'score': 100.0,
                            'found': True,
                            'position': word_pos
                        })
                else:
                    # Ищем похожие слова
                    best_word_score = 0
                    best_word_match = ""
                    for found_word in found_fragment.split():
                        score = fuzz.ratio(word, found_word, processor=utils.default_process)
                        if score > best_word_score:
                            best_word_score = score
                            best_word_match = found_word
                    word_scores.append({
                        'word': word,
                        'score': best_word_score,
                        'found': False,
                        'closest_match': best_word_match
                    })

            avg_word_score = sum(w['score'] for w in word_scores) / len(word_scores) if word_scores else 0

            # Определяем лучший процент схожести для этого вхождения
            best_match_score = max(ratio_score, partial_score, token_sort_score, token_set_score, avg_word_score)

            results.append({
                'position': start,
                'end_position': end,
                'context': context,
                'fragment': found_fragment[:150] + "..." if len(found_fragment) > 150 else found_fragment,
                'scores': {
                    'overall': overall_score,
                    'exact_ratio': ratio_score,
                    'partial_ratio': partial_score,
                    'token_sort': token_sort_score,
                    'token_set': token_set_score,
                    'average_word': avg_word_score
                },
                'word_matches': word_scores,
                'best_match_score': best_match_score,
                'match_quality': self._get_match_quality(best_match_score)
            })

        # Сортируем по лучшему проценту схожести
        results.sort(key=lambda x: x['best_match_score'], reverse=True)

        return results

    def _get_match_quality(self, score: float) -> str:
        """Определяет качество совпадения на основе процента схожести"""
        if score >= 95:
            return "отличное"
        elif score >= 85:
            return "хорошее"
        elif score >= 75:
            return "среднее"
        elif score >= 60:
            return "низкое"
        else:
            return "очень низкое"

    def fuzzy_search_best(self, text: str, search_text: str) -> float:
        """Лучший результат нечеткого поиска"""
        if not text or not search_text:
            return 0.0

        normalized_text = self.normalize_text(text)
        normalized_search = self.normalize_text(search_text)

        # Используем несколько методов для лучшего результата
        score1 = fuzz.partial_ratio(normalized_text, normalized_search)
        score2 = fuzz.token_sort_ratio(normalized_text, normalized_search)
        score3 = fuzz.token_set_ratio(normalized_text, normalized_search)

        # Возвращаем максимальный результат
        return max(score1, score2, score3)

    def extract_tables(self, document_text: str) -> List[str]:
        """Извлечение таблиц из текста документа"""
        tables = []
        lines = document_text.split('\n')

        current_table = []
        in_table = False

        for line in lines:
            # Эвристика для таблиц
            if (('\t' in line or '|' in line or
                 (line.count('  ') > 3 and any(c.isdigit() for c in line))) and
                    len(line.strip()) > 10):

                if not in_table:
                    in_table = True
                current_table.append(line)
            elif in_table:
                if current_table:
                    tables.append('\n'.join(current_table))
                    current_table = []
                in_table = False

        if current_table:
            tables.append('\n'.join(current_table))

        return tables

    def extract_paragraphs_after_tables(self, document_text: str) -> List[str]:
        """Извлечение абзацев после таблиц"""
        paragraphs = []
        lines = document_text.split('\n')

        for i, line in enumerate(lines):
            if (('\t' in line or '|' in line or line.count('  ') > 3) and
                    i < len(lines) - 1):

                next_line = lines[i + 1].strip()
                if (next_line and
                        '\t' not in next_line and
                        '|' not in next_line and
                        not re.match(r'^\s*$', next_line)):
                    paragraphs.append(next_line)

        return paragraphs

    def check_version_comparison(self, subcheck: Dict, document_text: str, page_info: List[Tuple[int, int]]) -> Dict[
        str, Any]:
        """
        Проверка показателей назначения (версий) с поддержкой различных операторов сравнения
        """
        name = subcheck.get('name', 'Неизвестная проверка')

        result = {
            'name': name,
            'type': 'version_comparison',
            'passed': False,
            'is_error': False,
            'needs_verification': False,
            'matches': [],
            'score': 0.0,
            'message': '',
            'details': '',
            'page': 0,
            'position': '',
            'start_pos': -1,
            'end_pos': -1,
            'context': '',
            'section_results': []
        }

        try:
            # Поддержка разных форматов структуры данных
            version_sections = []

            # Проверяем различные возможные форматы
            if 'version_sections' in subcheck:
                # Формат с version_sections (как ожидается в add_check_dialog)
                version_sections = subcheck.get('version_sections', [])
                required_total = subcheck.get('required_total_indicators', len(version_sections))
                strict_mode = subcheck.get('strict_mode', True)

                logger.info(f"Формат 1: version_sections с {len(version_sections)} показателями")

            elif 'indicators' in subcheck:
                # Альтернативный формат с indicators
                version_sections = subcheck.get('indicators', [])
                required_total = subcheck.get('required_indicators', len(version_sections))
                strict_mode = subcheck.get('strict_indicators_mode', True)

                logger.info(f"Формат 2: indicators с {len(version_sections)} показателями")

            elif 'sections' in subcheck:
                # Еще один возможный формат
                version_sections = subcheck.get('sections', [])
                required_total = subcheck.get('required_sections', len(version_sections))
                strict_mode = subcheck.get('strict_sections_mode', True)

                logger.info(f"Формат 3: sections с {len(version_sections)} показателями")

            else:
                # Если нет вложенных секций, пробуем найти прямые поля
                # Возможно, это одиночный показатель
                if 'required_version' in subcheck or 'version_patterns' in subcheck:
                    single_section = {
                        'name': subcheck.get('name', 'Показатель'),
                        'required_version': subcheck.get('required_version', '0.0'),
                        'operator': subcheck.get('operator', '>='),
                        'version_patterns': subcheck.get('version_patterns', [])
                    }
                    version_sections = [single_section]
                    required_total = subcheck.get('required_total_indicators', 1)
                    strict_mode = subcheck.get('strict_mode', True)

                    logger.info(f"Формат 4: одиночный показатель")
                else:
                    # Если вообще нет данных, пытаемся найти вложенные структуры в любых ключах
                    for key, value in subcheck.items():
                        if isinstance(value, list) and len(value) > 0:
                            # Проверяем, похоже ли это на список показателей
                            if all(isinstance(item, dict) for item in value):
                                # Проверяем, есть ли у элементов нужные поля
                                sample_item = value[0]
                                if ('required_version' in sample_item or
                                        'version_patterns' in sample_item or
                                        'name' in sample_item and 'required_version' in sample_item):
                                    version_sections = value
                                    required_total = subcheck.get('required_total', len(value))
                                    strict_mode = subcheck.get('strict_mode', True)
                                    logger.info(f"Формат 5: найдены показатели в ключе '{key}'")
                                    break

            if not version_sections:
                result['message'] = "Нет показателей для проверки"
                result['details'] = "Не настроены показатели назначения или неправильный формат данных"
                logger.error(f"Не удалось найти показатели в структуре: {subcheck}")
                return result

            logger.info(f"Начало проверки показателей назначения: {name}")
            logger.info(f"Количество показателей: {len(version_sections)}")
            logger.info(f"Требуется: {required_total}, строгий режим: {strict_mode}")

            passed_sections = 0
            failed_sections = 0
            needs_check_sections = 0
            all_matches = []

            section_results = []

            for section_idx, section in enumerate(version_sections):
                # Получаем данные показателя с проверкой на наличие ключей
                section_name = section.get('name', f'Показатель {section_idx + 1}')
                required_version = section.get('required_version', '0.0')
                operator = section.get('operator', '>=')  # Получаем оператор, по умолчанию '>='

                # Паттерны могут быть в разных форматах
                patterns = section.get('version_patterns', [])
                if not patterns and 'patterns' in section:
                    patterns = section.get('patterns', [])
                if not patterns and 'regex' in section:
                    patterns = section.get('regex', [])
                if not patterns and 'regex_patterns' in section:
                    patterns = section.get('regex_patterns', [])

                # Если patterns все еще не список, пробуем преобразовать
                if patterns and not isinstance(patterns, list):
                    patterns = [str(patterns)]

                # Преобразуем оператор в читаемый вид для отчета
                operator_display = {
                    'eq': '=', '=': '=', '==': '=',
                    'ne': '≠', '!=': '≠', '<>': '≠',
                    'gt': '>', '>': '>',
                    'lt': '<', '<': '<',
                    'ge': '≥', '>=': '≥',
                    'le': '≤', '<=': '≤'
                }.get(operator, operator)

                logger.info(
                    f"Проверка показателя: {section_name}, требуемая версия: {required_version}, оператор: {operator}")
                logger.info(f"Паттерны: {patterns}")

                if not patterns:
                    logger.warning(f"Нет паттернов для показателя: {section_name}")
                    section_result = {
                        'name': section_name,
                        'passed': False,
                        'needs_check': True,
                        'result': f"⚠ {section_name}: нет паттернов для поиска версий",
                        'found_version': None,
                        'required_version': required_version,
                        'operator': operator,
                        'operator_display': operator_display,
                        'position': -1
                    }
                    section_results.append(section_result)
                    needs_check_sections += 1
                    continue

                # Ищем версии по всем паттернам
                all_found_versions = []
                for pattern in patterns:
                    if not pattern or not isinstance(pattern, str) or not pattern.strip():
                        continue

                    found_versions = self.search_versions_with_regex(document_text, pattern)
                    all_found_versions.extend(found_versions)
                    logger.info(f"По паттерну '{pattern}' найдено версий: {len(found_versions)}")

                # Если ничего не нашли
                if not all_found_versions:
                    logger.info(f"Не найдено версий для показателя: {section_name}")
                    section_result = {
                        'name': section_name,
                        'passed': False,
                        'needs_check': True,
                        'result': f"❓ {section_name}: не найдена информация о версии",
                        'found_version': None,
                        'required_version': required_version,
                        'operator': operator,
                        'operator_display': operator_display,
                        'position': -1
                    }
                    section_results.append(section_result)
                    needs_check_sections += 1
                    continue

                # Проверяем каждую найденную версию
                best_match = None
                best_result = False
                best_version = None
                best_position = -1
                best_end_position = -1

                for version_info in all_found_versions:
                    found_version = version_info['version']

                    # Проверяем условие с оператором
                    condition_met = self.compare_versions_with_operator(found_version, required_version, operator)

                    logger.info(
                        f"Сравнение версий: найдено '{found_version}', требуется {operator_display} {required_version}, "
                        f"результат: {'✓' if condition_met else '✗'}"
                    )

                    # Если условие выполнено, это хорошее совпадение
                    if condition_met:
                        # Если это первое совпадение или версия больше (лучше)
                        if best_result is False:
                            best_result = True
                            best_version = found_version
                            best_position = version_info['position']
                            best_end_position = version_info['end_position']
                            best_match = version_info
                        else:
                            # Сравниваем версии, чтобы выбрать лучшую (большую)
                            if self.compare_versions_with_operator(found_version, best_version, '>'):
                                best_version = found_version
                                best_position = version_info['position']
                                best_end_position = version_info['end_position']
                                best_match = version_info
                    else:
                        # Если условие не выполнено, но это лучшее из найденных (для отчета об ошибке)
                        if best_result is False:
                            if best_match is None:
                                best_match = version_info
                                best_version = found_version
                                best_position = version_info['position']
                                best_end_position = version_info['end_position']
                            else:
                                # Сравниваем, какая версия ближе к требуемой
                                if self.compare_versions_with_operator(found_version, best_version, '>'):
                                    best_version = found_version
                                    best_position = version_info['position']
                                    best_end_position = version_info['end_position']
                                    best_match = version_info

                # Определяем результат для этого показателя
                section_passed = best_result

                if section_passed:
                    passed_sections += 1
                    status_icon = "✅"
                    result_text = f"{section_name}: найдена версия {best_version} {operator_display} {required_version}"
                else:
                    failed_sections += 1
                    status_icon = "❌"
                    if best_version:
                        result_text = f"{section_name}: найдена версия {best_version} (требуется {operator_display} {required_version})"
                    else:
                        result_text = f"{section_name}: версия не найдена (требуется {operator_display} {required_version})"

                # Добавляем в общие совпадения
                if best_position != -1:
                    all_matches.append((best_position, best_end_position,
                                        f"{section_name}: {best_version if best_version else 'не найдено'} {operator_display} {required_version}"))

                section_result = {
                    'name': section_name,
                    'passed': section_passed,
                    'needs_check': False,
                    'result': f"{status_icon} {result_text}",
                    'found_version': best_version,
                    'required_version': required_version,
                    'operator': operator,
                    'operator_display': operator_display,
                    'position': best_position,
                    'end_position': best_end_position
                }

                section_results.append(section_result)
                logger.info(f"Результат для {section_name}: {section_result['result']}")

            # Определяем общий результат
            logger.info(
                f"Итоги: пройдено {passed_sections}, провалено {failed_sections}, требует проверки {needs_check_sections}"
            )

            if strict_mode:
                # Строгий режим: все секции должны быть выполнены
                result['passed'] = (passed_sections == len(version_sections))
                result['is_error'] = not result['passed']
                result['needs_verification'] = False
                logger.info(f"Строгий режим: результат {'ПРОЙДЕН' if result['passed'] else 'ПРОВАЛЕН'}")
            else:
                # Нестрогий режим: достаточно required_total
                result['passed'] = (passed_sections >= required_total)
                result['is_error'] = (passed_sections < required_total)
                result['needs_verification'] = (passed_sections < required_total and
                                                passed_sections + needs_check_sections >= required_total)
                logger.info(
                    f"Нестрогий режим: требуется {required_total}, найдено {passed_sections}, "
                    f"результат {'ПРОЙДЕН' if result['passed'] else 'ПРОВАЛЕН'}"
                )

            result['matches'] = all_matches
            result['score'] = (passed_sections / len(version_sections)) * 100 if version_sections else 0
            result['section_results'] = section_results

            # Формируем детальное сообщение
            details_parts = []
            for sr in section_results:
                details_parts.append(sr['result'])

            result['message'] = (f"Показатели назначения: {passed_sections} из {len(version_sections)} "
                                 f"(требуется {required_total})")
            result['details'] = "\n".join(details_parts)

            # Определяем позицию первой найденной ошибки
            for sr in section_results:
                if not sr['passed'] and not sr['needs_check'] and sr.get('position', -1) != -1:
                    pos_info = self.find_position_in_document(sr['position'], document_text, page_info)
                    result.update(pos_info)
                    logger.info(
                        f"Найдена позиция ошибки: страница {result.get('page')}, позиция {result.get('position')}"
                    )
                    break

            logger.info(f"Проверка завершена: {result['message']}")

        except Exception as e:
            error_msg = str(e)
            result['message'] = f"Ошибка проверки: {error_msg}"
            result['passed'] = False
            result['is_error'] = True
            logger.error(f"Критическая ошибка в проверке показателей назначения '{name}': {error_msg}")
            logger.exception(e)

        return result

    def check_subcheck(self, subcheck: Dict, document_text: str, page_info: List[Tuple[int, int]] = None) -> Dict[
        str, Any]:
        """Улучшенная проверка с точным определением позиций и фильтрацией ошибок"""
        check_type = subcheck.get('type', '')
        name = subcheck.get('name', 'Неизвестная проверка')

        result = {
            'name': name,
            'type': check_type,
            'passed': False,
            'is_error': False,
            'needs_verification': False,
            'matches': [],
            'detailed_matches': [],  # Новое поле для детальной информации о совпадениях
            'score': 0.0,
            'message': '',
            'details': '',
            'found_text': '',
            'search_terms': subcheck.get('aliases', []) if 'aliases' in subcheck else [],
            'search_text': subcheck.get('text', '') if 'text' in subcheck else '',
            'page': 0,
            'position': '',
            'start_pos': -1,
            'end_pos': -1,
            'line_number': 0,
            'context': ''
        }

        try:
            # Для проверки показателей назначения используем отдельный метод
            if check_type == 'version_comparison':
                return self.check_version_comparison(subcheck, document_text, page_info)

            elif check_type == 'no_text_present':
                aliases = subcheck.get('aliases', [])
                matches = self.exact_search_with_context(document_text, aliases)
                result['passed'] = len(matches) == 0
                result['is_error'] = not result['passed']
                result['matches'] = matches

                if matches:
                    first_match = matches[0]
                    result['start_pos'] = first_match[0]
                    result['end_pos'] = first_match[1]
                    result['context'] = first_match[3]

                    # Точное определение позиции
                    pos_info = self.find_position_in_document(first_match[0], document_text, page_info)
                    result.update(pos_info)

                    result['found_text'] = f"Найдено запрещенное: '{first_match[2]}'"
                    result['message'] = f"Найдено запрещенных вхождений: {len(matches)}"
                    result['details'] = f"Запрещенные слова: {', '.join(aliases[:5])}" + (
                        "..." if len(aliases) > 5 else "")
                else:
                    result['message'] = "Запрещенные упоминания не найдены"

            elif check_type == 'text_present':
                aliases = subcheck.get('aliases', [])
                matches = self.exact_search_with_context(document_text, aliases)
                result['passed'] = len(matches) > 0
                result['is_error'] = not result['passed']
                result['matches'] = matches

                if matches:
                    first_match = matches[0]
                    result['start_pos'] = first_match[0]
                    result['end_pos'] = first_match[1]
                    result['context'] = first_match[3]

                    pos_info = self.find_position_in_document(first_match[0], document_text, page_info)
                    result.update(pos_info)

                    result['found_text'] = f"Найдено: '{first_match[2]}'"
                    result['message'] = f"Найдено обязательных вхождений: {len(matches)}"
                else:
                    result['message'] = "Обязательный текст не найден"
                    result['details'] = f"Искали обязательные слова: {', '.join(aliases[:5])}" + (
                        "..." if len(aliases) > 5 else "")

            elif check_type == 'text_present_without':
                aliases = subcheck.get('aliases', [])
                without_aliases = subcheck.get('without_aliases', [])

                positive_matches = self.exact_search_with_context(document_text, aliases)
                negative_matches = self.exact_search_with_context(document_text, without_aliases)

                result['passed'] = len(positive_matches) > 0 and len(negative_matches) == 0
                result['is_error'] = not result['passed']
                result['matches'] = positive_matches + negative_matches

                if positive_matches:
                    first_match = positive_matches[0]
                    result['start_pos'] = first_match[0]
                    result['end_pos'] = first_match[1]
                    result['context'] = first_match[3]

                    pos_info = self.find_position_in_document(first_match[0], document_text, page_info)
                    result.update(pos_info)

                result[
                    'found_text'] = f"Найдено основных: {len(positive_matches)}, исключающих: {len(negative_matches)}"
                result['message'] = f"Основных вхождений: {len(positive_matches)}, исключающих: {len(negative_matches)}"
                result['details'] = f"Основные: {', '.join(aliases[:3])}, Исключающие: {', '.join(without_aliases[:3])}"

            elif check_type == 'fuzzy_text_present':
                text = subcheck.get('text', '')
                threshold = float(subcheck.get('threshold', 70.0))
                trust_threshold = float(subcheck.get('trust_threshold', 85.0))
                show_detailed_scores = subcheck.get('show_detailed_scores', True)

                # Общая оценка схожести
                overall_score = self.fuzzy_search_best(document_text, text)
                result['score'] = overall_score
                result['passed'] = overall_score >= trust_threshold
                result['is_error'] = not result['passed'] and overall_score < threshold
                result['needs_verification'] = threshold <= overall_score < trust_threshold

                # Получаем детальную информацию о каждом совпадении
                detailed_matches = self.fuzzy_search_with_details(document_text, text, threshold)

                if detailed_matches:
                    # Сохраняем детальную информацию
                    result['detailed_matches'] = detailed_matches

                    # Формируем matches для обратной совместимости
                    matches_info = []
                    for match in detailed_matches:
                        match_info = (
                            match['position'],
                            match['end_position'],
                            f"Схожесть: {match['best_match_score']:.1f}% ({match['match_quality']}) | {match['fragment'][:100]}"
                        )
                        matches_info.append(match_info)

                    result['matches'] = matches_info

                    # Берем первое совпадение для позиционирования
                    first_match = detailed_matches[0]
                    result['start_pos'] = first_match['position']
                    result['end_pos'] = first_match['end_position']

                    pos_info = self.find_position_in_document(first_match['position'], document_text, page_info)
                    result.update(pos_info)

                    # Формируем детальное сообщение со всеми совпадениями
                    if show_detailed_scores:
                        details_lines = [f"Найдено совпадений: {len(detailed_matches)}"]
                        details_lines.append("")

                        for i, match in enumerate(detailed_matches[:10], 1):  # Показываем первые 10
                            scores = match['scores']
                            details_lines.append(
                                f"{i}. Схожесть: {match['best_match_score']:.1f}% ({match['match_quality']})"
                            )
                            details_lines.append(f"   Точная: {scores['exact_ratio']:.1f}%")
                            details_lines.append(f"   Частичная: {scores['partial_ratio']:.1f}%")
                            details_lines.append(f"   По словам: {scores['average_word']:.1f}%")

                            # Показываем информацию по словам
                            if match['word_matches']:
                                words_info = []
                                for w in match['word_matches'][:5]:
                                    if w['score'] >= 80:
                                        words_info.append(f"✓ {w['word']}")
                                    elif w['score'] >= 50:
                                        words_info.append(f"~ {w['word']} ({w['score']:.0f}%)")
                                    else:
                                        words_info.append(f"✗ {w['word']}")
                                details_lines.append(f"   Слова: {', '.join(words_info)}")

                            details_lines.append(f"   Контекст: {match['fragment'][:100]}...")
                            details_lines.append("")

                        if len(detailed_matches) > 10:
                            details_lines.append(f"... и еще {len(detailed_matches) - 10} совпадений")

                        result['details'] = "\n".join(details_lines)

                    result[
                        'found_text'] = f"Найдено совпадений: {len(detailed_matches)}, лучшее: {detailed_matches[0]['best_match_score']:.1f}%"

                result['message'] = f"Общая схожесть: {overall_score:.1f}% (порог: {threshold}/{trust_threshold}%)"

            elif check_type == 'no_fuzzy_text_present':
                text = subcheck.get('text', '')
                threshold = float(subcheck.get('threshold', 70.0))
                trust_threshold = float(subcheck.get('trust_threshold', 85.0))
                show_detailed_scores = subcheck.get('show_detailed_scores', True)

                overall_score = self.fuzzy_search_best(document_text, text)
                result['score'] = overall_score
                result['passed'] = overall_score < threshold
                result['is_error'] = not result['passed']
                result['needs_verification'] = threshold <= overall_score < trust_threshold

                # Получаем детальную информацию о каждом совпадении
                detailed_matches = self.fuzzy_search_with_details(document_text, text, threshold)

                if detailed_matches:
                    result['detailed_matches'] = detailed_matches

                    matches_info = []
                    for match in detailed_matches:
                        match_info = (
                            match['position'],
                            match['end_position'],
                            f"Схожесть: {match['best_match_score']:.1f}% | {match['fragment'][:100]}"
                        )
                        matches_info.append(match_info)

                    result['matches'] = matches_info

                    first_match = detailed_matches[0]
                    result['start_pos'] = first_match['position']
                    result['end_pos'] = first_match['end_position']

                    pos_info = self.find_position_in_document(first_match['position'], document_text, page_info)
                    result.update(pos_info)

                    if show_detailed_scores:
                        details_lines = [f"Найдено нежелательных совпадений: {len(detailed_matches)}"]
                        for i, match in enumerate(detailed_matches[:5], 1):
                            details_lines.append(
                                f"  {i}. Схожесть: {match['best_match_score']:.1f}%"
                            )
                            details_lines.append(f"     Текст: {match['fragment'][:100]}...")

                        if len(detailed_matches) > 5:
                            details_lines.append(f"  ... и еще {len(detailed_matches) - 5}")

                        result['details'] = "\n".join(details_lines)

                    result['found_text'] = f"Найдено похожих: {len(detailed_matches)}"

                result['message'] = f"Схожесть: {overall_score:.1f}% (порог: {threshold}%)"

            elif check_type == 'text_present_in_any_table':
                aliases = subcheck.get('aliases', [])
                tables = self.extract_tables(document_text)

                table_matches = []
                for table_idx, table in enumerate(tables):
                    matches = self.exact_search_with_context(table, aliases)
                    if matches:
                        table_start = document_text.find(table)
                        for match in matches:
                            global_start = table_start + match[0]
                            global_end = table_start + match[1]
                            table_matches.append((global_start, global_end,
                                                  f"Таблица {table_idx + 1}: {match[2]}",
                                                  match[3]))

                result['passed'] = len(table_matches) > 0
                result['is_error'] = not result['passed']
                result['matches'] = table_matches

                if table_matches:
                    first_match = table_matches[0]
                    result['start_pos'] = first_match[0]
                    result['end_pos'] = first_match[1]
                    result['context'] = first_match[3]

                    pos_info = self.find_position_in_document(first_match[0], document_text, page_info)
                    result.update(pos_info)

                    result['found_text'] = f"Найдено в {len(set([m[2].split(':')[0] for m in table_matches]))} таблицах"
                    result['message'] = f"Найдено в таблицах: {len(table_matches)} вхождений"

            elif check_type == 'fuzzy_text_present_after_any_table':
                text = subcheck.get('text', '')
                threshold = float(subcheck.get('threshold', 70.0))
                trust_threshold = float(subcheck.get('trust_threshold', 85.0))

                paragraphs = self.extract_paragraphs_after_tables(document_text)
                best_score = 0.0
                best_match = ""
                best_position = ""
                best_start_pos = -1

                for para_idx, paragraph in enumerate(paragraphs):
                    score = self.fuzzy_search_best(paragraph, text)
                    if score > best_score:
                        best_score = score
                        best_match = paragraph[:100] + "..." if len(paragraph) > 100 else paragraph
                        best_position = f"После таблицы {para_idx + 1}"

                        # Находим позицию абзаца в документе
                        para_pos = document_text.find(paragraph)
                        if para_pos != -1:
                            best_start_pos = para_pos

                result['score'] = best_score
                result['passed'] = best_score >= trust_threshold
                result['is_error'] = not result['passed'] and best_score < threshold
                result['needs_verification'] = threshold <= best_score < trust_threshold
                result['message'] = f"Схожесть: {best_score:.1f}% (порог: {threshold}/{trust_threshold}%)"
                result['details'] = f"Искали после таблиц: '{text[:50]}...'"
                result['position'] = best_position
                result['start_pos'] = best_start_pos

                if best_match:
                    result['found_text'] = f"Абзац: {best_match}"

                # Определяем страницу
                if best_start_pos != -1 and page_info:
                    for page_num, (start, end) in enumerate(page_info, 1):
                        if start <= best_start_pos < end:
                            result['page'] = page_num
                            break

            elif check_type == 'combined_check':
                """
                Комбинированная проверка с логическими операторами
                """
                conditions = subcheck.get('conditions', [])
                logic_operator = subcheck.get('logic_operator', 'AND')
                required_passed = subcheck.get('required_passed', len(conditions))

                condition_results = []
                passed_conditions = 0
                all_matches = []
                all_detailed_matches = []

                for condition in conditions:
                    # Проверяем каждое условие отдельно
                    cond_result = self.check_subcheck(condition, document_text, page_info)
                    condition_passed = cond_result.get('passed', False)

                    condition_results.append({
                        'name': condition.get('name', 'Неизвестное условие'),
                        'passed': condition_passed,
                        'message': cond_result.get('message', ''),
                        'details': cond_result.get('details', ''),
                        'matches': cond_result.get('matches', []),
                        'detailed_matches': cond_result.get('detailed_matches', [])
                    })

                    if condition_passed:
                        passed_conditions += 1

                    # Собираем все совпадения
                    all_matches.extend(cond_result.get('matches', []))
                    all_detailed_matches.extend(cond_result.get('detailed_matches', []))

                # Определяем общий результат в зависимости от оператора
                if logic_operator == 'AND':
                    result['passed'] = (passed_conditions == len(conditions))
                else:  # OR
                    result['passed'] = (passed_conditions >= required_passed)

                result['is_error'] = not result['passed']
                result['needs_verification'] = False
                result['matches'] = all_matches
                result['detailed_matches'] = all_detailed_matches
                result['score'] = (passed_conditions / len(conditions)) * 100 if conditions else 0
                result['condition_results'] = condition_results

                # Формируем детальное сообщение
                details_parts = []
                for cr in condition_results:
                    status = "✓" if cr['passed'] else "✗"
                    details_parts.append(f"{status} {cr['name']}: {cr['message']}")

                result[
                    'message'] = f"Комбинированная проверка: {passed_conditions} из {len(conditions)} условий выполнено"
                result['details'] = "\n".join(details_parts)

                # Определяем позицию первой найденной ошибки
                for cr in condition_results:
                    if not cr['passed'] and cr.get('matches'):
                        first_match = cr['matches'][0]
                        if len(first_match) >= 2:
                            pos_info = self.find_position_in_document(first_match[0], document_text, page_info)
                            result.update(pos_info)
                            break

        except Exception as e:
            result['message'] = f"Ошибка проверки: {str(e)}"
            result['passed'] = False
            result['is_error'] = True
            logger.error(f"Ошибка в проверке {name}: {str(e)}")
            logger.exception(e)

        return result

    def check_document(self, document_text: str, selected_checks: List[str] = None,
                       page_info: List[Tuple[int, int]] = None) -> List[Dict]:
        """Основная функция проверки документа"""
        if not self.config.get('checks'):
            return []

        results = []
        for check_group in self.config.get('checks', []):
            group_name = check_group.get('group', '')
            subchecks = check_group.get('subchecks', [])

            for subcheck in subchecks:
                check_name = subcheck.get('name', '')
                if selected_checks and check_name not in selected_checks:
                    continue

                result = self.check_subcheck(subcheck, document_text, page_info)
                result['group'] = group_name
                results.append(result)

        return results