import os
from datetime import datetime
from tqdm import tqdm
from typing import Optional, List, Tuple
import logging

# Флаг для активации записи логов в файл
enable_file_logging = False  # Поставьте True, если нужно записывать логи в файл

# Настройка логирования
log_handlers = [logging.StreamHandler()]  # Логи в консоль
if enable_file_logging:
    log_handlers.append(logging.FileHandler('debug.log', mode='w'))  # Логи в файл

logging.basicConfig(
    level=logging.DEBUG,  # Уровень логирования
    format='%(asctime)s - %(levelname)s - %(message)s',  # Формат сообщений
    handlers=log_handlers  # Обработчики (консоль и/или файл)
)

# Конфигурация
CONFIG = {
    'input_file': r'U:\ГЕНЕТИКА НОВОЕ ЗДАНИЕ\BASE\FAMILIAL SEARCH\familial_search_list.txt',
    'output_directory': r'U:\ГЕНЕТИКА НОВОЕ ЗДАНИЕ\BASE\FAMILIAL SEARCH',
    'conclusion_list_file': r'U:\ГЕНЕТИКА НОВОЕ ЗДАНИЕ\BASE\FAMILIAL SEARCH\Список.txt'
}

def extract_number(value: str) -> str:
    """
    Извлекает числовое значение до пробела или тире.
    
    Args:
        value: строка, из которой нужно извлечь номер
        
    Returns:
        str: извлеченный номер или пустая строка в случае ошибки
    """
    try:
        result = value.split()[0].split('-')[0]
        logging.debug(f"Извлечен номер {result} из значения {value}")
        return result
    except Exception as e:
        logging.error(f"Ошибка при извлечении номера из {value}: {e}")
        return ""

def extract_last_two_digits(value: str) -> Optional[str]:
    """
    Извлекает две последние цифры после тире (обычно означающие год).
    
    Args:
        value: строка вида "номер-год" (например, "123-23")
        
    Returns:
        Optional[str]: две последние цифры года или None в случае ошибки
    """
    try:
        parts = value.split('-')
        if len(parts) > 1 and len(parts[-1]) >= 2 and parts[-1][-2:].isdigit():
            result = parts[-1][-2:]
            logging.debug(f"Извлечены последние цифры {result} из значения {value}")
            return result
        
        logging.warning(f"Не удалось извлечь последние цифры из {value}")
        return None
        
    except Exception as e:
        logging.error(f"Ошибка при извлечении цифр из {value}: {e}")
        return None

def load_conclusion_list(file_path: str) -> dict:
    """
    Загружает данные из файла Список.txt и создает словарь для поиска эксперта по номеру заключения.
    
    Args:
        file_path: путь к файлу Список.txt
        
    Returns:
        dict: словарь вида {номер-год: эксперт}
    """
    logging.info(f"Загрузка данных из файла списка: {file_path}")
    conclusion_data = {}
    
    try:
        # Используем utf-8-sig для автоматического удаления BOM
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            lines = f.readlines()
            if not lines:
                logging.warning("Файл списка пуст")
                return conclusion_data
            
            # Парсим заголовок для определения индексов колонок
            headers = lines[0].strip().split('\t')
            try:
                number_idx = headers.index("Номер")
                date_idx = headers.index("Дата")
                expert_idx = headers.index("Эксперт")
                logging.debug(f"Найдены колонки: Номер - индекс {number_idx}, Дата - индекс {date_idx}, Эксперт - индекс {expert_idx}")
            except ValueError as e:
                logging.error(f"Не удалось найти необходимые колонки в заголовке: {e}")
                return conclusion_data
            
            # Обрабатываем строки с данными
            for line_num, line in enumerate(lines[1:], start=2):
                parts = line.strip().split('\t')
                if len(parts) <= max(number_idx, date_idx, expert_idx):
                    logging.warning(f"Пропущена строка {line_num}: недостаточно полей")
                    continue
                
                number = parts[number_idx].strip()
                date_str = parts[date_idx].strip()
                expert = parts[expert_idx].strip()
                
                if number and date_str and expert:
                    try:
                        # Удаляем начальные нули из номера
                        clean_number = str(int(number))
                        
                        # Извлекаем год из даты (формат: DD.MM.YYYY HH:MM:SS)
                        year_full = date_str.split('.')[2].split()[0]  # Берем часть года до пробела
                        year_suffix = year_full[-2:]  # Последние две цифры года
                        
                        # Формируем ключ для поиска
                        search_key = f"{clean_number}-{year_suffix}"
                        conclusion_data[search_key] = expert
                        logging.debug(f"Добавлена запись: ключ '{search_key}' -> эксперт '{expert}'")
                        
                    except (ValueError, IndexError) as e:
                        logging.warning(f"Ошибка обработки данных в строке {line_num}: номер '{number}', дата '{date_str}': {e}")
                else:
                    logging.warning(f"Пропущена строка {line_num}: пустые данные (номер: '{number}', дата: '{date_str}', эксперт: '{expert}')")
        
        logging.info(f"Успешно загружено {len(conclusion_data)} записей из файла списка")
        
    except Exception as e:
        logging.error(f"Ошибка при загрузке файла списка {file_path}: {e}")
    
    return conclusion_data

def find_expert_in_list(conclusion_number: str, conclusion_data: dict) -> str:
    """
    Ищет эксперта для указанного номера заключения в загруженных данных.
    
    Args:
        conclusion_number: номер заключения из основного файла (формат "1359 5I-25")
        conclusion_data: словарь с данными из файла списка
        
    Returns:
        str: фамилия эксперта или сообщение о невозможности определить
    """
    logging.debug(f"Поиск эксперта для номера: {conclusion_number}")
    
    # Извлекаем базовый номер (без возможных суффиксов)
    base_number = extract_number(conclusion_number)
    if not base_number:
        logging.warning(f"Не удалось извлечь номер из: {conclusion_number}")
        return "не найдено (ошибка формата)"
    
    # Извлекаем год из номера заключения
    year_suffix = extract_last_two_digits(conclusion_number)
    
    if year_suffix:
        # Номер содержит год - формируем ключ для поиска
        search_key = f"{base_number}-{year_suffix}"
        expert = conclusion_data.get(search_key)
        
        if expert:
            logging.info(f"Найден эксперт для {search_key}: {expert}")
            return expert
        else:
            logging.warning(f"Эксперт не найден для ключа: {search_key}")
            return "не найдено"
    else:
        # Номер без года - не можем определить эксперт
        logging.warning(f"Номер без года: {conclusion_number} - невозможно определить эксперта")
        return "(2023 г./PP/ошибка)"

def process_line(line: str) -> Optional[Tuple]:
    """
    Обработка одной строки из входного файла.
    
    Args:
        line: строка из входного файла
        
    Returns:
        Optional[Tuple]: кортеж с извлеченными данными или None в случае ошибки
    """
    logging.debug(f"Обработка строки: {line.strip()}")
    
    try:
        # Разбиваем строку на части по табуляции
        parts = line.strip().split('\t')
        
        # Проверяем количество полей
        if len(parts) < 12:
            logging.warning(
                f"Пропущена некорректная строка (недостаточно полей): {line.strip()}\n"
                f"Найдено полей: {len(parts)}, ожидалось: не менее 12"
            )
            return None
        
        # Извлекаем нужные поля
        result = (
            parts[0].strip(),  # profile_trace
            parts[1].strip(),  # candidate_hit
            parts[4].strip(),  # relationship
            parts[5].strip(),  # lr
            parts[6].strip(),  # exclusions
            parts[7].strip(),  # overlapping_markers
            parts[8].strip(),  # shared_alleles
        )
        
        # Проверяем, что все поля не пустые
        if not all(result):
            logging.warning(
                f"Обнаружены пустые поля в строке: {line.strip()}\n"
                f"Извлеченные значения: {result}"
            )
            return None
            
        logging.debug(
            f"Успешно обработана строка:\n"
            f"- profile_trace: {result[0]}\n"
            f"- candidate_hit: {result[1]}\n"
            f"- relationship: {result[2]}\n"
            f"- lr: {result[3]}\n"
            f"- exclusions: {result[4]}\n"
            f"- overlapping_markers: {result[5]}\n"
            f"- shared_alleles: {result[6]}"
        )
        
        return result
        
    except Exception as e:
        logging.error(
            f"Ошибка при обработке строки:\n"
            f"Строка: {line}\n"
            f"Ошибка: {str(e)}"
        )
        return None

def process_mismatch(profile_trace: str, candidate_hit: str, conclusion_data: dict) -> Tuple[str, str]:
    """
    Обработка пары несовпадающих заключений и поиск экспертов в файле списка.
    
    Args:
        profile_trace: номер первого заключения
        candidate_hit: номер второго заключения  
        conclusion_data: словарь с данными из файла списка
        
    Returns:
        Tuple[str, str]: пара строк с результатами поиска
    """
    logging.info(f"Обработка несовпадения между {profile_trace} и {candidate_hit}")
    
    # Ищем экспертов в загруженных данных
    profile_expert = find_expert_in_list(profile_trace, conclusion_data)
    candidate_expert = find_expert_in_list(candidate_hit, conclusion_data)
    
    logging.info(f"Для заключения {profile_trace} найден эксперт: {profile_expert}")
    logging.info(f"Для заключения {candidate_hit} найден эксперт: {candidate_expert}")
    
    profile_result = f"{profile_trace} {profile_expert}"
    candidate_result = f"{candidate_hit} {candidate_expert}"
    
    return profile_result, candidate_result

def get_normalized_key(profile: str, candidate: str, *rest) -> tuple:
    """
    Создает нормализованный ключ для сравнения "зеркальных" дубликатов.
    Ключ формируется как отсортированная пара (profile, candidate) + остальные данные.
    
    Args:
        profile: значение первого столбца (profile_trace)
        candidate: значение второго столбца (candidate_hit)
        *rest: остальные поля строки
        
    Returns:
        tuple: нормализованный ключ для сравнения
    """
    # Сортируем profile и candidate для создания уникального ключа
    sorted_pair = tuple(sorted((profile, candidate)))
    return (sorted_pair,) + rest

def main():
    """
    Основная функция программы.
    Выполняет следующие шаги:
    1. Читает входной файл
    2. Загружает данные из файла списка
    3. Находит несовпадающие записи
    4. Проводит многоуровневую фильтрацию дубликатов
    5. По запросу пользователя ищет экспертов в списке или просто выводит список несовпадений
    6. Сохраняет результаты в выходной файл
    """
    logging.info("Начало работы программы")

    # Загружаем данные из файла списка
    logging.info(f"Загрузка данных из файла списка: {CONFIG['conclusion_list_file']}")
    conclusion_data = load_conclusion_list(CONFIG['conclusion_list_file'])
    logging.info(f"Загружено {len(conclusion_data)} записей о заключениях")

    # Формируем имя выходного файла с текущей датой
    current_date = datetime.now().strftime('%d.%m.%Y')
    output_file_name = f"{current_date} несовпадения.txt"
    
    # Получаем путь к директории для сохранения выходного файла
    output_directory = CONFIG.get('output_directory', '.')
    
    # Полный путь к выходному файлу
    output_file_path = os.path.join(output_directory, output_file_name)
    
    mismatched_lines = []

    # Первый проход - поиск несовпадений
    try:
        logging.info(f"Чтение входного файла: {CONFIG['input_file']}")
        with open(CONFIG['input_file'], 'r', encoding='utf-8') as input_file:
            lines = input_file.readlines()
            logging.debug(f"Прочитано {len(lines)} строк")

            if len(lines) <= 1:
                logging.error("Входной файл пуст или содержит только заголовок")
                print("Ошибка: входной файл пуст или содержит только заголовок")
                return

    except Exception as e:
        logging.error(f"Ошибка при чтении входного файла: {e}")
        print(f"Ошибка при чтении входного файла: {e}")
        return

    # Поиск несовпадений
    logging.info("Начало поиска несовпадений")
    for line_num, line in enumerate(tqdm(lines[1:], desc="Поиск несовпадений"), start=2):
        result = process_line(line)
        if not result:
            logging.warning(f"Пропущена строка {line_num}: {line.strip()}")
            continue

        profile_trace, candidate_hit, relationship, lr, exclusions, overlapping_markers, shared_alleles = result
        if profile_trace != candidate_hit:
            mismatched_lines.append((line, profile_trace, candidate_hit, relationship, lr, exclusions, overlapping_markers, shared_alleles))
            logging.debug(f"Найдено несовпадение в строке {line_num}: {profile_trace} - {candidate_hit}")

    # Проверка наличия несовпадений
    if not mismatched_lines:
        logging.info("Все профили совпали")
        print("Все профили совпали")
        return

    logging.info(f"Найдено {len(mismatched_lines)} несовпадений")
    print(f"Найдено {len(mismatched_lines)} несовпадений")

    # МНОГОУРОВНЕВАЯ ФИЛЬТРАЦИЯ ДУБЛИКАТОВ:

    # 1. УДАЛЕНИЕ ПОЛНЫХ ДУБЛИКАТОВ
    # Удаляем строки, где все поля (кроме исходной строки) полностью идентичны
    unique_lines = []
    seen = set()
    for entry in mismatched_lines:
        # entry[1:] - берем все элементы, кроме исходной строки (line)
        key = tuple(entry[1:])  # (profile_trace, candidate_hit, relationship, lr, exclusions, overlapping_markers, shared_alleles)
        if key not in seen:
            seen.add(key)
            unique_lines.append(entry)
    
    duplicates_removed = len(mismatched_lines) - len(unique_lines)
    logging.info(f"Удалено полных дубликатов: {duplicates_removed}")
    if duplicates_removed > 0:
        print(f"Удалено полных дубликатов: {duplicates_removed}")

    # 2. ВЫБОР DIRECT-MATCH ВМЕСТО PARENT-CHILD
    # Если для одной и той же пары (A,B) есть и Direct-match и Parent-child, оставляем только Direct-match
    prioritized_lines = {}
    for entry in unique_lines:
        line, profile, candidate, rel, lr, exclusions, overlapping, shared = entry
        key = (profile, candidate)  # Ключ - конкретная пара заключений
        
        current_entry = {
            "key": key,
            "rel": rel,
            "entry": entry
        }

        if key in prioritized_lines:
            existing_rel = prioritized_lines[key]["rel"]
            # Если текущая строка имеет Direct-match, а существующая — Parent-child → заменяем
            if (rel == "Direct-match" and existing_rel == "Parent-child"):
                prioritized_lines[key] = current_entry
                logging.debug(f"Заменен Parent-child на Direct-match для пары {key}")
            # Если обе Direct-match или обе Parent-child - оставляем существующую (первую встреченную)
        else:
            prioritized_lines[key] = current_entry

    # Преобразуем словарь обратно в список
    filtered_lines = [v["entry"] for v in prioritized_lines.values()]
    direct_match_improvements = len(unique_lines) - len(filtered_lines)
    logging.info(f"После выбора Direct-match: {len(filtered_lines)} строк")
    if direct_match_improvements > 0:
        print(f"Улучшено за счет выбора Direct-match: {direct_match_improvements} пар")

    # 3. УДАЛЕНИЕ "ЗЕРКАЛЬНЫХ" ДУБЛИКАТОВ
    # Удаляем пары (A,B) и (B,A), которые представляют одно и то же несовпадение
    mirror_duplicates = {}
    final_lines = []
    
    logging.info("Начало удаления 'зеркальных' дубликатов")
    for entry in filtered_lines:
        line, profile, candidate, rel, lr, exclusions, overlapping, shared = entry
        
        # Создаем нормализованный ключ - сортируем пару (A,B) чтобы (A,B) и (B,A) стали одинаковыми
        key = get_normalized_key(profile, candidate, rel, lr, exclusions, overlapping, shared)
        
        if key in mirror_duplicates:
            original_profile, original_candidate = mirror_duplicates[key]
            logging.debug(
                f"Обнаружена 'зеркальная' пара:\n"
                f"  Уже есть: {original_profile}, {original_candidate}\n"
                f"  Пропускаем: {profile}, {candidate}"
            )
            continue
            
        # Сохраняем оригинальные значения для логирования
        mirror_duplicates[key] = (profile, candidate)
        final_lines.append(entry)
    
    # Подсчет удаленных "зеркальных" дубликатов
    mirror_duplicates_count = len(filtered_lines) - len(final_lines)
    if mirror_duplicates_count > 0:
        logging.info(f"Удалено 'зеркальных' дубликатов: {mirror_duplicates_count}")
        print(f"Удалено 'зеркальных' дубликатов: {mirror_duplicates_count}")
    else:
        logging.info("Не обнаружено 'зеркальных' дубликатов")

    # Итоговый подсчет
    total_removed = len(mismatched_lines) - len(final_lines)
    logging.info(f"Итоговая фильтрация: из {len(mismatched_lines)} осталось {len(final_lines)} уникальных несовпадений")
    print(f"После фильтрации: {len(final_lines)} уникальных несовпадений")

    # Запрос на поиск в списке
    search_list = input("Произвести поиск экспертов в списке? (да/нет): ").strip().lower()

    try:
        with open(output_file_path, 'w', encoding='utf-8') as output_file:
            if search_list == 'да':
                logging.info("Начало поиска экспертов в списке")
                print("Начинаем поиск экспертов в списке...")

                for idx, (_, profile_trace, candidate_hit, relationship, lr, exclusions, overlapping_markers, shared_alleles) in enumerate(
                    tqdm(final_lines, desc="Поиск экспертов"), start=1
                ):
                    logging.debug(f"Обработка пары {idx}/{len(final_lines)}")

                    profile_result, candidate_result = process_mismatch(
                        profile_trace, 
                        candidate_hit, 
                        conclusion_data
                    )

                    # Формируем и записываем строку результата
                    output_line = (f"{profile_result}, {candidate_result}, "
                                 f"{relationship}, {lr}, {exclusions}, "
                                 f"{overlapping_markers}, {shared_alleles}\n\n")
                    output_file.write(output_line)
                    logging.debug(f"Записана строка: {output_line.strip()}")
            else:
                logging.info("Поиск в списке отменен пользователем")
                print("Поиск в списке отменен. Вывод списка несовпадений...")

                for idx, (line, profile_trace, candidate_hit, relationship, lr, exclusions, overlapping_markers, shared_alleles) in enumerate(
                    final_lines, start=1
                ):
                    output_line = (f"{profile_trace}, {candidate_hit}, "
                                 f"{relationship}, {lr}, {exclusions}, "
                                 f"{overlapping_markers}, {shared_alleles}\n\n")
                    output_file.write(output_line)
                    logging.debug(f"Записана строка без поиска экспертов: {output_line.strip()}")

        logging.info(f"Результаты сохранены в файл: {output_file_path}")
        print(f"Результаты сохранены в файл: {output_file_path}")

    except Exception as e:
        logging.error(f"Ошибка при записи результатов: {e}")
        print(f"Ошибка при записи результатов: {e}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logging.critical(f"Критическая ошибка в программе: {e}")
        print(f"Произошла критическая ошибка: {e}")
    finally:
        logging.info("Завершение работы программы")
        print("Программа завершила работу.")