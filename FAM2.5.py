import win32com.client
import os
from datetime import datetime
from docx import Document
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
    'archive_base': r'U:\Архив\Архив 20',
    'surnames': [
        "Астраханцева", "Бевза", "Ермолаева", "Мамедова",
        "Разумов", "Сафарова", "Слепцова", "Игнатьева", "Угрюмова",
        "Хильман", "Хохлова"
    ],
    'output_directory': r'U:\ГЕНЕТИКА НОВОЕ ЗДАНИЕ\BASE\FAMILIAL SEARCH'  # Укажите нужный путь
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
        
    Example:
        >>> extract_last_two_digits("123-23")
        "23"
        >>> extract_last_two_digits("123-2023")
        "23"
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

def get_document_path(archive_path: str, number: str, year_suffix: str) -> Optional[str]:
    """
    Поиск документа с учетом формата номер-год.
    
    Args:
        archive_path: путь к папке архива
        number: номер заключения
        year_suffix: две цифры года
        
    Returns:
        Optional[str]: полный путь к найденному документу или None, если документ не найден
        
    Example:
        >>> get_document_path("C:/Archive/2023", "123", "23")
        "C:/Archive/2023/123-23.docx"
    """
    logging.info(f"Поиск документа в {archive_path} для номера {number}-{year_suffix}")
    
    # Поддерживаемые расширения файлов
    extensions = ['.doc', '.docx', '.DOC', '.DOCX']
    file_pattern = f"{number}-{year_suffix}"
    
    # Проверяем каждое возможное расширение
    for ext in extensions:
        path = os.path.join(archive_path, f"{file_pattern}{ext}")
        logging.debug(f"Проверяем путь: {path}")
        
        if os.path.exists(path):
            logging.info(f"Найден документ: {path}")
            return path
            
    logging.warning(f"Документ не найден в {archive_path} для номера {file_pattern}")
    return None

def read_doc_file(file_path: str) -> str:
    """
    Чтение старого формата .doc файлов через COM объект Word.
    Пытается получить текст несколькими способами в случае неудачи.
    
    Args:
        file_path: полный путь к .doc файлу
        
    Returns:
        str: извлеченный текст документа или пустая строка в случае ошибки
    """
    logging.info(f"Чтение .doc файла: {file_path}")
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(file_path)
        
        try:
            # Метод 1: Получаем весь текст напрямую
            logging.debug("Попытка получить текст через Content.Text")
            text = doc.Content.Text
            logging.debug(f"Успешно получен текст методом 1, длина: {len(text)}")
            
        except Exception as e1:
            logging.warning(f"Не удалось получить текст через Content.Text: {e1}")
            try:
                # Метод 2: Читаем по параграфам
                logging.debug("Попытка чтения по параграфам")
                text = ''
                for para in doc.Paragraphs:
                    text += para.Range.Text + '\n'
                logging.debug(f"Успешно получен текст методом 2, длина: {len(text)}")
                
            except Exception as e2:
                logging.warning(f"Не удалось прочитать по параграфам: {e2}")
                try:
                    # Метод 3: Через буфер обмена
                    logging.debug("Попытка чтения через буфер обмена")
                    doc.Content.Copy()
                    text = word.Selection.Text
                    logging.debug(f"Успешно получен текст методом 3, длина: {len(text)}")
                    
                except Exception as e3:
                    logging.error(f"Все методы чтения файла не удались: {e3}")
                    text = ""
        
        finally:
            # Закрываем документ и Word в любом случае
            doc.Close()
            word.Quit()
        
        # Проверяем результат
        if text:
            logging.debug(f"Успешно прочитано {len(text)} символов")
            logging.debug(f"Первые 100 символов: {text[:100]}")
        else:
            logging.warning("Получен пустой текст")
            
        return text
        
    except Exception as e:
        logging.error(f"Критическая ошибка при чтении .doc файла {file_path}: {e}")
        return ""

def safe_read_document(doc_path: str) -> Optional[str]:
    """
    Безопасное чтение документа Word обоих форматов (.doc и .docx) с учетом форматирования.
    
    Args:
        doc_path: полный путь к документу Word
        
    Returns:
        Optional[str]: текст документа или None в случае ошибки
        
    Example:
        >>> text = safe_read_document("path/to/document.docx")
        >>> if text:
        >>>     print(f"Документ содержит {len(text)} символов")
    """
    logging.info(f"Чтение документа: {doc_path}")
    try:
        # Определяем формат документа и выбираем соответствующий метод чтения
        if doc_path.lower().endswith('.doc'):
            logging.debug("Обнаружен формат .doc, используем COM объект")
            return read_doc_file(doc_path)
        else:
            logging.debug("Обнаружен формат .docx, используем python-docx")
            doc = Document(doc_path)
            full_text = []
            
            # Читаем документ по параграфам
            for i, paragraph in enumerate(doc.paragraphs):
                para_text = ''
                # Собираем текст из всех runs (фрагментов с разным форматированием)
                for run in paragraph.runs:
                    para_text += run.text
                    if run.text.strip():  # Логируем только непустые фрагменты
                        logging.debug(f"Параграф {i}, фрагмент текста: {run.text[:50]}...")
                
                if para_text.strip():  # Добавляем только непустые параграфы
                    full_text.append(para_text)
                    logging.debug(f"Добавлен параграф {i}: {para_text[:50]}...")
            
            result = '\n'.join(full_text)
            logging.info(f"Успешно прочитан документ, всего {len(result)} символов")
            logging.debug(f"Первые 200 символов: {result[:200]}")
            return result
            
    except Exception as e:
        logging.error(f"Ошибка при чтении документа {doc_path}: {str(e)}")
        return None

def find_surnames(doc_path: str, surnames: List[str]) -> str:
    """
    Поиск фамилий в документе с учётом различных вариантов написания и форматирования.
    
    Args:
        doc_path: путь к документу
        surnames: список фамилий для поиска
        
    Returns:
        str: найденная фамилия или "не найдено"
        
    Example:
        >>> find_surnames("path/to/doc.docx", ["Иванов", "Петров"])
        "Иванов"
    """
    logging.info(f"Поиск фамилий в документе: {doc_path}")
    
    text = safe_read_document(doc_path)
    if not text:
        logging.warning(f"Получен пустой текст из документа: {doc_path}")
        return "не найдено"
    
    try:
        # Выводим отладочную информацию
        logging.debug(f"Первые 200 символов текста: {text[:200]}")
        special_chars = [f"'{c}' ({ord(c)})" for c in text[:100] if ord(c) > 127]
        if special_chars:
            logging.debug(f"Найдены специальные символы: {special_chars}")
        
        # Нормализуем текст
        text = ' '.join(text.split())  # Убираем лишние пробелы
        text = text.lower()
        logging.debug(f"Нормализованный текст (первые 200 символов): {text[:200]}")
        
        # Ищем фамилии разными способами
        for surname in surnames:
            surname_lower = surname.lower()
            variations = [
                surname_lower,                    # Точное совпадение
                f" {surname_lower} ",            # С пробелами
                f"{surname_lower}.",             # С точкой
                f"{surname_lower},",             # С запятой
                f"/{surname_lower}/",            # В слэшах
                f"/{surname_lower}.",            # Слэш с точкой
                f"\n{surname_lower}\n",          # С переносами строк
                f" {surname_lower}\n",           # Пробел и перенос
                f"{surname_lower}:",             # С двоеточием
            ]
            
            for variant in variations:
                if variant in text:
                    logging.info(f"Найдена фамилия {surname} в варианте '{variant}'")
                    logging.debug(f"Контекст находки: {text[text.find(variant)-50:text.find(variant)+50]}")
                    return surname
        
        logging.warning(f"Фамилии не найдены в документе: {doc_path}")
        return "не найдено"
        
    except Exception as e:
        logging.error(f"Ошибка при поиске фамилий в {doc_path}: {e}")
        return "не найдено"

def process_line(line: str, surnames: List[str]) -> Optional[Tuple]:
    """
    Обработка одной строки из входного файла.
    
    Args:
        line: строка из входного файла
        surnames: список фамилий для поиска
        
    Returns:
        Optional[Tuple]: кортеж с извлеченными данными или None в случае ошибки
        Формат кортежа: (profile_trace, candidate_hit, relationship, lr, exclusions, overlapping_markers, shared_alleles)
        
    Example:
        >>> process_line("123-23\t456-23\t...\t0.5\t...", surnames_list)
        ('123-23', '456-23', 'Parent/Child', '0.5', '0', '15', '95.0')
    """
    logging.debug(f"Обработка строки: {line.strip()}")
    
    try:
        # Разбиваем строку на части по табуляции
        parts = line.strip().split('\t')
        
        # Проверяем количество полей
        if len(parts) < 12:  # Изменили минимальное количество полей на 12
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

def process_mismatch(profile_trace: str, candidate_hit: str, archive_base: str) -> Tuple[str, str]:
    """
    Обработка пары несовпадающих заключений и поиск фамилий в обоих документах.
    
    Args:
        profile_trace: номер первого заключения
        candidate_hit: номер второго заключения
        archive_base: базовый путь к архиву
        
    Returns:
        Tuple[str, str]: пара строк с результатами поиска
        Формат: (результат_для_profile, результат_для_candidate)
        
    Example:
        >>> process_mismatch("123-23", "456-23", "path/to/archive")
        ("123-23 Иванов", "456-23 Петров")
    """
    logging.info(f"Обработка несовпадения между {profile_trace} и {candidate_hit}")
    
    # Обработка первого заключения (profile_trace)
    logging.debug(f"Начало обработки первого заключения: {profile_trace}")
    num_profile = extract_number(profile_trace)
    year_profile = extract_last_two_digits(profile_trace)
    profile_surname = "не найдено"
    
    if year_profile:
        archive_path = archive_base + year_profile
        logging.debug(f"Путь к архиву для первого заключения: {archive_path}")
        
        if os.path.exists(archive_path):
            doc_path = get_document_path(archive_path, num_profile, year_profile)
            if doc_path:
                profile_surname = find_surnames(doc_path, CONFIG['surnames'])
                logging.info(f"Для заключения {profile_trace} найдена фамилия: {profile_surname}")
            else:
                logging.warning(f"Не найден документ для заключения {profile_trace}")
        else:
            logging.warning(f"Не найден путь к архиву: {archive_path}")
    else:
        logging.warning(f"Не удалось извлечь год из номера {profile_trace}")
    
    # Обработка второго заключения (candidate_hit)
    logging.debug(f"Начало обработки второго заключения: {candidate_hit}")
    num_candidate = extract_number(candidate_hit)
    year_candidate = extract_last_two_digits(candidate_hit)
    candidate_surname = "не найдено"
    
    if year_candidate:
        archive_path = archive_base + year_candidate
        logging.debug(f"Путь к архиву для второго заключения: {archive_path}")
        
        if os.path.exists(archive_path):
            doc_path = get_document_path(archive_path, num_candidate, year_candidate)
            if doc_path:
                candidate_surname = find_surnames(doc_path, CONFIG['surnames'])
                logging.info(f"Для заключения {candidate_hit} найдена фамилия: {candidate_surname}")
            else:
                logging.warning(f"Не найден документ для заключения {candidate_hit}")
        else:
            logging.warning(f"Не найден путь к архиву: {archive_path}")
    else:
        logging.warning(f"Не удалось извлечь год из номера {candidate_hit}")
    
    # Формирование результатов
    profile_result = f"{profile_trace} {profile_surname}"
    candidate_result = f"{candidate_hit} {candidate_surname}"
    
    logging.debug(f"Сформированы результаты:\n"
                 f"- Первое заключение: {profile_result}\n"
                 f"- Второе заключение: {candidate_result}")
    
    return profile_result, candidate_result

def main():
    """
    Основная функция программы.
    Выполняет следующие шаги:
    1. Читает входной файл
    2. Находит несовпадающие записи
    3. По запросу пользователя ищет фамилии в архиве или просто выводит список несовпадений
    4. Сохраняет результаты в выходной файл
    """
    logging.info("Начало работы программы")

    # Формируем имя выходного файла с текущей датой
    current_date = datetime.now().strftime('%d.%m.%Y')
    output_file_name = f"{current_date} несовпадения.txt"
    
    # Получаем путь к директории для сохранения выходного файла
    output_directory = CONFIG.get('output_directory', '.')  # Если путь не указан, используем текущую директорию
    
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
        result = process_line(line, CONFIG['surnames'])
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

    # Обработка дубликатов и выбор Direct-match
    # Сначала удаляем полные дубликаты
    unique_lines = []
    seen = set()
    for entry in mismatched_lines:
        # entry[1:] - берем все элементы, кроме исходной строки (line)
        # чтобы избежать хеширования строки, используем только обработанные данные
        key = tuple(entry[1:])  # (profile_trace, candidate_hit, relationship, ...)
        if key not in seen:
            seen.add(key)
            unique_lines.append(entry)
    logging.info(f"Удалено дубликатов: {len(mismatched_lines) - len(unique_lines)}")

    # Теперь выбираем строки с Direct-match вместо Parent-child
    # Создаем словарь: ключи — (profile_trace, candidate_hit), значения — лучшая строка
    prioritized_lines = {}
    for entry in unique_lines:
        line, profile, candidate, rel, lr, exclusions, overlapping, shared = entry
        key = (profile, candidate)
        current_entry = {
            "key": key,
            "rel": rel,
            "entry": entry
        }

        # Если ключ уже есть в словаре
        if key in prioritized_lines:
            existing_rel = prioritized_lines[key]["rel"]
            # Если текущая строка имеет Direct-match, а существующая — Parent-child → заменяем
            if (rel == "Direct-match" and existing_rel == "Parent-child"):
                prioritized_lines[key] = current_entry
            # Иначе оставляем существующую (например, обе Direct-match)
        else:
            prioritized_lines[key] = current_entry

    # Преобразуем словарь обратно в список
    filtered_lines = [v["entry"] for v in prioritized_lines.values()]
    logging.info(f"После выбора Direct-match: {len(filtered_lines)} строк")

    # Заменяем mismatched_lines на отфильтрованный список
    mismatched_lines = filtered_lines

    # Запрос на поиск в архиве
    search_archive = input("Произвести поиск в архиве? (да/нет): ").strip().lower()

    try:
        with open(output_file_path, 'w', encoding='utf-8') as output_file:  # Используем полный путь к файлу
            if search_archive == 'да':
                logging.info("Начало поиска в архиве")
                print("Начинаем поиск в архиве...")

                for idx, (_, profile_trace, candidate_hit, relationship, lr, exclusions, overlapping_markers, shared_alleles) in enumerate(
                    tqdm(mismatched_lines, desc="Поиск в архиве"), start=1
                ):
                    logging.debug(f"Обработка пары {idx}/{len(mismatched_lines)}")

                    profile_result, candidate_result = process_mismatch(
                        profile_trace, 
                        candidate_hit, 
                        CONFIG['archive_base']
                    )

                    # Формируем и записываем строку результата
                    output_line = (f"{profile_result}, {candidate_result}, "
                                 f"{relationship}, {lr}, {exclusions}, "
                                 f"{overlapping_markers}, {shared_alleles}\n\n")
                    output_file.write(output_line)
                    logging.debug(f"Записана строка: {output_line.strip()}")
            else:
                logging.info("Поиск в архиве отменен пользователем")
                print("Поиск в архиве отменен. Вывод списка несовпадений...")

                for idx, (line, profile_trace, candidate_hit, relationship, lr, exclusions, overlapping_markers, shared_alleles) in enumerate(
                    mismatched_lines, start=1
                ):
                    output_line = (f"{profile_trace}, {candidate_hit}, "
                                 f"{relationship}, {lr}, {exclusions}, "
                                 f"{overlapping_markers}, {shared_alleles}\n\n")
                    output_file.write(output_line)
                    logging.debug(f"Записана строка без поиска в архиве: {output_line.strip()}")

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