import os
import glob
import xml.dom.minidom
from datetime import datetime

def pretty_print_xml(file_path):
    # Парсинг XML файла
    dom = xml.dom.minidom.parse(file_path)
    # Приведение XML к красивому виду
    pretty_xml_as_string = dom.toprettyxml(indent="  ")
    return pretty_xml_as_string

def main():
    # Путь к директории с XML файлами
    xml_dir = r'U:\ГЕНЕТИКА НОВОЕ ЗДАНИЕ\BASE'
    # Получение текущей даты в формате ДД.ММ.ГГГГ
    current_date = datetime.now().strftime('%d.%m.%Y')
    # Создание новой папки с текущей датой в качестве имени
    new_folder_path = os.path.join(xml_dir, current_date)
    os.makedirs(new_folder_path, exist_ok=True)

    # Перемещение всех XML файлов в новую папку
    for xml_file in glob.glob(os.path.join(xml_dir, '*.xml')):
        os.rename(xml_file, os.path.join(new_folder_path, os.path.basename(xml_file)))

    # Имя новой папки
    folder_name = current_date
    # Установка имени выходного файла
    output_file = f'TOTAL {folder_name}.txt'
    output_file_path = os.path.join(new_folder_path, output_file)

    with open(output_file_path, 'w', encoding='utf-8') as f_out:
        for xml_file in glob.glob(os.path.join(new_folder_path, '*.xml')):
            pretty_xml = pretty_print_xml(xml_file)
            f_out.write(pretty_xml)
            f_out.write("\n\n")  # Добавление дополнительных новых строк для разделения файлов

    print(f"Извлечённая информация сохранена в {output_file_path}")

    # Путь к файлу TOTAL EXPERTISE 2025
    total_expertise_dir = os.path.join(xml_dir, 'TOTAL')
    total_expertise_file_path = os.path.join(total_expertise_dir, 'TOTAL EXPERTISE 2025.txt')  # Добавлено расширение .txt

    # Проверка существования файла TOTAL EXPERTISE 2025
    if not os.path.exists(total_expertise_file_path):
        print(f"Файл {total_expertise_file_path} не существует. Создайте его вручную.")
    else:
        # Добавление информации из нового файла в общий файл TOTAL EXPERTISE 2025
        with open(total_expertise_file_path, 'a', encoding='utf-8') as total_expertise_file:
            with open(output_file_path, 'r', encoding='utf-8') as new_data_file:
                total_expertise_file.write(new_data_file.read())
                total_expertise_file.write("\n\n")  # Добавление дополнительных новых строк для разделения данных

        print(f"Данные из {output_file_path} добавлены в {total_expertise_file_path}")

if __name__ == '__main__':
    main()