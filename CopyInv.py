import os
import shutil
import re


# Функция для получения пути из файла path.txt
def get_paths_from_txt():
    with open('path.txt', 'r') as file:
        paths = file.readlines()
    
    source = paths[0].strip()  # Первая строка - исходная папка
    target = paths[1].strip()  # Вторая строка - целевая папка
    invoice_range = paths[2].strip() if len(paths) > 2 else None  # Третья строка - диапазон
    
    return source, target, invoice_range


# Функция для получения диапазона номеров инвойсов
def get_invoice_range(invoice_range):
    if invoice_range:
        match = re.match(r"^(\d+)-(\d+)$", invoice_range)
        if match:
            start, end = int(match.group(1)), int(match.group(2))
            return set(range(start, end + 1))
    return None


# Функция для копирования файлов
def copy_invoice_files(source_dir, target_base_dir, template_dir, invoice_range):
    # Путь к файлу .xlsm в папке Template (предполагается, что он один)
    xlsm_file = None
    for file in os.listdir(template_dir):
        if file.lower().endswith(".xlsm"):
            xlsm_file = file
            break  # Прерываем цикл после нахождения первого xlsm файла

    if not xlsm_file:
        print(f'Ошибка: .xlsm файл не найден в папке {template_dir}')
        return

    # Полный путь к .xlsm файлу
    xlsm_file_path = os.path.join(template_dir, xlsm_file)

    # Регулярное выражение для поиска инвойсов в формате "invoice ЧИСЛО.xlsx"
    invoice_pattern = re.compile(r"^invoice (\d+)\.xlsx$", re.IGNORECASE)

    # Обрабатываем все папки в исходной директории
    for root, dirs, files in os.walk(source_dir):
        for file in files:
            match = invoice_pattern.match(file)  # Проверяем название файла
            if match:
                invoice_number = int(match.group(1))  # Извлекаем число из названия
                
                # Проверяем, входит ли номер в указанный диапазон (если он задан)
                if invoice_range and invoice_number not in invoice_range:
                    continue  # Пропускаем файлы, которые не входят в диапазон

                # Получаем имя текущей папки
                folder_name = root.split(os.sep)[-1]
                folder_parts = folder_name.split(',')

                # Проверяем, что в папке есть хотя бы 4 части (иначе пропускаем)
                if len(folder_parts) >= 4:
                    app_name = folder_parts[3].strip()  # Извлекаем название приложения (4-я часть после запятой)

                    # Формируем путь к целевой папке
                    target_folder = os.path.join(target_base_dir, app_name)
                    os.makedirs(target_folder, exist_ok=True)  # Создаём папку, если её нет

                    # Копируем файл инвойса
                    source_file_path = os.path.join(root, file)
                    target_file_path = os.path.join(target_folder, file)
                    shutil.copy(source_file_path, target_file_path)
                    print(f'✅ Файл {file} скопирован в {target_file_path}')

                    # Проверяем, есть ли уже .xlsm файл в целевой папке
                    xlsm_exists = any(f.lower().endswith(".xlsm") for f in os.listdir(target_folder))
                    
                    if not xlsm_exists:
                        target_xlsm_file_path = os.path.join(target_folder, xlsm_file)
                        shutil.copy(xlsm_file_path, target_xlsm_file_path)
                        print(f'✅ Файл {xlsm_file} скопирован в {target_folder}')
                    else:
                        print(f'⚠️ Внимание: В папке {target_folder} уже есть .xlsm файл. Копирование пропущено.')

                else:
                    print(f'⚠️ Ошибка: папка "{folder_name}" не содержит достаточно частей для извлечения приложения.')


if __name__ == '__main__':
    # Получаем пути и диапазон из path.txt
    source_directory, target_base_directory, invoice_range_str = get_paths_from_txt()
    
    # Преобразуем диапазон номеров в множество
    invoice_range_set = get_invoice_range(invoice_range_str)

    # Папка с шаблоном (ищем её в той же директории, где находится скрипт)
    template_directory = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Template')

    # Проверяем, существует ли папка с шаблоном
    if not os.path.exists(template_directory):
        print(f'❌ Ошибка: папка Template не найдена в {template_directory}')
    else:
        # Копируем файлы инвойсов и шаблон
        copy_invoice_files(source_directory, target_base_directory, template_directory, invoice_range_set)
