import os
import shutil


# Функция для получения пути из файла path.txt
def get_paths_from_txt():
    with open('path.txt', 'r') as file:
        paths = file.readlines()
    return paths[0].strip(), paths[1].strip()  # Первая строка - исходная папка, вторая - целевая


# Функция для перемещения файлов
def move_invoice_files(source_dir, target_base_dir, template_dir):
    # Путь к файлу .xlsm в папке Template (предполагается, что он один)
    xlsm_file = None
    for file in os.listdir(template_dir):
        if file.lower().endswith(".xlsm"):
            xlsm_file = file
            break  # Прерываем цикл после нахождения первого xlsm файла

    if not xlsm_file:
        print(f'Ошибка: .xlsm файл не найден в папке {template_dir}')
        return

    # Путь к полному файлу шаблона .xlsm
    xlsm_file_path = os.path.join(template_dir, xlsm_file)

    # Обрабатываем все папки в исходной директории
    for root, dirs, files in os.walk(source_dir):
        for file in files:
            # Проверяем, является ли файл инвойсом (Invoice НОМЕР.xlsx)
            if file.lower().startswith("invoice") and file.lower().endswith(".xlsx"):
                invoice_number = file.split()[1].replace('.xlsx', '')  # Извлекаем номер инвойса
                # Ищем папку, в которой будет храниться файл
                folder_name = root.split(os.sep)[-1]  # Извлекаем имя текущей папки
                folder_parts = folder_name.split(',')

                # Проверяем, что в папке есть хотя бы 4 части (если нет, пропускаем)
                if len(folder_parts) >= 4:
                    app_name = folder_parts[3].strip()  # Извлекаем название приложения (4-я часть после запятой)

                    # Формируем путь к целевой папке
                    target_folder = os.path.join(target_base_dir, app_name)
                    if not os.path.exists(target_folder):
                        os.makedirs(target_folder)  # Если папка не существует, создаем ее

                    # Путь для копирования файла инвойса
                    source_file_path = os.path.join(root, file)
                    target_file_path = os.path.join(target_folder, file)

                    # Копируем файл инвойса, а не перемещаем
                    shutil.copy(source_file_path, target_file_path)
                    print(f'Файл {file} скопирован в {target_file_path}')

                    # Копируем .xlsm файл из папки Template в папку приложения
                    target_xlsm_file_path = os.path.join(target_folder, xlsm_file)
                    shutil.copy(xlsm_file_path, target_xlsm_file_path)
                    print(f'Файл {xlsm_file} скопирован в {target_folder}')

                else:
                    print(f'Ошибка: папка "{folder_name}" не содержит достаточно частей для извлечения приложения.')


if __name__ == '__main__':
    # Получаем пути из файла path.txt
    source_directory, target_base_directory = get_paths_from_txt()

    # Папка с шаблоном (ищем её в той же директории, где находится скрипт)
    template_directory = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Template')

    # Проверяем, существует ли папка с шаблоном
    if not os.path.exists(template_directory):
        print(f'Ошибка: папка Template не найдена в {template_directory}')
    else:
        # Перемещаем файлы и копируем шаблон
        move_invoice_files(source_directory, target_base_directory, template_directory)
