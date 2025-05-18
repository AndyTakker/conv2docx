import sys
import os
import pypandoc
from importlib.metadata import version
import yaml
import json

def convert_yaml_to_json(input_file):
    """
    Конвертирует .yaml или .yml файл в формат JSON и сохраняет рядом с исходным файлом.

    :param input_file: Путь к .yaml/.yml файлу.
    :return: Путь к созданному JSON-файлу или None при ошибке.
    """
    base_name = os.path.splitext(input_file)[0]
    # json_file = f"{base_name}.json"
    json_file = f"{input_file}.json"

    try:
        with open(input_file, "r", encoding="utf-8") as src:
            data = yaml.safe_load(src)  # Загружаем YAML безопасно

        with open(json_file, "w", encoding="utf-8") as dst:
            json.dump(data, dst, indent=4, ensure_ascii=False)

        print(f"Создан временный JSON-файл: {json_file}")
        return json_file

    except Exception as e:
        print(f"Ошибка при конвертации YAML в JSON: {e}")
        return None

def prepare_markdown_file(input_file):
    """
    Создаёт временный .md файл на основе переданного JSON-файла.
    Добавляет маркеры ```json в начало и конец файла.

    :param input_file: Путь к исходному JSON-файлу.
    :return: Путь к созданному .md файлу или None, если произошла ошибка.
    """
    # Проверяем, существует ли исходный файл
    if not os.path.isfile(input_file):
        print(f"Файл {input_file} не найден.")
        return None

    # Формируем имя нового .md файла
    md_file = input_file + ".md"

    try:
        # Читаем содержимое оригинального файла
        with open(input_file, 'r', encoding='utf-8') as src:
            content = src.read()

        # Добавляем маркеры JSON в начало и конец
        new_content = "```json\n" + content + "\n```"

        # Записываем в новый .md файл
        with open(md_file, 'w', encoding='utf-8') as dst:
            dst.write(new_content)

        print(f"Создан временный файл: {md_file}")
        return md_file

    except Exception as e:
        print(f"Ошибка при подготовке файла: {e}")
        return None


def convert_md_to_docx(input_file):
    """
    Конвертирует указанный Markdown (.md) файл в формат .docx с помощью pypandoc.

    :param input_file: Путь к .md файлу.
    :return: True, если конвертация прошла успешно, иначе False.
    """

    base_name = os.path.splitext(os.path.basename(input_file))[0]
    output_file = f"{base_name}.docx"

    try:
        output = pypandoc.convert_file(input_file, 'docx', outputfile=output_file)
        print(f"Файл успешно сохранён как {output_file}")
        return True  # Успешно завершено
    except Exception as e:
        print(f"Ошибка при конвертации: {e}")
        return False  # Ошибка

def main():
    """
    Основное тело программы:
    
    - Если указан аргумент командной строки — обрабатывается один файл.
    - Если аргументов нет — обрабатываются все .json файлы в текущей директории.
    - Для каждого JSON-файла создаётся временный .md файл, выполняется конвертация,
      после чего временный файл удаляется.
    - Аргумент --keep-temp сохраняет временные файлы.  
    """
    print(version('conv2docx'))  # выведем версию модуля

    # Парсим аргументы
    args = sys.argv[1:]
    keep_temp = '--keep-temp' in args

    # Убираем служебные аргументы из списка файлов
    input_files = [f for f in args if f != '--keep-temp']

    # Если нет аргументов — берем все JSON/YAML файлы в текущей папке
    if not input_files:
        input_files = [f for f in os.listdir() if f.lower().endswith(('.json', '.yaml', '.yml'))]

    if not input_files:
        print("Нет подходящих файлов для обработки (.json, .yaml, .yml).")
        sys.exit(1)

    for file in input_files:
        print(f"\nОбработка файла: {file}")

        ext = os.path.splitext(file)[1].lower()

        # Шаг 1: Если это YAML/YML — конвертируем в JSON
        json_file = None
        if ext in ('.yaml', '.yml'):
            json_file = convert_yaml_to_json(file)
            if not json_file:
                print(f"Не удалось конвертировать {file} в JSON, пропускаем.")
                continue
        else:
            json_file = file

        # Шаг 2: Подготовить .md файл
        md_file = prepare_markdown_file(json_file)
        if not md_file:
            print(f"Не удалось подготовить файл {json_file}, пропускаем.")
            continue

        # Шаг 3: Конвертировать .md в .docx
        success = convert_md_to_docx(md_file)

        # Шаг 4: Удалить временные файлы (если не указан флаг --keep-temp)
        to_remove = []
        if json_file != file:  # удалить только если он был создан
            to_remove.append(json_file)
        if md_file:
            to_remove.append(md_file)

        if keep_temp:
            print(f"Временные файлы сохранены: {', '.join(to_remove)}")
        else:
            for temp_file in to_remove:
                if os.path.exists(temp_file):
                    try:
                        os.remove(temp_file)
                        print(f"Временный файл {temp_file} удалён.")
                    except Exception as e:
                        print(f"Не удалось удалить временный файл {temp_file}: {e}")

        if not success:
            print(f"Ошибка при обработке файла {file}")

if __name__ == "__main__":
    print(version('conv2docx'))  # выведем версию модуля
    main()
 