import argparse
import os
from collections import defaultdict
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler

import pandas as pd
from jinja2 import Environment, FileSystemLoader, select_autoescape

FOUNDATION_YEAR = 1920


def get_year_word(age):
    if age < 10:
        if age == 1:
            return 'год'
        elif 1 < age < 5:
            return 'года'
        else:
            return 'лет'

    if 10 <= age % 100 <= 20:
        return 'лет'

    last_digit = age % 10
    if last_digit == 1:
        return 'год'
    elif 1 < last_digit < 5:
        return 'года'
    else:
        return 'лет'


def get_age():
    current_year = datetime.now().year
    age = current_year - FOUNDATION_YEAR
    return age


def fill_missing_value(product, field, default_value=None):
    if pd.isna(product.get(field)):
        product[field] = default_value


def group_by_category(products_excel_data):
    categories = defaultdict(list)

    for product in products_excel_data.to_dict("records"):
        category = product.pop("Категория")

        for field in product:
            fill_missing_value(product, field)

        categories[category].append(product)

    return categories


def read_excel(file_path, sheet_name='Лист1'):
    return pd.read_excel(
        io=file_path,
        sheet_name=sheet_name,
        na_values='',
        keep_default_na=False
    )


def main():
    parser = argparse.ArgumentParser(
        description="Запуск программы с указанием пути к файлу данных"
    )
    parser.add_argument(
        "--file",
        default=os.getenv(
            "EXCEL_FILE_PATH",
            "wine_and_drinks_catalog.xlsx"
        ),
        help="Путь к файлу Excel с данными (по умолчанию 'wine_and_drinks_catalog.xlsx')"
    )
    parser.add_argument(
        "--template",
        default=os.getenv("TEMPLATE_PATH", "template.html"),
        help="Путь к шаблону HTML для рендеринга (по умолчанию 'template.html')"
    )
    parser.add_argument(
        "--port",
        type=int,
        default=int(os.getenv("PORT", 8000)),
        help="Порт для HTTP сервера (по умолчанию 8000)"
    )
    args = parser.parse_args()

    try:
        env = Environment(
            loader=FileSystemLoader('.'),
            autoescape=select_autoescape(['html', 'xml'])
        )
        template = env.get_template('template.html')

        excel_file_path = args.file
        try:
            wine_and_drinks_df = read_excel(excel_file_path)
        except Exception as e:
            print(f"Ошибка при чтении файла '{excel_file_path}': {str(e)}")
            return

        wine_and_drinks_by_category = group_by_category(wine_and_drinks_df)

        age = get_age()
        word = get_year_word(age)

        rendered_page = template.render(
            age=age,
            word=word,
            wine_and_drinks_by_category=wine_and_drinks_by_category
        )

        with open('index.html', 'w', encoding="utf8") as file:
            file.write(rendered_page)

        server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
        server.serve_forever()

    except Exception as e:
        print(f"Произошла ошибка: {str(e)}")


if __name__ == '__main__':
    main()
