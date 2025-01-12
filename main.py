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
    return current_year - FOUNDATION_YEAR


def group_by_category(products_excel_data):
    categories = defaultdict(list)

    for product in products_excel_data.to_dict("records"):
        category = product.pop("Категория")
        product = {
            field: value if not pd.isna(value) else None
            for field, value in product.items()
        }
        categories[category].append(product)

    return categories


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Скрипт для генерации HTML-страницы магазина авторского вина "
            "'Новое русское вино' из Excel-файла."
        )
    )
    parser.add_argument(
        "--file",
        default=os.getenv(
            "EXCEL_FILE_PATH",
            "wine_and_drinks_catalog.xlsx"
        ),
        help=(
            "Путь к файлу Excel с данными. "
            "По умолчанию: 'wine_and_drinks_catalog.xlsx'."
        ),
    )
    parser.add_argument(
        "--template",
        default=os.getenv(
            "TEMPLATE_PATH",
            "template.html"
        ),
        help=(
            "Путь к шаблону HTML для рендеринга. "
            "По умолчанию: 'template.html'."
        ),
    )
    parser.add_argument(
        "--port",
        type=int,
        default=int(os.getenv("PORT", 8000)),
        help="Порт для HTTP сервера. По умолчанию: 8000.",
    )
    args = parser.parse_args()

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template(args.template)

    wine_and_drinks_df = pd.read_excel(
        io=args.file,
        sheet_name='Лист1',
        na_values='',
        keep_default_na=False
    )

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

    server = HTTPServer(('0.0.0.0', args.port), SimpleHTTPRequestHandler)
    print(f"Сервер запущен на порту {args.port}...")
    server.serve_forever()


if __name__ == '__main__':
    main()
