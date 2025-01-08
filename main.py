import pandas as pd

from collections import defaultdict
from datetime import datetime
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape

FOUNDED_YEAR = 1920


def goda_or_let(age):
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
    age = current_year - FOUNDED_YEAR
    return age


def group_by_category(products_excel_data):
    categories = defaultdict(list)
    for product in products_excel_data.to_dict("records"):
        category = product.pop("Категория")

        if pd.isna(product.get("Сорт")):
            product["Сорт"] = None

        categories[category].append(product)
    return categories


def read_excel(file_path):
    return pd.read_excel(
        io=file_path,
        sheet_name='Лист1',
        na_values='',
        keep_default_na=False
    )


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    wine_excel_data = read_excel('wine3.xlsx')
    wine_info = group_by_category(wine_excel_data)

    age = get_age()
    word = goda_or_let(age)

    rendered_page = template.render(
        age=age,
        word=word,
        wine_info=wine_info
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
