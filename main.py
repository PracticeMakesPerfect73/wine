from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape
import datetime
import pandas
import argparse
import os
from dotenv import load_dotenv
from pathlib import Path
from collections import defaultdict

WINERY_OPENING_YEAR = 1920


def choose_year_word(years_since_opening):
    if 11 <= years_since_opening % 100 <= 14:
        return 'лет'
    elif years_since_opening % 10 == 1:
        return 'год'
    elif years_since_opening % 10 in [2, 3, 4]:
        return 'года'
    else:
        return 'лет'


def main():
    load_dotenv()

    parser = argparse.ArgumentParser(description='Введите путь до файла с винами')
    parser.add_argument(
        '-f', '--file',
        help='Путь до Excel-файла',
        default=os.getenv('WINES_DEFAULT_PATH')
    )
    args = parser.parse_args()

    file_name = args.file
    if not file_name:
        print('Ошибка: путь к файлу не указан.')
        return

    file_path = Path(file_name)
    if not file_path.exists():
        print(f'Ошибка: файл "{file_path}" не найден')
        return

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    current_year = datetime.datetime.now().year
    years_since_opening = current_year - WINERY_OPENING_YEAR
    years_word = choose_year_word(years_since_opening)

    excel_data_df = pandas.read_excel(file_path,
                                      sheet_name='Лист1').fillna('')
    wines_collection = excel_data_df.to_dict(orient='records')
    grouped_wines = defaultdict(list)
    for wine in wines_collection:
        category = wine.get('Категория', '').strip()
        grouped_wines[category].append(wine)

    rendered_page = template.render(
        grouped_wines=grouped_wines,
        winery_age=f'Уже {years_since_opening} {years_word} с нами',
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == "__main__":
    main()
