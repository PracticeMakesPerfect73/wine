from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape
import datetime
import pandas
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
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    current_year = datetime.datetime.now().year
    years_since_opening = current_year - WINERY_OPENING_YEAR
    years_word = choose_year_word(years_since_opening)

    excel_data_df = pandas.read_excel('wine.xlsx',
                                      sheet_name='Лист1').fillna('')
    wines_collection = excel_data_df.to_dict(orient='records')
    grouped_wines = defaultdict(list)
    for wine in wines_collection:
        category = wine.get('Категория', '').strip()
        grouped_wines[category].append(wine)

    rendered_page = template.render(
        wines=grouped_wines,
        feedback_sub_title=f'Уже {years_since_opening} {years_word} с нами',
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == "__main__":
    main()
