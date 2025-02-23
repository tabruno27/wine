from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
from datetime import datetime
import pandas as pd
from collections import defaultdict, OrderedDict

def get_year_word(n):
    if 11 <= n % 100 <= 14:
        return "лет"
    elif n % 10 == 1:
        return "год"
    elif n % 10 in [2, 3, 4]:
        return "года"
    else:
        return "лет"


def main():
    CREAT_YEAR = 1920
    delta = datetime.now().year - CREAT_YEAR
    
    parser = argparse.ArgumentParser(description='Обработка данных из Excel.')
    parser.add_argument('--data_path', type=str, default='Production.xlsx', help='Путь к файлу с данными')

    args = parser.parse_args()

    excel_data_df = pd.read_excel(args.data_path, sheet_name='Лист1', na_values=['N/A', 'NA'], keep_default_na=False)
    products = excel_data_df.to_dict(orient='records')

    year_word = get_year_word(delta)

    new_products = defaultdict(list)

    for product in products:
        category = product["Категория"]
        wine = {
            "Картинка": product["Картинка"],
            "Категория": product["Категория"],
            "Название": product["Название"],
            "Сорт": product["Сорт"] if pd.notna(product["Сорт"]) else "",
            "Цена": product["Цена"],
            "Акция": product["Акция"] if pd.notna(product["Акция"]) else ""
        }
        new_products[category].append(wine)
    
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    
    template = env.get_template('template.html')
    
    rendered_page = template.render(
        cap_title=f"Уже {delta} {year_word} с вами",
        new_products=new_products,
    )
    
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)
    
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()

if __main__=="__main__"
    main()
