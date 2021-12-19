from ctypes import string_at
import time
import json
import base64
from math import fabs
import requests
import xlsxwriter
from bs4 import BeautifulSoup   

def parse(url):
    # Сохраним исходный URL
    start_url = url

    # Создаём XLSX файл и лист в нём
    workbook = xlsxwriter.Workbook('result.xlsx')
    worksheet = workbook.add_worksheet()
    
    worksheet.write('A1', 'Имя')
    worksheet.write('B1', 'Телефон')
    worksheet.write('C1', 'Адрес')
    worksheet.write('D1', 'Текст объявления')
    worksheet.write('E1', 'Регион')

    # Индекс следующей строки в Excel
    row_index = 1

    # Номер страницы
    page_number = 1

    # Есть ли ещё страницы
    nextPageAvailable = True

    # Если есть следующая страница
    while(nextPageAvailable):
        
        html = requests.get(url=url, headers = {'User-agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/64.0.3282.186 Safari/537.36'})
        print(html.status_code)

        # Если страница получена
        if(html.status_code == 200):
            soup = BeautifulSoup(html.text, "html.parser")

            # Получаем родитель всех объявлений
            soup.find_all("div", class_="items-items-kAJAg")[0]
            # Получаем все объявления
            items = soup.findChildren("div", attrs={"data-marker": "item"})

            # Перебираем объявления
            for item in items:
                # Увеличиваем индекс строчки, в которую пишемся
                row_index += 1
                
                # Получаем ID объявления и ссылку на него
                item_id = item["data-item-id"]
                print(item_id)
                item_url = "https://www.avito.ru" + item.findChild("a", attrs={"data-marker": "item-title"})["href"]

                # По ссылке получаем имя, адрес и текст объявления
                item_page_soup = BeautifulSoup(requests.get(url=item_url).text, "html.parser")
                item_name = item_page_soup.find("div", attrs={"data-marker": "seller-info/name"}).text
                item_address = item_page_soup.find("div", attrs={"itemprop": "address"}).text
                item_text = item_page_soup.find("div", attrs={"itemprop": "description"}).text

                # Номер телефона как картинка
                item_phone_img_bs64 = json.loads(requests.get(url="https://www.avito.ru/items/phone/" + item_id).text)["image64"][22:]
                item_phone_img_path = "images/id"+item_id+".png"
                with open(item_phone_img_path, "wb") as fh:
                    fh.write(base64.decodebytes(item_phone_img_bs64.encode()))
                    fh.close()

                worksheet.write('A'+row_index, item_name)
                worksheet.insert_image('B'+row_index, item_phone_img_path)
                worksheet.write('C'+row_index, item_address)
                worksheet.write('D'+row_index, item_text)
                worksheet.write('E'+row_index, 'Регион')

                time.sleep(2)
            
            # Если больше нет - цикл закрывается
            if soup.find("span", attrs={"data-marker": "pagination-button/next"}) == None:
                nextPageAvailable = False
            else:
                page_number += 1;
                url = start_url + "&p=" + page_number
        else:
            print("Error parse:" + url)
            time.sleep(60)

        time.sleep(2)

    workbook.close()

parse(url=r'https://www.avito.ru/rossiya?q=продажа+евровагонки')