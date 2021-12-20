import time
import json
import base64
import requests
import xlsxwriter
from bs4 import BeautifulSoup   
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException

options =  Options()
driver = webdriver.Chrome(options=options)
driver.set_window_size(1024, 1024)

def parse(url):
    # Сохраним исходный URL
    start_url = url

    # Создаём XLSX файл и лист в нём
    workbook = xlsxwriter.Workbook('result.xlsx')
    worksheet = workbook.add_worksheet()
    
    worksheet.set_column_pixels(0, 0, 278)
    worksheet.write('A1', 'Имя')
    worksheet.set_column_pixels(1, 1, 165)
    worksheet.write('B1', 'Телефон')
    worksheet.write('C1', 'Адрес')
    worksheet.write('D1', 'Текст объявления')

    # Индекс следующей строки в Excel
    row_index = 1

    # Номер страницы
    page_number = 1

    # Есть ли ещё страницы
    nextPageAvailable = True

    # Если есть следующая страница
    while(nextPageAvailable):
        
        driver.get(url)
        soup = BeautifulSoup(driver.page_source , "html.parser")

        # Получаем родитель всех объявлений
        if soup.find("div", class_="items-items-kAJAg") == None:
            go_get_soup = True
            while(go_get_soup):
                soup = BeautifulSoup(driver.page_source , "html.parser")
                if soup.find("div", class_="items-items-kAJAg") != None:
                    go_get_soup = False
                else:
                    time.sleep(5)
        
        # Получаем все объявления
        items = soup.findChildren("div", attrs={"data-marker": "item"})

        print(items)

        # Перебираем объявления
        for item in items:
            # Увеличиваем индекс строчки, в которую пишемся
            row_index += 1
            
            # Получаем ID объявления и ссылку на него
            item_id = item["data-item-id"]
            print(item_id)
            item_url = "https://www.avito.ru" + item.findChild("a", attrs={"data-marker": "item-title"})["href"]

            # По ссылке получаем имя, адрес и текст объявления
            driver.get(item_url)
            item_page_soup = BeautifulSoup(driver.page_source, "html.parser")
            item_name = item_page_soup.find("div", attrs={"data-marker": "seller-info/name"})

            if(item_name != None):
                item_name = item_name.text
            
            item_address = item_page_soup.find("div", attrs={"itemprop": "address"})

            if(item_address != None):
                item_address = item_address.text

            item_text = item_page_soup.find("div", attrs={"itemprop": "description"})

            if(item_text != None):
                item_text = item_text.text

            # Номер телефона как картинка
            item_phone_img_bs64 = json.loads(requests.get("https://www.avito.ru/items/phone/" + item_id).text)["image64"][22:]
            item_phone_img_path = "images/id"+item_id+".png"
            with open(item_phone_img_path, "wb") as fh:
                fh.write(base64.decodebytes(item_phone_img_bs64.encode()))
                fh.close()

            worksheet.write('A'+str(row_index), item_name)
            worksheet.insert_image('B'+str(row_index), item_phone_img_path, {'x_scale': 0.35, 'y_scale': 0.35})
            worksheet.write('C'+str(row_index), item_address)
            worksheet.write('D'+str(row_index), item_text)

            time.sleep(2)
        
        # Если больше нет - цикл закрывается
        if soup.find("span", attrs={"data-marker": "pagination-button/next"}) == None:
            nextPageAvailable = False
        else:
            page_number += 1;
            url = start_url + "&p=" + str(page_number)

        time.sleep(2)

    workbook.close()

parse(url=r'https://www.avito.ru/rossiya?q=продажа+евровагонки')