# /////////       Importing         ///////////
import timeit
import math

import openpyxl
from xlsxwriter import Workbook
from selenium import webdriver

import time

from PIL import Image
from io import BytesIO

from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import TimeoutException


from selenium.webdriver.chrome.options import Options

# Настройка опций браузера
chrome_options = Options()
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome_options.add_argument('--headless')
chrome_options.add_argument('--log-level=3')

# //////////      Mian programm      //////////

# Ввод ссылки на поиск
url = input('Введите ссылку на поиск в формате \n "https://www.wildberries.ru/catalog/0/search.aspx?page=*Номер страницы*&sort=*метод сортировки*&search=*запрос*"\n: ')

# Обработка решения о скришотах
imgAnswer = ''
imgXLSXAnswer = ''
while imgAnswer != 'да' or 'нет':
    if imgAnswer == 'да':
        imgstate = True
        break
    elif imgAnswer == 'нет':
        imgstate = False
        break
    else:
        imgAnswer = input("Нужно ли сохранять скриншоты карточек? да/нет : ")

if imgstate == True:
    while imgXLSXAnswer != 'да' or 'нет':
        if imgXLSXAnswer == 'да':
            imgXLSXstate = True
            break
        elif imgXLSXAnswer == 'нет':
            imgXLSXstate = False
            break
        else:
            imgXLSXAnswer = input("Нужно ли сохранять скриншоты карточек в xlsx файл? да/нет : ")
else:
    imgXLSXstate = False



class WBParser(object):

    # Конструктор
    def __init__(self, driver):
        self.driver = driver

    # Функция модулей
    def parse(self):
        self.info_finder(url)
        self.save_exel()

    # Парсинг страницы
    def info_finder(self, url):

        # определние списков для информации
        self.productBrandNameList = []
        self.productLinkList = []
        self.productNameList = []
        self.productIdList = []
        self.productPriceList = []

        self.start = timeit.default_timer()

        self.driver.get(url)
        time.sleep(5)

        productCount = self.driver.find_element(By.XPATH, "/html/body/div[1]/main/div[2]/div/div/div[1]/div[1]/div/span/span[1]")
        productCount = productCount.text.replace(" ", "")
        pageCount = math.ceil(int(productCount)/100)

        for page in range(1, pageCount + 1):

            print('Парсинг страницы: ', page)

            startPageTime = timeit.default_timer()

            #ожидание загрузки карточек
            time.sleep(1)

            # поиск необходимых данных на странице
            product_href = self.driver.find_elements(By.CLASS_NAME, "product-card__main.j-card-link")
            product_brands = self.driver.find_elements(By.CLASS_NAME, "brand-name")
            product_name = self.driver.find_elements(By.CLASS_NAME, "goods-name")
            product_id = self.driver.find_elements(By.CLASS_NAME, "product-card.j-card-item")
            product_price = self.driver.find_elements(By.CLASS_NAME, "lower-price")

            # запись названия бренда
            for elem in product_brands:
                self.productBrandNameList.append(elem.text)
            
            counter = 0
            scrollRange = 0
            for elem in product_id:
                
                # запись артикула карточки
                self.productIdList.append(elem.get_attribute('data-popup-nm-id'))

                # скриншот карточек товара
                if imgstate == True:

                    location = elem.location
                    size = elem.size
                    png = self.driver.get_screenshot_as_png()

                    im = Image.open(BytesIO(png))

                    left = location['x']
                    top = 230
                    right = location['x'] + size['width']
                    bottom = 230 + size['height']

                    im = im.crop((left, top, right, bottom))
                    im.save('images/' + elem.get_attribute('data-popup-nm-id') + '.png')
                    counter += 1

                    if counter%4 == 0:
                        scrollRange = scrollRange + size["height"] + 24
                        self.driver.execute_script("window.scrollTo(0, " + str(scrollRange) + ")") 

            counter = 0
            scrollRange = 485

            #запись названия товара
            for elem in product_name:
                self.productNameList.append(elem.text)
            
            #запись ссылки на карточку
            for elem in product_href:
                self.productLinkList.append(elem.get_attribute('href'))

            #запись ссылки на товар
            for elem in product_price:
                self.productPriceList.append(((elem.text).replace(" ", "")).replace("₽", ""))
                
            # скролл страницы в низ для переключения на слудующую
            if  page > 2:
                self.driver.execute_script("window.scrollTo(0, 12200)")
            elif page > 10:
                self.driver.execute_script("window.scrollTo(0, 12000)")
            else:
                self.driver.execute_script("window.scrollTo(0, 12700)")


            
            #игнорирование ошибки загрузки элемента переключения на следующую страницу
            ignored_exceptions=(NoSuchElementException,StaleElementReferenceException, ElementClickInterceptedException)

            #обработка ошибки последней страницы поиска

            try:
                linkToNextPage = WebDriverWait(self.driver, 2,ignored_exceptions=ignored_exceptions).until(expected_conditions.presence_of_element_located((By.CLASS_NAME, "pagination-next.pagination__next")))    
            except TimeoutException:
                continue
            
            # переход на слелдующую страницу
            linkToNextPage.click()

            stopPageTime = timeit.default_timer()
            executionPagetime = stopPageTime - startPageTime
            print("Общее время обработки страницы поискового запроса: ", executionPagetime, " секунд")


        self.productBrandNameList = [sub[ : -2] for sub in self.productBrandNameList]
        
    # сохранение данных в эксель
    def save_exel(self):
        workbook = Workbook('data.xlsx')
        Report_Sheet = workbook.add_worksheet()

        # разметка столбцов
        Report_Sheet.write(0, 0, 'id карточки')
        Report_Sheet.write(0, 1, 'бренд')
        Report_Sheet.write(0, 2, 'название товара')
        Report_Sheet.write(0, 3, 'Цена')
        Report_Sheet.write(0, 4, 'ссылка на товар')
        Report_Sheet.write(0, 5, 'Картинка карточки')
        
        # запись столбцов
        Report_Sheet.write_column(1, 0, self.productIdList)
        Report_Sheet.write_column(1, 1, self.productBrandNameList)
        Report_Sheet.write_column(1, 2, self.productNameList)
        Report_Sheet.write_column(1, 3, self.productPriceList)
        Report_Sheet.write_column(1, 4, self.productLinkList)

        workbook.close()

        # Запись картинок в xlsx файл
        if imgXLSXstate == True:
            wb = openpyxl.load_workbook('data.xlsx')
            ws = wb['Sheet1']
            counter = 1
            for elem in self.productIdList:
                counter += 1
                img = openpyxl.drawing.image.Image('images/'+str(elem)+'.png')
                img.anchor = 'F'+str(counter) # Or whatever cell location you want to use.
                ws.add_image(img)

            counter = 0
            wb.save('data.xlsx')

        # Завершение таймера
        self.stop = timeit.default_timer()
        execution_time = self.stop - self.start
        print("Общее время обработки поискового запроса: ", round(execution_time), " секунд")

def main():
    driver = webdriver.Chrome()
    parser = WBParser(driver)
    parser.parse()

if __name__ == "__main__":
    main()