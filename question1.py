import cv2
import openpyxl
import pandas as pd
import re
import time
import urllib.request
import win32com.client as win32
from pytesseract import pytesseract, image_to_string
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException


pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
driver = webdriver.Firefox()
executor_url = driver.command_executor._url
site = "https://fssp.gov.ru"
driver.get(site)
time.sleep(3)

# Функция проверяет наличие капчи на странице


def captcha_find():
    try:
        captcha_find = driver.find_element_by_id('capchaVisual')
    except:
        captcha_find = 0
    return captcha_find

# Функция сохраняет капчу


def captcha_save(driver):
    try:
        url = driver.find_element_by_id('capchaVisual').get_attribute('src')
        name = "captcha.jpg"
        urllib.request.urlretrieve(url, name)
        return name
    except NoSuchElementException:
        time.sleep(5)
        return None

# Функция разгадывает капчу и отдает текст


def captcha_to_string(captcha):
    img = cv2.imread(captcha)
    time.sleep(2)
    gry = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    (h, w) = gry.shape[:2]
    gry = cv2.resize(gry, (w*2, h*2))
    cls = cv2.morphologyEx(gry, cv2.MORPH_CLOSE, None)
    thr = cv2.threshold(cls, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)[1]
    txt = re.sub("[^А-Яа-я0-9]", "", image_to_string(thr, lang='rus'))
    return txt

# Функция вставлчет решенную капчу в форму на сайте


def crack_captcha():
    captcha = captcha_save(driver)
    time.sleep(5)
    if captcha == 'captcha.jpg':
        txt = captcha_to_string(captcha)
        if txt == '':
            return 'error'
    else:
        try:
            error_find = driver.find_element_by_id('capchaVisual')
        except:
            error_find = "Done"
        return error_find
    print(txt)
    time.sleep(5)
    captcha_form = driver.find_element_by_id('captcha-popup-code')
    captcha_form.send_keys(txt)
    time.sleep(2)
    captcha_form = driver.find_element_by_id('ncapcha-submit').click()
    time.sleep(5)
    try:
        error_find = driver.find_element_by_id('capchaVisual')
    except:
        error_find = "Done"
    return error_find

# Функция циклом ищет, скачивает, решает и вставляет капчу


def captcha_resolve():
    captcha_find()
    while captcha_find != 0:
        crack = crack_captcha()
        time.sleep(5)
        while crack != "Done":
            crack = crack_captcha()
        return None

# Функция вставляет ФИО в поле поиска, или очищает поля
# фамилия, имя, отчество, дата рождения и вставляет данные в них


def search_name(personal_name):
    try:
        search_form = driver.find_element_by_id('debt-form01')
        search_form.send_keys(personal_name)
        search_form.submit()
        time.sleep(5)
    except:
        personal_name = personal_name.split(' ')
        driver.find_element_by_id('input01').clear()
        driver.find_element_by_id('input02').clear()
        driver.find_element_by_id('input05').clear()
        driver.find_element_by_id('input06').clear()
        search_form_second_name = driver.find_element_by_id('input01')
        search_form_first_name = driver.find_element_by_id('input02')
        search_form_fathers_name = driver.find_element_by_id('input05')
        search_form_date_birth = driver.find_element_by_id('input06')
        search_form_second_name.send_keys(personal_name[0])
        search_form_first_name.send_keys(personal_name[1])
        search_form_fathers_name.send_keys(personal_name[2])
        search_form_date_birth.send_keys(personal_name[3])
        time.sleep(1)
        search_form_second_name = driver.find_element_by_id('btn-sbm').click()

# Функция находит кнопку "следующая" для обхода всех страниц с результатом


def find_next_page(driver):
    try:
        driver.find_element_by_link_text('СЛЕДУЮЩАЯ').click()
        try:
            captcha_resolve()
        except:
            pass
        return 'button click'
    except:
        return None

# Функция находит табличную часть (результаты поиска) на странице и сохраняет ее в файл


def get_and_write_text(name):
    try:
        tbl = driver.find_element_by_xpath(
            '/html/body/div[3]/main/section/div/div/div[3]/div/div/div[2]').get_attribute('outerHTML')
        df = pd.concat(pd.read_html(tbl))
        # Попытается дозаписать данные в файл, если он существует. Если файла нет - создаст его
        try:
            writer = pd.ExcelWriter(dirpath_save+name+'.xlsx', engine="openpyxl", mode='a')
            df.to_excel(writer, sheet_name='Page')
            writer.save()
        except IOError:
            df.to_excel(dirpath_save+name+'.xlsx')
        return 'Table exist'
    except:
        return None


def main(name):
    search_name(name)  # Запускаем поиск по имени
    captcha_resolve()  # Решаем капчу
    time.sleep(5)
    get_and_write_text(name)  # Вытаскиваем текст и сохраняем его в таблицу
    time.sleep(5)
    # Ищем пагинатор, если он есть - проходим по всем страницам и дополняем таблицу данными
    next_page = find_next_page(driver)
    while next_page != None:
        get_and_write_text(name)
        time.sleep(5)
        next_page = find_next_page(driver)
    print('save done')


if __name__ == "__main__":
    dirpath_read = r'D:\WORK\MTS\\'  # Путь для файла с входящими данными
    dirpath_save = r'D:\WORK\MTS\output\\'  # Путь для файлов, где будут храниться результаты программы
    try:
        Excel = win32.Dispatch("Excel.Application")
        wb = Excel.Workbooks.Open(dirpath_read+'input_data.xlsx')
        sheet = wb.ActiveSheet
        row_count = sheet.UsedRange.Rows.Count
        # Циклично перебираем файл со списком имен, и зпередаем имена в функции поиска.
        # Момент с датой рождения решил самым простым путем, т.к. не нашел способ
        # как изменить форматирование ячейки с датами через pywin32
        for i in range(1, row_count):
            date = str(sheet.Cells(i, 4))[:-15]
            correct_date = date[8:] + "." + date[5:7] + "." + date[:4]
            name = str(sheet.Cells(i, 1)) + ' ' + str(sheet.Cells(i, 2)) + \
                ' ' + str(sheet.Cells(i, 3)) + ' ' + correct_date
            main(name)
        wb.Close()
        Excel.Quit()

    except Exception as e:
        print(e)
