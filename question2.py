import pandas as pd
import time
import win32com.client as win32
import openpyxl
from selenium import webdriver
from selenium.webdriver.support.select import Select


# Функция вставляет ФИО в поле поиска и выбирает Москву


def search_name(personal_name):
    try:
        moscow = driver.find_element_by_id("spSearchArea").find_element_by_id(
            'courtGuideTbl').find_element_by_id('court_subj')
        time.sleep(2)
        select_moscow = Select(moscow)
        select_moscow.select_by_index(11)
        search_form = driver.find_element_by_xpath('//*[@id="f_name"]')
        search_form.clear()
        search_form = driver.find_element_by_xpath('//*[@id="f_name"]')
        search_form.send_keys(personal_name)
        search_form.submit()
        time.sleep(10)
    except:
        pass


# Функция находит табличную часть (результаты поиска)
#  на странице и сохраняет ее в файл


def get_and_write_text(name):
    try:
        result_table = driver.find_element_by_xpath('//*[@id="resultTable"]').get_attribute('outerHTML')
        dataframe = pd.concat(pd.read_html(result_table))
        # Попытается дозаписать данные в файл, если он существует. Если файла нет - создаст его
        try:
            writer = pd.ExcelWriter(dirpath_save+name+'.xlsx', engine="openpyxl", mode='a')
            dataframe.to_excel(writer, sheet_name='Page')
            writer.save()
        except IOError:
            dataframe.to_excel(dirpath_save+name+'.xlsx')
        return 'Table exist'
    except:
        print('error write file')
        return None


def main(name):
    search_name(name)  # Запускаем поиск по имени
    time.sleep(5)
    get_and_write_text(name)  # Вытаскиваем текст и сохраняем его в таблицу
    time.sleep(5)
    print('Обработка данных завершена')


if __name__ == "__main__":
    dirpath_read = r'D:\WORK\MTS\\'  # Путь для файла с входящими данными
    dirpath_save = r'D:\WORK\MTS\output2\\'  # Путь для файлов, где будут храниться результаты программы
    driver = webdriver.Firefox()  # Настройки селениума
    executor_url = driver.command_executor._url
    site = "https://sudrf.ru/index.php?id=300#sp"
    driver.get(site)  # Запуск селениума
    time.sleep(3)
    try:
        Excel = win32.Dispatch("Excel.Application")
        wb = Excel.Workbooks.Open(dirpath_read+'input_data.xlsx')
        sheet = wb.ActiveSheet
        row_count = sheet.UsedRange.Rows.Count
        for i in range(1, row_count):
            name = str(sheet.Cells(i, 1)) + ' ' + str(sheet.Cells(i, 2)) + ' ' + str(sheet.Cells(i, 3))
            main(name)
        wb.Close()
        Excel.Quit()
        print('Обработка всех данных завершена')

    except Exception as e:
        print(e)
