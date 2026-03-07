import sys
import traceback
import pandas as pd
import sqlite3
# import re
from sqlite3 import Error
from bs4 import BeautifulSoup
import time
import xlsxwriter
import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

BasePath = 'D:\\Projects\\CedarPython\\ADIP-DZ2001\\'

######### Excel #########
File_path_Personnes_Physique = BasePath + 'OP\\ADIP_DZ2001_Personnes_Physique.xlsx'
File_path_Personnes_Morales = BasePath + 'OP\\ADIP_DZ2001_Personnes_Morales.xlsx'
File_path_Personnes_Physique_Arabic = BasePath + 'OP\\ADIP_DZ2001_Personnes_Physique_Arabic.xlsx'
File_path_Personnes_Morales_Arabic = BasePath + 'OP\\ADIP_DZ2001_Personnes_Morales_Arabic.xlsx'
######### Text #########
File_path_Personnes_Physique_txt = BasePath + 'Optxt\\ADIP_DZ2001_Personnes_Physique.txt'
File_path_Personnes_Morales_txt = BasePath + 'Optxt\\ADIP_DZ2001_Personnes_Morales.txt'
File_path_Personnes_Physique_Arabic_txt = BasePath + 'Optxt\\ADIP_DZ2001_Personnes_Physique_Arabic.txt'
File_path_Personnes_Morales_txt_Arabic = BasePath + 'Optxt\\ADIP_DZ2001_Personnes_Morales_Arabic.txt'
######### Text #########
File_path_error = BasePath + 'Error\\ADIP-DZ2001_Error.xlsx'
######### count #########
File_path_search_count= BasePath + 'Counts\\ADIP-DZ2001_Count.txt'

# Arabic_alphabet = ['بـ','تـ','ثـ','جـ','حـ','خـ','سـ','شـ','صـ','ضـ','طـ',
#                    'ظـ','عـ','غـ','فـ','قـ','كـ','لـ','مـ','نـ','هـ','يـ	']
arabic_letters = ['ا', 'ب', 'ت', 'ث', 'ج', 'ح', 'خ', 'د', 'ذ', 'ر', 'ز', 'س', 'ش', 
                  'ص', 'ض', 'ط', 'ظ', 'ع', 'غ', 'ف', 'ق', 'ك', 'ل', 'م', 'ن', 'ه', 'و', 'ي']
english_letters = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K',
                   'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']


def create_connection(db_file):
    """ create a database connection to the SQLite database
		specified by the db_file
	:param db_file: database file
	:return: Connection object or None
	"""
    conn = None
    try:
        conn = sqlite3.connect(db_file)
    except Error as e:
        print(e)
    # except Exception as e:
    # error = traceback.format_exc()
    # print(error)

    return conn


def delete_task(conn, Filepath):
    """
	Delete a task by task id
	:param conn:  Connection to the SQLite database
	:param id: id of the task
	:return:
	"""
    sql = 'delete from FileInfoOutput where Filepath=?'
    cur = conn.cursor()
    cur.execute(sql, (Filepath,))
    conn.commit()


def exception():
    global rowError
    error = traceback.format_exc()
    exception_type, exception_object, exception_traceback = sys.exc_info()
    worksheet_error.write(rowError, 0, Base_URL)
    worksheet_error.write(rowError, 1, "Not Responding")
    worksheet_error.write(rowError, 2, error)
    rowError += 1


def search(letter):
    try:
        search_bar = driver.find_element(By.ID, "critere")
        search_bar.clear()
        search_bar.send_keys(letter)
        search_bar.send_keys(webdriver.Keys.RETURN)
        time.sleep(1)
    except:
        exception()
        # error = traceback.format_exc()
        # exception_type, exception_object, exception_traceback = sys.exc_info()
        # worksheet_error.write(rowError, 0, Base_URL)
        # worksheet_error.write(rowError, 1, "Not Responding")
        # worksheet_error.write(rowError, 2, error)
        # rowError += 1

    
def Individual_data(sheet, row, File_path_list):
        
    wait = WebDriverWait(driver, 10)  # Wait for a maximum of 10 seconds
    ja_body_element = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.ja_body")))
    driver.implicitly_wait(1)
    #  driver.find_element(By.CSS_SELECTOR, "div.ja_body")
    if 'لا يوجد أي عنصر' in ja_body_element.text:
        return
    else:
        try:
            maintainece_element = driver.find_element(By.XPATH, "//div[@class='ja_body']/div/img")
            # under_construction = True
        except NoSuchElementException:
            try:
                soup = BeautifulSoup(driver.page_source, 'lxml')
                table1 = soup.find('div', id='tab1').find('table')
                if table1 is not None:
                    rows = table1.find('tbody').find_all('tr')
                    table_data(rows, sheet[0], 3, row[0], File_path_list[0])
                table2 = soup.find('div', id='tab2').find('table')
                if table2 is not None:
                    rows = table2.find('tbody').find_all('tr')
                    table_data(rows, sheet[1], 2, row[1], File_path_list[1])
                soup.decompose()
            except:
                exception()
        # else:
        #     print("Under Construction LOL")
            

def table_data(rows, sheet, num_data, book, File_path):
    try:
        for row in rows:
            Indi_data = []
            cells = row.find_all('td')
            for i in range(num_data):
                sheet.write(row_vars[book], i, cells[i].string.strip())
                Indi_data.append(cells[i].string.strip())
            row_vars[book] += 1
            try_count = 1
            while True:
                try:
                    with open(File_path_search_count, 'a') as fh:
                        fh.write('1\n')
                    break
                except:
                    if try_count > 5:
                        break
                    try_count += 1
            with open(File_path, 'a', encoding="utf-8") as fw:
                fw.write("\t".join(map(str, Indi_data)) + "\n")
    except:
        exception()


def duplicate(File_path):
    try:
        data = pd.read_excel(File_path)
        data_file = data.drop_duplicates()
        data_file.to_excel(File_path, index=False)
    except:
        pass


if __name__=='__main__':
    row_vars = {
        'book1': 1,
        'book2': 1,
        'book3': 1,
        'book4': 1
    }
    rowError = 1
    # under_construction = False

    # Creating the first workbook
    book1 = xlsxwriter.Workbook(File_path_Personnes_Physique)
    sheet1 = book1.add_worksheet()
    bold_format = book1.add_format({'bold': True})
    sheet1.write('A1', 'NRC', bold_format)
    sheet1.write('B1', 'Nom', bold_format)
    sheet1.write('C1', 'Prenom', bold_format)

    # Creating the second workbook
    book2 = xlsxwriter.Workbook(File_path_Personnes_Morales)
    sheet2 = book2.add_worksheet()
    bold_format = book2.add_format({'bold': True})
    sheet2.write('A1', 'NRC', bold_format)
    sheet2.write('B1', 'Raison Sociale', bold_format)

    # Creating the third workbook
    book3 = xlsxwriter.Workbook(File_path_Personnes_Physique_Arabic)
    sheet3 = book3.add_worksheet()
    bold_format = book3.add_format({'bold': True})
    sheet3.write('A1', 'NRC', bold_format)
    sheet3.write('B1', 'Nom (Arabic)', bold_format)
    sheet3.write('C1', 'Prenom (Arabic)', bold_format)

    # Creating the fourth workbook
    book4 = xlsxwriter.Workbook(File_path_Personnes_Morales_Arabic)
    sheet4 = book4.add_worksheet()
    bold_format = book4.add_format({'bold': True})
    sheet4.write('A1', 'NRC', bold_format)
    sheet4.write('B1', 'Raison Sociale (Arabic)', bold_format)
    
    # Creating the Error workbook
    workbook_error = xlsxwriter.Workbook(File_path_error)
    worksheet_error = workbook_error.add_worksheet()
    bold_format = workbook_error.add_format({'bold': True})

    worksheet_error.write('A1', 'URL', bold_format)
    worksheet_error.write('B1', 'Not Responding', bold_format)
    worksheet_error.write('C1', 'Error', bold_format)

    Personnes_Physique_headers = ['NRC', 'Nom', 'Prenom']
    Personnes_Physique_Arabic_headers = ['NRC', 'Nom (Arabic)', 'Prenom (Arabic)']
    Personnes_Morales_headers = ['NRC', 'Raison Sociale']
    Personnes_Morales_Arabic_headers = ['NRC', 'Raison Sociale (Arabic)']

    with open(File_path_search_count, "w")as f:
        f.write("")
    with open(File_path_Personnes_Physique_txt, "w")as f:
        f.write("\t".join(Personnes_Physique_headers)+"\n")
    with open(File_path_Personnes_Morales_txt, "w")as fw:
        fw.write("\t".join(Personnes_Morales_headers)+"\n")
    with open(File_path_Personnes_Physique_Arabic_txt, "w")as f:
        f.write("\t".join(Personnes_Physique_Arabic_headers)+"\n")
    with open(File_path_Personnes_Morales_txt_Arabic, "w")as fw:
        fw.write("\t".join(Personnes_Morales_Arabic_headers)+"\n")
        
    try:
        Base_URL = 'https://sidjilcom.cnrc.dz/web/cnrc/accueil'

        chromedriver_autoinstaller.install()

        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--incognito')
        options.add_argument('--start-maximized')
        # options.add_argument('--headless')

        driver = webdriver.Chrome(options=options)
        st = time.time()
        driver.get(Base_URL)
        time.sleep(1)
        
        try:
            # for letter1 in english_letters[26:]:
                # if under_construction:
                #     break
            for letter in english_letters[25:]:
                # letter = letter1 + letter2
                search(letter)
                Individual_data([sheet1, sheet2], ['book1', 'book2'], [
                                'File_path_Personnes_Physique_txt', 'File_path_Personnes_Morales_txt'])
                # if under_construction:
                #     break
                print(f'For Letter {letter}')
                close_button_element = driver.find_element(By.XPATH, "//div[contains(@class, 'closejAlert') and contains(text(), 'X')]")
                close_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(close_button_element))
                close_button.click()
            # print(f'Complete {letter1}\n\n')
            print(f'Complete {letter}\n\n')
            print('_____________________________________')

        finally:
            book1.close()
            book2.close()
            duplicate(File_path_Personnes_Physique)
            duplicate(File_path_Personnes_Morales)
        
        
        arabic_link_element = driver.find_element(By.XPATH, "//a[contains(@href, 'ar_SA') and contains(text(), 'العربية')]")
        arabic_link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(arabic_link_element))
        arabic_link.click()
        
        try:
            # for letter1 in arabic_letters[2:]:
                # if under_construction:
                #     break
            for letter in arabic_letters[:2]:
                # letter = letter1 + letter2
                search(letter)
                # under_construction = False
                Individual_data([sheet3, sheet4], [
                                'book3', 'book4'], ['File_path_Personnes_Physique_Arabic_txt', 'File_path_Personnes_Morales_txt_Arabic'])
                # if under_construction:
                #     break
                print(f'For Letter {letter}')
                close_button_element = driver.find_element(By.XPATH, "//div[contains(@class, 'closejAlert') and contains(text(), 'X')]")
                close_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(close_button_element))
                close_button.click()
            # print(f'Complete {letter1}\n\n')
            print(f'Complete {letter}\n\n')
            print('_____________________________________')
        finally:
            book3.close()
            book4.close()
            duplicate(File_path_Personnes_Physique_Arabic)
            duplicate(File_path_Personnes_Morales_Arabic)
            
    finally:
        workbook_error.close()
        driver.close()
        et = time.time()
        print(et-st)
        exit()

database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
	delete_task(conn, File_path_Personnes_Physique)
	delete_task(conn, File_path_Personnes_Morales)
	delete_task(conn, File_path_Personnes_Physique_Arabic)
	delete_task(conn, File_path_Personnes_Morales_Arabic)
	delete_task(conn, File_path_search_count)
	delete_task(conn, File_path_error)
