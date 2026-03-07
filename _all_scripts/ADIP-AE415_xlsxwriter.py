import csv
import os
import sys
import traceback
import pandas as pd
import sqlite3
import re
from sqlite3 import Error
from bs4 import BeautifulSoup
import time
import xlsxwriter
# import openpyxl
import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC

BasePath = 'D:\\Projects\\CedarPython\\ADIP-AE415\\'
# BasePath = 'F:\\CedarPython\\ADIP-AE415\\'

# File_path_English_excel = BasePath + 'Optxt\\ADIP-AE415_English_excel.xlsx'
# File_path_Arabic_excel = BasePath + 'Optxt\\ADIP-AE415_Arabic_excel.xlsx'
# File_path_Activity_excel = BasePath + 'Optxt\\ADIP-AE415_Activity_excel.xlsx'
######### Excel #########
File_path_English = BasePath + 'OP\\ADIP-AE415_English.xlsx'
File_path_Arabic = BasePath + 'OP\\ADIP-AE415_Arabic.xlsx'
File_path_Activity = BasePath + 'OP\\ADIP-AE415_Activity.xlsx'
File_path_Input = BasePath + 'InputFile\\ADIP-AE415-Input-Updated.xlsx'
######### CSV #########
File_path_English_CSV = BasePath + 'OPcsv\\ADIP-AE415_English.csv'
File_path_Arabic_CSV = BasePath + 'OPcsv\\ADIP-AE415_Arabic.csv'
File_path_Activity_CSV = BasePath + 'OPcsv\\ADIP-AE415_Activity.csv'
######### Text #########
File_path_English_txt = BasePath + 'Optxt\\ADIP-AE415_English.txt'
File_path_Arabic_txt = BasePath + 'Optxt\\ADIP-AE415_Arabic.txt'
File_path_Activity_txt = BasePath + 'Optxt\\ADIP-AE415_Activity.txt'
######### Error #########
File_path_error = BasePath + 'Error\\ADIP-AE415_Error.xlsx'
File_path_error_CSV = BasePath + 'Error\\ADIP-AE415_Error.csv'
######### Count #########
File_path_count = BasePath + 'Counts\\ADIP-AE415_Count.txt'
######### Log #########
File_path_log = BasePath + '\Log\ADIP-AE415_Log.txt'
File_path_log_index_English = BasePath + '\Log\ADIP-AE415_Log_Index_English.txt'
File_path_log_index_Arabic = BasePath + '\Log\ADIP-AE415_Log_Index_Arabic.txt'
######### Failed #########
# File_path_failed_English = BasePath + 'Error\\ADIP-AE415_Failed_English.xlsx'
# File_path_failed_Arabic = BasePath + 'Error\\ADIP-AE415_Failed_Arabic.xlsx'
File_path_failed_English_CSV = BasePath + 'Error\\ADIP-AE415_Failed_English.csv'
File_path_failed_Arabic_CSV = BasePath + 'Error\\ADIP-AE415_Failed_Arabic.csv'


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


def log_print(message):
    with open(File_path_log, 'a', encoding='utf-8') as file:
        file.write(message + '\n')
        file.flush()
    print(message)


def exception():
    global rowError
    error = traceback.format_exc()
    exception_type, exception_object, exception_traceback = sys.exc_info()
    worksheet_error.write(rowError, 0, Base_URL)
    worksheet_error.write(rowError, 1, "Not Responding")
    worksheet_error.write(rowError, 2, error)
    # worksheet_error.write(rowError, 3, exception_traceback)
    # worksheet_error.write(rowError, 4, exception_object)
    rowError += 1


def CountTxtFile(File_path, Indi_data):
    try_count = 1
    while True:
        try:
            with open(File_path_count, 'a') as fh:
                fh.write('1\n')
            break
        except:
            if try_count > 5:
                break
            try_count += 1

    with open(File_path, 'a', encoding="utf-8") as fw:
        fw.write("\t".join(map(str, Indi_data)) + "\n")
        fw.flush()


def scroll(input_element):
    script = "arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});"
    driver.execute_script(script, input_element)


def search_bar_click():
    try:
        search_bar = wait.until(EC.element_to_be_clickable((By.NAME, "licenceNumber")))
    except NoSuchElementException or TimeoutException:
        driver.refresh()
        driver.implicitly_wait(1)
        search_bar = wait.until(EC.element_to_be_clickable((By.NAME, "licenceNumber")))
    else:
        try:
            close_button_element = driver.find_element(
                By.XPATH, "//div[@role='button'][@aria-label='Close' or @aria-label='إغلاق']")
            close_button = wait.until(EC.element_to_be_clickable(close_button_element))
            close_button.click()
        except NoSuchElementException or TimeoutException:
            pass
        scroll(search_bar)
            # wait.until(EC.element_to_be_clickable(search_bar))
        search_bar.click()
        search_bar.send_keys(Keys.CONTROL + "a")
        search_bar.send_keys(licence_num)
        view_button_element = driver.find_element(By.XPATH, "//button[@type='button'][@aria-label='button-primary']")
        view_button = wait.until(
            EC.element_to_be_clickable(view_button_element))
        scroll(view_button_element)
        view_button.click()
    

def search(licence_num, sheet, book):
    try:
        invalid_flag = False
        search_bar_click()

        try:
            loading_element = driver.find_element(By.XPATH,
                                                  "//div[@class='ui-lib-spinner ui-lib-spinner_circle-image']")
            # wait.until(EC.invisibility_of_element(loading_element))
            retries = 0
            while retries < 3:
                try:
                    wait.until(EC.invisibility_of_element(loading_element))
                    break  
                except TimeoutException:
                    retries += 1
                    log_print(f"Timeout error occurred in view button. Retrying search...{retries}")
                    driver.refresh()
                    driver.implicitly_wait(1)
                    search_bar_click()
                    loading_element = driver.find_element(By.XPATH,
                                                          "//div[@class='ui-lib-spinner ui-lib-spinner_circle-image']")
                    wait.until(EC.invisibility_of_element(loading_element))
                    # continue
                    break
        except NoSuchElementException:
            pass
        # invalid_count = 0
        # while invalid_count < 1:
        try:
            invalid_element = driver.find_element(
                By.XPATH,
                "//div[contains(text(), 'Invalid licence number') or contains(text(), 'رقم الرخصة غير صحيح') or contains(@class, 'ui-lib-alert__text-description')]")
            invalid_flag = True
        except NoSuchElementException:
            return invalid_flag
            # driver.implicitly_wait(1)
        #         invalid_count += 1
        # except:
        #     exception()

        if invalid_flag:
            tradelicenceNumber = licence_num
            name_activity = ''
            if not arabic_flag:
                Indi_data = []
                sheet[0].write(row_vars[book[0]], 0, tradelicenceNumber)
                licenceStatus = 'Deleted'
                sheet[0].write(row_vars[book[0]], 10, licenceStatus)
                row_vars[book[0]] += 1
                Indi_data.append(tradelicenceNumber)
                Indi_data.append('')
                Indi_data.append('')
                Indi_data.append(licenceStatus)
                for i in range(4, 15):
                    Indi_data.append('')
                CountTxtFile(File_path_English_txt, Indi_data)

                Indi_data = []
                sheet[1].write(row_vars[book[1]], 0, tradelicenceNumber)
                sheet[1].write(row_vars[book[1]], 1, name_activity)
                sheet[1].write(row_vars[book[1]], 2, name_activity)
                row_vars[book[1]] += 1
                Indi_data.append(tradelicenceNumber)
                Indi_data.append('')
                Indi_data.append('')
                # CountTxtFile(File_path_Activity_txt, Indi_data)
                with open(File_path_Activity_txt, 'a', encoding="utf-8") as fw:
                    fw.write("\t".join(map(str, Indi_data)) + "\n")
                    fw.flush()
                with open(File_path_log_index_English, 'w', encoding='utf-8') as file:
                    file.write(licence_num)
                    file.flush()

            elif arabic_flag:
                Indi_data = []
                # tradelicenceNumber = licence_num
                sheet.write(row_vars[book], 0, tradelicenceNumber)
                sheet.write(row_vars[book], 1, name_activity)
                row_vars[book] += 1
                Indi_data.append(tradelicenceNumber)
                Indi_data.append('')
                # CountTxtFile(File_path_Arabic_txt, Indi_data)
                with open(File_path_Arabic_txt, 'a', encoding="utf-8") as fw:
                    fw.write("\t".join(map(str, Indi_data)) + "\n")
                    fw.flush()
                with open(File_path_log_index_Arabic, 'w', encoding='utf-8') as file:
                    file.write(licence_num)
                    file.flush()
            return invalid_flag

        # except NoSuchElementException or TimeoutException:
        #     # time.sleep(1)
    except:
        exception()


def Individual_data(sheet, book):
    try:
        if not arabic_flag:
            Indi_data = []
            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'table.ui-lib-table__root-grid')))
            soup = BeautifulSoup(driver.page_source, 'lxml')
            # time.sleep(5)
            tables = soup.find_all('table', class_='ui-lib-table__root-grid')
            rows = tables[0].find('tbody').find_all('tr')
            table_data1(rows, sheet[0], row_vars[book[0]], Indi_data)
            row_vars[book[0]] += 1
            CountTxtFile(File_path_English_txt, Indi_data)
            with open(File_path_log_index_English, 'w', encoding='utf-8') as file:
                file.write(licence_num)
                file.flush()

            # Indi_data = []
            # if tables[1]:
            rows = tables[1].find('tbody').find_all('tr')
            table_data2(rows, sheet[1], book[1])
            # row_vars[book[1]] += 1
            # CountTxtFile(File_path_Activity_txt, Indi_data)

            soup.decompose()
        elif arabic_flag:
            Indi_data = []
            wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'table.ui-lib-table__root-grid')))
            soup = BeautifulSoup(driver.page_source, 'lxml')
            cells = soup.find(
                'table', class_='ui-lib-table__root-grid').find('tbody').find_all('td', {"aria-label": 'details'})
            # table_data1(cells, sheet, row_vars[book], Indi_data)
            if len(cells) >= 3:
                tradelicenceNumber = cells[0].text.strip() if cells[0].text else ''
                sheet.write(row_vars[book], 0, tradelicenceNumber)
                name = cells[2].text.strip() if cells[2].text else ''
                sheet.write(row_vars[book], 1, name)
                Indi_data.append(tradelicenceNumber)
                Indi_data.append(name)
            row_vars[book] += 1
            # CountTxtFile(File_path_Arabic_txt, Indi_data)
            with open(File_path_Arabic_txt, 'a', encoding="utf-8") as fw:
                fw.write("\t".join(map(str, Indi_data)) + "\n")
                fw.flush()
            with open(File_path_log_index_Arabic, 'w', encoding='utf-8') as file:
                file.write(licence_num)
                file.flush()

            soup.decompose()
    except:
        exception()


def table_data1(rows, sheet, xlrow, Indi_data):
    i = 0
    for row in rows:
        try:
            cell = row.find('td', {"aria-label": 'details'})
            sheet.write(xlrow, i, cell.text.strip())
            Indi_data.append(cell.text.strip())
            i += 1

        except:
            exception()
            continue


def table_data2(rows, sheet, book):
    for row in rows:
        try:
            Indi_data = []
            cell = row.find('td', {"aria-label": 'description'})

            # Matches text between parentheses
            pattern = r'^(.*?)\s+\((\d+)\)$'

            cell_text = cell.text.strip()
            matches = re.match(pattern, cell_text)
            try:
                if matches:
                    tradeLicenceActivities = matches.group(1).strip()
                    tradeLicenceActivities_Code = matches.group(2).strip()
            except:
                tradeLicenceActivities = ''
                tradeLicenceActivities_Code = ''

            sheet.write(row_vars[book], 0, licence_num)
            sheet.write(row_vars[book], 1, tradeLicenceActivities)
            sheet.write(row_vars[book], 2, tradeLicenceActivities_Code)
            Indi_data.append(licence_num)
            Indi_data.append(tradeLicenceActivities)
            Indi_data.append(tradeLicenceActivities_Code)
            # xlrow += 1
            row_vars[book] += 1
            # CountTxtFile(File_path_Activity_txt, Indi_data)
            with open(File_path_Activity_txt, 'a', encoding="utf-8") as fw:
                fw.write("\t".join(map(str, Indi_data)) + "\n")
                fw.flush()
        except:
            exception()
            continue


# def convertTxtExcel(File_path_txt, File_path_excel):
#     dfConvert = pd.read_csv(File_path_txt, delimiter='\t')
#     dfConvert.to_excel(File_path_excel, index=False, engine='openpyxl')
#     duplicate(File_path_excel)

def duplicate(File_path):
    try:
        data = pd.read_excel(File_path)
        data_file = data.drop_duplicates()
        data_file.to_excel(File_path, index=False)
    except:
        pass


if __name__ == '__main__':
    row_vars = {
        'book1': 1,
        'book2': 1,
        'book3': 1,
    }
    rowError = 1
    arabic_flag = False
    if os.path.exists(File_path_log):
        os.remove(File_path_log)

    # Create directories if they don't exist
    directories = [
        BasePath + 'OP',
        BasePath + 'InputFile',
        BasePath + 'Optxt',
        BasePath + 'OPcsv',
        BasePath + 'Error',
        BasePath + 'Counts',
        BasePath + 'Log'
    ]

    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)

    # Creating the first workbook
    book1 = xlsxwriter.Workbook(File_path_English)
    sheet1 = book1.add_worksheet()
    bold_format = book1.add_format({'bold': True})
    sheet1.write('A1', 'Trade licence Number', bold_format)
    sheet1.write('B1', 'ADCCI Number', bold_format)
    sheet1.write('C1', 'Trade Name', bold_format)
    sheet1.write('D1', 'Legal Form', bold_format)
    sheet1.write('E1', 'Licence Type', bold_format)
    sheet1.write('F1', 'Branch', bold_format)
    sheet1.write('G1', 'Issuance Place', bold_format)
    sheet1.write('H1', 'Establishment Date', bold_format)
    sheet1.write('I1', 'Registration Date', bold_format)
    sheet1.write('J1', 'Expiry Date', bold_format)
    sheet1.write('K1', 'Licence Status', bold_format)
    sheet1.write('L1', 'Address', bold_format)
    sheet1.write('M1', 'Establishment Volume', bold_format)
    sheet1.write('N1', 'Social Media Account', bold_format)
    sheet1.write('O1', 'Social Media Type', bold_format)
    sheet1.write('P1', 'Web Site URL', bold_format)
    # sheet1.write('Q1', 'Trade Licence Activities', bold_format)
    # sheet1.write('R1', 'Trade Licence Activities - Code', bold_format)

    # Creating the second workbook
    book2 = xlsxwriter.Workbook(File_path_Arabic)
    sheet2 = book2.add_worksheet()
    bold_format = book2.add_format({'bold': True})
    sheet2.write('A1', 'Trade licence Number', bold_format)
    sheet2.write('B1', 'Trade Name', bold_format)

    # Creating the third workbook
    book3 = xlsxwriter.Workbook(File_path_Activity)
    sheet3 = book3.add_worksheet()
    bold_format = book3.add_format({'bold': True})
    sheet3.write('A1', 'Trade licence Number', bold_format)
    sheet3.write('B1', 'Trade Licence Activities', bold_format)
    sheet3.write('C1', 'Trade Licence Activities - Code', bold_format)

    # Creating the Error workbook
    workbook_error = xlsxwriter.Workbook(File_path_error)
    worksheet_error = workbook_error.add_worksheet()
    bold_format = workbook_error.add_format({'bold': True})

    worksheet_error.write('A1', 'URL', bold_format)
    worksheet_error.write('B1', 'Not Responding', bold_format)
    worksheet_error.write('C1', 'Error', bold_format)

    English_headers = ['Trade Licence Number', 'ADCCI Number', 'Trade Name', 'Legal Form', 'Licence Type',
                       'Branch', 'Issuance Place', 'Establishment Date', 'Registration Date', 'Expiry Date',
                       'Licence Status', 'Address', 'Establishment Volume', 'Social Media Account',
                       'Social Media Type', 'Web Site URL',]
    Arabic_headers = ['Trade Licence Number', 'Trade Name']
    Activity_headers = ['Trade Licence Number', 'Trade Licence Activities', 'Trade Licence Activities - Code']

    with open(File_path_count, "w") as f:
        f.write("")
    with open(File_path_English_txt, "w") as fw:
        fw.write("\t".join(English_headers) + "\n")
        fw.flush()
    with open(File_path_Arabic_txt, "w") as fw:
        fw.write("\t".join(Arabic_headers) + "\n")
        fw.flush()
    with open(File_path_Activity_txt, "w") as fw:
        fw.write("\t".join(Activity_headers) + "\n")
        fw.flush()

    try:
        Base_URL = 'https://www.tamm.abudhabi/services/business/ded/get-licence-details/licence-number'

        chromedriver_autoinstaller.install()

        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--incognito')
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-infobars")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--disable-web-security")
        options.add_argument("--allow-running-insecure-content")
        options.add_argument('--start-maximized')
        # options.add_argument('--headless')

        driver = webdriver.Chrome(options=options)
        st = time.time()
        driver.get(Base_URL)
        time.sleep(1)
        wait = WebDriverWait(driver, 50)

        # licence_num_list = ["CN-100002","CN-1000104","CN-1000116"]
        log_print('Data Importing...plz wait')
        df = pd.read_excel(File_path_Input, sheet_name='Sheet1')
        licence_num_list_og = df['number'].tolist()
        # licence_num_list = ["CN-100002", "CN-1000104", "CN-1000141"]
        log_print('Data Imported\n')
        
        log_index_flag = False
        if os.path.exists(File_path_log_index_English):
            log_index_flag = True
            with open(File_path_log_index_English, 'r', encoding='utf-8') as file:
                last_processed_licence = file.read().strip()
            
        # last_processed_index = licence_num_list.index(last_processed_licence)
        if log_index_flag:
            # Find the index of the log entry in the licence_num_list
            start_index = licence_num_list_og.index(last_processed_licence) + 1
            licence_num_list = licence_num_list_og[start_index:]
        else:
            licence_num_list = licence_num_list_og

        ########################################### License For English ###########################################

        # completed_without_errors = True
        try:
            for licence_num in licence_num_list[:]:
                try:
                    invalid_flag = search(licence_num, [sheet1, sheet3], ['book1', 'book3'])
                    if not invalid_flag:
                        Individual_data([sheet1, sheet3], ['book1', 'book3'])
                    log_print(f'Complete {licence_num} in English')
                    
                except:
                    exception()
                    # completed_without_errors = False
                    continue
            os.remove(File_path_log_index_English)
        # except:
            # completed_without_errors = False
        finally:
            book1.close()
            book3.close()
            # convertTxtExcel(File_path_English_txt, File_path_English_excel)
            # convertTxtExcel(File_path_Activity_txt, File_path_Activity_excel)
            duplicate(File_path_English)
            duplicate(File_path_Activity)
            # if completed_without_errors and os.path.exists(File_path_log_index_English):
            #     os.remove(File_path_log_index_English)

        arabic_link_element = driver.find_element(
            By.XPATH, "//button[@type='button'][@class='ui-lib-language-switcher']")
        scroll(arabic_link_element)
        arabic_link = wait.until(
            EC.element_to_be_clickable(arabic_link_element))
        arabic_link.click()
        arabic_flag = True
        time.sleep(1)
        log_print('==================================================')

        log_index_flag = False
        if os.path.exists(File_path_log_index_Arabic):
            log_index_flag = True
            with open(File_path_log_index_Arabic, 'r', encoding='utf-8') as file:
                last_processed_licence = file.read().strip()

        # last_processed_index = licence_num_list.index(last_processed_licence)
        if log_index_flag:
            # Find the index of the log entry in the licence_num_list
            start_index = licence_num_list_og.index(last_processed_licence) + 1
            licence_num_list = licence_num_list_og[start_index:]
        else:
            licence_num_list = licence_num_list_og
            
        ########################################### License For Arabic ###########################################
        
        # completed_without_errors = True
        try:
            for licence_num in licence_num_list[:]:
                try:
                    invalid_flag = search(licence_num, sheet2, 'book2')
                    if not invalid_flag:
                        Individual_data(sheet2, 'book2')
                    log_print(f'Complete {licence_num} in Arabic')
                    
                except:
                    exception()
                    # completed_without_errors = False
                    continue
            os.remove(File_path_log_index_Arabic)
        # except:
            # completed_without_errors = False
        finally:
            book2.close()
            # convertTxtExcel(File_path_Arabic_txt, File_path_Arabic_excel)
            # if completed_without_errors and os.path.exists(File_path_log_index_Arabic):
            #     os.remove(File_path_log_index_Arabic)
            
    except:
        exception()
        
    finally:
        workbook_error.close()
        driver.close()
        et = time.time()
        log_print(f'\n{et - st}')
        exit()

database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
	delete_task(conn, File_path_English)
	delete_task(conn, File_path_Arabic)
	delete_task(conn, File_path_Activity)
	# delete_task(conn, File_path_English_excel)
	# delete_task(conn, File_path_Arabic_excel)
	# delete_task(conn, File_path_Activity_excel)
	delete_task(conn, File_path_count)
	delete_task(conn, File_path_error)
