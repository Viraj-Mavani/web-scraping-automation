import csv
import os
import random
import string
import sys
import traceback
import pandas as pd
import sqlite3
import re
from sqlite3 import Error
from bs4 import BeautifulSoup
import time
import chromedriver_autoinstaller
import undetected_chromedriver as uc
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, TimeoutException, StaleElementReferenceException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC

# BasePath = 'D:\\Projects\\CedarPython\\ADIP-IL3701\\'
BasePath = os.getcwd()
# BasePath= 'E:\\ADIP-PY\\OP2'
Driver_path = r"D:\\Projects\\CedarPython\\ChromeDriver\\chromedriver.exe"
# Driver_path = r"E:\\ADIP-PY\\ChromeDriver\\chromedriver.exe"

######### Excel #########
File_path_search_page = BasePath + '\\OP\\ADIP-IL3701_search_page.xlsx'
File_path_company_details = BasePath + '\\OP\\ADIP-IL3701_company_details.xlsx'
File_path_address = BasePath + '\\OP\\ADIP-IL3701_address.xlsx'
######### CSV #########
File_path_search_page_CSV = BasePath + '\\OPcsv\\ADIP-IL3701_search_page.csv'
File_path_company_details_CSV = BasePath + '\\OPcsv\\ADIP-IL3701_company_details.csv'
File_path_address_CSV = BasePath + '\\OPcsv\\ADIP-IL3701_address.csv'
File_path_error_CSV = BasePath + '\\OPcsv\\ADIP-IL3701_Error.csv'
File_path_failed_CSV = BasePath + '\\OP\\ADIP-IL3701_Failed.csv'
######### Text #########
File_path_search_page_TXT = BasePath + '\\OPtxt\\ADIP-IL3701_search_page.txt'
File_path_company_details_TXT = BasePath + '\\OPtxt\\ADIP-IL3701_company_details.txt'
File_path_address_TXT = BasePath + '\\OPtxt\\ADIP-IL3701_address.txt'
######### Error #########
File_path_error = BasePath + '\\Error\\ADIP-IL3701_Error.xlsx'
######### Count #########
File_path_count = BasePath + '\\Counts\\ADIP-IL3701_Count.txt'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-IL3701_Log.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-IL3701_Run_Flag.txt'
File_path_log_index = BasePath + '\\Log\\ADIP-IL3701_Log_Index.txt'
# File_path_log_index_page = BasePath + '\\Log\\ADIP-IL3701_Log_Page.txt'
# File_path_log_index_company = BasePath + '\\Log\\ADIP-IL3701_Log_Company.txt'
# File_path_log_index_address = BasePath + '\\Log\\ADIP-IL3701_Log_Address.txt'


# English_alphabet_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K',
#             'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']


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
    Headers_Error = ['Letter', 'URL', 'Not Responding', 'Error']
    error = traceback.format_exc()
    exception_type, exception_object, exception_traceback = sys.exc_info()
    with open(File_path_error_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(Headers_Error)
        writer.writerow([letter, Base_URL, "Not Responding", error])
    df = pd.read_csv(File_path_error_CSV, encoding='utf-8')
    df.to_excel(File_path_error, index=False)


def count():
    try_count = 1
    while try_count <= 5:
        try:
            with open(File_path_count, 'a', encoding='utf-8') as fh:
                fh.write('1\n')
                fh.flush()
            break
        except Exception:
            pass
        try_count += 1


# def duplicate(File_path_EXL):
#     try:
#         data = pd.read_excel(File_path_EXL)
#         data_file = data.drop_duplicates()
#         data_file.to_excel(File_path_EXL, index=False)
#     except:
#         pass


def duplicateFromCSV(Csv_File_path):
    try:
        data = pd.read_csv(Csv_File_path)
        unique_data = data.drop_duplicates()
        unique_data.to_csv(Csv_File_path, index=False)
    except:
        pass


# def convertCSVExcel(File_path_CSV, File_path_EXL):
#     df = pd.read_csv(File_path_CSV, encoding='utf-8')
#     df.to_excel(File_path_EXL, index=False)


def convertCSVExcelExtended(File_path_CSV, File_path_EXL):
    chunk_size = 1000000  # Number of rows per Excel sheet (adjust as needed)

    # Try to read the entire CSV at once
    try:
        df = pd.read_csv(File_path_CSV, encoding='utf-8')
        df.to_excel(File_path_EXL, index=False)
        return None
    except (pd.errors.ParserError, pd.errors.EmptyDataError, ValueError):
        pass  # The CSV has more than 1000000 rows, so proceed with chunking
    except:
        exception()

    csv_reader = pd.read_csv(File_path_CSV, encoding='utf-8', chunksize=chunk_size)
    sheet_index = 1  # Index of the Excel sheet
    # excel_files = []  # List to store the names of generated Excel files

    for chunk in csv_reader:
        if len(chunk) > 0:  # Create Excel sheet only if chunk is not empty
            sheet_name = f'DataSet {sheet_index}'  # Generate a unique sheet name
            excel_file = f'{File_path_EXL[:-5]}_{sheet_index}.xlsx'  # Generate a unique Excel file name
            chunk.to_excel(excel_file, sheet_name=sheet_name, index=False)
            # excel_files.append(excel_file)
            sheet_index += 1

    # # Merge all Excel files into one
    # writer = pd.ExcelWriter(File_path_EXL, engine='xlsxwriter')
    # Sheet = 1
    # for file in excel_files:
    #     df = pd.read_excel(file)
    #     sheet_name = f'DataSet {Sheet}'
    #     df.to_excel(writer, sheet_name=sheet_name, index=False)
    #     Sheet += 1
    # writer.close()


def scrollnclick(input_element):
    scroll(input_element)
    input_element.click()
    
    
def scroll(input_element):
    script = "arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});"
    driver.execute_script(script, input_element)


def spinner():
    try:
        time.sleep(1)
        log_print("Loading...\t")
        spinner = main_form.find_element(By.XPATH, '//div[@class="k-loading-mask"]')
        # scroll(spinner)
        WebDriverWait(main_form, 150).until(EC.invisibility_of_element_located((spinner)))
        log_print("Done")
        # driver.implicitly_wait(1)
        # WebDriverWait(driver, 50).until(EC.invisibility_of_element_located((By.XPATH, "//div[@id='ctl00_cntMain_UpdateProgress1'][@style='display:none;']")))
        # WebDriverWait(driver, 50).until(EC.visibility_of_element_located((By.XPATH, "//div[@id='ctl00_cntMain_UpdateProgress1'][@style='display:block;']")))
        # WebDriverWait(driver, 50).until(lambda driver: driver.execute_script("return window.getComputedStyle(document.querySelector('#ctl00_cntMain_UpdateProgress1')).getPropertyValue('display') === 'none'"))
    except NoSuchElementException:
        driver.implicitly_wait(15)


def search(letter):
    try:
        search_flag = False
        search_bar = wait_form.until(EC.element_to_be_clickable((By.XPATH, '//div[@id="div_CorporationName"]/input')))
        # search_bar = driver.find_element(By.XPATH, '//input[@class="textBoxes"]')
        # search_bar.click()
        scrollnclick(search_bar) 
        search_bar.send_keys(Keys.CONTROL + "a")
        search_bar.send_keys(letter)
        search_button_element = wait_form.until(EC.element_to_be_clickable((By.XPATH, '//div[@id="div_btnSearchSearchCorporation1"]/input')))
        scrollnclick(search_button_element) 
        # search_button_element.click()
        try:
            spinner()
            search_flag = True
            return search_flag
        except TimeoutException:
            log_print(f"Timeout At Loading...for {letter}")
            driver.refresh()
            driver.implicitly_wait(1)
            
            search_flag = False
            with open(File_path_log_index, 'w', encoding='utf-8') as file:
                file.write(letter)
                file.flush()
            return search_flag
            # os._exit(1)
        # r_delay = random.uniform(0.5, 1.0)
        # time.sleep(r_delay)
    except:
        exception()

        
def individual_data():
    page = 0
    while True:
        try:
            data_div = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@id='divSearchCorporationResultList']/div/div")))
            data_table = data_div.find_element(By.XPATH, "//table/tbody")
            # master_rows = data_table.find_elements(By.XPATH, '//tr[@class="k-master-row" or @class="k-alt k-master-row" or @class="k-master-row k-state-selected"]')
            a_tags = data_table.find_elements(By.XPATH, "//tr//td[@class='k-hierarchy-cell']/a")
            for a_tag in a_tags:
                wait_form.until(EC.element_to_be_clickable((a_tag)))
                scrollnclick(a_tag)
                a_tag
                
            
            soup = BeautifulSoup(data_table.get_attribute('innerHTML'), 'lxml')
            master_rows = soup.find_all('tr', class_=re.compile(r'k-(master|alt k-master)-row(| k-state-selected)')) #//tr[@class="k-master-row" or @class="k-alt k-master-row" or @class="k-alt k-master-row k-state-selected"]
            details_rows = soup.find_all('tr', class_=re.compile(r'k-(detail|detail k-alt)-row')) #//tr[@class="k-master-row" or @class="k-alt k-master-row" or @class="k-alt k-master-row k-state-selected"]
            soup.decompose()

            totMaster_rows = len(master_rows)
            totDetails_rows = len(details_rows)
            
            log_print("Total Master Rows " + str(totMaster_rows) + " Found")
            log_print("Total Details Rows " + str(totDetails_rows) + " Found")
            if totMaster_rows==0:
                log_print("No Data Found!!")
                with open(File_path_log_index, 'w', encoding='utf-8') as file:
                    file.write(letter)
                    file.flush()
                return None

            indi_data_search = []
            indi_data_corporate = []
            indi_data_address = []
            
            for master_row in master_rows:
                master_cells = master_row.find_all('td')[1:-1]
                if len(master_cells) == 6:
                    try:
                        corporate_number = master_cells[0].text.strip()
                    except:
                        corporate_number = ''
                    try:
                        hebrew_name = master_cells[1].text.strip()
                    except:
                        hebrew_name = ''
                    try:
                        type = master_cells[2].text.strip()
                    except:
                        type = ''
                    try:
                        status = master_cells[3].text.strip()
                    except:
                        status = ''
                    try:
                        breaking_law = master_cells[4].text.strip()
                    except:
                        breaking_law = ''
                    try:
                        annual_report = master_cells[5].text.strip()
                    except:
                        annual_report = ''
                        
                indi_data_search.append([corporate_number, hebrew_name, type, status, breaking_law, annual_report])
                
            
            for details_row in details_rows:
                details_tr = details_row.find('td', class_="k-detail-cell").find('div')
                details_cells = details_tr.find_all('div', class_="moj-form-line")[:9]
                if len(details_cells) == 9:
                    hebrew_name = ''
                    try:
                        hebrew_name = details_cells[0].text.strip()
                        # TEXT ==> 'שם בעברית: קינן עעע השקעות בע"מ'
                        key, value = hebrew_name.split(":", 1)
                        if key.strip() == "שם בעברית":
                            hebrew_name = value.strip()
                    except:
                        pass
                    english_name = ''
                    try:
                        english_name = details_cells[1].text.strip()
                        # TEXT ==> 'שם באנגלית: KEINAN AAA INVESTMENTS LTD'
                        key, value = english_name.split(": ", 1)
                        if key.strip() == "שם באנגלית":
                            english_name = value.strip()
                    except:
                        pass
                    incorporation_date = ''
                    try:
                        incorporation_date = details_cells[2].text.strip()
                        # TEXT ==> 'תאריך התאגדות: 17/11/2015'
                        key, value = incorporation_date.split(": ", 1)
                        if key.strip() == "תאריך התאגדות":
                            incorporation_date = value.strip()
                    except:
                        pass
                    sub_status = ''
                    try:
                        sub_status = details_cells[3].text.strip()
                        # TEXT ==> 'תת סטטוס:'
                        key, value = sub_status.split(": ", 1)
                        if key.strip() == "תת סטטוס":
                            sub_status = value.strip()
                    except:
                        pass
                    description = ''
                    try:
                        description = details_cells[4].text.strip()
                        # TEXT ==> 'תיאור חברה:'
                        key, value = description.split(": ", 1)
                        if key.strip() == "תיאור חברה":
                            description = value.strip()
                    except:
                        pass
                    purpose = ''
                    try:
                        purpose = details_cells[5].text.strip()
                        # TEXT ==> 'מטרת החברה: לעסוק בכל עיסוק חוקי'
                        key, value = purpose.split(": ", 1)
                        if key.strip() == "מטרת החברה":
                            purpose = value.strip()
                    except:
                        pass
                    try:
                        govermental_restriction = details_cells[6].text.strip()
                        # TEXT ==> 'חברה ממשלתית: לא | מגבלה: מוגבלת'
                        # 'חברה ממשלתית: לא '
                        # ' מגבלה: מוגבלת'
                        if "|" in govermental_restriction:
                            govermental_restriction_parts = govermental_restriction.split("|")
                            for gov_part in govermental_restriction_parts:
                                key, value = gov_part.split(": ", 1)
                                govermental = ''
                                try:
                                    if key.strip() == "חברה ממשלתית":
                                        govermental = value.strip()
                                except:
                                    pass
                                restriction = ''
                                try:
                                    if key.strip() == "מגבלה":
                                        restriction = value.strip()
                                except:
                                    pass
                    except:
                        pass
                    try:
                        fee_2023_obligations = details_cells[7].text.strip()
                        # TEXT ==> 'חובות אגרה : אין | אגרת 2023 :'
                        if "|" in fee_2023_obligations:
                            fee_2023_obligations_parts = fee_2023_obligations.split("|")
                            for fee_part in fee_2023_obligations_parts:
                                key, value = fee_part.split(": ", 1)
                                fee_obligations = ''
                                try:
                                    if key.strip() == "חובות אגרה":
                                        fee_obligations = value.strip()
                                except:
                                    pass
                                fee_2023 = ''
                                try:
                                    if key.strip() == "אגרת 2023":
                                        fee_2023 = value.strip()
                                except:
                                    pass
                    except:
                        pass
                    try:
                        address = details_cells[8].text.strip()
                        # TEXT ==> 'יישוב: תל אביב - יפו | רחוב: צה"ל | מספר: 89 | מיקוד:  | ת.ד.:  | ארץ: ישראל | אצל:'
                        settlement = ''
                        street = ''
                        number = ''
                        postal = ''
                        pob = ''
                        country = ''
                        in_ = ''
                        if "|" in address:
                            address_parts = address.split("|")
                            for add_part in address_parts:
                                key, value = add_part.split(": ", 1)
                                if key.strip() == "יישוב":
                                    settlement = value.strip()
                                elif key.strip() == "רחוב":
                                    street = value.strip()
                                elif key.strip() == "מספר":
                                    number = value.strip()
                                elif key.strip() == "מיקוד":
                                    postal = value.strip()
                                elif key.strip() == "ת.ד.":
                                    pob = value.strip()
                                elif key.strip() == "ארץ":
                                    country = value.strip()
                                elif key.strip() == "אצל":
                                    in_ = value.strip()
                    except:
                        pass

                indi_data_corporate.append([hebrew_name, english_name, incorporation_date, sub_status, description, purpose, govermental, restriction, fee_obligations, fee_2023])
                indi_data_address.append([settlement, street, number, postal, pob, country, in_])
            
            try:
                # Write to CSV file
                with open(File_path_search_page_CSV, 'a', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerows(indi_data_search)
                # Write to TXT file
                with open(File_path_search_page_TXT, 'a', encoding="utf-8") as fw:
                    for row in indi_data_search:
                        fw.write("\t".join(map(str, row)) + "\n")
                    fw.flush() 
                    
                row_count = len(indi_data_search)
                for c in range(row_count):
                    count()

                # Write to CSV file
                with open(File_path_company_details_CSV, 'a', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerows(indi_data_corporate)
                # Write to TXT file
                with open(File_path_company_details_TXT, 'a', encoding="utf-8") as fw:
                    for row in indi_data_corporate:
                        fw.write("\t".join(map(str, row)) + "\n")
                    fw.flush() 

                # Write to CSV file
                with open(File_path_address_CSV, 'a', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerows(indi_data_address)
                # Write to TXT file
                with open(File_path_address_TXT, 'a', encoding="utf-8") as fw:
                    for row in indi_data_address:
                        fw.write("\t".join(map(str, row)) + "\n")
                    fw.flush() 
                
                # Write to INDEX file
                with open(File_path_log_index, 'w', encoding='utf-8') as file:
                    file.write(letter)
                    file.flush()
            except:
                exception()
            page+=1
            log_print(f"Page {page} Completed")
            
            # Navigate to the next page
            pagination = main_form.find_element(By.XPATH, "//div[@data-role='pager']")
            next_page = WebDriverWait(pagination, 50).until(EC.presence_of_element_located((By.XPATH, "//a[@aria-label='לעמוד הבא']")))
            next_page_class = next_page.get_attribute("class")
            # next_page = pagination.find_element(By.XPATH, "//a[text()=' Next']")
            # parent_class = next_page.find_element(By.XPATH, "..").get_attribute("class")
            
            if "disabled" not in next_page_class:
                scrollnclick(next_page)
                driver.implicitly_wait(0.5)
                continue
                # data_success = individual_data()
                # return data_success
            else:
                break
                data_success = True
                # return data_success

        except:
            exception()


if __name__ == '__main__':
    
    File_paths = [File_path_search_page_CSV, File_path_search_page_TXT, File_path_company_details_CSV, File_path_company_details_TXT, 
                  File_path_address_TXT, File_path_address_CSV, File_path_error_CSV]
    file_paths_logs = [File_path_log, File_path_log_index]
    
    # Create directories if they don't exist
    directories = [
        BasePath + '\\OP',
        BasePath + '\\OPtxt',
        BasePath + '\\OPcsv',
        BasePath + '\\Error',
        BasePath + '\\Counts',
        BasePath + '\\Log'
    ]

    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)
            
    # First_run = True
    # if First_run:
    if not os.path.exists(File_path_log_Run_Flag):
        with open(File_path_log_Run_Flag, "a", encoding='utf-8')as f:
            f.write("")
        for path_log in file_paths_logs:
            if os.path.exists(path_log):
                os.remove(path_log)
        for path_files in File_paths:
            if os.path.exists(path_files):
                os.remove(path_files)
        if os.path.exists(File_path_error_CSV):
            os.remove(File_path_error_CSV)

    HeadersF1 = ["Corporation number", "Name in Hebrew", "Company type", "Corporate status", "Violates of the law", "Annual Report"]
    with open(File_path_search_page_TXT, "a") as fw:
        if fw.tell() == 0:
            fw.write("\t".join(HeadersF1) + "\n")
            fw.flush()
    with open(File_path_search_page_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(HeadersF1)
    
    HeadersF2 = ["Name in Hebrew", "English Name", "Date of incorporation", "Sub-status", "Company description", 
                 "The purpose of the company", "Govermental company", "Limitation", "Toll obligations", "Fee 2023"]
    with open(File_path_company_details_TXT, "a") as fw:
        if fw.tell() == 0:
            fw.write("\t".join(HeadersF2) + "\n")
            fw.flush()
    with open(File_path_company_details_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(HeadersF2)
            
    HeadersF3 = ["Country", "P.O.B", "Zip code", "Street", "Number", "Settlement", "In"]
    with open(File_path_address_TXT, "a") as fw:
        if fw.tell() == 0:
            fw.write("\t".join(HeadersF3) + "\n")
            fw.flush()
    with open(File_path_address_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(HeadersF3)

    try:
        # user_agents = [
        #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36",
        #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:53.0) Gecko/20100101 Firefox/53.0",
        #     "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.0; Trident/5.0; Trident/5.0)",
        #     "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.2; Trident/6.0; MDDCJS)",
        #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.79 Safari/537.36 Edge/14.14393",
        #     "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1)",
        #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36",
        #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/89.0",
        #     "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Safari/537.36 Edg/92.0.902.55",
        #     "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36",
        #     "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/89.0",
        #     "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Safari/537.36 Edg/92.0.902.55"
        #     ]
        # my_user_agent = random.choice(user_agents)
        # headers = {'User-Agent': user_agent}
        # my_user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.159 Safari/537.36"
        
        Base_URL = 'https://ica.justice.gov.il/GenericCorporarionInfo/SearchCorporation?unit=8'

        options = webdriver.ChromeOptions()
        # options = uc.ChromeOptions() 
        # options.add_argument(f"user-agent={my_user_agent}")
        # options.headless = True 
        # # options.add_argument('--ignore-certificate-errors')
        # # options.add_argument('--incognito')
        # # options.add_argument("--no-sandbox")
        options.add_argument("--disable-infobars")
        # # options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-popup-blocking")
        # # options.add_argument("--disable-web-security")
        # # options.add_argument("--allow-running-insecure-content")
        # # options.add_argument('--start-maximized')
        # options.add_argument('--window-size=1920,1080') 
        options.add_argument('--headless')

        # driver = uc.Chrome(options=options) 
        
        ########### Auto chromedriver ###########
        # chromedriver_autoinstaller.install()
        # driver = webdriver.Chrome(options=options)
        
        ########## Manual chromedriver ##########
        service = Service(Driver_path)
        driver = webdriver.Chrome(service=service, options=options)
        
        st = time.time()
        driver.get(Base_URL)
        time.sleep(1)
        wait = WebDriverWait(driver, 100)
        
        t_letter_combinations = [f'{a}{b}{c}' for a in string.ascii_uppercase for b in string.ascii_uppercase for c in string.ascii_uppercase]

        # letterCombinationsList = number_combinations +  d_letter_combinations + dp_letter_combinations + t_letter_combinations + tp_letter_combinations

        log_index_flag = False
        if os.path.exists(File_path_log_index):
            log_index_flag = True
            with open(File_path_log_index, 'r', encoding='utf-8') as file:
                last_processed_letter = file.read().strip()

        if log_index_flag:
            start_index = t_letter_combinations.index(last_processed_letter) + 1
            letter_num_list = t_letter_combinations[start_index:]
        else:
            letter_num_list = t_letter_combinations


        for letter in letter_num_list[:]:
            main_form = wait.until(EC.presence_of_element_located((By.XPATH, '//form[@id="frmSearchCorporation"]')))
            wait_form = WebDriverWait(main_form, 100)
 
            search_return = search(letter)
            
            if search_return:
                individual_data()
            
            # time.sleep(300)
            # r_delay = random.uniform(200, 300)
            # time.sleep(r_delay)
            
            if os.path.exists(File_path_log_index):
                with open(File_path_log_index, 'r', encoding='utf-8') as file:
                    last_letter = file.read().strip()

            if letter == last_letter:
                log_print('Complete ' + letter)
            else:
                log_print('Failed!! ' + letter)
                with open(File_path_failed_CSV, 'a', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow([letter])

    except:
        exception()
        
    finally:
        driver.close()
        duplicateFromCSV(File_path_search_page_CSV)
        duplicateFromCSV(File_path_company_details_CSV)
        duplicateFromCSV(File_path_address_CSV)
        convertCSVExcelExtended(File_path_search_page_CSV, File_path_search_page)
        convertCSVExcelExtended(File_path_company_details_CSV, File_path_company_details)
        convertCSVExcelExtended(File_path_address_CSV, File_path_address)
        et = time.time()
        log_print(f'\n{et - st}')
        os._exit(0)

database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
    for File_path in File_paths:
        delete_task(conn, File_path)
    for File_path in file_paths_logs:
        delete_task(conn, file_paths_logs)
	# delete_task(conn, File_paths)
	# delete_task(conn, file_paths_logs)
