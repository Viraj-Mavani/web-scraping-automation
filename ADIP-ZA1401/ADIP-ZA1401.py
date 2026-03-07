import csv
import os
import random
import sys
import traceback
import pandas as pd
import sqlite3
import re
from sqlite3 import Error
from bs4 import BeautifulSoup
import time
import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC

# BasePath = 'D:\\Projects\\CedarPython\\ADIP-ZA1401\\'
BasePath = os.getcwd()
# BasePath= 'E:\\ADIP-PY\\OP2'
Driver_path = r"D:\\Projects\\CedarPython\\ChromeDriver\\chromedriver.exe"
# Driver_path = r"E:\\ADIP-PY\\ChromeDriver\\chromedriver.exe"

######### Excel #########
File_path = BasePath + '\\OP\\ADIP-ZA1401_Output.xlsx'
######### CSV #########
File_path_CSV = BasePath + '\\OPcsv\\ADIP-ZA1401_Output.csv'
File_path_error_CSV = BasePath + '\\OPcsv\\ADIP-ZA1401_Error.csv'
######### Text #########
File_path_txt = BasePath + '\\Optxt\\ADIP-ZA1401_Output.txt'
######### Error #########
File_path_error = BasePath + '\\Error\\ADIP-ZA1401_Error.xlsx'
######### Count #########
File_path_count = BasePath + '\\Counts\\ADIP-ZA1401_Count.txt'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-ZA1401_Log.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-ZA1401_Run_Flag.txt'
File_path_log_index_LetterE1 = BasePath + '\\Log\\ADIP-ZA1401_Log_Index_LetterE1.txt'
File_path_log_index_LetterE2 = BasePath + '\\Log\\ADIP-ZA1401_Log_Index_LetterE2.txt'
# File_path_log_index_LetterE3 = BasePath + '\\Log\\ADIP-ZA1401_Log_Index_LetterE3.txt'


English_alphabet_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K',
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


def convertCSVExcel(File_path_CSV, File_path_EXL):
    df = pd.read_csv(File_path_CSV, encoding='utf-8')
    df.to_excel(File_path_EXL, index=False)


def duplicate(File_path_EXL):
    try:
        data = pd.read_excel(File_path_EXL)
        data_file = data.drop_duplicates()
        data_file.to_excel(File_path_EXL, index=False)
    except:
        pass


def scrollnclick(input_element):
    script = "arguments[0].scrollIntoView({behavior: 'auto', block: 'center', inline: 'center'});"
    driver.execute_script(script, input_element)
    input_element.click()


def spinner():
    try:
        # WebDriverWait(driver, 50).until(EC.visibility_of_element_located((By.XPATH, "//div[@aria-hidden='true']")))
        # driver.implicitly_wait(1)
        # WebDriverWait(driver, 50).until(EC.invisibility_of_element_located((By.XPATH, "//div[@id='ctl00_cntMain_UpdateProgress1'][@style='display:none;']")))
        # WebDriverWait(driver, 50).until(EC.visibility_of_element_located((By.XPATH, "//div[@id='ctl00_cntMain_UpdateProgress1'][@style='display:block;']")))
        time.sleep(1)
        WebDriverWait(driver, 50).until(lambda driver: driver.execute_script("return window.getComputedStyle(document.querySelector('#ctl00_cntMain_UpdateProgress1')).getPropertyValue('display') === 'none'"))
    except NoSuchElementException:
        driver.implicitly_wait(5)


def search(letter):
    try:
        search_bar = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@class='textBoxes']")))
        # search_bar = driver.find_element(By.XPATH, '//input[@class="textBoxes"]')
        # search_bar.click()
        scrollnclick(search_bar) 
        search_bar.send_keys(Keys.CONTROL + "a")
        search_bar.send_keys(letter)
        search_button_element = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@name='ctl00$cntMain$btnSearch']")))
        scrollnclick(search_button_element) 
        # search_button_element.click()
        spinner()
        r_delay = random.uniform(0.5, 1.0)
        time.sleep(r_delay)
    except:
        exception()

        
def individual_data():
    try:
        data_success = False
        data_div = wait.until(EC.presence_of_element_located((By.XPATH, "//table[@class='mGrid']")))
        soup = BeautifulSoup(data_div.get_attribute('innerHTML'), 'lxml')
        rows = soup.find_all('tr')[1:]
        soup.decompose()

        totRows = len(rows)
        log_print(str(totRows) + "Data Found")
        if totRows==0:
            log_print("No Data Found!!")
            data_success = True
            return data_success

        indi_data = []

        for row in rows:
            try:
                cells = row.find_all('td')
                if len(cells)==3:
                    try:
                        name = cells[0].text.strip()
                    except:
                        name = ''
                    try:
                        number = cells[1].text.strip()
                    except:
                        number = ''
                    try:
                        status = cells[2].text.strip()
                    except:
                        status = ''
                else:
                    continue
                        
                indi_data = [name, number, status]
                # Write to CSV file
                with open(File_path_CSV, 'a', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow(indi_data)
                count()
                # Write to TXT file
                with open(File_path_txt, 'a', encoding="utf-8") as fw:
                    fw.write("\t".join(map(str, indi_data)) + "\n")
                    fw.flush()       
            except:
                exception()

        with open(File_path_log_index_LetterE2, 'w', encoding='utf-8') as file:
            file.write(letterE2)
            file.flush() 
        data_success = True
        return data_success

    except:
        exception()
        data_success = False
        return data_success


if __name__ == '__main__':
    
    File_paths = [File_path_CSV, File_path_txt, File_path_error_CSV]
    file_paths_logs = [File_path_log, File_path_log_index_LetterE1, File_path_log_index_LetterE2]
    File_path_log_index=[File_path_log_index_LetterE1, File_path_log_index_LetterE2]
    
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
        if os.path.exists(File_path_count):
            os.remove(File_path_count)
        if os.path.exists(File_path_CSV):
            os.remove(File_path_CSV)
        if os.path.exists(File_path_txt):
            os.remove(File_path_txt)
        if os.path.exists(File_path_error_CSV):
            os.remove(File_path_error_CSV)

    Headers = ['Enterprise Name', 'Enterprise / Tracking Number', 'Status']
    with open(File_path_count, "a") as f:
        f.write("")
    with open(File_path_txt, "a") as fw:
        if fw.tell() == 0:
            fw.write("\t".join(Headers) + "\n")
            fw.flush()
    with open(File_path_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file) 
        if file.tell() == 0:
            writer.writerow(Headers)

    try:
        Base_URL = 'https://eservices.cipc.co.za/NameSearch.aspx'

        options = webdriver.ChromeOptions()
        options.add_argument('--ignore-certificate-errors')
        options.add_argument('--incognito')
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-infobars")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-popup-blocking")
        options.add_argument("--disable-web-security")
        options.add_argument("--allow-running-insecure-content")
        # options.add_argument('--start-maximized')
        options.add_argument('--window-size=1920,1080') 
        options.add_argument('--headless')

        ########### Auto chromedriver ###########
        chromedriver_autoinstaller.install()
        driver = webdriver.Chrome(options=options)
        
        ########## Manual chromedriver ##########
        # service = Service(Driver_path)
        # driver = webdriver.Chrome(service=service, options=options)
        
        st = time.time()
        driver.get(Base_URL)
        time.sleep(1)
        wait = WebDriverWait(driver, 100)
        
        log_index_flag_LetterE1 = False
        log_index_flag_LetterE2 = False
        log_index_flag_LetterE3 = False
        
        if os.path.exists(File_path_log_index_LetterE1):
            with open(File_path_log_index_LetterE1, 'r', encoding='utf-8') as file:
                index_LetterE1 = file.read().strip()
            if index_LetterE1 != '':
                log_index_flag_LetterE1 = True

        if os.path.exists(File_path_log_index_LetterE2):
            with open(File_path_log_index_LetterE2, 'r', encoding='utf-8') as file:
                index_LetterE2 = file.read().strip()
            if index_LetterE2 != '' and index_LetterE2 != 'Z':
                log_index_flag_LetterE2 = True

        # if os.path.exists(File_path_log_index_LetterE3):
        #     with open(File_path_log_index_LetterE3, 'r', encoding='utf-8') as file:
        #         index_LetterE3 = file.read().strip()
        #     if index_LetterE3 != '' and index_LetterE3 != 'Z':
        #         log_index_flag_LetterE3 = True

        if log_index_flag_LetterE1:
            start_index_LetterE1 = English_alphabet_list.index(index_LetterE1) + 1
        else:
            start_index_LetterE1 = 0

        if log_index_flag_LetterE2:
            start_index_LetterE2 = English_alphabet_list.index(index_LetterE2) + 1
        else:
            start_index_LetterE2 = 0

        # if log_index_flag_LetterE3:
        #     start_index_LetterE3 = English_alphabet_list.index(index_LetterE3) + 1
        # else:
        #     start_index_LetterE3 = 0

        for indexE1 in range(start_index_LetterE1, len(English_alphabet_list)):
            letterE1 = English_alphabet_list[indexE1]
            for indexE2 in range(start_index_LetterE2, len(English_alphabet_list)):
                letterE2 = English_alphabet_list[indexE2]
                letter = (f'{letterE1}%{letterE2}%')
                search(letter)
                
                success = individual_data()
                
                if success:
                    log_print('Complete ' + letter)
                    with open(File_path_log_index_LetterE2, 'w', encoding='utf-8') as file:
                        file.write(letterE2)
                        file.flush()
                else:
                    log_print('Failed!! ' + letter)
            if success:
                with open(File_path_log_index_LetterE1, 'w', encoding='utf-8') as f1:
                    f1.write(letterE1)
                    f1.flush()
            start_index_LetterE2 = 0

    except:
        exception()
        
    finally:
        driver.close()
        convertCSVExcel(File_path_CSV, File_path)
        duplicate(File_path)
        et = time.time()
        log_print(f'\n{et - st}')
        if os.path.exists(File_path_log_index_LetterE1):
            with open(File_path_log_index_LetterE1, 'r', encoding='utf-8') as file:
                last_letter1 = file.read().strip()
        if os.path.exists(File_path_log_index_LetterE2):
            with open(File_path_log_index_LetterE2, 'r', encoding='utf-8') as file:
                last_letter2 = file.read().strip()
        last_letter = last_letter1 + last_letter2
        
        if last_letter.upper() == "ZZ":
            log_print('Script Completed')
            if os.path.exists(File_path_count):
                os.remove(File_path_count)
            if os.path.exists(File_path_log_Run_Flag):
                os.remove(File_path_log_Run_Flag)
        else:
            log_print('Script was Stopped')
        
        
        

database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
	delete_task(conn, File_paths)
	delete_task(conn, file_paths_logs)
