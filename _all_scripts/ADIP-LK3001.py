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
import chromedriver_autoinstaller
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC

# BasePath = 'D:\\Projects\\CedarPython\\ADIP-LK3001\\'
BasePath = os.getcwd()
# BasePath= 'E:\\ADIP-PY\\OP2'

######### Excel #########
File_path = BasePath + '\\OP\\ADIP-LK3001_Output.xlsx'
######### CSV #########
File_path_CSV = BasePath + '\\OPcsv\\ADIP-LK3001_Output.csv'
File_path_error_CSV = BasePath + '\\OPcsv\\ADIP-LK3001_Error.csv'
######### Text #########
File_path_txt = BasePath + '\\Optxt\\ADIP-LK3001_Output.txt'
######### Error #########
File_path_error = BasePath + '\\Error\\ADIP-LK3001_Error.xlsx'
######### Count #########
File_path_count = BasePath + '\\Counts\\ADIP-LK3001_Count.txt'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-LK3001_Log.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-LK3001_Run_Flag.txt'
File_path_log_index_LetterE1 = BasePath + '\\Log\\ADIP-LK3001_Log_Index_LetterE1.txt'
File_path_log_index_LetterE2 = BasePath + '\\Log\\ADIP-LK3001_Log_Index_LetterE2.txt'
File_path_log_index_LetterE3 = BasePath + '\\Log\\ADIP-LK3001_Log_Index_LetterE3.txt'


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
        loading_element = WebDriverWait(driver, 50).until(EC.invisibility_of_element_located((By.XPATH, "//div[@class='loading-text ng-star-inserted']")))
    except NoSuchElementException:
        driver.implicitly_wait(5)


def search(letter):
    try:
        search_element = driver.find_element(By.XPATH, "//div[@class='directory-searcher']")
        wait_search = WebDriverWait(search_element, 100)
        try:
            clear_element = wait_search.until(EC.element_to_be_clickable((By.XPATH, "//button[@mdbtooltip='Clear']")))
            # scroll(clear_element)
            # clear_element.click()
            scrollnclick(clear_element)
        except NoSuchElementException or TimeoutException:
            log_print(f"NoSuchElementException at Clear Button")
        search_bar = search_element.find_element(By.XPATH, "//input[@aria-label='Search']")
        search_bar.click()
        search_bar.send_keys(letter)
        view_button_element = wait_search.until(EC.element_to_be_clickable((By.XPATH, "//button[@mdbtooltip='Click']")))
        # scroll(view_button_element) 
        view_button_element.click()
        spinner()
    except:
        exception()

        
def individual_data():
    try:
        data_success = False
        data_div = wait.until(EC.presence_of_element_located((By.XPATH, "//app-search-bar//div[contains(@class, 'ng-star-inserted') and contains(@style, 'margin-bottom: 100px')]")))
        soup = BeautifulSoup(data_div.get_attribute('innerHTML'), 'lxml')
        job_listings = soup.select('section.ng-star-inserted div.job-listing-details')
        soup.decompose()

        if len(job_listings)==0:
            data_success = True
            return data_success

        indi_data = []

        for listing in job_listings:
            try:
                title = listing.select_one('h3.job-listing-title').text.strip()
            except:
                title = ''
            try:
                r_number = listing.select_one('div.mr-auto').text.strip()
            except:
                r_number = ''

            data = {
                'Title': title,
                'Registration Number': r_number
            }
            indi_data.append(data)
            count()
            # data_success = True

        # Write to CSV file
        csv_df = pd.DataFrame(indi_data)
        csv_df.to_csv(File_path_CSV, index=False, mode='a', header=not os.path.exists(File_path_CSV))

        # Write to TXT file
        with open(File_path_txt, "a") as f:
            for item in indi_data:
                try:
                    f.write("\t".join(map(str, item.values())) + "\n")
                except:
                    continue
                
        # Navigate to the next page
        pagination = data_div.find_element(By.XPATH, "//nav")
        next_page = WebDriverWait(pagination, 50).until(EC.presence_of_element_located((By.XPATH, "//a[text()=' Next']")))
        # next_page = pagination.find_element(By.XPATH, "//a[text()=' Next']")
        parent_class = next_page.find_element(By.XPATH, "..").get_attribute("class")
        
        if "disabled" not in parent_class:
            scrollnclick(next_page)
            spinner()
            # driver.implicitly_wait(0.5)
            data_success = individual_data()
            return data_success
        else:
            data_success = True
            return data_success

    except:
        exception()


if __name__ == '__main__':
    
    File_paths = [File_path_CSV, File_path_txt, File_path_error_CSV]
    file_paths_logs = [File_path_log, File_path_log_index_LetterE1, File_path_log_index_LetterE2, File_path_log_index_LetterE3]
    
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

    Headers = ['Company Name', 'Registration Number']
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
        Base_URL = 'https://eroc.drc.gov.lk/home/search'

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
        # options.add_argument('--start-maximized')
        options.add_argument('--window-size=1920,1080') 
        options.add_argument('--headless')

        driver = webdriver.Chrome(options=options)
        st = time.time()
        driver.get(Base_URL)
        time.sleep(1)
        wait = WebDriverWait(driver, 100)
        
        try:
            alert = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='alert alert-danger']/span")))
            alert.click()
        except:
            pass
        
        radio_label = driver.find_element(By.XPATH, "//label[@class='radio-inline']/input[@value='2']")
        # radio_label.click()
        scrollnclick(radio_label)
        
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

        if os.path.exists(File_path_log_index_LetterE3):
            with open(File_path_log_index_LetterE3, 'r', encoding='utf-8') as file:
                index_LetterE3 = file.read().strip()
            if index_LetterE3 != '' and index_LetterE3 != 'Z':
                log_index_flag_LetterE3 = True

        if log_index_flag_LetterE1:
            start_index_LetterE1 = English_alphabet_list.index(index_LetterE1) + 1
        else:
            start_index_LetterE1 = 0

        if log_index_flag_LetterE2:
            start_index_LetterE2 = English_alphabet_list.index(index_LetterE2) + 1
        else:
            start_index_LetterE2 = 0

        if log_index_flag_LetterE3:
            start_index_LetterE3 = English_alphabet_list.index(index_LetterE3) + 1
        else:
            start_index_LetterE3 = 0

        for indexE1 in range(start_index_LetterE1, len(English_alphabet_list)):
            letterE1 = English_alphabet_list[indexE1]
            for indexE2 in range(start_index_LetterE2, len(English_alphabet_list)):
                letterE2 = English_alphabet_list[indexE2]
                for indexE3 in range(start_index_LetterE3, len(English_alphabet_list)):
                    letterE3 = English_alphabet_list[indexE3]
                    letter = letterE1 + letterE2 + letterE3
                    search(letter)
                    
                    success = individual_data()
                    
                    if success:
                        log_print('Complete ' + letter)
                        with open(File_path_log_index_LetterE3, 'w', encoding='utf-8') as file:
                            file.write(letterE3)
                            file.flush()
                    else:
                        log_print('Failed!! ' + letter)
                if success:
                    with open(File_path_log_index_LetterE2, 'w', encoding='utf-8') as file:
                        file.write(letterE2)
                        file.flush()
                start_index_LetterE3 = 0
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
        exit()

database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
	delete_task(conn, File_paths)
	delete_task(conn, file_paths_logs)
