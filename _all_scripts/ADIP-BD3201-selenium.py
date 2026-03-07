import base64
import os
import pytesseract
import imageio.v3 as iio
import cv2
import csv
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
from selenium.common.exceptions import NoSuchElementException,StaleElementReferenceException,TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC

# BasePath = 'D:\\Projects\\CedarPython\\ADIP-BD3201\\'
BasePath = os.getcwd()
# BasePath= 'E:\\ADIP-PY\\OP2'

######### Excel #########
File_path = BasePath + '\\OP\\ADIP-BD3201_Output.xlsx'
######### CSV #########
File_path_CSV = BasePath + '\\OPcsv\\ADIP-BD3201_Output.csv'
File_path_error_CSV = BasePath + '\\OPcsv\\ADIP-BD3201_Error.csv'
######### Text #########
File_path_txt = BasePath + '\\Optxt\\ADIP-BD3201_Output.txt'
######### Error #########
File_path_error = BasePath + '\\Error\\ADIP-BD3201_Error.xlsx'
######### Count #########
File_path_count = BasePath + '\\Counts\\ADIP-BD3201_Count.txt'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-BD3201_Log.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-BD3201_Run_Flag.txt'
File_path_log_index_LetterE1 = BasePath + '\\Log\\ADIP-BD3201_Log_Index_LetterE1.txt'
File_path_log_index_LetterE2 = BasePath + '\\Log\\ADIP-BD3201_Log_Index_LetterE2.txt'
File_path_log_index_LetterE3 = BasePath + '\\Log\\ADIP-BD3201_Log_Index_LetterE3.txt'
######### IMG #########
og_image_path = BasePath + '\\Log\\ADIP-BD3201_Captcha.png'
# converted_image_path = BasePath + '\\Log\\ADIP-BD3201_Converted_img.png'


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


def img2txt():
    # log_print("Resampling the Image")
    # image = iio.imread(og_image_path)
    # iio.imwrite(converted_image_path, image)
    
    log_print('Resolving Captcha')
    img = cv2.imread(og_image_path)                                                                     #import image data
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)                                                        #convert to grayscale
    ret, thresh = cv2.threshold(gray, 190, 255, cv2.THRESH_BINARY_INV)                                  #threshold image
    return pytesseract.image_to_string(thresh,config="-c tessedit_char_whitelist=0123456789 --psm 10")  #img to txt


def loader():
    pass


def img_dl(img, image_path):
    img_captcha_base64 = driver.execute_script("""
        var ele = arguments[0];
        var cnv = document.createElement('canvas');
        cnv.width = 120; cnv.height = 50;
        cnv.getContext('2d').drawImage(ele, 0, 0);
        return cnv.toDataURL('image/png').substring(22);    
        """, img)
    with open(image_path, 'wb') as f:
        f.write(base64.b64decode(img_captcha_base64)) 



def captcha_input(obj, search_flag):
    captcha_img = obj.find_element(By.XPATH, "//div[@id='noprint']/img[@width='100']")
    img_dl(captcha_img, og_image_path) 
    captcha_text = img2txt()
    captcha_element = obj.find_element(By.XPATH, "//input[@name='p_captcha']")
    scrollnclick(captcha_element)
    captcha_element.send_keys(captcha_text)
    
    if search_flag:
        search_element = obj.find_element(By.XPATH, "//input[@name='btnsearch']")
        scrollnclick(search_element)
    else:
        next_page = obj.find_element(By.XPATH, "//input[@name='b_Next1']")
        scrollnclick(next_page)
        
    while True:
        obj = driver.find_element(By.XPATH, "//form[@action='nc_search']")
        try:
            error_element = obj.find_element(By.XPATH, "//b[contains(text(), 'Incorrect Code- Please try again')]")
        except NoSuchElementException:
            break
        if error_element.is_displayed():
            captcha_input(obj, search_flag)
            continue


def search(fe, letter):
    wait_search = WebDriverWait(fe, 100)
    try:
        entity_type_element = fe.find_element(By.XPATH, "//select[@name='entity_type']/option[text()='Company']")
        scrollnclick(entity_type_element)
        
        entity_name_element = fe.find_element(By.XPATH, "//input[@name='search_text']")
        scrollnclick(entity_name_element)
        entity_name_element.send_keys(Keys.CONTROL + "a")
        entity_name_element.send_keys(letter)
        
        captcha_input(fe, True)
        
    except:
        exception()


def individual_data(dt):
    try:
        data_success = False
        soup = BeautifulSoup(dt.get_attribute('innerHTML'), 'lxml')
        main_trs = soup.find_all('tr')[2]
        inner_table_rows = main_trs.find('table').find('tbody').find_all('tr')[1:]
        # main_trs = dt.find_elements(By.XPATH, "/tr")
        # inner_table = main_trs[1].find_element(By.XPATH, "/table")
        soup.decompose()
        
        if len(inner_table_rows)==0:
            data_success = True
            return data_success
        
        indi_data = []

        for row in inner_table_rows:
            cells = row.find_all('td')
            if cells[1].text.strip():
                entity_name = cells[1].text.strip()
            if cells[2].text.strip():
                entity_type = cells[2].text.strip()
            if cells[3].text.strip():
                status_registration = cells[3].find('font').text.strip()

                # Separating status and registration number using regex
                pattern = r'\s*(.*?)\s*\[ Reg\. No\. (.*?) \]'
                match = re.search(pattern, status_registration)

                if match:
                    status = match.group(1).strip()
                    registration_number = match.group(2).strip()
                else:
                    status = ""
                    registration_number = ""
                    
            data = {
            "Entity Name": entity_name,
            "Entity Type": entity_type,
            "Status": status,
            "Registration Number": registration_number
            }
            indi_data.append(data)
            
        # Write to CSV file
        csv_df = pd.DataFrame(indi_data)
        csv_df.to_csv(File_path_CSV, index=False, mode='a', header=not os.path.exists(File_path_CSV))
        count()
        # Write to TXT file
        with open(File_path_txt, "a") as f:
            for item in indi_data:
                f.write("\t".join(map(str, item.values())) + "\n")
        
        try:
            dt = driver.find_elements(By.XPATH, "//form[@action='nc_search']/table")
            captcha_input(dt[0], False)
            dt = driver.find_elements(By.XPATH, "//form[@action='nc_search']/table")
            data_success = individual_data(dt[1])
            return data_success
        except NoSuchElementException:
            data_success = True
            return data_success
        

    except:
        exception()


if __name__ == "__main__":
    File_paths = [File_path_CSV, File_path_txt, File_path_error_CSV]
    file_paths_logs = [File_path_log, File_path_log_index_LetterE1]
    
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
            
    First_run = True
    if First_run:
    # if not os.path.exists(File_path_log_Run_Flag):
    #     with open(File_path_log_Run_Flag, "a", encoding='utf-8')as f:
    #         f.write("")
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

    Headers = ['Entity Name', 'Entity Type', 'Status', 'Registration Number']
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
        Base_URL = 'https://app.roc.gov.bd/psp/nc_search'

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
        # options.add_argument('--headless')

        driver = webdriver.Chrome(options=options)
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
                            
                    form_element = driver.find_element(By.XPATH, "//form[@action='nc_search']")
                    search(form_element, letter)
                    form_tables = driver.find_elements(By.XPATH, "//form[@action='nc_search']/table")
                    success = individual_data(form_tables[1])
                    
                    if success:
                        log_print('Complete ' + letter)
                        with open(File_path_log_index_LetterE1, 'w', encoding='utf-8') as file:
                            file.write(letterE1)
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
