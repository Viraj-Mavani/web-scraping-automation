import ast
import csv
import os
import random
import sys
import traceback
import pandas as pd
import sqlite3
import re
from sqlite3 import Error
import requests
from bs4 import BeautifulSoup
import time
from requests.exceptions import RequestException
from urllib3.exceptions import ConnectTimeoutError
# import chromedriver_autoinstaller
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.common.exceptions import NoSuchElementException, TimeoutException
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.support import expected_conditions as EC
# from selenium.webdriver.common.action_chains import ActionChains

# BasePath = 'D:\\Projects\\CedarPython\\ADIP-BF2801\\'
BasePath = os.getcwd()

######### Excel #########
File_path_XL = BasePath + '\\OP\\ADIP-BF2801_Output.xlsx'
######### CSV #########
File_path_CSV = BasePath + '\\OPcsv\\ADIP-BF2801_Output.csv'
File_path_error_CSV = BasePath + '\\OPcsv\\ADIP-BF2801_Error.csv'
######### Text #########
File_path_TXT = BasePath + '\\OPtxt\\ADIP-BF2801_Output.txt'
######### Error #########
File_path_error = BasePath + '\\Error\\ADIP-BF2801_Error.xlsx'
######### Count #########
File_path_count = BasePath + '\\Counts\\ADIP-BF2801_Count.txt'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-BF2801_Log.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-BF2801_Run_Flag.txt'
File_path_log_index = BasePath + '\\Log\\ADIP-BF2801_Log_Index.txt'


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
    Headers_Error = ['URL', 'Not Responding', 'Error']
    error = traceback.format_exc()
    exception_type, exception_object, exception_traceback = sys.exc_info()
    with open(File_path_error_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(Headers_Error)
        writer.writerow([Base_URL, "Not Responding", error])
    df = pd.read_csv(File_path_error_CSV, encoding='utf-8')
    df.to_excel(File_path_error, index=False)


def Dereference(obj):
    del obj


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


def convertCSVExcel(csv, exl):
    df = pd.read_csv(csv, encoding='utf-8')
    df.to_excel(exl, index=False)
    

def duplicate(exl):
    try:
        data = pd.read_excel(exl)
        data_file = data.drop_duplicates()
        data_file.to_excel(exl, index=False)
    except:
        pass


def individual_data(rows):
    indi_data = []
    data_mapping = {
        'views-field-title': 'name',
        'views-field-field-forme-juridique': 'forme_juridique',
        'views-field-field-al-rccm': 'rccm',
        'views-field-field-ent-acte-rccm': 'acte_rccm',
        'views-field-field-al-date-rccm': 'date_rccm'
    }

    for row in rows:
        try:
            data = {}
            for class_name, var_name in data_mapping.items():
                element = row.find('td', class_=class_name)
                data[var_name] = element.get_text(strip=True) if element else ''

            indi_data.append(data)

        except:
            exception()

    # Write to CSV file
    with open(File_path_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        for data in indi_data:
            writer.writerow(list(data.values()))
            count()

    # Write to TXT file
    with open(File_path_TXT, "a", encoding='utf-8') as f:
        for data in indi_data:
            line = "\t".join(data.values())
            f.write(line + "\n")
            f.flush()
            
            
if __name__ == '__main__':
    
    File_paths = [File_path_XL, File_path_CSV, File_path_error_CSV, File_path_TXT, File_path_error,
                File_path_count, File_path_log, File_path_log_Run_Flag]

    directories = [
        BasePath + '\\Counts',
        BasePath + '\\Error',
        BasePath + '\\Log',
        BasePath + '\\OP',
        BasePath + '\\OPcsv',
        BasePath + '\\OPtxt'
    ]

    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)

    # First_run = True
    # if First_run:
    if not os.path.exists(File_path_log_Run_Flag):
        with open(File_path_log_Run_Flag, "a", encoding='utf-8')as f:
            f.write("")
        if os.path.exists(File_path_log):
            os.remove(File_path_log)
        if os.path.isfile(File_path_error_CSV):
            os.remove(File_path_error_CSV)
        if os.path.isfile(File_path_CSV):
            os.remove(File_path_CSV)
        if os.path.isfile(File_path_TXT):
            os.remove(File_path_TXT)
        if os.path.isfile(File_path_count):
            os.remove(File_path_count)
        if os.path.isfile(File_path_log_index):
            os.remove(File_path_log_index)

    Headers = ['Raison sociale(asc)', 'Forme juridique', 'RCCM', 'Acte RCCM', 'Date RCCM']
    with open(File_path_count, "a", encoding='utf-8')as f:
            f.write("")
    with open(File_path_TXT, "a", encoding='utf-8')as f:
        if f.tell() == 0:
            f.write("\t".join(Headers)+"\n")
            f.flush()
    with open(File_path_CSV, "a", newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if f.tell() == 0:
            writer.writerow(Headers)

    user_agents = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:53.0) Gecko/20100101 Firefox/53.0",
        "Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.0; Trident/5.0; Trident/5.0)",
        "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.2; Trident/6.0; MDDCJS)",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/51.0.2704.79 Safari/537.36 Edge/14.14393",
        "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1; SV1)",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/89.0",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Safari/537.36 Edg/92.0.902.55",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.107 Safari/537.36",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Firefox/89.0",
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Safari/537.36 Edg/92.0.902.55"
    ]
    user_agent = random.choice(user_agents)
    Headers = {'User-Agent': user_agent}

    Base_URL = 'http://www.me.bf/en/annonces-legales?keys=&order=title&sort=asc'
    retry_attempts = 5
    retry_delay = 2
        
    try:
        ##################### Last Page Count #####################
        tempRetry = 1
        error_message_flag = False
        while tempRetry <= retry_attempts:
            try:
                # obj_temp = requests.get(Base_url.format(letter, 1), proxies={'http': proxy_url, 'https': proxy_url})
                obj_temp = requests.get(Base_URL)
                break
            except (ConnectTimeoutError, RequestException) as e:
                exception()
                delay = retry_delay * (2 ** tempRetry)
                log_print(f'Retrying in {delay} seconds...{tempRetry}')
                time.sleep(delay)
                tempRetry += 1
                continue
        soup_temp = BeautifulSoup(obj_temp.content, 'html.parser')
        
        pagination_element = soup_temp.find('section', class_ = 'col-sm-8').find('ul', class_ = 'pagination')
        if pagination_element is not None:
            last_page_element = pagination_element.find('li', class_='pager-last').find('a')
            last_page_link = last_page_element['href']
            last_page_number = re.search(r'page=(\d+)', last_page_link)
            if last_page_number:
                last_page_number = int(last_page_number.group(1))
            else:
                log_print("Last page number not found.")
                exit(1)
        
        del obj_temp
        del soup_temp
        
        ##################### last Processed Page #####################
        
        log_index_flag = False
        if os.path.exists(File_path_log_index):
            log_index_flag = True
            with open(File_path_log_index, 'r', encoding='utf-8') as file:
                last_processed_page = int(file.read().strip())

        if log_index_flag:
            start_index = last_processed_page + 1
        else:
            start_index = 0
        
        ##################### Pages #####################
        
        for index in range(start_index, last_page_number + 1):
            # time.sleep(10)
            user_agent = random.choice(user_agents)
            Headers = {'User-Agent': user_agent}
            Page_URL = Base_URL + f'&page={index}'
            if index == 0:
                Page_URL = Base_URL
            innerRetry = 1
            while innerRetry <= retry_attempts:
                try:
                    obj = requests.get(Page_URL, headers=Headers, timeout=200)
                    break
                except Exception as e:
                    log_print(f"Error occurred")
                    Headers = {'User-Agent': user_agent}
                    delay = retry_delay * (2 ** innerRetry)
                    log_print(f'Retrying in {delay} seconds...RETRY: {innerRetry}')
                    time.sleep(delay)
                    innerRetry += 1
                    continue
            else:
                exit(1)

            soup = BeautifulSoup(obj.content, 'html.parser')
            res = soup.find('table', class_='views-table cols-6 table table-hover table-striped').find('tbody')
            rows = res.find_all('tr')
            individual_data(rows)
            with open(File_path_log_index, 'w', encoding='utf-8') as file:
                file.write(str(index))
                file.flush()
            log_print(f'Completed Page {index+1}')
        
        

    finally:
        convertCSVExcel(File_path_CSV, File_path_XL)
        duplicate(File_path_XL)
        if os.path.exists(File_path_log_index):
            with open(File_path_log_index, 'r', encoding = 'utf-8') as file:
                last_processed_page = file.read().strip()
        if int(last_processed_page) == last_page_number:
            log_print('Script Completed')
            if os.path.exists(File_path_log_Run_Flag):
                os.remove(File_path_log_Run_Flag)
            if os.path.exists(File_path_count):
                os.remove(File_path_count)
        else:
            log_print('Stopped')
        
database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
    for File_path in File_paths:
        delete_task(conn, File_path)
