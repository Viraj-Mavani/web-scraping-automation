import csv
from http.client import RemoteDisconnected
import random
import time
import pandas as pd
import requests
import os
import sys
import sqlite3
import traceback
import re
import string
from sqlite3 import Error
from bs4 import BeautifulSoup
from selenium import webdriver
# import openpyxl
from requests_toolbelt import MultipartEncoder
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from requests.exceptions import RequestException
from urllib3.exceptions import ConnectTimeoutError

# BasePath= 'E:\\ADIP-PY\\'
# BasePath = 'D:\\Projects\\CedarPython\\ADIP-DZ2001-ByCRNumber'
BasePath = os.getcwd()

chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

######### Excel #########
File_path_Personnes_Physique = BasePath + \
    '\\OP\\ADIP-DZ2001-ByCRNumber_Personnes_Physique.xlsx'
File_path_Personnes_Morales = BasePath + \
    '\\OP\\ADIP-DZ2001-ByCRNumber_Personnes_Morales.xlsx'
File_path_Personnes_Physique_Arabic = BasePath + \
    '\\OP\\ADIP-DZ2001-ByCRNumber_Personnes_Physique_Arabic.xlsx'
File_path_Personnes_Morales_Arabic = BasePath + \
    '\\OP\\ADIP-DZ2001-ByCRNumber_Personnes_Morales_Arabic.xlsx'
######### Text #########
File_path_Personnes_Physique_txt = BasePath + \
    '\\OPtxt\\ADIP-DZ2001-ByCRNumber_Personnes_Physique.txt'
File_path_Personnes_Morales_txt = BasePath + \
    '\\OPtxt\\ADIP-DZ2001-ByCRNumber_Personnes_Morales.txt'
File_path_Personnes_Physique_Arabic_txt = BasePath + \
    '\\OPtxt\\ADIP-DZ2001-ByCRNumber_Personnes_Physique_Arabic.txt'
File_path_Personnes_Morales_Arabic_txt = BasePath + \
    '\\OPtxt\\ADIP-DZ2001-ByCRNumber_Personnes_Morales_Arabic.txt'
######### Input #########
# File_path_Input = BasePath + '\\InputFile\\AlgeriaInputTest.csv'
File_path_Input = BasePath + '\\InputFile\\AlgeriaInput.csv'
# File_path_Input = BasePath + '\\InputFile\\DataSet1.xlsx'
######### Failed #########
File_path_failed_English = BasePath + \
    '\\OP\\ADIP-DZ2001-ByCRNumber_Failed_English.xlsx'
File_path_failed_Arabic = BasePath + \
    '\\OP\\ADIP-DZ2001-ByCRNumber_Failed_Arabic.xlsx'
File_path_failed_English_CSV = BasePath + \
    '\\OPcsv\\ADIP-DZ2001-ByCRNumber_Failed_English.csv'
File_path_failed_Arabic_CSV = BasePath + \
    '\\OPcsv\\ADIP-DZ2001-ByCRNumber_Failed_Arabic.csv'
######### Error #########
Error_File = BasePath + '\\Error\\ADIP-DZ2001-ByCRNumber_Error.xlsx'
######### Count #########
File_path_search_count = BasePath + '\\Counts\\ADIP-DZ2001-ByCRNumber_Count.txt'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-DZ2001-ByCRNumber_Log.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-DZ2001-ByCRNumber_Run_Flag.txt'
File_path_log_index_English = BasePath + \
    '\\Log\\ADIP-DZ2001-ByCRNumber_Log_Index_English.txt'
File_path_log_index_Arabic = BasePath + \
    '\\Log\\ADIP-DZ2001-ByCRNumber_Log_Index_Arabic.txt'
######### CSV #########
Error_File_CSV = BasePath + '\\OPcsv\\ADIP-DZ2001-ByCRNumber_Error.csv'
File_path_Personnes_Physique_CSV = BasePath + \
    '\\OPcsv\\ADIP-DZ2001-ByCRNumber_Personnes_Physique.csv'
File_path_Personnes_Physique_Arabic_CSV = BasePath + \
    '\\OPcsv\\ADIP-DZ2001-ByCRNumber_Personnes_Physique_Arabic.csv'
File_path_Personnes_Morales_CSV = BasePath + \
    '\\OPcsv\\ADIP-DZ2001-ByCRNumber_Personnes_Morales.csv'
File_path_Personnes_Morales_Arabic_CSV = BasePath + \
    '\\OPcsv\\ADIP-DZ2001-ByCRNumber_Personnes_Morales_Arabic.csv'


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


def Dereference(obj):
    del obj


def log_print(message):
    with open(File_path_log, 'a', encoding='utf-8') as file:
        file.write(message + '\n')
        file.flush()
    print(message)


def exception(URL):
    # global rowError
    Headers_Error = ['URL', 'Not Responding', 'Error']
    error = traceback.format_exc()
    exception_type, exception_object, exception_traceback = sys.exc_info()
    # worksheet_error.write(rowError, 0, Base_URL)
    # worksheet_error.write(rowError, 1, "Not Responding")
    # worksheet_error.write(rowError, 2, error)
    # rowError += 1
    with open(Error_File_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(Headers_Error)
        writer.writerow([URL, "Not Responding", str(error)])
    df = pd.read_csv(Error_File_CSV, encoding='utf-8')
    df.to_excel(Error_File, index=False)


def count():
    try_count = 1
    while try_count <= 5:
        try:
            with open(File_path_search_count, 'a', encoding='utf-8') as fh:
                fh.write('1\n')
                fh.flush()
            break
        except Exception:
            pass
        try_count += 1


def duplicateFromCSV(Csv_File_path):
    try:
        data = pd.read_csv(Csv_File_path)
        unique_data = data.drop_duplicates()
        unique_data.to_csv(Csv_File_path, index=False)
    except:
        pass


def convertCSVExcel(File_path_CSV, File_path_EXL):
    chunk_size = 1000000  # Number of rows per Excel sheet (adjust as needed)
    csv_reader = pd.read_csv(
        File_path_CSV, encoding='utf-8', chunksize=chunk_size)
    sheet_index = 1  # Index of the Excel sheet
    excel_files = []  # List to store the names of generated Excel files

    for chunk in csv_reader:
        if len(chunk) > 0:  # Create Excel sheet only if chunk is not empty
            # Generate a unique sheet name
            sheet_name = f'DataSet {sheet_index}'
            # Generate a unique Excel file name
            excel_file = f'{File_path_EXL[:-5]}_{sheet_index}.xlsx'
            chunk.to_excel(excel_file, sheet_name=sheet_name, index=False)
            excel_files.append(excel_file)
            sheet_index += 1

    # Merge all Excel files into one
    writer = pd.ExcelWriter(File_path_EXL, engine='xlsxwriter')
    Sheet = 1
    for file in excel_files:
        df = pd.read_excel(file)
        sheet_name = f'DataSet {Sheet}'
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        Sheet += 1
    writer.close()


def duplicate(File_path):
    try:
        data = pd.read_excel(File_path)
        data_file = data.drop_duplicates()
        data_file.to_excel(File_path, index=False)
    except:
        pass


if __name__ == '__main__':
    File_paths = [File_path_Personnes_Physique, File_path_Personnes_Morales,
                  File_path_Personnes_Physique_Arabic, File_path_Personnes_Morales_Arabic]

    # First_run = False
    # if First_run:
    if not os.path.exists(File_path_log_Run_Flag):
        with open(File_path_log_Run_Flag, "a", encoding='utf-8')as f:
            f.write("")
        File_paths_csv = [File_path_Personnes_Physique_CSV, File_path_Personnes_Morales_CSV,
                          File_path_Personnes_Physique_Arabic_CSV, File_path_Personnes_Morales_Arabic_CSV]
        File_paths_txt = [File_path_Personnes_Physique_txt, File_path_Personnes_Morales_txt,
                          File_path_Personnes_Physique_Arabic_txt, File_path_Personnes_Morales_Arabic_txt]
        File_path_index = [File_path_log_index_English,
                           File_path_log_index_Arabic]
        File_paths_error = [
            Error_File_CSV, File_path_failed_English_CSV, File_path_failed_Arabic_CSV]
        if os.path.exists(File_path_search_count):
            os.remove(File_path_search_count)
        if os.path.exists(File_path_log):
            os.remove(File_path_log)
        for path_csv in File_paths_csv:
            if os.path.exists(path_csv):
                os.remove(path_csv)
        for Path_txt in File_paths_txt:
            if os.path.exists(Path_txt):
                os.remove(Path_txt)
        for Path_ere in File_paths_error:
            if os.path.exists(Path_ere):
                os.remove(Path_ere)
        for Path_index in File_path_index:
            if os.path.exists(Path_index):
                os.remove(Path_index)

    # Create directories if they don't exist
    directories = [
        BasePath + '\\OP',
        BasePath + '\\OPtxt',
        BasePath + '\\OPcsv',
        BasePath + '\\InputFile',
        BasePath + '\\Error',
        BasePath + '\\Counts',
        BasePath + '\\Log'
    ]

    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)

    Personnes_Physique_headers = ['NRC', 'Nom', 'Prenom']
    Personnes_Physique_Arabic_headers = [
        'NRC', 'Nom (Arabic)', 'Prenom (Arabic)']
    Personnes_Morales_headers = ['NRC', 'Raison Sociale']
    Personnes_Morales_Arabic_headers = ['NRC', 'Raison Sociale (Arabic)']

    if not os.path.exists(File_path_search_count):
        with open(File_path_search_count, "a", encoding='utf-8')as f:
            f.write("")
    with open(File_path_Personnes_Physique_txt, "a", encoding='utf-8')as f:
        if f.tell() == 0:
            f.write("\t".join(Personnes_Physique_headers)+"\n")
            f.flush()
    with open(File_path_Personnes_Morales_txt, "a", encoding='utf-8')as fw:
        if fw.tell() == 0:
            fw.write("\t".join(Personnes_Morales_headers)+"\n")
            fw.flush()
    with open(File_path_Personnes_Physique_Arabic_txt, "a", encoding='utf-8')as f:
        if f.tell() == 0:
            f.write("\t".join(Personnes_Physique_Arabic_headers)+"\n")
            f.flush()
    with open(File_path_Personnes_Morales_Arabic_txt, "a", encoding='utf-8')as fw:
        if fw.tell() == 0:
            fw.write("\t".join(Personnes_Morales_Arabic_headers)+"\n")
            fw.flush()

    with open(File_path_Personnes_Physique_CSV, "a", newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if f.tell() == 0:
            writer.writerow(Personnes_Physique_headers)
    with open(File_path_Personnes_Morales_CSV, "a", newline='', encoding='utf-8') as fw:
        writer = csv.writer(fw)
        if fw.tell() == 0:
            writer.writerow(Personnes_Morales_headers)
    with open(File_path_Personnes_Physique_Arabic_CSV, "a", newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if f.tell() == 0:
            writer.writerow(Personnes_Physique_Arabic_headers)
    with open(File_path_Personnes_Morales_Arabic_CSV, "a", newline='', encoding='utf-8') as fw:
        writer = csv.writer(fw)
        if fw.tell() == 0:
            writer.writerow(Personnes_Morales_Arabic_headers)

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

    Home_URL = 'https://sidjilcom.cnrc.dz/web/cnrc/accueil'
    Arabic_url = 'https://sidjilcom.cnrc.dz/accueil?p_p_id=82&p_p_lifecycle=1&p_p_state=normal&p_p_mode=view&p_p_col_id=column-5&p_p_col_count=8&_82_struts_action=%2Flanguage%2Fview&_82_redirect=%2Faccueil&_82_languageId=ar_SA'

    retry_attempts = 3
    retry_delay = 2

    log_print('Data Importing...plz wait')
    df = pd.read_csv(File_path_Input, header=None)
    number_list = df.iloc[:, 0].tolist()
    # df = pd.read_excel(File_path_Input, sheet_name='Sheet1', header=None)
    # number_list = df.iloc[:, 0].tolist()
    log_print('Data Imported\n')
    # number_list = ["CN-100002", "CN-1000104", "CN-1000141"]

    ########################################### Number For English ###########################################

    try:
        Driver = webdriver.Chrome(service=Service(
            ChromeDriverManager().install()), options=chrome_options)
        Driver.get(Home_URL)

        sid = Driver.get_cookie(
            'SID')['name'] + '=' + Driver.get_cookie('SID')['value'] + ';'
        gid = Driver.get_cookie(
            '_gid')['name'] + '=' + Driver.get_cookie('_gid')['value'] + ';'
        ga = Driver.get_cookie('_ga')['name'] + '=' + \
            Driver.get_cookie('_ga')['value'] + ';'
        session = Driver.get_cookie('cookiesession1')[
            'name'] + '=' + Driver.get_cookie('cookiesession1')['value'] + ';'
        support = Driver.get_cookie('COOKIE_SUPPORT')[
            'name'] + '=' + Driver.get_cookie('COOKIE_SUPPORT')['value'] + ';'
        lang = Driver.get_cookie('GUEST_LANGUAGE_ID')[
            'name'] + '=' + Driver.get_cookie('GUEST_LANGUAGE_ID')['value'] + ';'
        Jsess = Driver.get_cookie('JSESSIONID')[
            'name'] + '=' + Driver.get_cookie('JSESSIONID')['value'] + ';'

        Home_soup = BeautifulSoup(Driver.page_source.encode(), 'html.parser')
        Link_Form = Home_soup.find('form', id='f1').get('action')
        Driver.close()
        Driver.quit()

        log_index_flag = False
        if os.path.exists(File_path_log_index_English):
            log_index_flag = True
            with open(File_path_log_index_English, 'r', encoding='utf-8') as file:
                last_processed_number = file.read().strip()

        if log_index_flag:
            start_index = number_list.index(last_processed_number) + 1
            numbers = number_list[start_index:]
        else:
            numbers = number_list

        for numberE in numbers[:]:
            js_found_flag = False
            try:
                fields = {'hidden': 'goRecherche', 'critere': numberE}
                Form_Data = MultipartEncoder(
                    fields=fields, boundary='----WebKitFormBoundaryTdS000SBKNSpDEkf')
                Headers = {'Cookie': support + ' ' + lang + ' ' + session + ' ' + ga + ' ' + gid + ' ' + Jsess + ' ' + sid,
                           'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
                           'Accept-Encoding': 'gzip, deflate, br',
                           'Accept-Language': 'en-IN,en-GB;q=0.9,en-US;q=0.8,en;q=0.7',
                           'Cache-Control': 'max-age=0',
                           'Connection': 'keep-alive',
                           'Content-Length': '243',
                           'Content-Type': 'multipart/form-data; boundary=----WebKitFormBoundaryTdS000SBKNSpDEkf',
                           'Host': 'sidjilcom.cnrc.dz',
                           'Origin': 'https://sidjilcom.cnrc.dz',
                           'Referer': Link_Form,
                           'sec-ch-ua': '\"Google Chrome\";v=\"111\", \"Not(A:Brand\";v=\"8\", \"Chromium\";v=\"111\"',
                           'sec-ch-ua-mobile': '?0',
                           'sec-ch-ua-platform': '\"Windows\"',
                           'Sec-Fetch-Dest': 'document',
                           'Sec-Fetch-Mode': 'navigate',
                           'Sec-Fetch-Site': 'same-origin',
                           'Sec-Fetch-User': '?1',
                           'Upgrade-Insecure-Requests': '1',
                           # 'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36'}
                           'User-Agent': user_agent}

                engRetry = 1
                while engRetry <= retry_attempts:
                    try:
                        Info = requests.post(
                            Link_Form, data=Form_Data, headers=Headers, timeout=200)
                        break
                    # except Exception as e:
                    # 	log_print(f"Error occurred for {letter}")
                    # 	exception(Home_URL)
                    # 	log_print(str(e))
                    # 	os._exit(1)
                    except Exception as e:
                        log_print(
                            f"Error occurred for {numberE}...Retrying in 2 min")
                        exception(Home_URL)
                        time.sleep(120)
                        engRetry += 1
                        break
                        # os._exit(1)

                Data_soup = BeautifulSoup(Info.content, 'html.parser')
                pattern = r'\$\(\s*function\(\)\s*{\s*\$\.\w+\({'

                if len(Data_soup.find_all('script', type='text/javascript')) > 1:
                    for scripts in Data_soup.find_all('script', type='text/javascript'):
                        if scripts.string == None:
                            js_found_flag = False
                            continue
                        if re.match(pattern, scripts.string):
                            Script = scripts.string
                            if Script == None:
                                js_found_flag = False
                                continue
                            Script_soup = BeautifulSoup(Script, 'html.parser')
                            js_found_flag = True
                            table1 = Script_soup.find(
                                'div', id='tab1').find('table')
                            if table1 is not None:
                                tab1_rows = table1.find('tbody').find_all('tr')
                                for row1 in tab1_rows:
                                    Personnes_Physique_data = []
                                    cells = row1.find_all('td')
                                    Personnes_Physique_data.append(
                                        cells[0].string.strip() if cells[0].string else '')
                                    Personnes_Physique_data.append(
                                        cells[1].string.strip() if cells[1].string else '')
                                    Personnes_Physique_data.append(
                                        cells[2].string.strip() if cells[2].string else '')
                                    with open(File_path_Personnes_Physique_CSV, 'a', newline='', encoding='utf-8') as file:
                                        writer = csv.writer(file)
                                        writer.writerow(
                                            Personnes_Physique_data)
                                    with open(File_path_Personnes_Physique_txt, "a", encoding='utf-8')as f:
                                        f.write(
                                            "\t".join(map(str, Personnes_Physique_data))+"\n")
                                        f.flush()
                                    count()
                            table2 = Script_soup.find(
                                'div', id='tab2').find('table')
                            if table2 is not None:
                                tab2_rows = table2.find('tbody').find_all('tr')
                                for row2 in tab2_rows:
                                    cells = row2.find_all('td')
                                    Personnes_Morales_data = []
                                    Personnes_Morales_data.append(
                                        cells[0].string.strip() if cells[0].string else '')
                                    Personnes_Morales_data.append(
                                        cells[1].string.strip() if cells[1].string else '')
                                    with open(File_path_Personnes_Morales_CSV, 'a', newline='', encoding='utf-8') as file:
                                        writer = csv.writer(file)
                                        writer.writerow(Personnes_Morales_data)
                                    with open(File_path_Personnes_Morales_txt, "a", encoding='utf-8')as f:
                                        f.write(
                                            "\t".join(map(str, Personnes_Morales_data))+"\n")
                                        f.flush()
                                    count()
                            if table1 == None and table2 == None:
                                # js_found_flag = False
                                break
                            with open(File_path_log_index_English, 'w', encoding='utf-8') as file:
                                file.write(numberE)
                                file.flush()
                            break
                    if not js_found_flag:
                        log_print(f"Failed!! Couldn't find JS for {numberE}")
                        os._exit(1)
                    # else:
                    #     log_print(f"Couldn't find JS for {numberE}")
                    #     os._exit(1)
                if os.path.exists(File_path_log_index_English):
                    with open(File_path_log_index_English, 'r', encoding='utf-8') as file:
                        last_number = file.read().strip()
                if numberE == last_number:
                    log_print(f'Complete {numberE} in English')
                else:
                    log_print(f'Failed!! {numberE} in English')
                    if not js_found_flag:
                        log_print(f"Failed!! Couldn't find JS for {numberE}")
                        os._exit(1)
                    with open(File_path_failed_English_CSV, 'a', newline='', encoding='utf-8') as file:
                        writer = csv.writer(file)
                        writer.writerow([numberE])
            except:
                exception(Home_URL)

    except:
        exception(Home_URL)

    finally:
        duplicateFromCSV(File_path_Personnes_Physique_CSV)
        duplicateFromCSV(File_path_Personnes_Morales_CSV)
        convertCSVExcel(File_path_Personnes_Physique_CSV,
                        File_path_Personnes_Physique)
        convertCSVExcel(File_path_Personnes_Morales_CSV,
                        File_path_Personnes_Morales)
        Home_soup.decompose()
        # duplicate(File_path_Personnes_Physique)
        # duplicate(File_path_Personnes_Morales)

    ########################################### Number For Arabic ###########################################

    try:
        Arabic_Driver = webdriver.Chrome(service=Service(
            ChromeDriverManager().install()), options=chrome_options)
        Arabic_Driver.get(Arabic_url)
        Arabic_url = Arabic_Driver.current_url
        Arabic_Driver.get(Arabic_url)

        Arabic_sid = Arabic_Driver.get_cookie(
            'SID')['name'] + '=' + Arabic_Driver.get_cookie('SID')['value'] + ';'
        Arabic_gid = Arabic_Driver.get_cookie(
            '_gid')['name'] + '=' + Arabic_Driver.get_cookie('_gid')['value'] + ';'
        Arabic_ga = Arabic_Driver.get_cookie(
            '_ga')['name'] + '=' + Arabic_Driver.get_cookie('_ga')['value'] + ';'
        Arabic_session = Arabic_Driver.get_cookie('cookiesession1')[
            'name'] + '=' + Arabic_Driver.get_cookie('cookiesession1')['value'] + ';'
        Arabic_support = Arabic_Driver.get_cookie('COOKIE_SUPPORT')[
            'name'] + '=' + Arabic_Driver.get_cookie('COOKIE_SUPPORT')['value'] + ';'
        Arabic_lang = Arabic_Driver.get_cookie('GUEST_LANGUAGE_ID')[
            'name'] + '=' + Arabic_Driver.get_cookie('GUEST_LANGUAGE_ID')['value'] + ';'
        Arabic_Jsess = Arabic_Driver.get_cookie('JSESSIONID')[
            'name'] + '=' + Arabic_Driver.get_cookie('JSESSIONID')['value'] + ';'

        Ar_Home_soup = BeautifulSoup(
            Arabic_Driver.page_source.encode(), 'html.parser')
        Ar_Link_Form = Ar_Home_soup.find('form', id='f1').get('action')
        Arabic_Driver.close()
        Arabic_Driver.quit()

        log_index_flag = False
        if os.path.exists(File_path_log_index_Arabic):
            log_index_flag = True
            with open(File_path_log_index_Arabic, 'r', encoding='utf-8') as file:
                last_processed_number = file.read().strip()

        if log_index_flag:
            start_index = number_list.index(last_processed_number) + 1
            numbers = number_list[start_index:]
        else:
            numbers = number_list

        for numberA in numbers[:]:
            try:
                fields = {'hidden': 'goRecherche', 'critere': numberA}
                Form_Data = MultipartEncoder(
                    fields=fields, boundary='----WebKitFormBoundaryTdS000SBKNSpDEkf')
                Headers = {'Cookie': Arabic_support + ' ' + Arabic_lang + ' ' + Arabic_session + ' ' + Arabic_ga + ' ' + Arabic_gid + ' ' + Arabic_Jsess + ' ' + Arabic_sid,
                           'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
                           'Accept-Encoding': 'gzip, deflate, br',
                           'Accept-Language': 'en-IN,en-GB;q=0.9,en-US;q=0.8,en;q=0.7',
                           'Cache-Control': 'max-age=0',
                           'Connection': 'keep-alive',
                           'Content-Length': '243',
                           'Content-Type': 'multipart/form-data; boundary=----WebKitFormBoundaryTdS000SBKNSpDEkf',
                           'Host': 'sidjilcom.cnrc.dz',
                           'Origin': 'https://sidjilcom.cnrc.dz',
                           'Referer': Ar_Link_Form,
                           'sec-ch-ua': '\"Google Chrome\";v=\"111\", \"Not(A:Brand\";v=\"8\", \"Chromium\";v=\"111\"',
                           'sec-ch-ua-mobile': '?0',
                           'sec-ch-ua-platform': '\"Windows\"',
                           'Sec-Fetch-Dest': 'document',
                           'Sec-Fetch-Mode': 'navigate',
                           'Sec-Fetch-Site': 'same-origin',
                           'Sec-Fetch-User': '?1',
                           'Upgrade-Insecure-Requests': '1',
                           'User-Agent': random.choice(user_agents)}

                araRetry = 1
                while araRetry <= retry_attempts:
                    try:
                        Ar_Info = requests.post(
                            Ar_Link_Form, data=Form_Data, headers=Headers, timeout=200)
                        break
                    # except Exception as e:
                    # 	log_print(f"Error occurred for {letter}")
                    # 	exception(Home_URL)
                    # 	log_print(str(e))
                    # 	os._exit(1)
                    except Exception as e:
                        log_print(
                            f"Error occurred for {numberA}...Retrying in 2 min")
                        exception(Arabic_url)
                        time.sleep(120)
                        break
                        # os._exit(1)

                Ar_Data_soup = BeautifulSoup(Ar_Info.content, 'html.parser')

                pattern = r'\$\(\s*function\(\)\s*{\s*\$\.\w+\({'
                if len(Ar_Data_soup.find_all('script', type='text/javascript')) > 1:
                    for scripts in Ar_Data_soup.find_all('script', type='text/javascript'):
                        if scripts.string == None:
                            continue
                        if re.match(pattern, scripts.string):
                            Ar_Script = scripts.string
                            if Ar_Script == None:
                                continue
                            Ar_Script_soup = BeautifulSoup(
                                Ar_Script, 'html.parser')
                            # Ar_Data = Ar_Script_soup.find_all('tr')
                            table1 = Ar_Script_soup.find(
                                'div', id='tab1').find('table')
                            if table1 is not None:
                                tab1_rows = table1.find('tbody').find_all('tr')
                                for row1 in tab1_rows:
                                    Personnes_Physique_data = []
                                    cells = row1.find_all('td')
                                    Personnes_Physique_data.append(
                                        cells[0].string.strip() if cells[0].string else '')
                                    Personnes_Physique_data.append(
                                        cells[1].string.strip() if cells[1].string else '')
                                    Personnes_Physique_data.append(
                                        cells[2].string.strip() if cells[2].string else '')
                                    with open(File_path_Personnes_Physique_Arabic_CSV, 'a', newline='', encoding='utf-8') as file:
                                        writer = csv.writer(file)
                                        writer.writerow(
                                            Personnes_Physique_data)
                                    with open(File_path_Personnes_Physique_Arabic_txt, "a", encoding='utf-8')as f:
                                        f.write(
                                            "\t".join(map(str, Personnes_Physique_data))+"\n")
                                        f.flush()
                                    count()
                            table2 = Ar_Script_soup.find(
                                'div', id='tab2').find('table')
                            if table2 is not None:
                                tab2_rows = table2.find('tbody').find_all('tr')
                                for row2 in tab2_rows:
                                    cells = row2.find_all('td')
                                    Personnes_Morales_data = []
                                    Personnes_Morales_data.append(
                                        cells[0].string.strip() if cells[0].string else '')
                                    Personnes_Morales_data.append(
                                        cells[1].string.strip() if cells[1].string else '')
                                    with open(File_path_Personnes_Morales_Arabic_CSV, 'a', newline='', encoding='utf-8') as file:
                                        writer = csv.writer(file)
                                        writer.writerow(Personnes_Morales_data)
                                    with open(File_path_Personnes_Morales_Arabic_txt, "a", encoding='utf-8')as f:
                                        f.write(
                                            "\t".join(map(str, Personnes_Morales_data))+"\n")
                                        f.flush()
                                    count()
                            with open(File_path_log_index_Arabic, 'w', encoding='utf-8') as file:
                                file.write(numberA)
                                file.flush()
                            break
                    else:
                        log_print(f"Couldn't find JS for {numberE}")
                        os._exit(1)
                log_print('Complete ' + numberA)
                if os.path.exists(File_path_log_index_Arabic):
                    with open(File_path_log_index_Arabic, 'r', encoding='utf-8') as file:
                        last_number = file.read().strip()
                if numberA == last_number:
                    log_print(f'Complete {numberA} in Arabic')
                else:
                    log_print(f'Failed!! {numberA} in Arabic')
                    with open(File_path_failed_Arabic_CSV, 'a', newline='', encoding='utf-8') as file:
                        writer = csv.writer(file)
                        writer.writerow([numberA])
            except:
                exception(Arabic_url)
    except:
        exception(Arabic_url)

    finally:
        duplicateFromCSV(File_path_Personnes_Physique_Arabic_CSV)
        duplicateFromCSV(File_path_Personnes_Morales_Arabic_CSV)
        convertCSVExcel(File_path_Personnes_Physique_Arabic_CSV,
                        File_path_Personnes_Physique_Arabic)
        convertCSVExcel(File_path_Personnes_Morales_Arabic_CSV,
                        File_path_Personnes_Morales_Arabic)
        Ar_Home_soup.decompose()
        # duplicate(File_path_Personnes_Physique_Arabic)
        # duplicate(File_path_Personnes_Morales_Arabic)
