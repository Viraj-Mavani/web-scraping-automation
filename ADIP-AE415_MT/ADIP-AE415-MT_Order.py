import csv
import json
import os
import random
import subprocess
import sys
import traceback
import pandas as pd
import sqlite3
import re
from sqlite3 import Error
# from bs4 import BeautifulSoup
import time
import requests
import threading


# BasePath= 'E:\\ADIP-PY'
BasePath = 'D:\\Projects\\CedarPython\\ADIP-AE415_Copy'
# BasePath = os.getcwd()

######### Excel #########
File_path_English = BasePath + '\\OP\\ADIP-AE415_Details.xlsx'
# File_path_Arabic = BasePath + '\\OP\\ADIP-AE415_Arabic.xlsx'
File_path_Activity = BasePath + '\\OP\\ADIP-AE415_Activity.xlsx'
######### Failed #########
File_path_failed = BasePath + '\\OP\\ADIP-AE415_Failed.xlsx'
File_path_failed_CSV = BasePath + '\\OPcsv\\ADIP-AE415_Failed.csv'
######### Input #########
# File_path_Input = BasePath + '\\InputFile\\ADIP-AE415-Input-Updated.xlsx'
File_path_Input = BasePath + '\\InputFile\\ADIP-AE415-Input-sample.xlsx'
######### CSV #########
File_path_English_CSV = BasePath + '\\OPcsv\\ADIP-AE415_Details.csv'
# File_path_Arabic_CSV = BasePath + '\\OPcsv\\ADIP-AE415_Arabic.csv'
File_path_Activity_CSV = BasePath + '\\OPcsv\\ADIP-AE415_Activity.csv'
File_path_error_CSV = BasePath + '\\OPcsv\\ADIP-AE415_Error.csv'
######### Text #########
File_path_English_txt = BasePath + '\\Optxt\\ADIP-AE415_Details.txt'
# File_path_Arabic_txt = BasePath + '\\Optxt\\ADIP-AE415_Arabic.txt'
File_path_Activity_txt = BasePath + '\\Optxt\\ADIP-AE415_Activity.txt'
######### Error #########
File_path_error = BasePath + '\\Error\\ADIP-AE415_Error.xlsx'
######### Count #########
File_path_count = BasePath + '\\Counts\\ADIP-AE415_Count.txt'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-AE415_Log.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-AE415_Run_Flag.txt'
File_path_log_index = BasePath + '\\Log\\ADIP-AE415_Log_Index.txt'


def create_connection(db_file):
    """ create a database connection to the SQLite database
                specified by the db_file
        :param db_file: database file
        :return: Connection object or None
        """
    conn = None
    try:
        conn = sqlite3.connect(db_file)
    except Exception as e:
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


def exception(licence_num):
    Headers_Error = ['licence number', 'URL', 'Not Responding', 'Error']
    error = traceback.format_exc()
    exception_type, exception_object, exception_traceback = sys.exc_info()
    with open(File_path_error_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(Headers_Error)
        writer.writerow([licence_num, Base_URL, "Not Responding", error])
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


# def restart_script():
#     python = sys.executable
#     subprocess.call([python] + sys.argv)


def request(url, Header):
    
    try:
        Retry = 1
        while Retry <= retry_attempts:
            try:
                # url = Base_URL.format('CN-1000141')
                r_delay = random.uniform(0.5, 3.0)
                time.sleep(r_delay)
                obj = requests.post(url, timeout=200)
                json_data = obj.json()
                return json_data
            except Exception:
                exception()
                log_print(f"Error occurred in Request")
                user_agent = random.choice(user_agents)
                Header = {'User-Agent': user_agent}
                delay = retry_delay * (2 ** Retry)
                log_print(f'Retrying in {delay} seconds...{Retry}')
                time.sleep(delay)
                Retry += 1
                continue
        else:
            log_print('\n\Requests Failed!!\nTerminating the script...\n===========================================================')
            os._exit(1)
            # log_print('\n\Request Failed!!\nRestarting the script in 5 min...\n===========================================================')
            # time.sleep(300)
            # restart_script()
    except:
        exception()


def individual_data(lnum, data_json, invalid_flag):
    indi_data_EN = ['']*17
    # indi_data_AR = ['']*2
    
    if not invalid_flag:
        try:
            if data_json.get("message") == 'Success':
                data = data_json.get("data", {})
                results = data.get("result", {})
                for result in results:
                    if 'licenseNo' in result:
                        indi_data_EN[0] = results['licenseNo']
                        # indi_data_AR[0] = results['licenseNo']
                    elif 'adcciNo' in result:
                        indi_data_EN[1] = results['adcciNo']
                    elif 'businessNameEn' in result:
                        indi_data_EN[2] = results['businessNameEn']
                    elif 'businessNameAr' in result:
                        indi_data_EN[3] = results['businessNameAr']
                    elif 'legalFormEn' in result:
                        indi_data_EN[4] = results['legalFormEn']
                    elif 'licenseTypeEn' in result:
                        indi_data_EN[5] = results['licenseTypeEn']
                    elif 'isBranch' in result:
                        if results['isBranch'] == "Y":
                            indi_data_EN[6] = "Yes"
                        elif results['isBranch'] == "N":
                            indi_data_EN[6] = "No"
                    elif 'issuePlaceEn' in result:
                        indi_data_EN[7] = results['issuePlaceEn']
                    elif 'establishmentDate' in result:
                        indi_data_EN[8] = results['establishmentDate']
                    elif 'issueDate' in result:
                        indi_data_EN[9] = results['issueDate']
                    elif 'expiryDate' in result:
                        indi_data_EN[10] = results['expiryDate']
                    elif 'licenseStatusEn' in result:
                        indi_data_EN[11] = results['licenseStatusEn']
                    elif 'businessAddressEn' in result:
                        indi_data_EN[12] = results['businessAddressEn']
                    elif 'establishmentVolumeEn' in result:
                        indi_data_EN[13] = results['establishmentVolumeEn']
                    elif 'socialMediaAcount' in result:
                        indi_data_EN[14] = results['socialMediaAcount']
                    elif 'socialMediaType' in result:
                        indi_data_EN[15] = results['socialMediaType']
                    elif 'socialMediaWeb' in result:
                        indi_data_EN[16] = results['socialMediaWeb']
                
                with open(File_path_English_CSV, 'a', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow(indi_data_EN)
                count()
                with open(File_path_English_txt, 'a', encoding="utf-8") as fw:
                    fw.write("\t".join(map(str, indi_data_EN)) + "\n")
                    fw.flush()
                
                # with open(File_path_Arabic_CSV, 'a', newline='', encoding='utf-8') as file:
                #     writer = csv.writer(file)
                #     writer.writerow(indi_data_AR)
                # with open(File_path_Arabic_txt, 'a', encoding="utf-8") as fw:
                #     fw.write("\t".join(map(str, indi_data_AR)) + "\n")
                #     fw.flush()
                
                activities = results.get("activities", [])
                for activity in activities:
                    indi_data_AC = ['']*3
                    indi_data_AC[0] = lnum
                    indi_data_AC[1] = activity.get("activityCode", "")
                    indi_data_AC[2] = activity.get("activityNameEng", "")
                    with open(File_path_Activity_CSV, 'a', newline='', encoding='utf-8') as file:
                        writer = csv.writer(file)
                        writer.writerow(indi_data_AC)
                    with open(File_path_Activity_txt, 'a', encoding="utf-8") as fw:
                        fw.write("\t".join(map(str, indi_data_AC)) + "\n")
                        fw.flush()
                        
                with open(File_path_log_index, 'w', encoding='utf-8') as file:
                    file.write(lnum)
                    file.flush()

        except:
            exception()

    
    elif invalid_flag:
        try:
            indi_data_AC = ['']*3
            indi_data_EN[0] = lnum
            indi_data_EN[11] = 'Deleted'
            indi_data_AC[0] = lnum
            # indi_data_AR[0] = lnum
            
            with open(File_path_English_CSV, 'a', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(indi_data_EN)
            count()
            with open(File_path_English_txt, 'a', encoding="utf-8") as fw:
                fw.write("\t".join(map(str, indi_data_EN)) + "\n")
                fw.flush()
            # with open(File_path_Arabic_CSV, 'a', newline='', encoding='utf-8') as file:
            #     writer = csv.writer(file)
            #     writer.writerow(indi_data_AR)
            # with open(File_path_Arabic_txt, 'a', encoding="utf-8") as fw:
            #     fw.write("\t".join(map(str, indi_data_AR)) + "\n")
            #     fw.flush()
            with open(File_path_Activity_CSV, 'a', newline='', encoding='utf-8') as file:
                writer = csv.writer(file)
                writer.writerow(indi_data_AC)
            with open(File_path_Activity_txt, 'a', encoding="utf-8") as fw:
                fw.write("\t".join(map(str, indi_data_AC)) + "\n")
                fw.flush()
            
            with open(File_path_log_index, 'w', encoding='utf-8') as file:
                file.write(lnum)
                file.flush()
        except:
            exception()


def convertCSVExcel(File_path_CSV, File_path_EXL):
    df = pd.read_csv(File_path_CSV, encoding='utf-8', low_memory=False)
    df.to_excel(File_path_EXL, index=False)


def duplicate(File_path):
    try:
        data = pd.read_excel(File_path)
        data_file = data.drop_duplicates()
        data_file.to_excel(File_path, index=False)
    except:
        pass
    
    
def process_tasks(licence_num):
    try:
        
            try:
                st = time.time()
                invalid_flag = False
                user_agent = random.choice(user_agents)
                Headers = {'User-Agent': user_agent}
                req_url = Base_URL.format(licence_num)
                json_data = request(req_url, Headers)
                # success = individual_data(json_data)
                
                if "data" in json_data and json_data["data"]["status"] == "error":
                    invalid_flag = True
                    error_description = json_data["data"]["error"]["errorDescriptionEn"]
                    if "not exist" in error_description.lower():
                        individual_data(licence_num, json_data, invalid_flag)
                    else:
                        log_print(f'Error for {licence_num}: {error_description}')
                else:
                    individual_data(licence_num, json_data, invalid_flag)

                # invalid_flag = request(req_url)
                # if not invalid_flag:
                #     individual_data()
                # log_print(f'Complete {licence_num} in English')
                
                et = time.time()
                tm = et - st

                if os.path.exists(File_path_log_index):
                    with open(File_path_log_index, 'r', encoding='utf-8') as file:
                        last_licence = file.read().strip()
                # CN-1093928 CN-1053845
                if licence_num == last_licence:
                    if not invalid_flag:
                        log_print(f'Complete {licence_num} {tm: .2f}')
                    else:
                        log_print(f'Invalid {licence_num} {tm: .2f}')
                else:
                    log_print(f'Failed!! {licence_num} {tm: .2f}')
                    with open(File_path_failed_CSV, 'a', newline='', encoding='utf-8') as file:
                        writer = csv.writer(file)
                        writer.writerow([licence_num])
                    # df = pd.read_csv(File_path_failed_English_CSV, encoding='utf-8')
                    # df.to_excel(File_path_, index=False)

            except:
                exception(licence_num)
                
            time.sleep(1)

    finally:
        convertCSVExcel(File_path_English_CSV, File_path_English)
        convertCSVExcel(File_path_Activity_CSV, File_path_Activity)
        duplicate(File_path_English)
        duplicate(File_path_Activity)
        # convertCSVExcel(File_path_Arabic_CSV, File_path_Arabic)
        # duplicate(File_path_Arabic)
        if os.path.exists(File_path_failed_CSV):
            convertCSVExcel(File_path_failed_CSV, File_path_failed)
            
def process_data_sequentially(data):
    lock = threading.Lock()

    def process_wrapper():
        while True:
            with lock:
                if not data:
                    break
                item = data.pop(0)  
            process_tasks(item)

    num_threads = 5  # Number of threads
    threads = []

    # Create threads
    for _ in range(num_threads):
        thread = threading.Thread(target=process_wrapper)
        thread.start()
        threads.append(thread)

    # Wait for all threads to complete
    for thread in threads:
        thread.join()
        
    while data:
        threads = [thread for thread in threads if thread.is_alive()]

        for thread in threads:
            if not data:
                break
            item = data.pop(0)
            process_tasks(item)
        
        
if __name__ == '__main__':

    # Create directories if they don't exist
    directories = [
        BasePath + '\\OP',
        BasePath + '\\InputFile',
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
        File_paths_csv = [File_path_English_CSV, File_path_Activity_CSV]
        File_paths_txt = [File_path_English_txt, File_path_Activity_txt]
        File_paths_error = [File_path_error_CSV, File_path_failed_CSV]
        File_paths_log = [File_path_log, File_path_log_index]
        if os.path.exists(File_path_count):
            os.remove(File_path_count)
        for path_log in File_paths_log:
            if os.path.exists(path_log):
                os.remove(path_log)
        for path_csv in File_paths_csv:
            if os.path.exists(path_csv):
                os.remove(path_csv)
        for Path_txt in File_paths_txt:
            if os.path.exists(Path_txt):
                os.remove(Path_txt)
        for Path_ere in File_paths_error:
            if os.path.exists(Path_ere):
                os.remove(Path_ere)

    English_headers = ['Trade Licence Number', 'ADCCI Number', 'Trade Name in English', 'Trade Name in Arabic', 'Legal Form', 
                        'Licence Type', 'Branch', 'Issuance Place', 'Establishment Date', 'Registration Date', 'Expiry Date',
                        'Licence Status', 'Address', 'Establishment Volume', 'Social Media Account','Social Media Type', 'Web Site URL',]
    Activity_headers = ['Trade Licence Number','Trade Licence Activities', 'Trade Licence Activities - Code']

    # English_headers = ['Trade Licence Number', 'ADCCI Number', 'Trade Name', 'Legal Form', 'Licence Type',
    #                    'Branch', 'Issuance Place', 'Establishment Date', 'Registration Date', 'Expiry Date',
    #                    'Licence Status', 'Address', 'Establishment Volume', 'Social Media Account',
    #                    'Social Media Type', 'Web Site URL',]
    # Arabic_headers = ['Trade Licence Number', 'Trade Name']
    # Activity_headers = ['Trade Licence Number','Trade Licence Activities', 'Trade Licence Activities - Code']

    with open(File_path_count, "a") as f:
        f.write("")
    with open(File_path_English_txt, "a") as fw:
        if fw.tell() == 0:
            fw.write("\t".join(English_headers) + "\n")
            fw.flush()
    # with open(File_path_Arabic_txt, "a") as fw:
    #     if fw.tell() == 0:
    #         fw.write("\t".join(Arabic_headers) + "\n")
    #         fw.flush()
    with open(File_path_Activity_txt, "a") as fw:
        if fw.tell() == 0:
            fw.write("\t".join(Activity_headers) + "\n")
            fw.flush()

    with open(File_path_English_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(English_headers)
    # with open(File_path_Arabic_CSV, 'a', newline='', encoding='utf-8') as file:
    #     writer = csv.writer(file)
    #     if file.tell() == 0:
    #         writer.writerow(Arabic_headers)
    with open(File_path_Activity_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(Activity_headers)

    try:
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
        
        Base_URL = 'https://www.tamm.abudhabi/services/business/ded/pub/proxy/ms-call/gateway/TammJourneyDed/2.0/dedBusiness/businessLicense/tahaqaq?applicationNo={}&inquiryType=&'
        st = time.time()
        
        retry_attempts = 16
        retry_delay = 2

        log_print('Data Importing...plz wait')
        df = pd.read_excel(File_path_Input, sheet_name='Sheet1')
        licence_num_list_og = df['number'].tolist()
        # licence_num_list_og = ["CN-100002", "CN-1000104", "CN-1000141"]
        log_print('Data Imported\n')

        log_index_flag = False
        if os.path.exists(File_path_log_index):
            log_index_flag = True
            with open(File_path_log_index, 'r', encoding='utf-8') as file:
                last_processed_licence = file.read().strip()

        if log_index_flag:
            start_index = licence_num_list_og.index(last_processed_licence) + 1
            licence_num_list = licence_num_list_og[start_index:]
        else:
            licence_num_list = licence_num_list_og
            
        process_data_sequentially(licence_num_list)
        # process_tasks(licence_num_list)
        ########################################### License Number ###########################################
        
            # os.remove(File_path_log_index_English)
            # with open(File_path_log_index_English, 'w', encoding='utf-8') as f1:
            #     f1.write('')
            #     f1.flush()
            

    except:
        exception()
        
    et = time.time()
    log_print(f'\n{et - st}')
    os._exit(0)

    # finally:
        # convertCSVExcel(File_path_English_CSV, File_path_English)
        # convertCSVExcel(File_path_Activity_CSV, File_path_Activity)
        # convertCSVExcel(File_path_Arabic_CSV, File_path_Arabic)

database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
    delete_task(conn, File_path_English)
    # delete_task(conn, File_path_Arabic)
    delete_task(conn, File_path_Activity)
    # delete_task(conn, File_path_English_excel)
    # delete_task(conn, File_path_Arabic_excel)
    # delete_task(conn, File_path_Activity_excel)
    delete_task(conn, File_path_count)
    delete_task(conn, File_path_error)
