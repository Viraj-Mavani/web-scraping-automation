import csv
import json
import math
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
BasePath = 'D:\\Projects\\CedarPython\\ADIP-AE415_MT'
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
File_path_log_index_Thread1 = BasePath + '\\Log\\ADIP-AE415_Log_Index_Thread1.txt'
File_path_log_index_Thread2 = BasePath + '\\Log\\ADIP-AE415_Log_Index_Thread2.txt'
File_path_log_index_Thread3 = BasePath + '\\Log\\ADIP-AE415_Log_Index_Thread3.txt'
File_path_log_index_Thread4 = BasePath + '\\Log\\ADIP-AE415_Log_Index_Thread4.txt'
File_path_log_index_Thread5 = BasePath + '\\Log\\ADIP-AE415_Log_Index_Thread5.txt'


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


def exception(licence_num=""):
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


def individual_data(lnum, data_json, invalid_flag, index_log):
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
                        
                with open(index_log, 'w', encoding='utf-8') as file:
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
            
            with open(index_log, 'w', encoding='utf-8') as file:
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
    
    
def process_tasks(num_list_og):
    current_thread = threading.current_thread().name
    index_set = {
        "T1": File_path_log_index_Thread1,
        "T2": File_path_log_index_Thread2,
        "T3": File_path_log_index_Thread3,
        "T4": File_path_log_index_Thread4,
        "T5": File_path_log_index_Thread5
    }
    if current_thread in index_set:
        index_path = index_set[current_thread]

    
    if index_path and os.path.exists(index_path):
        with open(index_path, 'r', encoding='utf-8') as file:
            last_processed_licence = file.read().strip()
            if last_processed_licence in num_list_og:
                start_index = num_list_og.index(last_processed_licence) + 1
                num_list = num_list_og[start_index:]
            else:
                num_list = "" 
    else:
        num_list = num_list_og
        
    try:
        for licence_num in num_list[:]:
            try:
                log_print(f"Processing Record {licence_num} by Thread {current_thread}")
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
                        individual_data(licence_num, json_data, invalid_flag, index_path)
                    else:
                        log_print(f'Error for {licence_num}: {error_description}')
                        with open(index_path, 'w', encoding='utf-8') as file:
                            file.write(licence_num)
                            file.flush()
                else:
                    individual_data(licence_num, json_data, invalid_flag, index_path)

                # invalid_flag = request(req_url)
                # if not invalid_flag:
                #     individual_data()
                # log_print(f'Complete {licence_num} in English')
                
                et = time.time()
                tm = et - st

                if os.path.exists(index_path):
                    with open(index_path, 'r', encoding='utf-8') as file:
                        last_licence = file.read().strip()
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
            
                with open(index_path, 'w', encoding='utf-8') as file:
                    file.write(str(licence_num))
                    file.flush()
                    
            except:
                exception(licence_num)
                continue
    except:
        exception()
        
    finally:
        if os.path.exists(index_path):
            with open(index_path, 'r', encoding='utf-8') as file:
                Thread_last_licence = file.read().strip()
        if Thread_last_licence == num_list_og[-1]:
            log_print(f'==============================================\n{current_thread} Completed\n==============================================')
            completed_threads.append(current_thread)
        else:
            log_print(f'==============================================\nStopped from {Thread_last_licence} in {current_thread}\n==============================================')
            
            
# def divide_data_into_chunks(data, chunk_size):
#     chunks = [data[i:i + chunk_size] for i in range(0, len(data), chunk_size)]
#     threads = []

#     for chunk in chunks:
#         thread = threading.Thread(target=process_tasks, args=(chunk,))
#         threads.append(thread)
#         thread.start()

#     # Wait for all threads to complete
#     for thread in threads:
#         thread.join()


def parallel_processing(records, num_threads):
    # Calculate the number of records per thread
    records_per_thread = math.ceil(len(records) / num_threads)

    # Create threads and start processing
    thread_strs = ["T1", "T2", "T3", "T4", "T5"]
    threads = []
    for i,thread_str in enumerate(thread_strs):
        start = i * records_per_thread
        end = (i + 1) * records_per_thread if i < num_threads - 1 else len(records)
        chunk = records[start:end]
        thread = threading.Thread(target=process_tasks, args=(chunk,))
        thread.name = thread_str
        threads.append(thread)
        thread.start()

    # Wait for all threads to complete
    for thread in threads:
        thread.join()

        
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

    # First_run = True
    # if First_run:
    if not os.path.exists(File_path_log_Run_Flag):
        with open(File_path_log_Run_Flag, "a", encoding='utf-8')as f:
            f.write("")
        File_paths_csv = [File_path_English_CSV, File_path_Activity_CSV]
        File_paths_txt = [File_path_English_txt, File_path_Activity_txt]
        File_paths_error = [File_path_error_CSV, File_path_failed_CSV]
        File_paths_log = [File_path_log, File_path_log_index_Thread1, File_path_log_index_Thread2, File_path_log_index_Thread3, File_path_log_index_Thread4, File_path_log_index_Thread5]
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
        
        numberOfThreads = 5
        completed_threads = []
        retry_attempts = 16
        retry_delay = 2

        log_print('Data Importing...plz wait')
        df = pd.read_excel(File_path_Input, sheet_name='Sheet1')
        licence_num_list_og = df['number'].tolist()
        licence_num_list_og = licence_num_list_og[:51]
        # licence_num_list_og = ["CN-100002", "CN-1000104", "CN-1000141"]
        log_print('Data Imported\n')

        parallel_processing(licence_num_list_og, numberOfThreads)
        # divide_data_into_chunks(licence_num_list, 41000)
        # process_tasks(licence_num_list)

        convertCSVExcel(File_path_English_CSV, File_path_English)
        convertCSVExcel(File_path_Activity_CSV, File_path_Activity)
        duplicate(File_path_English)
        duplicate(File_path_Activity)
        if os.path.exists(File_path_failed_CSV):
            convertCSVExcel(File_path_failed_CSV, File_path_failed)
            
        et = time.time()
        log_print(f'\n{et - st}')

    except:
        exception()

    finally:
        if len(completed_threads) == numberOfThreads:
            log_print('Script Completed')
        else:
            log_print('Script was Stopped')

database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# # create a database connection
# conn = create_connection(database)
# with conn:
#     delete_task(conn, File_path_English)
#     # delete_task(conn, File_path_Arabic)
#     delete_task(conn, File_path_Activity)
#     # delete_task(conn, File_path_English_excel)
#     # delete_task(conn, File_path_Arabic_excel)
#     # delete_task(conn, File_path_Activity_excel)
#     delete_task(conn, File_path_count)
#     delete_task(conn, File_path_error)
