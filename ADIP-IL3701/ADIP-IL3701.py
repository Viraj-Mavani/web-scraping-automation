import csv
import json
import os
import random
import string
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
File_path_failed_CSV = BasePath + '\\OPcsv\\ADIP-IL3701_Failed.csv'
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


# def restart_script():
#     python = sys.executable
#     subprocess.call([python] + sys.argv)


def request(pl, headers):
    json_data = None
    try:
        Retry = 1
        while Retry <= retry_attempts:
            try:
                # url = Base_URL.format('CN-1000141')
                r_delay = random.uniform(0.5, 3.0)
                time.sleep(r_delay)
                obj = requests.post(Base_URL, json=pl, headers=headers, timeout=200)
                json_data = obj.json()
                return json_data
            except requests.JSONDecodeError:
                Retry += 1
                if Retry>2:
                    break
                else:
                    log_print(f"Retry {Retry-1}...for JSON")
                    continue
            except Exception:
                exception()
                log_print(f"Error occurred in Request")
                user_agent = random.choice(user_agents)
                headers = {
                    'User-Agent': user_agent,
                    'Content-Type': 'application/json'
                }
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

        
def individual_data(lnum, data_json):
    indi_data_search = ['']*6
    indi_data_corporate = ['']*9
    indi_data_address = ['']*8
    try:
        if data_json and data_json.get("Success") == True and len(data_json['Data']) > 0:
            main_data = data_json.get("Data", [])
            for companies in main_data:
                try:
                    for com in companies :
                        try:
                            if 'DisplayId' in com:
                                c_num = companies['DisplayId']
                                indi_data_search[0] = c_num
                                indi_data_corporate[8] = c_num
                                indi_data_address[7] = c_num
                            elif 'DisplayName' in com:
                                hebrew_name = companies['DisplayName']
                                indi_data_search[1] = hebrew_name
                                indi_data_corporate[0] = hebrew_name
                            elif 'DisplayCompanyType' in com:
                                indi_data_search[2] = companies['DisplayCompanyType']
                            elif 'StatusString' in com:
                                indi_data_search[3] = companies['StatusString']
                            elif 'DisplayCompanyViolates' in com:
                                indi_data_search[4] = companies['DisplayCompanyViolates']
                            elif 'LastYearlyReport' in com:
                                indi_data_search[5] = companies['LastYearlyReport']
                            
                            elif 'DisplayEnglishName' in com:
                                english_name = companies['DisplayEnglishName']
                                indi_data_corporate[1] = english_name  
                            elif 'DisplayCompanyRegistrationDate' in com:
                                indi_data_corporate[2] = companies['DisplayCompanyRegistrationDate']
                            elif 'DisplayCompanySubStatus' in com:
                                if companies['DisplayCompanySubStatus'] is not None:
                                    indi_data_corporate[3] = companies['DisplayCompanySubStatus']
                            elif 'PurposeDescription' in com:
                                if companies['PurposeDescription'] is not None:
                                    indi_data_corporate[4] = companies['PurposeDescription']
                            elif 'DisplayCompanyPurpose' in com:
                                indi_data_corporate[5] = companies['DisplayCompanyPurpose']   
                            elif 'IsGovernmental' in com:
                                IsGovernmental = companies['IsGovernmental']
                                if IsGovernmental == True:
                                    indi_data_corporate[6] = "Yes"
                                elif IsGovernmental == False:
                                    indi_data_corporate[6] = "No"
                            elif 'DisplayCompanyLimitType' in com:
                                indi_data_corporate[7] = companies['DisplayCompanyLimitType']
                                
                            elif ('Address' in com) and companies['Address'] and len(companies['Address']) > 0:
                                Addresses = companies['Address']
                                for Address in Addresses:
                                    try:
                                        if 'CityName' in Address:
                                            indi_data_address[5] = Addresses['CityName']
                                        elif 'StreetName' in Address:
                                            indi_data_address[3] = Addresses['StreetName']
                                        elif 'HouseNumber' in Address:
                                            indi_data_address[4] = Addresses['HouseNumber']
                                        elif 'ZipCode' in Address:
                                            indi_data_address[2] = Addresses['ZipCode']
                                        elif 'PostBox' in Address:
                                            indi_data_address[1] = Addresses['PostBox']
                                        elif 'CountryName' in Address:
                                            indi_data_address[0] = Addresses['CountryName']
                                        elif 'AtAddress' in Address:
                                            indi_data_address[6] = Addresses['AtAddress']
                                    except:
                                        exception()
        
                        except:
                            exception()
                        
                    try:
                        # Write to CSV file
                        with open(File_path_search_page_CSV, 'a', newline='', encoding='utf-8') as file:
                            writer = csv.writer(file)
                            writer.writerow(indi_data_search)
                        # Write to TXT file
                        with open(File_path_search_page_TXT, 'a', encoding="utf-8") as fw:
                            fw.write("\t".join(map(str, indi_data_search)) + "\n")
                            fw.flush() 
                            
                        count()

                        # Write to CSV file
                        with open(File_path_company_details_CSV, 'a', newline='', encoding='utf-8') as file:
                            writer = csv.writer(file)
                            writer.writerow(indi_data_corporate)
                        # Write to TXT file
                        with open(File_path_company_details_TXT, 'a', encoding="utf-8") as fw:
                            fw.write("\t".join(map(str, indi_data_corporate)) + "\n")
                            fw.flush() 

                        # Write to CSV file
                        with open(File_path_address_CSV, 'a', newline='', encoding='utf-8') as file:
                            writer = csv.writer(file)
                            writer.writerow(indi_data_address)
                        # Write to TXT file
                        with open(File_path_address_TXT, 'a', encoding="utf-8") as fw:
                            fw.write("\t".join(map(str, indi_data_address)) + "\n")
                            fw.flush() 
                        
                        log_print(f"Added {english_name}")
                    except:
                        exception()
                except:
                    exception()
            # Write to INDEX file
            with open(File_path_log_index, 'w', encoding='utf-8') as file:
                file.write(lnum)
                file.flush()
        else:
            log_print(f"No data Found for {lnum}")
            with open(File_path_log_index, 'w', encoding='utf-8') as file:
                file.write(lnum)
                file.flush()
    except:
        exception()


if __name__ == '__main__':
    
    File_paths = [File_path_search_page_CSV, File_path_search_page_TXT, File_path_company_details_CSV, File_path_company_details_TXT, 
                  File_path_address_TXT, File_path_address_CSV, File_path_error_CSV]
    file_paths_logs = [File_path_log, File_path_log_index, File_path_failed_CSV]
    
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
    HeadersF2 = ["Name in Hebrew", "English Name", "Date of incorporation", "Sub-status", "Company description", 
                 "The purpose of the company", "Govermental company", "Limitation", "Corporation number"]
    HeadersF3 = ["Country", "P.O.B", "Zip code", "Street", "Number", "Settlement", "In", "Corporation number"]
    
    with open(File_path_search_page_TXT, "a") as fw:
        if fw.tell() == 0:
            fw.write("\t".join(HeadersF1) + "\n")
            fw.flush()
    with open(File_path_search_page_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(HeadersF1)
    
    with open(File_path_company_details_TXT, "a") as fw:
        if fw.tell() == 0:
            fw.write("\t".join(HeadersF2) + "\n")
            fw.flush()
    with open(File_path_company_details_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(HeadersF2)
            
    with open(File_path_address_TXT, "a") as fw:
        if fw.tell() == 0:
            fw.write("\t".join(HeadersF3) + "\n")
            fw.flush()
    with open(File_path_address_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(HeadersF3)

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
        headers = {
            'User-Agent': user_agent,
            'Content-Type': 'application/json'
        }

        # Base_URL = 'https://ica.justice.gov.il/GenericCorporarionInfo/SearchCorporation?unit=8'
        Base_URL = 'https://ica.justice.gov.il/GenericCorporarionInfo/SearchGenericCorporation'
        
        st = time.time()
                
        retry_attempts = 15
        retry_delay = 2

        t_letter_combinations = [f'{a}{b}{c}' for a in string.ascii_lowercase for b in string.ascii_lowercase for c in string.ascii_lowercase]

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
            try:
                success_flag = True
                user_agent = random.choice(user_agents)
                headers = {
                    'User-Agent': user_agent,
                    'Content-Type': 'application/json'
                }
                # req_url = Base_URL.format(letter)
                pay_load = {
                    "corporationType":3,
                    "CorporationName":letter
                }
                json_data = request(pay_load, headers)
                
                if json_data is not None:
                    individual_data(letter, json_data)
                elif json_data is None:
                    log_print(f"No JSON Found for {letter}")
                    success_flag = False
                    with open(File_path_log_index, 'w', encoding='utf-8') as file:
                        file.write(letter)
                        file.flush()

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
                continue
    except:
        exception()
        
    finally:
        duplicateFromCSV(File_path_search_page_CSV)
        duplicateFromCSV(File_path_company_details_CSV)
        duplicateFromCSV(File_path_address_CSV)
        convertCSVExcelExtended(File_path_search_page_CSV, File_path_search_page)
        convertCSVExcelExtended(File_path_company_details_CSV, File_path_company_details)
        convertCSVExcelExtended(File_path_address_CSV, File_path_address)
    
    et = time.time()
    log_print(f'\n{et - st}')
    if os.path.exists(File_path_log_index):
        with open(File_path_log_index, 'r', encoding='utf-8') as file:
            last_letter = file.read().strip()
    if last_letter == letter_num_list[-1]:
        log_print("Success")
        if os.path.exists(File_path_log_Run_Flag):
            os.remove(File_path_log_Run_Flag)
        if os.path.exists(File_path_count):
            os.remove(File_path_count)
        
    else:
        log_print(f"Stopped at {last_letter}")

# database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# # create a database connection
# conn = create_connection(database)
# with conn:
#     for File_path in File_paths:
#         delete_task(conn, File_path)
#     for File_path in file_paths_logs:
#         delete_task(conn, file_paths_logs)
# 	# delete_task(conn, File_paths)
# 	# delete_task(conn, file_paths_logs)
