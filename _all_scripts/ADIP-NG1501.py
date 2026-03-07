import csv
import json
import os
import random
import sys
import re
import sqlite3
from sqlite3 import Error
import time
import traceback
import pandas as pd
import requests
import string
# import xlsxwriter
# from openpyxl.styles import Font

# BasePath = 'E:\\ADIP-PY\\'
# BasePath = 'D:\\Projects\\CedarPython\\ADIP-NG1501'
BasePath = os.getcwd()
Total_URL = 0
######### Excel #########
File_path_Search_Page_Info = BasePath + '\\OP\\ADIP-NG1501_Search_Page_Info.xlsx'
File_path_Shareholders = BasePath + '\\OP\\ADIP-NG1501_Shareholders.xlsx'
######### Text #########
File_path_Search_Page_Info_txt = BasePath + '\\OPtxt\\ADIP-NG1501_Search_Page_Info.txt'
File_path_Shareholders_txt = BasePath + '\\OPtxt\\ADIP-NG1501_Shareholders.txt'
######### CSV #########
Error_File_CSV = BasePath + '\\Error\\ADIP-NG1501_Error.csv'
File_path_Search_Page_Info_CSV = BasePath + '\\OPcsv\\ADIP-NG1501_Search_Page_Info.csv'
File_path_Shareholders_CSV = BasePath + '\\OPcsv\\ADIP-NG1501_Shareholders.csv'
######### Count #########
File_path_search_count = BasePath + '\\Counts\\ADIP-NG1501_Count.txt'
######### Error #########
Error_File = BasePath + '\\Error\\ADIP-NG1501_Error.xlsx'
######### Input #########
# File_path_Input = BasePath + 'Proxy\\http_proxies.xlsx'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-NG1501_Log.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-NG1501_Run_Flag.txt'
File_path_log_index_LetterE1 = BasePath + '\\Log\\ADIP-NG1501_Log_Index_LetterE1.txt'
File_path_log_index_LetterE2 = BasePath + '\\Log\\ADIP-NG1501_Log_Index_LetterE2.txt'


English_alphabet_list = list(string.ascii_lowercase)


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


def log_print(message):
    with open(File_path_log, 'a', encoding='utf-8') as file:
        file.write(message + '\n')
        file.flush()
    print(message)


def exception():
    Headers_Error = ['URL', 'Not Responding', 'Error']
    error = traceback.format_exc()
    exception_type, exception_object, exception_traceback = sys.exc_info()
    with open(Error_File_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(Headers_Error)
        writer.writerow([Search_Page_URL, "Not Responding", str(error)])
    df = pd.read_csv(Error_File_CSV, encoding='utf-8')
    df.to_excel(Error_File, index=False)


def convertCSVExcel(File_path_CSV, File_path_EXL):
    df = pd.read_csv(File_path_CSV, encoding='utf-8')
    df.to_excel(File_path_EXL, index=False)


def duplicate(File_path):
    try:
        data = pd.read_excel(File_path)
        data_file = data.drop_duplicates()
        data_file.to_excel(File_path, index=False)
    except:
        pass


def Dereference(obj):
    del obj


if __name__ == '__main__':

    ############################################# Writing Headers for Excel Files #############################################
    File_paths = [File_path_Search_Page_Info, File_path_Shareholders]

    # First_run = False
    # if First_run:
    if not os.path.exists(File_path_log_Run_Flag):
        with open(File_path_log_Run_Flag, "a", encoding='utf-8')as f:
            f.write("")
        File_paths_csv = [File_path_Search_Page_Info_CSV, File_path_Shareholders_CSV]
        File_paths_txt = [File_path_Search_Page_Info_txt, File_path_Shareholders_txt]
        File_path_index = [File_path_log_index_LetterE1, File_path_log_index_LetterE2]
        if os.path.exists(File_path_log):
            os.remove(File_path_log)
        if os.path.exists(File_path_search_count):
            os.remove(File_path_search_count)
        for path_csv in File_paths_csv:
            if os.path.exists(path_csv):
                os.remove(path_csv)
        for Path_txt in File_paths_txt:
            if os.path.exists(Path_txt):
                os.remove(Path_txt)
        for Path_index in File_path_index:
            if os.path.exists(Path_index):
                os.remove(Path_index)

    # Create directories if they don't exist
    directories = [
        BasePath + '\\OP',
        BasePath + '\\OPtxt',
        BasePath + '\\OPcsv',
        BasePath + '\\Proxy',
        BasePath + '\\Error',
        BasePath + '\\Counts',
        BasePath + '\\Log']

    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)

    Search_Page_Headers = ['RC Number', 'Company Name',
                           'Address', 'Status', 'Date of Registration']
    Shareholders_Headers = ['Shareholder Name', 'Address', 'Date of PSC', 'Legal form', 'governing Law', 'Register', 'Country of registration', 'Email',
                            'Does the PSC directly or indirectly hold at least 5% of the shares or interest in a company or limited liability partnership?',
                            'Does the PSC directly or indirectly hold at least 5% of the voting rights in a company or limited liability partnership?',
                            'Does the PSC directly or indirectly hold the right to appoint or remove a majority of the directors or partners in a company or limited liability partnership?',
                            'Does the PSC otherwise have the right to exercise or is actually exercising significant influence or control over a company or limited liability partnership?',
                            'Does the PSC have the right to exercise, or actually exercise significant influence or control over the activities of a trust or firm, whether or not it is a legal entity, but would itself satisfy any of the first four conditions if it were an individual?',
                            'Company RC Number', 'Main Company Name']

    with open(File_path_search_count, "a", encoding='utf-8')as f:
        f.write("")
    with open(File_path_Search_Page_Info_txt, "a", encoding='utf-8')as f:
        if f.tell() == 0:
            f.write("\t".join(Search_Page_Headers)+"\n")
            f.flush()
    with open(File_path_Shareholders_txt, "a", encoding='utf-8')as f:
        if f.tell() == 0:
            f.write("\t".join(Shareholders_Headers)+"\n")
            f.flush()

    with open(File_path_Search_Page_Info_CSV, "a", newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if f.tell() == 0:
            writer.writerow(Search_Page_Headers)
    with open(File_path_Shareholders_CSV, "a", newline='', encoding='utf-8') as fw:
        writer = csv.writer(fw)
        if fw.tell() == 0:
            writer.writerow(Shareholders_Headers)

    Search_Page_URL = 'https://searchapp.cac.gov.ng/searchapp/api/public-search/company-business-name-it'

    retry_attempts = 5
    retry_delay = 2

    # # Proxy Credentials
    # username = 'CedarRose-res-NG'
    # password = 'DZjZc7NJi9F8pi0'
    # server = 'gw.ntnt.io'
    # port = '5959'
    # proxy = {
    # 'http': f'http://{username}:{password}@{server}:{port}',
    # 'https': f'http://{username}:{password}@{server}:{port}'
    # }
    # proxies = [
    #     '95.216.189.78:8080',
    #     '65.21.0.216:8080',
    #     '65.109.236.232:8080',
    #     '120.197.219.82:9091',
    #     '5.161.78.209:8080']
    # df = pd.read_excel(File_path_Input, sheet_name='Sheet1')
    # proxies = df['proxies'].tolist()

    log_index_flag_LetterE1 = False
    log_index_flag_LetterE2 = False
    if os.path.exists(File_path_log_index_LetterE1):
        with open(File_path_log_index_LetterE1, 'r', encoding='utf-8') as file:
            index_LetterE1 = file.read().strip()
            if index_LetterE1 != '':
                log_index_flag_LetterE1 = True

    if os.path.exists(File_path_log_index_LetterE2):
        with open(File_path_log_index_LetterE2, 'r', encoding='utf-8') as file:
            index_LetterE2 = file.read().strip()
            if index_LetterE2 != '':
                log_index_flag_LetterE2 = True

    if log_index_flag_LetterE1:
        start_index_LetterE1 = English_alphabet_list.index(index_LetterE1) + 1
    else:
        start_index_LetterE1 = 0

    if log_index_flag_LetterE2:
        start_index_LetterE2 = English_alphabet_list.index(index_LetterE2) + 1
    else:
        start_index_LetterE2 = 0

    try:
        for indexE1 in range(start_index_LetterE1, len(English_alphabet_list)):
            letterE1 = English_alphabet_list[indexE1]
            for indexE2 in range(start_index_LetterE2, len(English_alphabet_list)):
                letterE2 = English_alphabet_list[indexE2]
                letter = letterE1+letterE2
                log_print('\n\nScraping for Letter : ' + letter)
                payload = json.dumps({"searchTerm": letter})
                headers = {"content-type": "application/json"}
                # proxy = random.choice(proxies)
                # proxy_url = f'http://{proxy}'
                # log_print(f"Using proxy: {proxy}")

                try:
                    searchpageRetry = 1
                    while searchpageRetry <= retry_attempts:
                        try:
                            # Search_Page = requests.post(Search_Page_URL, data=payload, headers=headers, proxies=proxy)
                            Search_Page = requests.post(Search_Page_URL, data=payload, headers=headers, timeout=100)
                            break
                        except Exception as e:
                            if isinstance(e, ConnectionError):
                                log_print(f"ConnectionError occurred in Search_Page for {letter}")
                            else:
                                log_print(f"Error occurred in Search_Page for {letter}")
                            delay = retry_delay * \
                                (2 ** searchpageRetry)
                            log_print(f'Retrying in {delay} seconds...{searchpageRetry}')
                            time.sleep(delay)
                            searchpageRetry += 1
                            continue
                    else:
                        exception()
                        # raise e
                        os._exit(1)
                    if re.search(r'"data":\[\{"state"', Search_Page.text):
                        res_data = json.loads(Search_Page.text)
                        search_data = res_data['data']
                    else:
                        search_data = None
                    if search_data == None:
                        continue
                    Status_ids = []
                    for ids in search_data:
                        Status_ids.append(ids['id'])
                    Status_url = 'https://searchapp.cac.gov.ng/searchapp/api/public-search/check-company-status'
                    Status_payload = {"companyIds": Status_ids}
                    Status_headers = {"content-type": "application/json",
                                      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36"}

                    statusRetry = 1
                    while statusRetry <= retry_attempts:
                        try:
                            Status = requests.post(Status_url, data=json.dumps(Status_payload), headers=Status_headers, timeout=100)
                            break
                        except Exception as e:
                            if isinstance(e, ConnectionError):
                                log_print(f"ConnectionError occurred in status for {letter}")
                            else:
                                log_print(f"Error occurred in status for {letter}")
                            delay = retry_delay * \
                                (2 ** statusRetry)
                            log_print(f'Retrying in {delay} seconds...{statusRetry}')
                            time.sleep(delay)
                            statusRetry += 1
                            continue
                    else:
                        exception()
                        # raise e
                        exit()

                    if re.search(r'"data":\{"\d+"', Status.text):
                        status_data = json.loads(Status.text)
                        status_info = status_data['data']
                    else:
                        status_info = None
                    if status_info == None:
                        continue
                    for i in search_data:
                        if i['classificationId'] == 2:
                            ID = i['id']
                            Search_data = ['']*5
                            log_print('Adding ' + i['approvedName'].strip())
                            with open(File_path_search_count, "a", encoding='utf-8')as fh:
                                fh.write("1\n")

                            Search_data[0] = i['rcNumber'] if i['rcNumber'] != None else 'NOT YET ASSIGNED'
                            RC = i['rcNumber'] if i['rcNumber'] != None else 'NOT YET ASSIGNED'

                            company_name = i['approvedName'].strip() if i['approvedName'] != None else ''
                            Search_data[1] = i['approvedName'].strip() if i['approvedName'] != None else ''

                            Search_data[2] = ' '.join(i['address'].split()) if i['address'] != None else ''

                            Search_data[3] = i['companyStatus'] if i['companyStatus'] != None else 'INACTIVE'

                            Search_data[4] = i['registrationDate'].split('T')[0] if i['registrationDate'] != None else 'UNDER REGISTRATION'

                            with open(File_path_Search_Page_Info_CSV, 'a', newline='', encoding='utf-8') as file:
                                writer = csv.writer(file)
                                writer.writerow(Search_data)
                            with open(File_path_Search_Page_Info_txt, "a", encoding='utf-8') as fw:
                                fw.write("\t".join(map(str, Search_data))+"\n")

                            try:
                                if ID:
                                    Shareholders_URL = 'https://searchapp.cac.gov.ng/searchapp/api/status-report/find/company-affiliates/{id}'.format(id=ID)
                                    Share_headers = {
                                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.0.0 Safari/537.36",
                                        "Accept-Encoding": "gzip, deflate, br",
                                        "Connection": "keep-alive",
                                        "content-type": "application/json",
                                        "accept": "*/*"
                                    }
                                    shareholdersRetry = 1
                                    while shareholdersRetry <= retry_attempts:
                                        try:
                                            Shareholders_Info = requests.get(Shareholders_URL, headers=Share_headers, timeout=100)
                                            break
                                        except Exception as e:
                                            if isinstance(e, ConnectionError):
                                                log_print(f"ConnectionError occurred in Shareholders_Info for {letter}")
                                            else:
                                                log_print(f"Error occurred in Shareholders_Info for {letter}")
                                            delay = retry_delay * (2 ** shareholdersRetry)
                                            log_print(f'Retrying in {delay} seconds...{shareholdersRetry}')
                                            time.sleep(delay)
                                            shareholdersRetry += 1
                                            continue
                                    else:
                                        exception()
                                        # raise e
                                        exit()
                                    # try_count = 1
                                    # while True:
                                    #     try:
                                    #         Shareholders_Info = requests.get(Shareholders_URL, headers=Share_headers, timeout=100)
                                    #         break
                                        # except:
                                        #     if try_count > 3:
                                        #         break
                                        #     try_count += 1
                                    if Shareholders_Info.status_code == 200:
                                        res = json.loads(Shareholders_Info.text)
                                        data = res['data']
                                        Shareholders_Data = ['']*15
                                        for item in data:
                                            if item['affiliatesPscInformation']:
                                                Shareholders_Data[13] = RC
                                                Shareholders_Data[14] = company_name
                                                if item['isCorporate'] == None or item['isCorporate'] == False:
                                                    Name = item['surname'].strip() + ' ' if item['surname'] else ''
                                                    Name += item['firstname'].strip() + ' ' if item['firstname'] else ''
                                                    Name += item['otherName'].strip() + ' ' if item['otherName'] else ''
                                                elif item['isCorporate'] == True:
                                                    Name = item['corporationName'].strip() + ' ' if item['corporationName'] else ''
                                                    Name += item['rcNumber'].strip() + ' ' if item['rcNumber'] else ''

                                                Shareholders_Data[0] = Name

                                                Address = item['streetNumber'] + ', ' if item['streetNumber'] else ''
                                                Address += item['address'] + ', ' if item['address'] else ''
                                                Address += item['city'] + ', ' if item['city'] else ''
                                                if item['state'] == "FCT":
                                                    Address += item['state']
                                                elif item['state']:
                                                    Address += item['state'] + ' STATE, '
                                                else:
                                                    Address += ''
                                                Shareholders_Data[1] = Address

                                                Shareholders_Data[2] = item['affiliatesPscInformation']['dateOfPsc'] if item['affiliatesPscInformation']['dateOfPsc'] else ''

                                                Shareholders_Data[3] = item['affiliatesPscInformation']['legalForm'] if item['affiliatesPscInformation']['legalForm'] else ''

                                                Shareholders_Data[4] = item['affiliatesPscInformation']['governingLaw'] if item['affiliatesPscInformation']['governingLaw'] else ''

                                                Shareholders_Data[5] = item['affiliatesPscInformation']['register'] if item['affiliatesPscInformation']['register'] else ''

                                                Shareholders_Data[6] = item['affiliatesPscInformation']['taxResidencyOrJurisdiction'] if item['affiliatesPscInformation']['taxResidencyOrJurisdiction'] else ''

                                                Shareholders_Data[7] = item['email'] if item['email'] else ''

                                                if item['affiliatesPscInformation']['pscHoldsSharesOrInterest']:
                                                    Shareholders_Data[8] = 'YES Directly: [{d}%] and Indirectly: [{ind}%]'.format(
                                                        d=item['affiliatesPscInformation']['pscHoldsSharesOrInterestPercentageHeldDirectly'], ind=item['affiliatesPscInformation']['pscHoldsSharesOrInterestPercentageHeldIndirectly'])
                                                else:
                                                    Shareholders_Data[8] = 'NO Directly: [0%] and Indirectly: [0%]'
                                                if item['affiliatesPscInformation']['pscVotingRights']:
                                                    Shareholders_Data[9] = 'YES Directly: [{d}%] and Indirectly: [{ind}%]'.format(
                                                        d=item['affiliatesPscInformation']['pscVotingRightsPercentageHeldDirectly'], ind=item['affiliatesPscInformation']['pscVotingRightsPercentageHeldIndirectly'])
                                                else:
                                                    Shareholders_Data[9] = 'NO Directly: [0%] and Indirectly: [0%]'
                                                if item['affiliatesPscInformation']['pscRightToAppoints']:
                                                    Shareholders_Data[10] = 'YES'
                                                else:
                                                    Shareholders_Data[10] = 'NO'
                                                if item['affiliatesPscInformation']['pscSignificantInfluence']:
                                                    Shareholders_Data[11] = 'YES'
                                                else:
                                                    Shareholders_Data[11] = 'NO'
                                                if item['affiliatesPscInformation']['pscExeriseSignificantInfluence']:
                                                    Shareholders_Data[12] = 'YES'
                                                else:
                                                    Shareholders_Data[12] = 'NO'

                                                with open(File_path_Shareholders_CSV, 'a', newline='', encoding='utf-8') as file:
                                                    writer = csv.writer(file)
                                                    writer.writerow(
                                                        Shareholders_Data)
                                                with open(File_path_Shareholders_txt, "a", encoding='utf-8') as fw:
                                                    fw.write(
                                                        "\t".join(map(str, Shareholders_Data))+"\n")
                            except Exception as e:
                                exception()
                    with open(File_path_log_index_LetterE2, 'w', encoding='utf-8') as file:
                        file.write(letterE2)
                        file.flush()
                except Exception as e:
                    exception()
            with open(File_path_log_index_LetterE1, 'w', encoding='utf-8') as f1:
                f1.write(letterE1)
                f1.flush()
            convertCSVExcel(File_path_Search_Page_Info_CSV, File_path_Search_Page_Info)
            convertCSVExcel(File_path_Shareholders_CSV, File_path_Shareholders)
            duplicate(File_path_Search_Page_Info)
            duplicate(File_path_Shareholders)
            start_index_LetterE2 = 0
        # with open([File_path_log_index_LetterE1, File_path_log_index_LetterE2], 'w', encoding='utf-8') as f1:
        #     f1.write('')
        #     f1.flush()
        log_print('Script Complete')
    finally:
        file_paths = [File_path_log_index_LetterE1, File_path_log_index_LetterE2]
        if all(os.path.exists(file_path) for file_path in file_paths):
            letters = []
            for file_path in file_paths:
                with open(file_path, 'r', encoding='utf-8') as file:
                    letter = file.read().strip()
                    letters.append(letter)
            if all(letter == 'z' for letter in letters):
                log_print('English Letters Complete')
        exit()
database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
    for File_path in File_paths:
        delete_task(conn, File_path)
