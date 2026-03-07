import csv
import random
import subprocess
import sys
import traceback
import pandas as pd
import os
import sqlite3
from sqlite3 import Error
from bs4 import BeautifulSoup
import time
import requests

# BasePath = 'D:\\Projects\\CedarPython\\ADIP-TA201' 
BasePath = os.getcwd()
# BasePath= 'E:\\ADIP-PY\\OP2'


######### Excel #########
File_path_search_page = BasePath + '\\OP\\ADIP-TA201_search_page.xlsx'
File_path_details_page = BasePath + '\\OP\\ADIP-TA201_details_page.xlsx'
######### CSV #########
File_path_search_page_CSV = BasePath + '\\OPcsv\\ADIP-TA201_search_page.csv'
File_path_details_page_CSV = BasePath + '\\OPcsv\\ADIP-TA201_details_page.csv'
File_path_error_CSV = BasePath + '\\OPcsv\\ADIP-TA201_Error.csv'
######### Text #########
File_path_search_page_TXT = BasePath + '\\OPtxt\\ADIP-TA201_search_page.txt'
File_path_details_page_TXT = BasePath + '\\OPtxt\\ADIP-TA201_details_page.txt'
######### Error #########
File_path_error = BasePath + '\\Error\\ADIP-TA201_Error.xlsx'
######### Count #########
File_path_count = BasePath + '\\Counts\\ADIP-TA201_Count.txt'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-TA201_Log.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-TA201_Run_Flag.txt'
File_path_log_index_page = BasePath + '\\Log\\ADIP-TA201_Log_Page.txt'
File_path_log_index_company = BasePath + '\\Log\\ADIP-TA201_Log_Company.txt'


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


def exception(url):
    Headers_Error = ['URL', 'Not Responding', 'Error']
    error = traceback.format_exc()
    exception_type, exception_object, exception_traceback = sys.exc_info()
    with open(File_path_error_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(Headers_Error)
        writer.writerow([url, "Not Responding", error])
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


def convertCSVExcelExtended(File_path_CSV, File_path_EXL):
    chunk_size = 1000000 

    try:
        df = pd.read_csv(File_path_CSV, encoding='utf-8')
        df.to_excel(File_path_EXL, index=False)
        return None
    except (pd.errors.ParserError, pd.errors.EmptyDataError, ValueError):
        pass 
    except:
        exception(Base_URL)

    csv_reader = pd.read_csv(File_path_CSV, encoding='utf-8', chunksize=chunk_size)
    sheet_index = 1 
    # excel_files = []

    for chunk in csv_reader:
        if len(chunk) > 0:
            sheet_name = f'DataSet {sheet_index}'
            excel_file = f'{File_path_EXL[:-5]}_{sheet_index}.xlsx'
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


def request(req):
    user_agent = random.choice(user_agents)
    headers = {'User-Agent': user_agent}
    try:
        Retry = 1
        while Retry <= retry_attempts:
            try:
                r_delay = random.uniform(1.5, 3.0)
                time.sleep(r_delay)
                obj = requests.get(req, timeout=200)
                # obj = requests.post(req, headers, timeout=200)
                soup = BeautifulSoup(obj.content, 'html.parser')
                return soup
            except Exception:
                exception(req)
                log_print(f"Error occurred in Request")
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
        exception(req)
        

def search_data(search_sp):
    try:
        search_table = search_sp.find('table', class_='table table-hover table-sm table-bordered table-striped').find('tbody')
        rows = search_table.find_all('tr')[1:]
        
        totRows = len(rows)
        if totRows==0:
            log_print("No Data Found!!")
            # with open(File_path_log_index1, 'w', encoding='utf-8') as file:
            #     file.write(page)
            #     file.flush()
            return None
        
        indi_data = []
        
        for row in rows:
            try:
                cells = row.find_all('td')[1:-1]
                if len(cells) == 9:
                    try:
                        company_name = cells[0].text.strip()
                    except:
                        company_name = ''
                    try:
                        r_number = cells[1].text.strip()
                    except:
                        r_number = ''
                    try:
                        director = cells[2].text.strip()
                    except:
                        director = ''
                    try:
                        contractor_type = cells[3].text.strip()
                    except:
                        contractor_type = ''
                    try:
                        contractor_class = cells[4].text.strip()
                    except:
                        contractor_class = ''
                    try:
                        category = cells[5].text.strip()
                    except:
                        category = ''
                    try:
                        physical_address = cells[6].text.strip()
                    except:
                        physical_address = ''
                    try:
                        postal_address = cells[7].text.strip()
                    except:
                        postal_address = ''
                    try:
                        town = cells[8].text.strip()
                    except:
                        town = ''
                else:
                    log_print("Not enough data!!")
                
                indi_data = [company_name, r_number, director, contractor_type, contractor_class, 
                             category, physical_address, postal_address, town]
                # Write to CSV file
                with open(File_path_search_page_CSV, 'a', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow(indi_data)
                count()
                # Write to TXT file
                with open(File_path_search_page_TXT, 'a', encoding="utf-8") as fw:
                    fw.write("\t".join(map(str, indi_data)) + "\n")
                    fw.flush()
            except:
                exception(search_page_URL)
        with open(File_path_log_index_page, 'w', encoding='utf-8') as file:
            file.write(str(page))
            file.flush()
        log_print(f"Page {page} Compeleted")
        
    except:
        exception(search_page_URL)


def individual_data(details_sp):
    try:
        details_divs = details_sp.find('div', class_='col-sm-9').find_all('div', class_='container')
        # rows = search_table.find_all('tr')[1:]
        
        totDivs = len(details_divs)
        if totDivs==0:
            log_print("No Data Found!!")
            # with open(File_path_log_index1, 'w', encoding='utf-8') as file:
            #     file.write(page)
            #     file.flush()
            return None
        
        indi_data = []
        
        if totDivs == 11:
            try:
                company_name = details_divs[0].text.strip()
            except:
                company_name = ''
            try:
                r_number = details_divs[1].find('b')
                r_number = r_number.text.strip()
            except:
                r_number = ''
            try:
                director = details_divs[2].find('b')
                director = director.text.strip()
            except:
                director = ''
            try:
                contractor_type = details_divs[3].find('b')
                contractor_type = contractor_type.text.strip()
            except:
                contractor_type = ''
            try:
                contractor_class = details_divs[4].find('b')
                contractor_class = contractor_class.text.strip()
            except:
                contractor_class = ''
            try:
                category = details_divs[5].find('b')
                category = category.text.strip()
            except:
                category = ''
            try:
                physical_address = details_divs[6].find('b')
                physical_address = physical_address.text.strip()
            except:
                physical_address = ''
            try:
                postal_address = details_divs[7].find('b')
                postal_address = postal_address.text.strip()
            except:
                postal_address = ''
            try:
                town = details_divs[8].find('b')
                town = town.text.strip()
            except:
                town = ''
            try:
                phone = details_divs[9].find('b')
                phone = phone.text.strip()
            except:
                phone = ''
            try:
                email = details_divs[10].find('b')
                email = email.text.strip()
            except:
                email = ''
        else:
            log_print("Not enough data!!")
        
        indi_data = [company_name, r_number, director, contractor_type, contractor_class, 
                        category, physical_address, postal_address, town, phone, email]
        # Write to CSV file
        with open(File_path_details_page_CSV, 'a', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(indi_data)
        count()
        # Write to TXT file
        with open(File_path_details_page_TXT, 'a', encoding="utf-8") as fw:
            fw.write("\t".join(map(str, indi_data)) + "\n")
            fw.flush()
        with open(File_path_log_index_company, 'w', encoding='utf-8') as file:
            file.write(str(index))
            file.flush()
        log_print(f"Added {company_name}")
    except:
        exception(details_page_URL)


if __name__ == "__main__":
    File_paths = [File_path_search_page_CSV, File_path_search_page_TXT, File_path_details_page_CSV, File_path_details_page_TXT, File_path_error_CSV]
    file_paths_logs = [File_path_log, File_path_log_index_page, File_path_log_index_company]

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
        if os.path.exists(File_path_count):
            os.remove(File_path_count)

    with open(File_path_count, "a") as f:
        f.write("")
        
    HeadersF1 = ["Company Name", "Registration Number", "Managing Director", "Type of contractor", "Class", 
                 "Category", "Physical Address", "Postal address", "Town"]
    with open(File_path_search_page_TXT, "a") as fw:
        if fw.tell() == 0:
            fw.write("\t".join(HeadersF1) + "\n")
            fw.flush()
    with open(File_path_search_page_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(HeadersF1)
    
    HeadersF2 = ["Company Name", "Registration Number", "Managing Director", "Type of contractor", "Class", 
                 "Category", "Physical Address", "Postal address", "Town", "Phone", "Email"]
    with open(File_path_details_page_TXT, "a") as fw:
        if fw.tell() == 0:
            fw.write("\t".join(HeadersF2) + "\n")
            fw.flush()
    with open(File_path_details_page_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(HeadersF2)

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
    headers = {'User-Agent': user_agent}

    Base_URL = 'https://www.crb.go.tz/contractordetails'
    Page_URL = Base_URL + '?page={}'
    Indi_URL = Base_URL + '/home/show/{}'
    st = time.time()
    
    retry_attempts = 5
    retry_delay = 2
    
    try:
        ####################################### Phase 1 #######################################
        
        try:
            soup_temp = request(Base_URL)
            pagination = soup_temp.find('ul', class_='pagination')
            total_pages = int(pagination.find_all('li', class_='page-item')[-2].get_text(strip=True))
            soup_temp.decompose()
            
            log_index_flag = False
            if os.path.exists(File_path_log_index_page):
                log_index_flag = True
                with open(File_path_log_index_page, 'r', encoding='utf-8') as file:
                    try:
                        last_processed_page = int(file.read().strip())
                    except:
                        last_processed_page = ''

            if log_index_flag and last_processed_page != '':
                start_index = last_processed_page + 1
            else:
                start_index = 1
            
            for page in range(start_index, total_pages+1):
                try:
                    search_page_URL = Page_URL.format(page)
                    search_soup = request(search_page_URL)
                    
                    search_data(search_soup)
                    
                except:
                    exception(Page_URL.format(page))
        except:
            exception(Page_URL.format(0))
            
        ####################################### Phase 2 #######################################
        
        try:
            soup_temp = request(Page_URL.format(total_pages))
            temp_table = soup_temp.find('table', class_='table table-hover table-sm table-bordered table-striped')
            temp_rows = temp_table.find_all('tr')[-1]
            last_company_link = temp_rows.find('a')['href']
            last_company_num = int(last_company_link.replace('/contractordetails/home/show/', ''))
            soup_temp.decompose()
            
            log_index_flag = False
            if os.path.exists(File_path_log_index_company):
                log_index_flag = True
                with open(File_path_log_index_company, 'r', encoding='utf-8') as file:
                    try:
                        last_processed_company = int(file.read().strip())
                    except:
                        last_processed_company = ''

            if log_index_flag and last_processed_company != '':
                start_index = last_processed_company + 1
            else:
                start_index = 1
            
            for index in range(start_index, last_company_num+1):
                try:
                    details_page_URL = Indi_URL.format(index)
                    details_soup = request(details_page_URL)
                    
                    individual_data(details_soup)
                except:
                    exception(Indi_URL.format(index))
        except:
            exception(Indi_URL.format(0))
    
    except:
        exception(Base_URL)
        
    finally:
        duplicateFromCSV(File_path_search_page_CSV)
        duplicateFromCSV(File_path_details_page_CSV)
        convertCSVExcelExtended(File_path_search_page_CSV, File_path_search_page)
        convertCSVExcelExtended(File_path_details_page_CSV, File_path_details_page)
        with open(File_path_log_index_company, 'r', encoding='utf-8') as file:
            last_processed_company = file.read().strip()
        if int(last_processed_company) == last_company_num:
            log_print('Script Completed')
            if os.path.exists(File_path_error):
                os.remove(File_path_error)
            if os.path.exists(File_path_log_Run_Flag):
                os.remove(File_path_log_Run_Flag)
            if os.path.exists(File_path_count):
                os.remove(File_path_count)
        et = time.time()
        log_print(f'\n{et - st}')

database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
    for File_path in File_paths:
        delete_task(conn, File_path)
    for File_path in file_paths_logs:
        delete_task(conn, file_paths_logs)
    # delete_task(conn, File_path_search_page)
    # delete_task(conn, File_path_details_page)