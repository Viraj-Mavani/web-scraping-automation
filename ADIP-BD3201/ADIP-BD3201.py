# import base64
import os
import pytesseract
# import imageio.v3 as iio
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
import requests

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

def index_file_read(File_index):
    if os.path.exists(File_index):
        with open(File_index, 'r', encoding='utf-8') as file:
            return file.read().strip()

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


def img2txt():
    # log_print("Resampling the Image")
    # image = iio.imread(og_image_path)
    # iio.imwrite(converted_image_path, image)
    
    img = cv2.imread(og_image_path)                                                                         #import image data
    gray_img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)                                                        #convert to grayscale
    _, threshold_img = cv2.threshold(gray_img, 190, 255, cv2.THRESH_BINARY_INV)                             #threshold image
    blurred_img = cv2.GaussianBlur(threshold_img, (5, 5), 0)                                                #BLUR image
    return pytesseract.image_to_string(blurred_img,config="-c tessedit_char_whitelist=0123456789 --psm 10")      #img to txt


def loader():
    pass


def img_dl(image_path, img_url):
    response = requests.get(img_url, timeout=200, verify=False)
    # img_soup = BeautifulSoup(response.content, 'html.parser')
    # img_tag = img_soup.find('img', src=img_url)
    with open(image_path, 'wb') as handler:
        handler.write(response.content)


def captcha(obj):
    p_hash = obj.select_one('input[name="p_hash"]').get('value')
    captcha_img_url = 'https://app.roc.gov.bd/psp/nc_cap?p_hash=' + p_hash
    img_dl(og_image_path, captcha_img_url) 
    captcha_text = img2txt()
    return p_hash, captcha_text.strip()


def request(payload):
    # retry_attempts = 1
    # retry_delay = 2
    while True:
        try:
            # Retry = 1
            # while Retry <= retry_attempts:
            try:
                obj = requests.post(Base_URL, data=payload, timeout=200, verify=False)
                # break
            except Exception as e:
                log_print(f"Error occurred in Request")
                # delay = retry_delay * (2 ** Retry)
                # log_print(f'Retrying in {delay} seconds...{Retry}')
                # time.sleep(delay)
                # Retry += 1
                # continue
                exception()
                return None
            # else:
                # os._exit(1)
            soup = BeautifulSoup(obj.content,'html.parser')
            form_element = soup.find('form', action='nc_search')
            error_element = form_element.find('b', string=' Incorrect Code- Please try again.')

            if error_element is None:
                soup.decompose()
                break
            else:
                p_hash = soup.select_one('input[name="p_hash"]').get('value')
                captcha_img_url = 'https://app.roc.gov.bd/psp/nc_cap?p_hash=' + p_hash
                img_dl(og_image_path, captcha_img_url) 
                captcha_text = img2txt()
                captcha_value = captcha_text.strip()
                payload['p_captcha'] = captcha_value
                payload['p_hash'] = p_hash
                continue
        except:
            exception()
            time.sleep(200)
    return form_element

def individual_data(rows):
    try:
        data_success = False
        indi_data = []
        
        for row in rows:
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
                    status = status_registration
                    registration_number = ""
                    
            data = {
            "Entity Name": entity_name,
            "Entity Type": entity_type,
            "Status": status,
            "Registration Number": registration_number
            }
            indi_data.append(data)
            count()
        
        # Write to CSV file
        csv_df = pd.DataFrame(indi_data)
        csv_df.to_csv(File_path_CSV, index=False, mode='a', header=not os.path.exists(File_path_CSV))
        data_success = True
        # Write to TXT file
        with open(File_path_txt, "a") as f:
            for item in indi_data:
                f.write("\t".join(map(str, item.values())) + "\n")
        
        return data_success
        

    except:
        data_success = False
        exception()
        return data_success
    

if __name__ == "__main__":
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
        st = time.time()
        
        f_req = requests.get(Base_URL, timeout=200, verify=False)
        f_soup = BeautifulSoup(f_req.content,'html.parser')
        p_hash, captcha_value = captcha(f_soup)

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
                    try:
                        letter = letterE1 + letterE2 + letterE3
                        success = False
                        
                        fields = {
                            'entity_type':'1',
                            'search_text': letter,
                            'CB':'1',
                            'p_captcha':captcha_value,
                            'p_hash':p_hash,
                            'result_type':'0',
                            'p_entry_mode':'3',
                            'page_no':'1'
                            }
                        form_element = request(fields)
                        
                        if form_element is None:
                            log_print(f'Failed!! {letter}\n')
                            with open(File_path_log_index_LetterE3, 'w', encoding='utf-8') as file:
                                file.write(letterE3)
                                file.flush()
                            continue
                        
                        table2 = form_element.find('table', id='AutoNumber2')
                        
                        index_table = table2.find('table', id='AutoNumber3')
                        try:
                            page_element = index_table.find('font')
                            page_text = page_element.get_text(strip=True)
                            total_page = int(page_text.split()[-1])
                        except:
                            total_page = 1
                        
                        sl_element = table2.find('b', string='SL.')
                        data_table_rows = sl_element.find_parents('td')[1].find('table').find_all('tr')[1:]
                        
                        if total_page < 2 and len(data_table_rows)==0:
                            log_print(f'No Data for {letter}')
                            success = True
                        else:
                            success = individual_data(data_table_rows)
                            if success:
                                log_print('Complete ' + letter + ' for page: 1')
                            else:
                                log_print('Failed!! ' + letter + ' for page: 1')
                            for index in range(2, total_page+1):
                                p_hash, captcha_value = captcha(form_element)
                                page_fields = {
                                    'entity_type':'1',
                                    'search_text': letter,
                                    'CB':'1',
                                    'p_captcha':captcha_value,
                                    'p_hash':p_hash,
                                    'result_type':'0',
                                    'p_entry_mode':'3',
                                    'page_no':index
                                    }
                                form_element = request(page_fields)
                                if form_element is not None:
                                    table2 = form_element.find('table', id='AutoNumber2')
                                    sl_element = table2.find('b', string='SL.')
                                    data_table_rows = sl_element.find_parents('td')[1].find('table').find_all('tr')[1:]
                                    success = individual_data(data_table_rows)
                                    if success:
                                        log_print(f'Complete {letter} for page: {str(index)}')
                                    else:
                                        log_print(f'Failed!! {letter} for page: {str(index)}')
                                else:
                                    success = False
                                    break
                                
                        if success:
                            log_print(f'Complete {letter}\n')
                        else:
                            log_print(f'Failed!! {letter}\n')

                        with open(File_path_log_index_LetterE3, 'w', encoding='utf-8') as file:
                            file.write(letterE3)
                            file.flush()
                    except:
                        exception()
                        # log_print(f'Failed!! {letter}\n')
                        # with open(File_path_log_index_LetterE3, 'w', encoding='utf-8') as file:
                        #     file.write(letterE3)
                        #     file.flush()
                with open(File_path_log_index_LetterE2, 'w', encoding='utf-8') as file:
                    file.write(letterE2)
                    file.flush()
                start_index_LetterE3 = 0
            with open(File_path_log_index_LetterE1, 'w', encoding='utf-8') as f1:
                f1.write(letterE1)
                f1.flush()
            start_index_LetterE2 = 0
        
    except:
        exception()
        
    finally:
        # driver.close()
        convertCSVExcel(File_path_CSV, File_path)
        duplicate(File_path)
        et = time.time()
        log_print(f'\n{et - st}')
        last_index_LetterE1 = index_file_read(File_path_log_index_LetterE1)
        last_index_LetterE2 = index_file_read(File_path_log_index_LetterE2)
        last_index_LetterE3 = index_file_read(File_path_log_index_LetterE3)
        last_letter = last_index_LetterE1 + last_index_LetterE2 + last_index_LetterE3
        if last_letter == "ZZZ":
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
	delete_task(conn, File_paths)
	delete_task(conn, file_paths_logs)
