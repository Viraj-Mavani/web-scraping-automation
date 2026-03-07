# import base64
import os
import string
# import pytesseract
# import imageio.v3 as iio
# import cv2
import csv
import sys
import traceback
import pandas as pd
import sqlite3
from sqlite3 import Error
import time
import requests
import json

BasePath = 'D:\\Projects\\CedarPython\\ADIP-UG3101\\'
# BasePath = os.getcwd()
# BasePath= 'E:\\ADIP-PY\\OP2'

######### Excel #########
File_path = BasePath + '\\OP\\ADIP-UG3101_Output.xlsx'
######### CSV #########
File_path_CSV = BasePath + '\\OPcsv\\ADIP-UG3101_Output.csv'
File_path_error_CSV = BasePath + '\\OPcsv\\ADIP-UG3101_Error.csv'
######### Text #########
File_path_txt = BasePath + '\\Optxt\\ADIP-UG3101_Output.txt'
######### Error #########
File_path_error = BasePath + '\\Error\\ADIP-UG3101_Error.xlsx'
######### Count #########
File_path_count = BasePath + '\\Counts\\ADIP-UG3101_Count.txt'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-UG3101_Log.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-UG3101_Run_Flag.txt'
File_path_log_index = BasePath + '\\Log\\ADIP-UG3101_Log_Index.txt'
######### IMG #########
# og_image_path = BasePath + '\\Log\\ADIP-UG3101_Captcha.png'
# converted_image_path = BasePath + '\\Log\\ADIP-UG3101_Converted_img.png'


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


# def img2txt():
    # log_print("Resampling the Image")
    # image = iio.imread(og_image_path)
    # iio.imwrite(converted_image_path, image)

    # img = cv2.imread(og_image_path)  # import image data
    # gray_img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)  # convert to grayscale
    # _, threshold_img = cv2.threshold(
    #     gray_img, 190, 255, cv2.THRESH_BINARY_INV)  # threshold image
    # blurred_img = cv2.GaussianBlur(threshold_img, (5, 5), 0)  # BLUR image
    # # img to txt
    # return pytesseract.image_to_string(blurred_img, config="-c tessedit_char_whitelist=0123456789 --psm 10")


# def img_dl(image_path, img_url):
#     response = requests.get(img_url, timeout=200, verify=False)
#     # img_soup = BeautifulSoup(response.content, 'html.parser')
#     # img_tag = img_soup.find('img', src=img_url)
#     with open(image_path, 'wb') as handler:
#         handler.write(response.content)


# def captcha():
#     # captcha_img_url = 'https://app.roc.gov.bd/psp/nc_cap?p_hash=' 
#     # img_dl(og_image_path, captcha_img_url)
#     # captcha_text = img2txt()
#     captcha_text = 11111
#     return captcha_text.strip()


def request(pl):
    try:
        Retry = 1
        while Retry <= retry_attempts:
            try:
                obj = requests.post(Base_URL, data=json.dumps(pl), timeout=200)
                json_data = obj.json()
                return json_data
            except Exception:
                exception()
                log_print(f"Error occurred in Request")
                delay = retry_delay * (2 ** Retry)
                log_print(f'Retrying in {delay} seconds...{Retry}')
                time.sleep(delay)
                Retry += 1
                continue
        else:
            os._exit(1)
    except:
        exception()


def individual_data(json):
    try:
        data_success = False
        indi_data = []
        data = {}
        if json.get("TotalRecordCount") > 0:
            records = json.get("Records", [])
            
            status_mapping = {
                "NRS-CH": "Changed",
                "NRS-USE": "Registered",
                "NRS-OVR": "Overdue",
                "NRS-CS": "Ceased",
                "NRS-IN": "In use",
                "NRS-RS": "Reserved"
            }

            entity_type_mapping = {
                'EST-BN': "Business name",
                'EST-FC': "Foreign company",
                'EST-GP': "General Partnership",
                'EST-JV': "Joint Venture",
                'EST-LP': "Limited Partnership",
                'EST-PCLG': "Private / Company limited by guarantee",
                'EST-PCLS': "Private / Company limited by shares",
                'EST-PC': "Public company"
            }

            for record in records:
                data = {
                    "Nr" : record[6], 
                    "Registration / Reservation date" : record[5], 
                    "Name" : record[1], 
                    "Similarity" : record[9],
                    "Business/entity type": entity_type_mapping.get(record[4], record[4]), 
                    "Status": status_mapping.get(record[2], record[2]),
                    "Dissolve / Expiry date" : record[7], 
                    "Name changed" : record[8], 
                    "Is compliant" : record[3]
                }
                
                indi_data.append(data)
                count()
        else:
            # log_print(f'No results found for {letter} & Type: {type}')
            data_success = True
            return data_success

        # Write to CSV file
        csv_df = pd.DataFrame(indi_data)
        csv_df.to_csv(File_path_CSV, index=False, mode='a', header=not os.path.exists(File_path_CSV))
        # Write to TXT file
        with open(File_path_txt, "a") as f:
            for item in indi_data:
                f.write("\t".join(map(str, item.values())) + "\n")
        data_success = True
        return data_success

    except:
        data_success = False
        exception()
        return data_success


if __name__ == "__main__":
    File_paths = [File_path_CSV, File_path_txt, File_path_error_CSV]
    file_paths_logs = [File_path_log, File_path_log_index]
    
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

    Headers = ["Nr", "Registration / Reservation date", "Name", "Similarity", "Business/entity type", 
               "Status", "Dissolve / Expiry date", "Name changed", "Is compliant"]
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
        Base_URL = 'https://brs.ursb.go.ug/brs/pro/bnr/list/pub/search/name'
        st = time.time()
        
        retry_attempts = 5
        retry_delay = 2
        split=5
        
        number_combinations = [str(a) + str(b) + str(c) for a in range(10) for b in range(10) for c in range(10)]
        letter_combinations = [a + b + c for a in string.ascii_lowercase for b in string.ascii_lowercase for c in string.ascii_lowercase]
        
        letter_num_list = letter_combinations
        if os.path.exists(File_path_log_index):
            with open(File_path_log_index, 'r', encoding = 'utf-8') as file:
                last_letter  =  file.read().strip()
            if last_letter in letter_combinations:
                start  =  letter_combinations.index(last_letter)+1
                letter_num_list =  letter_combinations[start:]
              
        for letter in letter_num_list:
            try:  
                ent_type = ['EST-BN', 'EST-FC', 'EST-GP', 'EST-JV', 'EST-LP', 'EST-PCLG', 'EST-PCLS', 'EST-PC']
                # sorting = ['reg_nr ASC', 'reg_nr DESC', 'reg_date ASC', 'reg_date DESC',
                #             'name ASC', 'name DESC', 'similarity DESC', 'similarity DESC']
                
                for type in ent_type:
                    # for sort in sorting:
                    payload = {
                        "ent_name": letter,
                        # "Sorting": sort,
                        "ent_type": type
                    }
                    
                    json_data = request(payload)
                    success = individual_data(json_data)

                    if success:
                        log_print(f'{letter} for Type: {type}')
                    else:
                        log_print(f'Failed!! {letter} for Type: {type}')

                if success:
                    log_print(f'Complete {letter}\n')
                else:
                    log_print(f'Failed!! {letter}\n')
                    
                with open(File_path_log_index, 'w', encoding='utf-8') as file:
                    file.write(letter)
                    file.flush()
            except:
                exception()
    except:
        exception()
        
    finally:
        convertCSVExcel(File_path_CSV, File_path)
        duplicate(File_path)
        et = time.time()
        log_print(f'\n{et - st}')
        
        if os.path.exists(File_path_log_index):
            with open(File_path_log_index, 'r', encoding = 'utf-8') as file:
                last_letter = file.read().strip()
        if last_letter == letter_combinations[-1]:
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
