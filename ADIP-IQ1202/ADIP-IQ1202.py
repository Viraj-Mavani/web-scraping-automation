import csv
import os
import random
import sys
import requests
import re
import time
from multipledispatch import dispatch
# try:
# 	from PIL import Image
# except ImportError:
# 	import Image
# import pytesseract
# import cv2
import traceback
import sqlite3
from sqlite3 import Error
import pandas as pd
import json
from bs4 import BeautifulSoup
os.environ['PYDEVD_WARN_SLOW_RESOLVE_TIMEOUT'] = '2.0'

arabic_letters = ['ا', 'ب', 'ت', 'ث', 'ج', 'ح', 'خ', 'د', 'ذ', 'ر', 'ز', 'س', 'ش', 
                  'ص', 'ض', 'ط', 'ظ', 'ع', 'غ', 'ف', 'ق', 'ك', 'ل', 'م', 'ن', 'ه', 'و', 'ي']

# BasePath = 'D:\\Projects\\CedarPython\\ADIP-IQ1202' 
BasePath = os.getcwd()
# BasePath= 'E:\\ADIP-PY\\OP2'


######### Excel #########
File_path_search_page = BasePath + '\\OP\\ADIP-IQ1202_search_page.xlsx'
# File_path_details_page = BasePath + '\\OP\\ADIP-IQ1202_details_page.xlsx'
######### CSV #########
File_path_search_page_CSV = BasePath + '\\OPcsv\\ADIP-IQ1202_search_page.csv'
# File_path_details_page_CSV = BasePath + '\\OPcsv\\ADIP-IQ1202_details_page.csv'
File_path_error_CSV = BasePath + '\\OPcsv\\ADIP-IQ1202_Error.csv'
######### Text #########
File_path_search_page_TXT = BasePath + '\\OPtxt\\ADIP-IQ1202_search_page.txt'
# File_path_details_page_TXT = BasePath + '\\OPtxt\\ADIP-IQ1202_details_page.txt'
######### Error #########
File_path_error = BasePath + '\\Error\\ADIP-IQ1202_Error.xlsx'
######### Count #########
File_path_count = BasePath + '\\Counts\\ADIP-IQ1202_Count.txt'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-IQ1202_Log.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-IQ1202_Run_Flag.txt'
File_path_log_index = BasePath + '\\Log\\ADIP-IQ1202_Log_index.txt'


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
    Headers_Error = ['Letter','URL', 'Not Responding', 'Error']
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


def regex_match(regex,content):
	match = re.search(regex,content,re.I)
	if match:
		return match.group(1)
	else:
		return ''


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


def extract_numbers(text):
    numbers = re.findall(r'\d+', text)
    if numbers:
        return int(numbers[0])
    else:
        return 0


@dispatch(str)
def request(req):
	# user_agent = random.choice(user_agents)
	# headers = {'User-Agent': user_agent}
	try:
		Retry = 1
		while Retry <= retry_attempts:
			try:
				r_delay = random.uniform(1.5, 3.0)
				time.sleep(r_delay)
				obj = requests.get(req, timeout=500)
				# obj = requests.get(req, headers, timeout=500)
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
  
  
@dispatch(str, dict)
def request(req, payload):
	# user_agent = random.choice(user_agents)
	# headers = {'User-Agent': user_agent}
	try:
		Retry = 1
		while Retry <= retry_attempts:
			try:
				r_delay = random.uniform(1.5, 3.0)
				time.sleep(r_delay)
				obj = requests.post(req, data=payload, timeout=2500)
				# obj = requests.post(req, headers, timeout=2500)
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

		
def listPageCollector(content):
	try:
		list_count = content.find('span', id='lbl_count').text.strip()
		input_num = extract_numbers(list_count)
		if input_num == 0:
			log_print(f"No Data Found!! for {letter}")
			with open(File_path_log_index, 'w', encoding='utf-8') as file:
				file.write(str(letter))
				file.flush()	
			return None

		try:
			rows = content.find('table', id='GridView1').find_all('tr', recursive=False)[1:]
		except:
			log_print(f"Data table not Found!! for {letter}")
			with open(File_path_log_index, 'w', encoding='utf-8') as file:
				file.write(str(letter))
				file.flush()	
			return None
		
		for row in rows:
			try:
				cells = row.find_all('td')
				indi_data = ['']*15
				totcell = len(cells)
				if totcell == 15:
					for i in range(totcell):
						indi_data[i] = cells[totcell - 1 - i].text.strip() if cells[totcell - 1 - i].text else ''

					# indi_data[0] = cells[14].text.strip() if cells[14].text else ''
					# indi_data[1] = cells[13].text.strip() if cells[13].text else ''
					# indi_data[2] = cells[12].text.strip() if cells[12].text else ''
					# indi_data[3] = cells[11].text.strip() if cells[11].text else ''
					# indi_data[4] = cells[10].text.strip() if cells[10].text else ''
					# indi_data[5] = cells[9].text.strip() if cells[9].text else ''
					# indi_data[6] = cells[8].text.strip() if cells[8].text else ''
					# indi_data[7] = cells[7].text.strip() if cells[7].text else ''
					# indi_data[8] = cells[6].text.strip() if cells[6].text else ''
					# indi_data[9] = cells[5].text.strip() if cells[5].text else ''
					# indi_data[10] = cells[4].text.strip() if cells[4].text else ''
					# indi_data[11] = cells[3].text.strip() if cells[3].text else ''
					# indi_data[12] = cells[2].text.strip() if cells[2].text else ''
					# indi_data[13] = cells[1].text.strip() if cells[1].text else ''
					# indi_data[14] = cells[0].text.strip() if cells[0].text else ''

				else:
					log_print(f"Not enough data!! for {letter}")

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
				exception()	
		log_print(f"Completed {letter}")
		with open(File_path_log_index, 'w', encoding='utf-8') as file:
			file.write(str(letter))
			file.flush()
	except:
		exception()	


if __name__ == "__main__":	
	# File_paths= ['E:\ADIP-PY\OP\ADIP-IQ1202_Searchpage.xlsx','E:\ADIP-PY\OPtxt\ADIP-IQ1202_Searchpage.txt','E:\ADIP-PY\Error\ADIP-IQ1202_Error.xlsx','E:\ADIP-PY\Counts\ADIP-IQ1202_Count.txt']
	# cachePath = 'Cache_IQ1202/'
	# if not os.path.isdir(cachePath):
	# 	os.makedirs(cachePath)
	
	File_paths = [File_path_search_page_CSV, File_path_search_page_TXT, File_path_error_CSV]
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
		for path_files in File_paths:
			if os.path.exists(path_files):
				os.remove(path_files)
		if os.path.exists(File_path_count):
			os.remove(File_path_count)

	with open(File_path_count, "a") as f:
		f.write("")
		
	HeadersF1 = ["File Number","Company Name","Money capital","Date of certification","Certificate Number","Registered in Kurdistan",
              "Blacklist","Filter","Merge","Delete","reservation","Number of shares","Name","Adjective","Address"]
	with open(File_path_search_page_TXT, "a") as fw:
		if fw.tell() == 0:
			fw.write("\t".join(HeadersF1) + "\n")
			fw.flush()
	with open(File_path_search_page_CSV, 'a', newline='', encoding='utf-8') as file:
		writer = csv.writer(file)
		if file.tell() == 0:
			writer.writerow(HeadersF1)
 
	# with open('E:\ADIP-PY\Counts\ADIP-IQ1202_Count.txt',"w")as f:
	# 	f.write("")	
	# sess=requests.session()
	# sess.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.183 Safari/537.36'
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
		headers = {'User-Agent': user_agent}

		Base_URL = "https://tasjeel.mot.gov.iq/search_n1/national_page"
		st = time.time()
		
		retry_attempts = 5
		retry_delay = 2
	
		t_letter_combinations = [f'{a}{b}{c}' for a in arabic_letters for b in arabic_letters for c in arabic_letters]
		# t_letter_combinations = [f'{a}{b}{c}{d}' for a in arabic_letters for b in arabic_letters for c in arabic_letters for d in arabic_letters]

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

		success_flag = False

		for letter in letter_num_list[:]:
			try:
				home_content = request(Base_URL)
				VIEWSTATE = home_content.find('input', id='__VIEWSTATE')['value']
				VIEWSTATEGENERATOR = home_content.find('input', id='__VIEWSTATEGENERATOR')['value']
				EVENTVALIDATION = home_content.find('input', id='__EVENTVALIDATION')['value']
				VIEWSTATE = home_content.find('input', id='__VIEWSTATE')['value']
				pay_load={
					"__VIEWSTATE": VIEWSTATE,
					"__VIEWSTATEGENERATOR": VIEWSTATEGENERATOR,
					"__EVENTVALIDATION": EVENTVALIDATION,
					"btn_search": "بحث",
					"txt_search": letter}
		
				listPage = request(Base_URL, pay_load)
				listPageCollector(listPage) 
			except:
				exception()
	
		duplicateFromCSV(File_path_search_page_CSV)
		convertCSVExcelExtended(File_path_search_page_CSV, File_path_search_page)
  
	except:
		exception()
  
	finally:
		if os.path.exists(File_path_log_index):
			with open(File_path_log_index, 'r', encoding='utf-8') as file:
				last_letter = file.read().strip()
		if last_letter == letter_num_list[-1]:
			log_print('Script Completed')
			if os.path.exists(File_path_count):
				os.remove(File_path_count)
			if os.path.exists(File_path_log_Run_Flag):
				os.remove(File_path_log_Run_Flag)
		else:
			log_print('Script was Stopped')
  
database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
	for File_path in File_paths:
		delete_task(conn, File_path)