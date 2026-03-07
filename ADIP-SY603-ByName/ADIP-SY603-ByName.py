import csv
import os
import random
import string
import time
import pandas as pd
# import time
import requests
# import re
from bs4 import BeautifulSoup
import sqlite3
from sqlite3 import Error
import traceback
import sys
from requests.adapters import HTTPAdapter
# from requests.packages.urllib3.util.retry import Retry
from requests.exceptions import RequestException
from urllib3.exceptions import ConnectTimeoutError

# BasePath = 'D:\Projects\CedarPython\ADIP-SY603-ByName'
BasePath = os.getcwd()
######### Excel #########
File_path= BasePath +'\OP\ADIP-SY603-ByName_Output.xlsx'
######### CSV #########
File_path_CSV = BasePath +'\OPcsv\ADIP-SY603-ByName_Output.csv'
File_path_error_CSV = BasePath + '\OPcsv\ADIP-SY603-ByName_Error.csv'
######### Text #########
File_path_txt= BasePath +'\OPtxt\ADIP-SY603-ByName_Output.txt'
######### Input #########
File_path_Input = BasePath + '\\InputFile\\ADIP-SY603-ByName-Input.xlsx'
######### Count #########
File_path_count= BasePath +'\Counts\ADIP-SY603-ByName_Count.txt'
######### Failed #########
File_path_failed_Arabic = BasePath + '\\OP\\ADIP-SY603-ByName_Failed_Arabic.xlsx'
File_path_failed_Arabic_CSV = BasePath + '\\OPcsv\\ADIP-SY603-ByName_Failed_Arabic.csv'
######### Error #########
File_path_error= BasePath +'\Error\ADIP-SY603-ByName_Error.xlsx'
######### Log #########
File_path_log = BasePath + '\Log\ADIP-SY603-ByName_Log.txt'
File_path_log_index_Arabic = BasePath + '\\Log\\ADIP-SY603-ByName_Log_Index_Arabic.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-SY603-ByName_Run_Flag.txt'
######### Proxy #########
# File_path_Input = BasePath + '\Proxy\http_proxies.xlsx'


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
	# global rowError
	headers = ['URL', 'Not Responding', 'Error']
	error = traceback.format_exc()
	exception_type, exception_object, exception_traceback = sys.exc_info()
	with open(File_path_error_CSV, 'a', newline='', encoding='utf-8') as file:
		writer = csv.writer(file)
		if file.tell() == 0:
			writer.writerow(headers)
		writer.writerow([Base_url, "Not Responding", error])
	df = pd.read_csv(File_path_error_CSV, encoding='utf-8')
	df.to_excel(File_path_error, index=False)


def Individual_data(data):
	headers = ['Name', 'Phone', 'Address', 'Activity']
	
	for item in data:
		Indi_data = []
		Name_tag = item.find('i', {"class": ['fa', 'fa-user']})
		name = Name_tag.next if Name_tag.next else ''
		Body_data = item.find_all('li')

		phone = ''
		address = ''
		activity = ''

		try:
			for li in Body_data:
				if li.next.next == " هاتف : ":
					phone = li.find('span').string if li.find('span') else ''
				elif li.next.next == " العنوان: ":
					address = li.find('span').string if li.find('span') else ''
				else:
					activity = li.string if li.string else ''

			Indi_data.append(name)
			Indi_data.append(phone)
			Indi_data.append(address)
			Indi_data.append(activity)

			with open(File_path_CSV, 'a', newline='', encoding='utf-8') as file:
				writer = csv.writer(file)
				if file.tell() == 0: 
					writer.writerow(headers)
				writer.writerow([name, phone, address, activity])
			count()
			with open(File_path_txt,'a',encoding="utf-8") as fw:
				fw.write("\t".join(map(str,Indi_data))+"\n")
				fw.flush()
			with open(File_path_log_index_Arabic, 'w', encoding='utf-8') as f3:
				f3.write(indexA1)
				f3.flush()
		except:
			exception()


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


def duplicate(File_path):
    try:
        data = pd.read_excel(File_path)
        data_file = data.drop_duplicates()
        data_file.to_excel(File_path, index=False)
    except:
        pass
		

if __name__=='__main__':

	# Create directories if they don't exist
	directories = [
        BasePath + '\Counts',
        BasePath + '\Error',
        BasePath + '\Log',
        BasePath + '\OP',
        BasePath + '\InputFile',
        BasePath + '\OPcsv',
        BasePath + '\Optxt',
        BasePath + '\Proxy'
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
		if os.path.isfile(File_path_log_index_Arabic):
			os.remove(File_path_log_index_Arabic)
		if os.path.isfile(File_path_txt):
			os.remove(File_path_txt)
		if os.path.isfile(File_path_count):
			os.remove(File_path_count)
		if os.path.isfile(File_path_failed_Arabic_CSV):
			os.remove(File_path_failed_Arabic_CSV)

	Search_headers = ['Company Name','Phone','Adrress','Activity']
	if not os.path.exists(File_path_txt):
		with open(File_path_txt,"a")as f:
			f.write("\t".join(Search_headers)+"\n")
			f.flush()
	if not os.path.exists(File_path_count):
		with open(File_path_count,"a")as f:
			f.write("")

	# df = pd.read_excel(File_path_Input, sheet_name='Sheet1')
	# proxies = df['proxies'].tolist()
 
	Base_url = 'http://hamachamber.com/members-index/?ftxt={}&ftxt2=&searchType=1&count={}'
 
	log_print('Data Importing...plz wait')
	df = pd.read_excel(File_path_Input, sheet_name='Sheet1')
	names_list = df['nameLocal'].tolist()
	log_print('Data Imported\n')

	retry_attempts = 5
	retry_delay = 2

	log_index_flag = False
	if os.path.exists(File_path_log_index_Arabic):
		log_index_flag = True
		with open(File_path_log_index_Arabic, 'r', encoding='utf-8') as file:
			last_processed_name = file.read().strip()

	if log_index_flag:
		start_index = names_list.index(last_processed_name) + 1
		names_list = names_list[start_index:]
	else:
		names_list = names_list

	try:
		for indexA1 in names_list[:]:
			# proxy = random.choice(proxies)
			# proxy_url = f'http://{proxy}'
			# log_print(f"Using proxy: {proxy}")
			outerRetry = 1
			error_message_flag = False
			while outerRetry <= retry_attempts:
				try:
					# obj_temp = requests.get(Base_url.format(indexA1, 1), proxies={'http': proxy_url, 'https': proxy_url})
					obj_temp = requests.get(Base_url.format(indexA1, 1))
				except (ConnectTimeoutError, RequestException) as e:
					log_print(f"Error occurred for {indexA1}")
					exception()
					delay = retry_delay * (2 ** outerRetry)
					log_print(f'Retrying in {delay} seconds...{outerRetry}')
					time.sleep(delay)
					outerRetry += 1
					continue
				else:
					soup_temp = BeautifulSoup(obj_temp.content, 'html.parser')
					
					error_message = soup_temp.find('div', class_='alert alert-danger')
					if error_message:
						error_message_flag = True
						break
					
					if len(soup_temp.find_all('a', class_='page-link')) > 0:
						last_page_element = soup_temp.find_all('a', class_='page-link')[-2]
						last_page_number = int(last_page_element.get_text(strip=True))
						del obj_temp
						del soup_temp
					else:
						last_page_number = 1
					for index in range(1, last_page_number + 1):
						innerRetry = 1
						while innerRetry <= retry_attempts:
							try:
								obj = requests.get(Base_url.format(indexA1, index), timeout=300)
							except (ConnectTimeoutError, RequestException) as e:
								log_print(f"Error occurred for {indexA1}")
								exception()
								delay = retry_delay * (2 ** innerRetry)
								log_print(f'Retrying in {delay} seconds...RETRY: {innerRetry}')
								time.sleep(delay)
								innerRetry += 1
								continue
							else:
								soup = BeautifulSoup(obj.content, 'html.parser')
								res = soup.find_all('div', class_='mycalls')[1:]
								Individual_data(res)

								if os.path.exists(File_path_log_index_Arabic):
									with open(File_path_log_index_Arabic, 'r', encoding='utf-8') as file:
										last_processed_name = file.read().strip()
								if indexA1 == last_processed_name:
									log_print(f'Success {indexA1} {index}')
								else:
									log_print(f'Failed!! {indexA1} {index}')
									with open(File_path_failed_Arabic_CSV, 'a', newline='', encoding='utf-8') as file:
										writer = csv.writer(file)
										writer.writerow([indexA1])
								break
						else:
							with open(File_path_failed_Arabic_CSV, 'a', newline='', encoding='utf-8') as file:
								writer = csv.writer(file)
								writer.writerow([indexA1])
							log_print(f'Failed!! {indexA1} {index}')
							continue
					break
			else:
				log_print(f'{indexA1} Failed')
				continue
			if error_message_flag:
				error_message_flag = False
				log_print(f'Data Not Found {indexA1}')
				with open(File_path_log_index_Arabic, 'w', encoding='utf-8') as f3:
					f3.write(indexA1)
					f3.flush()
				continue

	except:
		exception()
	finally:
		convertCSVExcel(File_path_CSV, File_path)
		data = pd.read_excel(File_path)
		data_file = data.drop_duplicates()
		data_file.to_excel(File_path, index=False)
		log_print("\nComplete")	
		exit()
	
database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
	delete_task(conn, File_path)
	delete_task(conn, File_path_CSV)
	delete_task(conn, File_path_txt)
	delete_task(conn, File_path_count)
	delete_task(conn, File_path_error)
	delete_task(conn, File_path_error_CSV)