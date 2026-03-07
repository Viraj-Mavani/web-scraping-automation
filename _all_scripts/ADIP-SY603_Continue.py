import os
import random
import string
import time
import pandas as pd
# import time
import requests
# import re
import xlsxwriter
from bs4 import BeautifulSoup
# import openpyxl
# from openpyxl.styles import Font
# from requests_toolbelt.multipart.encoder import MultipartEncoder
import sqlite3
from sqlite3 import Error
import traceback
import sys
from requests.adapters import HTTPAdapter
# from requests.packages.urllib3.util.retry import Retry
from requests.exceptions import RequestException
from urllib3.exceptions import ConnectTimeoutError

BasePath = 'D:\Projects\CedarPython\ADIP-SY603'

File_path= BasePath +'\OP\ADIP-SY603_Output.xlsx'
File_path_txt= BasePath +'\OPtxt\ADIP-SY603_Output.txt'
File_path_count= BasePath +'\Counts\ADIP-SY603_Count.txt'
File_path_error= BasePath +'\Error\ADIP-SY603_Error.xlsx'
######### Log #########
File_path_log = BasePath + '\Log\Log.txt'
File_path_log_index_LetterA1 = BasePath + '\Log\Log_Index_LetterA1.txt'
File_path_log_index_LetterA2 = BasePath + '\Log\Log_Index_LetterA2.txt'
File_path_log_index_LetterA3 = BasePath + '\Log\Log_Index_LetterA3.txt'
File_path_log_index_LetterE1 = BasePath + '\Log\Log_Index_LetterE1.txt'
# File_path_Input = BasePath + '\Proxy\http_proxies.xlsx'

persian_alphabet_list = [
    "ا", "ب","پ","ت","ث","ج","چ","ح","خ","د","ذ","ر","ز","ژ","س","ش",
	"ص","ض","ط","ظ","ع","غ","ف","ق","ک","گ","ل","م","ن","و","ه","ی"
]

english_alphabet_list = list(string.ascii_lowercase)

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
	global rowError
	error = traceback.format_exc()
	exception_type, exception_object, exception_traceback = sys.exc_info()
	worksheet_error.write(rowError, 0, Base_url)
	worksheet_error.write(rowError, 1, "Not Responding")
	worksheet_error.write(rowError, 2, error)
	rowError += 1


def Individual_data(data):
	global row1, rowError

	for item in data:
		Indi_data = ['']*4
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
	
			sheet1.write(row1,0 , name)
			sheet1.write(row1,1 , phone)
			sheet1.write(row1,2 , address)
			sheet1.write(row1,3 , activity)
			row1 += 1

			try_count=1
			while True:
				try:
					with open(File_path_count,'a') as fh:
						fh.write('1\n')
						fh.flush()
					break
				except:
					if try_count>5:
						break
					try_count+=1

			with open(File_path_txt,'a',encoding="utf-8") as fw:
				fw.write("\t".join(map(str,Indi_data))+"\n")
				fw.flush()
		except:
			exception()

if __name__=='__main__':
	
	row1=1
	rowError=1

	# Create directories if they don't exist
	directories = [
        BasePath + '\Counts',
        BasePath + '\Error',
        BasePath + '\Log',
        BasePath + '\OP',
        BasePath + '\Optxt',
        BasePath + '\Proxy'
    ]

	for directory in directories:
		if not os.path.exists(directory):
			os.makedirs(directory)
	
	if os.path.isfile(File_path_log):
		os.remove(File_path_log)

	book1 = xlsxwriter.Workbook(File_path)
	# book1 = openpyxl.Workbook()
	sheet1 = book1.add_worksheet()
	bold_format = book1.add_format({'bold': True})
	# sheet1 = book1.active
	sheet1.write('A1', 'Company Name', bold_format)
	sheet1.write('B1', 'Phone!', bold_format)
	sheet1.write('C1', 'Address', bold_format)
	sheet1.write('D1', 'Acitivity', bold_format)
	
	workbook_error = xlsxwriter.Workbook(File_path_error)
	worksheet_error = workbook_error.add_worksheet()

	bold_format = workbook_error.add_format({'bold': True})

	worksheet_error.write('A1', 'URL', bold_format)
	worksheet_error.write('B1', 'Not Responding', bold_format)
	worksheet_error.write('C1', 'Error', bold_format)

	Search_headers = ['Company Name','Phone','Adrress','Activity']
	with open(File_path_txt,"w")as f:
		f.write("\t".join(Search_headers)+"\n")
		f.flush()
	with open(File_path_count,"w")as f:
		f.write("")
		f.flush()

	# df = pd.read_excel(File_path_Input, sheet_name='Sheet1')
	# proxies = df['proxies'].tolist()
 
	Base_url = 'http://hamachamber.com/members-index/?ftxt={}&ftxt2=&searchType=1&count={}'

	retry_attempts = 5
	retry_delay = 2

	log_index_flag_LetterA1 = False
	log_index_flag_LetterA2 = False
	log_index_flag_LetterA3 = False
 
	if os.path.exists(File_path_log_index_LetterA1):
		with open(File_path_log_index_LetterA1, 'r', encoding='utf-8') as file:
			index_LetterA1 = file.read().strip()
			if index_LetterA1 != '':
				log_index_flag_LetterA1 = True

	if os.path.exists(File_path_log_index_LetterA2):
		with open(File_path_log_index_LetterA2, 'r', encoding='utf-8') as file:
			index_LetterA2 = file.read().strip()
			if index_LetterA2 != '':
				log_index_flag_LetterA2 = True

	if os.path.exists(File_path_log_index_LetterA3):
		with open(File_path_log_index_LetterA3, 'r', encoding='utf-8') as file:
			index_LetterA3 = file.read().strip()
			if index_LetterA3 != '':
				log_index_flag_LetterA3 = True
		
	if log_index_flag_LetterA1:
		start_index = persian_alphabet_list.index(index_LetterA1)
		persian_alphabet1 = persian_alphabet_list[start_index:]
		log_index_flag_LetterA1 = False
	else:
		persian_alphabet1 = persian_alphabet_list
  
	if log_index_flag_LetterA2:
		start_index = persian_alphabet_list.index(index_LetterA2)
		persian_alphabet2 = persian_alphabet_list[start_index:]
		log_index_flag_LetterA2 = False
	else:
		persian_alphabet2 = persian_alphabet_list
  
	if log_index_flag_LetterA3:
		start_index = persian_alphabet_list.index(index_LetterA3)
		persian_alphabet3 = persian_alphabet_list[start_index:]
		log_index_flag_LetterA3 = False
	else:
		persian_alphabet3 = persian_alphabet_list

	try:
		for letter1 in persian_alphabet1[:]:
			for letter2 in persian_alphabet2[:]:
				for letter3 in persian_alphabet3[:]:
					letter = letter1 + letter2 + letter3
					# proxy = random.choice(proxies)
					# proxy_url = f'http://{proxy}'
					# log_print(f"Using proxy: {proxy}")
					outerRetry = 1
					error_message_flag = False
					while outerRetry <= retry_attempts:
						if error_message_flag:
							error_message_flag = False
							break
						try:
							# obj_temp = requests.get(Base_url.format(letter, 1), proxies={'http': proxy_url, 'https': proxy_url})
							obj_temp = requests.get(Base_url.format(letter, 1))
						except (ConnectTimeoutError, RequestException) as e:
							log_print(f"Error occurred for {letter1} {letter2} {letter3}")
							exception()
							delay = retry_delay * (2 ** outerRetry)
							log_print(f'Retrying in {delay} seconds...{outerRetry}')
							time.sleep(delay)
							outerRetry += 1
							continue

							# if retry < retry_attempts - 1:
							# else:
							# 	raise e
						else:
							soup_temp = BeautifulSoup(obj_temp.content, 'html.parser')
							
							error_message = soup_temp.find('div', class_='alert alert-danger')
							if error_message:
								error_message_flag = True
							
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
										obj = requests.get(Base_url.format(letter, index), timeout=300)
									except (ConnectTimeoutError, RequestException) as e:
										log_print(f"Error occurred for {letter1} {letter2} {letter3}")
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
										log_print(f'Success {letter1} {letter2} {letter3} {index}')
										break
								else:
									log_print(f'Failed {letter1} {letter2} {letter3} {index}')
									continue
							log_print(f'{letter1} {letter2} {letter3} Complete\n\n')
							with open(File_path_log_index_LetterA3, 'w', encoding='utf-8') as f3:
								f3.write(letter3)
								f3.flush()
							break
					else:
						log_print(f'{letter1} {letter2} {letter3} Failed\n\n')
						continue
				with open(File_path_log_index_LetterA3, 'w', encoding='utf-8') as f3:
					f3.write('')
					f3.flush()
				# os.remove(File_path_log_index_LetterA3)
				with open(File_path_log_index_LetterA2, 'w', encoding='utf-8') as f2:
					f2.write(letter2)
					f2.flush()
			# os.remove(File_path_log_index_LetterA2)
			with open(File_path_log_index_LetterA2, 'w', encoding='utf-8') as f2:
				f2.write('')
				f2.flush()
			with open(File_path_log_index_LetterA1, 'w', encoding='utf-8') as f1:
				f1.write(letter1)
				f1.flush()
		with open(File_path_log_index_LetterA1, 'w', encoding='utf-8') as f1:
			f1.write('')
			f1.flush()
		# os.remove(File_path_log_index_LetterA1)

		log_index_flag_LetterE1 = False
		if os.path.exists(File_path_log_index_LetterE1):
			with open(File_path_log_index_LetterE1, 'r', encoding='utf-8') as file:
				index_LetterE1 = file.read().strip()
				if index_LetterE1 != '':
					log_index_flag_LetterE1 = True

		if log_index_flag_LetterE1:
			start_index = english_alphabet_list.index(index_LetterE1)
			english_alphabet1 = english_alphabet_list[start_index:]
			log_index_flag_LetterE1 = False
		else:
			english_alphabet1 = english_alphabet_list

		for letter1 in english_alphabet1:
			outerRetry = 1
			error_message_flag = False
			while outerRetry <= retry_attempts:
				if error_message_flag:
					error_message_flag = False
					break
				try:
					# proxy = random.choice(proxies)
					# proxy_url = f'http://{proxy}'
					# log_print(f"Using proxy: {proxy}")
					obj_temp = requests.get(Base_url.format(letter1, 1))
				except (ConnectTimeoutError, RequestException) as e:
					log_print(f"Error occurred for {letter1}")
					exception()
					delay = retry_delay * (2 ** outerRetry)
					log_print(f'Retrying in {delay} seconds...RETRY: {outerRetry}')
					time.sleep(delay)
					outerRetry += 1
					continue
				else:
					soup_temp = BeautifulSoup(obj_temp.content, 'html.parser')

					error_message = soup_temp.find('div', class_='alert alert-danger')
					if error_message:
						error_message_flag = True
					
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
								obj = requests.get(Base_url.format(letter1, index), timeout=300)
							except (ConnectTimeoutError, RequestException) as e:
								log_print(f"Error occurred for {letter1}")
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
								log_print(f'Success {letter1} {index}')
								break
						else:
							log_print(f'Failed {letter1} {index}')
							continue
					log_print(f'{letter1} Complete\n\n')
					with open(File_path_log_index_LetterE1, 'w', encoding='utf-8') as f4:
						f4.write(letter1)
						f4.flush()
					break
			else:
				log_print(f'{letter1} Failed\n\n')
				continue
		with open(File_path_log_index_LetterE1, 'w', encoding='utf-8') as f4:
			f4.write('')
			f4.flush()
	except:
		exception()
	finally:
		book1.close()
		workbook_error.close()
		data = pd.read_excel(File_path)
		data_file = data.drop_duplicates()
		data_file.to_excel(File_path, index=False)
		log_print("\nComplete")	# Restore the default print function
		exit()
	
database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
	delete_task(conn, File_path)
	delete_task(conn, File_path_txt)
	delete_task(conn, File_path_count)
	delete_task(conn, File_path_error)