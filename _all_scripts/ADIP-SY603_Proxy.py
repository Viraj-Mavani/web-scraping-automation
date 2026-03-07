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
File_path_Input = BasePath + '\Proxy\http_proxies.xlsx'

persian_alphabet = [
    "ا", "ب","پ","ت","ث","ج","چ","ح","خ","د","ذ","ر","ز","ژ","س","ش",
	"ص","ض","ط","ظ","ع","غ","ف","ق","ک","گ","ل","م","ن","و","ه","ی"
]

english_alphabet = list(string.ascii_lowercase)

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
					break
				except:
					if try_count>5:
						break
					try_count+=1

			with open(File_path_txt,'a',encoding="utf-8") as fw:
				fw.write("\t".join(map(str,Indi_data))+"\n")
		except:
			exception()

if __name__=='__main__':
	
	row1=1
	rowError=1
	
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
	with open(File_path_count,"w")as f:
		f.write("")

	df = pd.read_excel(File_path_Input, sheet_name='Sheet1')
	proxies = df['proxies'].tolist()
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
	] # List of user-agent headers

	Base_url = 'http://hamachamber.com/members-index/?ftxt={}&ftxt2=&searchType=1&count={}'

	retry_attempts = 4
	retry_delay = 2

	########################## Persian_Alphabet ##########################

	try:
		for letter1 in persian_alphabet:
			for letter2 in persian_alphabet:
				for letter3 in persian_alphabet[:]:
					letter = letter1 + letter2 + letter3
					# letter = letter1 + letter2
					proxy = random.choice(proxies)
					proxy_url = f'http://{proxy}'
					user_agent = random.choice(user_agents)
					headers = {'User-Agent': user_agent}
					print(f"Proxy: {proxy} | User-Agent: {user_agent}")
					for retry in range(retry_attempts):
						try:
							obj_temp = requests.get(Base_url.format(letter, 1), proxies={'http': proxy_url, 'https': proxy_url}, headers=headers)
						except (ConnectTimeoutError, RequestException) as e:
							print(f"Error occurred for {letter}: {e}")
							exception()

							if retry < retry_attempts - 1:
								proxy = random.choice(proxies)
								user_agent = random.choice(user_agents)
								headers = {'User-Agent': user_agent}
								delay = retry_delay * (2 ** retry)
								time.sleep(delay)
								print(f'Retrying in {delay} seconds...{retry}')
								continue
							else:
								raise e
							# obj_temp = requests.get(Base_url.format(letter, 1))
						soup_temp = BeautifulSoup(obj_temp.content, 'html.parser')

						error_message = soup_temp.find('div', class_='alert alert-danger')
						if error_message:
							continue
						
						if len(soup_temp.find_all('a', class_='page-link')) > 0:
							last_page_element = soup_temp.find_all('a', class_='page-link')[-2]
							last_page_number = int(last_page_element.get_text(strip=True))
								
							del obj_temp
							del soup_temp
						else:
							last_page_number = 1

						########################## Individual_data ##########################

						for index in range(1, last_page_number + 1):
							for retry in range(retry_attempts):
								try:
									obj = requests.get(Base_url.format(letter, index), timeout=300)
								except (ConnectTimeoutError, RequestException) as e:
									print(f"Error occurred for {letter}: {e}")
									exception()

									if retry < retry_attempts - 1:
										proxy = random.choice(proxies)
										user_agent = random.choice(user_agents)
										headers = {'User-Agent': user_agent}
										delay = retry_delay * (2 ** retry)
										time.sleep(delay)
										print(f'Retrying in {delay} seconds...{retry}')
										continue
									else:
										raise e
								soup = BeautifulSoup(obj.content, 'html.parser')
								res = soup.find_all('div', class_='mycalls')[1:]
								Individual_data(res)
								print(f'Success {letter} {index}')
								break
						print(f'{letter} Complete\n\n')
						break

		########################## English_Alphabet ##########################

		for letter in english_alphabet:
			# letter = letter1 + letter2
			proxy = random.choice(proxies)
			proxy_url = f'http://{proxy}'
			user_agent = random.choice(user_agents)
			headers = {'User-Agent': user_agent}
			print(f"Proxy: {proxy} | User-Agent: {user_agent}")
			for retry in range(retry_attempts):
				try:
					obj_temp = requests.get(Base_url.format(letter, 1), proxies={'http': proxy, 'https': proxy}, headers=headers)
				except (ConnectTimeoutError, RequestException) as e:
					print(f"Error occurred for {letter}: {e}")
					exception()

					if retry < retry_attempts - 1:
						proxy = random.choice(proxies)
						user_agent = random.choice(user_agents)
						headers = {'User-Agent': user_agent}
						delay = retry_delay * (2 ** retry)
						time.sleep(delay)
						print(f'Retrying in {delay} seconds...{retry}')
						continue
					else:
						raise e
				soup_temp = BeautifulSoup(obj_temp.content, 'html.parser')

				error_message = soup_temp.find('div', class_='alert alert-danger')
				if error_message:
					continue
				
				if len(soup_temp.find_all('a', class_='page-link')) > 0:
					last_page_element = soup_temp.find_all('a', class_='page-link')[-2]
					last_page_number = int(last_page_element.get_text(strip=True))
						
					del obj_temp
					del soup_temp
				else:
					last_page_number = 1

				########################## Individual_data ##########################

				for index in range(1, last_page_number + 1):
					for retry in range(retry_attempts):
						try:
							obj = requests.get(Base_url.format(letter, index), timeout=300, proxies={'http': proxy, 'https': proxy}, headers=headers)
						except (ConnectTimeoutError, RequestException) as e:
							print(f"Error occurred for {letter}: {e}")
							exception()

							if retry < retry_attempts - 1:
								proxy = random.choice(proxies)
								user_agent = random.choice(user_agents)
								headers = {'User-Agent': user_agent}
								delay = retry_delay * (2 ** retry)
								time.sleep(delay)
								print(f'Retrying in {delay} seconds...{retry}')
								continue
							else:
								raise e
						soup = BeautifulSoup(obj.content, 'html.parser')
						res = soup.find_all('div', class_='mycalls')[1:]
						Individual_data(res)
						print(f'Success {letter} {index}')
						break
				print(f'{letter} Complete\n\n')
				break
	except:
		exception()
	finally:
		book1.close()
		workbook_error.close()
		data = pd.read_excel(File_path)
		data_file = data.drop_duplicates()
		data_file.to_excel(File_path, index=False)
		exit()
	
database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
	delete_task(conn, File_path)
	delete_task(conn, File_path_txt)
	delete_task(conn, File_path_count)
	delete_task(conn, File_path_error)