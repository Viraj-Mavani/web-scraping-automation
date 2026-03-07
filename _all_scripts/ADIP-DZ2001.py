import csv
from http.client import RemoteDisconnected
import random
import time
import pandas as pd
import requests
import os
import sys
import sqlite3
import traceback
import re
import string
from sqlite3 import Error
from bs4 import BeautifulSoup
from selenium import webdriver
from requests_toolbelt import MultipartEncoder
from openpyxl.styles import Font
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from requests.exceptions import RequestException
from urllib3.exceptions import ConnectTimeoutError


# BasePath= 'E:\\ADIP-PY\\'
# BasePath = 'D:\\Projects\\CedarPython\\ADIP-DZ2001'
BasePath = os.getcwd()

chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

######### Excel #########
File_path_Personnes_Physique = BasePath + '\\OP\\ADIP-DZ2001_Personnes_Physique.xlsx'
File_path_Personnes_Morales = BasePath + '\\OP\\ADIP-DZ2001_Personnes_Morales.xlsx'
File_path_Personnes_Physique_Arabic = BasePath + '\\OP\\ADIP-DZ2001_Personnes_Physique_Arabic.xlsx'
File_path_Personnes_Morales_Arabic = BasePath + '\\OP\\ADIP-DZ2001_Personnes_Morales_Arabic.xlsx'
######### Text #########
File_path_Personnes_Physique_txt = BasePath + '\\OPtxt\\ADIP-DZ2001_Personnes_Physique.txt'
File_path_Personnes_Morales_txt = BasePath + '\\OPtxt\\ADIP-DZ2001_Personnes_Morales.txt'
File_path_Personnes_Physique_Arabic_txt = BasePath + '\\OPtxt\\ADIP-DZ2001_Personnes_Physique_Arabic.txt'
File_path_Personnes_Morales_Arabic_txt = BasePath + '\\OPtxt\\ADIP-DZ2001_Personnes_Morales_Arabic.txt'
######### Error #########
Error_File = BasePath + '\\Error\\ADIP-DZ2001_Error.xlsx'
######### Input #########
File_path_proxy = BasePath + '\\InputFile\\Proxies.xlsx'
######### Failed #########
File_path_failed_English = BasePath + '\\OP\\ADIP-DZ2001_Failed_English.xlsx'
File_path_failed_Arabic = BasePath + '\\OP\\ADIP-DZ2001_Failed_Arabic.xlsx'
File_path_failed_English_CSV = BasePath + '\\OPcsv\\ADIP-DZ2001_Failed_English.csv'
File_path_failed_Arabic_CSV = BasePath + '\\OPcsv\\ADIP-DZ2001_Failed_Arabic.csv'
######### Count #########
File_path_search_count= BasePath + '\\Counts\\ADIP-DZ2001_Count.txt'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-DZ2001_Log.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-DZ2001_Run_Flag.txt'
File_path_log_index_LetterE1 = BasePath + '\\Log\\ADIP-DZ2001_Log_Index_LetterE1.txt'
File_path_log_index_LetterE2 = BasePath + '\\Log\\ADIP-DZ2001_Log_Index_LetterE2.txt'
File_path_log_index_LetterE3 = BasePath + '\\Log\\ADIP-DZ2001_Log_Index_LetterE3.txt'
File_path_log_index_LetterA1 = BasePath + '\\Log\\ADIP-DZ2001_Log_Index_LetterA1.txt'
File_path_log_index_LetterA2 = BasePath + '\\Log\\ADIP-DZ2001_Log_Index_LetterA2.txt'
File_path_log_index_LetterA3 = BasePath + '\\Log\\ADIP-DZ2001_Log_Index_LetterA3.txt'
######### CSV #########
Error_File_CSV = BasePath + '\\OPcsv\\ADIP-DZ2001_Error.csv'
File_path_Personnes_Physique_CSV = BasePath + '\\OPcsv\\ADIP-DZ2001_Personnes_Physique.csv'
File_path_Personnes_Physique_Arabic_CSV = BasePath + '\\OPcsv\\ADIP-DZ2001_Personnes_Physique_Arabic.csv'
File_path_Personnes_Morales_CSV = BasePath + '\\OPcsv\\ADIP-DZ2001_Personnes_Morales.csv'
File_path_Personnes_Morales_Arabic_CSV = BasePath + '\\OPcsv\\ADIP-DZ2001_Personnes_Morales_Arabic.csv'
# File_path_Personnes_Physique_CSV_wo_Duplicate = BasePath + '\\OPcsv\\ADIP-DZ2001_Personnes_Physique_wo_Duplicate.csv'
# File_path_Personnes_Physique_Arabic_CSV_wo_Duplicate = BasePath + '\\OPcsv\\ADIP-DZ2001_Personnes_Physique_Arabic_wo_Duplicate.csv'
# File_path_Personnes_Morales_CSV_wo_Duplicate = BasePath + '\\OPcsv\\ADIP-DZ2001_Personnes_Morales.csv'
# File_path_Personnes_Morales_Arabic_CSV_wo_Duplicate = BasePath + '\\OPcsv\\ADIP-DZ2001_Personnes_Morales_Arabic_wo_Duplicate.csv'



English_alphabet_list = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K',
            'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

Arabic_alphabet_list = ['ا', 'ب', 'ت', 'ث', 'ج', 'ح', 'خ', 'د', 'ذ', 'ر', 'ز', 'س', 'ش',
                   'ص', 'ض', 'ط', 'ظ', 'ع', 'غ', 'ف', 'ق', 'ك', 'ل', 'م', 'ن', 'ه', 'و', 'ي']


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


def Dereference(obj):
	del obj


def log_print(message):
    with open(File_path_log, 'a', encoding='utf-8') as file:
        file.write(message + '\n')
        file.flush()
    print(message)


def exception(URL):
    # global rowError
    Headers_Error = ['URL', 'Not Responding', 'Error']
    error = traceback.format_exc()
    exception_type, exception_object, exception_traceback = sys.exc_info()
    with open(Error_File_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(Headers_Error)
        writer.writerow([URL, "Not Responding", str(error)])
    df = pd.read_csv(Error_File_CSV, encoding='utf-8')
    df.to_excel(Error_File, index=False)


def duplicateFromCSV(Csv_File_path):
    try:
        data = pd.read_csv(Csv_File_path)
        unique_data = data.drop_duplicates()
        unique_data.to_csv(Csv_File_path, index=False)
    except:
        pass


def convertCSVExcel(File_path_CSV, File_path_EXL):
    chunk_size = 1000000  # Number of rows per Excel sheet (adjust as needed)
    csv_reader = pd.read_csv(File_path_CSV, encoding='utf-8', chunksize=chunk_size)
    sheet_index = 1  # Index of the Excel sheet
    excel_files = []  # List to store the names of generated Excel files

    for chunk in csv_reader:
        if len(chunk) > 0:  # Create Excel sheet only if chunk is not empty
            sheet_name = f'DataSet {sheet_index}'  # Generate a unique sheet name
            excel_file = f'{File_path_EXL[:-5]}_{sheet_index}.xlsx'  # Generate a unique Excel file name
            chunk.to_excel(excel_file, sheet_name=sheet_name, index=False)
            excel_files.append(excel_file)
            sheet_index += 1

    # Merge all Excel files into one
    writer = pd.ExcelWriter(File_path_EXL, engine='xlsxwriter')
    Sheet = 1
    for file in excel_files:
        df = pd.read_excel(file)
        sheet_name = f'DataSet {Sheet}'
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        Sheet += 1
    writer.close()


# def duplicate(File_path):
#     try:
#         data = pd.read_excel(File_path)
#         data_file = data.drop_duplicates()
#         data_file.to_excel(File_path, index=False)
#     except:
#         pass


def count():
    try_count = 1
    while try_count <= 5:
        try:
            with open(File_path_search_count, 'a', encoding='utf-8') as fh:
                fh.write('1\n')
                fh.flush()
            break
        except Exception:
            pass
        try_count += 1


if __name__=='__main__':
	File_paths= [File_path_Personnes_Physique,File_path_Personnes_Morales, File_path_Personnes_Physique_Arabic, File_path_Personnes_Morales_Arabic]
    
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
	
	# First_run = False
    # if First_run:
	if not os.path.exists(File_path_log_Run_Flag):
		with open(File_path_log_Run_Flag, "a", encoding='utf-8')as f:
			f.write("")
		File_paths_csv = [File_path_Personnes_Physique_CSV, File_path_Personnes_Morales_CSV,
					File_path_Personnes_Physique_Arabic_CSV, File_path_Personnes_Morales_Arabic_CSV, File_path_failed_English_CSV, File_path_failed_Arabic_CSV]
		File_paths_txt = [File_path_Personnes_Physique_txt, File_path_Personnes_Morales_txt,
					File_path_Personnes_Physique_Arabic_txt, File_path_Personnes_Morales_Arabic_txt]
		File_path_index = [File_path_log_index_LetterE1, File_path_log_index_LetterE2, File_path_log_index_LetterE3, 
                    File_path_log_index_LetterA1, File_path_log_index_LetterA2, File_path_log_index_LetterA3]
		if os.path.exists(File_path_search_count):
			os.remove(File_path_search_count)
		if os.path.exists(File_path_log):
			os.remove(File_path_log)
		for path_csv in File_paths_csv:
			if os.path.exists(path_csv):
				os.remove(path_csv)
		for Path_txt in File_paths_txt:
			if os.path.exists(Path_txt):
				os.remove(Path_txt)
		for Path_index in File_path_index:
			if os.path.exists(Path_index):
				os.remove(Path_index)
	
	Personnes_Physique_headers = ['NRC','Nom','Prenom']
	Personnes_Physique_Arabic_headers = ['NRC','Nom (Arabic)','Prenom (Arabic)']
	Personnes_Morales_headers = ['NRC','Raison Sociale']
	Personnes_Morales_Arabic_headers = ['NRC','Raison Sociale (Arabic)']
	
	if not os.path.exists(File_path_search_count):
		with open(File_path_search_count, "a", encoding='utf-8')as f:
			f.write("")
	with open(File_path_Personnes_Physique_txt, "a", encoding='utf-8')as f:
		if f.tell() == 0:
			f.write("\t".join(Personnes_Physique_headers)+"\n")
			f.flush()
	with open(File_path_Personnes_Morales_txt,"a", encoding='utf-8')as fw:
		if fw.tell() == 0:
			fw.write("\t".join(Personnes_Morales_headers)+"\n")
			fw.flush()
	with open(File_path_Personnes_Physique_Arabic_txt, "a", encoding='utf-8')as f:
		if f.tell() == 0:
			f.write("\t".join(Personnes_Physique_Arabic_headers)+"\n")
			f.flush()
	with open(File_path_Personnes_Morales_Arabic_txt, "a", encoding='utf-8')as fw:
		if fw.tell() == 0:
			fw.write("\t".join(Personnes_Morales_Arabic_headers)+"\n")
			fw.flush()

	with open(File_path_Personnes_Physique_CSV, "a", newline='', encoding='utf-8') as f:
		writer = csv.writer(f)
		if f.tell() == 0:
			writer.writerow(Personnes_Physique_headers)
	with open(File_path_Personnes_Morales_CSV, "a", newline='', encoding='utf-8') as fw:
		writer = csv.writer(fw)
		if fw.tell() == 0:
			writer.writerow(Personnes_Morales_headers)
	with open(File_path_Personnes_Physique_Arabic_CSV, "a", newline='', encoding='utf-8') as f:
		writer = csv.writer(f)
		if f.tell() == 0:
			writer.writerow(Personnes_Physique_Arabic_headers)
	with open(File_path_Personnes_Morales_Arabic_CSV, "a", newline='', encoding='utf-8') as fw:
		writer = csv.writer(fw)
		if fw.tell() == 0:
			writer.writerow(Personnes_Morales_Arabic_headers)
	
	# df = pd.read_csv(File_path_proxy, header=None)
	# proxy_list = df.iloc[:, 0].tolist()
	df = pd.read_excel(File_path_proxy, sheet_name='Sheet1')
	proxy_list = df['proxies'].tolist()
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
	proxy = random.choice(proxy_list)

	Home_URL = 'https://sidjilcom.cnrc.dz/web/cnrc/accueil'
	Arabic_url = 'https://sidjilcom.cnrc.dz/accueil?p_p_id=82&p_p_lifecycle=1&p_p_state=normal&p_p_mode=view&p_p_col_id=column-5&p_p_col_count=8&_82_struts_action=%2Flanguage%2Fview&_82_redirect=%2Faccueil&_82_languageId=ar_SA'
	
	retry_attempts = 5
	retry_delay = 2

	try:
		Driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=chrome_options)
		Driver.get(Home_URL)
		
		sid = Driver.get_cookie('SID')['name'] + '=' + Driver.get_cookie('SID')['value'] + ';'
		gid = Driver.get_cookie('_gid')['name'] + '=' + Driver.get_cookie('_gid')['value'] + ';'
		ga = Driver.get_cookie('_ga')['name'] + '=' + Driver.get_cookie('_ga')['value'] + ';'
		session = Driver.get_cookie('cookiesession1')['name'] + '=' + Driver.get_cookie('cookiesession1')['value'] + ';'
		support = Driver.get_cookie('COOKIE_SUPPORT')['name'] + '=' + Driver.get_cookie('COOKIE_SUPPORT')['value'] + ';'
		lang = Driver.get_cookie('GUEST_LANGUAGE_ID')['name'] + '=' + Driver.get_cookie('GUEST_LANGUAGE_ID')['value'] + ';'
		Jsess = Driver.get_cookie('JSESSIONID')['name'] + '=' + Driver.get_cookie('JSESSIONID')['value'] + ';'
		
		Home_soup = BeautifulSoup(Driver.page_source.encode(), 'html.parser')
		Link_Form = Home_soup.find('form',id='f1').get('action')
		Driver.close()
		Driver.quit()

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
					letter = letterE1 + letterE2 + letterE3
					js_found_flag = False
					fields = {
						'hidden': 'goRecherche',
						'critere': letter
						}
					Form_Data = MultipartEncoder(fields=fields,boundary='----WebKitFormBoundaryTdS000SBKNSpDEkf')
					Headers = {'Cookie':support + ' ' + lang + ' ' + session + ' ' + ga + ' ' + gid + ' ' + Jsess + ' ' + sid,
							'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
							'Accept-Encoding':'gzip, deflate, br',
							'Accept-Language':'en-IN,en-GB;q=0.9,en-US;q=0.8,en;q=0.7',
							'Cache-Control':'max-age=0',
							'Connection':'keep-alive',
							'Content-Length':'243',
							'Content-Type':'multipart/form-data; boundary=----WebKitFormBoundaryTdS000SBKNSpDEkf',
							'Host':'sidjilcom.cnrc.dz',
							'Origin':'https://sidjilcom.cnrc.dz',
							'Referer':Link_Form,
							'sec-ch-ua':'\"Google Chrome\";v=\"111\", \"Not(A:Brand\";v=\"8\", \"Chromium\";v=\"111\"',
							'sec-ch-ua-mobile':'?0',
							'sec-ch-ua-platform':'\"Windows\"',
							'Sec-Fetch-Dest':'document',
							'Sec-Fetch-Mode':'navigate',
							'Sec-Fetch-Site':'same-origin',
							'Sec-Fetch-User':'?1',
							'Upgrade-Insecure-Requests':'1',
							# 'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36'}
							'User-Agent': user_agent}
					try:
						engRetry = 1
						while engRetry <= retry_attempts:
							try:
								# Info = requests.post(Link_Form, data=Form_Data, headers=Headers, proxies={'http': proxy, 'https': proxy}, timeout=200)
								Info = requests.post(Link_Form, data=Form_Data, headers=Headers, timeout=200)
								break
							# except Exception as e:
							# 	log_print(f"Error occurred for {letter}")
							# 	exception(Home_URL)
							# 	log_print(str(e))
							# 	os._exit(1)
							except Exception as e:
								log_print(f"Error occurred for {letter}...Retrying in 2 min")
								time.sleep(120)
								try:
									Driver.get(Home_URL)
								except Exception as e:
									log_print(str(e))
									exception(Home_URL)
									os._exit(1)

								sid = Driver.get_cookie('SID')['name'] + '=' + \
											Driver.get_cookie('SID')['value'] + ';'
								gid = Driver.get_cookie('_gid')['name'] + '=' + \
											Driver.get_cookie('_gid')['value'] + ';'
								ga = Driver.get_cookie('_ga')['name'] + '=' + \
											Driver.get_cookie('_ga')['value'] + ';'
								session = Driver.get_cookie('cookiesession1')[
											'name'] + '=' + Driver.get_cookie('cookiesession1')['value'] + ';'
								support = Driver.get_cookie('COOKIE_SUPPORT')[
											'name'] + '=' + Driver.get_cookie('COOKIE_SUPPORT')['value'] + ';'
								lang = Driver.get_cookie('GUEST_LANGUAGE_ID')[
											'name'] + '=' + Driver.get_cookie('GUEST_LANGUAGE_ID')['value'] + ';'
								Jsess = Driver.get_cookie('JSESSIONID')[
											'name'] + '=' + Driver.get_cookie('JSESSIONID')['value'] + ';'

								Home_soup = BeautifulSoup(Driver.page_source.encode(), 'html.parser')
								Link_Form = Home_soup.find('form', id='f1').get('action')
								Driver.close()
								Driver.quit()
								fields = {'hidden': 'goRecherche', 'critere': letter}
								Form_Data = MultipartEncoder(
								fields=fields, boundary='----WebKitFormBoundaryTdS000SBKNSpDEkf')
								proxy = random.choice(proxy_list)
								user_agent = random.choice(user_agents)
								log_print(f'Using {proxy} & {user_agent}')
								Headers = {'Cookie': support + ' ' + lang + ' ' + session + ' ' + ga + ' ' + gid + ' ' + Jsess + ' ' + sid,
										'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
										'Accept-Encoding': 'gzip, deflate, br',
										'Accept-Language': 'en-IN,en-GB;q=0.9,en-US;q=0.8,en;q=0.7',
										'Cache-Control': 'max-age=0',
										'Connection': 'keep-alive',
										'Content-Length': '243',
										'Content-Type': 'multipart/form-data; boundary=----WebKitFormBoundaryTdS000SBKNSpDEkf',
										'Host': 'sidjilcom.cnrc.dz',
										'Origin': 'https://sidjilcom.cnrc.dz',
										'Referer': Link_Form,
										'sec-ch-ua': '\"Google Chrome\";v=\"111\", \"Not(A:Brand\";v=\"8\", \"Chromium\";v=\"111\"',
										'sec-ch-ua-mobile': '?0',
										'sec-ch-ua-platform': '\"Windows\"',
										'Sec-Fetch-Dest': 'document',
										'Sec-Fetch-Mode': 'navigate',
										'Sec-Fetch-Site': 'same-origin',
										'Sec-Fetch-User': '?1',
										'Upgrade-Insecure-Requests': '1',
										'User-Agent': user_agent}
								continue
						else:
							log_print(str(e))
							exception(Home_URL)
							os._exit(1)
						Data_soup = BeautifulSoup(Info.content,'html.parser')

						pattern = r'\$\(\s*function\(\)\s*{\s*\$\.\w+\({'

						if len(Data_soup.find_all('script',type='text/javascript'))>1:
							for scripts in Data_soup.find_all('script',type='text/javascript'):
								if scripts.string == None:
									continue
								if re.match(pattern, scripts.string):
									Script = scripts.string
									if Script == None:
										continue
									Script_soup = BeautifulSoup(Script,'html.parser')
									js_found_flag = True
									Data = Script_soup.find_all('tr')
									Personnes_Physique_count = int(Script_soup.find_all('b')[2].string)+1
									i=1
									for items in Data:
										if i<Personnes_Physique_count:
											if items.find_all('td'):
												for j in items.find_all('td')[1].attrs:
													if j == 'style' and items.find_all('td')[1].attrs['style'] == 'text-align:left':
														# log_print('Record '+ str(i) + ' Added')
														count()
														Personnes_Physique_data = []
														Personnes_Physique_data.append(items.find_all('td')[0].string.strip() if items.find_all('td')[0].string else '')
														Personnes_Physique_data.append(items.find_all('td')[1].string.strip() if items.find_all('td')[1].string else '')
														Personnes_Physique_data.append(items.find_all('td')[2].string.strip() if items.find_all('td')[2].string else '')
														with open(File_path_Personnes_Physique_CSV, 'a', newline='', encoding='utf-8') as file:
															writer = csv.writer(file)
															writer.writerow(Personnes_Physique_data)
														with open(File_path_Personnes_Physique_txt, "a", encoding='utf-8')as f:
															f.write("\t".join(map(str,Personnes_Physique_data))+"\n")
															f.flush()
														int(i)
														i+=1
										else:
											if items.find_all('td'):
												for j in items.find_all('td')[1].attrs:
													if j == 'style' and items.find_all('td')[1].attrs['style'] == 'text-align:left':
														# log_print('Record ' + str(i) + ' Added')
														count()
														Personnes_Morales_data = []
														Personnes_Morales_data.append(items.find_all('td')[0].string.strip() if items.find_all('td')[0].string else '')
														Personnes_Morales_data.append(items.find_all('td')[1].string.strip() if items.find_all('td')[1].string else '')
														with open(File_path_Personnes_Morales_CSV, 'a', newline='', encoding='utf-8') as file:
															writer = csv.writer(file)
															writer.writerow(Personnes_Morales_data)
														with open(File_path_Personnes_Morales_txt, "a", encoding='utf-8')as f:
															f.write("\t".join(map(str,Personnes_Morales_data))+"\n")
															f.flush()
														int(i)
														i+=1
							if not js_found_flag:
								log_print(f"Failed!! Couldn't find JS for {letter}")
								user_agent = random.choice(user_agents)
								with open(File_path_log_index_LetterE3, 'w', encoding='utf-8') as file:
									file.write(letterE3)
									file.flush()
								with open(File_path_failed_English_CSV, 'a', newline='', encoding='utf-8') as file:
									writer = csv.writer(file)
									writer.writerow([letter])
								continue
								# os._exit(1)
						log_print('Complete ' + letter)
						with open(File_path_log_index_LetterE3, 'w', encoding='utf-8') as file:
							file.write(letterE3)
							file.flush()
					except:
						exception(Home_URL)
					finally:
						Dereference(Info)
						try:
							Data_soup.decompose()
							Script_soup.decompose()
						except:
							pass
						# if not js_found_flag:
						# 	exit(1)
				with open(File_path_log_index_LetterE2, 'w', encoding='utf-8') as file:
					file.write(letterE2)
					file.flush()
				start_index_LetterE3 = 0
			with open(File_path_log_index_LetterE1, 'w', encoding='utf-8') as f1:
				f1.write(letterE1)
				f1.flush()
			start_index_LetterE2 = 0

	except:
		exception(Home_URL)
	finally:
		Home_soup.decompose()
		# duplicateFromCSV(File_path_Personnes_Physique_CSV)
		# duplicateFromCSV(File_path_Personnes_Morales_CSV)
		# convertCSVExcel(File_path_Personnes_Physique_CSV, File_path_Personnes_Physique)
		# convertCSVExcel(File_path_Personnes_Morales_CSV, File_path_Personnes_Morales)
		# if os.path.exists(File_path_failed_English_CSV):
		# 	df = pd.read_csv(File_path_failed_English_CSV, encoding='utf-8', low_memory=False)
		# 	df.to_excel(File_path_failed_English, index=False)
		# # duplicate(File_path_Personnes_Physique)
		# # duplicate(File_path_Personnes_Morales)

	try:
		Arabic_Driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=chrome_options)
		Arabic_Driver.get(Arabic_url)
		Arabic_url = Arabic_Driver.current_url
		Arabic_Driver.get(Arabic_url)

		Arabic_sid = Arabic_Driver.get_cookie('SID')['name'] + '=' + Arabic_Driver.get_cookie('SID')['value'] + ';'
		Arabic_gid = Arabic_Driver.get_cookie('_gid')['name'] + '=' + Arabic_Driver.get_cookie('_gid')['value'] + ';'
		Arabic_ga = Arabic_Driver.get_cookie('_ga')['name'] + '=' + Arabic_Driver.get_cookie('_ga')['value'] + ';'
		Arabic_session = Arabic_Driver.get_cookie('cookiesession1')['name'] + '=' + Arabic_Driver.get_cookie('cookiesession1')['value'] + ';'
		Arabic_support = Arabic_Driver.get_cookie('COOKIE_SUPPORT')['name'] + '=' + Arabic_Driver.get_cookie('COOKIE_SUPPORT')['value'] + ';'
		Arabic_lang = Arabic_Driver.get_cookie('GUEST_LANGUAGE_ID')['name'] + '=' + Arabic_Driver.get_cookie('GUEST_LANGUAGE_ID')['value'] + ';'
		Arabic_Jsess = Arabic_Driver.get_cookie('JSESSIONID')['name'] + '=' + Arabic_Driver.get_cookie('JSESSIONID')['value'] + ';'
		
		Ar_Home_soup = BeautifulSoup(Arabic_Driver.page_source.encode(), 'html.parser')
		Ar_Link_Form = Ar_Home_soup.find('form',id='f1').get('action')
		Arabic_Driver.close()
		Arabic_Driver.quit()

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
				if index_LetterA2 != '' and index_LetterA2 != 'ي':
					log_index_flag_LetterA2 = True
     
		if os.path.exists(File_path_log_index_LetterA3):
			with open(File_path_log_index_LetterA3, 'r', encoding='utf-8') as file:
				index_LetterA3 = file.read().strip()
				if index_LetterA3 != '' and index_LetterA3 != 'ي':
					log_index_flag_LetterA3 = True

		if log_index_flag_LetterA1:
			start_index_LetterA1 = Arabic_alphabet_list.index(index_LetterA1) + 1
		else:
			start_index_LetterA1 = 0

		if log_index_flag_LetterA2:
			start_index_LetterA2 = Arabic_alphabet_list.index(index_LetterA2) + 1
		else:
			start_index_LetterA2 = 0
   
		if log_index_flag_LetterA3:
			start_index_LetterA3 = Arabic_alphabet_list.index(index_LetterA3) + 1
		else:
			start_index_LetterA3 = 0


		for indexA1 in range(start_index_LetterA1, len(Arabic_alphabet_list)):
			letterA1 = Arabic_alphabet_list[indexA1]
			for indexA2 in range(start_index_LetterA2, len(Arabic_alphabet_list)):
				letterA2 = Arabic_alphabet_list[indexA2]
				for indexA3 in range(start_index_LetterA3, len(Arabic_alphabet_list)):
					letterA3 = Arabic_alphabet_list[indexA3]
					letterA = letterA1 + letterA2 + letterA3
					js_found_flag = False
					fields = {
						'hidden': 'goRecherche',
						'critere': letterA
						}
					Form_Data = MultipartEncoder(fields=fields,boundary='----WebKitFormBoundaryTdS000SBKNSpDEkf')
					Headers = {'Cookie':Arabic_support + ' ' + Arabic_lang + ' ' + Arabic_session + ' ' + Arabic_ga + ' ' + Arabic_gid + ' ' + Arabic_Jsess + ' ' + Arabic_sid,
							'Accept':'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
							'Accept-Encoding':'gzip, deflate, br',
							'Accept-Language':'en-IN,en-GB;q=0.9,en-US;q=0.8,en;q=0.7',
							'Cache-Control':'max-age=0',
							'Connection':'keep-alive',
							'Content-Length':'243',
							'Content-Type':'multipart/form-data; boundary=----WebKitFormBoundaryTdS000SBKNSpDEkf',
							'Host':'sidjilcom.cnrc.dz',
							'Origin':'https://sidjilcom.cnrc.dz',
							'Referer':Ar_Link_Form,
							'sec-ch-ua':'\"Google Chrome\";v=\"111\", \"Not(A:Brand\";v=\"8\", \"Chromium\";v=\"111\"',
							'sec-ch-ua-mobile':'?0',
							'sec-ch-ua-platform':'\"Windows\"',
							'Sec-Fetch-Dest':'document',
							'Sec-Fetch-Mode':'navigate',
							'Sec-Fetch-Site':'same-origin',
							'Sec-Fetch-User':'?1',
							'Upgrade-Insecure-Requests':'1',
							'User-Agent': random.choice(user_agents)}
					try:
						areRetry = 1
						while areRetry <= retry_attempts:
							try:
								Ar_Info = requests.post(Ar_Link_Form, data=Form_Data, headers=Headers, timeout=200)
								# Ar_Info = requests.post(Ar_Link_Form, data=Form_Data, headers=Headers, proxies={'http': proxy, 'https': proxy}, timeout=200)
								break
							# except Exception as e:
							# 	log_print(f"Error occurred for {letterA}")
							# 	exception(Arabic_url)
							# 	log_print(str(e))
							# 	os._exit(1)
							except Exception as e:
								log_print(f"Error occurred for {letterA}...Retrying in 2 min")
								time.sleep(120)
								try: 
									Arabic_Driver.get(Arabic_url)
									Arabic_url = Arabic_Driver.current_url
									Arabic_Driver.get(Arabic_url)
								except Exception as e:
									exception(Arabic_url)
									log_print(str(e))
									os._exit(1)

								Arabic_sid = Arabic_Driver.get_cookie('SID')['name'] + '=' + Arabic_Driver.get_cookie('SID')['value'] + ';'
								Arabic_gid = Arabic_Driver.get_cookie('_gid')['name'] + '=' + Arabic_Driver.get_cookie('_gid')['value'] + ';'
								Arabic_ga = Arabic_Driver.get_cookie('_ga')['name'] + '=' + Arabic_Driver.get_cookie('_ga')['value'] + ';'
								Arabic_session = Arabic_Driver.get_cookie('cookiesession1')['name'] + '=' + Arabic_Driver.get_cookie('cookiesession1')['value'] + ';'
								Arabic_support = Arabic_Driver.get_cookie('COOKIE_SUPPORT')['name'] + '=' + Arabic_Driver.get_cookie('COOKIE_SUPPORT')['value'] + ';'
								Arabic_lang = Arabic_Driver.get_cookie('GUEST_LANGUAGE_ID')['name'] + '=' + Arabic_Driver.get_cookie('GUEST_LANGUAGE_ID')['value'] + ';'
								Arabic_Jsess = Arabic_Driver.get_cookie('JSESSIONID')['name'] + '=' + Arabic_Driver.get_cookie('JSESSIONID')['value'] + ';'
								
								Ar_Home_soup = BeautifulSoup(Arabic_Driver.page_source.encode(), 'html.parser')
								Ar_Link_Form = Ar_Home_soup.find('form',id='f1').get('action')
								Arabic_Driver.close()
								Arabic_Driver.quit()
								fields = {'hidden': 'goRecherche', 'critere': letterA}
								Form_Data = MultipartEncoder(
								fields=fields, boundary='----WebKitFormBoundaryTdS000SBKNSpDEkf')
								# proxy = random.choice(proxy_list)
								user_agent = random.choice(user_agents)
								log_print(f'Using {proxy} & {user_agent}')
								Headers = {'Cookie': Arabic_support + ' ' + Arabic_lang + ' ' + Arabic_session + ' ' + Arabic_ga + ' ' + Arabic_gid + ' ' + Arabic_Jsess + ' ' + Arabic_sid,
										'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
										'Accept-Encoding': 'gzip, deflate, br',
										'Accept-Language': 'en-IN,en-GB;q=0.9,en-US;q=0.8,en;q=0.7',
										'Cache-Control': 'max-age=0',
										'Connection': 'keep-alive',
										'Content-Length': '243',
										'Content-Type': 'multipart/form-data; boundary=----WebKitFormBoundaryTdS000SBKNSpDEkf',
										'Host': 'sidjilcom.cnrc.dz',
										'Origin': 'https://sidjilcom.cnrc.dz',
										'Referer': Ar_Link_Form,
										'sec-ch-ua': '\"Google Chrome\";v=\"111\", \"Not(A:Brand\";v=\"8\", \"Chromium\";v=\"111\"',
										'sec-ch-ua-mobile': '?0',
										'sec-ch-ua-platform': '\"Windows\"',
										'Sec-Fetch-Dest': 'document',
										'Sec-Fetch-Mode': 'navigate',
										'Sec-Fetch-Site': 'same-origin',
										'Sec-Fetch-User': '?1',
										'Upgrade-Insecure-Requests': '1',
										'User-Agent': user_agent}
								continue
						else:
							exception(Arabic_url)
							log_print(str(e))
							os._exit(1)
							# except Exception as e:
							# 	if isinstance(e, RemoteDisconnected):
							# 		log_print(f"ConnectionError occurred for {letter}")
							# 	else:
							# 		log_print(f"Error occurred for {letter}")
							# 	delay = retry_delay * (2 * areRetry)
							# 	log_print(f'Retrying in {delay} seconds...{areRetry}')
							# 	time.sleep(delay)
							# 	areRetry += 1
							# 	continue
						Ar_Data_soup = BeautifulSoup(Ar_Info.content,'html.parser')

						pattern = r'\$\(\s*function\(\)\s*{\s*\$\.\w+\({'

						if len(Ar_Data_soup.find_all('script',type='text/javascript'))>1:
							for scripts in Ar_Data_soup.find_all('script',type='text/javascript'):
								if scripts.string == None:
									continue
								if re.match(pattern, scripts.string):
									Ar_Script = scripts.string
									if Ar_Script == None:
										continue
									Ar_Script_soup = BeautifulSoup(Ar_Script,'html.parser')
									js_found_flag = True
									Ar_Data = Ar_Script_soup.find_all('tr')
									Personnes_Physique_count_ar = int(Ar_Script_soup.find_all('b')[2].string)+1
									i=1
									for items in Ar_Data:
										if i<Personnes_Physique_count_ar:
											if items.find_all('td'):
												# log_print('Record ' + str(i) + ' Added')
												try_count=1
												count()
												Personnes_Physique_data = []
												Personnes_Physique_data.append(items.find_all('td')[0].string.strip() if items.find_all('td')[0].string else '')
												Personnes_Physique_data.append(items.find_all('td')[1].string.strip() if items.find_all('td')[1].string else '')
												Personnes_Physique_data.append(items.find_all('td')[2].string.strip() if items.find_all('td')[2].string else '')
												with open(File_path_Personnes_Physique_Arabic_CSV, 'a', newline='', encoding='utf-8') as file:
													writer = csv.writer(file)
													writer.writerow(Personnes_Physique_data)
												with open(File_path_Personnes_Physique_Arabic_txt,"a",encoding='utf-8')as f:
													f.write("\t".join(map(str,Personnes_Physique_data))+"\n")
													f.flush()
												int(i)
												i+=1
										else:
											if items.find_all('td'):
												# log_print('Record ' + str(i) + ' Added')
												count()
												Personnes_Morales_data = []
												Personnes_Morales_data.append(items.find_all('td')[0].string.strip() if items.find_all('td')[0].string else '')
												Personnes_Morales_data.append(items.find_all('td')[1].string.strip() if items.find_all('td')[1].string else '')
												with open(File_path_Personnes_Morales_Arabic_CSV, 'a', newline='', encoding='utf-8') as file:
													writer = csv.writer(file)
													writer.writerow(Personnes_Morales_data)
												with open(File_path_Personnes_Morales_Arabic_txt,"a",encoding='utf-8')as f:
													f.write("\t".join(map(str,Personnes_Morales_data))+"\n")
													f.flush()
												int(i)
												i+=1
							if not js_found_flag:
								log_print(f"Failed!! Couldn't find JS for {letterA}")
								with open(File_path_log_index_LetterA3, 'w', encoding='utf-8') as file:
									file.write(letterA3)
									file.flush()
								with open(File_path_failed_Arabic_CSV, 'a', newline='', encoding='utf-8') as file:
									writer = csv.writer(file)
									writer.writerow([letterA])
								continue
								# os._exit(1)
						log_print('Complete ' + letterA)
						with open(File_path_log_index_LetterA3, 'w', encoding='utf-8') as file:
							file.write(letterA3)
							file.flush()
					except:
						exception(Arabic_url)
					finally:
						Dereference(Ar_Info)
						try:
							Ar_Data_soup.decompose()
							Ar_Script_soup.decompose()
						except:
							pass
						# if not js_found_flag:
						# 	exit(1)
				with open(File_path_log_index_LetterA2, 'w', encoding='utf-8') as file:
					file.write(letterA2)
					file.flush()
				start_index_LetterA3 = 0
			with open(File_path_log_index_LetterA1, 'w', encoding='utf-8') as f1:
				f1.write(letterA1)
				f1.flush()
			start_index_LetterA2 = 0
	except:
		exception(Arabic_url)
	finally:
		Ar_Home_soup.decompose()
		duplicateFromCSV(File_path_Personnes_Physique_Arabic_CSV)
		duplicateFromCSV(File_path_Personnes_Morales_Arabic_CSV)
		convertCSVExcel(File_path_Personnes_Physique_Arabic_CSV, File_path_Personnes_Physique_Arabic)
		convertCSVExcel(File_path_Personnes_Morales_Arabic_CSV, File_path_Personnes_Morales_Arabic)
		if os.path.exists(File_path_failed_Arabic_CSV):
			df = pd.read_csv(File_path_failed_Arabic_CSV, encoding='utf-8', low_memory=False)
			df.to_excel(File_path_failed_Arabic, index=False)
	file_pathsE = [File_path_log_index_LetterE1, File_path_log_index_LetterE2, File_path_log_index_LetterE3]
	if all(os.path.exists(file_path) for file_path in file_pathsE):
		letters = []
		for file_path in file_pathsE:
			with open(file_path, 'r', encoding='utf-8') as file:
				letter = file.read().strip()
				letters.append(letter)
		if all(letter == 'Z' for letter in letters):
			log_print('English Letters Complete')
	file_pathsA = [File_path_log_index_LetterA1, File_path_log_index_LetterA2, File_path_log_index_LetterA3]
	if all(os.path.exists(file_path) for file_path in file_pathsA):
		letters = []
		for file_path in file_pathsA:
			with open(file_path, 'r', encoding='utf-8') as file:
				letter = file.read().strip()
				letters.append(letter)
		if all(letter == 'ي' for letter in letters):
			log_print('Arabic Letters Complete')
	log_print('Script Complete')
	exit()
  
database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
	for File_path in File_paths:
		delete_task(conn, File_path)