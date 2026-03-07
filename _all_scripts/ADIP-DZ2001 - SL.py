import csv
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
# import openpyxl
from requests_toolbelt import MultipartEncoder
from openpyxl.styles import Font
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options


# BasePath= 'E:\\ADIP-PY\\'
BasePath = 'D:\\Projects\\CedarPython\\ADIP-DZ2001'

chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

######### Excel #########
File_path_Personnes_Physique = BasePath + '\\OP\\ADIP_DZ2001_Personnes_Physique.xlsx'
File_path_Personnes_Morales = BasePath + '\\OP\\ADIP_DZ2001_Personnes_Morales.xlsx'
File_path_Personnes_Physique_Arabic = BasePath + '\\OP\\ADIP_DZ2001_Personnes_Physique_Arabic.xlsx'
File_path_Personnes_Morales_Arabic = BasePath + '\\OP\\ADIP_DZ2001_Personnes_Morales_Arabic.xlsx'
######### Text #########
File_path_Personnes_Physique_txt = BasePath + '\\OPtxt\\ADIP_DZ2001_Personnes_Physique.txt'
File_path_Personnes_Morales_txt = BasePath + '\\OPtxt\\ADIP_DZ2001_Personnes_Morales.txt'
File_path_Personnes_Physique_Arabic_txt = BasePath + '\\OPtxt\\ADIP_DZ2001_Personnes_Physique_Arabic.txt'
File_path_Personnes_Morales_Arabic_txt = BasePath + '\\OPtxt\\ADIP_DZ2001_Personnes_Morales_Arabic.txt'
######### Error #########
Error_File = BasePath + '\\Error\\ADIP-DZ2001_Error.xlsx'
Error_File_CSV = BasePath + '\\Error\\ADIP-DZ2001_Error.csv'
######### Count #########
File_path_search_count= BasePath + '\\Counts\\ADIP-DZ2001_Count.txt'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-DZ2001_Log.txt'
File_path_log_index_English = BasePath + '\\Log\\ADIP-DZ2001_Log_Index_English.txt'
File_path_log_index_Arabic = BasePath + '\\Log\\ADIP-DZ2001_Log_Index_Arabic.txt'
######### CSV #########
File_path_Personnes_Physique_CSV = BasePath + '\\OPcsv\\ADIP_DZ2001_Personnes_Physique.csv'
File_path_Personnes_Physique_Arabic_CSV = BasePath + '\\OPcsv\\ADIP_DZ2001_Personnes_Physique_Arabic.csv'
File_path_Personnes_Morales_CSV = BasePath + '\\OPcsv\\ADIP_DZ2001_Personnes_Morales.csv'
File_path_Personnes_Morales_Arabic_CSV = BasePath + '\\OPcsv\\ADIP_DZ2001_Personnes_Morales_Arabic.csv'


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
    # worksheet_error.write(rowError, 0, Base_URL)
    # worksheet_error.write(rowError, 1, "Not Responding")
    # worksheet_error.write(rowError, 2, error)
    # rowError += 1
    with open(Error_File, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(Headers_Error)
        writer.writerow([URL, "Not Responding", str(error)])
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


if __name__=='__main__':
	File_paths= [File_path_Personnes_Physique,File_path_Personnes_Morales, File_path_Personnes_Physique_Arabic, File_path_Personnes_Morales_Arabic]
	
	First_run = False
	if First_run:
		File_paths_csv = [File_path_Personnes_Physique_CSV, File_path_Personnes_Morales_CSV,
					File_path_Personnes_Physique_Arabic_CSV, File_path_Personnes_Morales_Arabic_CSV]
		File_paths_txt = [File_path_Personnes_Physique_txt, File_path_Personnes_Morales_txt,
					File_path_Personnes_Physique_Arabic_txt, File_path_Personnes_Morales_Arabic_txt]
		if os.path.exists(File_path_log):
			os.remove(File_path_log)
		for path_csv in File_paths_csv:
			if os.path.exists(path_csv):
				os.remove(path_csv)
		for Path_txt in File_paths_txt:
			if os.path.exists(Path_txt):
				os.remove(Path_txt)
	
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

	Personnes_Physique_headers = ['NRC','Nom','Prenom']
	Personnes_Physique_Arabic_headers = ['NRC','Nom (Arabic)','Prenom (Arabic)']
	Personnes_Morales_headers = ['NRC','Raison Sociale']
	Personnes_Morales_Arabic_headers = ['NRC','Raison Sociale (Arabic)']
	
	with open(File_path_search_count,"w")as f:
		f.write("")
	with open(File_path_Personnes_Physique_txt,"a")as f:
		if f.tell() == 0:
			f.write("\t".join(Personnes_Physique_headers)+"\n")
			f.flush()
	with open(File_path_Personnes_Morales_txt,"a")as fw:
		if fw.tell() == 0:
			fw.write("\t".join(Personnes_Morales_headers)+"\n")
			fw.flush()
	with open(File_path_Personnes_Physique_Arabic_txt,"a")as f:
		if f.tell() == 0:
			f.write("\t".join(Personnes_Physique_Arabic_headers)+"\n")
			f.flush()
	with open(File_path_Personnes_Morales_Arabic_txt,"a")as fw:
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
	
	Home_URL = 'https://sidjilcom.cnrc.dz/web/cnrc/accueil'
	Arabic_url = 'https://sidjilcom.cnrc.dz/accueil?p_p_id=82&p_p_lifecycle=1&p_p_state=normal&p_p_mode=view&p_p_col_id=column-5&p_p_col_count=8&_82_struts_action=%2Flanguage%2Fview&_82_redirect=%2Faccueil&_82_languageId=ar_SA'
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

		log_index_English_flag = False
		if os.path.exists(File_path_log_index_English):
			log_index_English_flag = True
			with open(File_path_log_index_English, 'r', encoding='utf-8') as file:
				last_English = file.read().strip()

		if log_index_English_flag:
			start_index = English_alphabet_list.index(last_English) + 1
			alphabet = English_alphabet_list[start_index:]
		else:
			alphabet = English_alphabet_list

		for letter in alphabet[:]:
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
					'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36'}
			try:
				Info = requests.post(Link_Form,data=Form_Data,headers=Headers)
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
							Data = Script_soup.find_all('tr')
							Personnes_Physique_count = int(Script_soup.find_all('b')[2].string)+1
							i=1
							for items in Data:
								if i<Personnes_Physique_count:
									if items.find_all('td'):
										for j in items.find_all('td')[1].attrs:
											if j == 'style' and items.find_all('td')[1].attrs['style'] == 'text-align:left':
												# log_print('Record '+ str(i) + ' Added')
												try_count=1
												while True:
													try:
														with open(File_path_search_count,'a') as fh:
															fh.write('1\n')
															fh.flush()
														break
													except:
														if try_count>5:
															break
														try_count+=1
												# with open(File_path_search_count,"a")as f:
												# 	f.write("1\n")
												Personnes_Physique_data = []
												Personnes_Physique_data.append(items.find_all('td')[0].string.strip() if items.find_all('td')[0].string else '')
												Personnes_Physique_data.append(items.find_all('td')[1].string.strip() if items.find_all('td')[1].string else '')
												Personnes_Physique_data.append(items.find_all('td')[2].string.strip() if items.find_all('td')[2].string else '')
												with open(File_path_Personnes_Physique_CSV, 'a', newline='', encoding='utf-8') as file:
													writer = csv.writer(file)
													writer.writerow(Personnes_Physique_data)
												with open(File_path_Personnes_Physique_txt,"a")as f:
													f.write("\t".join(map(str,Personnes_Physique_data))+"\n")
													f.flush()
												int(i)
												i+=1
								else:
									if items.find_all('td'):
										for j in items.find_all('td')[1].attrs:
											if j == 'style' and items.find_all('td')[1].attrs['style'] == 'text-align:left':
												# log_print('Record ' + str(i) + ' Added')
												try_count=1
												while True:
													try:
														with open(File_path_search_count,'a') as fh:
															fh.write('1\n')
															fh.flush()
														break
													except:
														if try_count>5:
															break
														try_count+=1
												# with open(File_path_search_count,"a")as f:
												# 	f.write("1\n")
												Personnes_Morales_data = []
												Personnes_Morales_data.append(items.find_all('td')[0].string.strip() if items.find_all('td')[0].string else '')
												Personnes_Morales_data.append(items.find_all('td')[1].string.strip() if items.find_all('td')[1].string else '')
												with open(File_path_Personnes_Morales_CSV, 'a', newline='', encoding='utf-8') as file:
													writer = csv.writer(file)
													writer.writerow(Personnes_Morales_data)
												with open(File_path_Personnes_Morales_txt,"a")as f:
													f.write("\t".join(map(str,Personnes_Morales_data))+"\n")
													f.flush()
												int(i)
												i+=1
				log_print('Complete ' + letter)
				with open(File_path_log_index_English, 'w', encoding='utf-8') as file:
					file.write(letter)
					file.flush()
				# data = pd.read_excel(File_path_Personnes_Physique)
				# file_df_first_record = data.drop_duplicates(subset=["NRC"])
				# file_df_first_record.to_excel(File_path_Personnes_Physique, index=False)
				
				# data2 = pd.read_excel(File_path_Personnes_Morales)
				# file_df_first_record2 = data2.drop_duplicates(subset=["NRC"])
				# file_df_first_record2.to_excel(File_path_Personnes_Morales, index=False)
			except:
				exception(Home_URL)
			finally:
				Dereference(Info)
				Data_soup.decompose()
				Script_soup.decompose()
				convertCSVExcel(File_path_Personnes_Physique_CSV,File_path_Personnes_Physique)
				convertCSVExcel(File_path_Personnes_Morales_CSV, File_path_Personnes_Morales)
		
		

	except:
		exception(Home_URL)
	finally:
		Home_soup.decompose()
		duplicate(File_path_Personnes_Physique)
		duplicate(File_path_Personnes_Morales)

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

		log_index_Arabic_flag = False
		if os.path.exists(File_path_log_index_Arabic):
			log_index_Arabic_flag = True
			with open(File_path_log_index_Arabic, 'r', encoding='utf-8') as file:
				last_Arabic = file.read().strip()

		if log_index_Arabic_flag:
			start_index = Arabic_alphabet_list.index(last_Arabic) + 1
			Arabic_alphabet = Arabic_alphabet_list[start_index:]
		else:
			Arabic_alphabet = Arabic_alphabet_list

		for letter in Arabic_alphabet[:]:
			fields = {
				'hidden': 'goRecherche',
				'critere': letter
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
					'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36'}
			try:
				try_count_ar = 1
				while True:
					try:
						Ar_Info = requests.post(Ar_Link_Form,data=Form_Data,headers=Headers)
						break
					except:
						if try_count_ar>3:
							break
						try_count_ar += 1

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
							Ar_Data = Ar_Script_soup.find_all('tr')
							Personnes_Physique_count_ar = int(Ar_Script_soup.find_all('b')[2].string)+1
							i=1
							for items in Ar_Data:
								if i<Personnes_Physique_count_ar:
									if items.find_all('td'):
										# log_print('Record ' + str(i) + ' Added')
										try_count=1
										while True:
											try:
												with open(File_path_search_count,'a') as fh:
													fh.write('1\n')
													fh.flush()
												break
											except:
												if try_count>5:
													break
												try_count+=1
										# with open(File_path_search_count,"a")as f:
										# 	f.write("1\n")
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
										try_count=1
										while True:
											try:
												with open(File_path_search_count,'a') as fh:
													fh.write('1\n')
													fh.flush()
												break
											except:
												if try_count>5:
													break
												try_count+=1
										# with open(File_path_search_count,"a")as f:
										# 	f.write("1\n")
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
				log_print('Complete ' + letter)
				with open(File_path_log_index_Arabic, 'w', encoding='utf-8') as file:
					file.write(letter)
					file.flush()
				# data = pd.read_excel(File_path_Personnes_Physique)
				# file_df_first_record = data.drop_duplicates(subset=["NRC"])
				# file_df_first_record.to_excel(File_path_Personnes_Physique, index=False)
				
				# data2 = pd.read_excel(File_path_Personnes_Morales)
				# file_df_first_record2 = data2.drop_duplicates(subset=["NRC"])
				# file_df_first_record2.to_excel(File_path_Personnes_Morales, index=False)
			except:
				exception(Arabic_url)
			finally:
				Dereference(Ar_Info)
				Ar_Data_soup.decompose()
				Ar_Script_soup.decompose()
				convertCSVExcel(File_path_Personnes_Physique_Arabic_CSV, File_path_Personnes_Physique_Arabic)
				convertCSVExcel(File_path_Personnes_Morales_Arabic_CSV, File_path_Personnes_Morales_Arabic)

	except:
		exception(Arabic_url)
	finally:
		Ar_Home_soup.decompose()
		duplicate(File_path_Personnes_Physique_Arabic)
		duplicate(File_path_Personnes_Morales_Arabic)
	with open(File_path_log_index_English, 'w', encoding='utf-8') as f1:
		f1.write('')
		f1.flush()
	with open(File_path_log_index_Arabic, 'w', encoding='utf-8') as f1:
		f1.write('')
		f1.flush()
	log_print('Success')
	exit()
  
database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
	for File_path in File_paths:
		delete_task(conn, File_path)