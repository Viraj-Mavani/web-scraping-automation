import os
import requests
import re
import time
# try:
	# from PIL import Image
# except ImportError:
	# import Image
# import pytesseract
# import cv2
import traceback
import xlsxwriter
import sqlite3
from sqlite3 import Error
import warnings
import json
warnings.filterwarnings("ignore")

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

def regex_match(regex,content):
	match = re.search(regex,content,re.I)
	if match:
		return match.group(1)
	else:
		return ''
		
def attribute_replace(content):
	data = str(content)
	data = re.sub(r'<[^>]*?>',' ', data, flags = re.I|re.M)
	data = re.sub(r'&amp;','&', data, flags = re.I|re.M)
	data = re.sub(r'&nbsp;','', data, flags = re.I|re.M)
	data = re.sub(r'\s+',' ', data, flags = re.I|re.M)
	data = re.sub(r'^\s+','', data, flags = re.I|re.M)
	data = re.sub(r'\s+$','', data, flags = re.I|re.M)
	data = re.sub(r'None','', data, flags = re.I|re.M)
	data = re.sub(r'\|$','', data, flags = re.I|re.M)
	data = re.sub(r'^\|','', data, flags = re.I|re.M)
	# data = re.sub(r',\s+,','', data, flags = re.I|re.M)
	# data = re.sub(r',$','', data, flags = re.I|re.M)
	return data
	
def listPageCollector(content):
	try:
		global row1
		
		tableBlock = re.findall(r'<td>\s*([^>]*?)\s*<\/td>\s*<td>\s*([^>]*?)\s*<\/td>\s*<td>\s*([^>]*?)\s*<\/td>\s*<td>\s*([^>]*?)\s*<\/td>\s*<td>\s*([^>]*?)\s*<\/td>\s*<td>\s*([^>]*?)\s*<\/td>\s*<td>\s*([^>]*?)\s*<\/td>\s*<td>\s*([^>]*?)\s*<\/td>\s*<td>\s*([^>]*?)\s*<\/td>',content)	
		
		for tableRow in tableBlock:
			to_Merge = attribute_replace(tableRow[0])
			filtering = attribute_replace(tableRow[1])
			cross_Off = attribute_replace(tableRow[2])
			certificate_Date = attribute_replace(tableRow[3])
			certificate_Number = attribute_replace(tableRow[4])
			capital = attribute_replace(tableRow[5])
			activity = attribute_replace(tableRow[6])
			company_Name = attribute_replace(tableRow[7])
			file_Number = attribute_replace(tableRow[8])
			search_page = [file_Number,company_Name,activity,capital,certificate_Number,certificate_Date,cross_Off,filtering,to_Merge]
			# time.sleep(5)
			with open('E:\ADIP-PY\OPtxt\ADIP-IQ1202_Searchpage.txt',"a",encoding='utf8')as f:
				f.write("\t".join(map(str,search_page))+"\n")
			try_count=1
			while True:
				try:
					with open('E:\ADIP-PY\Counts\ADIP-IQ1202_Count.txt','a') as fh:
						fh.write('1\n')
					break
				except:
					if try_count>5:
						with open('Log_File.txt','a',encoding='utf-8') as logfile:
							logfile.write(traceback.format_exc())
						break
					time.sleep(1)
					try_count+=1
			# try:
				# with open('E:\ADIP-PY\Counts\ADIP-IQ1202_Count.txt',"a")as f:
					# f.write("1\n")
			# except:
				# pass
			# time.sleep(5)
			worksheet1.write_string(row1, 0, file_Number)
			worksheet1.write_string(row1, 0 + 1, company_Name)
			worksheet1.write_string(row1, 0 + 2, activity)
			worksheet1.write_string(row1, 0 + 3, capital)
			worksheet1.write_string(row1, 0 + 4, certificate_Number)
			worksheet1.write_string(row1, 0 + 5, certificate_Date)
			worksheet1.write_string(row1, 0 + 6, cross_Off)
			worksheet1.write_string(row1, 0 + 7, filtering)
			worksheet1.write_string(row1, 0 + 8, to_Merge)
			row1 +=1
	except Exception as e:
		error = traceback.format_exc()
		print(error)	
		
if __name__ == "__main__":	
	File_paths= ['E:\ADIP-PY\OP\ADIP-IQ1202_Searchpage.xlsx','E:\ADIP-PY\OPtxt\ADIP-IQ1202_Searchpage.txt','E:\ADIP-PY\Error\ADIP-IQ1202_Error.xlsx','E:\ADIP-PY\Counts\ADIP-IQ1202_Count.txt']
	cachePath = 'Cache_IQ1202/'
	if not os.path.isdir(cachePath):
		os.makedirs(cachePath)
	
	with open('E:\ADIP-PY\Counts\ADIP-IQ1202_Count.txt',"w")as f:
		f.write("")	
	sess=requests.session()
	sess.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.183 Safari/537.36'
	
	##############Output in Text#############
	search_page = ["File Number","Company Name","Activity","Capital","Registration Number","Registration Date","Deletion","Liquidation","Merge"]
	with open('E:\ADIP-PY\OPtxt\ADIP-IQ1202_Searchpage.txt',"a")as f:
		f.write("\t".join(search_page)+"\n")
		
	##############Output in Excel#############	
	row1=1
	workbook1 = xlsxwriter.Workbook('E:\ADIP-PY\OP\ADIP-IQ1202_Searchpage.xlsx')
	worksheet1 = workbook1.add_worksheet()
	bold = workbook1.add_format({'bold': 1})
	worksheet1.write('A1', 'File Number', bold)
	worksheet1.write('B1', 'Company Name', bold)
	worksheet1.write('C1', 'Activity', bold)
	worksheet1.write('D1', 'Capital', bold)
	worksheet1.write('E1', 'Registration Number', bold)
	worksheet1.write('F1', 'Registration Date', bold)
	worksheet1.write('G1', 'Deletion', bold)
	worksheet1.write('H1', 'Liquidation', bold)
	worksheet1.write('I1', 'Merge', bold)
	
	###############################################
	
	row2=1
	workbook2 = xlsxwriter.Workbook('E:\ADIP-PY\Error\ADIP-IQ1202_Error.xlsx')
	worksheet2 = workbook2.add_worksheet()
	bold = workbook2.add_format({'bold': 1})
	worksheet2.write('A1', 'URL', bold)
	worksheet2.write('B1', 'Responding status', bold)
	worksheet2.write('C1', 'Error', bold)
	try:
		obj = sess.get('http://tasjeel.mot.gov.iq/search/national_page',headers={'Host': 'tasjeel.mot.gov.iq','Upgrade-Insecure-Requests': '1','User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36'})
		with open('{}home_content.html'.format(cachePath),'wb') as fh:
			fh.write(obj.content)
		home_content=obj.text
		
		VIEWSTATE = regex_match('<input[^>]*?id="__VIEWSTATE"\s*value="([^>]*?)"[^>]*?>',home_content)
		EVENTVALIDATION = regex_match('<input[^>]*?id="__EVENTVALIDATION"\s*value="([^>]*?)"[^>]*?>',home_content)
		VIEWSTATEGENERATOR = regex_match('<input[^>]*?id="__VIEWSTATEGENERATOR"\s*value="([^>]*?)"[^>]*?>',home_content)
		paramblock={'__EVENTTARGET': '','__EVENTARGUMENT': '','__VIEWSTATE': VIEWSTATE,'__VIEWSTATEGENERATOR': VIEWSTATEGENERATOR,'__SCROLLPOSITIONX': '0','__SCROLLPOSITIONY': '0','__EVENTVALIDATION': EVENTVALIDATION,'btn_search': 'بحث','txt_search': 'ال'}
		# print(paramblock)
		obj1=sess.post('http://tasjeel.mot.gov.iq/search/national_page',data=paramblock,headers={'Content-Type':'application/x-www-form-urlencoded','Host':'tasjeel.mot.gov.iq','Referer':'http://tasjeel.mot.gov.iq/search/national_page','User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36','Upgrade-Insecure-Requests': '1','Origin': 'http://tasjeel.mot.gov.iq'})
		time.sleep(30)
		with open('{}search_result_content_ال_1.html'.format(cachePath),'wb') as fh:
			fh.write(obj1.content)
		listPage = obj1.text
		listPageCollector(listPage) 
	except:
		print("Error")
		error = traceback.format_exc()
		print(error)
		worksheet2.write_string(row2, 0, 'http://tasjeel.mot.gov.iq/search/national_page')
		worksheet2.write_string(row2, 0 + 1, "Not_responding")
		worksheet2.write_string(row2, 0 + 13, "TimeoutError")
		row2 +=1
		workbook2.close()
	finally:
		workbook1.close()
		# if os.path.isfile('E:\ADIP-PY\Counts\ADIP-IQ1202_Processed_Input.txt'):
			# os.remove('E:\ADIP-PY\Counts\ADIP-IQ1202_Processed_Input.txt')
database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
	for File_path in File_paths:
		delete_task(conn, File_path)