import csv
import sys
import requests
import re
import os
import traceback
# import xlsxwriter
# from requests_toolbelt.multipart.encoder import MultipartEncoder
import sqlite3
from sqlite3 import Error
from bs4 import BeautifulSoup 
import pandas as pd


BasePath = 'D:\Projects\CedarPython\ADIP-SY602'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-SY602_Log.txt'
# File_path_log_index_English = BasePath + '\\Log\\ADIP-SY602_Log_Index_English.txt'
File_path_log_index_Arabic = BasePath + '\\Log\\ADIP-SY602_Log_Index_Arabic.txt'
######### Excel #########
File_path = BasePath + '\\OP\\ADIP-SY602_Output.xlsx'
######### CSV #########
File_path_CSV = BasePath + '\\OPcsv\\ADIP-SY602_Output.csv'
######### Text #########
File_path_txt= BasePath + '\\OPtxt\\ADIP-SY602_Output.txt'
######### Count #########
File_path_count= BasePath + '\\Counts\\ADIP-SY602_Count.txt'
######### Error #########
Error_File = BasePath + '\\Error\\ADIP-SY602_Error.xlsx'
Error_File_CSV = BasePath + '\\Error\\ADIP-SY602_Error.csv'


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
#######################


def log_print(message):
    with open(File_path_log, 'a', encoding='utf-8') as file:
        file.write(message + '\n')
        file.flush()
    print(message)


def exception():
    # global rowError
    Headers_Error = ['URL', 'Not Responding', 'Error']
    error = traceback.format_exc()
    exception_type, exception_object, exception_traceback = sys.exc_info()
    with open(Error_File, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(Headers_Error)
        writer.writerow([homeurl, "Not Responding", str(error)])
    df = pd.read_csv(Error_File_CSV, encoding='utf-8')
    df.to_excel(Error_File, index=False)


def convertCSVExcel(File_path_CSV, File_path_EXL):
    df = pd.read_csv(File_path_CSV, encoding='utf-8')
    df.to_excel(File_path_EXL, index=False)


def regex_match(regex,content):
	match = re.search(regex,content,flags=re.I)
	if match:
		return match.group(1)
	else:
		return ''
	
	
def attribute_replace(content):
	data = str(content)
	data = re.sub(r'<\/li>','listclose', data, flags = re.I|re.M)
	data = re.sub(r'<[^>]*?>',' ', data, flags = re.I|re.M)
	data = re.sub(r'listclose','|', data, flags = re.I|re.M)
	data = re.sub(r'&amp;','&', data, flags = re.I|re.M)
	data = re.sub(r'&nbsp;',' ', data, flags = re.I|re.M)
	data = re.sub(r'\s+',' ', data, flags = re.I|re.M)
	data = re.sub(r'^\s+','', data, flags = re.I|re.M)
	data = re.sub(r'\s+$','', data, flags = re.I|re.M)
	data = re.sub(r'None','', data, flags = re.I|re.M)
	data = re.sub(r'\|$','', data, flags = re.I|re.M)
	data = re.sub(r'^\|','', data, flags = re.I|re.M)
	# data = re.sub(r',\s+,','', data, flags = re.I|re.M)
	# data = re.sub(r',$','', data, flags = re.I|re.M)
	return data
	
 
def searchpage_Collector(Content,arabicletter,page):
	try:
		global homeurl
		mainBlock = re.findall(r'(<div\s*class="panel\s*panel-default\s*memberpanel">[\w\W]*?<\/div>\s*<\/div>)',Content)
		data = BeautifulSoup(Content, 'html.parser')
		Info = data.find_all('div',class_='memberpanel')
		i=0
		
		for block in mainBlock:
			# time.sleep(0.1)
			Company_Name = attribute_replace(regex_match('<div\s*class="panel-heading">\s*([\w\W]*?)\s*<\/div>',block))
			# print(page)
			Phone1 = attribute_replace(regex_match('>[^>]*?\s*هاتف\s*1\s*:\s*([\d\-]+)\s*[^>]*?<',block))
			Phone1 = re.sub(r'([\d\s]+)(-)([\d\s]+)',r'\3\2\1',Phone1)
			Phone2 = attribute_replace(regex_match('>\s*هاتف\s*2\s*:\s*([\d\-]+)\s*[^>]*?<',block))
			Phone2 = re.sub(r'([\d\s]+)(-)([\d\s]+)',r'\3\2\1',Phone2)
			Fax = attribute_replace(regex_match('>\s*فاكس\s*:\s*([\d\-]+)\s*[^>]*?<',block))
			Fax = re.sub(r'([\d\s]+)(-)([\d\s]+)',r'\3\2\1',Fax)
			Activity = attribute_replace(regex_match('<h5>\s*النشاط\s*<\/h5>\s*([\w\W]*?)\s*<\/ul>',block))
			Mobile = Info[i].find('i',class_='fa-mobile').next.next.string if Info[i].find('i',class_='fa-mobile') else ''
			Email = Info[i].find('i',class_='fa-envelope').next.string if Info[i].find('i',class_='fa-envelope') else ''
			if Email != '':
				Email = Email.split(':')[1]
			
			if Company_Name:
				try_count=1
				while True:
					try:
						with open(File_path_count,'a') as fh:
							fh.write('1\n')
						break
					except:
						if try_count>3:
							break
						# time.sleep(1)
						try_count+=1
			detail_page = [Company_Name,Phone1,Phone2,Mobile,Email,Fax,Activity]
			with open(File_path_CSV, 'a', newline='', encoding='utf-8') as file:
				writer = csv.writer(file)
				writer.writerow(detail_page)
			with open(File_path_txt,"a",encoding='utf8')as f:
				f.write("\t".join(map(str,detail_page))+"\n")
				f.flush()
			i+=1
		log_print(f'{arabicletter} {page}')
		pagination = data.find('ul',class_='pagination')
		
		if pagination:
			nextpage = pagination.find_all('a',class_='page-link')[-1]
			nextpagelink = nextpage.attrs['href']
		else:
			nextpagelink = None
		# nextpage = regex_match('<a\s*class="(page-link)"\s*href="(?!\#)([^>]*?)"[^>]*?>\s*<span[^>]*?>»<',Content)
		data.decompose()
		del Info
		if nextpagelink:
			# time.sleep(5)
			if os.path.exists('{}nextpage_{}_{}.html'.format(cachePath,arabicletter,page)):
				with open('{}nextpage_{}_{}.html'.format(cachePath,arabicletter,page),'r',encoding='utf-8') as fh:
					content_1=fh.read()
			else:
				# time.sleep(0.5)
				obj = sess.get('http://homschamber.com'+nextpagelink,timeout = 300)
				with open('{}nextpage_{}_{}.html'.format(cachePath,arabicletter,page),'wb') as fh:
					fh.write(obj.content)
					content_1=obj.text
			page+=1
			# nextpageCon = obj.text
			searchpage_Collector(content_1,arabicletter,page)
	except Exception as e:
		exception()


if __name__=='__main__':
    
	# Create directories if they don't exist
	directories = [BasePath + '\\OP',
					BasePath + '\\Optxt',
					BasePath + '\\OPcsv',
					BasePath + '\\Error',
					BasePath + '\\Counts',
					BasePath + '\\Log']
	for directory in directories:
		if not os.path.exists(directory):
			os.makedirs(directory)
	
	First_run = True
	if First_run:
		if os.path.exists(File_path_log):
			os.remove(File_path_log)

		if os.path.exists(File_path_CSV):
			os.remove(File_path_CSV)

		if os.path.exists(File_path_txt):
			os.remove(File_path_txt)

		if os.path.exists(Error_File_CSV):
			os.remove(Error_File_CSV)

	sess = requests.Session()
	sess.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.183 Safari/537.36'
	
	cachePath = BasePath + '\\CacheSY602\\'
	if not os.path.isdir(cachePath):
		os.makedirs(cachePath)
  
	detail_page_headers = ["Company Name","Phone1","Phone2","Mobile","Email","Fax","Activity"]
	
	with open(File_path_count,"w")as f:
		f.write("")
	with open(File_path_txt,"a")as f:
		if f.tell() == 0:
			f.write("\t".join(detail_page_headers)+"\n")

	with open(File_path_CSV, "a", newline='', encoding='utf-8') as f:
		writer = csv.writer(f)
		if f.tell() == 0:
			writer.writerow(detail_page_headers)

	try:
		arabicWords = ["ي","و","ه","ن","م","ل","ك","ق","ف","غ","ع","ظ","ط","ض","ص","ش","س","ز","ر","ذ","د","خ","ح","ج","ث","ت","ب","ا"]
		# arabicWords = ["ب"]
		
		page=1
		for i in arabicWords:
			for j in arabicWords:
				for k in arabicWords:
					arabicletter = i+j+k
					# time.sleep(5)
					homeurl = 'https://homschamber.com/members-index/?ftxt='+arabicletter+'&searchType=1'
					# obj=sess.get('http://homschamber.com/members-index/?ftxt='+arabicletter+'&ftxt2=',timeout = 300)
					if os.path.exists('{}Searchpage_{}.html'.format(cachePath,arabicletter)):
						with open('{}Searchpage_{}.html'.format(cachePath,arabicletter),'r',encoding='utf-8') as fh:
							companycontent=fh.read()
					else:
						# time.sleep(1)
						obj=sess.get('https://homschamber.com/members-index/?ftxt='+arabicletter+'&searchType=1', verify=False, timeout = 300)
						with open('{}Searchpage_{}.html'.format(cachePath,arabicletter),'wb') as fh:
							fh.write(obj.content)
						companycontent = obj.text
					
					searchpage_Collector(companycontent,arabicletter,page)
					log_print(f'Complete {arabicletter}\n')
			convertCSVExcel(File_path_CSV, File_path)
		
	except Exception as e:
		exception()
	finally:
		data = pd.read_excel(File_path)
		data_file =data.drop_duplicates()
		data_file.to_excel(File_path,index=False)
		exit()
database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
	delete_task(conn, File_path)
	delete_task(conn, File_path_txt)
	delete_task(conn, File_path_count)