import requests
import re
import os
import time
import traceback
import xlsxwriter
from requests_toolbelt.multipart.encoder import MultipartEncoder
import sqlite3
from sqlite3 import Error


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
		global row1
		mainBlock = re.findall(r'(<div\s*class="panel\s*panel-default\s*memberpanel">[\w\W]*?<\/div>\s*<\/div>)',Content)
		
		for block in mainBlock:
			Company_Name = attribute_replace(regex_match('<div\s*class="panel-heading">\s*([\w\W]*?)\s*<\/div>',block))
			#print(Company_Name)
			Phone1 = attribute_replace(regex_match('>[^>]*?\s*هاتف\s*1\s*:\s*([\d\-]+)\s*[^>]*?<',block))
			Phone1 = re.sub(r'([\d\s]+)(-)([\d\s]+)',r'\3\2\1',Phone1)
			Phone2 = attribute_replace(regex_match('>\s*هاتف\s*2\s*:\s*([\d\-]+)\s*[^>]*?<',block))
			Phone2 = re.sub(r'([\d\s]+)(-)([\d\s]+)',r'\3\2\1',Phone2)
			Fax = attribute_replace(regex_match('>\s*فاكس\s*:\s*([\d\-]+)\s*[^>]*?<',block))
			Fax = re.sub(r'([\d\s]+)(-)([\d\s]+)',r'\3\2\1',Fax)
			Activity = attribute_replace(regex_match('<h5>\s*النشاط\s*<\/h5>\s*([\w\W]*?)\s*<\/ul>',block))
			
			detail_page = [Company_Name,Phone1,Phone2,Fax,Activity]
			with open(File_path_txt,"a",encoding='utf8')as f:
				f.write("\t".join(map(str,detail_page))+"\n")
			# worksheet1.write_string(row1, 0, Company_Number)
			if Company_Name:
				with open(File_path_count,"a")as fh:
					fh.write("1\n")
			worksheet1.write_string(row1, 0, Company_Name)
			worksheet1.write_string(row1, 0 + 1, Phone1)
			worksheet1.write_string(row1, 0 + 2, Phone2)
			worksheet1.write_string(row1, 0 + 3, Fax)
			worksheet1.write_string(row1, 0 + 4, Activity)
			row1 +=1
		nextpage = regex_match('<a\s*href="(?!\#)([^>]*?)"[^>]*?>\s*<span[^>]*?>»<\/span>\s*<\/a>',Content)
		
		if nextpage:
			# time.sleep(5)
			if os.path.exists('{}nextpage_{}_{}.html'.format(cachePath,arabicletter,page)):
				with open('{}nextpage_{}_{}.html'.format(cachePath,arabicletter,page),'r',encoding='utf-8') as fh:
					content_1=fh.read()
			else:
				obj = sess.get('http://homschamber.com'+nextpage,timeout = 300)
				with open('{}nextpage_{}_{}.html'.format(cachePath,arabicletter,page),'wb') as fh:
					fh.write(obj.content)
					content_1=obj.text
			page+=1
			nextpageCon = obj.text
			searchpage_Collector(content_1,arabicletter,page)
	except Exception as e:
		error = traceback.format_exc()
		print(error)

if __name__=='__main__':
	File_path = 'D:\Projects\CedarPython\ADIP-SY602\OP\ADIP-SY602_Output.xlsx'
	File_path_txt = 'D:\Projects\CedarPython\ADIP-SY602\OPtxt\ADIP-SY602_Output.txt'
	File_path_count = 'D:\Projects\CedarPython\ADIP-SY602\Counts\ADIP-SY602_Count.txt'
	sess = requests.Session()
	sess.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.183 Safari/537.36'
	
	cachePath = 'CacheSY602/'
	if not os.path.isdir(cachePath):
		os.makedirs(cachePath)
	
	
	detail_page = ["Company Name","Phone1","Phone2","Fax","Activity"]
	
	error_file = 'D:\Projects\CedarPython\ADIP-SY602\Error\ADIP-SY602_Error.txt'
	
	error_page = ["URL","Responding status","Error"]
	
	with open(File_path_txt,"w")as f:
		f.write("\t".join(detail_page)+"\n")
	with open(File_path_count,"w")as f:
		f.write("")
	
	with open(error_file,"w")as f:
		f.write("\t".join(error_page)+"\n")
	row1=1
	homeurl = ['']
	workbook1 = xlsxwriter.Workbook(
		'D:\Projects\CedarPython\ADIP-SY602\OP\ADIP-SY602_Output.xlsx')
	worksheet1 = workbook1.add_worksheet()
	bold = workbook1.add_format({'bold': 1})
	worksheet1.write('A1', 'Company Name', bold)
	worksheet1.write('B1', 'Phone1', bold)
	worksheet1.write('C1', 'Phone2', bold)
	worksheet1.write('D1', 'Fax', bold)
	worksheet1.write('E1', 'Activity', bold)
	
	row2=1
	workbook2 = xlsxwriter.Workbook(
		'D:\Projects\CedarPython\ADIP-SY602\Error\ADIP-SY602_Error.xlsx')
	worksheet2 = workbook2.add_worksheet()
	bold = workbook2.add_format({'bold': 1})
	worksheet2.write('A1', 'URL', bold)
	worksheet2.write('B1', 'Responding status', bold)
	worksheet2.write('C1', 'Error', bold)
	try:
		arabicWords = ["ي","و","ه","ن","م","ل","ك","ق","ف","غ","ع","ظ","ط","ض","ص","ش","س","ز","ر","ذ","د","خ","ح","ج","ث","ت","ب","ا"]

		page=1
		for i in arabicWords:
			for j in arabicWords:
				for k in arabicWords:
					arabicletter = i+j+k
					print(arabicletter)
					# time.sleep(5)
					homeurl = 'https://homschamber.com/members-index/?ftxt='+arabicletter+'&searchType=1'
					# obj=sess.get('http://homschamber.com/members-index/?ftxt='+arabicletter+'&ftxt2=',timeout = 300)
					if os.path.exists('{}Searchpage_{}.html'.format(cachePath,arabicletter)):
						with open('{}Searchpage_{}.html'.format(cachePath,arabicletter),'r',encoding='utf-8') as fh:
							companycontent=fh.read()
					else:
						obj=sess.get('https://homschamber.com/members-index/?ftxt='+arabicletter+'&searchType=1',timeout = 300)
						with open('{}Searchpage_{}.html'.format(cachePath,arabicletter),'wb') as fh:
							fh.write(obj.content)
						companycontent = obj.text
					
					searchpage_Collector(companycontent,arabicletter,page)
		workbook1.close()
	except:
		# print("Error")
		worksheet2.write_string(row2, 0, homeurl)
		worksheet2.write_string(row2, 0 + 1, "Not_responding")
		worksheet2.write_string(row2, 0 + 2, "TimeoutError")
		row2 +=1
		workbook2.close()
database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
	delete_task(conn, File_path)
	delete_task(conn, File_path_txt)
	delete_task(conn, File_path_count)