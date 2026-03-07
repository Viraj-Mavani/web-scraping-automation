import os
import requests
import re
from requests_toolbelt.multipart.encoder import MultipartEncoder
import traceback
import xlsxwriter
import sqlite3
from sqlite3 import Error
import time
import warnings
import json
import html
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
	data = re.sub(r'&nbsp;',' ', data, flags = re.I|re.M)
	data = re.sub(r'\"','', data, flags = re.I|re.M)
	data = re.sub(r'\s+',' ', data, flags = re.I|re.M)
	data = re.sub(r'^\.','', data, flags = re.I|re.M)
	data = re.sub(r'^\s+','', data, flags = re.I|re.M)
	data = re.sub(r'\s+$','', data, flags = re.I|re.M)
	data = re.sub(r'None','', data, flags = re.I|re.M)
	data = re.sub(r'\|$','', data, flags = re.I|re.M)
	data = re.sub(r'^\|','', data, flags = re.I|re.M)
	# data = re.sub(r',\s+,','', data, flags = re.I|re.M)
	# data = re.sub(r',$','', data, flags = re.I|re.M)
	return data

def detailpageCollector(detailUrl,nif):
	try:
		global row2
		obj = sess.get(detailUrl)
		detailcontent = obj.text
		detailcontent = re.sub('&amp;','&',detailcontent)
		detailcontent = html.unescape(detailcontent)
		with open(f'{cachePath}detailPage_{nif}.html','w',encoding='utf8') as fh:
			fh.write(detailcontent)
		accessCode = ''
		companyNumber = attribute_replace(regex_match('<[^>]*?>\s*Matrícula\s*:\s*([\w\W]*?)\s*<br[^>]*?>',detailcontent))
		firm = attribute_replace(regex_match('<[^>]*?>\s*Firma\s*:\s*([\w\W]*?)\s*<br[^>]*?>',detailcontent))
		nif = attribute_replace(regex_match('<[^>]*?>\s*NIF\s*:\s*([\w\W]*?)\s*<br[^>]*?>',detailcontent))
		print(nif)
		titlematch = re.search('<[^>]*?>\s*([^>]*?)\s*\,\s*([^>]*?)\s*\,\s*assinado\s*em\s*([\d\-]+)[^>]*?<',detailcontent)
		[name ,title, date, folioregistryNumber, legalForm, currency, capital] = ['']*7
		if titlematch:
			title = attribute_replace(titlematch.group(1))
			name = attribute_replace(titlematch.group(2))
			date = attribute_replace(titlematch.group(3))
		elif re.search('<b>\s*([^>]*?)\s*<br>\s*Assinado\s*electronicamente\s*em\s*([\d\-]+)[^>]*?<',detailcontent):
			titlematch = re.search('<b>\s*([^>]*?)\s*<br>\s*Assinado\s*electronicamente\s*em\s*([\d\-]+)[^>]*?<',detailcontent)
			title = attribute_replace(titlematch.group(1))
			# name = attribute_replace(titlematch.group(2))
			date = attribute_replace(titlematch.group(2))
		inscMatch = re.search('>\s*Insc.([^>]*?)\s*-\s*([\w\W]*?)\s*<br',detailcontent)
		if inscMatch:
			folioregistryNumber = attribute_replace(inscMatch.group(1))
			legalForm  = attribute_replace(inscMatch.group(2))
		address = attribute_replace(regex_match('>\s*SEDE:\s*([\w\W]*?)\s*<\/p>',detailcontent))
		activity = attribute_replace(regex_match('>\s*OBJECTO:\s*([\w\W]*?)\s*<\/p>',detailcontent))
		capitalmatch = re.search('>\s*CAPITAL:\s*([\w\W]*?)\s*([\d\s\,\.]+)\s*[^>]*?<\/p>',detailcontent)
		if capitalmatch:
			currency = attribute_replace(capitalmatch.group(1))
			capital = attribute_replace(capitalmatch.group(2))
		shareholderBlock = re.findall(r'>\s*([\d]+º)\s*<strong>\s*([^>]*?)\s*</strong>\,\s*([^>]*?)\,\s*com\s*uma\s*quota\s*no[^>]*?\s*([\d\,\.]+)[^>]*?<',detailcontent)
		# print(shareholderBlock)
		shareholdernumbers = []
		shareholdernames = []
		shareholderdess = []
		shareholdervalues = []
		for shareholders in shareholderBlock:
			shareholdernumber = attribute_replace(shareholders[0])
			shareholdername = attribute_replace(shareholders[1])
			shareholderdes = attribute_replace(shareholders[2])
			shareholdervalue = attribute_replace(shareholders[3])
			shareholdernumbers.append(shareholdernumber)
			shareholdernames.append(shareholdername)
			shareholderdess.append(shareholderdes)
			shareholdervalues.append(shareholdervalue)
		shareholdernum = ("|".join(shareholdernumbers))
		shareholdernam = ("|".join(shareholdernames))
		shareholderdescription = ("|".join(shareholderdess))
		shareholderval = ("|".join(shareholdervalues))
		managementDes = attribute_replace(regex_match('>\s*GERÊNCIA\s*:\s*<\/strong>\s*([\w\W]*?)\s*<\/p>',detailcontent))
		managerBlock = re.findall(r'<strong>\s*([^>]*?)\s*<\/strong>\s*-\s*Contribuinte\s*nº\s*([^>]*?)\s*<\/p>',detailcontent)
		managernames = []
		managernumbers = []
		for managers in managerBlock:
			managername = attribute_replace(managers[0])
			managernumber = attribute_replace(managers[1])
			managernames.append(managername)
			managernumbers.append(managernumber)
		managernom = ("|".join(managernames))
		managerno = ("|".join(managernumbers))
		Waytooblige = attribute_replace(regex_match('<strong>\s*FORMA\s*DE\s*OBRIGAR\:\s*<\/strong>\s*([^>]*?)\s*<br[^>]*?>',detailcontent))
		
		
		detail_page = [accessCode,companyNumber,firm,nif,title,name,date,folioregistryNumber,legalForm,address,activity,capital,currency,shareholdernum,shareholdernam,shareholderdescription,shareholderval,managementDes,managernom,managerno,Waytooblige]
		with open('D:\Projects\CedarPython\ADIP-AO2202\OPtxt\ADIP-AO2202_Companyinfo.txt',"a",encoding='utf8')as f:
			f.write("\t".join(map(str,detail_page))+"\n")
		worksheet2.write_string(row2, 0, accessCode)
		worksheet2.write_string(row2, 0 + 1, companyNumber)
		worksheet2.write_string(row2, 0 + 2, firm)
		worksheet2.write_string(row2, 0 + 1, nif)
		worksheet2.write_string(row2, 0 + 4, title)
		worksheet2.write_string(row2, 0 + 5, name)
		worksheet2.write_string(row2, 0 + 6, date)
		worksheet2.write_string(row2, 0 + 7, folioregistryNumber)
		worksheet2.write_string(row2, 0 + 8, legalForm)
		worksheet2.write_string(row2, 0 + 9, address)
		worksheet2.write_string(row2, 0 + 10, activity)
		worksheet2.write_string(row2, 0 + 11, capital)
		worksheet2.write_string(row2, 0 + 12, currency)
		worksheet2.write_string(row2, 0 + 13, shareholdernum)
		worksheet2.write_string(row2, 0 + 14, shareholdernam)
		worksheet2.write_string(row2, 0 + 15, shareholderdescription)
		worksheet2.write_string(row2, 0 + 16, shareholderval)
		worksheet2.write_string(row2, 0 + 17, managementDes)
		worksheet2.write_string(row2, 0 + 18, managernom)
		worksheet2.write_string(row2, 0 + 19, managerno)
		worksheet2.write_string(row2, 0 + 20, Waytooblige)
		row2 +=1
	except Exception as e:
		error = traceback.format_exc()
		print(error)
def searchPageCollector(content):
	try:
		global row1
		searchBlock = re.findall('<tr[^>]*?>([\w\W]*?)<\/tr>',content)
		
		if searchBlock:
			for block in searchBlock:
				searchitems = re.search('<td[^>]*?>\s*([^>]*?)\s*<\/td>\s*<td[^>]*?>\s*([^>]*?)\s*<br>\s*([^>]*?)\s*<\/td>\s*<td[^>]*?>\s*\:?\s*([^>]*?)\s*<\/td>\s*<td[^>]*?>\s*([^>]*?)\s*<\/td>[\w\W]*?<a\s*href="([^>]*?)"[^>]*?>',block)
				
				if searchitems:
					denomination = searchitems.group(2)
					date = searchitems.group(3)
					nif = searchitems.group(4)
					origin = searchitems.group(5)
					detailUrl = searchitems.group(6)
					
					while True:
						try:
							with open('D:\Projects\CedarPython\ADIP-AO2202\Counts\ADIP-AO2202_Count.txt','a') as fh:
								fh.write('1\n')
							break
						except:
							if try_count>5:
								with open('Log_File.txt','a',encoding='utf-8') as logfile:
									logfile.write(traceback.format_exc())
								break
							time.sleep(1)
							try_count+=1
						
					search_page = [denomination,date,nif,origin]
					with open('D:\Projects\CedarPython\ADIP-AO2202\OPtxt\ADIP-AO2202_Searchpage.txt',"a",encoding='utf8')as f:
						f.write("\t".join(map(str,search_page))+"\n")
					worksheet1.write_string(row1, 0, denomination)
					worksheet1.write_string(row1, 0 + 1, date)
					worksheet1.write_string(row1, 0 + 2, nif)
					worksheet1.write_string(row1, 0 + 3, origin)
					row1 +=1
					detailpageCollector(detailUrl,nif)
		nextpageUrl = regex_match('href="([^>]*?)"\s*rel="next"',content)
		
		if nextpageUrl:
			obj = sess.get(nextpageUrl)
			nextpagecontent = html.unescape(obj.text)
			pagenumber = regex_match('page=([\d]+)$',nextpageUrl)
			with open(f'{cachePath}searchPage{pagenumber}.html','w',encoding='utf8') as fh:
				fh.write(nextpagecontent)
			searchPageCollector(nextpagecontent)
				
	except Exception as e:
		error = traceback.format_exc()
		print(error)
if __name__ == "__main__":	
	File_paths= ['D:\Projects\CedarPython\ADIP-AO2202\OP\ADIP-AO2202_Searchpage.xlsx','D:\Projects\CedarPython\ADIP-AO2202\OPtxt\ADIP-AO2202_Searchpage.txt','D:\Projects\CedarPython\ADIP-AO2202\OP\ADIP-AO2202_Companyinfo.xlsx','D:\Projects\CedarPython\ADIP-AO2202\OPtxt\ADIP-AO2202_Companyinfo.txt','D:\Projects\CedarPython\ADIP-AO2202\Error\ADIP-AO2202_Error.xlsx','D:\Projects\CedarPython\ADIP-AO2202\Counts\ADIP-AO2202_Count.txt']
	cachePath = 'Cache_AO2202/'
	if not os.path.isdir(cachePath):
		os.makedirs(cachePath)
	
	with open('D:\Projects\CedarPython\ADIP-AO2202\Counts\ADIP-AO2202_Count.txt',"w")as f:
			f.write("")	
	
	##############Output in Text#############
	search_page = ['Firm/Denomination','Date','NIF','Origin','Action']
	with open('D:\Projects\CedarPython\ADIP-AO2202\OPtxt\ADIP-AO2202_Searchpage.txt',"a")as f:
		f.write("\t".join(search_page)+"\n")
	
	row1 = 1
				
	workbook1 = xlsxwriter.Workbook('D:\Projects\CedarPython\ADIP-AO2202\OP\ADIP-AO2202_Searchpage.xlsx')
	worksheet1 = workbook1.add_worksheet()
	bold = workbook1.add_format({'bold': 1})
	worksheet1.write('A1', 'Firm/Denomination', bold)
	worksheet1.write('B1', 'Date', bold)
	worksheet1.write('C1', 'NIF', bold)
	worksheet1.write('D1', 'Origin', bold)
	worksheet1.write('E1', 'Action', bold)
		
	detail_page = ['Access code','Company number','Firm','NIF','Authority representitive - title','Authority representitive - name','Date','Folio registry number','Legal Form','Address of the headquarter','Activity description','Capital','Currency','Shareholders and quotas','Shareholder name','Shareholder information','Quota','Management description','Managers','Contributor no.','Way to oblige']
	with open('D:\Projects\CedarPython\ADIP-AO2202\OPtxt\ADIP-AO2202_Companyinfo.txt',"a")as f:
		f.write("\t".join(detail_page)+"\n")
	
	row2 = 1
				
	workbook2 = xlsxwriter.Workbook('D:\Projects\CedarPython\ADIP-AO2202\OP\ADIP-AO2202_Companyinfo.xlsx')
	worksheet2 = workbook2.add_worksheet()
	bold = workbook1.add_format({'bold': 1})
	worksheet2.write('A1', 'Access code', bold)
	worksheet2.write('B1', 'Company number', bold)
	worksheet2.write('C1', 'Firm', bold)
	worksheet2.write('D1', 'NIF', bold)
	worksheet2.write('E1', 'Authority representitive - title', bold)
	worksheet2.write('F1', 'Authority representitive - name', bold)
	worksheet2.write('G1', 'Date', bold)
	worksheet2.write('H1', 'Folio registry number', bold)
	worksheet2.write('I1', 'Legal Form', bold)
	worksheet2.write('J1', 'Address of the headquarter', bold)
	worksheet2.write('K1', 'Activity description', bold)
	worksheet2.write('L1', 'Capital', bold)
	worksheet2.write('M1', 'Currency', bold)
	worksheet2.write('N1', 'Shareholders and quotas', bold)
	worksheet2.write('O1', 'Shareholder name', bold)
	worksheet2.write('P1', 'Shareholder information', bold)
	worksheet2.write('Q1', 'Quota', bold)
	worksheet2.write('R1', 'Management description', bold)
	worksheet2.write('S1', 'Managers', bold)
	worksheet2.write('T1', 'Contributor no.', bold)
	worksheet2.write('U1', 'Way to oblige', bold)
	###############################################
	
	row2=1
	workbook2 = xlsxwriter.Workbook('D:\Projects\CedarPython\ADIP-AO2202\Error\ADIP-AO2202_Error.xlsx')
	worksheet2 = workbook2.add_worksheet()
	bold = workbook2.add_format({'bold': 1})
	worksheet2.write('A1', 'URL', bold)
	worksheet2.write('B1', 'Responding status', bold)
	worksheet2.write('C1', 'Error', bold)
	try:
		sess = requests.Session()
		sess.headers['User-Agent']='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36'
		mainUrl = 'https://gue.gov.ao/portal/publicacao?empresa=a%25a%25'
		obj = sess.get(mainUrl)
		searchcontent = html.unescape(obj.text)
		with open(f'{cachePath}searchPage1.html','w',encoding='utf8') as fh:
			fh.write(searchcontent)
		searchPageCollector(searchcontent) 
	except Exception as e:
		error = traceback.format_exc()
		print(error)
		worksheet2.write_string(row2, 0, 'https://gue.gov.ao/portal/publicacao?empresa=a%25a%25')
		worksheet2.write_string(row2, 0 + 1, "Not_responding")
		worksheet2.write_string(row2, 0 + 13, "TimeoutError")
		row2 +=1
		# workbook2.close()
	finally:
		workbook1.close()
		workbook2.close()
		# if os.path.isfile('ADIP-AO2202_Processed_Input.txt'):
			# os.remove('ADIP-AO2202_Processed_Input.txt')
database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
	for File_path in File_paths:
		delete_task(conn, File_path)