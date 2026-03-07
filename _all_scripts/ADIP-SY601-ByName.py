import csv
import sys
from bs4 import BeautifulSoup
import requests
import re
import os
import time
import traceback
import pandas as pd
import xlsxwriter
from requests_toolbelt.multipart.encoder import MultipartEncoder


# BasePath = 'D:\Projects\CedarPython\ADIP-SY601-ByName'
BasePath = os.getcwd()
######### Excel #########
File_path = BasePath + '\\OP\\ADIP-SY601-ByName_Company_Info.xlsx'
######### CSV #########
File_path_CSV = BasePath + '\\OPcsv\ADIP-SY601-ByName_Output.csv'
File_path_error_CSV = BasePath + '\\OPcsv\ADIP-SY601-ByName_Error.csv'
######### Text #########
File_path_txt = BasePath + '\\OPtxt\ADIP-SY601-ByName_Output.txt'
######### Input #########
File_path_Input = BasePath + '\\InputFile\\ADIP-SY601-ByName-Input.xlsx'
######### Count #########
File_path_count = BasePath + '\\Counts\\ADIP-SY601-ByName_Count.txt'
######### Failed #########
File_path_failed_Arabic = BasePath + '\\OP\\ADIP-SY601-ByName_Failed_Arabic.xlsx'
File_path_failed_Arabic_CSV = BasePath + '\\OPcsv\\ADIP-SY601-ByName_Failed_Arabic.csv'
######### Error #########
File_path_error = BasePath + '\\Error\\ADIP-SY601-ByName_Error.xlsx'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-SY601-ByName_Log.txt'
File_path_log_index_Arabic = BasePath + '\\Log\\ADIP-SY601-ByName_Log_Index_Arabic.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-SY601-ByName_Run_Flag.txt'


def log_print(message):
    with open(File_path_log, 'a', encoding='utf-8') as file:
        file.write(message + '\n')
        file.flush()
    print(message)


def exception():
	headers = ['URL', 'Not Responding', 'Error']
	error = traceback.format_exc()
	exception_type, exception_object, exception_traceback = sys.exc_info()
	with open(File_path_error_CSV, 'a', newline='', encoding='utf-8') as file:
		writer = csv.writer(file)
		if file.tell() == 0:
			writer.writerow(headers)
		writer.writerow([BaseUrl, "Not Responding", error])
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


def content_read_from_file(filename):
	try:
		with open(filename,'r') as fh:
			content = fh.read()
		return content
	except:
		try:
			with open(filename,'rb') as fh:
				content = fh.read().decode('utf-8')
			return content
		except:
			try:
				with open(filename,'r') as fh:
					content = fh.read().decode('utf-8')
				return content
			except:
				with open('Error_Files.txt','a') as fh:
					fh.write('{}\n'.format(filename))
				return ''
				
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
	
def searchpage_Collector(Content):
	try:
		# global row1
		companyInfoBlocks = re.findall(r'<div\s*class="w3-panel\s*w3-card-2\s*w3-hover-light-grey\s*w3-padding-8">\s*([\w\W]*?)\s*<\/div>\s*<\/a>',Content)
		Indi_data = []
		for companyInfoBlock in companyInfoBlocks:
			FacilityName = attribute_replace(regex_match('<div[^>]*?>\s*اسم\s*المنشأة\s*\:\s*<\/div>\s*<div[^>]*?>\s*([^>]*?)\s*<\/div>',companyInfoBlock))
			RegisteredActivity = attribute_replace(regex_match('<div[^>]*?>\s*الصناعة\s*وفق\s*الأرومة\s*\:\s*<\/div>\s*<div[^>]*?>\s*([^>]*?)\s*<\/div>',companyInfoBlock))
			MainIndustry = attribute_replace(regex_match('<div[^>]*?>\s*الصناعة\s*الرئيسية\s*\:\s*<\/div>\s*<div[^>]*?>\s*([^>]*?)\s*<\/div>',companyInfoBlock))
			Mobile = attribute_replace(regex_match('<div[^>]*?>\s*الجوال\s*\:\s*<\/div>\s*<div[^>]*?>\s*([^>]*?)\s*<\/div>',companyInfoBlock))
			Phone = attribute_replace(regex_match('<div[^>]*?>\s*الهاتف\s*\:\s*<\/div>\s*<div[^>]*?>\s*([^>]*?)\s*<\/div>',companyInfoBlock))
			Email = attribute_replace(regex_match('<div[^>]*?>\s*البريد\s*الالكتروني\s*\:\s*<\/div>\s*<div[^>]*?>\s*([^>]*?)\s*<\/div>',companyInfoBlock))
			
			Indi_data = [FacilityName, RegisteredActivity, MainIndustry, Mobile, Phone, Email]

			with open(File_path_CSV, 'a', newline='', encoding='utf-8') as file:
				writer = csv.writer(file)
				writer.writerow(Indi_data)
			count()
			with open(File_path_txt, 'a', encoding="utf-8") as fw:
				fw.write("\t".join(map(str, Indi_data))+"\n")
				fw.flush()
			with open(File_path_log_index_Arabic, 'w', encoding='utf-8') as f3:
				f3.write(indexA1)
				f3.flush()
	except Exception as e:
		exception()

if __name__=='__main__':
    
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
   
	try:
		sess = requests.Session()
		sess.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.183 Safari/537.36'
		
		# Create directories if they don't exist

		File_path_headers = ['Name', 'Registered Activity', 'Main Industry', 'Mobile', 'Phone', 'Email']
		if not os.path.exists(File_path_count):
			with open(File_path_count, "a")as f:
				f.write("")
		if not os.path.exists(File_path_txt):
			with open(File_path_txt, "a")as f:
				f.write("\t".join(File_path_headers)+"\n")
				f.flush()
		with open(File_path_CSV, 'a', newline='', encoding='utf-8') as file:
			writer = csv.writer(file)
			if file.tell() == 0:
				writer.writerow(File_path_headers)
		# arabicWords = ["ي","و","ه","ن","م","ل","ك","ق","ف","غ","ع","ظ","ط","ض","ص","ش","س","ز","ر","ذ","د","خ","ح","ج","ث","ت","ب","ا"]
		# arabicWords = ["ب","ا"]

		BaseUrl = 'https://www.dci-syria.org/?tns=&search_key={}&industry='
  
		log_print('Data Importing...plz wait')
		df = pd.read_excel(File_path_Input, sheet_name='Sheet1')
		names_list = df['nameLocal'].tolist()
		log_print('Data Imported\n')
		
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

		for indexA1 in names_list[:]:
			error_message_flag = False
			# log_print("Getting Data for : " + indexA1)
			obj = sess.get(BaseUrl.format(indexA1))
			# with open('{}Listpage_{}.html'.format(cachePath,arabicLetters),'wb') as fh:
			# 	fh.write(obj.content)
			listpagecontent = obj.text
			soup = BeautifulSoup(obj.content, 'html.parser')

			search_number_element = soup.find(
			    'div', class_='block-search-detail').find("h2").find("span", class_='w3-text-red')
			if search_number_element.string.strip() == '(0)':
				error_message_flag = True
			soup.decompose()
			if error_message_flag:
				error_message_flag = False
				log_print(f'Data Not Found {indexA1}')
				with open(File_path_log_index_Arabic, 'w', encoding='utf-8') as f3:
					f3.write(indexA1)
					f3.flush()
				continue

			searchpage_Collector(listpagecontent)

			if os.path.exists(File_path_log_index_Arabic):
				with open(File_path_log_index_Arabic, 'r', encoding='utf-8') as file:
					last_processed_name = file.read().strip()
			if indexA1 == last_processed_name:
				log_print(f'Success {indexA1}')
			else:
				log_print(f'Failed!! {indexA1}')
				with open(File_path_failed_Arabic_CSV, 'a', newline='', encoding='utf-8') as file:
					writer = csv.writer(file)
					writer.writerow([indexA1])
		log_print('Script Completed')

	except:
		exception()
	finally:
		convertCSVExcel(File_path_CSV, File_path)
		duplicate(File_path)