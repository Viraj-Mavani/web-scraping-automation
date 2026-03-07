import requests
import re
import os
import time
import traceback
import xlsxwriter
from requests_toolbelt.multipart.encoder import MultipartEncoder

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
		global row1
		companyInfoBlocks = re.findall(r'<div\s*class="w3-panel\s*w3-card-2\s*w3-hover-light-grey\s*w3-padding-8">\s*([\w\W]*?)\s*<\/div>\s*<\/a>',Content)
		
		for companyInfoBlock in companyInfoBlocks:
			FacilityName = attribute_replace(regex_match('<div[^>]*?>\s*اسم\s*المنشأة\s*\:\s*<\/div>\s*<div[^>]*?>\s*([^>]*?)\s*<\/div>',companyInfoBlock))
			RegisteredActivity = attribute_replace(regex_match('<div[^>]*?>\s*الصناعة\s*وفق\s*الأرومة\s*\:\s*<\/div>\s*<div[^>]*?>\s*([^>]*?)\s*<\/div>',companyInfoBlock))
			MainIndustry = attribute_replace(regex_match('<div[^>]*?>\s*الصناعة\s*الرئيسية\s*\:\s*<\/div>\s*<div[^>]*?>\s*([^>]*?)\s*<\/div>',companyInfoBlock))
			Mobile = attribute_replace(regex_match('<div[^>]*?>\s*الجوال\s*\:\s*<\/div>\s*<div[^>]*?>\s*([^>]*?)\s*<\/div>',companyInfoBlock))
			Phone = attribute_replace(regex_match('<div[^>]*?>\s*الهاتف\s*\:\s*<\/div>\s*<div[^>]*?>\s*([^>]*?)\s*<\/div>',companyInfoBlock))
			Email = attribute_replace(regex_match('<div[^>]*?>\s*البريد\s*الالكتروني\s*\:\s*<\/div>\s*<div[^>]*?>\s*([^>]*?)\s*<\/div>',companyInfoBlock))
			
			worksheet1.write_string(row1, 0, FacilityName)
			worksheet1.write_string(row1, 0 + 1, RegisteredActivity)
			worksheet1.write_string(row1, 0 + 2, MainIndustry)
			worksheet1.write_string(row1, 0 + 3, Mobile)
			worksheet1.write_string(row1, 0 + 4, Phone)
			worksheet1.write_string(row1, 0 + 5, Email)
			row1 += 1
	except Exception as e:
		error = traceback.format_exc()
		print(error)

if __name__=='__main__':
	try:
		sess = requests.Session()
		sess.headers['User-Agent'] = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.183 Safari/537.36'
		
		cachePath = 'Cache/'
		if not os.path.isdir(cachePath):
			os.makedirs(cachePath)
		
		row1 = 1
		
		workbook1 = xlsxwriter.Workbook('D:\Projects\CedarPython\ADIP-SY601\OP\ADIP_SY601_Company_Info' + '.xlsx')
		worksheet1 = workbook1.add_worksheet()
		bold = workbook1.add_format({'bold': 1})
		worksheet1.write('A1', 'Name', bold)
		worksheet1.write('B1', 'Registered Activity', bold)
		worksheet1.write('C1', 'Main Industry', bold)
		worksheet1.write('D1', 'Mobile', bold)
		worksheet1.write('E1', 'Phone', bold)
		worksheet1.write('F1', 'Email', bold)
		
		arabicWords = ["ي","و","ه","ن","م","ل","ك","ق","ف","غ","ع","ظ","ط","ض","ص","ش","س","ز","ر","ذ","د","خ","ح","ج","ث","ت","ب","ا"]
		# arabicWords = ["ب","ا"]
		
		for i in arabicWords:
			for j in arabicWords:
				arabicLetters = i+j
				print("Getting Data for : " + arabicLetters)
				obj=sess.get('https://www.dci-syria.org/?tns=&search_key='+arabicLetters+'&industry=')
				with open('{}Listpage_{}.html'.format(cachePath,arabicLetters),'wb') as fh:
					fh.write(obj.content)
				listpagecontent = obj.text
				searchpage_Collector(listpagecontent)
	except Exception as e:
		error = traceback.format_exc()
		print(error)
	finally:
		workbook1.close()