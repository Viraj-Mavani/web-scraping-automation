import os
import requests
import re
from requests_toolbelt.multipart.encoder import MultipartEncoder
import traceback
import sqlite3
from sqlite3 import Error
import pandas as pd
import sys
import time
import random
from bs4 import BeautifulSoup
import warnings
import json
import html
import csv
warnings.filterwarnings("ignore")


# BasePath = 'D:\\Projects\\CedarPython\\ADIP-AO2202\\'
BasePath = os.getcwd()
# BasePath= 'E:\\ADIP-PY\\OP2'

######### Excel #########
File_path_searchpage = BasePath + '\\OP\\ADIP-AO2202_Searchpage.xlsx'
File_path_companyinfo = BasePath + '\\OP\\ADIP-AO2202_Companyinfo.xlsx'
######### CSV #########
File_path_searchpage_CSV = BasePath + '\\OPcsv\\ADIP-AO2202_Searchpage.csv'
File_path_companyinfo_CSV = BasePath + '\\OPcsv\\ADIP-AO2202_Companyinfo.csv'
File_path_error_CSV = BasePath + '\\OPcsv\\ADIP-AO2202_Error.csv'
# File_path_failed_CSV = BasePath + '\\OP\\ADIP-AO2202_Failed.csv'
######### Text #########
File_path_searchpage_TXT = BasePath + '\\Optxt\\ADIP-AO2202_Searchpage.txt'
File_path_companyinfo_TXT = BasePath + '\\Optxt\\ADIP-AO2202_Companyinfo.txt'
######### Error #########
File_path_error = BasePath + '\\Error\\ADIP-AO2202_Error.xlsx'
######### Count #########
File_path_count = BasePath + '\\Counts\\ADIP-AO2202_Count.txt'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-AO2202_Log.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-AO2202_Run_Flag.txt'
File_path_log_index_page = BasePath + '\\Log\\ADIP-AO2202_Log_Page.txt'
File_path_log_index = BasePath + '\\Log\\ADIP-AO2202_Log_Index.txt'


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


def log_print(message):
	with open(File_path_log, 'a', encoding='utf-8') as file:
		file.write(message + '\n')
		file.flush()
	print(message)


def exception(url):
    Headers_Error = ['URL', 'Not Responding', 'Error']
    error = traceback.format_exc()
    exception_type, exception_object, exception_traceback = sys.exc_info()
    with open(File_path_error_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(Headers_Error)
        writer.writerow([url, "Not Responding", error])
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


# def restart_script():
#     python = sys.executable
#     subprocess.call([python] + sys.argv)
    

# def duplicate(File_path_EXL):
#     try:
#         data = pd.read_excel(File_path_EXL)
#         data_file = data.drop_duplicates()
#         data_file.to_excel(File_path_EXL, index=False)
#     except:
#         pass


def duplicateFromCSV(Csv_File_path):
	try:
		data = pd.read_csv(Csv_File_path)
		unique_data = data.drop_duplicates()
		unique_data.to_csv(Csv_File_path, index=False)
	except:
		pass


def convertCSVExcelExtended(File_path_CSV, File_path_EXL):
	chunk_size = 1000000 

	try:
		df = pd.read_csv(File_path_CSV, encoding='utf-8')
		df.to_excel(File_path_EXL, index=False)
		return None
	except (pd.errors.ParserError, pd.errors.EmptyDataError, ValueError):
		pass 
	except:
		exception(Base_URL)

	csv_reader = pd.read_csv(File_path_CSV, encoding='utf-8', chunksize=chunk_size)
	sheet_index = 1 
	# excel_files = []

	for chunk in csv_reader:
		if len(chunk) > 0:
			sheet_name = f'DataSet {sheet_index}'
			excel_file = f'{File_path_EXL[:-5]}_{sheet_index}.xlsx'
			chunk.to_excel(excel_file, sheet_name=sheet_name, index=False)
			# excel_files.append(excel_file)
			sheet_index += 1

	# # Merge all Excel files into one
	# writer = pd.ExcelWriter(File_path_EXL, engine='xlsxwriter')
	# Sheet = 1
	# for file in excel_files:
	#     df = pd.read_excel(file)
	#     sheet_name = f'DataSet {Sheet}'
	#     df.to_excel(writer, sheet_name=sheet_name, index=False)
	#     Sheet += 1
	# writer.close()


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


def request(req):
	# user_agent = random.choice(user_agents)
	# headers = {'User-Agent': user_agent}
	try:
		Retry = 1
		while Retry <= retry_attempts:
			try:
				sess = requests.Session()
				sess.headers['User-Agent']='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/111.0.0.0 Safari/537.36'
				obj = sess.get(req)
				searchcontent = html.unescape(obj.text)
				return searchcontent
			except Exception:
				exception(req)
				log_print(f"Error occurred in Request")
				delay = retry_delay * (2 ** Retry)
				log_print(f'Retrying in {delay} seconds...{Retry}')
				time.sleep(delay)
				Retry += 1
				continue
		else:
			log_print('\n\Requests Failed!!\nTerminating the script...\n===========================================================')
			os._exit(1)
			# log_print('\n\Request Failed!!\nRestarting the script in 5 min...\n===========================================================')
			# time.sleep(300)
			# restart_script()
	except:
		exception(req)


# def request(req):
# 	user_agent = random.choice(user_agents)
# 	headers = {'User-Agent': user_agent}
# 	try:
# 		Retry = 1
# 		while Retry <= retry_attempts:
# 			try:
# 				r_delay = random.uniform(1.5, 3.0)
# 				time.sleep(r_delay)
# 				# obj = requests.get(req, timeout=200)
# 				obj = requests.get(req, headers, timeout=200)
# 				soup = BeautifulSoup(obj.content, 'html.parser')
# 				return soup
# 			except Exception:
# 				exception(req)
# 				log_print(f"Error occurred in Request")
# 				delay = retry_delay * (2 ** Retry)
# 				log_print(f'Retrying in {delay} seconds...{Retry}')
# 				time.sleep(delay)
# 				Retry += 1
# 				continue
# 		else:
# 			log_print('\n\Requests Failed!!\nTerminating the script...\n===========================================================')
# 			os._exit(1)
# 			# log_print('\n\Request Failed!!\nRestarting the script in 5 min...\n===========================================================')
# 			# time.sleep(300)
# 			# restart_script()
# 	except:
# 		exception(req)


def extract_socios_quotas(content):
    result = []

    # Match the content inside the <strong>SÓCIOS E QUOTAS:</strong> tag
    match = re.search(r'<strong>SÓCIOS E QUOTAS:</strong>\s*(.*?)<\/p>', content)
    
    if match:
        socios_quotas_content = match.group(1)

        # Use a modified regular expression to match each shareholder
        shareholder_matches = re.finditer(r'<strong>(.*?)<\/strong>,\s*(.*?)(?:<br>|$)', socios_quotas_content)

        for shareholder_match in shareholder_matches:
            name = shareholder_match.group(1).strip()
            info = shareholder_match.group(2).strip()

            # Extract quota and currency information
            quota_match = re.search(r'(\d[\d\s\,\.]+)\s*(?:\(.*?"(.*?)"\))?', info)

            if quota_match:
                quota = quota_match.group(1).replace(',', '').strip()
                currency = quota_match.group(2).strip() if quota_match.group(2) else ""

                result.append({
                    'Shareholder Name': name,
                    'Quota': quota,
                    'Currency': currency
                })

    return result


def romoveDotAfter(data_str):
    if data_str != "" and data_str[-1] == ".":
        return data_str[0:-1]
    else:
        return data_str
    
    
def romoveDotBefore(data_str):
    if  data_str != "" and data_str[0] == ".":
        return data_str[1:]
    else:
        return data_str


def individual_data(NIF, action_sp, indi_url):
	try:
		details_section = action_sp.find('section', id='cabecalho')
		if details_section == None:
			log_print(f"{NIF} has No Data / Server Error: 500")
			return
		section_tables = details_section.find_all('table')
		rowsT1 = section_tables[0].find('tbody').find_all('tr')[1].find('td')
		rowsT2 = section_tables[1].find('tbody').find_all('tr')[1].find('td')
		all_p_tags = rowsT2.find_all('p')

		indi_data = ['']*15

		if rowsT1:
			data_pattern = rowsT1.get_text(strip=True)

			matricula_match = re.search(r'Matrícula[:\s]+\s*(.*?)\s*Firma', data_pattern)
			matricula = matricula_match.group(1) if matricula_match else ""
			indi_data[0] = romoveDotAfter(matricula)
   
			# if indi_data[0] == "3.214-20/201021":
			# 	pass

			firma_match = re.search(r'Firma[:\s]+\s*(.*?)\s*NIF+.*?\s*(\d+)', data_pattern)
			indi_data[1] = firma_match.group(1) if firma_match else ""
			indi_data[2] = firma_match.group(2) if firma_match else ""

			representative_match = re.search(r'NIF+.*?\s*\d+\s*(.*)', data_pattern)
			if representative_match:
				representative = representative_match.group(1).strip() if representative_match else ""
				indi_data[3] = romoveDotBefore(representative)
				# representatives = representative.split(',')
				# indi_data[3] = representatives[0].strip()
				# indi_data[4] = representatives[1].strip()
				# indi_data[5] = representatives[2].strip()
		if rowsT2:
			rowsT2content = str(rowsT2)
   
			socios_soup = gerencia_soup = obrigar_soup = None
			contribuinter_soups = []
			for p_tag in all_p_tags:
				capital_soup_value = p_tag.text.strip()
				capital_soup_match = re.search(r'(CAPITAL\s*:)', capital_soup_value)
				if "SEDE" in p_tag.get_text():
					sede_soup = p_tag
				elif "OBJECTO" in p_tag.get_text():
					obj_soup = p_tag
				elif (capital_soup_match) and (capital_soup_match.group(1) in p_tag.get_text()):
					capital_soup = p_tag
				elif "SÓCIOS" in p_tag.get_text() or "SÓCIO" in p_tag.get_text():
					socios_soup = p_tag
				elif "GERÊNCIA" in p_tag.get_text():
					gerencia_soup = p_tag
				elif "Contribuinte nº" in p_tag.get_text():
					contribuinter_soups.append(p_tag)
				elif "FORMA DE OBRIGAR" in p_tag.get_text():
					obrigar_soup = p_tag
   
			folio_match = re.search(r'Insc.\s*(.*?) -\s*<strong>(.*?)<\/strong>', rowsT2content)
			indi_data[4] = folio_match.group(1).strip() if folio_match else ""
			indi_data[5] = folio_match.group(2).strip() if folio_match else ""

			if sede_soup:
				sede_contents = sede_soup.text.strip()

				# if len(sede_contents)==2:
				# 	for content in sede_contents:
				# 		if "<strong" in str(content):
				# 			if content.next_sibling.text != None:
				# 				indi_data[8] = content.next_sibling.strip()
				# else:
				# 	pass
			sede_match = re.search(r'SEDE:\s*(.*)', sede_contents)
			indi_data[6] = sede_match.group(1).strip() if sede_match else ""

			if obj_soup:
				obj_contents = obj_soup.text.strip()

				# if len(obj_contents)==2:
				# 	for content in obj_contents:
				# 		if "<strong" in str(content):
				# 			if content.next_sibling.text != None:
				# 				indi_data[9] = content.next_sibling.strip()
				# else:
				# 	pass
			objecto_match = re.search(r'OBJECTO:\s*(.*)', obj_contents)
			indi_data[7] = objecto_match.group(1).strip() if objecto_match else ""

			if capital_soup:
				capital_contents = capital_soup.contents

				# if len(capital_contents)==2:
				# 	for content in capital_contents:
				# 		if "<strong" in str(content):
				# 			if content.next_sibling.text != None:
				# 				capital_value = content.next_sibling.strip()
				# 				capital_value_match = re.search(r'[\'"]+([A-Za-z])+[\'"]?\s*([\d\.,]+)', capital_value)
        
				if len(capital_contents)>1:
					if capital_soup.text != None:
						capital_code_value = capital_soup.text.strip()

						# Extract currency code
						currency_code_match = re.search(r'[\'"]*\b([A-Za-z]{2})\b[\'"]*', capital_code_value)
						# Extract numerical value with dots and commas, starting with a number
						value_match = re.search(r'\b(\d[\s\d\s\.\s,\s]*)\b', capital_code_value)

						indi_data[8] = value_match.group(1).strip() if value_match else ""
						indi_data[9] = currency_code_match.group(1).strip() if currency_code_match else ""

				else:
					pass

			# capital_match = re.search(r'CAPITAL:\s*<\/strong>(.*?)<\/p>', rowsT2content)
			# if capital_match:
			# 	capital_value_match = re.search(r'([A-Za-z]+)\.?([\d\s\,\.]+)', capital_match.group(1))

			# shareholdernumbers = []
			# shareholdernames = []
			# shareholderdesc = []
			# shareholdervalues = []
 
			if socios_soup:
				socios_contents = socios_soup.text.strip()
    
				# socios_contents = socios_soup.contents[1:]
				# current_number = current_name = current_description = None
				# current_quota_value = ""
				# for content in socios_contents:
					# if "<strong" in str(content):
					# 	if ";" in content.previous_sibling.strip():
					# 		current_number_match = re.search(r';\s*(.*)', content.previous_sibling.strip())
					# 		current_number = current_number_match.group(1)
					# 	else:
					# 		current_number = content.previous_sibling.strip() if content.previous_sibling else ""
					# 	current_name = content.text.strip()
					# 	if content.next_sibling.text != None:
					# 		full_description = content.next_sibling.strip()
					# 		full_description = full_description[2:] if full_description[:2]==", " else full_description
					# 	if "com uma quota" in full_description:
					# 		current_description_match = re.search(r'\s*(.*?), com uma quota', full_description)
					# 		quota_match = re.search(r'quota no valor nominal de\s*de\s*[A-Za-z]+ ([\d\.,]+)', full_description)
					# 		current_quota_value = quota_match.group(1) if quota_match else ""
					# 	else:
					# 		current_description_match = re.search(r'\s*(.*?);', full_description)
					# 	current_description = current_description_match.group(1) if current_description_match else ""
      
				# shareholdernumbers.append(current_number)
				# shareholdernames.append(current_name)
				# shareholderdesc.append(current_description)
				# shareholdervalues.append(current_quota_value)
				# indi_data[12] = ("|".join(shareholdernumbers))
				# indi_data[13] = ("|".join(shareholdernames))
				# indi_data[14] = ("|".join(shareholderdesc))
				# indi_data[15] = ("|".join(shareholdervalues))
    
				socios_match = re.search(r'SÓCIOS? E QUOTAS?\s*:\s*(.*)', socios_contents)
				indi_data[10] = socios_match.group(1).strip() if socios_match else ""
      
      

			contribuinter_name = []
			contribuinter_num = []
   
			if gerencia_soup:
				gerencia_contents = gerencia_soup.contents

				if len(gerencia_contents)==2:
					gerencia_txt = gerencia_soup.text.strip() if gerencia_soup.text else ''
					gerencia_match = re.search(r'GERÊNCIA:\s*(.*)', gerencia_txt)
					indi_data[11] = gerencia_match.group(1).strip() if gerencia_match else ""
					# indi_data[17] = content.text.strip()
				elif len(gerencia_contents)>2:
					for content in gerencia_contents:
						if "<strong>GERÊNCIA" in str(content):
							continue
						elif "<strong" not in str(content) and "Contribuinte nº" not in str(content):
							indi_data[11] = content.text.strip()
						elif "<strong" in str(content):
							contribuinter_name.append(content.text.strip())
						elif "Contribuinte nº" in str(content):
							contribuinter_num_match = re.search(r'Contribuinte nº ([A-Za-z0-9]+)', content.text.strip())
							contribuinter_num.append(contribuinter_num_match.group(1) if contribuinter_num_match else "")
					indi_data[12] = ("|".join(contribuinter_name))
					indi_data[13] = ("|".join(contribuinter_num))
				else:
					pass

			if len(contribuinter_soups)>0:
				for contribuinter_soup in contribuinter_soups:
					contribuinter_contents = contribuinter_soup.contents
					if len(contribuinter_contents)==2:
						for content in contribuinter_contents:
							if "<strong" in str(content):
								contribuinter_name.append(content.text.strip())
							else:
								contribuinter_num_match = re.search(r'Contribuinte nº ([A-Za-z0-9]+)', content.text.strip())
								contribuinter_num.append(contribuinter_num_match.group(1) if contribuinter_num_match else "")
					else:
						pass
				indi_data[12] = ("|".join(contribuinter_name))
				indi_data[13] = ("|".join(contribuinter_num))

			if obrigar_soup:
				obrigar_contents = obrigar_soup.contents

				if len(obrigar_contents)>1:
					for content in obrigar_contents:
						if "<strong" in str(content):
							if content.next_sibling.text != None:
								indi_data[14] = content.next_sibling.text.strip()
								break
				else:
					pass


		# indi_data = [matricula,firma,nif,re_title,re_name,re_signon,folioregistryNumber,legalForm,address,activity,capital,currency,
        #     shareholdernum,shareholdernam,shareholderdescription,shareholderval,managementDes,managernom,managerno,formadeobrigar]
		# Write to CSV file
		with open(File_path_companyinfo_CSV, 'a', newline='', encoding='utf-8') as file:
			writer = csv.writer(file)
			writer.writerow(indi_data)
		count()
		# Write to TXT file
		with open(File_path_companyinfo_TXT, 'a', encoding="utf-8") as fw:
			fw.write("\t".join(map(str, indi_data)) + "\n")
			fw.flush()
		log_print(f"Added {NIF}")
	except:
		exception(indi_url)


def search_data(search_sp):
	try:
		searchBlock = re.findall('<tr[^>]*?>([\w\W]*?)<\/tr>',searchcontent)		
		indi_data = []
		
		if searchBlock:
			for block in searchBlock:
				try:
					searchitems = re.search('<td[^>]*?>\s*([^>]*?)\s*<\/td>\s*<td[^>]*?>\s*([^>]*?)\s*<br>\s*([^>]*?)\s*<\/td>\s*<td[^>]*?>\s*\:?\s*([^>]*?)\s*<\/td>\s*<td[^>]*?>\s*([^>]*?)\s*<\/td>[\w\W]*?<a\s*href="([^>]*?)"[^>]*?>',block)
					
					if searchitems:
						Firm = searchitems.group(2)
						date = searchitems.group(3).replace('(',"").replace(')',"")
						nif = searchitems.group(4)
						Origin = searchitems.group(5)
						Action_URL = searchitems.group(6)
					else:
						log_print("Not enough data!!")
					
					indi_data = [Firm, date, nif, Origin, Action_URL]
					# Write to CSV file
					with open(File_path_searchpage_CSV, 'a', newline='', encoding='utf-8') as file:
						writer = csv.writer(file)
						writer.writerow(indi_data)
					count()
					# Write to TXT file
					with open(File_path_searchpage_TXT, 'a', encoding="utf-8") as fw:
						fw.write("\t".join(map(str, indi_data)) + "\n")
						fw.flush()

					Action_URL = "https://gue.gov.ao/portal/publicacao/ver/poIJE"
					ActionContent = request(Action_URL)
					ActionSoup = BeautifulSoup(ActionContent, 'html.parser')
					individual_data(nif, ActionSoup, Action_URL)
		
				except:
					exception(search_page_URL)
		with open(File_path_log_index_page, 'w', encoding='utf-8') as file:
			file.write(str(page))
			file.flush()
		log_print(f"Page {page} Compeleted")
		
	except:
		exception(search_page_URL)

  
if __name__ == "__main__":
	File_paths = [File_path_searchpage_CSV,File_path_companyinfo_CSV,File_path_searchpage_TXT,File_path_companyinfo_TXT, File_path_error_CSV]
	file_paths_logs = [File_path_log, File_path_log_index_page]

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

	# First_run = True
	# if First_run:
	if not os.path.exists(File_path_log_Run_Flag):
		with open(File_path_log_Run_Flag, "a", encoding='utf-8')as f:
			f.write("")
		for path_log in file_paths_logs:
			if os.path.exists(path_log):
				os.remove(path_log)
		for path_files in File_paths:
			if os.path.exists(path_files):
				os.remove(path_files)
		if os.path.exists(File_path_count):
			os.remove(File_path_count)

	with open(File_path_count, "a") as f:
		f.write("")

	HeadersF1 = ['Firm/Denomination','Date','NIF','Origin','Action']
	with open(File_path_searchpage_TXT, "a") as fw:
		if fw.tell() == 0:
			fw.write("\t".join(HeadersF1) + "\n")
			fw.flush()
	with open(File_path_searchpage_CSV, 'a', newline='', encoding='utf-8') as file:
		writer = csv.writer(file)
		if file.tell() == 0:
			writer.writerow(HeadersF1)
 
	HeadersF2 = ['Company number','Firm','NIF','Authority representitive','Folio registry number','Legal Form','Address of the headquarter','Activity description','Capital','Currency','Shareholders and quotas','Management description','Managers','Contributor no.','Way to oblige']
	with open(File_path_companyinfo_TXT, "a") as fw:
		if fw.tell() == 0:
			fw.write("\t".join(HeadersF2) + "\n")
			fw.flush()
	with open(File_path_companyinfo_CSV, 'a', newline='', encoding='utf-8') as file:
		writer = csv.writer(file)
		if file.tell() == 0:
			writer.writerow(HeadersF2)
 
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
	headers = {'User-Agent': user_agent}

	Base_URL = 'https://gue.gov.ao/portal/publicacao?empresa=a%25a%25'
	Page_URL = Base_URL + '&page={}'
	# company_URL = 'https://gue.gov.ao/portal/publicacao/ver/{}'
	st = time.time()

	retry_attempts = 5
	retry_delay = 2

	try:
		searchcontent_temp = request(Base_URL)
  
		search_soup_temp = BeautifulSoup(searchcontent_temp, 'html.parser')	
		pagination = search_soup_temp.find('ul', class_='pagination')
		total_pages = int(pagination.find_all('li', class_='page-item')[-2].get_text(strip=True))
		search_soup_temp.decompose()
		
		log_index_flag = False
		if os.path.exists(File_path_log_index_page):
			log_index_flag = True
			with open(File_path_log_index_page, 'r', encoding='utf-8') as file:
				try:
					last_processed_page = int(file.read().strip())
				except:
					last_processed_page = ''

		if log_index_flag and last_processed_page != '':
			start_index = last_processed_page + 1
		else:
			start_index = 1
		
		for page in range(start_index, total_pages+1)[:]:
			try:
				search_page_URL = Page_URL.format(page)
				searchcontent = request(search_page_URL)
				search_soup = BeautifulSoup(searchcontent, 'html.parser')			
    
				search_data(search_soup)
				
			except:
				exception(Page_URL.format(page))
    
		duplicateFromCSV(File_path_searchpage_CSV)
		duplicateFromCSV(File_path_companyinfo_CSV)
		convertCSVExcelExtended(File_path_searchpage_CSV, File_path_searchpage)
		convertCSVExcelExtended(File_path_companyinfo_CSV, File_path_companyinfo)
		et = time.time()
		log_print(f'\n{et - st}')
		if os.path.exists(File_path_count):
			os.remove(File_path_count)
		exit()
	except:
		exception(Base_URL)

# database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# # create a database connection
# conn = create_connection(database)
# with conn:
# 	for File_path in File_paths:
# 		delete_task(conn, File_path)