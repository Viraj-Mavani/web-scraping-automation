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

BasePath = 'D:\Projects\CedarPython\ADIP-SY603'

File_path= BasePath +'\OP\ADIP-SY603_Output.xlsx'
File_path_txt= BasePath +'\OPtxt\ADIP-SY603_Output.txt'
File_path_count= BasePath +'\Counts\ADIP-SY603_Count.txt'
File_path_error= BasePath +'\Error\ADIP-SY603_Error.xlsx'

persian_alphabet = [
    "ا", "ب","پ","ت","ث","ج","چ","ح","خ","د","ذ","ر","ز","ژ","س","ش",
	"ص","ض","ط","ظ","ع","غ","ف","ق","ک","گ","ل","م","ن","و","ه","ی"
]


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

def Individual_data(data):
	# book1 = openpyxl.load_workbook(File_path)
	# sheet1 = book1.active
	global row1, rowError
	# unique_rows = set()

	for item in data:
		Indi_data = ['']*4
		Name_tag = item.find('i', {"class": ['fa', 'fa-user']})
		name = Name_tag.next if Name_tag.next else ''
		Body_data = item.find_all('li')

		# if len(Body_data) < 3:  # Check if Body_data has at least 3 elements
		# 	continue

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

			# Create a tuple with all column values
			Indi_data.append(name)
			Indi_data.append(phone)
			Indi_data.append(address)
			Indi_data.append(activity)

			# Check if the row data is already in the set
			# if row_data not in unique_rows:
			# sheet1['A{col}'.format(col=row1)] = name
			# sheet1['B{col}'.format(col=row1)] = phone
			# sheet1['C{col}'.format(col=row1)] = address
			# sheet1['D{col}'.format(col=row1)] = activity
	
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
			error = traceback.format_exc()
			exception_type, exception_object, exception_traceback = sys.exc_info()
			# book2 = openpyxl.load_workbook(File_path_error)
			# sheet2 = book2.active
			worksheet_error.write('A{}'.format(rowError), Base_url)
			worksheet_error.write('B{}'.format(rowError), "Not Responding")
			worksheet_error.write('C{}'.format(rowError), error)
			# book2.save(File_path_error)
			# book2.close()
			rowError+=1

	# book1.save(File_path)
	# book1.close()		

if __name__=='__main__':
	
	row1=2
	rowError=2
	
	book1 = xlsxwriter.Workbook(File_path)
	# book1 = openpyxl.Workbook()
	sheet1 = book1.add_worksheet()
	bold_format = book1.add_format({'bold': True})
	# sheet1 = book1.active
	sheet1.write('A1', 'Company Name', bold_format)
	sheet1.write('B1', 'Phone!', bold_format)
	sheet1.write('C1', 'Address', bold_format)
	sheet1.write('D1', 'Acitivity', bold_format)
	# sheet1['A1'] = 'Company Name'
	# sheet1['A1'].font = Font(bold=True)
	# sheet1['B1'] = 'Phone'
	# sheet1['B1'].font = Font(bold=True)
	# sheet1['C1'] = 'Address'
	# sheet1['C1'].font = Font(bold=True)
	# sheet1['D1'] = 'Acitivity'
	# sheet1['D1'].font = Font(bold=True)
	# book1.save(File_path)
	# book1.close()
	
	workbook_error = xlsxwriter.Workbook(File_path_error)
	worksheet_error = workbook_error.add_worksheet()

	bold_format = workbook_error.add_format({'bold': True})

	worksheet_error.write('A1', 'URL', bold_format)
	worksheet_error.write('B1', 'Not Responding', bold_format)
	worksheet_error.write('C1', 'Error', bold_format)
	# book2 = openpyxl.Workbook()
	# sheet2 = book2.active
	# sheet2['A1'] = 'URL'
	# sheet2['A1'].font = Font(bold=True)
	# sheet2['B1'] = 'Not Responding'
	# sheet2['B1'].font = Font(bold=True)
	# sheet2['C1'] = 'Error'
	# sheet2['C1'].font = Font(bold=True)
	# book2.save(File_path_error)
	# book2.close()

	Search_headers = ['Company Name','Phone','Adrress','Activity']
	with open(File_path_txt,"w")as f:
		f.write("\t".join(Search_headers)+"\n")
	with open(File_path_count,"w")as f:
		f.write("")

	# Get the second-to-last <a> element
	Base_url = 'http://hamachamber.com/members-index/?ftxt={}&ftxt2=&searchType=1&count={}'

	try:
		for letter in persian_alphabet:
			obj_temp = requests.get(Base_url.format(letter, 1))
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
				

			for index in range(1, last_page_number + 1):
				obj = requests.get(Base_url.format(letter, index), timeout=300)
				soup = BeautifulSoup(obj.content, 'html.parser')
				res = soup.find_all('div', class_='mycalls')[1:]
				Individual_data(res)
				print(f'Success {letter} {index}')
			print(f'{letter} Complete\n\n')
	except:
		error = traceback.format_exc()
		exception_type, exception_object, exception_traceback = sys.exc_info()
		# book2 = openpyxl.load_workbook(File_path_error)
		# sheet2 = book2.active
		worksheet_error.write(rowError ,0 , Base_url)
		worksheet_error.write(rowError ,1 , "Not Responding")
		worksheet_error.write(rowError ,2 , error)
		# book2.save(File_path_error)
		# book2.close()
		rowError+=1
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