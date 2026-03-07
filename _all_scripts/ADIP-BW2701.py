import csv
import os
import json
import random
import sqlite3
from sqlite3 import Error
import subprocess
import sys
import traceback
from bs4 import BeautifulSoup
import pandas as pd
import requests
import math
import time
import string
from selenium import webdriver
# import openpyxl
# from openpyxl.styles import Font
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options

chrome_options = Options()
chrome_options.add_argument('--headless')
chrome_options.add_argument('--no-sandbox')
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])


######### Excel #########
BasePath= 'D:\\Projects\\CedarPython\\ADIP-BW2701'
Total_URL = 0
File_path_General_Details= BasePath + '\\OP\\ADIP-BW2701_General_Details.xlsx'
File_path_Auditors= BasePath + '\\OP\\ADIP-BW2701_Auditors.xlsx'
File_path_Addresses= BasePath + '\\OP\\ADIP-BW2701_Addresses.xlsx'
File_path_Directors= BasePath + '\\OP\\ADIP-BW2701_Directors.xlsx'
File_path_Secretaries= BasePath + '\\OP\\ADIP-BW2701_Secretaries.xlsx'
File_path_Shareholders= BasePath + '\\OP\\ADIP-BW2701_Shareholders.xlsx'
File_path_Share_Allocations= BasePath + '\\OP\\ADIP-BW2701_Share_Allocations.xlsx'
File_path_Proprietors= BasePath + '\\OP\\ADIP-BW2701_Proprietors.xlsx'
File_path_Search_Page_Info= BasePath + '\\OP\\ADIP-BW2701_Search_Page_Info.xlsx'
######### Text #########
File_path_General_Details_txt= BasePath + '\\OPtxt\\ADIP-BW2701_General_Details.txt'
File_path_Auditors_txt= BasePath + '\\OPtxt\\ADIP-BW2701_Auditors.txt'
File_path_Addresses_txt= BasePath + '\\OPtxt\\ADIP-BW2701_Addresses.txt'
File_path_Directors_txt= BasePath + '\\OPtxt\\ADIP-BW2701_Directors.txt'
File_path_Secretaries_txt= BasePath + '\\OPtxt\\ADIP-BW2701_Secretaries.txt'
File_path_Shareholders_txt= BasePath + '\\OPtxt\\ADIP-BW2701_Shareholders.txt'
File_path_Share_Allocations_txt= BasePath + '\\OPtxt\\ADIP-BW2701_Share_Allocations.txt'
File_path_Proprietors_txt= BasePath + '\\OPtxt\\ADIP-BW2701_Proprietors.txt'
File_path_Search_Page_Info_txt= BasePath + '\\OPtxt\\ADIP-BW2701_Search_Page_Info.txt'
######### CSV #########
File_path_General_Details_csv= BasePath + '\\OPcsv\\ADIP-BW2701_General_Details.csv'
File_path_Auditors_csv= BasePath + '\\OPcsv\\ADIP-BW2701_Auditors.csv'
File_path_Addresses_csv= BasePath + '\\OPcsv\\ADIP-BW2701_Addresses.csv'
File_path_Directors_csv= BasePath + '\\OPcsv\\ADIP-BW2701_Directors.csv'
File_path_Secretaries_csv= BasePath + '\\OPcsv\\ADIP-BW2701_Secretaries.csv'
File_path_Shareholders_csv= BasePath + '\\OPcsv\\ADIP-BW2701_Shareholders.csv'
File_path_Share_Allocations_csv= BasePath + '\\OPcsv\\ADIP-BW2701_Share_Allocations.csv'
File_path_Proprietors_csv= BasePath + '\\OPcsv\\ADIP-BW2701_Proprietors.csv'
File_path_Search_Page_Info_csv= BasePath + '\\OPcsv\\ADIP-BW2701_Search_Page_Info.csv'
######### Error #########
File_path_error= BasePath + '\\Error\\ADIP-BW2701_Error.xlsx'
File_path_error_CSV= BasePath + '\\Error\\ADIP-BW2701_Error.csv'
######### Count #########
File_path_count= BasePath + '\\Counts\\ADIP-BW2701_Count.txt'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-BW2701_Log.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-BW2701_Run_Flag.txt'
File_path_log_index = BasePath + '\\Log\\ADIP-BW2701_Log_Index.txt'
File_path_log_index_LetterE1 = BasePath + '\\Log\\ADIP-BW2701_Log_Index_LetterE1.txt'
File_path_log_index_LetterE2 = BasePath + '\\Log\\ADIP-BW2701_Log_Index_LetterE2.txt'
File_path_log_index_LetterE3 = BasePath + '\\Log\\ADIP-BW2701_Log_Index_LetterE3.txt'


alphabet = list(string.ascii_lowercase)


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


def exception():
    Headers_Error = ['URL', 'Not Responding', 'Error']
    error = traceback.format_exc()
    exception_type, exception_object, exception_traceback = sys.exc_info()
    with open(File_path_error_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(Headers_Error)
        writer.writerow([New_SearchPage_Url, "Not Responding", error])
    df = pd.read_csv(File_path_error_CSV, encoding='utf-8')
    df.to_excel(File_path_error, index=False)


def Dereference(obj):
    del obj


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


# def convertCSVExcel(csv, exl):
#     df = pd.read_csv(csv, encoding='utf-8')
#     df.to_excel(exl, index=False)
    

# def duplicate(exl):
#     try:
#         data = pd.read_excel(exl)
#         data_file = data.drop_duplicates()
#         data_file.to_excel(exl, index=False)
#     except:
#         pass


def General_Details(ID,data,isCompany):
	General_Details_data = ['']*21
	Main_data = data[data[data[ID]['children'][0]]['children'][0]]['children']
	for i in Main_data:
		if 'attribute' in data[i] and data[i]['attribute'] == '_businessIdentifier':
			General_Details_data[9] = data[i]['attributeValue'] if 'attributeValue' in data[i] else ''
		elif 'attribute' in data[i] and data[i]['attribute'] == 'ForeignCompanyDocumentLodgedYn':
			General_Details_data[14] = data[i]['attributeValue'] if 'attributeValue' in data[i] else ''
		elif 'attribute' in data[i] and data[i]['attribute'] == 'Exempt':
			General_Details_data[15] = data[i]['text']['dcText'] if ('text' in data[i] and 'dcText' in data[i]['text']) else ''
		elif 'attribute' in data[i] and data[i]['attribute'] == 'RegistrationDate':
			if 'text' in data[i] and 'label' in data[i]['text'] and data[i]['text']['label'] == 'Incorporation Date':
				General_Details_data[16] = data[i]['attributeValue'] if 'attributeValue' in data[i] else ''
			if 'text' in data[i] and 'label' in data[i]['text'] and data[i]['text']['label'] == 'Registration Date':
				General_Details_data[4] = data[i]['attributeValue'] if 'attributeValue' in data[i] else ''			
		elif 'attribute' in data[i] and data[i]['attribute'] == 'ReregistrationDate':
			General_Details_data[17] = data[i]['attributeValue'] if 'attributeValue' in data[i] else ''
		elif 'attribute' in data[i] and data[i]['attribute'] == 'OwnConstitutionYn':
			General_Details_data[18] = data[i]['text']['dcText'] if ('text' in data[i] and 'dcText' in data[i]['text']) else ''
		elif 'attribute' in data[i] and data[i]['attribute'] == 'FilingMonth':
			General_Details_data[19] = data[i]['text']['dcText'] if ('text' in data[i] and 'dcText' in data[i]['text']) else ''
		elif 'attribute' in data[i] and data[i]['attribute'] == 'ARLastFiledDate':
			General_Details_data[20] = data[i]['attributeValue'] if 'attributeValue' in data[i] else ''
		elif 'text' in data[i] and 'label' in data[i]['text'] and data[i]['text']['label'] == 'Status History':
			Status_All = data[data[i]['children'][0]]['text']['row']
			Status = Status_All.split('(')[0].strip()
			Dates = Status_All.split('(')[1]
			Start = Dates.split('to')[0].strip()
			End = Dates.split('to')[1].strip()
			General_Details_data[11] = Status
			General_Details_data[12] = Start
			General_Details_data[13] = End[:-1]
		elif 'children' in data[i]:
			if 'domain' in data[data[i]['children'][0]] and data[data[i]['children'][0]]['domain'] == 'Status':
				if 'text' in data[data[i]['children'][0]] and 'row' in data[data[i]['children'][0]]['text']:
					if isCompany == 1:
						General_Details_data[10] = data[data[i]['children'][0]]['text']['row']
					else:
						General_Details_data[3] = data[data[i]['children'][0]]['text']['row']
		elif 'attribute' in data[i] and data[i]['attribute'] == 'RenewalMonth':
			General_Details_data[8] = data[i]['text']['dcText'] if ('text' in data[i] and 'dcText' in data[i]['text']) else ''
	for j in data:
		if 'DC_' in j:
			break
		else:
			if 'attribute' in data[j] and data[j]['attribute'] == '_ClassificationDescription':
				General_Details_data[6] = data[j]['attributeValue'] if 'attributeValue' in data[j] else ''
			elif 'attribute' in data[j] and data[j]['attribute'] == 'CommencementDate':
				General_Details_data[7] = data[j]['attributeValue'] if 'attributeValue' in data[j] else ''
			elif 'widget' in data[j] and data[j]['widget'] == 'attribute-value-list':
				if 'text' in data[j] and 'label' in data[j]['text'] and data[j]['text']['label'] == 'Name History':
					Name_History = data[data[j]['children'][0]]['text']['row']
					Name = ' '.join(Name_History.split(' ')[:-3])
					Date = Name_History.split(' ')
					StartDate = Date[-3]
					EndDate = Date[-1]
					General_Details_data[0] = Name
					General_Details_data[1] = StartDate
					General_Details_data[2] = EndDate[:-1]
	with open(File_path_General_Details_csv, 'a', newline='', encoding='utf-8') as file:
		writer = csv.writer(file)
		writer.writerow(General_Details_data)
	with open(File_path_General_Details_txt,"a", encoding='utf-8')as f:
		f.write("\t".join(map(str,General_Details_data))+"\n")
		f.flush()

def Addresses(data,UIN):
	Addresses_Data = ['']*7
	Addresses_Data[6] = UIN
	for i in data:
		if 'DC_' in i:
			break
		else:
			if 'domain' in data[i] and data[i]['domain'] == 'RegisteredOfficeAddress':
				if 'text' in data[i] and 'singleline' in data[i]['text']:
					Addresses_Data[0] = data[i]['text']['row'].strip()
			elif 'domain' in data[i] and data[i]['domain'] == 'EntityPostalAddress':
				if 'text' in data[i] and 'row' in data[i]['text']:
					Postal = data[i]['text']['row']
					if '(' in Postal:
						Addresses_Data[3] = Postal.strip()
					else:
						Addresses_Data[2] = Postal.strip()
			elif 'domain' in data[i] and data[i]['domain'] == 'PrincipalPlaceOfBusinessAddress':
				if 'text' in data[i] and 'singleline' in data[i]['text']:
					Addresses_Data[4] = data[i]['text']['row'].strip()
			elif 'widget' in data[i] and data[i]['widget'] == 'attribute-value-list':
				if 'text' in data[i] and 'label' in data[i]['text'] and data[i]['text']['label'] == 'Previous Principal Places of Business':
					Prev_postal = data[data[i]['children'][0]]['text']['row']
					Prev_postal = Prev_postal.split(' ')
					Addresses_Data[5] = ' '.join(Prev_postal[:-3])
				elif 'text' in data[i] and 'label' in data[i]['text'] and data[i]['text']['label'] == 'Previous Registered Office Addresses':
					Prev = data[data[i]['children'][0]]['text']['row']
					Prev = Prev.split(' ')
					Addresses_Data[1] = ' '.join(Prev[:-3])
	with open(File_path_Addresses_csv, 'a', newline='', encoding='utf-8') as file:
		writer = csv.writer(file)
		writer.writerow(Addresses_Data)
	with open(File_path_Addresses_txt,"a", encoding='utf-8')as f:
		f.write("\t".join(map(str,Addresses_Data))+"\n")

def Directors(data,UIN):
	Directors_Data = ['']*7
	collections = []
	for i in data:
		if 'DC_' in i:
			break
		else: 
			if 'domain' in data[i] and data[i]['domain'] == 'IndividualDirector':
				collections.append(i)
	for id in collections:
		Directors_Data[6] = UIN
		if 'text' in data[id] and 'singleline' in data[id]['text']:
			Directors_Data[0] = data[id]['text']['singleline']
		Detail_ID = data[id]['children'][1]
		Nationality_ID = data[Detail_ID]['children'][0]
		Nationality_ID = data[Nationality_ID]['children'][1]
		if 'attribute' in data[Nationality_ID] and 'text' in data[Nationality_ID] and data[Nationality_ID]['attribute'] == 'Nationality':
			Directors_Data[1] = data[Nationality_ID]['text']['dcText']
		Resident = data[Detail_ID]['children'][2]
		Resident = data[Resident]['children'][0]
		Resident = data[Resident]['children'][0]
		if 'domain' in data[Resident] and 'text' in data[Resident] and data[Resident]['domain'] == 'ResidentialAddress':
			Directors_Data[2] = data[Resident]['text']['row'].strip()
		Postal = data[Detail_ID]['children'][3]
		Postal = data[Postal]['children'][0]
		Postal = data[Postal]['children'][0]
		if 'domain' in data[Postal] and 'text' in data[Postal] and data[Postal]['domain'] == 'ServiceAddress':
			Directors_Data[3] = data[Postal]['text']['row'].strip()
		Appointment = data[Detail_ID]['children'][4]
		Appointment = data[Appointment]['children']
		for ids in Appointment:
			if 'children' in data[ids]:
				for j in range(0,len(data[ids]['children'])):
					Date_ID = data[ids]['children'][j]
					if 'attribute' in data[Date_ID] and 'attributeValue' in data[Date_ID]:
						if data[Date_ID]['attribute'] == 'StartDate':
							Directors_Data[4] = data[Date_ID]['attributeValue'].strip()
						elif data[Date_ID]['attribute'] == 'EndDate':
							Directors_Data[5] = data[Date_ID]['attributeValue'].strip()
	
	with open(File_path_Directors_csv, 'a', newline='', encoding='utf-8') as file:
		writer = csv.writer(file)
		writer.writerow(Directors_Data)
	with open(File_path_Directors_txt,"a", encoding='utf-8')as f:
		f.write("\t".join(map(str,Directors_Data))+"\n")
		f.flush()

def Secretaries(data,UIN):
	collection = {}
	isCompany = 0
	for i in data:
		if 'DC_' in i:
			break
		else:
			if 'domain' in data[i] and data[i]['domain'] == 'EntitySecretary':
				isCompany = 1
				collection[i] = isCompany
			elif 'domain' in data[i] and data[i]['domain'] == 'IndividualSecretary':
				isCompany = 0
				collection[i] = isCompany
	for key,value in collection.items():
			Secretaries_Data = ['']*14
			Secretaries_Data[13] = UIN
			if value == 1:
				Details_ID = data[key]['children'][1]
				Secretary = data[Details_ID]['children'][0]
				Secretary = data[Secretary]['children'][0]
				for id in data[Secretary]['children']:
					if 'attribute' in data[id] and 'attributeValue' in data[id] and data[id]['attribute'] == 'Name':
						Secretaries_Data[0] = data[id]['attributeValue']
					elif 'attribute' in data[id] and 'attributeValue' in data[id] and data[id]['attribute'] == 'EntityNumber':
						Secretaries_Data[1] = data[id]['attributeValue']
					elif 'text' in data[id] and 'label' in data[id]['text'] and data[id]['text']['label'] == 'Registered Office Address':
						ad = data[id]['children'][0]
						Secretaries_Data[2] = data[ad]['text']['row'].strip()
				Representative = data[Details_ID]['children'][1]
				for j in data[Representative]['children']:
					if 'attribute' in data[j] and 'attributeValue' in data[j] and data[j]['attribute'] == 'Name':
						Secretaries_Data[3] = data[j]['attributeValue'].strip()
					elif 'domain' in data[j] and 'text' in data[j] and data[j]['domain'] == 'ServiceAddress':
						Secretaries_Data[4] = data[j]['text']['row'].strip()
				Additional_Details = data[Details_ID]['children'][2]
				for k in data[Additional_Details]['children']:
					if 'children' in data[k]:
						for l in data[k]['children']:
							if data[l]['attribute'] == 'StartDate':
								Secretaries_Data[5] = data[l]['attributeValue'].strip()
				Secretaries_Data[11] = isCompany
				with open(File_path_Secretaries_csv, 'a', newline='', encoding='utf-8') as file:
					writer = csv.writer(file)
					writer.writerow(Secretaries_Data)
				with open(File_path_Secretaries_txt,"a", encoding='utf-8')as f:
					f.write("\t".join(map(str,Secretaries_Data))+"\n")
					f.flush()
			else:
				Secretaries_Data[6] = data[key]['text']['row'].strip()
				Details = data[key]['children'][1]
				Nationality = data[Details]['children'][0]
				Nationality = data[Nationality]['children'][1]
				if 'text' in data[Nationality]:
					Secretaries_Data[7] = data[Nationality]['text']['dcText'].strip()
				Address = data[Details]['children'][2]
				Address = data[data[Address]['children'][0]]['children'][0]
				Secretaries_Data[8] = data[Address]['text']['row'].strip()
				Postal = data[Details]['children'][3]
				Postal = data[data[Postal]['children'][0]]['children'][0]
				Secretaries_Data[9] = data[Postal]['text']['row'].strip()
				Additional = data[Details]['children'][4]
				for a in data[Additional]['children']:
					if 'children' in data[a]:
						for b in data[a]['children']:
							if 'attribute' in data[b] and data[b]['attribute'] == 'StartDate':
								Secretaries_Data[10] = data[b]['attributeValue'].strip()
							elif 'attribute' in data[b] and data[b]['attribute'] == 'EndDate':
								Secretaries_Data[12] = data[b]['attributeValue'].strip()
				Secretaries_Data[11] = isCompany
				with open(File_path_Secretaries_csv, 'a', newline='', encoding='utf-8') as file:
					writer = csv.writer(file)
					writer.writerow(Secretaries_Data)
				with open(File_path_Secretaries_txt,"a", encoding='utf-8')as f:
					f.write("\t".join(map(str,Secretaries_Data))+"\n")
					f.flush()


def Auditors(data,UIN):
	Auditors_Data = ['']*6
	collections = []
	for i in data:
		if 'DC_' in i:
			break
		else: 
			if 'domain' in data[i] and data[i]['domain'] == 'IndividualAuditor':
				collections.append(i)

	for id in collections:
		Auditors_Data[5] = UIN
		if 'text' in data[id] and 'singleline' in data[id]['text']:
			Auditors_Data[0] = data[id]['text']['singleline']
		Detail_ID = data[id]['children'][1]
		Nationality_ID = data[Detail_ID]['children'][0]
		Nationality_ID = data[Nationality_ID]['children'][1]
		if 'attribute' in data[Nationality_ID] and 'text' in data[Nationality_ID] and data[Nationality_ID]['attribute'] == 'Nationality':
			Auditors_Data[1] = data[Nationality_ID]['text']['dcText']
		Resident = data[Detail_ID]['children'][2]
		Resident = data[Resident]['children'][0]
		Resident = data[Resident]['children'][0]
		if 'domain' in data[Resident] and 'text' in data[Resident] and data[Resident]['domain'] == 'ResidentialAddress':
			Auditors_Data[2] = data[Resident]['text']['row'].strip()
		Appointment = data[Detail_ID]['children'][3]
		Appointment = data[Appointment]['children']
		for ids in Appointment:
			if 'children' in data[ids]:
				for j in range(0,len(data[ids]['children'])):
					Date_ID = data[ids]['children'][j]
					if 'attribute' in data[Date_ID] and 'attributeValue' in data[Date_ID]:
						if data[Date_ID]['attribute'] == 'StartDate':
							Auditors_Data[3] = data[Date_ID]['attributeValue'].strip()
						elif data[Date_ID]['attribute'] == 'EndDate':
							Auditors_Data[4] = data[Date_ID]['attributeValue'].strip()
		with open(File_path_Auditors_csv, 'a', newline='', encoding='utf-8') as file:
			writer = csv.writer(file)
			writer.writerow(Auditors_Data)
		with open(File_path_Auditors_txt,"a", encoding='utf-8')as f:
			f.write("\t".join(map(str,Auditors_Data))+"\n")
			f.flush()

def Shareholders(data,UIN):
	collection = {}
	isCompany = 0
	for ids in data:
		if 'DC_' in ids:
			break
		else:
			if 'domain' in data[ids] and data[ids]['domain'] == 'IndividualShareholder':
				isCompany = 0
				collection[ids] = isCompany
			elif 'domain' in data[ids] and data[ids]['domain'] == 'EntityShareholder':
				isCompany = 1
				collection[ids] = isCompany
			elif 'domain' in data[ids] and data[ids]['domain'] == 'OtherShareholder':
				isCompany = 2
				collection[ids] = isCompany
	for key,value in collection.items():
		Shareholders_Data = ['']*18
		Shareholders_Data[17] = UIN
		if value == 0:
			Shareholders_Data[10] = data[key]['text']['singleline'].strip()
			Details = data[key]['children'][1]
			Nationality = data[Details]['children'][0]
			Nationality = data[Nationality]['children'][1]
			if 'text' in data[Nationality]:
				Shareholders_Data[11] = data[Nationality]['text']['dcText']
			Address = data[Details]['children'][2]
			Address = data[data[Address]['children'][0]]['children'][0]
			Shareholders_Data[12] = data[Address]['text']['row'].strip()
			Postal = data[Details]['children'][3]
			Postal = data[data[Postal]['children'][0]]['children'][0]
			Shareholders_Data[13] = data[Postal]['text']['row'].strip()
			Nominee = data[Details]['children'][4]
			Nominee = data[Nominee]['children'][0]
			Shareholders_Data[14] = data[Nominee]['text']['dcText'].strip()
			Additional = data[Details]['children'][5]
			for i in data[Additional]['children']:
				if 'children' in data[i]:
					for m in data[i]['children']:
						if 'attribute' in data[m] and data[m]['attribute'] == 'StartDate':
							Shareholders_Data[15] = data[m]['attributeValue']
			Shareholders_Data[16] = value
			with open(File_path_Shareholders_txt, 'a', newline='', encoding='utf-8') as file:
				writer = csv.writer(file)
				writer.writerow(Shareholders_Data)
			with open(File_path_Shareholders_txt,"a", encoding='utf-8')as f:
				f.write("\t".join(map(str,Shareholders_Data))+"\n")
				f.flush()
		elif value == 1:
			Details = data[key]['children'][1]
			Share = data[Details]['children'][0]
			for j in data[Share]['children']:
				if 'attribute' in data[j] and data[j]['attribute'] == 'Name':
					Shareholders_Data[0] = data[j]['attributeValue'].strip()
				elif 'attribute' in data[j] and data[j]['attribute'] == 'EntityNumber':
					if 'attributeValue' in data[j]:
						Shareholders_Data[1] = data[j]['attributeValue'].strip()
				elif 'text' in data[j] and 'label' in data[j]['text'] == 'Registered Office Address':
					Ad = data[j]['children'][0]
					Shareholders_Data[2] = data[Ad]['text']['row'].strip()
			Postal = data[Details]['children'][1]
			Postal = data[Postal]['children'][0]
			Shareholders_Data[3] = data[Postal]['text']['row'].strip()
			Additional_Details = data[Details]['children'][3]
			for k in data[Additional_Details]['children']:
				if 'attribute' in data[k] and data[k]['attribute'] == 'NomineeYn':
					Shareholders_Data[4] = data[k]['text']['dcText'].strip() if 'dcText' in data[k]['text'] else ''
				elif 'children' in data[k]:
					for l in data[k]['children']:
						if 'attribute' in data and data[l]['attribute'] == 'StartDate':
							Shareholders_Data[5] = data[l]['attributeValue'].strip()
			Shareholders_Data[16] = value
			with open(File_path_Shareholders_csv, 'a', newline='', encoding='utf-8') as file:
				writer = csv.writer(file)
				writer.writerow(Shareholders_Data)
			with open(File_path_Shareholders_txt,"a", encoding='utf-8')as f:
				f.write("\t".join(map(str,Shareholders_Data))+"\n")
				f.flush()
		elif value == 2:
			Details = data[key]['children'][1]
			Shareholders_Details = data[Details]['children'][0]
			for i in data[Shareholders_Details]['children']:
				if 'attribute' in data[i] and data[i]['attribute'] == 'Name':
					Shareholders_Data[6] = data[i]['attributeValue'].strip()
				elif 'text' in data[i] and data[i]['text']['label'] == 'Address':
					Shareholders_Data[7] = data[data[i]['children'][0]]['text']['singleline'].strip()
				elif 'text' in data[i] and data[i]['text']['label'] == 'Postal Address':
					Shareholders_Data[13] = data[data[i]['children'][0]]['text']['singleline'].strip()
				elif 'attribute' in data[i] and data[i]['attribute'] == 'EntityNumber':
					Shareholders_Data[8] = data[i]['attributeValue'].strip() if 'attributeValue' in data[i] else ''
				elif 'attribute' in data[i] and data[i]['attribute'] == 'CountryOfOrigin':
					Shareholders_Data[9] = data[i]['text']['dcText'].strip() if 'dcText' in data[i]['text'] else ''
			Nominee = data[Details]['children'][2]
			if 'text' in data[Nominee] and data[Nominee]['text']['label'] == 'Nominee and Beneficial Owner Details':
					Shareholders_Data[14] = data[data[Nominee]['children'][0]]['text']['dcText'].strip()
			Appointment = data[Details]['children'][3]
			for m in data[Appointment]['children']:
				if 'children' in data[m]:
					for j in data[m]['children']:
						if 'attribute' in data[j] and data[j]['attribute'] == 'StartDate':
							Shareholders_Data[15] = data[j]['attributeValue'].strip()
			Shareholders_Data[16] = value
			with open(File_path_Shareholders_csv, 'a', newline='', encoding='utf-8') as file:
				writer = csv.writer(file)
				writer.writerow(Shareholders_Data)
			with open(File_path_Shareholders_txt,"a", encoding='utf-8')as f:
				f.write("\t".join(map(str,Shareholders_Data))+"\n")
				f.flush()


def Share_Allocations(data,UIN):
	Share_Allocations_Data = ['']*5
	Type = ''
	Total_Shares = ''
	for i in data:
		if 'DC_' in i:
			break
		else:
			if 'attribute' in data[i] and data[i]['attribute'] == 'ShareAllocationType':
				Type += data[i]['text']['dcText'].strip() if 'dcText' in data[i]['text'] else ''
			elif 'attribute' in data[i] and data[i]['attribute'] == 'TotalShares':
				Total_Shares += data[i]['attributeValue'].strip()
	for j in data:
		if 'DC_' in j:
			break
		else:
			if 'domain' in data[j] and data[j]['domain'] == 'OwnershipBundle':
				Share_Allocations_Data[4] = UIN
				Shares = data[j]['children'][0]
				Share_Allocations_Data[1] = data[Shares]['attributeValue'].strip()
				Name = data[data[data[j]['children'][1]]['children'][0]]['children'][0]
				Share_Allocations_Data[2] = data[Name]['attributeValue'].strip()
				Share_Allocations_Data[0] = Type
				Share_Allocations_Data[3] = Total_Shares
				with open(File_path_Share_Allocations_csv, 'a', newline='', encoding='utf-8') as file:
					writer = csv.writer(file)
					writer.writerow(Share_Allocations_Data)
				with open(File_path_Share_Allocations_txt,"a", encoding='utf-8')as f:
					f.write("\t".join(map(str,Share_Allocations_Data))+"\n")
					f.flush()


def Proprietors(data,UIN):
	collection = {}
	isCompany = 0
	for i in data:
		if 'DC_' in i:
			break
		else:
			if 'domain' in data[i] and data[i]['domain'] == 'EntityProprietor':
				isCompany = 1
				collection[i] = isCompany
			elif 'domain' in data[i] and data[i]['domain'] == 'IndividualProprietor':
				isCompany = 0
				collection[i] = isCompany
	for key,value in collection.items():
		Proprietors_Data = ['']*12
		Proprietors_Data[11] = UIN
		if value == 1:
			Info = data[key]['children'][1]
			Details = data[Info]['children'][0]
			for k in data[Details]['children']:
				if 'attribute' in data[k] and data[k]['attribute'] == 'Name':
					Proprietors_Data[0] = data[k]['attributeValue'].strip()
				elif 'attribute' in data[k] and data[k]['attribute'] == 'EntityNumber':
					Proprietors_Data[1] = data[k]['attributeValue'].strip() if 'attributeValue' in data[k] else ''
				elif 'text' in data[k] and data[k]['text']['label'] == 'Registered Office Address':
					Proprietors_Data[2] = data[data[k]['children'][0]]['text']['singleline'].strip()
				elif 'text' in data[k] and data[k]['text']['label'] == 'Postal Address':
					Proprietors_Data[3] = data[data[k]['children'][0]]['text']['singleline'].strip()
			Appointment = data[Info]['children'][1]
			Appointment = data[Appointment]['children'][0]
			for l in data[Appointment]['children']:
				if 'attribute' in data[l] and data[l]['attribute'] == 'StartDate':
					Proprietors_Data[4] = data[l]['attributeValue'].strip()
			Proprietors_Data[10] = value
			with open(File_path_Proprietors_csv, 'a', newline='', encoding='utf-8') as file:
				writer = csv.writer(file)
				writer.writerow(Proprietors_Data)
			with open(File_path_Proprietors_txt,"a", encoding='utf-8')as f:
				f.write("\t".join(map(str,Proprietors_Data))+"\n")
				f.flush()
		else:
			Proprietors_Data[5] = data[key]['text']['singleline'].strip()
			Info = data[key]['children'][1]
			for j in data[Info]['children']:
				if 'text' in data[j] and data[j]['text']['label'] == 'Proprietor\'s Details':
					Proprietors_Data[6] = data[data[j]['children'][1]]['text']['dcText'].strip()
				elif 'text' in data[j] and data[j]['text']['label'] == 'Residential Address':
					Proprietors_Data[7] = data[data[data[j]['children'][0]]['children'][0]]['text']['singleline'].strip()
				elif 'text' in data[j] and data[j]['text']['label'] == 'Postal Address':
					Proprietors_Data[8] = data[data[data[j]['children'][0]]['children'][0]]['text']['singleline'].strip()
				elif 'text' in data[j] and data[j]['text']['label'] == 'Additional Details':
					for k in data[data[j]['children'][0]]['children']:
						if 'attribute' in data[k] and data[k]['attribute'] == 'StartDate':
							Proprietors_Data[9] = data[k]['attributeValue'].strip()
			Proprietors_Data[10] = value
			with open(File_path_Proprietors_csv, 'a', newline='', encoding='utf-8') as file:
				writer = csv.writer(file)
				writer.writerow(Proprietors_Data)
			with open(File_path_Proprietors_txt,"a", encoding='utf-8')as f:
				f.write("\t".join(map(str,Proprietors_Data))+"\n")
				f.flush()


def restart_script():
    python = sys.executable
    subprocess.call([python] + sys.argv)


def Individual_Company(Company_ID,Company_Name,UIN):
	Indi_payload = json.dumps({"returnRootHtmlOnChange":"false","returnChangesOnly":"true","commands":[{"type":"conflict-check"},{"type":"view-node-button-click","id":Company_ID}]})
	Indi_headers = {"cookie":Cookie,'User-Agent': user_agent}
	try:
		try:
			Individual = requests.post(New_SearchPage_Url,data=Indi_payload,headers=Indi_headers)
			Indi_data = Individual.json()
		except requests.exceptions.RequestException as e:
			exception()
			log_print('\n\nRequest Failed\nRestarting the script in 5 min...\n===========================================================')
			time.sleep(300)
			restart_script()
		try:
			Indi_URL = Indi_data['redirect']
		except KeyError:
			log_print('\n\nCannot find redirect link\nRestarting the script in 5 min...\n===========================================================')
			time.sleep(300)
			restart_script()
			# os._exit(1)
			# time.sleep(120)
			# return None
		try:
			Indi_GET = requests.get(Indi_URL)
			Indi_soup = BeautifulSoup(Indi_GET.content, 'html.parser')
			Title = Indi_soup.find('title')
			if Title.string == 'Error 500':
				return
			Indi_script = Indi_soup.find_all('script',type='text/javascript')[0].string
			Indi_View_Tree = Indi_script.split('var viewTree = ')
			Indi_View_Tree = Indi_View_Tree[1].split('for (key in viewTree)')
			Indi_View_Tree = Indi_View_Tree[0].strip()
			json_data = json.loads(Indi_View_Tree[:-1])
			Individual.close()
			Indi_GET.close()
			Dereference(Individual)
			Dereference(Indi_GET)
			Indi_soup.decompose()
			for it in json_data:
				if 'DC_' in it:
					break
				elif (it != 'root') and ('widget' in json_data[it]) and (json_data[it]['widget'] == 'wizard') and (json_data[it]['text']['shortlabel'] == 'Company Details' or json_data[it]['text']['shortlabel'] == 'Business Name'):
					Totaltabchild =json_data[it]['children']
					for currentitem in Totaltabchild:
						if(json_data[currentitem]['text']['label'] == 'General Details'):
							General_Details(it,json_data,1) if json_data[it]['text']['shortlabel'] == 'Company Details' else General_Details(it,json_data,0)
						elif(json_data[currentitem]['text']['label'] == 'Addresses'):
							Addresses(json_data,UIN)
						elif(json_data[currentitem]['text']['label'] == 'Directors'):
							Directors(json_data,UIN)
						elif(json_data[currentitem]['text']['label'] == 'Secretaries'):
							Secretaries(json_data,UIN)
						elif(json_data[currentitem]['text']['label'] == 'Auditors'):
							Auditors(json_data,UIN)
						elif(json_data[currentitem]['text']['label'] == 'Shareholders'):
							Shareholders(json_data,UIN)
						elif(json_data[currentitem]['text']['label'] == 'Share Allocations'):
							Share_Allocations(json_data,UIN)
						elif(json_data[currentitem]['text']['label'] == 'Proprietors'):
							Proprietors(json_data,UIN)
		except Exception as e:
			exception()
		
	except Exception as e:
		exception()

def Search_Page_Info(ID_Arr,collection):
	# For Name & Status
	for NS in ID_Arr:
		Search_Page_data = []
		hasUIN = 0
		UIN_test = collection[NS[1]]['children']
		for k in range(0,len(UIN_test)):
			if ('attribute' in collection[UIN_test[k]]) and (collection[UIN_test[k]]['attribute']=="businessIdentifier"):
				if ('attributeValue' in collection[UIN_test[k]]):
					hasUIN = 1
		if hasUIN == 1:
			Name_Status = collection[NS[0]]['children']
			Name = collection[Name_Status[0]]['text']['label'].strip()
			Search_Page_data.append(collection[Name_Status[0]]['text']['label'].strip())
			Search_Page_data.append(collection[Name_Status[1]]['text']['singleline'].strip())
		# For Registration No. , Company Type & Registration Date
			CRC = collection[NS[1]]['children']
			for i in range(0,len(CRC)):
				if ('attribute' in collection[CRC[i]]) and (collection[CRC[i]]['attribute']=="businessIdentifier"):
					Search_Page_data.append(collection[CRC[i]]['attributeValue'].strip() if 'attributeValue' in collection[CRC[i]] else '')
					UIN_no = collection[CRC[i]]['attributeValue'].strip()
				elif ('attribute' in collection[CRC[i]]) and (collection[CRC[i]]['attribute']=="Type"):
					Search_Page_data.append(collection[CRC[i]]['attributeValue'].strip() if 'attributeValue' in collection[CRC[i]] else '')
				elif ('attribute' in collection[CRC[i]]) and (collection[CRC[i]]['attribute']=="RegistrationDate"):
					Search_Page_data.append(collection[CRC[i]]['attributeValue'].strip() if 'attributeValue' in collection[CRC[i]] else '')
		# For Address
			if len(NS)==3:
				if ('children' in collection[NS[2]]):
					Add = collection[NS[2]]['children']
					if ('domain' in collection[Add[0]]) and (collection[Add[0]]['domain']=="PrimaryAddress"):
						Search_Page_data.append(collection[Add[0]]['text']['singleline'].strip() if collection[Add[0]]['text']['singleline'] else '')
			# indi_ts_start = time.time()
			Individual_Company(collection[Name_Status[0]]['id'],Name,UIN_no)
			# indi_ts_stop = time.time()
			# log_print(f"Added {collection[Name_Status[0]]['text']['label'].strip()} {indi_ts_stop - indi_ts_start:.2f}")
			with open(File_path_Search_Page_Info_csv, 'a', newline='', encoding='utf-8') as file:
				writer = csv.writer(file)
				writer.writerow(Search_Page_data)
			with open(File_path_Search_Page_Info_txt,"a", encoding='utf-8')as f:
				f.write("\t".join(map(str,Search_Page_data))+"\n")
				f.flush()
			count()


if __name__=='__main__':

############################################# Writing Headers for Excel Files #############################################
	File_paths= [File_path_Search_Page_Info,File_path_Addresses,File_path_Auditors,File_path_Directors,File_path_General_Details,File_path_Proprietors,File_path_Secretaries,File_path_Share_Allocations,File_path_Shareholders]
	File_paths_csv= [File_path_Search_Page_Info_csv,File_path_Addresses_csv,File_path_Auditors_csv,File_path_Directors_csv,File_path_General_Details_csv,File_path_Proprietors_csv,File_path_Secretaries_csv,File_path_Share_Allocations_csv,File_path_Shareholders_csv]
	File_paths_txt= [File_path_Search_Page_Info_txt,File_path_Addresses_txt,File_path_Auditors_txt,File_path_Directors_txt,File_path_General_Details_txt,File_path_Proprietors_txt,File_path_Secretaries_txt,File_path_Share_Allocations_txt,File_path_Shareholders_txt]
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
		File_path_index = [File_path_log_index_LetterE1, File_path_log_index_LetterE2, File_path_log_index_LetterE3, File_path_log_index]
		if os.path.exists(File_path_count):
			os.remove(File_path_count)
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

	Search_Page_Headers = ['Company name','Status','Registration number','Company Type','Registration date','Company address']
	General_Details_Headers = ['Previous Name','Name Date from','Name Date to','Business Name Status','Registration Date','Business Activities',
			    'Business Activity','Date of Commencement of Business Activity','Renewal Filing Month','UIN','Company Status','Status',
				'Status Date From','Status Date to','Foreign Company','Exempt','Incorporation Date','Re-Registration Date','Have own constitution','Annual Return Filing Month','Annual Return last filed on']
	Addresses_Headers = ['Registered Office Address','Previous Registered Office Address','Postal Address','Previous Postal Addresses','Principal Place of Business']
	Proprietors_Headers = ['Company Name','UIN','Registered Office Address','Postal address','Company Appointment Date','Name','Nationality','Residential Address','Postal address','Individual Appointment Date']
	Directors_Headers = ['Name','Nationality','Residential Address','Postal address','Appointment Date','ceased Date']
	Secretaries_Headers = ['Company Name','UIN','Registered Office Address','Representative Name','Representative Postal address','Appointment Date','Name','Nationality','Residential Address','Postal address','Appointment Date']
	Shareholders_Headers = ['Company Name','UIN','Registered Office Address','Company Postal Address','Company Nominee shareholder','Company Appointment Date','Entity Name','Address','Registration Number','Country of Registration',
				'Name of Shareholder','Nationality','Residential Address','Individual Postal Address','Individual Nominee shareholder','Individual Appointment Date']
	Auditors_Headers = ['Name','Nationality','Residential Address','Appointment Date','Ceased date']
	Share_Allocations_Headers = ['Share allocation type','Number of Shares','Shareholder Name','Total number of shares']

	with open(File_path_count,"a", encoding='utf-8')as f:
		f.write("")
  
	txt_files = [
		(File_path_Search_Page_Info_txt, Search_Page_Headers),
		(File_path_General_Details_txt, General_Details_Headers),
		(File_path_Addresses_txt, Addresses_Headers),
		(File_path_Proprietors_txt, Proprietors_Headers),
		(File_path_Directors_txt, Directors_Headers),
		(File_path_Secretaries_txt, Secretaries_Headers),
		(File_path_Shareholders_txt, Shareholders_Headers),
		(File_path_Share_Allocations_txt, Share_Allocations_Headers),
		(File_path_Auditors_txt, Auditors_Headers)
	]

	for file_path, headers in txt_files:
		with open(file_path, "a", encoding='utf-8') as f:
			if f.tell() == 0:
				f.write("\t".join(headers) + "\n")
				f.flush()

	csv_files = [
		(File_path_Search_Page_Info_csv, Search_Page_Headers),
		(File_path_General_Details_csv, General_Details_Headers),
		(File_path_Addresses_csv, Addresses_Headers),
		(File_path_Proprietors_csv, Proprietors_Headers),
		(File_path_Directors_csv, Directors_Headers),
		(File_path_Secretaries_csv, Secretaries_Headers),
		(File_path_Shareholders_csv, Shareholders_Headers),
		(File_path_Share_Allocations_csv, Share_Allocations_Headers),
		(File_path_Auditors_csv, Auditors_Headers)
	]

	for file_path, headers in csv_files:
		with open(file_path, "a", newline='', encoding='utf-8') as f:
			writer = csv.writer(f)
			if f.tell() == 0:
				writer.writerow(headers)
	
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

	HomeURL = 'https://www.cipa.co.bw/'
	SearchPageURL = HomeURL + 'ng-cipa-master/ui/start/entitySearch'
	
	try:
		Driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()),options=chrome_options)
		Driver.get(SearchPageURL)
		Cookies = Driver.get_cookies()
		Cookie = Cookies[0]['name'] + '=' + Cookies[0]['value']
		New_SearchPage_Url = Driver.current_url
		
		soup = BeautifulSoup(Driver.page_source.encode(), 'html.parser')
		ID_Script = soup.find_all('script',type='text/javascript')
		Data = ID_Script[0].string
		soup.decompose()
		Driver.close()
		Driver.quit()
		View_Tree = Data.split('var viewTree = ')
		json_obj = View_Tree[1].split('for ')
		JSON_text = json_obj[0] + '\"}'
		s = json.loads(JSON_text)
		Key1 = s[s[s[s[s['root']]['children'][0]]['children'][0]]['children'][0]]['children'][1]
		# Key2 = s[s[s[s[s['root']]['children'][0]]['children'][0]]['children'][0]]['children'][2]


		log_index_flag_LetterE1 = False
		if os.path.exists(File_path_log_index_LetterE1):
			with open(File_path_log_index_LetterE1, 'r', encoding='utf-8') as file:
				index_LetterE1 = file.read().strip()
				if index_LetterE1 != '':
					log_index_flag_LetterE1 = True
		if log_index_flag_LetterE1:
			start_index_LetterE1 = alphabet.index(index_LetterE1) + 1
		else:
			start_index_LetterE1 = 0

		log_index_flag_LetterE2 = False
		if os.path.exists(File_path_log_index_LetterE2):
			with open(File_path_log_index_LetterE2, 'r', encoding='utf-8') as file:
				index_LetterE2 = file.read().strip()
				if index_LetterE2 != '' and index_LetterE2 != 'z':
					log_index_flag_LetterE2 = True
		if log_index_flag_LetterE2:
			start_index_LetterE2 = alphabet.index(index_LetterE2) + 1
		else:
			start_index_LetterE2 = 0


		log_index_flag_LetterE3 = False
		if os.path.exists(File_path_log_index_LetterE3):
			with open(File_path_log_index_LetterE3, 'r', encoding='utf-8') as file:
				index_LetterE3 = file.read().strip()
				if index_LetterE3 != '' and index_LetterE3 != 'z':
					log_index_flag_LetterE3 = True
		if log_index_flag_LetterE3:
			start_index_LetterE3 = alphabet.index(index_LetterE3) + 1
		else:
			start_index_LetterE3 = 0


		# combinations = [a + b for a in string.ascii_lowercase for b in string.ascii_lowercase]
		# for combination in combinations:
		# for letter in alphabet:
		for letterE1 in range(start_index_LetterE1, len(alphabet)):
			letterE1 = alphabet[letterE1]
			for indexE2 in range(start_index_LetterE2, len(alphabet)):
				letterE2 = alphabet[indexE2]
				for indexE3 in range(start_index_LetterE3, len(alphabet)):
					letterE3 = alphabet[indexE3]
					letter = letterE1+letterE2+letterE3
					log_print('Searching for Letter : ' + letter.upper())
					letter_ts_start = time.time()
					payload_1 = json.dumps({"returnRootHtmlOnChange":'false',"returnChangesOnly":'true',"commands":[{"type":"view-node-set-attribute-value","id":Key1,"value":letter}]})
					user_agent = random.choice(user_agents)
					header_1 = {
						'content-type': 'application/json',
						'cookie': Cookie,
						'User-Agent': user_agent 
						}
					try:
						try:
							response_1 = requests.post(New_SearchPage_Url, data=payload_1, headers=header_1)
							res_data_1 = response_1.json()
							# Process the response here
						except requests.exceptions.RequestException as e:
							exception()
							log_print('\n\nRequest Failed\nRestarting the script in 5 min...\n===========================================================')
							time.sleep(300)
							restart_script()
						data_1 = res_data_1['state']
						search_ID = ''
						for ids in data_1:
							if 'widget' in data_1[ids] and data_1[ids]['widget'] == 'search-results':
								search_ID = data_1[ids]['id']
						Dereference(response_1)
						Dereference(res_data_1)
						Dereference(data_1)
						
						Total_page = 0
						# letterTemp = alphabet[start_index_LetterE1] + alphabet[start_index_LetterE2] + alphabet[start_index_LetterE3]
						# if letter == letterTemp:
						payload_page = json.dumps({"returnRootHtmlOnChange":"false","returnChangesOnly":"true","commands":[{"type":"pagination-update","id":search_ID,"page":1,"size":200},{"type":"view-node-execute-rule","id":search_ID,"scope":"page-change"}]})
						header_page = {'cookie':Cookie,'User-Agent': user_agent}
						try:
							try:
								response_page = requests.post(New_SearchPage_Url,data=payload_page,headers=header_page)
								page_data = response_page.json()
							except requests.exceptions.RequestException as e:
								exception()
								log_print('\n\nRequest Failed\nRestarting the script in 5 min...\n===========================================================')
								time.sleep(300)
								restart_script()
							data_pg = page_data['state']
							for pg in data_pg:
								# log_print(str(data_pg))
								if data_pg[pg] is not None:
									if ('widget' in data_pg[pg]) and (data_pg[pg]['widget'] == 'search-results'):
										Total_page = int(data_pg[pg]['kv']['ui-total'])/int(data_pg[pg]['kv']['ui-size'])
										break
							Dereference(response_page)
							Dereference(page_data)
							Dereference(data_pg)
							Total_page = math.ceil(Total_page) if Total_page < 50 else 50
						except Exception as e:
							exception()
						# else:
						# 	for ids in data_1:
						# 		# log_print(str(data_1))
						# 		if 'widget' in data_1[ids] and data_1[ids]['widget'] == 'search-results':
						# 			try:
						# 				Total_page = int(data_1[ids]['kv']['ui-total'])/int(data_1[ids]['kv']['ui-size'])
						# 			except ZeroDivisionError:
						# 				Total_page = 0
						# 			break
							# Total_page = math.ceil(Total_page) if Total_page < 50 else 50

						log_index_flag = False
						if os.path.exists(File_path_log_index):
							log_index_flag = True
							with open(File_path_log_index, 'r', encoding='utf-8') as file:
								try:
									last_processed_page = int(file.read().strip())
								except:
									last_processed_page = ''

						if log_index_flag and last_processed_page != '':
							start_index = last_processed_page + 1
						else:
							start_index = 1

						for page in range(1, Total_page+1):		##########################TODO##########################
							log_print('-------------------- Page : ' + str(page) +'/' + str(Total_page) + ' --------------------')
							pg_ts_start = time.time()
							payload_2 = json.dumps({"returnRootHtmlOnChange":"false","returnChangesOnly":"true","commands":[{"type":"pagination-update","id":search_ID,"page":int(page),"size":200},{"type":"view-node-execute-rule","id":search_ID,"scope":"page-change"}]})
							user_agent = random.choice(user_agents)
							header_2 = {'cookie':Cookie,'User-Agent': user_agent}
							try:
								try:
									response_2 = requests.post(New_SearchPage_Url,data=payload_2,headers=header_2)
									res_data = response_2.json()
								except requests.exceptions.RequestException as e:
									exception()
									log_print('\n\nRequest Failed\nRestarting the script in 5 min...\n===========================================================')
									time.sleep(300)
									restart_script()
								data = res_data['state']
								Master = []
								for item in data:
									if data[item] != None:
										if ('dos' in data[item]) and (data[item]['dos'][0]=="css-entity-search-result"):
											Master.append(data[item]['children'])
								Search_Page_Info(Master,data)
								Dereference(response_2)
								# Dereference(res_data)
								# Dereference(data)
							except Exception as e:
								exception()
							pg_ts_stop = time.time()
							log_print(f'Page Time: {pg_ts_stop - pg_ts_start: .2f}')
							with open(File_path_log_index, 'w', encoding='utf-8') as file:
								file.write(str(page))
								file.flush()
					except Exception as e:
						exception()
					letter_ts_stop = time.time()
					with open(File_path_log_index_LetterE3, 'w', encoding='utf-8') as file:
						file.write(letterE3)
						file.flush()
					with open(File_path_log_index, 'w', encoding='utf-8') as file:
						file.write('')
						file.flush()
					log_print(f'Letter Time: {letter_ts_stop - letter_ts_start: .2f}')
				start_index_LetterE3 = 0
				with open(File_path_log_index_LetterE2, 'w', encoding='utf-8') as file:
					file.write(letterE2)
					file.flush()
			start_index_LetterE2 = 0
			with open(File_path_log_index_LetterE1, 'w', encoding='utf-8') as file:
				file.write(letterE1)
				file.flush()
		start_index_LetterE1 = 0
	except Exception as e:
		exception()

	for file in File_paths_csv:
		duplicateFromCSV(file)
	for file_index in range(0,8):
		convertCSVExcel(File_paths_csv[file_index], File_paths[file_index])
	log_print('Success')
	exit()
	
database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
	for File_path in File_paths:
		delete_task(conn, File_path)