import json
import os
import random
import sys
import re
import sqlite3
from sqlite3 import Error
import pandas as pd
import requests
import string
import xlsxwriter
# from openpyxl.styles import Font

# BasePath = 'E:\\ADIP-PY'
# BasePath = 'D:\\Projects\\CedarPython\\ADIP-NG1501'
BasePath = os.getcwd()
Total_URL = 0

File_path_Search_Page_Info = BasePath + '\\OP\\ADIP-NG1501_Search_Page_Info.xlsx'
File_path_Search_Page_Info_txt = BasePath + '\\OPtxt\\ADIP-NG1501_Search_Page_Info.txt'
File_path_Shareholders = BasePath + '\\OP\\ADIP-NG1501_Shareholders.xlsx'
File_path_Shareholders_txt = BasePath + '\\OPtxt\\ADIP-NG1501_Shareholders.txt'
File_path_search_count = BasePath + '\\Counts\\ADIP-NG1501_Count.txt'
Error_File = BasePath + '\\Error\\ADIP-NG1501_Error.xlsx'
File_path_Input = BasePath + '\\Proxy\\http_proxies.xlsx'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-NG1501_Log.txt'

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


def Dereference(obj):
    del obj


if __name__ == '__main__':

    ############################################# Writing Headers for Excel Files #############################################
    File_paths = [File_path_Search_Page_Info, File_path_Shareholders]

    directories = [
        BasePath + '\\OP',
        BasePath + '\\OPtxt',
        # BasePath + '\\OPcsv',
        BasePath + '\\Error',
        BasePath + '\\Counts',
        BasePath + '\\Proxy',
        BasePath + '\\Log'
    ]

    for directory in directories:
        if not os.path.exists(directory):
            os.makedirs(directory)

    row1 = 1

    book1 = xlsxwriter.Workbook(File_path_Search_Page_Info)
    sheet1 = book1.add_worksheet()

    bold_format = book1.add_format({'bold': True})

    sheet1.write('A1', "RC Number", bold_format)
    sheet1.write('B1', "Company Name", bold_format)
    sheet1.write('C1', "Address", bold_format)
    sheet1.write('D1', "Status", bold_format)
    sheet1.write('E1', "Date of Registration", bold_format)

    # book1 = openpyxl.Workbook()
    # sheet1 = book1.active

    # sheet1['A1'] = "RC Number"
    # sheet1['A1'].font = Font(bold=True)
    # sheet1['B1'] = "Company Name"
    # sheet1['B1'].font = Font(bold=True)
    # sheet1['C1'] = "Address"
    # sheet1['C1'].font = Font(bold=True)
    # sheet1['D1'] = "Status"
    # sheet1['D1'].font = Font(bold=True)
    # sheet1['E1'] = "Date of Registration"
    # sheet1['E1'].font = Font(bold=True)
    # book1.save(File_path_Search_Page_Info)
    # Dereference(sheet1)
    # Dereference(book1)

    row2 = 1

    book2 = xlsxwriter.Workbook(File_path_Shareholders)
    sheet2 = book2.add_worksheet()

    bold_format = book2.add_format({'bold': True})

    sheet2.write('A1', "Shareholder Name", bold_format)
    sheet2.write('B1', "Address", bold_format)
    sheet2.write('C1', "Date of PSC", bold_format)
    sheet2.write('D1', "Legal form", bold_format)
    sheet2.write('E1', "Governing Law", bold_format)
    sheet2.write('F1', "Register", bold_format)
    sheet2.write('G1', "Country of registration", bold_format)
    sheet2.write('H1', "Email", bold_format)
    sheet2.write('I1', "Does the PSC directly or indirectly hold at least 5% of the shares or interest in a company or limited liability partnership?", bold_format)
    sheet2.write('J1', "Does the PSC directly or indirectly hold at least 5% of the voting rights in a company or limited liability partnership?", bold_format)
    sheet2.write('K1', "Does the PSC directly or indirectly hold the right to appoint or remove a majority of the directors or partners in a company or limited liability partnership?", bold_format)
    sheet2.write('L1', "Does the PSC otherwise have the right to exercise or is actually exercising significant influence or control over a company or limited liability partnership?", bold_format)
    sheet2.write('M1', "Does the PSC have the right to exercise, or actually exercise significant influence or control over the activities of a trust or firm, whether or not it is a legal entity, but would itself satisfy any of the first four conditions if it were an individual?", bold_format)
    sheet2.write('N1', "Company RC Number", bold_format)
    sheet2.write('O1', "Main Company Name", bold_format)

    # book2 = openpyxl.Workbook()
    # sheet2 = book2.active

    # sheet2['A1'] = "Shareholder Name"
    # sheet2['A1'].font = Font(bold=True)
    # sheet2['B1'] = "Address"
    # sheet2['B1'].font = Font(bold=True)
    # sheet2['C1'] = "Date of PSC"
    # sheet2['C1'].font = Font(bold=True)
    # sheet2['D1'] = "Legal form"
    # sheet2['D1'].font = Font(bold=True)
    # sheet2['E1'] = "Governing Law"
    # sheet2['E1'].font = Font(bold=True)
    # sheet2['F1'] = "Register"
    # sheet2['F1'].font = Font(bold=True)
    # sheet2['G1'] = "Country of registration"
    # sheet2['G1'].font = Font(bold=True)
    # sheet2['H1'] = "Email"
    # sheet2['H1'].font = Font(bold=True)
    # sheet2['I1'] = "Does the PSC directly or indirectly hold at least 5% of the shares or interest in a company or limited liability partnership?"
    # sheet2['I1'].font = Font(bold=True)
    # sheet2['J1'] = "Does the PSC directly or indirectly hold at least 5% of the voting rights in a company or limited liability partnership?"
    # sheet2['J1'].font = Font(bold=True)
    # sheet2['K1'] = "Does the PSC directly or indirectly hold the right to appoint or remove a majority of the directors or partners in a company or limited liability partnership?"
    # sheet2['K1'].font = Font(bold=True)
    # sheet2['L1'] = "Does the PSC otherwise have the right to exercise or is actually exercising significant influence or control over a company or limited liability partnership?"
    # sheet2['L1'].font = Font(bold=True)
    # sheet2['M1'] = "Does the PSC have the right to exercise, or actually exercise significant influence or control over the activities of a trust or firm, whether or not it is a legal entity, but would itself satisfy any of the first four conditions if it were an individual?"
    # sheet2['M1'].font = Font(bold=True)
    # sheet2['N1'] = "Company RC Number"
    # sheet2['N1'].font = Font(bold=True)
    # sheet2['O1'] = "Main Company Name"
    # sheet2['O1'].font = Font(bold=True)
    # book2.save(File_path_Shareholders)
    # # book2.close()
    # Dereference(sheet2)
    # Dereference(book2)

    row3 = 1
    book3 = xlsxwriter.Workbook(Error_File)
    sheet3 = book3.add_worksheet()

    bold_format = book3.add_format({'bold': True})

    sheet3.write('A1', "URL", bold_format)
    sheet3.write('B1', "Responding Status", bold_format)
    sheet3.write('C1', "Error", bold_format)

    # book3 = openpyxl.Workbook()
    # sheet3 = book3.active

    # sheet3['A1'] = "URL"
    # sheet3['A1'].font = Font(bold=True)
    # sheet3['B1'] = "Responding Status"
    # sheet3['B1'].font = Font(bold=True)
    # sheet3['C1'] = "Error"
    # sheet3['C1'].font = Font(bold=True)
    # book3.save(Error_File)
    # book3.close()
    # Dereference(sheet3)
    # Dereference(book3)

    Search_Page_Headers = ['RC Number', 'Company Name',
                           'Address', 'Status', 'Date of Registration']
    Shareholders_Headers = ['Shareholder Name', 'Address', 'Date of PSC', 'Legal form', 'governing Law', 'Register', 'Country of registration', 'Email',
                            'Does the PSC directly or indirectly hold at least 5% of the shares or interest in a company or limited liability partnership?',
                            'Does the PSC directly or indirectly hold at least 5% of the voting rights in a company or limited liability partnership?',
                            'Does the PSC directly or indirectly hold the right to appoint or remove a majority of the directors or partners in a company or limited liability partnership?',
                            'Does the PSC otherwise have the right to exercise or is actually exercising significant influence or control over a company or limited liability partnership?',
                            'Does the PSC have the right to exercise, or actually exercise significant influence or control over the activities of a trust or firm, whether or not it is a legal entity, but would itself satisfy any of the first four conditions if it were an individual?',
                            'Company RC Number', 'Main Company Name']

    with open(File_path_search_count, "w", encoding='utf-8')as f:
        f.write("")
    with open(File_path_Search_Page_Info_txt, "w", encoding='utf-8')as f:
        f.write("\t".join(Search_Page_Headers)+"\n")
    with open(File_path_Shareholders_txt, "w", encoding='utf-8')as fw:
        fw.write("\t".join(Shareholders_Headers)+"\n")

    Search_Page_URL = 'https://searchapp.cac.gov.ng/searchapp/api/public-search/company-business-name-it'

    # # Proxy Credentials
    # username = 'CedarRose-res-NG'
    # password = 'DZjZc7NJi9F8pi0'
    # server = 'gw.ntnt.io'
    # port = '5959'
    # proxy = {
    # 'http': f'http://{username}:{password}@{server}:{port}',
    # 'https': f'http://{username}:{password}@{server}:{port}'
    # }
    # proxies = [
    #     '95.216.189.78:8080',
    #     '65.21.0.216:8080',
    #     '65.109.236.232:8080',
    #     '120.197.219.82:9091',
    #     '5.161.78.209:8080']

    df = pd.read_excel(File_path_Input, sheet_name='Sheet1')
    proxies = df['proxies'].tolist()
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
	] # List of user-agent headers


    try:
        for l1 in alphabet:
            for l2 in alphabet:
                letter = l1+l2
                print('\n\nScraping for Letter : ' + letter)
                payload = json.dumps({"searchTerm": letter})
                headers = {"content-type": "application/json"}
                proxy = random.choice(proxies)
                proxy_url = f'http://{proxy}'
                print(f"Using proxy: {proxy}")
                proxies = {'http': proxy, 'https': proxy}
                try:
                    Search_Page = requests.post(
                        Search_Page_URL, data=payload, headers=headers, proxies=proxies)
                    # Search_Page = requests.post(
                    #     Search_Page_URL, data=payload, headers=headers)
                    if re.search(r'"data":\[\{"state"', Search_Page.text):
                        res_data = json.loads(Search_Page.text)
                        search_data = res_data['data']
                    else:
                        search_data = None
                    if search_data == None:
                        continue
                    Status_ids = []
                    for ids in search_data:
                        Status_ids.append(ids['id'])
                    Status_url = 'https://searchapp.cac.gov.ng/searchapp/api/public-search/check-company-status'
                    Status_payload = {"companyIds": Status_ids}
                    Status_headers = {"content-type": "application/json",
                                      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/113.0.0.0 Safari/537.36"}
                    Status = requests.post(Status_url, data=json.dumps(
                        Status_payload), headers=Status_headers, proxies=proxies)
                    if re.search(r'"data":\{"\d+"', Status.text):
                        status_data = json.loads(Status.text)
                        status_info = status_data['data']
                    else:
                        status_info = None
                    if status_info == None:
                        continue
                    for i in search_data:
                        if i['classificationId'] == 2:
                            ID = i['id']
                            Search_data = ['']*5
                            print('Adding ' + i['approvedName'].strip())
                            with open(File_path_search_count, "a", encoding='utf-8')as fh:
                                fh.write("1\n")

                            sheet1.write(
                                row1, 0, i['rcNumber'] if i['rcNumber'] != None else 'NOT YET ASSIGNED')
                            Search_data[0] = i['rcNumber'] if i['rcNumber'] != None else 'NOT YET ASSIGNED'
                            RC = i['rcNumber'] if i['rcNumber'] != None else 'NOT YET ASSIGNED'

                            # sheet1['B{col}'.format(col=row1)] = i['approvedName'].strip() if i['approvedName'] != None else ''
                            sheet1.write(row1, 1, i['approvedName'].strip(
                            ) if i['approvedName'] is not None else '')
                            company_name = i['approvedName'].strip(
                            ) if i['approvedName'] != None else ''
                            Search_data[1] = i['approvedName'].strip(
                            ) if i['approvedName'] != None else ''

                            # sheet1['C{col}'.format(col=row1)] = ' '.join(i['address'].split()) if i['address'] != None else ''
                            sheet1.write(row1, 2, ' '.join(
                                i['address'].split()) if i['address'] is not None else '')
                            Search_data[2] = ' '.join(
                                i['address'].split()) if i['address'] != None else ''

                            # sheet1['D{col}'.format(col=row1)] = i['companyStatus'] if i['companyStatus'] != None else 'INACTIVE'
                            sheet1.write(
                                row1, 3, i['companyStatus'] if i['companyStatus'] is not None else 'INACTIVE')
                            Search_data[3] = i['companyStatus'] if i['companyStatus'] != None else 'INACTIVE'

                            # sheet1['E{col}'.format(col=row1)] = i['registrationDate'].split('T')[0] if i['registrationDate'] != None else 'UNDER REGISTRATION'
                            sheet1.write(row1, 4, i['registrationDate'].split('T')[
                                         0] if i['registrationDate'] is not None else 'UNDER REGISTRATION')
                            Search_data[4] = i['registrationDate'].split(
                                'T')[0] if i['registrationDate'] != None else 'UNDER REGISTRATION'
                            row1 += 1

                            with open(File_path_Search_Page_Info_txt, "a", encoding='utf-8') as fw:
                                fw.write("\t".join(map(str, Search_data))+"\n")

                            # book1.save(File_path_Search_Page_Info)
                            # book1.close()
                            try:
                                if ID:
                                    Shareholders_URL = 'https://searchapp.cac.gov.ng/searchapp/api/status-report/find/company-affiliates/{id}'.format(
                                        id=ID)
                                    Share_headers = {
                                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.0.0 Safari/537.36",
                                        "Accept-Encoding": "gzip, deflate, br",
                                        "Connection": "keep-alive",
                                        "content-type": "application/json",
                                        "accept": "*/*"
                                    }
                                    try_count = 1
                                    while True:
                                        try:
                                            Shareholders_Info = requests.get(
                                                Shareholders_URL, headers=Share_headers, proxies=proxies)
                                            break
                                        except:
                                            if try_count > 3:
                                                break
                                            try_count += 1
                                    if Shareholders_Info.status_code == 200:
                                        res = json.loads(
                                            Shareholders_Info.text)
                                        data = res['data']
                                        Shareholders_Data = ['']*15
                                        for item in data:
                                            if item['affiliatesPscInformation']:
                                                sheet2.write(row2, 13, RC)
                                                sheet2.write(
                                                    row2, 14, company_name)
                                                # sheet2['N{col}'.format(col=row2)] = RC
                                                # sheet2['O{col}'.format(col=row2)] = company_name
                                                Shareholders_Data[13] = RC
                                                Shareholders_Data[14] = company_name
                                                if item['isCorporate'] == None or item['isCorporate'] == False:
                                                    Name = item['surname'].strip(
                                                    ) + ' ' if item['surname'] else ''
                                                    Name += item['firstname'].strip() + \
                                                        ' ' if item['firstname'] else ''
                                                    Name += item['otherName'].strip() + \
                                                        ' ' if item['otherName'] else ''
                                                elif item['isCorporate'] == True:
                                                    Name = item['corporationName'].strip(
                                                    ) + ' ' if item['corporationName'] else ''
                                                    Name += item['rcNumber'].strip() + \
                                                        ' ' if item['rcNumber'] else ''

                                                sheet2.write(row2, 0, Name)
                                                # sheet2['A{col}'.format(col=row2)] = Name
                                                Shareholders_Data[0] = Name

                                                Address = item['streetNumber'] + \
                                                    ', ' if item['streetNumber'] else ''
                                                Address += item['address'] + \
                                                    ', ' if item['address'] else ''
                                                Address += item['city'] + \
                                                    ', ' if item['city'] else ''
                                                if item['state'] == "FCT":
                                                    Address += item['state']
                                                elif item['state']:
                                                    Address += item['state'] + \
                                                        ' STATE, '
                                                else:
                                                    Address += ''
                                                sheet2.write(row2, 1, Address)
                                                # sheet2['B{col}'.format(col=row2)] = Address
                                                Shareholders_Data[1] = Address

                                                sheet2.write(
                                                    row2, 2, item['affiliatesPscInformation']['dateOfPsc'] if item['affiliatesPscInformation']['dateOfPsc'] else '')
                                                # sheet2['C{col}'.format(col=row2)] = item['affiliatesPscInformation']['dateOfPsc'] if item['affiliatesPscInformation']['dateOfPsc'] else ''
                                                Shareholders_Data[2] = item['affiliatesPscInformation'][
                                                    'dateOfPsc'] if item['affiliatesPscInformation']['dateOfPsc'] else ''

                                                sheet2.write(
                                                    row2, 3, item['affiliatesPscInformation']['legalForm'] if item['affiliatesPscInformation']['legalForm'] else '')
                                                Shareholders_Data[3] = item['affiliatesPscInformation'][
                                                    'legalForm'] if item['affiliatesPscInformation']['legalForm'] else ''

                                                sheet2.write(
                                                    row2, 4, item['affiliatesPscInformation']['governingLaw'] if item['affiliatesPscInformation']['governingLaw'] else '')
                                                Shareholders_Data[4] = item['affiliatesPscInformation'][
                                                    'governingLaw'] if item['affiliatesPscInformation']['governingLaw'] else ''

                                                sheet2.write(
                                                    row2, 5, item['affiliatesPscInformation']['register'] if item['affiliatesPscInformation']['register'] else '')
                                                Shareholders_Data[5] = item['affiliatesPscInformation'][
                                                    'register'] if item['affiliatesPscInformation']['register'] else ''

                                                sheet2.write(row2, 6, item['affiliatesPscInformation']['taxResidencyOrJurisdiction']
                                                             if item['affiliatesPscInformation']['taxResidencyOrJurisdiction'] else '')
                                                Shareholders_Data[6] = item['affiliatesPscInformation']['taxResidencyOrJurisdiction'] if item[
                                                    'affiliatesPscInformation']['taxResidencyOrJurisdiction'] else ''

                                                sheet2.write(
                                                    row2, 7, item['email'] if item['email'] else '')
                                                Shareholders_Data[7] = item['email'] if item['email'] else ''

                                                # sheet2['D{col}'.format(col=row2)] = item['affiliatesPscInformation']['legalForm'] if item['affiliatesPscInformation']['legalForm'] else ''
                                                # Shareholders_Data[3] = item['affiliatesPscInformation']['legalForm'] if item['affiliatesPscInformation']['legalForm'] else ''

                                                # sheet2['E{col}'.format(col=row2)] = item['affiliatesPscInformation']['governingLaw'] if item['affiliatesPscInformation']['governingLaw'] else ''
                                                # Shareholders_Data[4] = item['affiliatesPscInformation']['governingLaw'] if item['affiliatesPscInformation']['governingLaw'] else ''

                                                # sheet2['F{col}'.format(col=row2)] = item['affiliatesPscInformation']['register'] if item['affiliatesPscInformation']['register'] else ''
                                                # Shareholders_Data[5] = item['affiliatesPscInformation']['register'] if item['affiliatesPscInformation']['register'] else ''

                                                # sheet2['G{col}'.format(col=row2)] = item['affiliatesPscInformation']['taxResidencyOrJurisdiction'] if item['affiliatesPscInformation']['taxResidencyOrJurisdiction'] else ''
                                                # Shareholders_Data[6] = item['affiliatesPscInformation']['taxResidencyOrJurisdiction'] if item['affiliatesPscInformation']['taxResidencyOrJurisdiction'] else ''

                                                # sheet2['H{col}'.format(col=row2)] = item['email'] if item['email'] else ''
                                                # Shareholders_Data[7] = item['email'] if item['email'] else ''

                                                if item['affiliatesPscInformation']['pscHoldsSharesOrInterest']:
                                                    sheet2.write(row2, 8, "YES Directly: [{d}%] and Indirectly: [{ind}%]".format(
                                                        d=item['affiliatesPscInformation']['pscHoldsSharesOrInterestPercentageHeldDirectly'], ind=item['affiliatesPscInformation']['pscHoldsSharesOrInterestPercentageHeldIndirectly']))
                                                    # sheet2['I{col}'.format(col=row2)] = 'YES Directly: [{d}%] and Indirectly: [{ind}%]'.format(d=item['affiliatesPscInformation']['pscHoldsSharesOrInterestPercentageHeldDirectly'],ind=item['affiliatesPscInformation']['pscHoldsSharesOrInterestPercentageHeldIndirectly'])
                                                    Shareholders_Data[8] = 'YES Directly: [{d}%] and Indirectly: [{ind}%]'.format(
                                                        d=item['affiliatesPscInformation']['pscHoldsSharesOrInterestPercentageHeldDirectly'], ind=item['affiliatesPscInformation']['pscHoldsSharesOrInterestPercentageHeldIndirectly'])
                                                else:
                                                    sheet2.write(
                                                        row2, 8, "NO Directly: [0%] and Indirectly: [0%]")
                                                    # sheet2['I{col}'.format(col=row2)] = 'NO Directly: [0%] and Indirectly: [0%]'
                                                    Shareholders_Data[8] = 'NO Directly: [0%] and Indirectly: [0%]'
                                                if item['affiliatesPscInformation']['pscVotingRights']:
                                                    sheet2.write(row2, 9, "YES Directly: [{d}%] and Indirectly: [{ind}%]".format(
                                                        d=item['affiliatesPscInformation']['pscVotingRightsPercentageHeldDirectly'], ind=item['affiliatesPscInformation']['pscVotingRightsPercentageHeldIndirectly']))
                                                    # sheet2['J{col}'.format(col=row2)] = 'YES Directly: [{d}%] and Indirectly: [{ind}%]'.format(d=item['affiliatesPscInformation']['pscVotingRightsPercentageHeldDirectly'],ind=item['affiliatesPscInformation']['pscVotingRightsPercentageHeldIndirectly'])
                                                    Shareholders_Data[9] = 'YES Directly: [{d}%] and Indirectly: [{ind}%]'.format(
                                                        d=item['affiliatesPscInformation']['pscVotingRightsPercentageHeldDirectly'], ind=item['affiliatesPscInformation']['pscVotingRightsPercentageHeldIndirectly'])
                                                else:
                                                    sheet2.write(
                                                        row2, 9, "NO Directly: [0%] and Indirectly: [0%]")
                                                    # sheet2['J{col}'.format(col=row2)] = 'NO Directly: [0%] and Indirectly: [0%]'
                                                    Shareholders_Data[9] = 'NO Directly: [0%] and Indirectly: [0%]'
                                                if item['affiliatesPscInformation']['pscRightToAppoints']:
                                                    sheet2.write(
                                                        row2, 10, "YES")
                                                    # sheet2['K{col}'.format(col=row2)] = 'YES'
                                                    Shareholders_Data[10] = 'YES'
                                                else:
                                                    sheet2.write(
                                                        row2, 10, "NO")
                                                    # sheet2['K{col}'.format(col=row2)] = 'NO'
                                                    Shareholders_Data[10] = 'NO'
                                                if item['affiliatesPscInformation']['pscSignificantInfluence']:
                                                    sheet2.write(
                                                        row2, 11, "YES")
                                                    # sheet2['L{col}'.format(col=row2)] = 'YES'
                                                    Shareholders_Data[11] = 'YES'
                                                else:
                                                    sheet2.write(
                                                        row2, 11, "NO")
                                                    # sheet2['L{col}'.format(col=row2)] = 'NO'
                                                    Shareholders_Data[11] = 'NO'
                                                if item['affiliatesPscInformation']['pscExeriseSignificantInfluence']:
                                                    sheet2.write(
                                                        row2, 12, "YES")
                                                    # sheet2['M{col}'.format(col=row2)] = 'YES'
                                                    Shareholders_Data[12] = 'YES'
                                                else:
                                                    sheet2.write(
                                                        row2, 12, "NO")
                                                    # sheet2['M{col}'.format(col=row2)] = 'NO'
                                                    Shareholders_Data[12] = 'NO'

                                                row2 += 1
                                                with open(File_path_Shareholders_txt, "a", encoding='utf-8') as fw:
                                                    fw.write(
                                                        "\t".join(map(str, Shareholders_Data))+"\n")
                                                # book2.save(File_path_Shareholders)
                                                # book2.close()
                                                # Dereference(sheet2)
                                                # Dereference(book2)
                            except Exception as e:
                                exception_type, exception_object, exception_traceback = sys.exc_info()
                                filename = exception_traceback.tb_frame.f_code.co_filename
                                line_number = exception_traceback.tb_lineno
                                # book3 = openpyxl.load_workbook(Error_File)
                                # sheet3 = book3.active
                                sheet3.write(row3, 0, Search_Page_URL)
                                sheet3.write(row3, 1, 'Not Responding')
                                sheet3.write(row3, 2, str(e))
                                # sheet3['A{col}'.format(col=row3)] = Search_Page_URL
                                # sheet3['B{col}'.format(col=row3)] = 'Not Responding'
                                # sheet3['C{col}'.format(col=row3)] = str(e)
                                # book3.save(Error_File)
                                # Dereference(sheet3)
                                # Dereference(book3)
                                row3 += 1
                        Dereference(sheet1)
                        Dereference(book1)
                except Exception as e:
                    exception_type, exception_object, exception_traceback = sys.exc_info()
                    filename = exception_traceback.tb_frame.f_code.co_filename
                    line_number = exception_traceback.tb_lineno
                    # book3 = openpyxl.load_workbook(Error_File)
                    # sheet3 = book3.active
                    sheet3.write(row3, 0, Search_Page_URL)
                    sheet3.write(row3, 1, 'Not Responding')
                    sheet3.write(row3, 2, str(e))
                    # book3.save(Error_File)
                    # book3.close()
                    # Dereference(sheet3)
                    # Dereference(book3)
                    row3 += 1
    finally:
        book1.close()
        book2.close()
        book3.close()
        print('Success')
        exit()
database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
    for File_path in File_paths:
        delete_task(conn, File_path)
