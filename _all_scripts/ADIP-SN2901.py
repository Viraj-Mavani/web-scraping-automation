import csv
import sys
import traceback
import pandas as pd
import os
import sqlite3
from sqlite3 import Error
from bs4 import BeautifulSoup
import time
import requests

BasePath = 'D:\\Projects\\CedarPython\\ADIP-SN2901'
Listing = {}

######### Excel #########
File_path_search = BasePath + '\\OP\\ADIP-SN2901_Search_Page.xlsx'
File_path_company_details = BasePath + '\\OP\\ADIP-SN2901_Company_Info.xlsx'
######### Text #########
File_path_search_txt = BasePath + '\\OPtxt\\ADIP-SN2901_Search_Page.txt'
File_path_company_details_txt = BasePath + \
    '\\OPtxt\\ADIP-SN2901_Company_Info.txt'
######### Error #########
Error_File = BasePath + '\\Error\\ADIP-SN2901_Error.xlsx'
######### Count #########
File_path_search_count = BasePath + '\\Counts\\ADIP-SN2901_Count.txt'
######### Log #########
File_path_log = BasePath + '\\Log\\ADIP-SN2901_Log.txt'
File_path_log_Run_Flag = BasePath + '\\Log\\ADIP-SN2901_Run_Flag.txt'
File_path_log_index = BasePath + '\\Log\\ADIP-SN2901_Log_Index.txt'
######### CSV #########
Error_File_CSV = BasePath + '\\OPcsv\\ADIP-SN2901_Error.csv'
File_path_search_CSV = BasePath + '\\OPcsv\\ADIP-SN2901_Search_Page.csv'
File_path_company_details_CSV = BasePath + \
    '\\OPcsv\\ADIP-SN2901_Company_Info.csv'


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


def Dereference(obj):
    del obj


def log_print(message):
    with open(File_path_log, 'a', encoding='utf-8') as file:
        file.write(message + '\n')
        file.flush()
    print(message)


def exception(URL):
    Headers_Error = ['URL', 'Not Responding', 'Error']
    error = traceback.format_exc()
    exception_type, exception_object, exception_traceback = sys.exc_info()
    with open(Error_File_CSV, 'a', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)
        if file.tell() == 0:
            writer.writerow(Headers_Error)
        writer.writerow([URL, "Not Responding", str(error)])
    df = pd.read_csv(Error_File_CSV, encoding='utf-8')
    df.to_excel(Error_File, index=False)


def count():
    try_count = 1
    while try_count <= 5:
        try:
            with open(File_path_search_count, 'a', encoding='utf-8') as fh:
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


def get_field_value(field_item):
    try:
        return field_item.string.strip()
    except AttributeError:
        return ''


def CompanyInfo(data, DCompanyName, IndividualURL):
    for Info in data:
        field_item = Info.find('div', class_=['field-item'])
        try:
            ClassName = Info['class'][1]
            if ClassName == 'field-name-field-siege-societe':
                Siège_social = get_field_value(field_item)
            elif ClassName == 'field-name-field-rc-societe':
                Régistre_de_commerce = get_field_value(field_item)
            elif ClassName == 'field-name-field-ninea-societe':
                Ninéa = get_field_value(field_item)
            elif ClassName == 'field-name-field-date-crea-societe':
                Date_Création = get_field_value(field_item)
            elif ClassName == 'field-name-field-localite':
                Localité = get_field_value(field_item)
            elif ClassName == 'field-name-field-gerance-societe':
                Gérance = get_field_value(field_item)
            elif ClassName == 'field-name-field-secteur':
                Secteur_d_activité = get_field_value(field_item)
            elif ClassName == 'field-name-field-forme-juriduqe':
                Forme_Juridique = get_field_value(field_item)
            elif ClassName == 'field-name-field-objet-societe':
                Objet_social = get_field_value(field_item)
                Objet_social = Objet_social.encode("ascii", "ignore")
            elif ClassName == 'field-name-field-exercice-societe':
                Exercice_social = get_field_value(field_item)
            elif ClassName == 'field-name-field-duree-societe':
                Durée = get_field_value(field_item)
            elif ClassName == 'field-name-field-region-societe':
                Région = get_field_value(field_item)
            elif ClassName == 'field-name-field-capital-societe':
                Capital = get_field_value(field_item)
            elif ClassName == 'field-name-field-apports':
                Montant_des_apports_en_numéraires = get_field_value(field_item)
            elif ClassName == 'field-name-field-description-sommaire':
                Description_sommaire_et_lévaluation_des_apports_en_nature = get_field_value(field_item)
            elif ClassName == 'field-name-field-nom-prenom-associes':
                Nom_prénoms_usuels_et_domicile_des_associés_tenus_indéfiniment_des_dettes_sociales = get_field_value(field_item)
            elif ClassName == 'field-name-field-nom-prenom-dirigeants':
                Nom_prénoms_et_domicile_des_premiers_dirigeants_et_des_premiers_commissaires_aux_comptes = get_field_value(field_item)
            elif ClassName == 'field-name-field-references-immatriclation':
                Références_de_limmatriculation_au_registre_du_commerce_et_du_crédit_mobilier = get_field_value(field_item)
            elif ClassName == 'field-name-field-date-commencement':
                Date_effective_ou_prévue_du_commencement_d_activité = get_field_value(field_item)
            elif ClassName == 'field-name-field-valeur-nominale-actions':
                Nombre_et_la_valeur_nominale_des_actions_souscrites_en_numéraire = get_field_value(field_item)
            elif ClassName == 'field-name-field-valeur-nom-actions-attr':
                Nombre_et_la_valeur_nominale_des_actions_attribuées_en_rémunération_de_chaque_apport_en_nature = get_field_value(field_item)
            elif ClassName == 'field-name-field-montant-partie-lib':
                Montant_de_la_partie_libérée = get_field_value(field_item)
            elif ClassName == 'field-name-field-disposition-statutaires':
                Dispositions_statutaires_relatives_à_la_constitution_des_réserves_et_à_la_répartition_des_bénéfices_et_du_boni_de_liquidation = get_field_value(field_item)
            elif ClassName == 'field-name-field-avantages-particuliers':
                Avantages_particuliers_stipulés = get_field_value(field_item)
            elif ClassName == 'field-name-field-conditions-admission':
                Conditions_d_admission_aux_assemblées_dactionnaires_et_dexercice_du_droit_de_vote = get_field_value(field_item)
            elif ClassName == 'field-name-field-existence-clause':
                Existence_de_clauses_relatives_à_l_agrément_des_cessionnaires_d_actions = get_field_value(field_item)
            elif ClassName == 'field-name-field-type-annonce':
                Type_Annonces = get_field_value(field_item)

            # if ClassName == 'field-name-field-siege-societe':
            #     Siège_social = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-rc-societe':
            #     Régistre_de_commerce = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-ninea-societe':
            #     Ninéa = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-date-crea-societe':
            #     Date_Création = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-localite':
            #     Localité = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-gerance-societe':
            #     Gérance = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-secteur':
            #     Secteur_d_activité = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-forme-juriduqe':
            #     Forme_Juridique = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-objet-societe':
            #     Objet_social = field_item.string.strip() if field_item.string is not None else ''
            #     Objet_social = Objet_social.encode("ascii", "ignore")
            # elif ClassName == 'field-name-field-exercice-societe':
            #     Exercice_social = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-duree-societe':
            #     Durée = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-region-societe':
            #     Région = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-capital-societe':
            #     Capital = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-apports':
            #     Montant_des_apports_en_numéraires = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-description-sommaire':
            #     Description_sommaire_et_lévaluation_des_apports_en_nature = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-nom-prenom-associes':
            #     Nom_prénoms_usuels_et_domicile_des_associés_tenus_indéfiniment_des_dettes_sociales = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-nom-prenom-dirigeants':
            #     Nom_prénoms_et_domicile_des_premiers_dirigeants_et_des_premiers_commissaires_aux_comptes = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-references-immatriclation':
            #     Références_de_limmatriculation_au_registre_du_commerce_et_du_crédit_mobilier = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-date-commencement':
            #     Date_effective_ou_prévue_du_commencement_d_activité = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-valeur-nominale-actions':
            #     Nombre_et_la_valeur_nominale_des_actions_souscrites_en_numéraire = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-valeur-nom-actions-attr':
            #     Nombre_et_la_valeur_nominale_des_actions_attribuées_en_rémunération_de_chaque_apport_en_nature = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-montant-partie-lib':
            #     Montant_de_la_partie_libérée = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-disposition-statutaires':
            #     Dispositions_statutaires_relatives_à_la_constitution_des_réserves_et_à_la_répartition_des_bénéfices_et_du_boni_de_liquidation = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-avantages-particuliers':
            #     Avantages_particuliers_stipulés = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-conditions-admission':
            #     Conditions_d_admission_aux_assemblées_dactionnaires_et_dexercice_du_droit_de_vote = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-existence-clause':
            #     Existence_de_clauses_relatives_à_l_agrément_des_cessionnaires_d_actions = field_item.string.strip() if field_item.string is not None else ''
            # elif ClassName == 'field-name-field-type-annonce':
            #     Type_Annonces = field_item.string.strip() if field_item.string is not None else ''

            # if ClassName == 'field-name-field-siege-societe':
            #     Siège_social = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-rc-societe':
            #     Régistre_de_commerce = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-ninea-societe':
            #     Ninéa = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-date-crea-societe':
            #     Date_Création = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-localite':
            #     Localité = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-gerance-societe':
            #     Gérance = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-secteur':
            #     Secteur_d_activité = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-forme-juriduqe':
            #     Forme_Juridique = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-objet-societe':
            #     Objet_social = field_item.string.strip() if field_item.string.strip() else ''
            #     Objet_social = Objet_social.encode("ascii", "ignore")
            # elif ClassName == 'field-name-field-exercice-societe':
            #     Exercice_social = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-duree-societe':
            #     Durée = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-region-societe':
            #     Région = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-capital-societe':
            #     Capital = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-apports':
            #     Montant_des_apports_en_numéraires = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-nom-prenom-associes':
            #     Nom_prénoms_usuels_et_domicile_des_associés_tenus_indéfiniment_des_dettes_sociales = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-nom-prenom-dirigeants':
            #     Nom_prénoms_et_domicile_des_premiers_dirigeants_et_des_premiers_commissaires_aux_comptes = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-references-immatriclation':
            #     Références_de_limmatriculation_au_registre_du_commerce_et_du_crédit_mobilier = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-date-commencement':
            #     Date_effective_ou_prévue_du_commencement_d_activité = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-valeur-nominale-actions':
            #     Nombre_et_la_valeur_nominale_des_actions_souscrites_en_numéraire = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-valeur-nom-actions-attr':
            #     Nombre_et_la_valeur_nominale_des_actions_attribuées_en_rémunération_de_chaque_apport_en_nature = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-montant-partie-lib':
            #     Montant_de_la_partie_libérée = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-disposition-statutaires':
            #     Dispositions_statutaires_relatives_à_la_constitution_des_réserves_et_à_la_répartition_des_bénéfices_et_du_boni_de_liquidation = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-avantages-particuliers':
            #     Avantages_particuliers_stipulés = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-conditions-admission':
            #     Conditions_d_admission_aux_assemblées_dactionnaires_et_dexercice_du_droit_de_vote = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-existence-clause':
            #     Existence_de_clauses_relatives_à_l_agrément_des_cessionnaires_d_actions = field_item.string.strip() if field_item.string.strip() else ''
            # elif ClassName == 'field-name-field-type-annonce':
            #     Type_Annonces = field_item.string.strip() if field_item.string.strip() else ''

                Company_Details_data = [DCompanyName, Siège_social, Régistre_de_commerce, Ninéa, Date_Création, Localité, Gérance, Secteur_d_activité, Forme_Juridique, Objet_social, Exercice_social, Durée, Région, Capital, Montant_des_apports_en_numéraires, Description_sommaire_et_lévaluation_des_apports_en_nature, Nom_prénoms_usuels_et_domicile_des_associés_tenus_indéfiniment_des_dettes_sociales, Nom_prénoms_et_domicile_des_premiers_dirigeants_et_des_premiers_commissaires_aux_comptes, Références_de_limmatriculation_au_registre_du_commerce_et_du_crédit_mobilier, Date_effective_ou_prévue_du_commencement_d_activité, Nombre_et_la_valeur_nominale_des_actions_souscrites_en_numéraire, Nombre_et_la_valeur_nominale_des_actions_attribuées_en_rémunération_de_chaque_apport_en_nature, Montant_de_la_partie_libérée, Dispositions_statutaires_relatives_à_la_constitution_des_réserves_et_à_la_répartition_des_bénéfices_et_du_boni_de_liquidation, Avantages_particuliers_stipulés, Conditions_d_admission_aux_assemblées_dactionnaires_et_dexercice_du_droit_de_vote, Existence_de_clauses_relatives_à_l_agrément_des_cessionnaires_d_actions, Type_Annonces, IndividualURL]
                with open(File_path_company_details_CSV, 'a', newline='', encoding='utf-8') as file:
                    writer = csv.writer(file)
                    writer.writerow(Company_Details_data)
                with open(File_path_company_details_txt, "a")as f:
                    f.write("\t".join(map(str, Company_Details_data))+"\n")
                    f.flush()
                Company_Details_data = []
        except:
            exception(IndividualURL)


if __name__ == '__main__':
    File_paths = [File_path_search, File_path_company_details]

    # Create directories if they don't exist
    directories = [
        BasePath + '\\OP',
        BasePath + '\\OPtxt',
        BasePath + '\\OPcsv',
        BasePath + '\\InputFile',
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
        File_paths_csv = [File_path_search_CSV, File_path_company_details_CSV]
        File_paths_txt = [File_path_search_txt, File_path_company_details_txt]
        if os.path.exists(File_path_search_count):
            os.remove(File_path_search_count)
        if os.path.exists(File_path_log):
            os.remove(File_path_log)
        for path_csv in File_paths_csv:
            if os.path.exists(path_csv):
                os.remove(path_csv)
        for Path_txt in File_paths_txt:
            if os.path.exists(Path_txt):
                os.remove(Path_txt)
        if os.path.exists(Error_File_CSV):
            os.remove(Error_File_CSV)
        if os.path.exists(File_path_log_index):
            os.remove(File_path_log_index)

    Search_page_txt = ["Name", "Year Of Creation",
                       "Headquaters", "Legal Form", "Main Activity", "Link"]
    Company_Details_txt = ["Name", "Siège social", "Régistre de commerce", "Ninéa", "Date Création", "Localité", "Gérance", "Secteur d'activité", "Forme Juridique", "Objet social", "Exercice social", "Durée", "Région", "Capital", "Montant des apports en numéraires", "Description sommaire et lévaluation des apports en nature", "Nom prénoms usuels et domicile des associés tenus indéfiniment des dettes sociales", "Nom prénoms et domicile des premiers dirigeants et des premiers commissaires aux comptes", "Références de limmatriculation au registre du commerce et du crédit mobilier",
                           "Date effective ou prévue du commencement d'activité", "Nombre et la valeur nominale des actions souscrites en numéraire", "Nombre et la valeur nominale des actions attribuées en rémunération de chaque apport en nature", "Montant de la partie libérée", "Dispositions statutaires relatives à la constitution des réserves et à la répartition des bénéfices et du boni de liquidation", "Avantages particuliers stipulés", "Conditions d'admission aux assemblées dactionnaires et dexercice du droit de vote", "Existence de clauses relatives à l'agrément des cessionnaires d'actions", "Type Annonces", "Link"]
    if not os.path.exists(File_path_search_count):
        with open(File_path_search_count, "a", encoding='utf-8')as f:
            f.write("")
    with open(File_path_search_txt, "a", encoding='utf-8')as f:
        if f.tell() == 0:
            f.write("\t".join(Search_page_txt)+"\n")
            f.flush()
    with open(File_path_company_details_txt, "a", encoding='utf-8')as fw:
        if fw.tell() == 0:
            fw.write("\t".join(Company_Details_txt)+"\n")
            fw.flush()

    with open(File_path_search_CSV, "a", newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        if f.tell() == 0:
            writer.writerow(Search_page_txt)
    with open(File_path_company_details_CSV, "a", newline='', encoding='utf-8') as fw:
        writer = csv.writer(fw)
        if fw.tell() == 0:
            writer.writerow(Company_Details_txt)

    retry_attempts = 6
    retry_delay = 2

    Home_URL = 'https://creationdentreprise.sn/en/finding-business'

    try:
        homeRetry = 1
        while homeRetry <= retry_attempts:
            try:
                Driver = requests.get(Home_URL)
                break
            except Exception as e:
                log_print(f"Error occurred in status for Home URL")
                delay = retry_delay * (2 ** homeRetry)
                log_print(f'Retrying in {delay} seconds...{homeRetry}')
                time.sleep(delay)
                homeRetry += 1
                continue
        else:
            exception()
            os._exit(1)
        if Driver.text.__len__() == 0:
            log_print('API Response ERROR in Company Get')
        else:
            soup = BeautifulSoup(Driver.text.encode(), 'html.parser')
            Dereference(Driver)
            Page = soup.find('li', class_='pager-current')
            Total_Pages = int(Page.getText().split(' ')[-1])
            soup.decompose()

            log_index_flag = False
            if os.path.exists(File_path_log_index):
                log_index_flag = True
                with open(File_path_log_index, 'r', encoding='utf-8') as file:
                    last_processed_page = int(file.read().strip())

            if log_index_flag:
                start_index = last_processed_page + 1
            else:
                start_index = 1

            for page in range(start_index, Total_Pages+1):
                st = time.time()
                int(page)
                if page == 1:
                    NextPageURL = 'https://creationdentreprise.sn/en/finding-business?field_rc_societe_value=&field_ninea_societe_value=&denomination=&field_localite_nid=All&field_siege_societe_value=&field_forme_juriduqe_nid=All&field_secteur_nid=All&field_date_crea_societe_value='
                else:
                    NextPageURL = 'https://creationdentreprise.sn/en/finding-business?field_rc_societe_value=&field_ninea_societe_value=&denomination=&field_localite_nid=All&field_siege_societe_value=&field_forme_juriduqe_nid=All&field_secteur_nid=All&field_date_crea_societe_value=&page={temp_page}'.format(
                        temp_page=page-1)
                try:
                    nextPageRetry = 1
                    while nextPageRetry <= retry_attempts:
                        try:
                            CompanyListDriver = requests.get(NextPageURL)
                            break
                        except Exception as e:
                            log_print(
                                f"Error occurred in status for Page {page}")
                            delay = retry_delay * (2 ** nextPageRetry)
                            log_print(
                                f'Retrying in {delay} seconds...{nextPageRetry}')
                            time.sleep(delay)
                            nextPageRetry += 1
                            continue
                    else:
                        exception()
                        os._exit(1)
                    Dereference(NextPageURL)
                    if CompanyListDriver.text.__len__() == 0:
                        log_print('Error while getting Next Page')
                    else:
                        CompanyListsoup = BeautifulSoup(
                            CompanyListDriver.text.encode(), 'html.parser')
                        Dereference(CompanyListDriver)
                        Company_Names = CompanyListsoup.select('td')
                        for a in Company_Names:
                            if a.attrs['class'][1] == 'views-field-title':
                                SName = a.get_text(strip=True)
                                SLink = 'https://creationdentreprise.sn' + \
                                    a.find('a')['href']
                            elif a.attrs['class'][1] == 'views-field-field-date-crea-societe':
                                SYear = a.get_text(strip=True)
                            elif a.attrs['class'][1] == 'views-field-field-siege-societe':
                                SHead = a.get_text(strip=True)
                            elif a.attrs['class'][1] == 'views-field-field-forme-juriduqe':
                                SLegal = a.get_text(strip=True)
                            elif a.attrs['class'][1] == 'views-field-field-secteur':
                                SActivity = a.get_text(strip=True)
                                Search_page_txt = [
                                    SName, SYear, SHead, SLegal, SActivity, SLink]

                                with open(File_path_search_CSV, 'a', newline='', encoding='utf-8') as file:
                                    writer = csv.writer(file)
                                    writer.writerow(Search_page_txt)
                                with open(File_path_search_txt, "a", encoding='utf-8')as f:
                                    f.write(
                                        "\t".join(map(str, Search_page_txt))+"\n")
                                    f.flush()
                                with open(File_path_search_count, "a", encoding='utf-8')as fh:
                                    fh.write("1\n")
                        Companies = CompanyListsoup.find_all(
                            'td', class_='views-field-title')
                        for URLs in Companies:
                            CompanyURL = URLs.find('a')['href']
                            IndividualURL = 'https://creationdentreprise.sn'+CompanyURL
                            try:
                                individualRetry = 1
                                while individualRetry <= retry_attempts:
                                    try:
                                        IndiCompanyDriver = requests.get(
                                            IndividualURL)
                                        break
                                    except Exception as e:
                                        log_print(
                                            f"Error occurred in status for {CompanyURL}")
                                        delay = retry_delay * \
                                            (2 ** individualRetry)
                                        log_print(
                                            f'Retrying in {delay} seconds...{individualRetry}')
                                        time.sleep(delay)
                                        individualRetry += 1
                                        continue
                                else:
                                    exception()
                                    os._exit(1)
                                time.sleep(2)
                                Dereference(IndividualURL)
                                if IndiCompanyDriver.text.__len__() == 0:
                                    log_print(
                                        'Erro while getting Individual Company')
                                else:
                                    IndiCompanysoup = BeautifulSoup(
                                        IndiCompanyDriver.text.encode(), 'html.parser')
                                    Dereference(IndiCompanyDriver)
                                    CompanyName = IndiCompanysoup.find(
                                        'h1', class_=['title-page', 'title-page-societe'])
                                    DCompanyName = CompanyName.get_text().replace('"', '')
                                    Company_Details = IndiCompanysoup.select(
                                        'div[class*="field-name-field-"]')
                                    CompanyInfo(Company_Details,
                                                DCompanyName, IndividualURL)

                                    log_print('Info Added: ' +
                                              CompanyName.get_text())
                            except Exception as e:
                                exception(IndividualURL)
                                Dereference(IndividualURL)

                        CompanyListsoup.decompose()
                        IndiCompanysoup.decompose()
                    with open(File_path_log_index, 'w', encoding='utf-8') as file:
                        file.write(str(page))
                        file.flush()
                    log_print('Page Completed: ' + str(page))
                    et = time.time()
                    log_print('Last Page Timing: ' +
                              str(round(et-st, 2)) + '\n')
                except Exception as e:
                    exception(NextPageURL)
                    Dereference(exception(NextPageURL))
    except Exception as e:
        exception(Home_URL)
        Dereference(Home_URL)

    finally:
        convertCSVExcel(File_path_search_CSV, File_path_search)
        convertCSVExcel(File_path_company_details_CSV,
                        File_path_company_details)
        duplicate(File_path_search)
        duplicate(File_path_company_details)

    log_print('Success')
    exit()
database = r"E:\ADIP Schedulers\NewWorkingDataBase\WorkingDB\InventoryDB.sqldb"

# create a database connection
conn = create_connection(database)
with conn:
    for File_path in File_paths:
        delete_task(conn, File_path)
