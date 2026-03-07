import pandas as pd
import os

BasePath = os.getcwd()
File_path_excels = BasePath + '\\InputFile\\'
File_path_merged_excel = BasePath + '\\OP\\MergedExcel_Option1.xlsx'

def merge_sheets_option1():
    excel_files = [f for f in os.listdir(File_path_excels) if f.endswith('.xlsx')]
    merged_data = pd.DataFrame()  # Initialize an empty DataFrame to store merged data

    for file in excel_files:
        excel_file = pd.ExcelFile(os.path.join(File_path_excels, file))
        for sheet_name in excel_file.sheet_names:
            if sheet_name!="Index":
                data = excel_file.parse(sheet_name)
                # Add the data to the merged_data DataFrame, filling missing columns with null
                merged_data = pd.concat([merged_data, data], axis=0, ignore_index=True)

    # Save the merged data to a new Excel file
    merged_data.to_excel(File_path_merged_excel, index=False)
    print(f'Merged data saved to {File_path_merged_excel}')

if __name__ == '__main__':
    merge_sheets_option1()
