import pandas as pd
import sys
import traceback
import os

# BasePath = os.getcwd()
BasePath= 'D:\\Projects\\CedarPython\\ADIP-ExcelMerge'
File_path_excels = BasePath + '\\InputFile\\'
File_path_merged_excel = BasePath + '\\OP\\MergedExcel.xlsx'


def exception():
    error = traceback.format_exc()
    exception_type, exception_object, exception_traceback = sys.exc_info()
    print(error)


def find_excel_files(directory):
    excel_files = []
    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith('.xlsx'):
                excel_files.append(os.path.join(root, file))
    return excel_files


if __name__ == '__main__':
    try: 
        directories = [
            BasePath + '\\OP',
            BasePath + '\\InputFile'
        ]

        for directory in directories:
            if not os.path.exists(directory):
                os.makedirs(directory)
        
        excel_files = find_excel_files(File_path_excels)
        data_frames = {}
        
        for file in excel_files:
            excel_file = pd.ExcelFile(os.path.join(File_path_excels, file))
            for sheet_name in excel_file.sheet_names:
                if sheet_name!="Index":
                    if sheet_name not in data_frames:
                        data_frames[sheet_name] = []
                    data = excel_file.parse(sheet_name)
                    data_frames[sheet_name].append(data)

        merged_data_frames = {}
        for sheet_name, data_list in data_frames.items():
            merged_data_frames[sheet_name] = pd.concat(data_list, ignore_index=True)

        with pd.ExcelWriter(File_path_merged_excel) as writer:
            for sheet_name, data in merged_data_frames.items():
                data.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f'Merged data saved to {File_path_merged_excel}')
    except:
        exception()