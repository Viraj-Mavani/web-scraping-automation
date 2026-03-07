import pandas as pd
import sys
import traceback
import os

BasePath = os.getcwd()
# BasePath= 'E:\\ADIP-PY\\OP2'
File_path_excels = BasePath + '\\InputFile\\'
File_path_merged_excel = BasePath + '\\OP\\MergedExcel.xlsx'
File_path_modified_excel = BasePath + '\\OP\\ModifiedExcel.xlsx'


def exception():
    error = traceback.format_exc()
    exception_type, exception_object, exception_traceback = sys.exc_info()
    print(error)


def merge_excel(excel_files):
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
    
    
def merge_sheets():
    # Read all sheets into a dictionary with sheet name as key and header as value
    merged_excel = pd.ExcelFile(File_path_merged_excel)
    sheet_headers = {}

    for sheet_name in merged_excel.sheet_names:
        data = merged_excel.parse(sheet_name)
        sheet_headers[sheet_name] = tuple(data.columns)

    # Merge sheets with similar headers
    merged_data = {}  # Reinitialize the dictionary for merged data

    while sheet_headers:
        sheet_name, headers = list(sheet_headers.items())[0]
        similar_sheets = {sheet_name}
        del sheet_headers[sheet_name]

        for other_sheet, other_headers in sheet_headers.items():
            if headers == other_headers:
                similar_sheets.add(other_sheet)

        merged_data[sheet_name] = pd.DataFrame(columns=headers)  # Initialize with the first sheet

        for sheet_to_merge in similar_sheets:
            data_to_merge = merged_excel.parse(sheet_to_merge)
            merged_data[sheet_name] = pd.concat([merged_data[sheet_name], data_to_merge], ignore_index=True)

    # Save the final merged data to a new Excel file (overwrite the existing one)
    with pd.ExcelWriter(File_path_modified_excel) as writer:
        for sheet_name, data in merged_data.items():
            data.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f'Final merged sheets saved to {File_path_modified_excel}')



if __name__ == '__main__':
    try: 
        directories = [
            BasePath + '\\OP',
            BasePath + '\\InputFile'
        ]

        for directory in directories:
            if not os.path.exists(directory):
                os.makedirs(directory)
        
        excel_files = [f for f in os.listdir(File_path_excels) if f.endswith('.xlsx')]
        merge_excel(excel_files)
        
        merge_sheets()
    except:
        exception()