import openpyxl
import pandas as pd
import os.path

# FOR USE IN ETS.PY FOR CREATING SINGLE DAILY REPORT - CREATES A COPY OF TEMPLATE EXCEL FILE
def create_new_workbook(file_path,template):
    wb = openpyxl.load_workbook(template) #OPENS TEMPLATE EXCEL FILE
    wb.save(file_path) #SAVES AS NEW FILE

def fill_workbook(data,filepath):
    if os.path.isfile(data):
        writable_data = pd.read_csv(data) #READS CSV AND STORES AS DATAFRAME
        with (pd.ExcelWriter(path=filepath, #FILEPATH IS WHERE NEW REPORT WILL BE SAVED
                            mode="a", #APPENDING THE COPIED EXCEL FILE
                            engine="openpyxl",
                            if_sheet_exists="overlay")
              as writer):

            writable_data.to_excel(writer,
                                   sheet_name="WorkFlow",
                                   startrow=1,
                                   header=False,
                                   index=False)

def generate_report(data,filepath,template):
    if os.path.isfile(data):
        create_new_workbook(filepath,template)
        fill_workbook(data,filepath)

## FOR USE IN REPORT GENERATOR FOR MULTIPLE REPORTS
def create_new_multi_workbook(file_path,template):
    wb = openpyxl.load_workbook(template) #OPENS TEMPLATE EXCEL FILE
    wb.save(file_path) #SAVES AS NEW FILE

def fill_multi_workbook(data,filepath):
    writable_data = data #TAKES DATAFRAME AND SAVES AS TEMP VARIABLE
    with (pd.ExcelWriter(path=filepath,
                        mode="a",
                        engine="openpyxl",
                        if_sheet_exists="overlay")
          as writer):
        writable_data.to_excel(writer,
                               sheet_name="WorkFlow",
                               startrow=1, header=False,
                               index=False)

def generate_multi_report(data,filepath,template):
        create_new_multi_workbook(filepath, template)
        try:
            fill_multi_workbook(data, filepath) #FILLS WORKBOOK IF DATA FOUND
        except AttributeError:
            os.remove(filepath) #DELETES THE WORKBOOK IF NO DATA FOUND














