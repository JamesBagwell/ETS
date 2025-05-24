from datetime import timedelta,datetime,date
import pandas as pd
import tkinter as tk
from tkinter import messagebox
from tkcalendar import Calendar
import Writer
import os

#GET DATES BETWEEEN TWO DATES INCLUDING START AND END DATES
def get_dates_between(start_date, end_date):
    return [start_date + timedelta(days=i)
            for i in range((end_date - start_date).days + 1)]

#GETS LIST OF FILES BETWEEN DATES
def get_all_files(dates):
    files = []
    for dates in dates:
        files.append(f"{base}/Reports/CSV Files/ETS {dates} WF")
    return files

#GETS DATAFRAMES FOR ALL FILES IN FILE LIST AND CONCATS INTO SINGLE DATAFRAME
def convert_to_df(list_of_files):
    start_date = start_datetime() #GET START DATE FOR MESSAGEBOX
    end_date = end_datetime() #GET END DAATE FOR MESSAGEBOX
    df_list = []
    files_found = 0
    files_not_found = 0
    for filename in list_of_files:
        try:
            df = pd.read_csv(filename, index_col=None, header=0)  #ADDS ALL DATAFRAMES TO LIST OF DATAFRAMES
            df_list.append(df)
            files_found += 1
        except FileNotFoundError:
            files_not_found += 1
    try:
        grouped_data = pd.concat(df_list, axis=0, ignore_index=True) #CONCATS ALL DATAFRAMES FROM LIST INTO ONE
        tk.messagebox.showinfo(title="File Summary",
                            message=f"Report created for dates: {start_date} to {end_date}\n"
                                    f"Usable files found: {files_found}\n"
                                    f"Files missing: {files_not_found}")
        return grouped_data  # RETURNS SINGLE DATAFRAME
    except ValueError: tk.messagebox.showerror(title="No Data Available",message="No data available for these dates")
    return False



#TAKES START DATE FROM WIDGET AND RETURNS IN CORRECT FORMAT
def start_datetime():
    start = datetime.strptime(start_date_calendar.get_date(), '%m/%d/%y').strftime("%Y/%m/%d")
    start_tuple = tuple(start.split("/"))
    start_date = date(int(start_tuple[0]), int(start_tuple[1]), int(start_tuple[2]))
    return start_date

#TAKES END DATE FROM WIDGET AND RETURNS IN CORRECT FORMAT
def end_datetime():
    end = datetime.strptime(end_date_calendar.get_date(), '%m/%d/%y').strftime("%Y/%m/%d")
    end_tuple = tuple(end.split("/"))
    end_date = date(int(end_tuple[0]), int(end_tuple[1]), int(end_tuple[2]))
    return end_date

#GET DATES FROM WIDGETS AND CREATE DATAFRAME
def get_dataframe_from_dates():
    date_range = get_dates_between(start_date=start_datetime(), #START DATE OF LIST OF DATES
                                   end_date=end_datetime()) #END DATE OF LIST OF DATES
    files_for_dates = get_all_files(date_range) #GETS ALL FILES FOR DATES IN THE LIST
    main_df = convert_to_df(files_for_dates) # CONCATS ALL FILES COLLECTED INTO SINGLE DATAFRAME
    return main_df #RETURNS AS SINGLE DATAFRAME

def generate_multiple_report():
    template_path = f"{base}/Reports/Templates/template_report.xlsx" #LINK TO TEMPLATE FILE
    start_date = start_datetime() #GET START DATE BASED ON USER INPUT FOR FILENAME CREATION
    end_date = end_datetime() #GET END DATE BASED ON USER INPUT FOR FILENAME CREATION
    filepath = f"{base}/Reports/Custom Reports/ETS {start_date} to {end_date}.xlsx" # USE TWO ABOVE VARIBLES
    main_csv = get_dataframe_from_dates() #CREATES USABLE VARIABLE FROM SELECTED DATAFRAMES
    Writer.generate_multi_report(data=main_csv, #DATA = MAIN DATAFRAME
                                 filepath=filepath, #FILEPATH OF WHERE NEW EXCEL IS CREATED
                                 template=template_path) #TEMPLATE FILEPATH

#TODAY'S DATE
current_datetime = datetime.now() #RETURN FULL DATE
day = current_datetime.day #TAKES DAY ONLY FROM FULL DATE
month = current_datetime.month #TAKES MONTH ONLY FROM FULL DATE
year = current_datetime.year #TAKES YEAR ONLY FROM FULL DATE

#BASE FOLDER
base = os.path.dirname(os.path.abspath(__file__)) #USED TO MAKE MAINTAIN FILE STRUCTURE ON ALL COMPUTERS

#GUI SET UP
root = tk.Tk()
root.title("Engineering report generator")
root.minsize(height=350,width=600)
root.maxsize(height=350,width=600)
root.grid_columnconfigure(index=(0,1,2),weight=1)
root.grid_rowconfigure(index=(0,1,2),weight=1)

#FONT VARIABLES
main_font = "Arial"
label_font_size = 18
main_pad = 5

#CALANDAR WIDGETS
start_date_calendar = Calendar(root,selectmode="day", #CALENDAR FOR START DATE SELECTION
                               year=year,
                               month=month,
                               day=day)

end_date_calendar = Calendar(root,selectmode="day", #CALANDAR FOR END DATE SELECTION
                             year=year,
                             month=month,
                             day=day)

#LABELS
date1_label = tk.Label(text="Start Date",font=(main_font,label_font_size),) #START DATE LABEL
date2_label = tk.Label(text="End Date",font=(main_font,label_font_size)) #END DATE LAVEL

#BUTTONS
button1 = tk.Button(text="Generate Report",command=generate_multiple_report,font=(main_font,18))
button1.config(height=2,width=30,padx=main_pad,pady=main_pad,bg="Grey")

#GRID LAYOUT
start_date_calendar.grid(column=0,row=1)
date1_label.grid(column=0,row=2,sticky="new")
end_date_calendar.grid(column=2,row=1)
date2_label.grid(column=2,row=2,sticky="new")
button1.grid(column=0,row=3,columnspan=3)

root.mainloop()