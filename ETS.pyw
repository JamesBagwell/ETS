import tkinter as tk
from tkinter import StringVar
from tkinter import messagebox
import datetime
import csv
import os.path
import Writer

class EtsGui:
    def __init__(self):

        #Main window Config
        self.root = tk.Tk()
        self.root.title("Engineering Time Sheet")
        self.root.minsize(width=1000,height=400)
        self.root.maxsize(width=1000,height=400)

        #Grid configuration
        self.root.rowconfigure(index=(0,1,2,3,4,5,6),weight=1)
        self.root.columnconfigure(index=(0,1,2,3,4,5),weight=2)
        self.root.columnconfigure(index=(0, 1, 2, 3, 4, 5), weight=2)

        #Variables
        self.designer = StringVar()
        self.designer.set("Designer")
        self.tooltype = StringVar()
        self.tooltype.set("Tool Type")
        self.statusvar = StringVar()
        self.statusvar.set("Order Status")
        self.repeatvar = StringVar()
        self.repeatvar.set("Repeat?")
        self.addtypevar = StringVar()
        self.addtypevar.set("Task")
        self.latevar = StringVar()
        self.latevar.set("Late order?")

        #CSV WF Variables
        self.production_order_number = 0
        self.sales_order_number = 0
        self.date = datetime.date.today()
        self.time = 0
        self.tool_type = ""
        self.tool_designer = ""
        self.order_status = ""
        self.repeat_type = ""
        self.wfcomments = ""
        self.addcomments = ""
        self.task = ""
        self.late = ""

        #CSV Filenames
        self.base = os.path.dirname(os.path.abspath(__file__))
        #Workflow File
        self.wf_file = f"{self.base}/Reports/CSV Files/ETS {self.date} WF"

        #Excel File location
        existing_file = f"{self.base}/Reports/Daily Reports/ETS {datetime.date.today()}.xlsx"
        template_file = f"{self.base}/Reports/Templates/template_report.xlsx"

        #GLOBAL VARIABLES
        self.PAD = 10
        self.FONT = "Courier"
        self.DROPDOWN_SIZE = 16
        self.LABEL_SIZE = 14
        self.INPUT_SIZE = 14
        self.TITLE_SIZE = 20
        self.SUBMISSION_BUTTTON_SIZE = 16
        self.BUTTON_SIZE = 14

        #Adds WF data to WF CSV file
        def add_wf_to_csv(ponumber,salesnumber,designer,date,time,status,ordertype,repeat,comment,late):
            tempdict = {"Production Order Number": ponumber,
                        "Sales Number": salesnumber,
                        "Designer": designer,
                        "Date": date,
                        "Time Taken": time,
                        "Order Status": status,
                        "Tool Type": ordertype,
                        "Repeat?": repeat,
                        "Late": late,
                        "Comments": comment,
                        }

            if os.path.isfile(self.wf_file):
                with open(file=self.wf_file,mode="a",newline="") as file:
                    writer = csv.writer(file)
                    writer.writerow(tempdict.values())
            else:
                with open(file=self.wf_file,mode="w",newline="") as file:
                    writer = csv.writer(file)
                    writer.writerow(tempdict.keys())
                    writer.writerow(tempdict.values())


        def getwfdata():
            po = self.ponumber.get().strip()
            sales = self.salesnumber.get().strip()
            time = self.approxmins.get().strip()
            tool = self.tooltype.get()
            date = self.dateinput.get()
            desginer = self.designer.get()
            status = self.statusvar.get()
            repeat = self.repeatvar.get()
            comment = self.worflowcomments.get("1.0", "end-1c")
            late = self.latevar.get()
            self.production_order_number = po
            self.sales_order_number = sales
            self.time = time
            self.tool_type = tool
            self.date = date
            self.tool_designer = desginer
            self.repeat_type = repeat
            self.order_status = status
            self.wfcomments = str(comment)
            self.late = late

            if self.production_order_number == "":
                tk.messagebox.showerror(title = "Missing PO#",message="Please enter production order number.")
                self.ponumber.focus()
            elif self.sales_order_number == "":
                tk.messagebox.showerror(title="Missing Sales#", message="Please enter sales order number")
                self.salesnumber.focus()
            elif self.time == "":
                tk.messagebox.showerror(title="No Time Input", message="Please enter approximately how long you have spent on the order (minutes) ")
                self.approxmins.focus()
            elif self.tool_type == "Tool Type":
                tk.messagebox.showerror(title="Missing Tool Type", message="Please enter a tool type.")
                self.ordertype.focus()
            elif self.tool_designer == "Designer":
                tk.messagebox.showerror(title = "Missing Designer",message="Please enter designer name.")
                self.userselection.focus()
            elif self.order_status == "Order Status":
                tk.messagebox.showerror(title="Missing Order Status", message="Please state current status of order.")
                self.status_menu.focus()
            elif self.repeat_type == "Repeat?":
                tk.messagebox.showerror(title="Missing Repeat Status", message="Please state if order was a repeat.")
                self.repeat.focus()
            elif self.late == "Late order?":
                tk.messagebox.showerror(title="Missing Late Status", message="Please state if order was cleared from the workflow within 24 hours.")
                self.late_order.focus()

            else:
                add_wf_to_csv(ponumber=self.production_order_number,
                           salesnumber=self.sales_order_number,
                           time=self.time,
                           ordertype=self.tool_type,
                           date=self.date,
                           designer=self.tool_designer,
                           status=self.order_status,
                           repeat=self.repeat_type,
                            late=self.late,
                           comment=self.wfcomments,
                            )
                clearalldata()

        def clearalldata():
            self.tooltype.set("Tool Type")
            self.statusvar.set("Fax")
            self.repeatvar.set("Repeat?")
            self.addtypevar.set("Task")
            self.latevar.set("Late order?")
            self.ponumber.delete("0", tk.END)
            self.salesnumber.delete("0", tk.END)
            self.approxmins.delete("0", tk.END)
            self.worflowcomments.delete("1.0", "end-1c")
            self.root.focus()

        def new_report():
            confirmation = tk.messagebox.askyesno(title="Confirm report creation", message=f"Create report for {self.date}?")
            if confirmation:
                try:
                    if os.path.isfile(existing_file):
                        Writer.generate_report(template=template_file,data=self.wf_file,filepath=existing_file)
                        tk.messagebox.showinfo(title="Success",message=f"Report for {self.date} already exists. Report has been updated.")
                    else:
                        Writer.generate_report(template=template_file,data=self.wf_file,filepath=existing_file)
                        tk.messagebox.showinfo(title="Success", message=f"New report generated for {self.date}")

                except PermissionError: tk.messagebox.showerror(title="Report already open",message=f"Report for {self.date} already exists and is open. Please close before trying to update.")

        def confirm_user_selection(user):
            temp_user = ""

            if user == "GD":
                temp_user = "Garrat Duke"
            elif user == "JB":
                temp_user = "James Bagwell"
            elif user == "SRS":
                temp_user = "Scott Stevens"
            elif user == "ODM":
                temp_user = "Owen Mcdonald"

            answer = tk.messagebox.askyesno(title="Confirm User Selection", message=f"Confirm {temp_user} as designer?")

            if not answer:
                self.designer.set("Designer")

        #Labels
        #TITLE LABELS
        self.workflowlabel = tk.Label(self.root,text="Workflow Order",font=(self.FONT,self.TITLE_SIZE,"bold"))

        self.ponumberlabel = tk.Label(self.root,text="Production #: ", font=(self.FONT,self.LABEL_SIZE))
        self.salesnumberlabel = tk.Label(self.root,text="Sales #: ", font=(self.FONT,self.LABEL_SIZE))
        self.timespentlabel = tk.Label(self.root,text="Time Spent: ",justify="right", font=(self.FONT,self.LABEL_SIZE))
        self.datelabel = tk.Label(self.root,text="Date: ", font=(self.FONT,self.LABEL_SIZE))
        self.additionalcommentslabel = tk.Label(self.root,text="Comment: ", font=(self.FONT,self.LABEL_SIZE))
        self.workflowcommmentlabel = tk.Label(self.root,text="Comment: ", font=(self.FONT,self.LABEL_SIZE))

        #Input Box's
        self.ponumber = tk.Entry(justify="center", font=(self.FONT,self.INPUT_SIZE))
        self.salesnumber = tk.Entry(justify="center", font=(self.FONT,self.INPUT_SIZE))
        self.approxmins = tk.Entry(justify="center", font=(self.FONT,self.INPUT_SIZE))
        self.dateinput = tk.Entry(justify="center", font=(self.FONT,self.INPUT_SIZE),width=10)
        self.dateinput.insert("0",f"{self.date}")
        self.worflowcomments = tk.Text(height=1, width=1, font=(self.FONT,self.INPUT_SIZE ))

        #Dropdowns
        self.userselection = tk.OptionMenu(self.root, self.designer, "JB", "GD","SRS","ODM",command=confirm_user_selection)
        self.userselection.config(width=10,font=(self.FONT,self.DROPDOWN_SIZE))
        self.ordertype = tk.OptionMenu(self.root,self.tooltype,"2D","Form")
        self.ordertype.config(width = 15,font=(self.FONT,self.DROPDOWN_SIZE))
        self.status_menu = tk.OptionMenu(self.root, self.statusvar, "FFA ", "Acknowledgement", "Design Request")
        self.status_menu.config(font=(self.FONT,self.DROPDOWN_SIZE))
        self.repeat = tk.OptionMenu(self.root, self.repeatvar, "Yes", "No")
        self.repeat.config(font=(self.FONT,self.DROPDOWN_SIZE))
        self.late_order = tk.OptionMenu(self.root,self.latevar,"Late", "On Time")
        self.late_order.config(font=(self.FONT,self.DROPDOWN_SIZE))

        #Buttons
        #Submission Buttons
        self.submitworkflow = tk.Button(text="Submit Workflow Entry", font=(self.FONT,self.SUBMISSION_BUTTTON_SIZE),command=getwfdata)

        #Other Buttons
        self.resetbutton = tk.Button(text="Clear All",justify="center", font=(self.FONT,self.BUTTON_SIZE,),width=10,command=clearalldata)
        self.new_report = tk.Button(text="Generate Daily Report",justify="center", font=(self.FONT,self.BUTTON_SIZE),command=new_report)

        #Layout
        self.workflowlabel.grid(row=0,column=0,columnspan=3,rowspan=2,sticky="nesw",padx=self.PAD,pady=self.PAD)
        self.ponumberlabel.grid(row=2,column=0,sticky="nesw",padx=self.PAD,pady=self.PAD)
        self.ponumber.grid(row=2,column=1,sticky="nesw",padx=self.PAD,pady=self.PAD)
        self.salesnumberlabel.grid(row=3,column=0,sticky="nesw",padx=self.PAD,pady=self.PAD)
        self.salesnumber.grid(row=3,column=1,sticky="nesw",padx=self.PAD,pady=self.PAD)
        self.workflowcommmentlabel.grid(row=4,column=0,sticky="nesw",padx=self.PAD,pady=self.PAD)
        self.worflowcomments.grid(row=4,column=1,sticky="nesw",padx=self.PAD,pady=self.PAD)
        self.status_menu.grid(row=2,column=2,sticky="nesw",padx=self.PAD,pady=self.PAD)
        self.repeat.grid(row=4,column=2,sticky="nesw",padx=self.PAD,pady=self.PAD)
        self.ordertype.grid(row=3,column=2,sticky="nesw",padx=self.PAD,pady=self.PAD)
        self.datelabel.grid(row=0,column=4,sticky="e",padx=self.PAD,pady=self.PAD)
        self.dateinput.grid(row=0,column=5,sticky="nesw",padx=self.PAD,pady=self.PAD)
        self.userselection.grid(row=1,column=5,sticky="nesw",padx=self.PAD,pady=self.PAD)
        self.timespentlabel.grid(row=5,column=0,sticky="news",padx=self.PAD,pady=self.PAD)
        self.approxmins.grid(row=5,column=1,sticky="nesw",padx=self.PAD,pady=self.PAD)
        self.submitworkflow.grid(row=4,column=4,sticky="nse",columnspan=2,rowspan=1,padx=self.PAD,pady=self.PAD)
        self.resetbutton.grid(row=2,column=5,sticky="nesw",padx=self.PAD,pady=self.PAD)
        self.new_report.grid(row=5, column=4,columnspan=2,rowspan=1, sticky="nse", padx=self.PAD, pady=self.PAD)
        self.late_order.grid(row=5,column=2,sticky="news",pady=self.PAD,padx=self.PAD)

        self.root.mainloop()

EtsGui()








