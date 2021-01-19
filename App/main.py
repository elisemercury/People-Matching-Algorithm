import tkinter as tk
from tkinter import ttk 
from tkinter import messagebox
from tkinter import filedialog
from tkinter import *
from ttkthemes import ThemedTk
import textwrap
import time
import algorithm
from datetime import date
import os

global filename_mentee, filename_mentor
filename_mentee = None
filename_mentor = None

# main window
root = ThemedTk(theme="plastik") #arc, clearlooks, keramik (elegance)
root.title("MMM Tool - Mentee/Mentor Matcher") 
root.minsize(580, 450)
tabControl = ttk.Notebook(root) 

photo = PhotoImage(file = "icon.png")
root.iconphoto(False, photo)
root.iconbitmap(default='icon.png')
  
tab1 = ttk.Frame(tabControl) 

tab2 = ttk.Frame(tabControl) 
tab2.place(x = 0, y = 0, relwidth = 1, relheight = 1)

tabControl.add(tab1, text ='Upload') 
tabControl.add(tab2, text ='Help') 
tabControl.pack(expand = 5, fill ="both") 


############################# tab 1 ###########################
# tab 1 title
lbl = tk.Label(tab1, text="Hello, welcome to the Mentee/Mentor Matcher.")#, bg ="gray90")
lbl.config(font=("Arial", 13))
lbl.grid(column=0, row=0, padx=15, pady=5, columnspan=2, sticky="W")

ttk.Separator(tab1).place(x=0, y=35, relwidth=1)

# tab 1 subtitle
lbl = tk.Label(tab1, text="To start, please upload mentor and the mentee survey data.")
lbl.config(font=("Arial", 10))
lbl.grid(column=0, row=1, padx=15, pady=12, columnspan=2, sticky="W")


# upload mentee survey
lbl = tk.Label(tab1, text="Step 1:")
lbl.config(font=("Arial", 12, "bold"))
lbl.grid(column=0, row=2, padx=15, pady=0, columnspan=1, sticky="W")

lbl = tk.Label(tab1, text="Please upload your mentee survey data file:")
lbl.grid(column=0, row=3, padx=15, pady=0, columnspan = 1, sticky="W")
lbl.config(font=("Arial", 10))

mentee_file_txt = tk.StringVar()
mentee_file_deposit = tk.Label(tab1, textvariable=mentee_file_txt, bg ="gray100", font=("Arial", 10))
mentee_file_deposit.grid(column=0, row=4, padx=21, pady=10, columnspan=2, sticky="W")
mentee_file_txt.set("... /path/mentee_file.xlsx")

btn = tk.Button(tab1, text="Browse...", command = UploadAction_mentee)
btn.grid(column=0, row=5, padx=16, sticky="W")


# upload mentor survey
lbl = tk.Label(tab1, text="Step 2:")
lbl.config(font=("Arial", 12, "bold"))
lbl.grid(column=1, row=2, padx=15, pady=0, columnspan=1, sticky="W")

lbl = tk.Label(tab1, text="Please upload your mentor survey data file:")
lbl.grid(column=1, row=3, padx=15, pady=0, columnspan = 1, sticky="W")
lbl.config(font=("Arial", 10))

mentor_file_txt = tk.StringVar()
mentor_file_deposit = tk.Label(tab1, textvariable=mentor_file_txt, bg ="gray100", font=("Arial", 10))
mentor_file_deposit.grid(column=1, row=4, padx=21, pady=10, columnspan = 2, sticky="W")
mentor_file_txt.set("... /path/mentor_file.xlsx")

btn2 = tk.Button(tab1, text="Browse...", command = UploadAction_mentor)
btn2.grid(column=1, row=5, padx=21, sticky="W")


# optional config
lbl = tk.Label(tab1, text="\nOptional:")
lbl.config(font=("Arial", 12, "bold"))
lbl.grid(column=0, row=10, padx=15, pady=0, columnspan=2, sticky="W")
# checkbutton make unique
makeUnique = IntVar(value=True)
tk.Checkbutton(tab1, text="Make first rank unique", 
               variable=makeUnique).grid(row=11, padx=21, pady=10, sticky="W", columnspan = 2)
# spinbox number of ranks
lbl = tk.Label(tab1, text="Number of ranks:")
lbl.grid(column=0, row=12, padx=21, pady=0, columnspan = 2, sticky="W")
lbl.config(font=("Arial", 11))
ranking = StringVar(tab1)
ranking.set("5")
spin = Spinbox(tab1, from_= 1, to = 10, width=15, textvariable=ranking)  
spin.grid(column=0, row=12, padx=150, columnspan=2, sticky="W")

empty_row = tk.Label(tab1, text=" ")
empty_row.grid(column=0, row=13, padx=15, pady=7, columnspan = 2, sticky="W")


# progress bar
progBar = set_progBar(length=260)

# please wait placeholder
global wait_txt
wait_txt = tk.StringVar()
wait_deposit = tk.Label(tab1, textvariable=wait_txt)
wait_deposit.grid(column=0, row=14, padx=21, sticky="W", columnspan=2)
wait_txt.set("        ")

# submit
btn2 = tk.Button(tab1, text="Submit & Run", command=lambda:[clear_output(), Action_submit()])#.place(x=435, y=430)
btn2.grid(column=0, row=14, padx=20, sticky="E", columnspan=2)
# reset
lbl = tk.Label(tab1, text="Reset")
lbl.grid(column=0, row=15, padx=20, sticky="NE", columnspan=2)
#lbl.place(x=505, y=460)
lbl.config(font=("Arial", 8, "underline"))
lbl.bind( "<Button>", reset ) 

# output message
#global success_msg, success_msg_deposit, output_file, output_file_deposit

output_msg = tk.StringVar()
output_msg_deposit = tk.Label(tab1, textvariable=output_msg, font=("Arial", 11))
output_msg_deposit.grid(column=0, row=17, padx=21, pady=0, columnspan = 2, sticky="W")
output_msg.set("")
# open file
output_file = tk.StringVar()
output_file_deposit = tk.Label(tab1, textvariable=output_file, font=("Arial", 10))
output_file_deposit.grid(column=0, row=18, padx=21, pady=0, columnspan = 2, sticky="W")
output_file.set("")

############################# tab 2 ###########################
# tab 2 title
lbl = tk.Label(tab2, text="Help")#, bg ="gray90")
lbl.config(font=("Arial", 13))
lbl.grid(column=0, row=0, padx=15, pady=5, columnspan=2, sticky="W")

ttk.Separator(tab2).place(x=0, y=35, relwidth=1)

# description
lbl = tk.Label(tab2, text=str("""Upload your mentor and mentee survey data into the MMM tool to help find the best pairings based on the answers given in the survey data."""),
               wraplength=550, justify=LEFT)
lbl.grid(column=0, row=1, padx=15, pady=10, rowspan = 1, sticky="W")
lbl.config(font=("Arial", 10))

# generate template
lbl = tk.Label(tab2, text="Uploading survey data:")
lbl.config(font=("Arial", 12, "bold"))
lbl.grid(column=0, row=2, padx=15, pady=0, columnspan=1, sticky="W")

lbl = tk.Label(tab2, text=str("""The MMM supports only a standardized .xlsx file format. The column names in the template must match the column names in the uploaded files (case sensitive).\n 
You can generate template files below:"""), wraplength=550, justify=LEFT)
lbl.grid(column=0, row=3, padx=15, pady=0, rowspan = 1, sticky="W")
lbl.config(font=("Arial", 10))

btn2 = tk.Button(tab2, text="Generate template file", command = download_template)
btn2.grid(column=0, row=4, padx=21, pady=5, sticky="W")

empty_row = tk.Label(tab2, text=" ")
empty_row.grid(column=0, row=5, padx=15, pady=0, columnspan = 2, sticky="W")

empty_row = tk.Label(tab2, text=" ")
empty_row.grid(column=0, row=6, padx=15, pady=0, columnspan = 2, sticky="W")


# view example survey
lbl = tk.Label(tab2, text="Conducting survey:")
lbl.config(font=("Arial", 12, "bold"))
lbl.grid(column=0, row=8, padx=15, pady=5, columnspan=1, sticky="W")

lbl = tk.Label(tab2, text=str(
"""You can view an example of a mentor/mentee survey below: """),
               wraplength=390, justify=LEFT)
lbl.grid(column=0, row=9, padx=15, pady=5, rowspan = 1, sticky="W")
lbl.config(font=("Arial", 10))

btn2 = tk.Button(tab2, text="View example survey", command = open_survey_pdf)
btn2.grid(column=0, row=10, padx=21, pady=5, sticky="W")

#  version
lbl = tk.Label(tab2, text=str("""v.1.0"""),
               wraplength=550, justify=LEFT)
lbl.grid(column=0, row=11, padx=15, pady=5, rowspan = 1, sticky="E")
lbl.config(font=("Arial", 8))




# run 
root.geometry("420x505")  
root.mainloop()   

def UploadAction_mentee():
    global filename_mentee
    filename_mentee = filedialog.askopenfilename(title="Select mentee survey data")
    if len(filename_mentee) == 0:
        return
    if ".xlsx" in filename_mentee:
        if len(filename_mentee) > 50:
            mentee_file_txt.set("... " + filename_mentee[50:])
        else:
            mentee_file_txt.set(filename_mentee)
    else:
        tk.messagebox.showerror("Whoops!", "Please upload a file in .xlsx Excel format.")

def UploadAction_mentor(event=None):
    global filename_mentor
    filename_mentor = filedialog.askopenfilename(title="Select mentor survey data")
    if len(filename_mentor) == 0:
        return
    if ".xlsx" in filename_mentor:
        if len(filename_mentor) > 50:
            mentor_file_txt.set("... " + filename_mentor[50:])
        else:
            mentor_file_txt.set(filename_mentor)
    else:
        tk.messagebox.showerror("Whoops!", "Please upload a file in .xlsx Excel format.")

def Action_submit(event=None):   
    # check if files selected
    if filename_mentee == None or filename_mentor == None:
        tk.messagebox.showerror("Whoops!", "Please upload files before submitting.")    
        return

    wait_txt.set("Please wait...")
   
    try:
        if int(spin.get()) > 10:
            set_progBar(length=260)
            wait_txt.set("")
            return tk.messagebox.showerror("Whoops!", "Maximum number of ranks supported is 10. Please choose a value equal to or smaller than 10.")    
    except: 
        set_progBar(length=260)
        wait_txt.set("")
        return tk.messagebox.showerror("Whoops!", "Number of ranks should be a number between 1 and 10.")    

    if makeUnique.get() == 0:
        assignment = algorithm.run_algorithm(filename_mentee, filename_mentor, ranks=int(spin.get()), unique=False, progbar_window=tab1)
    else:
        assignment = algorithm.run_algorithm(filename_mentee, filename_mentor, ranks=int(spin.get()), progbar_window=tab1)
    
    if "Failed" in assignment:
        tk.messagebox.showerror("Whoops!", "Something went wrong. " + str(assignment))    
        wait_txt.set("            ")
        wait_deposit.grid(column=0, row=15, padx=21, sticky="W", columnspan=2)
        set_progBar(length=260)
        return
    else: 
        wait_txt.set("Done!")
        
    if assignment == "Not unique":
        output_msg.set("Warning: could not create unique ranking. Please try again.")
    elif assignment == "Success not unique":
        output_msg.set("Success: ranking successfully created.")
    else:
        output_msg.set("Success: unique ranking successfully created.")
    
    today = date.today()
    today = today.strftime("%d%m%Y")
    global output_file_name    
    output_file_name = "MMM_ranking_" + str(today) + ".xlsx"
    output_file.set("File name: " + output_file_name)
    
    global open_file_button
    open_file_button = tk.Button(tab1, text="Open file", command = open_output)
    open_file_button.grid(column=0, row=18, padx=200, sticky="E", columnspan=2)
               
                        
def set_progBar(col=0, row=14, length=260):
    bar = ttk.Progressbar(tab1, length=length, style='grey.Horizontal.TProgressbar')
    bar['value'] = 0
    bar.grid(column=col, row=row, padx=150, sticky="W", columnspan=2)
    return bar

def update_progbar(bar, value):
    bar['value'] = value
    tab1.update_idletasks() 
    time.sleep(1)        
    
def download_template():
    template = algorithm.create_template()   
    file_info = tk.Label(tab2, text="Template files generated in: ")
    file_info.config(font=("Arial", 10))
    file_info.grid(column=0, row=5, padx=15, pady=1, columnspan = 2, sticky="W")
    
    file_link_txt = str(os.getcwd())
    if len(file_link_txt) > 200:
        file_link = tk.Label(tab2, text=file_link_txt[100:])
    else:
        file_link = tk.Label(tab2, text=file_link_txt)
    file_link.config(font=("Arial", 10, "underline"))
    file_link.grid(column=0, row=6, padx=21, pady=0, columnspan = 2, sticky="W")
    file_link.bind( "<Button>", open_folder)    
        
    print("Placeholder")
    
def reset(event=None):
    global filename_mentee, filename_mentor
    filename_mentee = None
    filename_mentor = None
    mentee_file_txt.set("...\path\mentee_file.xlsx")
    mentor_file_txt.set("...\path\mentor_file.xlsx")
    wait_txt.set("")
    output_msg.set("")
    output_file.set("")
    open_file_button.destroy()
    set_progBar()
    tab1.update_idletasks() 
    
def clear_output(event=None):
    try:
        wait_txt.set("")
        output_msg.set("")
        output_file.set("")
        open_file_button.destroy()
    except:
        return
    
def open_output(event=None):
    folder = os.getcwd()
    open_link = str(folder) + "\\" + str(output_file_name)
    os.startfile(open_link, 'open')
                         
def open_folder(event=None):
    folder = os.getcwd()
    open_link = str(folder)
    os.startfile(open_link, 'open')                 
    print("opened")
    
def open_survey_pdf(event=None):
    folder = os.getcwd()
    open_link = str(folder) + "\\MMM Survey Templates.pdf"
    os.startfile(open_link, 'open')    