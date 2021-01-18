import pandas as pd
import scipy.spatial
import random
import xlsxwriter
from datetime import date
import time
import tkinter as ttk
from tkinter import *
import tkinter as tk

def read_file(xlsx_file, group=None):
    if group == "mentor":
        df_mentor_raw = pd.read_excel(xlsx_file, sheet_name=0)
        if len(df_mentor_raw) == 2:
            df_mentor_raw = pd.read_excel(xlsx_file, sheet_name=1)
        if df_mentor_raw.shape != (0, 0):
            return df_mentor_raw
        else:
            return "Failed"
            # error: no data found within your file
    elif group == "mentee":
        df_mentee_raw = pd.read_excel(xlsx_file, sheet_name=0)
        if len(df_mentee_raw) == 2:
            df_mentee_raw = pd.read_excel(xlsx_file, sheet_name=1)
        if df_mentee_raw.shape != (0, 0):
            return df_mentee_raw
        else:
            return "Failed"
            # error: no data found within your file  

def extract_choices(df_col):
    choices = []
    cols = set()
    data = set(df_col)
    for choice in data:
        choice = choice.split(";")
        choices.append(choice)
    for innerlist in choices:
        for choice in innerlist:
            if choice != "":
                cols.add(choice)    
        
    return list(cols)

def encode_onehot(df_input, df_output):
    count = 0
    for answer in df_output.columns:
        for column in df_input.columns:
            if column in ["What are you most likely to give in this mentoring engagement?", "What are your general strengths?", "What are your personal interests?", "What are you most likely to want out of this mentoring engagement?", "What are the strengths you would like to further develop?"]:
                for result in df_input[column]:
                    if not pd.isna(result):
                        if type(result) == str:
                            if answer in result: 
                                df_output.loc[df_output.index[count], answer] = 1                           
                            if count == (len(df_input.index)-1):
                                count = 0
                            else:
                                count += 1

    df_output = df_output.fillna(0)
    return df_output

def add_missing(columnlist1, columnlist2):
    for choice in columnlist1:
            if choice not in columnlist2:
                columnlist2.append(choice)
    for choice in columnlist2:
            if choice not in columnlist1:
                columnlist1.append(choice)

def calculate_similarity(df_onehot_mentor, df_onehot_mentee):
    jac = scipy.spatial.distance.cdist(df_onehot_mentor.iloc[:,1:], df_onehot_mentee.iloc[:,1:], metric='jaccard')
    people_similarity = pd.DataFrame(jac, columns=df_onehot_mentee.index.values, index=df_onehot_mentor.index.values)
    return people_similarity

# function that assigns the list of mentors ranked by similarity to each mentee in a dict
def create_cluster(df):
    clusters = {}
    for mentee in df.columns:
        mentee_cluster = df[mentee].nsmallest(len(df))
        data = {mentee : list(mentee_cluster.index)}
        clusters.update(data)
    return clusters

def prepare_data():
    # theatre data
    df_mentor_theatre = df_mentor_raw[df_mentor_raw["Would you prefer a mentee in the same theatre?"]!="No preference"]
    df_mentee_theatre = df_mentee_raw[df_mentee_raw["If possible, would you prefer a mentor in the same theatre?"]!="No preference"]

    # role data
    df_mentor_role = df_mentor_raw[df_mentor_raw["Would you prefer a mentee in a:"]!="No preference"]
    df_mentee_role = df_mentee_raw[df_mentee_raw["If possible, would you prefer a mentor in a:"]!="No preference"]

def filter_by_pref(mentee, filter_by):
    if filter_by == "Theatre":
        mentee_theatre = df_mentee_raw["Your theatre?"].loc[mentee]
        for mentor in df_mentor_raw.index:
            mentor_theatre = df_mentor_raw["Your theatre?"].loc[mentor] 
            if mentee_theatre != mentor_theatre:
                try:
                    people_cluster[mentee].remove(mentor)
                except:
                    continue
    elif filter_by == "Role":
        mentee_role = df_mentee_raw["If possible, would you prefer a mentor in a:"].loc[mentee]
        for mentor in df_mentor_raw.index:
            mentor_role = df_mentor_raw["Your role?"].loc[mentor]
            if mentee_role != mentor_role:
                try:
                    people_cluster[mentee].remove(mentor)
                except:
                    continue
    elif filter_by == "Segment":
        mentee_segment = df_mentee_raw["If possible, which segment would you like your mentor to be from?"].loc[mentee]
        for mentor in df_mentor_raw.index:
            mentor_segment = df_mentor_raw["Your segment?"].loc[mentor]
            if mentee_segment.lower() not in mentor_segment.lower():
                try:
                    people_cluster[mentee].remove(mentor)
                except:
                    continue
    elif filter_by == "Gender":
        mentee_gender = df_mentee_raw["Your gender?"].loc[mentee]
        for mentor in df_mentor_raw.index:
            mentor_gender = df_mentor_raw["Your gender?"].loc[mentor]
            if mentee_gender != mentor_gender:
                try:
                    people_cluster[mentee].remove(mentor)
                except:
                    continue   

def mentor_filter(col_ask_preference, col_ask_own, method="yes"):
    for mentor in df_mentor_raw.index:
        if method == "yes":
            if df_mentor_raw[col_ask_preference].loc[mentor]== "Yes":
                mentor_theatre = df_mentor_raw[col_ask_own].loc[mentor]
                for mentee in people_cluster:
                    mentee_theatre = df_mentee_raw[col_ask_own].loc[mentee]
                    if mentor_theatre != mentee_theatre:
                        try:
                            people_cluster[mentee].remove(mentor)
                        except:
                            continue 
        elif method == "no preference":                         
            if df_mentor_raw[col_ask_preference].loc[mentor]!= "No preference":
                mentor_theatre = df_mentor_raw[col_ask_own].loc[mentor]
                for mentee in people_cluster:
                    mentee_theatre = df_mentee_raw[col_ask_own].loc[mentee]
                    if mentor_theatre != mentee_theatre:
                        try:
                            people_cluster[mentee].remove(mentor)
                        except:
                            continue 
# test to check if each mentee has at least 1 metor assigned to it
def test_assignment(info="default"):
    for mentee in people_cluster:
        if len(people_cluster[mentee]) < 1:
            raise ValueError(mentee, "has less than 1 mentor assigned! Info:", info)

def sort_unique(cluster, max_rounds=50, method="limited"):
    i = 0
    j = 1
    count = []
    for number in range(0,max_rounds):
        count.append(0)
    rounds = 0
    mentees = list(cluster.keys())
    while rounds < max_rounds:
        while i < len(cluster):
            while j < len(cluster):
                to_compare = list(zip(cluster[mentees[i]], cluster[mentees[j]]))[0]
                if to_compare[0] == to_compare[1]:
                    count[rounds] += 1
                    if rounds > 1:
                        if count[rounds-1] <= count[rounds-2]:
                            if rounds % 5 == 0:
                                random_position = random.randint(0, 1)
                                if random_position == 0:
                                    try:
                                        cluster[mentees[j]][0], cluster[mentees[j]][2] = cluster[mentees[j]][2], cluster[mentees[j]][0]
                                        print("1")
                                    except:
                                        continue
                                else:
                                    try:
                                        cluster[mentees[i]][0], cluster[mentees[i]][2] = cluster[mentees[i]][2], cluster[mentees[i]][0]
                                        print("2")          
                                    except:
                                        continue
                            else:
                                random_position = random.randint(0, 1)
                                if random_position == 0:
                                    try:
                                        cluster[mentees[j]][0], cluster[mentees[j]][1] = cluster[mentees[j]][1], cluster[mentees[j]][0]
                                        print("3")
                                    except:
                                        continue
                                else:
                                    try:
                                        cluster[mentees[i]][0], cluster[mentees[i]][1] = cluster[mentees[i]][1], cluster[mentees[i]][0]
                                        print("4")
                                    except: 
                                        continue
                        else:
                            if method == "limited":
                                random_position = random.randint(0, 1)
                                if random_position == 0:
                                    try:
                                        cluster[mentees[j]][0], cluster[mentees[j]][3] = cluster[mentees[j]][3], cluster[mentees[j]][0]
                                        print("5")
                                    except:
                                        continue
                                else:
                                    try:    
                                        cluster[mentees[i]][0], cluster[mentees[i]][3] = cluster[mentees[i]][3], cluster[mentees[i]][0] 
                                        print("6")
                                    except:
                                        continue
                            elif method == "medium":
                                random_position = random.randint(0, 1)
                                if random_position == 0:
                                    try:
                                        cluster[mentees[j]][2], cluster[mentees[j]][4] = cluster[mentees[j]][4], cluster[mentees[j]][2]
                                        print("7")
                                    except:
                                        continue
                                else:
                                    try:
                                        cluster[mentees[i]][2], cluster[mentees[i]][4] = cluster[mentees[i]][4], cluster[mentees[i]][2] 
                                        print("8")
                                    except:
                                        continue                      
                            elif method == "hard":
                                random_position = random.randint(0, 1)                            
                                if len(cluster[mentees[i]]) > 5 or len(cluster[mentees[j]]) > 5:
                                    random_position = random.randint(0, 1)
                                    if random_position == 0:
                                        try:
                                            cluster[mentees[j]][3], cluster[mentees[j]][5] = cluster[mentees[j]][5], cluster[mentees[j]][3]
                                            print("9")
                                        except:
                                            cluster[mentees[i]][3], cluster[mentees[i]][5] = cluster[mentees[i]][5], cluster[mentees[i]][3]
                                            print("10")
                                    else:
                                        try:
                                            cluster[mentees[i]][3], cluster[mentees[i]][5] = cluster[mentees[i]][5], cluster[mentees[i]][3]  
                                            print("11")
                                        except:
                                            cluster[mentees[j]][3], cluster[mentees[j]][5] = cluster[mentees[j]][5], cluster[mentees[j]][3] 
                                            print("12")
                                else:
                                    random_position = random.randint(0, 1)
                                    if random_position == 0:
                                        cluster[mentees[j]][0], cluster[mentees[j]][-1] = cluster[mentees[j]][-1], cluster[mentees[j]][0]
                                        print("13")
                                    else:
                                        cluster[mentees[i]][0], cluster[mentees[i]][-1] = cluster[mentees[i]][-1], cluster[mentees[i]][0] 
                                        print("14")
                    else:
                        random_position = random.randint(0, 1)
                        if random_position == 0:
                            cluster[mentees[j]][0], cluster[mentees[j]][1] = cluster[mentees[j]][1], cluster[mentees[j]][0]
                            print("15")
                        else:
                            cluster[mentees[i]][0], cluster[mentees[i]][1] = cluster[mentees[i]][1], cluster[mentees[i]][0]   
                            print("16")             
                j+=1
            i += 1
            j = i+1
        if count[rounds] > 0:
            if rounds == max_rounds-1:
                return "Failed"
            else:
                rounds += 1
                i = 0
                j = 1
        else:
            break
    return "Success"

def filter_mentee_mind(df_mentee_raw, df_mentor_raw, poeple_cluster):
    pairs = {}
    for mentee in df_mentee_raw.index:
        if mentee in list(df_mentor_raw["Let us know if you already have a CSAP mentee in mind"]):
            mentor = df_mentor_raw.index[df_mentor_raw["Let us know if you already have a CSAP mentee in mind"] == mentee]
            pairs[mentee] = mentor[0]  
    
    for mentee_mind in pairs:
        mentor_mind = pairs[mentee_mind]
        for mentee in poeple_cluster:
            if mentor_mind in poeple_cluster[mentee]:
                poeple_cluster[mentee].remove(mentor_mind)
            if mentee == mentee_mind:
                poeple_cluster[mentee].insert(0, mentor_mind)

def export_excel(people_cluster, ranks):
    # Create an new Excel file and add a worksheet.    
    today = date.today()
    today = today.strftime("%d%m%Y")
    workbook = xlsxwriter.Workbook('MMM_ranking_' + today + ".xlsx")
    worksheet = workbook.add_worksheet()

    # format the sheet
    worksheet.set_column('A:A', 30)
    worksheet.set_column('B:K', 20)
    # mentee cell (0,0)
    mentee_cell = workbook.add_format({'bold': True})
    mentee_cell.set_right()
    mentee_cell.set_bottom()
    # vertical
    vertical = workbook.add_format({'bold': True})
    vertical.set_right()
    # horzontal
    horizontal = workbook.add_format({'bold': True})
    horizontal.set_bottom()

    # Write "Mentee" cell
    worksheet.write(0, 0, "Mentee", mentee_cell)

    # Write mentee names
    i = 1
    for mentee in people_cluster:
        worksheet.write(i, 0, str(mentee), vertical)
        i+=1

    # Write ranks
    i = 1
    for j in range(0,ranks):
        worksheet.write(0, i, str("Rank " + str(i)), horizontal)
        i+=1

    # Write mentor names
    row = 1
    for mentee in people_cluster:
        mentors = people_cluster[mentee]
        col = 1
        for mentor in mentors:
            worksheet.write(row, col, mentor)
            col+=1
        row +=1
        
    workbook.close()
def create_template():
    # MENTEE template
    workbook = xlsxwriter.Workbook("MMM_template_MENTEE.xlsx")
    worksheet = workbook.add_worksheet()
    
    worksheet.set_column('A:A', 10)
    worksheet.set_column('B:M', 30)

    cell_format = workbook.add_format({'bold': True})
    cell_format.set_bottom()

    mentee_columns = ['Name',
                    'What are you most likely to want out of this mentoring engagement?',
                    'What are the strengths you would like to further develop?',
                    'What are your personal interests?', 'Your role?', 'Your theatre?',
                    'Your gender?',
                    'If possible, which segment would you like your mentor to be from?',
                    'If possible, would you prefer a mentor in a:',
                    'If possible, what is your first priority in matching you with a mentor?']

    count = 0
    while count < len(mentee_columns):
        worksheet.write(0, count, mentee_columns[count], cell_format)
        count += 1
    workbook.close()

    # MENTOR template
    workbook = xlsxwriter.Workbook("MMM_template_MENTOR.xlsx")
    worksheet = workbook.add_worksheet()

    worksheet.set_column('A:A', 10)
    worksheet.set_column('B:M', 30)

    cell_format = workbook.add_format({'bold': True})
    cell_format.set_bottom()

    mentor_columns = ['Name',
                    'What are you most likely to give in this mentoring engagement?',
                    'What are your general strengths?', 'What are your personal interests?',
                    'Your role?', 'Your theatre?', 'Your segment?', 'Your gender?',
                    'Would you prefer a mentee in a:',
                    'Would you prefer a mentee in the same theatre?',
                    'Would you prefer a mentee of the same gender?',
                    'Let us know if you already have a CSAP mentee in mind']

    count = 0
    while count < len(mentor_columns):
        worksheet.write(0, count, mentor_columns[count], cell_format)
        count += 1

    workbook.close()

def update_progressbar(progbar_window, value):
    bar = ttk.Progressbar(progbar_window, length=260, style='grey.Horizontal.TProgressbar')
    bar['value'] = value
    bar.grid(column=0, row=14, padx=150, sticky="W", columnspan=2)
    progbar_window.update_idletasks() 
    time.sleep(0.2)        

def run_algorithm(mentee_file, mentor_file, ranks, unique=True, progbar_window=None):
    global df_mentor_raw, df_mentee_raw
    df_mentee_raw = read_file(mentee_file, group="mentee")
    if type(df_mentee_raw) == str:
        tk.messagebox.showerror("Whoops!", "No data found within the uploaded mentee data file. Make sure the data is located on the first sheet of the .xlsx file, or select a different file.")
        return    
    df_mentee_raw = pd.DataFrame(df_mentee_raw).set_index('Name', drop=True)
    
    df_mentor_raw = read_file(mentor_file, group="mentor")
    if type(df_mentor_raw) == str:
        tk.messagebox.showerror("Whoops!", "No data found within the uploaded mentor data file. Make sure the data is located on the first sheet of the .xlsx file, or select a different file.")
        return
    df_mentor_raw = pd.DataFrame(df_mentor_raw).set_index('Name', drop=True)
    try:
        give_mentor = extract_choices(df_mentor_raw["What are you most likely to give in this mentoring engagement?"])
        strengths_mentor = extract_choices(df_mentor_raw["What are your general strengths?"])
        interests_mentor = extract_choices(df_mentor_raw["What are your personal interests?"])

        want_mentee = extract_choices(df_mentee_raw["What are you most likely to want out of this mentoring engagement?"])
        strengths_mentee = extract_choices(df_mentee_raw["What are the strengths you would like to further develop?"])
        interests_mentee = extract_choices(df_mentee_raw["What are your personal interests?"])
    except:
        return "Failed: error 01"
    add_missing(give_mentor, want_mentee)
    add_missing(strengths_mentor, strengths_mentee)
    add_missing(interests_mentor, interests_mentee)

    try:
        # 1 | mentor likely to give
        df_give_mentor = encode_onehot(df_mentor_raw, pd.DataFrame([], index=df_mentor_raw.index, columns=give_mentor))
        # 2 | mentor general strengths
        df_strengths_mentor = encode_onehot(df_mentor_raw, pd.DataFrame([], index=df_mentor_raw.index, columns=strengths_mentor))
        # 3 | mentor personal interests
        df_interests_mentor = encode_onehot(df_mentor_raw, pd.DataFrame([], index=df_mentor_raw.index, columns=interests_mentor))
    except:
        return "Failed: error 02"
    # combine all three dataframes
    df_onehot_mentor = pd.concat([df_give_mentor, df_strengths_mentor,df_interests_mentor], axis=1)
    del df_give_mentor, df_strengths_mentor, df_interests_mentor

    if progbar_window!=None:
        update_progressbar(progbar_window, 20)
    try:
        # 4 | mentee want 
        df_want_mentee = encode_onehot(df_mentee_raw, pd.DataFrame([], index=df_mentee_raw.index, columns=want_mentee))
        # 5 | mentee strengths to develop
        df_strengths_mentee = encode_onehot(df_mentee_raw, pd.DataFrame([], index=df_mentee_raw.index, columns=strengths_mentee))
        # 6 | mentee personal interests
        df_interests_mentee = encode_onehot(df_mentee_raw, pd.DataFrame([], index=df_mentee_raw.index, columns=interests_mentee))
    except:
        return "Failed: error 03"
    # combine all three dataframes
    df_onehot_mentee = pd.concat([df_want_mentee, df_strengths_mentee, df_interests_mentee], axis=1)
    del df_want_mentee, df_strengths_mentee, df_interests_mentee

    if progbar_window!=None:
        update_progressbar(progbar_window, 40)
    try:
    # dataframe
        people_similarity = calculate_similarity(df_onehot_mentor, df_onehot_mentee)
        # dictionary
        global people_cluster
        people_cluster = create_cluster(people_similarity)
    except:
        return "Failed: error 04"

    if progbar_window!=None:
        update_progressbar(progbar_window, 60)

    # assignments start here
    # 7 | filter by mentee's first preference
    try:
        col ="If possible, what is your first priority in matching you with a mentor?"
        for mentee in df_mentee_raw.index:
            if df_mentee_raw[col].loc[mentee] == "Theatre":
                filter_by_pref(mentee, "Theatre")
            elif df_mentee_raw[col].loc[mentee] == "Role (Tech/Sales)":
                filter_by_pref(mentee, "Role")
            elif df_mentee_raw[col].loc[mentee] == "Segment":
                filter_by_pref(mentee, "Segment")
            elif df_mentee_raw[col].loc[mentee] == "Gender":
                filter_by_pref(mentee, "Gender")
    except:
        return "Failed: error 05"

    

    test_assignment("7")
    # 8 | filter by mentor's preferred theatre
    mentor_filter("Would you prefer a mentee in the same theatre?", "Your theatre?", method="yes")
    test_assignment("8")

    # 9 | filter by mentor's preferred role
    mentor_filter("Would you prefer a mentee of the same gender?", "Your gender?", method="yes")
    test_assignment("9")   

    # 10 | filter by mentor's preferred role
    mentor_filter("Would you prefer a mentee in a:", "Your role?", method="no preference")
    test_assignment("10")   

    if progbar_window!=None:
        update_progressbar(progbar_window, 80)

    # 11 | make rankings unique
    sorting = "unchecked"
    if unique == True:
        trys = 1
        sorting = sort_unique(people_cluster, method="limited")
        trys += 1
        while trys < 100:
            if sorting == "Success":
                print("Converged at try:", trys)
                trys = 100
                break
            elif trys < 10:
                sorting = sort_unique(people_cluster, method="limited")
                trys += 1
            else:
                choice_list = ["limited"]*1 + ["medium"]*1 + ["hard"]*1
                random_method = random.choice(choice_list)
                print(random_method)
                sorting = sort_unique(people_cluster, method=random_method)
                trys += 1 
            #print(warning_unique)
    test_assignment("11")   
    
    # 12 | filter if a mentor already has a mentee in mind
    filter_mentee_mind(df_mentee_raw, df_mentor_raw, people_cluster)
    test_assignment("12")   

    # crop by number of desired rankings, default = 10
    people_cluster_cropped = {}
    for mentee in people_cluster.keys():
        people_cluster_cropped[mentee] = people_cluster[mentee][0:ranks]
    try:
        # export results as Excel sheet
        export_excel(people_cluster_cropped, ranks)
    except:
        return "Failed: error 06. Could not save ranking file."

    if progbar_window!=None:
        update_progressbar(progbar_window, 100)

    if sorting == "Failed":
        return "Not unique"
    elif sorting == "unchecked":
        return "Success not unique"
    else:
        return "Success"
