# Compare two spreadsheets to create a third spreadsheet, having a list of commands
# Only add competencies from req spreadsheet that do not already exist
# requires pandas, xlsxwriter

from datetime import datetime

import pandas as pd
import sys

p_record_file = ""
c5_req_file = ""
g_list = []
g_comp_key = []
g_comp_type = []

class c5_item:
    def __init__(this):
        this.posn = ""
        this.div = ""
        this.branch = ""
        this.section = ""
        this.title = ""
        this.pay_grade = ""
        this.osms = ""
        this.da_comp = []
        this.da_code = []
        this.priority = []
        this.training_priority = []

    def add_comp(this, c):
        this.da_comp.append(c)

    def add_da(this, c):
        this.da_code.append(c)

    def add_priority(this, c):
        this.priority.append(c)

    def add_training_priority(this, c):
        this.training_priority.append(c)

    def __str__(this):
        s = ""
        if (len(this.posn) < 1):
            return s
        s = this.posn+"\n\n"
        s += this.div + ", "+this.branch+", "+this.section+", "+this.title+", "+this.pay_grade+", "+this.osms+"\n\n"
        for i in range(0, len(this.da_comp)):
            s += this.da_code[i]+"\t"+this.da_comp[i]+"\t\t"+this.priority[i]+"\t"+this.training_priority[i]+"\n"
        return s

    # spreadsheet will contain some variation of these for positions that are not valid
    def valid_title(title):
        title = str(title)
        if len(title) < 1:
            return False
        if "DO NOT" in title:
            return False
        if "NOT IN" in title:
            return False
        if "NO DATA" in title:
            return False
        return True

    def valid_comp(comp):
        comp = str(comp)
        if len(comp) < 1:
            return False
        if "NO COMP" in comp:
            return False
        return True


class p_record_item:
    def __init__(this):
        this.posn = ""
        this.organization_level = ""
        this.department = ""
        this.source = ""
        this.emp_salary_plan = ""
        this.title = ""
        this.pay_grade = ""
        this.rank = ""
        this.name = ""
        this.emp_id = ""
        this.comp_key = []
        this.tmt_key = []
        this.comp = []
        this.comp_type = []
        this.cert_date = []
        this.required = []
        this.qualified = []

    def __str__(this):
        s = ""
        if len(this.posn) < 1:
            return s
        s = this.posn+"\n\n"
        s += this.name+", "+this.organization_level+", "+this.department+", "+this.source+"\n"
        s += this.emp_salary_plan+", "+this.title+", "+this.pay_grade+", "+this.emp_id+"\n"
        s +="\n"
        for i in range(0, len(this.comp_key)):
            s+=this.comp_key[i]+"\t\t\t"+this.comp[i]+"\t\t\t"+this.required[i]+"\t\t"+this.qualified[i]+"\n";
        return s

    def add_comp_key(this, c):
        this.comp_key.append(c)

    def add_tmt_key(this, c):
        this.tmt_key.append(c)

    def add_comp(this, c):
        this.comp.append(c)

    def add_comp_type(this, c):
        this.comp_type.append(c)

    def add_cert_date(this, c):
        this.cert_date.append(c)

    def add_required(this, c):
        this.required.append(c)

    def add_qualified(this, c):
        this.qualified.append(c)


def err(disp):
	now = datetime.now() 
	# dd/mm/YY H:M:S
	dt_string = now.strftime("%Y/%m/%d %H:%M:%S")
	print(dt_string+": "+disp);
	return 1


def get_rec(val1, val2):
    if val1 == "nan":
        return val2
    return val1


'''
create and insert item at the correct position of the list

list1: list to search through
posn: field to search for
item: new class to insert, if posn is not found
return position of posn in list
'''
def list_find_index(list1, posn, item):
    i = 0
    for i in range(0, len(list1)):
        if list1[i].posn < posn:
            continue
        if list1[i].posn == posn:
            return i
        break
    list1.insert(i, item)
    return i 


def lookup_comp_type(da_comp):
    i = 0
    for i in range(0, len(g_comp_key)):
        if find_in_str(g_comp_key[i], da_comp):
            return g_comp_type[i]
    return "Unknown"


def add_comp_lookup(comp_key, comp_type):
    i = 0
    for i in range(0, len(g_comp_key)):
        if g_comp_key[i] < comp_key:
            continue
        if g_comp_key[i] == comp_key:
            return i
        break
    g_comp_key.insert(i, comp_key)
    g_comp_type.insert(i, comp_type)
    return


def verify_not_nan(str1):
    if str1 != str1:
        str1 = ""
    elif len(str1) < 1:
        str1 = ""
    elif str1 == "nan":
        str1 = ""
    return str1


def read_c5_req():
    try:
        df = pd.read_excel(c5_req_file, skiprows=[0], dtype=str)
    except:
    	exit(err("file "+c5_req_file+" could not be opened"))

    ml = []
    last_div = ""
    last_branch = ""
    last_section = ""
    last_title = ""
    last_pay_grade = ""
    last_osms = ""
    last_posn = ""
    priority = ""
    training_priority = ""
    # Some of the columns contain rows which are multiline, so go through carefully
    for index, row in df.iterrows():
        posn = row['POSN#']
        if pd.isna(posn):
            posn = ""
        da_code = row['Direct Access Code']
        if pd.isna(da_code):
            da_code = ""

        if posn == "":
            if da_code == "":
                continue
        else: # Deal with blank lines
            last_posn = posn
            last_div = verify_not_nan(row['DIV'])
            last_section = verify_not_nan(row['SECTION'])
            last_title = verify_not_nan(row['WORKROLE / POSN TITLE'])
            last_pay_grade = verify_not_nan(row['PAY GRADE'])
            last_branch = verify_not_nan(row['BRANCH'])
            last_osms = verify_not_nan(row['OSMS'])

        if not c5_item.valid_title(last_title):
            continue
        comp = row['Compentency']
        if not c5_item.valid_comp(comp):
            continue
        priority = row['Priority']
        training_priority = row['Training Priority']

        q = list_find_index(ml, last_posn, c5_item())
        ml[q].posn = last_posn
        ml[q].div = last_div
        ml[q].branch = last_branch
        ml[q].section = last_section
        ml[q].title = last_title
        ml[q].pay_grade = last_pay_grade
        ml[q].osms = last_osms
        ml[q].add_comp(str(comp))
        ml[q].add_da(str(da_code))
        ml[q].add_priority(str(priority))
        ml[q].add_training_priority(str(training_priority))
    return ml


def read_p_record():
    try:
        df = pd.read_excel(p_record_file, dtype=str)
    except:
        exit(err("file "+p_record_file+" could not be opened"))

    ml = []
    organization_level = ""
    department = ""
    source = ""
    emp_salary_plan = ""
    posn = ""
    title = ""
    pay_grade = ""
    rank = ""
    name = ""
    emp_id = ""
    comp_key = ""
    tmt_key = ""
    comp = ""
    comp_type = ""
    cert_date = ""
    required = ""
    qualified = ""
    for i, row in df.iterrows():
        posn = verify_not_nan(row['Position Number'])
        if len(posn) < 1:
            continue

        q = list_find_index(ml, posn, p_record_item())
        organization_level = verify_not_nan(row['Organization Level '])
        department = verify_not_nan(row['Department'])
        source = verify_not_nan(row['Source'])
        emp_salary_plan = verify_not_nan(row['Employee Salary Admin Plan Info'])
        title = verify_not_nan(row['Title'])
        pay_grade = verify_not_nan(row['Position Grade'])
        rank = verify_not_nan(row['Position Rank'])
        name = verify_not_nan(row['Employee'])
        emp_id = verify_not_nan(row['Employee ID'])
        comp_key = verify_not_nan(row['Comp Key'])
        tmt_key = verify_not_nan(row['TMT Comp Key'])
        comp = verify_not_nan(row['Competency'])
        comp_type = verify_not_nan(row['Type'])
        cert_date = verify_not_nan(row['Certified Date'])
        required = verify_not_nan(row['Position Required'])
        if row['DA Qualified'] == "Yes" or row['TMT Certified'] == "Yes":
            qualified = "qualified"
        else:
            qualified = "not qualified"

        ml[q].posn = posn
        ml[q].organization_level = organization_level
        ml[q].department = department
        ml[q].source = source
        ml[q].emp_salary_plan = emp_salary_plan
        ml[q].title = title
        ml[q].pay_grade = pay_grade
        ml[q].rank = rank
        ml[q].name = name
        ml[q].emp_id = emp_id
        ml[q].add_comp_key(comp_key)
        ml[q].add_tmt_key(tmt_key)
        ml[q].add_comp(comp)
        ml[q].add_comp_type(comp_type)
        ml[q].add_cert_date(cert_date)
        ml[q].add_required(required)
        ml[q].add_qualified(qualified)
        add_comp_lookup(comp_key, comp_type)
    return ml


# Try different variations to see if the strings match enough
def find_in_str(str1, str2):
    if str1 in str2:
        return True
    if str2 in str1:
        return True
    return False


# Given the two lists of same position numbers, find if the comp already exists
# if not, add it to the request list
def match_comps(c5_req_item, p_record_item):
    for i in range(0, len(c5_req_item.da_code)):
        for q in range(0, len(p_record_item.comp_key)):
            found = False
            if find_in_str(c5_req_item.da_code[i], p_record_item.comp_key[q]):
                found = True
                break
        if not found:
            add_to_list(c5_req_item, i, p_record_item, q)
    return


# if the position number matches, and the comp is not in the list, add to the request list
def match_lists(p_record_list, c5_req_list):
    for i in range(0, len(c5_req_list)):
        for q in range(0, len(p_record_list)):
            if c5_req_list[i].posn == p_record_list[q].posn:
                match_comps(c5_req_list[i], p_record_list[q])
                break
    return



'''
Unit Name
Department ID
Position Number
Rate
Status
Competency Description
Competency Code
Type
Action
Importance
Comments
CG-1B1 Comments
'''
# Create row to enter into request list
def add_to_list(c5_req_item, c5_index, p_record_item, p_record_index):
    i = c5_index
    q = p_record_index
    sl = []
    p1 = ""
    p2 = ""
    idx = p_record_item.department.find('(')

    p1 = p_record_item.department[0:idx-1]
    p2 = p_record_item.department[idx+1:p_record_item.department.find(')')]

    sl.append(p1)
    sl.append(p2)
    sl.append(c5_req_item.posn)
    sl.append(p_record_item.rank)
    sl.append(p_record_item.source)
    sl.append(c5_req_item.da_comp[i])
    sl.append(c5_req_item.da_code[i])
    sl.append(lookup_comp_type(c5_req_item.da_code[i]))
    sl.append("add")
    sl.append(c5_req_item.priority[i])
    sl.append("")
    sl.append("")
    g_list.append(sl)
    return



# main

if len(sys.argv) < 3:
    err("Usage: "+sys.argv[0]+" <p_record_file.xlsx>   <req_file.xlsx>")
    exit(1)

p_record_file = sys.argv[1]
c5_req_file = sys.argv[2]

err("start")

p_record_list = []
c5_req_list = []

p_record_list = read_p_record()
c5_req_list = read_c5_req()

match_lists(p_record_list, c5_req_list)

df = pd.DataFrame(g_list, columns=['Unit Name', 'Department ID', 'Position Number', 'Rate', 'Status', 'Competency Description', 'Competency Code', 'Type', 'Action', 'Importance', 'Comments', 'CG-1B1 Comments'])

print("Creating request for adding "+str(len(df))+" items")

now = datetime.now()
dt_string = now.strftime("%Y_%m_%d")
with pd.ExcelWriter("out_req_"+dt_string+".xlsx") as writer:
	df.to_excel(writer, sheet_name=f'Sheet1', index=False, header=True, startrow=4)
	for column in df:
		column_width = max(df[column].astype(str).map(len).max(), len(column))
		new_width = column_width+(column_width*.1)
		col_idx = df.columns.get_loc(column)
		writer.sheets['Sheet1'].set_column(col_idx, col_idx, new_width)

err("finish")
print("")
