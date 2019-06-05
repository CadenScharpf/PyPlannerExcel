import os
import shutil
import datetime
from constants import *

month = 6
year = 2019

CONFIG = {
    "from_dt": None,
    "to_dt": None
} 

#TODO: create templates for months of different # days and modify individual sheets
#xlrd, xlutils and xlwt modules need to be installed.  
#Can be done via pip install <module>
#from xlrd import open_workbook
#from xlutils.copy import copy

#rb = open_workbook("names.xls")
#wb = copy(rb)

#s = wb.get_sheet(0)
#s.write(0,0,'A1')
#wb.save('names.xls')

def generate_files(month, year, **kwargs):
    src_dir = os.curdir
    src_file = os.path.join(src_dir, "template.xlsx")
    dst_dir = os.path.join(src_dir, "output")

    num_days = get_month_days(month, year)
    for day in range(num_days):
        print(dst_dir)
        date = datetime.date(year, month, day+1)
        filename = "output/"+date.isoformat()+".xlsx"
        shutil.copy2(src_file, dst_dir)
        dst_file = os.path.join(dst_dir, "template.xlsx")
        os.rename(dst_file, filename)

    

def get_month_days(month, year):
    if month == 2:
        if year%4 == 0:
            return 29 
        else:
            return MONTH_DAYS[str(month)]
    try:
        return MONTH_DAYS[str(month)]
    except:
        print("enter a valid month")

generate_files(month, year, **CONFIG)