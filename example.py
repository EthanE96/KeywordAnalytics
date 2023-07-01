import eel
import shutil
import os
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import FORMULAE


#Declare Varible
myTimesheets = []

#Look in Folder
directory = r'C:\Users\Ethan\Test'
for filename in os.listdir(directory):
    if filename.endswith(".xlsx"):
        #Adds Excel Files to List
        myTimesheets.append(filename)

print(myTimesheets)
