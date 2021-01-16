#GoogleMaps API key import
from dotenv import load_dotenv
load_dotenv()
import os
KEY = os.getenv("api-key")

#Pandas, and Excel enviroment import
import pandas as pd
from openpyxl import load_workbook

#File managment Import
from tkinter import filedialog
filename = filedialog.askopenfilename()
#outFilename = filedialog.asksaveasfilename(defaultextension='.xlsx')

#Main 
months = ['JAN', 'FEB', 'MAR', 'APR', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']


book = load_workbook(filename)
writer = pd.ExcelWriter(filename, engine = 'openpyxl')
writer.book = book

for month in months:
    excelSheet = pd.read_excel(writer, sheet_name = month)
    addressCol = pd.DataFrame(excelSheet, columns= ['ADDRESS FOUND']) 
    
    #Do stuff

    addressCol.to_excel(writer, sheet_name = month, columns= ['ADDRESS FOUND'], index = False )


writer.save()
writer.close()

# psedoCode

# for month in months 
#     Import data of that months sheet  as data frame
#     Pull out Address column as dataframe

#     Do the thing required in google maps API

#     Save Address column original excel sheet 