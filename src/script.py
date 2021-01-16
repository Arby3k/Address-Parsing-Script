import pandas as pd
from tkinter.filedialog import askopenfilename
filename = askopenfilename()


x=0
months = ['JAN', 'FEB', 'MAR', 'APR', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
excelSheet = [0,0,0,0,0,0,0,0,0,0,0,0]
ColAddress = [0,0,0,0,0,0,0,0,0,0,0,0]

for month in months:
    excelSheet[x] = pd.read_excel (filename, sheet_name = month)
    ColAddress[x] = pd.DataFrame(excelSheet[x], columns= ['ADDRESS FOUND'])
    x = x+1

