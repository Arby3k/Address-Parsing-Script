from dotenv import load_dotenv
load_dotenv()
import os
KEY = os.getenv("api-key")

import pandas as pd
from tkinter.filedialog import askopenfilename
filename = askopenfilename()

months = ['JAN', 'FEB', 'MAR', 'APR', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']

excelSheet = []
ColAddress = []

for i, month in enumerate(months):
    excelSheet.append( pd.read_excel (filename, sheet_name = month) )
    ColAddress.append( pd.DataFrame(excelSheet[i], columns= ['ADDRESS FOUND']) )
