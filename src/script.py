from dotenv import load_dotenv
load_dotenv()
import os
KEY = os.getenv("api-key")

import pandas as pd
from tkinter.filedialog import askopenfilename
filename = askopenfilename()

months = ['JAN', 'FEB', 'MAR', 'APR', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']

excelSheet = [0,0,0,0,0,0,0,0,0,0,0,0]
ColAddress = [0,0,0,0,0,0,0,0,0,0,0,0]

for i, month in enumerate(months):
    excelSheet[i] = pd.read_excel (filename, sheet_name = month)
    ColAddress[i] = pd.DataFrame(excelSheet[i], columns= ['ADDRESS FOUND']) 

