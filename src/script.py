from dotenv import load_dotenv
load_dotenv()
import os
KEY = os.getenv("api-key")

import pandas as pd
from tkinter.filedialog import askopenfilename
filename = askopenfilename()

months = ['JAN', 'FEB', 'MAR', 'APR', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']

for month in months
    excelSheet = pd.read_excel (filename, sheet_name = month)
    adressCol = pd.DataFrame(excelSheet, columns= ['ADDRESS FOUND']) 


    addressCol.to_excel(filename, sheet_name = month, columns= ['ADDRESS FOUND'] )


# psedoCode

# for month in months 
#     Import data of that months sheet  as data frame
#     Pull out Address column as dataframe

#     Do the thing required in google maps API

#     Save Address column original excel sheet 