#GoogleMaps API key import
from dotenv import load_dotenv
load_dotenv()
import os
KEY = os.getenv("api-key")

#GoogleMaps Services for Python Import
import googlemaps
gmaps = googlemaps.Client(key=KEY)

#Pandas, and Excel enviroment import
import pandas as pd
import openpyxl as op


#File managment Import
from tkinter import filedialog
filename = filedialog.askopenfilename()
#outFilename = filedialog.asksaveasfilename(defaultextension='.xlsx')

#Main 
months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']


for n, month in enumerate(months):
    excelSheet = pd.read_excel(filename, sheet_name = month)
    addressCol = pd.DataFrame(excelSheet, columns= ['ADDRESS FOUND']) 
    addressCol.dropna(inplace=True)
    rowCount = addressCol.shape[0]

    wb = op.load_workbook(filename)
    ws = wb[month]

    for i in range(0,rowCount):
        addStr = addressCol.iloc[i][0]
        #print(addStr)
        if addStr != 'UNK':
            geocodeResult = gmaps.geocode(addStr)
            #print(type(geocodeResult))
            if geocodeResult:
                addressCol.iloc[i][0] = geocodeResult[0]["formatted_address"]
            else:
                addressCol.iloc[i][0] = 'MANUAL CHECK ' + addressCol.iloc[i][0]
        
        #print(addressCol.iloc[i][0])
    print(month)
    for index, row in addressCol.iterrows():
        cell = 'J%d'  % (index + 2)
        ws[cell] = row[0]
    
    wb.save(filename)


# psedoCode

# for month in months 
#     Import data of that months sheet  as data frame
#     Pull out Address column as dataframe

#     Do the thing required in google maps API

#     Save Address column original excel sheet 