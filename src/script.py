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
months = ['JAN']#, 'FEB', 'MAR', 'APR', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']


for i, month in enumerate(months):
    excelSheet = pd.read_excel(filename, sheet_name = month)
    addressCol = pd.DataFrame(excelSheet, columns= ['ADDRESS FOUND']) 
    addressCol.dropna(inplace=True)
    rowCount = addressCol.shape[0]
    #print(addressCol)
    print(addressCol.iloc[0][0])
    addStr = addressCol.iloc[0][0]
    geocode_result = gmaps.geocode(addStr)
    fullAddress = geocode_result[0]["formatted_address"]
    print(fullAddress)


    # for 
    #     geocode_result = gmaps.geocode('1600 Amphitheatre Parkway, Mountain View, CA')





# psedoCode

# for month in months 
#     Import data of that months sheet  as data frame
#     Pull out Address column as dataframe

#     Do the thing required in google maps API

#     Save Address column original excel sheet 