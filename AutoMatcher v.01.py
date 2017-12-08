# -*- coding: utf-8 -*-
"""
Created on Wed Oct 25 14:23:11 2017

@author: achang
"""

from time import strftime
import pandas as pd
import sys
import os 
import xlrd 
import numpy as np
import re
import xlsxwriter
from fuzzywuzzy import fuzz 
#from mypackages import FunctionToolbox as tbox


#This function cleans the names 
def cleanName(name):
    
    try:
        name = name.strip().lower()
        name = name.replace("&"," and ").replace("professional corporation","")
        name = name.replace(" pc "," ").replace(" lp "," ").replace(" llc "," ")
        name = name.replace("incorporated","").replace(" inc."," ").replace(" inc "," ")
        name = re.sub('[^A-Za-z0-9\s]+', '', name)
        name = re.sub( '\s+', ' ', name)
    except AttributeError:
        name = ""
    return name

#This function cleans the addresses
def cleanAddress(address):
    try:
        address = address.strip().lower()
        address = address.replace("&"," and ").replace(".","")
        address = address.replace(" avenue "," ave ").replace(" boulevard "," blvd ").replace(" bypass "," byp ")
        address = address.replace(" circle "," cir ").replace(" drive "," dr ").replace(" expressway "," expy ")
        address = address.replace(" highway "," hwy ").replace(" parkway "," pkwy ").replace(" road "," rd ")
        address = address.replace(" street "," st ").replace(" turnpike "," tpke ")
        address = re.sub('[^A-Za-z0-9\s]+', '', address)
        address = re.sub( '\s+', ' ', address)
    except AttributeError:
        address = ""
    return address
    
#This function cleans the city
def cleanCity(city):
    try:
        city = city.strip().lower()
        city = city.replace("&"," and ").replace("saint ","st ")
        city = re.sub('[^A-Za-z0-9\s]+', '', city)
        city = re.sub( '\s+', ' ', city)
    except AttributeError:
        city = ""
    return city

#This function compares the names in the file
def compareName(df,IndustryType):
    df['Cleaned Location Name'] = df['Location Name'].apply(cleanName) 
    df['Cleaned Listing Name'] = df['Listing Name'].apply(cleanName)
    
    averagenamescore = []
    average=0
    #Agent Names matching
    if IndustryType=='5':
        inputName2=''
        busR=0
        busPR=0
        businessNames=[]
        inputName=raw_input("Enter Business Name:")
        businessNames.append(cleanName(inputName))
        while inputName2 !='0':
            inputName2=raw_input("Enter Alternative Business Name.  If no other alternatives, enter 0:\n")
            businessNames.append(cleanName(inputName2))
        
        for index, row in df.iterrows(): 
            busR=0
            busPR=0
            for bName in businessNames:
                busR=max(busR,fuzz.ratio(bName,row['Cleaned Listing Name']))
                busPR=max(busPR,fuzz.partial_ratio(bName,row['Cleaned Listing Name']))
            nsr = fuzz.ratio(row['Cleaned Location Name'], row['Cleaned Listing Name'])
            ntpr = fuzz.partial_ratio(row['Cleaned Location Name'], row['Cleaned Listing Name'])
            average = max(np.mean([busR,busPR]),np.mean([nsr,ntpr]))
            averagenamescore.append(average)
    else:       
        for index, row in df.iterrows(): 
            nsr = fuzz.ratio(row['Cleaned Location Name'], row['Cleaned Listing Name'])
            ntpr = fuzz.partial_ratio(row['Cleaned Location Name'], row['Cleaned Listing Name'])
            average = np.mean([nsr,ntpr])
            averagenamescore.append(average)
    df['Name Score'] = averagenamescore
    
#This function compares the countries in the file
def compareStateCountry(df):
    statematch = []
    for index, row in df.iterrows(): 
        # Drop different countries
        if (row['Input Location Country'] != row['Duplicate Location Country']):
            df.drop(index,inplace=True) 
        else:
            # Flag international locations
            if row['Input Location Country'] != "US":
                statematch.append("International")
            else:
                # Drop different states within US
                if (row['Location State'] != row['Listing State']):
                    df.drop(index,inplace=True)
                else:
                    statematch.append("US")
    df['US?'] = statematch

#This function compares the locationIds in the file
def compareId(df):
    IDmatch = []
    for index, row in df.iterrows(): 
        if (row['Input Location Id'] == row['Duplicate Location Id']):
            df.drop(index,inplace=True)
        else:
            IDmatch.append("Unique Pair")                
    df['Location ID Match'] = IDmatch

#This function removes matches that are not 1
def userMatch(df):
    for index, row in df.iterrows(): 
        if (row['Match \n1 = yes, 0 = no'] != 1):
            df.drop(index,inplace=True)

#This function compares the phones in the file                
def comparePhone(df):
    df['Phone Match'] = df.apply(lambda x: '0' if x['Location Phone'] == x['Listing Phone'] else '1', axis=1)

#This function compares the addresses in the file                
def compareAddress(df):
    df['Cleaned Input Address'] = df['Location Address'].apply(cleanAddress) 
    df['Cleaned Listing Address'] = df['Listing Address'].apply(cleanAddress)
    
    averageaddressscore = []
    for index, row in df.iterrows(): 
        
        #just international?
        inputStreetNumber=''
        inputStreetName=''
        ListingStreetNumber=''
        ListingStreetName=''
        if row['Cleaned Input Address'].split()[0].isdigit:    
            inputStreetNumber=row['Cleaned Input Address'].split()[0]
            inputStreetName=row['Cleaned Input Address'].split(None,1)[1]
        elif row['Cleaned Input Address'].split()[-1].isdigit:
            inputStreetNumber=row['Cleaned Input Address'].split()[-1]
            inputStreetName=row['Cleaned Input Address'].rsplit(None,1)[0]
        if (row['Cleaned Listing Address'] !='' and len(row['Cleaned Listing Address'].split())>1):
            if row['Cleaned Listing Address'].split()[0].isdigit:    
                ListingStreetNumber=row['Cleaned Listing Address'].split()[0]
                ListingStreetName=row['Cleaned Listing Address'].split(None,1)[1]
            elif row['Cleaned Listing Address'].split()[-1].isdigit:
                ListingStreetNumber=row['Cleaned Listing Address'].split()[-1]
                ListingStreetName=row['Cleaned Listing Address'].rsplit(None,1)[0]
        if (inputStreetNumber!=''and inputStreetName!='' and ListingStreetNumber!='' and  ListingStreetName!=''):
            numberR=fuzz.ratio(inputStreetNumber,ListingStreetNumber)
            nameR=fuzz.ratio(inputStreetName,ListingStreetName)
            averageaddressscore.append(np.mean([numberR,nameR]))
        else:
            asr = fuzz.ratio(row['Cleaned Input Address'], row['Cleaned Listing Address'])
            averageaddressscore.append(asr)
    df['Address Score'] = averageaddressscore            

#This function compares the cities in the file                
def compareCity(df):
    df['Cleaned Input City'] = df['Location City'].apply(cleanCity) 
    df['Cleaned Listing City'] = df['Listing City'].apply(cleanCity)
    averagecityscore = []
    for index, row in df.iterrows(): 
        csr = fuzz.ratio(row['Cleaned Input City'], row['Cleaned Listing City'])
        ctpr = fuzz.partial_ratio(row['Cleaned Input City'], row['Cleaned Listing City'])
        
        average = np.mean([csr,ctpr])
        averagecityscore.append(average)   
    df['City Score'] = averagecityscore

#This function compares the Country in the file                        
def compareCountry(df):
     df['Country Match'] = df.apply(lambda x: '0' if x['Input Location Country'] == x['Duplicate Location Country'] else '1', axis=1)

#This function compares the Zips in the file                   
def compareZip(df):
     df['Zip Match'] = df.apply(lambda x: '0' if x['Location Zip'] == x['Listing Zip'] else '1', axis=1)

#This function compares the data by calling on all the functions
def compareData(IndustryType,df):
    #compareId(df)
    comparePhone(df)
    #compareCountry(df)
    compareZip(df)
    compareName(df,IndustryType)
    compareAddress(df)
    #compareStateCountry(df)
    suggestedmatch(df, IndustryType)


#This function provides a suggested match based on certain name/address thresholds
def suggestedmatch(df, IndustryType):  
    robotmatch = []

    #Hotel Type
    if IndustryType == '2':
        for index, row in df.iterrows(): 
            #If phone matches
            if df['Phone Match']==1:
                if row['Address Score'] < 70:
                    robotmatch.append("No Match - Address")
                else:
                    #Need to add excluded words to leave out
                    robotmatch.append("Match Suggested")                                                
            else:
                if row['Address Score'] < 70:
                    robotmatch.append("No Match - Address")
                else:
                    #Need to add certain hotels and excluded words
                    if 60 < row['Name Score'] < 80:
                        robotmatch.append("Check")
                    else:
                        robotmatch.append("Match Suggested")                                                                    
        df['Robot Suggestion'] = robotmatch
        df['Match \n1 = yes, 0 = no'] = ""
    else:
        for index, row in df.iterrows(): 
            if row['Name Score'] <= 60:
                robotmatch.append("No Match - Name")
            else:
                if row['Address Score'] < 70:
                    robotmatch.append("No Match - Address")
                else:
                    if 60 < row['Name Score'] < 80:
                        robotmatch.append("Check")
                    else:
                        robotmatch.append("Match Suggested")                                                
        df['Robot Suggestion'] = robotmatch
        df['Match \n1 = yes, 0 = no'] = ""
    
def main():    
    
    xlsFile = raw_input("\nPlease input your file for matching. Make sure your file is saved as an XLSX"
                        "\n\nEnter File Path Here: ").replace('""','').lstrip("\"").rstrip("\"")
    directory = xlsFile[:xlsFile.rindex("\\")]
                        
    os.chdir(directory)
    
    IndustryType = raw_input("\nPlease input which industry you're matching Normal = 0, Auto = 1, Hotel = 2, Healthcare Doctor = 3, Healthcare Facility = 4, Agent = 5, International = 6\n")

    #xlsFile = r"C:\Users\achang\Downloads\test.xlsx"
    wb = xlrd.open_workbook(xlsFile, on_demand=True)
    sNames = wb.sheet_names()        
    wsTitle = "none"
    for name in sNames:
         wsTitle = name
                  
    row = 0 
    
    df = pd.ExcelFile(xlsFile).parse(wsTitle)


    #wb = xlrd.open_workbook(xlsFile, on_demand=True)
    #sNames = wb.sheet_names()        
    #for name in sNames:
    #     wsTitle = name
                  
    row = 0 
    # Read in data set
    lastcol = len(df.columns)
    
    for x in df['Link ID']:
        row += 1
    compareData(IndustryType,df)
            
    FilepathMatch =  os.path.expanduser("~\Documents\Python Scripts\AutoMatcher Output.xlsx")

    writer = pd.ExcelWriter(FilepathMatch, engine='xlsxwriter')
    df.to_excel(writer,sheet_name=wsTitle, index=False)
    workbook  = writer.book
    worksheet = writer.sheets[wsTitle]

    headerformat = workbook.add_format({
    'bold': True,
    'text_wrap': True})
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, headerformat)
        
    # Define column formatting
    formatBlue = workbook.add_format({'bg_color': '#77e8da', 'border' : 1, 'border_color': '#c0c0c0'}) #
    formatRed = workbook.add_format({'bg_color': '#f47676', 'border' : 1, 'border_color': '#c0c0c0'}) #
    formatYellow = workbook.add_format({'bg_color': '#f5cb70', 'border' : 1, 'border_color': '#c0c0c0'}) #
    formatPurp = workbook.add_format({'bg_color': '#d9b3ff', 'border' : 1, 'border_color': '#c0c0c0'}) #
    formatOrange = workbook.add_format({'bg_color': '#ffaa80', 'border' : 1, 'border_color': '#c0c0c0'}) #
   
    namescorecol = df.columns.get_loc("Name Score")
    addressscorecol = df.columns.get_loc("Address Score")
    cleannamecol = df.columns.get_loc("Cleaned Location Name")
    addresscol =  df.columns.get_loc("Cleaned Input Address")
    robotcol =  df.columns.get_loc("Robot Suggestion")
    lastgencol = df.columns.get_loc("Match \n1 = yes, 0 = no")
    Lat =  df.columns.get_loc("Listing Latitude")
    LastPostDate=  df.columns.get_loc("Listing Latitude")
    
    
    worksheet.set_row(0, 29.4)

    worksheet.set_column(Lat , LastPostDate, None, None, {'hidden': 1})
    worksheet.set_column(0, lastcol, None, None, {'hidden': 1})
    worksheet.set_column(lastcol+2, lastcol+3, None, None, {'hidden': 1})
    worksheet.set_column(cleannamecol, cleannamecol+1, 45)
    worksheet.set_column(namescorecol, namescorecol, 7.5)
    worksheet.set_column(addressscorecol, addressscorecol, 7.5)
    worksheet.set_column(addresscol, addresscol+1, 27)
    worksheet.set_column(robotcol, robotcol, 17.33)
    worksheet.set_column(robotcol+1, robotcol+1, 14.22)
    
    worksheet.autofilter(0,0,0,lastgencol)

    # Format Match columns
    worksheet.conditional_format(0, lastcol, 0, lastgencol, {'type':'text',
                                'criteria': 'containing',
                                'value':    "Name",
                                'format':   formatOrange})
    worksheet.conditional_format(0, lastcol, 0, lastgencol, {'type':'text',
                                'criteria': 'containing',
                                'value':    "Address",
                                'format':   formatPurp})
    # Format Score columns
    worksheet.conditional_format(0, lastcol, 0, lastgencol, {'type':'text',
                                'criteria': 'containing',
                                'value':    "Score",
                                'format':   headerformat })
    
    worksheet.conditional_format(1, lastcol, row, lastgencol, {'type':'text',
                                        'criteria': 'begins with',
                                        'value':    "Match",
                                        'format':   formatBlue})
    worksheet.conditional_format(1, lastcol, row, lastgencol, {'type':'text',
                                        'criteria': 'begins with',
                                        'value':    "Check",
                                        'format':   formatYellow})
    worksheet.conditional_format(1, lastcol, row, lastgencol, {'type':'text',
                                        'criteria': 'begins with',
                                        'value':    "Unique",
                                        'format':   formatBlue})

    #Not match coloring
    worksheet.conditional_format(1, lastcol, row, lastgencol, {'type':'text',
                                        'criteria': 'begins with',
                                        'value':    "No Match",
                                        'format':   formatRed})
    worksheet.conditional_format(1, lastcol, row, lastgencol, {'type':'text',
                                        'criteria': 'begins with',
                                        'value':    "Same Location",
                                        'format':   formatRed})
        
    try:
        writer.save()    
        print "\nDone! Results have been wrizzled to your Excel file. 1love <3"
        print "\nMatching Template here:"
        print FilepathMatch 
    except IOError:
        print "\nIOError: Make sure your Excel file is closed before re-running the script."          
            
main()