# -*- coding: utf-8 -*-
"""
Created on Wed Oct 25 14:23:11 2017

@author: achang, plasser, storres
"""

from time import strftime
import pandas as pd
import timeit
import sys
import os 
import xlrd 
import numpy as np
import re
import csv
import xlsxwriter
from fuzzywuzzy import fuzz 
import MySQLdb
import xlwings
from math import radians, cos, sin, asin, sqrt
from Tkinter import *   
import tkFileDialog
import Tkinter

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
def compareName(df,IndustryType,bid):
    df['Cleaned Location Name'] = df['Location Name'].apply(cleanName) 
    df['Cleaned Listing Name'] = df['Listing Name'].apply(cleanName)
    
    df['No Name']=''
    for index,row in df.iterrows():
        #If name is blank, fills in last part of URL
        if row['Listing Name']==None and row['Listing URL']!=None:
            df.loc[index,'Cleaned Listing Name']=cleanName(row['Listing URL'].split('/')[-1])
            df.loc[index,'No Name']='URL for name'
        elif row['Listing Name']==None:
            df.loc[index,'No Name']='No Name'
    for index,row in df.iterrows():    
        #Removes City name if in Listing name
        if cleanName(row['Location City']) in row['Cleaned Listing Name']:
            df.loc[index,'Cleaned Listing Name']=row['Cleaned Listing Name'].replace(cleanName(row['Location City']),'')
            
    df['Cleaned Location Name'] = df['Location Name'].apply(cleanName) 
    df['Cleaned Listing Name'] = df['Listing Name'].apply(cleanName)
    
    averagenamescore = []
    average=0
    
    #Populates businessNames with Account Name and alt Name Policies
    inputName=''
    businessNames=[]
    businessNames.append(cleanName(getBusName(bid)))
    altNames=getAltName(bid)
    for name in altNames:
        businessNames.append(cleanName(name))

    print '\nCurrent Business Names:'
    print businessNames
    #Get additional good match names for business
    while inputName !='0':
        inputName=raw_input("Enter additional business name, or 0 to continue:\n")
        businessNames.append(cleanName(inputName))
    
    #Hotel
    if IndustryType=="2":
        OtherHotelMatch = []
        HotelBrands = ["test"]
        HotelBrands=["Grill", "bar", "starbucks", "electric", "wedding", "gym", "pool", "restaurant", "bistro"]     

        #HotelBrands=["AC hotels", "aloft", "America's Best", "americas best value", "ascend", "autograph", "baymont", "best western", "cambria", "canadas best value", "candlewood", "clarion", "comfort inn", "comfort suites", "Country Hearth", "courtyard", "crowne plaza", "curio", "days inn", "doubletree", "econo lodge", "econolodge", "edition", "Element", "embassy", "even", "fairfield inn", "four points", "garden inn", "Gaylord", "hampton inn", "hilton", "holiday inn", "homewood", "howard johnson", "hyatt", "indigo", "intercontinental", "Jameson", "JW", "la quinta", "Le Meridien", "Le Méridien", "Lexington", "luxury collection", "mainstay", "marriott", "microtel", "motel 6", "palace inn", "premier inn", "quality inn", "quality suites", "ramada", "red roof", "renaissance", "residence", "ritz", "rodeway", "sheraton", "Signature Inn", "sleep inn", "springhill", "st regis", "st. regis", "starwood", "staybridge", "studio 6", "super 8", "towneplace", "Value Hotel", "Value Inn", "W hotel", "westin", "wingate", "wyndham"]        
        for index, row in df.iterrows(): 
            businessRatio=0
            businessPartialRatio=0
            OtherHotel = 0
            for brands in HotelBrands:
                businessRatio=max(businessRatio,fuzz.ratio(brands, row['Cleaned Listing Name']))
                businessPartialRatio=max(businessPartialRatio,fuzz.partial_ratio(brands, row['Cleaned Listing Name']))            
            #Check listing name against location name
            nsr = fuzz.ratio(row['Cleaned Location Name'], row['Cleaned Listing Name'])
            ntpr = fuzz.partial_ratio(row['Cleaned Location Name'], row['Cleaned Listing Name'])
            #returns Max of Business Match or Location Name Match

            #Pass a boolean of if it's part of the brand match
            if np.mean([businessRatio,businessPartialRatio]) > np.mean([nsr,ntpr]):
                OtherHotel = 1
            else:
                OtherHotel = 0
            OtherHotelMatch.append(OtherHotel)
            average = max(np.mean([businessRatio,businessPartialRatio]),np.mean([nsr,ntpr]))
            averagenamescore.append(average)
        df['Other Hotel Match'] = OtherHotelMatch
        print len(OtherHotelMatch)
        df['Name Score'] = averagenamescore
        print len(averagenamescore)
        return
    #df['Name Score'] = averagenamescore
    #Agent Names matching
    if IndustryType=="5":
        for index, row in df.iterrows(): 
            businessRatio=0
            businessPartialRatio=0
            #Check listing name against business names
            for bName in businessNames:
                businessRatio=max(businessRatio,fuzz.ratio(bName,row['Cleaned Listing Name']))
                businessPartialRatio=max(businessPartialRatio,fuzz.partial_ratio(bName,row['Cleaned Listing Name']))
            #Check listing name against location name
            nsr = fuzz.ratio(row['Cleaned Location Name'], row['Cleaned Listing Name'])
            ntpr = fuzz.partial_ratio(row['Cleaned Location Name'], row['Cleaned Listing Name'])
            #returns Max of Business Match or Location Name Match
            average = max(np.mean([businessRatio,businessPartialRatio]),np.mean([nsr,ntpr]))
            averagenamescore.append(average)
        df['Name Score'] = averagenamescore
        return
    else:       
        for index, row in df.iterrows(): 
            businessRatio=0
            businessPartialRatio=0
            #Check listing name against business names
            for bName in businessNames:
                businessRatio=max(businessRatio,fuzz.ratio(bName,row['Cleaned Listing Name']))
                businessPartialRatio=max(businessPartialRatio,fuzz.partial_ratio(bName,row['Cleaned Listing Name']))
            #Check listing name against location name
            nsr = fuzz.ratio(row['Cleaned Location Name'], row['Cleaned Listing Name'])
            ntpr = fuzz.partial_ratio(row['Cleaned Location Name'], row['Cleaned Listing Name'])
            #returns Max of Business Match or Location Name Match
            average = max(np.mean([businessRatio,businessPartialRatio]),np.mean([nsr,ntpr]))
            averagenamescore.append(average)
        df['Name Score'] = averagenamescore
        return   
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
    
    
    try:
        df['Phone Match'] = df.apply(lambda x: True if (x['Location Phone'] == x['Listing Phone'] and x['Location Phone']!="") else (True if (x['Location Local Phone'] == x['Listing Phone']and x['Location Local Phone']!="") else False), axis=1)
    except:
        df['Phone Match'] = 'x'
    
#
#    for index, row in df.iterrows(): 
#        try:
#            if row['Location Phone'] == ['Listing Phone']:
#                row['Phone Match']='1'
#            elif row['Local Phone']==['Listing Phone']:
#                row['Phone Match']='1'
#            else:
#                row ['Phone Match']='0'
#        except:
#            row['Phone Match'] = '0'  

#This function compares the addresses in the file                
def compareAddress(df,IndustryType):
    #International 
    if IndustryType=='6':
        #Combine Address 1, Address 2
        df['Cleaned Input Address'] = df['Location Address'].apply(cleanAddress)+' '+df['Location Address 2'].apply(cleanAddress) 
        df['Cleaned Listing Address'] = df['Listing Address'].apply(cleanAddress)+' '+df['Listing Address 2'].apply(cleanAddress)
        #removes extra space where necessary
        df['Cleaned Input Address']=df['Cleaned Input Address'].apply(cleanAddress)
        df['Cleaned Listing Address'] =df['Cleaned Listing Address'].apply(cleanAddress)
        averageaddressscore = []
#        asrSeries=[]
#        sortSeries=[]
#        aprSeries=[]
        for index, row in df.iterrows(): 
                #Finds best match between normal ratio, sorted ratio
                asr = fuzz.ratio(row['Cleaned Input Address'], row['Cleaned Listing Address'])
                addressSortRatio=fuzz.token_sort_ratio(row['Cleaned Input Address'], row['Cleaned Listing Address'])
                apsr = fuzz.partial_token_sort_ratio(row['Cleaned Input Address'], row['Cleaned Listing Address'])

                averageaddressscore.append(max(asr,addressSortRatio,apsr))
#                asrSeries.append(asr)
#                sortSeries.append(addressSortRatio)
#                aprSeries.append(apr)
    else:
        df['Cleaned Input Address'] = df['Location Address'].apply(cleanAddress) 
        df['Cleaned Listing Address'] = df['Listing Address'].apply(cleanAddress)
        averageaddressscore = []
        for index, row in df.iterrows(): 
                asr = fuzz.ratio(row['Cleaned Input Address'], row['Cleaned Listing Address'])
                averageaddressscore.append(asr)
    df['Address Score'] = averageaddressscore
    #df['Address Score'] = asrSeries
    #df['Sort'] = sortSeries
    #df['APR'] = aprSeries

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
     df['Country Match'] = df.apply(lambda x: '0' if x['Input Location Country'] != x['Duplicate Location Country'] else '1', axis=1)

#This function compares the Zips in the file                   
def compareZip(df):
     df['Zip Match'] = df.apply(lambda x: True if x['Location Zip'] == x['Listing Zip'] else False, axis=1)
     
def calculateDistance(row):
    try: locationLat = float(row['Location Latitude'])
    except: locationLat = 0.0
    try: locationLong = float(row['Location Longitude'])
    except: locationLong = 0.0
    try: listingLat = float(row['Listing Latitude'])
    except: listingLat = 90.0
    try: listingLong = float(row['Listing Longitude'])
    except: listingLong = 90.0


    if locationLat == 0.0 or listingLat == 0.0: return "n/a"

    # convert decimal degrees to radians
    locationLat, locationLong, listingLat, listingLong = map(radians, [locationLat, locationLong, listingLat, listingLong])

    # haversine formula
    dlon = listingLong - locationLong
    dlat = listingLat - locationLat
    a = sin(dlat/2)**2 + cos(locationLat) * cos(listingLat) * sin(dlon/2)**2
    c = 2 * asin(sqrt(a))
    r = 6371 # Radius of earth in kilometers. Use 3956 for miles
    return c * r * 1000

    

#This function compares the data by calling on all the functions
def compareData(df,IndustryType,bid):
    #compareId(df)
    print 'comparing phones'
    comparePhone(df)
    #compareCountry(df)
    print 'comparing zips'
    compareZip(df)
    print 'comparing names'
    compareName(df,IndustryType,bid)
    print 'comparing addresses'
    compareAddress(df,IndustryType)
    #compareStateCountry(df)
    df['Distance (M)'] = df.apply(lambda row: calculateDistance(row), axis=1) 
    print 'suggesting matches'
    suggestedmatch(df, IndustryType)

#This function provides a suggested match based on certain name/address thresholds
def suggestedmatch(df, IndustryType):  
    robotmatch = []
    
    df['Address Match'] = df.apply(lambda x: True if x['Address Score'] >= 70 else False, axis=1)
    df['Name Match'] = df.apply(lambda x: 1 if x['Name Score'] >= 70 else (2 if 70 > x['Name Score'] > 60 else 0), axis=1)
    df['Geocode Match'] = df.apply(lambda x: True if x['Distance (M)']<=200 else False, axis=1)
    
    matchText='Match Suggested'
    noName='No Match - Name'
    noAddress='No Match - Address'
    check='Check Name'
    
    
    
    

    #Hotel Type
    if IndustryType == '2':
        for index, row in df.iterrows(): 
            #If hotel matches another brand better
            if row['Other Hotel Match'] == "1":                
                if row['Phone Match']:
                    if row['Address Score'] < 70:
                        robotmatch.append("No Match - Address")
                    else:
                        if 60 < row['Name Score'] < 80:
                            robotmatch.append("Check") 
                        elif 80 <= row['Name Score']:
                            robotmatch.append("Match Suggested") 
                        else: 
                            robotmatch.append("No Match - Name")                         
                else:
                    if row['Address Score'] < 70:
                        robotmatch.append("No Match - Address")
                    else:
                        if 60 < row['Name Score']:
                            robotmatch.append("No Match - Name")                         
                        else:
                            robotmatch.append("Check")                         
            else:              
                if row['Phone Match']:
                    if row['Address Score'] < 70:
                        robotmatch.append("No Match - Address")
                    else:
                        if 60 < row['Name Score'] < 80:
                            robotmatch.append("Check") 
                        elif 80 <= row['Name Score']:
                            robotmatch.append("Match Suggested") 
                        else: 
                            robotmatch.append("No Match - Name")                         
                else:
                    if row['Address Score'] < 70:
                        robotmatch.append("No Match - Address")
                    else:
                        if 60 < row['Name Score'] < 80:
                            robotmatch.append("Check")                         
                        elif 80 <= row['Name Score']:
                            robotmatch.append("Match Suggested")                                                                    
                        else:
                            robotmatch.append("No Match - Name")  
        df['Robot Suggestion'] = robotmatch

    #International
    elif IndustryType=='6':
        for index, row in df.iterrows(): 
            if row['Name Score'] <= 60 and row['No Name']!='URL for name':
                robotmatch.append("No Match - Name")
            else:
                if row['Phone Match']:
                    if 60 < row['Name Score'] < 80 or row['Cleaned Listing Name']is None:
                        robotmatch.append("Check")
                    elif row['Name Score']<60:
                        robotmatch.append('Check - URL Name')
                    else:
                        robotmatch.append("Match Suggested") 
                elif row['Country']=='GB':
                    #if GB zip matches, then address match
                    if row['Address Score'] < 70 and not row['Zip Match']:
                        if row['Distance (M)']<200:
                            if 60 < row['Name Score'] < 80 or row['Cleaned Listing Name']is None:
                                robotmatch.append("Check")
                            elif row['Name Score']<60:
                                robotmatch.append('Check - URL Name')
                            else:
                                robotmatch.append("Match Suggested - Geocode")
                        else:
                            robotmatch.append("No Match - Address")
                    elif row['Zip Match']:
                        if 60 < row['Name Score'] < 80 or row['Cleaned Listing Name']is None:
                            robotmatch.append("Check")
                        elif row['Name Score']<60:
                            robotmatch.append('Check - URL Name')
                        else:
                            robotmatch.append("Match Suggested")
                    
                    else:
                        if 60 < row['Name Score'] < 80 or row['Cleaned Listing Name']is None:
                            robotmatch.append("Check")
                        elif row['Name Score']<60:
                            robotmatch.append('Check - URL Name')
                        else:
                            robotmatch.append("Match Suggested")
                else:
                    if row['Address Score'] < 70:
                        if row['Distance (M)']<200:
                            if 60 < row['Name Score'] < 80 or row['Cleaned Listing Name']is None:
                                robotmatch.append("Check")
                            elif row['Name Score']<60:
                                robotmatch.append('Check - URL Name')

                            else:
                                robotmatch.append("Match Suggested - Geocode")
                        else:
                            robotmatch.append("No Match - Address")
                    else:
                        if 60 < row['Name Score'] < 80 or row['Cleaned Listing Name']is None:
                            robotmatch.append("Check")
                        elif row['Name Score']<60:
                            robotmatch.append('Check - URL Name')
                        else:
                            robotmatch.append("Match Suggested")                                                
        df['Robot Suggestion'] = robotmatch    
    #All other industries
    else:
        
         
        
        df['Robot Suggestion'] = df.apply(lambda x: matchText if x['Name Match']==1 and (x['Phone Match'] or x['Address Match'] or x['Geocode Match']) else (check if x['Name Match']==2 else (noName if x['Name Match']==0 else (noAddress if not x['Address Match'] else 'uh oh'))) , axis=1)

        
        #
#        for index, row in df.iterrows(): 
#            if row['Phone Match']=='1':
#                if 60 < row['Name Score'] < 80 or row['Cleaned Listing Name']is None:
#                    robotmatch.append("Check") 
#                elif 80 <= row['Name Score']:
#                    robotmatch.append("Match Suggested") 
#                else: 
#                    if row['No Name']=='URL for name':
#                        robotmatch.append('Check - URL name')
#                    else:
#                        robotmatch.append("No Match - Name")                         
#            elif row['Address Score'] < 70:
#                if row['Distance (M)']<200:
#                    if 60 < row['Name Score'] < 80 or row['Cleaned Listing Name']is None:
#                        robotmatch.append("Check") 
#                    elif 80 <= row['Name Score']:
#                        robotmatch.append("Match Suggested - Geocode") 
#                    else: 
#                        if row['No Name']=='URL for name':
#                            robotmatch.append('Check - URL name')
#                        else:
#                            robotmatch.append("No Match - Name")
#                else:
#                    robotmatch.append("No Match - Address")
#            else:    
#                if 60 < row['Name Score'] < 80 or row['Cleaned Listing Name']is None:
#                    robotmatch.append("Check") 
#                elif 80 <= row['Name Score']:
#                    robotmatch.append("Match Suggested") 
#                else: 
#                    if row['No Name']=='URL for name':
#                        robotmatch.append('Check - URL name')
#                    else:
#                        robotmatch.append("No Match - Name")                         
#        
#        df['Robot Suggestion'] = robotmatch
#        df['Match \n1 = yes, 0 = no'] = ""


    df['Match \n1 = yes, 0 = no'] = ""
def main():
    getInput()    
    runProg()
    
def getInput():    
    #Gets all inputs
    IndustryType = raw_input("\nPlease input which industry you're matching Normal = 0, Auto = 1, Hotel = 2, Healthcare Doctor = 3, Healthcare Facility = 4, Agent = 5, International = 6\n")                
    inputChoice=raw_input("Do you want to pull data from SQL or give input file? 0=SQL, 1=File \n")
    
    
    if inputChoice=='0':
        bid = raw_input("Input Business ID: ")
        
        df=pd.DataFrame()
        df=sqlPull(bid)
    else:
        xlsFile = raw_input("\nPlease input your file for matching."
                            "\n\nEnter File Path Here: ").replace('""','').lstrip("\"").rstrip("\"")
        directory = xlsFile[:xlsFile.rindex("\\")]
                            
        os.chdir(directory)
        
        
        #xlsFile = r"C:\Users\achang\Downloads\test.xlsx"
        print 'reading file'
        if xlsFile[-4:]=='xlsx':
            wb = xlrd.open_workbook(xlsFile, on_demand=True)
            sNames = wb.sheet_names()        
            wsTitle = "none"
            for name in sNames:
                 wsTitle = name
            df = pd.ExcelFile(xlsFile).parse(wsTitle)

        elif xlsFile[-3:]=='csv':
            df = pd.read_csv(xlsFile)
        else:
            raise Exception('What kind of file did you give me, bro?')
        #wb = xlrd.open_workbook(xlsFile, on_demand=True)
        #f = open(xlsFile, 'rb')
        #reader = csv.reader(f)
        #sNames = wb.sheet_names()        
        #wsTitle = "none"
        #for name in sNames:
        #     wsTitle = name
        row = 0 
        
        #Finds Business ID
        bid=getBusIDfromLoc(df.loc[0,'Location ID'])
        
        
        #timeit.timeit('"-".join(str(n) for n in range(100))', number=10000)
        #df = pd.ExcelFile(xlsFile).parse(wsTitle)
    
    #wb = xlrd.open_workbook(xlsFile, on_demand=True)
    #sNames = wb.sheet_names()        
    #for name in sNames:
    #     wsTitle = name
    
    
def runProg(df,IndustryType,bid):    
    print 'runprog'
    print df
    row = 0 
    # Read in data set
   
    #Gets Providers First and Last name. Saves to column 'Provider Name'
    DoctorNameDF=getProviderName(df)
    df=df.merge(DoctorNameDF,on='Location ID', how='left')


    lastcol=df.shape[1]
    row=df.shape[0]

    compareData(df,IndustryType,bid)
    
    matchingNameQs=matchingQuestions(df,row)
        
    FilepathMatch =  os.path.expanduser("~\Documents\Python Scripts\AutoMatcher Output.xlsx")

    print 'writing file'
    writer = pd.ExcelWriter(FilepathMatch, engine='xlsxwriter')
    df.to_excel(writer,sheet_name="Result", index=False)
    matchingNameQs.to_excel(writer,sheet_name="Matching Questions", index=False)
    workbook  = writer.book
    worksheet = writer.sheets["Result"]

    print 'formatting file'
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
    worksheet.set_column(0, lastcol-1, None, None, {'hidden': 1})
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
    worksheet.conditional_format(1, lastgencol-1, row, lastgencol, {'type':'text',
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
     
    print 'saving'
    try:
        writer.save()    
        print "\nDone! Results have been wrizzled to your Excel file. 1love <3"
        print "\nMatching Template here:"
        print FilepathMatch 
    except IOError:
        print "\nIOError: Make sure your Excel file is closed before re-running the script."          

#Pulls all location and listing match data
def sqlPull(bid):
    print 'pulling data'
    #Pull Location info and Listing IDs
    SQL_QueryMatches = open(os.path.expanduser("~/Documents/Changing-the-World/SQL Data Pull/1. Pull Matches.sql")).read()
    SQL_QueryMatches=SQL_QueryMatches.splitlines()
    
    
    for index, line in enumerate(SQL_QueryMatches):
        SQL_QueryMatches[index]=line.replace('@bizid', str(bid))
        
    SQL_QueryMatches = [x for x in SQL_QueryMatches if x.startswith("--") is False]
    FinalSQL_QueryMatches = []
    for x in SQL_QueryMatches:
        if '--' in x:
            FinalSQL_QueryMatches.append(x[0:x.index('-')])
        else:
            FinalSQL_QueryMatches.append(x)
        
        #Convert Query into string from list
    FinalSQL_QueryMatches = ' '.join(FinalSQL_QueryMatches)
    
    Yext_Mat_DB = MySQLdb.connect(host="127.0.0.1", port=5009, db="alpha")
    SQL_DataMatches = pd.read_sql(FinalSQL_QueryMatches, con=Yext_Mat_DB)
    
    #Clean IDs
    ListingIDs=SQL_DataMatches['Listing ID']
    ListingIDs=ListingIDs.map(lambda x: x.lstrip('\''))
    
    #Gets Listing Data from Listing IDs
    SQL_QueryListings = open(os.path.expanduser("~/Documents/Changing-the-World/SQL Data Pull/2. Pull Listings Data.sql")).read()
    SQL_QueryListings=SQL_QueryListings.splitlines()
    
    for index, line in enumerate(SQL_QueryListings):
        SQL_QueryListings[index]=line.replace('@ListingIDs', ','.join(map(str, ListingIDs)) )
                
                
    SQL_QueryListings = [x for x in SQL_QueryListings if x.startswith("--") is False]
    FinalQueryListings = []
    for x in SQL_QueryListings:
        if '--' in x:
            FinalQueryListings.append(x[0:x.index('-')])
        else:
            FinalQueryListings.append(x)
        
        #Convert Query into string from list
    FinalQueryListings = ' '.join(FinalQueryListings)
    
    Yext_Prod_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
    SQL_DataListings = pd.read_sql(FinalQueryListings, con=Yext_Prod_DB)    
    
    
    #Combine results
    df=SQL_DataListings.merge(SQL_DataMatches,on='Listing ID',how='outer')
    return df

#Gets Alt Name policies at business level
def getAltName(bid):    
    SQL_AltNameQuery = open(os.path.expanduser("~/Documents/Changing-the-World/SQL Data Pull/Alt Name policies.sql")).read()

    SQL_AltNameQuery=SQL_AltNameQuery.replace('@bizid', str(bid))
    
    Yext_SMS_DB = MySQLdb.connect(host="127.0.0.1", port=5007, db="alpha")
    AltNames = pd.read_sql(SQL_AltNameQuery, con=Yext_SMS_DB)['policyString']

    BusNames=[]
    for name in AltNames:
        BusNames.append(name)
    
    return BusNames      
    
#Gets Business Account name    
def getBusName(bid):    

    SQL_BusQuery='Select name from alpha.businesses where id='+str(bid)+';'
    
    Yext_OPS_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
    busName = pd.read_sql(SQL_BusQuery, con=Yext_OPS_DB)['name'][0]
    
    return busName  
    
#Gets Business ID from location ID
def getBusIDfromLoc(locationID):
    SQL_BusIDQuery='SELECT business_id from alpha.locations where id='+str(locationID)+';'
    Yext_OPS_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
    busID = pd.read_sql(SQL_BusIDQuery, con=Yext_OPS_DB)['business_id'][0]
    
    return busID    

   #Pulls First Name, and Last Name for Healthcare Professionals 
def getProviderName(df):
    print 'getting doctor names'
    print df
    SQL_DoctorNameQuery = open(os.path.expanduser("~/Documents/Changing-the-World/SQL Data Pull/Doctor Name.sql")).read()    
    locationIDs=df['Location ID']
            
    SQL_DoctorNameQuery=SQL_DoctorNameQuery.replace('@locationIDs', ','.join(map(str, locationIDs)))    
    Yext_OPS_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
    DoctorNamesIDs = pd.read_sql(SQL_DoctorNameQuery, con=Yext_OPS_DB)

    return DoctorNamesIDs    
    
    
 #Finds Listing names that could be good to match to based on prevalence    
      
def matchingQuestions(df,numLinks):
    df=df[df['Robot Suggestion'].isin(['No Match - Name','Check'])]
    pivot=pd.pivot_table(df,values='Link ID',index='Listing Name',aggfunc='count')
    listingNames=pd.DataFrame(pivot)
    listingNames.columns=['Count']    
    listingNames=listingNames.sort_values(by=['Count'],ascending=False)
    listingNames=listingNames.reset_index() 
    listingNames['Listing Name']=listingNames['Listing Name'].apply(cleanName)

    listingNames= listingNames[listingNames['Count'] > 5]  

    return listingNames




from Tkinter import Frame
from Tkinter import *

#Your app is a subclass of the Tkinter class Frame.
class MatchingInput(Tkinter.Frame):

    #Constructor for our app.  The master parameter represents the root object
    #that will hold out application.  It is basically a window, with our application
    #existing in a frame inside that window.
 def __init__(self, master):
        #Call the constructor of the Frame superclass.  The pad options create padding
        #between the edge of the Frame and the Widgits inside.
        Tkinter.Frame.__init__(self, master, padx=10, pady=10)
        #Give our window a title by calling the .title() method on the master object.
        master.title("AutoMatcher Setup")
        #Set a minimum size through the master object, just to make our UI a little
        #nicer to look at.
        master.minsize(width=1000, height=1000)
        #.pack() is a necessary method to get our app ready to be displayed.  Packing
        #objects puts them in a simple column.  For a more complex way to arrange your
        #widgets, see .grid():
        #http://effbot.org/tkinterbook/grid.htm
        #self.pack()

        #Now let's make some buttons using the Tkinter class Button.  The "text" parameter
        #indicates the text to be displayed in the button and the "command" parameter
        #specifies a procedure to execute if the button is clicked.  In this app, we will
        #have buttons that increase and decrease a variable.
#        self.upButton = Tkinter.Button(self, text="Up", command=self.increment)
#        self.upButton.pack() #Individual objects also must be packed to appear.
#        self.downButton = Tkinter.Button(self, text="Down", command=self.decrement)
#        self.downButton.pack()
#
#        #The variable that we'll be incrementing and decrementing.
#        self.value = 0
#        #When you want to integrate a variable with your Widgets (eg buttons, labels, etc),
#        #you make it a special type of Tkinter variable.  In this case, a StringVar.  There
#        #is also IntVar, DoubleVar, and BooleanVar.  These are essentially mutable versions
#        #of primitive types.  If we assign a normal string varaible to be the text of a button,
#        #then change that string variable, the text of the button would be unchanged.  If we
#        #instead use a StringVar, the button text will update automatically.  You'll see!
#        self.value_str = Tkinter.StringVar()
#        self.value_str.set("0") #You set all Tkinter variable objects with the .set() method.
#
#        #A Label to display our value.  Labels are like buttons except with no click effect.
#        #Note that we use the textvariable parameter instead of text so that the text on
#        #this label will automatically update with our StringVar.
#        self.valueLabel = Tkinter.Label(self, textvariable=self.value_str)
#        self.valueLabel.pack()
#
#        #Lastly, a quit button, which will call the .create_quit_window() method defined below,
        #which displays a new window asking whether the user want's to quit.
        #self.quitButton = Tkinter.Button(master, text="Quit", command=self.quit)
       
        self.IndustryType=IntVar()
        self.IndustryType.set(-1)
        self.IndustryLabel=Label(master,text="Select Industry Type").grid(row=0,column=0)

        self.Normal=Radiobutton(master, text="Normal", variable=self.IndustryType,value=0).grid(row=1,column=0)
        self.Normal=Radiobutton(master, text="Auto", variable=self.IndustryType,value=1).grid(row=1,column=1)
        self.Normal=Radiobutton(master, text="Hotel", variable=self.IndustryType,value=2).grid(row=1,column=2)
        self.Normal=Radiobutton(master, text="Healthcare Doctor", variable=self.IndustryType,value=3).grid(row=1,column=3)
        self.Normal=Radiobutton(master, text="Healthcare Facility", variable=self.IndustryType,value=4).grid(row=1,column=4)
        self.Normal=Radiobutton(master, text="Agent", variable=self.IndustryType,value=5).grid(row=1,column=5)
        self.Normal=Radiobutton(master, text="International", variable=self.IndustryType,value=6).grid(row=1,column=6)
        
        
        
        
        
        
#        self.quitButton.pack()
        self.dataInput=IntVar()
        self.dataInput.set(0)
        self.inputType=Label(master,text="Select data input type:")

        self.SQL=Radiobutton(master, text="Pull Data from SQL", variable=self.dataInput,value=2)
        self.file=Radiobutton(master, text="Input File", variable=self.dataInput,value=1)
        
        
        
        self.nextButton = Button(master, text="Next", command=self.detailsWindow)
        
        self.inputType.grid(row=2,column=0)
        self.SQL.grid(row=3,column=0, sticky=W)
        self.file.grid(row=4,column=0, sticky=W)
        self.nextButton.grid(row=5,column=0,sticky=W)

        global master1
        master1=master
 def detailsWindow(self):
        if self.dataInput.get()==1:
            fname = tkFileDialog.askopenfilename(initialdir = "/",title = "Select file",filetypes=(("CSV", "*.csv"), ("Excel files", "*.xlsx;*.xls") ))
        elif self.dataInput.get()==2:
            vcmd = master1.register(self.validate)
            self.busIDLabel=Label(master1,text="Enter Business ID").grid(row=6,column=0,sticky=W)
            self.folderIDLabel=Label(master1,text="Enter Folder ID").grid(row=7,column=0,sticky=W)
            self.labelIDLabel=Label(master1,text="Enter Label ID").grid(row=8,column=0,sticky=W)
            
            self.bizID=StringVar()
            busID=Entry(master1,validate="key", validatecommand=(vcmd, '%P'),textvariable=self.bizID).grid(row=6,column=1,sticky=W)
            self.folderID=Entry(master1,validate="key", validatecommand=(vcmd, '%P')).grid(row=7,column=1,sticky=W)
            self.labelID=Entry(master1,validate="key", validatecommand=(vcmd, '%P')).grid(row=8,column=1,sticky=W)
        
            self.pullButton = Button(master1, text="Pull Data", command=self.pullSQLRun).grid(row=9,column=1,sticky=W)
            
            
 def pullSQLRun(self):
        print self.bizID.get()
        df=sqlPull(self.bizID.get())
        runProg(df,self.IndustryType.get(),self.bizID.get())
            
def namesWindow(self,busNames):
        self.CurrentNameLabel=Label(master1,text="Current Business Names: "+ busNames).grid(row=10,column=0)
        self.AddName=Entry(master1,text="Enter additional Business Name:").grid(row=11,column=0)
        self.AddMore=Button(master1,text="Add more",command=lambda:[busNames.append(cleanName(AddName)),self.AddName.clear_text]).grid(row=12,column=1)
        self.Done=Button(master1,text="Done!",command=self.Done).grid(row=12,column=2)

def AddMore(self,busNames,AddName):
        busNames.append(cleanName(AddName))
        self.AddName.clear_text
 
def Done(self)        
def validate(self, new_text):
        if not new_text: # the field is being cleared
            self.entered_number = 0
            return True

        try:
            self.entered_number = int(new_text)
            return True
        except ValueError:
            return False

global df
df=pd.DataFrame        
#We make a Tk object to serve as the root of our interface.  We're making 
#something to use as an argument for the master parameter, to be the "parent"
#of our frame.  We can use it to do things like set the size.  Other than
##that, you don't have to worry too much about this step.
root = Tkinter.Tk()
###Initialize our app object and run the Frame method .mainloop() to begin!
###You should see a small window with up and down buttons, a label displaying
###a number, and a quit button when you run this program.
app = MatchingInput(root)
app.mainloop()


#main()
