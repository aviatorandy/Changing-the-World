# -*- coding: utf-8 -*-
"""
Created on Wed Oct 25 14:23:11 2017

@author: achang, plasser, storres
"""

from time import strftime
import pandas as pd
import os 
import xlrd 
import numpy as np
import re
from fuzzywuzzy import fuzz 
import MySQLdb
from math import radians, cos, sin, asin, sqrt, isnan
from Tkinter import *   
from ttk import *
import tkFileDialog
import Tkinter
import subprocess
import time
import datetime



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
            df.loc[index,'Cleaned Listing Name']=row['Cleaned Listing Name']\
            .replace(cleanName(row['Location City']),'')
            
    df['Cleaned Location Name'] = df['Location Name'].apply(cleanName) 
    df['Cleaned Listing Name'] = df['Listing Name'].apply(cleanName)
    
    averagenamescore = []
    average=0
    
    #Populates businessNames with Account Name and alt Name Policies
    inputName=''
    global businessNames
    businessNames=[]
    businessNames.append(cleanName(getBusName(bid)))
    altNames=getAltName(bid)
    for name in altNames:
        businessNames.append(cleanName(name))

    global namesComplete
 
    #calls Tkinter input window for more business names. Waits for it to complete
    app.namesWindow(businessNames)
    app.wait_window(app.nameW)
   
#start of comparisons, broken out by industry
    #Hotel
    if IndustryType=="2":
        OtherHotelMatch = []
        HotelBrands = ["test"]
        HotelBrands=["Grill", "bar", "starbucks", "electric", "wedding", "gym",\
                     "pool", "restaurant", "bistro"]     

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
    
    #Healthcare Professional matching
    if IndustryType=="3":
        ProviderScore = []
        for index, row in df.iterrows():
            ntpr = fuzz.token_set_ratio(row['Provider Name'], row['Cleaned Listing Name'])
            ProviderScore.append(ntpr)
        df['Name Score'] = ProviderScore
        return
        
    #Healthcare Facility matching
    if IndustryType=="4":
        return

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

    #Auto Name Matching
    if IndustryType=="6":
        return
        
    #Normal/International    
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
        df['Phone Match'] = df.apply(lambda x: True \
            if (x['Location Phone'] == x['Listing Phone'] and x['Location Phone'] != "")\
            else (True if (x['Location Local Phone'] == x['Listing Phone'] \
                           and x['Location Local Phone'] != "") else False), axis=1)
    except:
        df['Phone Match'] = 'x'    

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
        for index, row in df.iterrows(): 
                
                #Finds best match between normal ratio, sorted ratio
                asr = fuzz.ratio(row['Cleaned Input Address'], row['Cleaned Listing Address'])
                addressSortRatio=fuzz.token_sort_ratio(row['Cleaned Input Address'], row['Cleaned Listing Address'])
                apsr = fuzz.partial_token_sort_ratio(row['Cleaned Input Address'], row['Cleaned Listing Address'])

                averageaddressscore.append(max(asr,addressSortRatio,apsr))

    #All other industries    
    else:
        
        df['Cleaned Input Address'] = df['Location Address'].apply(cleanAddress) 
        df['Cleaned Listing Address'] = df['Listing Address'].apply(cleanAddress)
        averageaddressscore = []
        for index, row in df.iterrows(): 
               
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
     df['Country Match'] = df.apply(lambda x: '0' if x['Input Location Country'] != x['Duplicate Location Country']\
                                     else '1', axis=1)

#This function compares the Zips in the file                   
def compareZip(df):
     df['Zip Match'] = df.apply(lambda x: True if x['Location Zip'] == x['Listing Zip']\
                                 else False, axis=1)
     
#This function compares the NPIs in the file
def compareNPI(df):
     df['NPI Match'] = df.apply(lambda x: True if x['Location NPI'] == x['Listing NPI'] else False, axis=1)

#copied from old template. Calculates metric distance between geocodes     
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
    
#Potential Speed Optimizer of Passing in Rows instead of dataframes.    
#def compareEverything(df,IndustryType, bid):
#    averageaddressscore = []
#    phonescore = []
#    cleanedlocadd = []
#    cleanedlistingadd = []
#    print "working"
#    if IndustryType=='3':
#         NPIscore = []   
#    #    averagenamescore = []
#    for index, row in df.iterrows(): 

def compareData(df, IndustryType, bid):
    
    #compareId(df)
    print 'comparing phones'
    comparePhone(df)
    #compareCountry(df)
    print 'comparing zips'
    compareZip(df)
    if IndustryType=='3':
        print 'comparing NPIs'
        compareNPI(df)
    print 'comparing names'
    compareName(df,IndustryType,bid)
    print 'comparing addresses'
    compareAddress(df,IndustryType)
    #compareStateCountry(df)
    df['Distance (M)'] = df.apply(lambda row: calculateDistance(row), axis=1) 
    print 'suggesting matches'
    calculateTotalScore(df)
    suggestedmatch(df, IndustryType)

#This function provides a suggested match based on certain name/address thresholds
def suggestedmatch(df, IndustryType):  
    robotmatch = []
    
    #Creates new columns that are easier to read in code, and for users when outputed - Mostly boolean variables
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

        df['Match \n1 = yes, 0 = no'] = ""
   
    #Healthcare Professional
    if IndustryType == '3': 
        for index, row in df.iterrows(): 
            if row ['NPI Match'] :
                robotmatch.append("Match - NPI")
            elif row['Phone Match']=='1':
                if 66 < row['Name Score'] < 76 or row['Cleaned Listing Name']is None:
                    robotmatch.append("Check") 
                elif 76 <= row['Name Score']:
                    robotmatch.append("Match Suggested") 
                else: 
                    if row['No Name']=='URL for name':
                        robotmatch.append('Check - URL name')
                    else:
                        robotmatch.append("No Match - Name")                         
            elif row['Address Score'] < 70:
                if row['Distance (M)']<200:
                    if 66 < row['Name Score'] < 76 or row['Cleaned Listing Name']is None:
                        robotmatch.append("Check") 
                    elif 76 <= row['Name Score']:
                        robotmatch.append("Match Suggested - Geocode") 
                    else: 
                        if row['No Name']=='URL for name':
                            robotmatch.append('Check - URL name')
                        else:
                            robotmatch.append("No Match - Name")
                else:
                    robotmatch.append("No Match - Address")
            else:    
                if 66 < row['Name Score'] < 76 or row['Cleaned Listing Name']is None:
                    robotmatch.append("Check") 
                elif 76 <= row['Name Score']:
                    robotmatch.append("Match Suggested") 
                else: 
                    if row['No Name']=='URL for name':
                        robotmatch.append('Check - URL name')
                    else:
                        robotmatch.append("No Match - Name")                         
        
        df['Robot Suggestion'] = robotmatch
        df['Match \n1 = yes, 0 = no'] = ""
                        
                    
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
        df['Match \n1 = yes, 0 = no'] = ""        
    
 
     #All other industries
    else:
        #Applies Match rules based on new columns.
        df['Robot Suggestion'] = df.apply(lambda x: matchText if x['Name Match']==1 and \
        (x['Phone Match'] or x['Address Match'] or x['Geocode Match']) else (check if x['Name Match']==2 \
        else (noName if x['Name Match']==0 else (noAddress if not x['Address Match'] else 'uh oh'))) , axis=1)

    df['Name Match'] = df.apply(lambda x: 'Good' if x['Name Match'] == 1 else ('Check' if x['Name Match'] ==2 else 'Bad'), axis=1)
    df['Match \n1 = yes, 0 = no'] = ""

def calculateTotalScore(df):
#GET ALL THE SCORE THEN GIVE THEM WEIGHTING THEN CREATE A NEW TOTAL SCORE COLUMN    
    totalscore= []
    for index,row in df.iterrows():
        totalscore.append(row['Name Score']*.7 + row['Address Score']*.3)
    df['Total Score'] = totalscore
    return df
    
def ExternalID_De_Dupe(df):    
    #If Listing ID is matched to more than one location 
    df=df.sort_values(['Match','Listing ID','Total Score'], ascending=[True, True, True] )
    df = df.reset_index(drop=True)
    for index,row in df.iterrows():
      
        if index < df.shape[0]-1:
            if row['Match'] == 1 and df.iloc[index+1]['Match'] == 1:
                if row['Listing ID'] == df.iloc[index+1]['Listing ID']:
                    row['Match'] = 0
    return df

#Now not called at all
def main():
    getInput()    
    runProg()

#Now not called at all.    
def getInput():    
    #Gets all inputs
    IndustryType = raw_input("\nPlease input which industry you're matching Normal = 0, Auto = 1, Hotel = 2, \
                            Healthcare Doctor = 3, Healthcare Facility = 4, Agent = 5, International = 6\n")                
    inputChoice=raw_input("Do you want to pull data from SQL or give input file? 0=SQL, 1=File \n")
    
    
    if inputChoice=='0':
        bid = raw_input("Input Business ID: ")
        
        df=pd.DataFrame()
        df=sqlPull(bid)
    else:
        xlsFile = raw_input("\nPlease input your file for matching."
                            "\n\nEnter File Path Here: ").replace('""','').lstrip("\"").rstrip("\"")
        df,bid=readFile(xlsFile)
        runProg(df,IndustryType,bid)
        
#Reads CSV or XLSX files        
def readFile(xlsFile):
        
        print 'reading file'
        if xlsFile[-4:]=='xlsx':
            wb = xlrd.open_workbook(xlsFile, on_demand=True,encoding_override="utf-8")
            sNames = wb.sheet_names()        
            wsTitle = "none"
#            for name in sNames:
#                 wsTitle = name
      #Takes first sheet name, not last     
            wsTitle=sNames[0]
            df = pd.ExcelFile(xlsFile).parse(wsTitle)

        elif xlsFile[-3:]=='csv':
            df = pd.read_csv(xlsFile, encoding ='utf-8')
        else:
            raise Exception('What kind of file did you give me, bro?')
       
        #Finds Business ID
        bid=getBusIDfromLoc(df.loc[0,'Location ID'])
        
        return df,bid
    
    
 #New main runtime function - broken out as previous functions now handled in Tkinter   
def runProg(df,IndustryType,bid):    
    print 'runprog'
    row = 0 
    if IndustryType =='3':
    #Gets Providers First and Last name. Saves to column 'Provider Name'
        DoctorNameDF=getProviderName(df)
        df=df.merge(DoctorNameDF,on='Location ID', how='left')

    lastcol=df.shape[1]
    row=df.shape[0]

#Compares all, suggests matches    
    compareData(df,IndustryType,bid)
    
 #Completes Matching Question sheet   
    matchingNameQs=matchingQuestions(df,row)     
        
    FilepathMatch =  os.path.expanduser("~\Documents\Python Scripts\AutoMatcher Output.xlsx")

    df=df.sort_values(by='Robot Suggestion')
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
        app.AllDone("\nDone! Results have been wrizzled to your Excel file. 1love <3"+"\nMatching Template here:\n"+ FilepathMatch )
        print "\nDone! Results have been wrizzled to your Excel file. 1love <3"
        print "\nMatching Template here:"
        print FilepathMatch
   #Opens explorer window to path of output     
        subprocess.Popen(r'explorer /select,'+os.path.expanduser("~\Documents\Python Scripts\AutoMatcher Output.xlsx"))
    except IOError:
        app.AllDone("\nIOError: Make sure your Excel file is closed before re-running the script.")
        print "\nIOError: Make sure your Excel file is closed before re-running the script."          

#Pulls all location and listing match data
def sqlPull(bid,folderID,labelID):
    print 'pulling data'
    #Pull Location info and Listing IDs
    #IF LISTINGS:
    SQL_QueryMatches = open(os.path.expanduser("~/Documents/Changing-the-World/SQL Data Pull/1. Pull Matches.sql")).read()
    #IF SUPPRESSION:
    #SQL_QueryMatches = open(os.path.expanduser("~/Documents/Changing-the-World/SQL Data Pull/1. Pull Matches.sql")).read()    
    SQL_QueryMatches=SQL_QueryMatches.splitlines()
    
 #Subs out variables for Account ID numbers   
    for index, line in enumerate(SQL_QueryMatches):
        SQL_QueryMatches[index]=line.replace('@bizid', str(bid))
    
    if folderID !=0:
        for index, line in enumerate(SQL_QueryMatches):
            SQL_QueryMatches[index]=line.replace('--left join alpha.location_tree_nodes ltn on ltn.id=l.treeNode_id',\
                    'left join alpha.location_tree_nodes ltn on ltn.id=l.treeNode_id')\
                    .replace('--and l.treeNode_id=@folderid', 'and l.treeNode_id=@folderid').replace('@folderid', str(folderID))
    
    if labelID !=0:
        for index, line in enumerate(SQL_QueryMatches):
            SQL_QueryMatches[index]=line.replace('--left join alpha.location_labels ll on ll.location_id=l.id', \
            'left join alpha.location_labels ll on ll.location_id=l.id')
        for index, line in enumerate(SQL_QueryMatches):   
            SQL_QueryMatches[index]=line.replace('--and ll.label_id=@labelid', 'and ll.label_id=@labelid')
        for index, line in enumerate(SQL_QueryMatches):   
            SQL_QueryMatches[index]=line.replace('@labelid', str(labelID))
    
            
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
    #locationID=locationID.replace("\'","")
    SQL_BusIDQuery='SELECT business_id from alpha.locations where id='+str(locationID).replace("\'","")+';'
    Yext_OPS_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
    busID = pd.read_sql(SQL_BusIDQuery, con=Yext_OPS_DB)['business_id'][0]
    
    return busID    

   #Pulls First Name, and Last Name for Healthcare Professionals 
def getProviderName(df):
    print 'getting doctor names'
    SQL_DoctorNameQuery = open(os.path.expanduser("~/Documents/Changing-the-World/SQL Data Pull/Doctor Name.sql")).read()    
    locationIDs=df['Location ID']
            
    SQL_DoctorNameQuery=SQL_DoctorNameQuery.replace('@locationIDs', ','.join(map(str, locationIDs)))    
    Yext_OPS_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
    DoctorNamesIDs = pd.read_sql(SQL_DoctorNameQuery, con=Yext_OPS_DB)

    return DoctorNamesIDs        
    
 #Finds Listing names that could be good to match to based on prevalence    
      
def matchingQuestions(df,numLinks):
    df=df[df['Robot Suggestion'].isin(['No Match - Name','Check'])]
   
    pivot=pd.pivot_table(df,values=0,index='Listing Name',aggfunc='count')
    listingNames=pd.DataFrame(pivot)
    #print df[0]
    if not listingNames.empty:
        listingNames.columns=['Count']    
        listingNames=listingNames.sort_values(by=['Count'],ascending=False)
        listingNames=listingNames.reset_index() 
        listingNames['Listing Name']=listingNames['Listing Name'].apply(cleanName)
    
        listingNames= listingNames[listingNames['Count'] > 5]  
    
        return listingNames
    else:
        return pd.DataFrame([{'Listing Name' : 'None', 'Count' : '0'}])

    #GUI Tkinter section!

#from Tkinter import Frame
#from Tkinter import *

#Your app is a subclass of the Tkinter class Frame.


def writeUploadFile(uploadDF):
    filePath =  os.path.expanduser("~\Documents\Python Scripts\UploadLinkages.csv")

    print 'writing file'
    #writer = pd.ExcelWriter(filePath, engine='xlsxwriter')
    uploadDF.to_csv(filePath, encoding='utf-8',index=False)
    
    
    
class MatchingInput(Tkinter.Frame):

    def __init__(self, master):

        Tkinter.Frame.__init__(self, master, padx=10, pady=10)

        master.title("AutoMatcher Setup")

        master.minsize(width=500, height=300)
        
       # style = Style()
       # style.theme_use('classic')
        
        
#First screen - needs to explain what's going on, get input on where in process they are        
        self.IntroLabel=Label(master,text="Welcome to ze AutoMatcher! It be cool. Enjoy. \n Pick what you want to do, man.").grid(row=0,column=0,pady=10,padx=0)
        self.processChoice=IntVar()
        self.processChoice.set(-1)
        self.StartMatch=Radiobutton(master,text="Start AutoMatcher - pull data/enter file",variable=self.processChoice,value=0).grid(row=1,column=0,sticky=W)
        self.ChecksChoice=Radiobutton(master,text="Input Completed Manual Checks, create upload",variable=self.processChoice,value=1)\
                                                                .grid(row=1,column=1)
        
        self.Next=Button(master,text="Next",command=lambda: [self.initialSettingsWindow() if self.processChoice.get()==0 else \
                                                             (self.inputChecks() if self.processChoice.get()==1 else self.processChoice.set(-1))])\
                                                             .grid(row=2,column=0,sticky=W,pady=25,padx=10)
        self.Quit=Button(master,text="Quit",command= lambda: [root.destroy()]).grid(row=3,column=0, sticky=W,pady=10,padx=10)
        
#If the user has manually checked the matches, this will take those in, determine Match statuses, and produce upload document        
    def inputChecks(self):
        self.master.withdraw()
#Takes in completed matches file with checks filled out
        checkedFile=tkFileDialog.askopenfilename(initialdir = "/",title = "Select completed matching file with Check column filled out",\
                                                 defaultextension="*.xlsx;*.xls", filetypes=( ("Excel files", "*.xlsx;*.xls"), ("CSV", "*.csv"),('All files','*.*') ))        
        checkedDF,bid=readFile(checkedFile)
        
#Checks to see if all rows asking for a check have manual review        
        allChecksComplete=True
        for index, row in checkedDF.iterrows(): 
            if  ('Check' in row['Robot Suggestion'] and (isnan(row['Match \n1 = yes, 0 = no']))):
                allChecksComplete=False
                
#Exits if manual review incomplete                
        if not allChecksComplete:
            self.errorBox=Toplevel()
            self.errorMsg=Label(self.errorBox,text="Please complete all checks first. Bye.").pack()
            self.okButton=Button(self.errorBox,text="OK",command= lambda:[root.destroy()]).pack()
#If complete, determines matches, creates upload        
        else:
            checkedDF['Match']=checkedDF.apply(lambda x: 1 if 'Match Suggested' in x['Robot Suggestion'] else 0,axis=1)
            checkedDF['Match']=checkedDF.apply(lambda x: 1 if x['Match \n1 = yes, 0 = no']==1 else (0 if x['Match \n1 = yes, 0 = no']==0 else x['Match']),axis=1)

            #EXTERNAL ID DEDPUE
            #print checkedDF
            checkedDF['override']=checkedDF.apply(lambda x: 'Match' if x['Match']==1 else 'AntiMatch',axis=1)
            checkedDF['PL Status']=checkedDF.apply(lambda x: 'Sync' if x['override']=='Match' else 'NoPowerListing',axis=1)
               
           
            #checkedDF=calculateTotalScore(checkedDF)
            checkedDF=ExternalID_De_Dupe(checkedDF)
            uploadDF=checkedDF[['Publisher ID','Location ID','Listing ID','override','PL Status']]

            
            writeUploadFile(uploadDF)
            print t0
            t1=time.time()
            print t1
            print t1-t0
       #remove this once we do something here     
            root.destroy()
                    
 
            #Gets user input for how to set up matcher           
    def initialSettingsWindow(self):
        
        self.master.withdraw()
        self.settingWindow=Toplevel()
        self.IndustryType=IntVar()
        self.IndustryType.set(-1)
        self.IndustryLabel=Label(self.settingWindow,text="Select Industry Type").grid(row=0,column=0)

        self.Normal=Radiobutton(self.settingWindow, text="Normal", variable=self.IndustryType,value=0).grid(row=1,column=7,sticky=W)
        self.Auto=Radiobutton(self.settingWindow, text="Auto", variable=self.IndustryType,value=1).grid(row=1,column=1,sticky=W)
        self.Hotel=Radiobutton(self.settingWindow, text="Hotel", variable=self.IndustryType,value=2).grid(row=1,column=2,sticky=W)
        self.Doctor=Radiobutton(self.settingWindow, text="Healthcare Doctor", variable=self.IndustryType,value=3).grid(row=1,column=3,sticky=W)
        self.Facility=Radiobutton(self.settingWindow, text="Healthcare Facility", variable=self.IndustryType,value=4).grid(row=1,column=4,sticky=W)
        self.Agent=Radiobutton(self.settingWindow, text="Agent", variable=self.IndustryType,value=5).grid(row=1,column=5,sticky=W)
        self.International=Radiobutton(self.settingWindow, text="International", variable=self.IndustryType,value=6).grid(row=1,column=6,sticky=W)

#        self.quitButton.pack()
        self.dataInput=IntVar()
        self.dataInput.set(0)
        self.inputType=Label(self.settingWindow,text="Select data input type:")

        self.SQL=Radiobutton(self.settingWindow, text="Pull Data from SQL", variable=self.dataInput,value=2)
        self.file=Radiobutton(self.settingWindow, text="Input File", variable=self.dataInput,value=1)
        
        
        
        self.nextButton = Button(self.settingWindow, text="Next", command=self.detailsWindow)
        
        self.inputType.grid(row=2,column=0)
        self.SQL.grid(row=3,column=1, sticky=W)
        self.file.grid(row=4,column=1, sticky=W)
        self.nextButton.grid(row=5,column=0,sticky=W,pady=25)

 #Gets File path or business, folder, and label ids to pull from SQL       
    def detailsWindow(self):
        self.settingWindow.destroy()
        #File pull
        if self.dataInput.get()==1:
            fname = tkFileDialog.askopenfilename(initialdir = "/",title = "Select file",filetypes=(("CSV", "*.csv"), ("Excel files", "*.xlsx;*.xls") ))
            df,bid=readFile(fname)
            
            runProg(df,str(self.IndustryType.get()),bid)
        #SQL
        elif self.dataInput.get()==2:
            self.detailsW=Tkinter.Toplevel(self)
            vcmd = self.master.register(self.validate)
            self.InputExplanation=Label(self.detailsW,text="Please enter account info \
                                        for data to pull. Business ID is required. You can leave blank \
                                        or put 0 for Folder ID and Label ID if not needed").grid(row=5,column=0)
            self.busIDLabel=Label(self.detailsW,text="Enter Business ID").grid(row=6,column=0,sticky=W)
            self.folderIDLabel=Label(self.detailsW,text="Enter Folder ID").grid(row=7,column=0,sticky=W)
            self.labelIDLabel=Label(self.detailsW,text="Enter Label ID").grid(row=8,column=0,sticky=W)
            
            self.bizID=StringVar()
            self.folderID=StringVar()
            self.labelID=StringVar()
            self.busIDentry=Entry(self.detailsW,validate="key", validatecommand=(vcmd, '%P'),textvariable=self.bizID).grid(row=6,column=1,sticky=W)
            self.folderIDentry=Entry(self.detailsW,validate="key", validatecommand=(vcmd, '%P'),textvariable=self.folderID).grid(row=7,column=1,sticky=W)
            self.labelIDentry=Entry(self.detailsW,validate="key", validatecommand=(vcmd, '%P'),textvariable=self.labelID).grid(row=8,column=1,sticky=W)
        
            self.pullButton = Button(self.detailsW, text="Pull Data", command=self.pullSQLRun).grid(row=9,column=1,sticky=W)
            
#Runs SQl Pull            
    def pullSQLRun(self):
                
        if self.bizID.get():
            
            self.detailsW.destroy()
            if  self.folderID.get():
                folderID=self.folderID.get()
            else: 
                folderID=0
            if self.labelID.get():
                labelID=self.labelID.get()
            else:
                labelID=0
            df=sqlPull(self.bizID.get(),folderID,labelID)
            runProg(df,self.IndustryType.get(),self.bizID.get())
    
#button function to ammend businessNames list            
    def AddMore(self):
        global businessNames
        for i in self.NewName.get().split(","):
            businessNames.append(cleanName(i)) 
        global namesComplete
        namesComplete = 1
       # businessNames.append(cleanName(NewName))
        #AddNameBox.delete(0, END)
        self.nameW.destroy()
       # self.namesWindow(businessNames)
 
  #Prints out current business names, asks for additional names  
    def namesWindow(self,busNames):
        global businessNames
        global namesComplete
        self.complete=BooleanVar()
        self.complete=False
        self.nameW=Toplevel()
        self.nameW.minsize(width=300, height=200)
        self.NewName=StringVar()
        self.CurrentNameLabel=Label(self.nameW,text="Current Business Names:\n"+ ", ".join([str(i) for i in businessNames]) ).grid(row=10,column=0)
        self.AddMoreLabel=Label(self.nameW,text="Enter comma separated list of additional business names:").grid(row=11,column=0)
        self.AddName=Entry(self.nameW,textvariable=self.NewName).grid(row=12,column=0)
        self.AddMore=Button(self.nameW,text="Add names",command=self.AddMore).grid(row=13,column=0, sticky=W)
        self.Done=Button(self.nameW,text="No more names needed",command= lambda: [self.nameW.destroy()]).grid(row=13,column=1, sticky=W)  
         
#Checks if input is a number
    def validate(self, new_text):
        if not new_text: # the field is being cleared
            self.entered_number = 0
            return True

        try:
            self.entered_number = int(new_text)
            return True
        except ValueError:
            return False

  #Final screen to close program.          
    def AllDone(self,msg):
        global t0
        t1=time.time()
        print "start: "+datetime.datetime.fromtimestamp(t0).strftime('%Y-%m-%d %H:%M:%S')
        print "end: "+ datetime.datetime.fromtimestamp(t1).strftime('%Y-%m-%d %H:%M:%S')
        print str(t1-t0)+" seconds"
        self.completeWindow=Toplevel()  
        self.completeMsg=Label(self.completeWindow,text="time to complete: "+str(t1-t0)+"  "+msg).grid(row=1,column=1)
        self.DoneButton=Button(self.completeWindow,text="Exit", command= lambda:[root.destroy()]).grid(row=2,column=1)

#Defining this globally helps being able to call it within the Tkinter class
global df
df=pd.DataFrame
global checkedDF
global root
checkedDF=pd.DataFrame     
global t0
t0=time.time()
#starts Tkinter
root = Tkinter.Tk()
app = MatchingInput(root)
app.mainloop()