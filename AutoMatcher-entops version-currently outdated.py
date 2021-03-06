# -*- coding: utf-8 -*-
"""
Created on Wed Oct 25 14:23:11 2017

@author: achang, plasser, storres
"""

from time import strftime
import pandas as pd
import os 
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
from datetime import date
import sys
import math
import win32com.client
import warnings
import itertools




warnings.simplefilter("ignore", UserWarning)
warnings.filterwarnings('ignore', category=MySQLdb.Warning)
#import win32com.client


#This function cleans the names 
def cleanName(name):    
    try:
      #  name=name.encode('utf-8')           #this is erroring out for some reason. Maybe if blank? nicodeDecodeError: 'charmap' codec can't decode byte 0x81 in position 30: character maps to <undefined>
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
        address = address.replace(" street "," st ").replace(" turnpike "," tpke ")\
    .replace(" road"," rd")\
    .replace(" street"," st")\
    .replace(" place"," pl")\
    .replace(" drive"," dr")\
    .replace(" i-h "," i")\
    .replace(" ih "," i")\
    .replace(" interstate "," i")\
    .replace(" boulevard"," blvd")\
    .replace(" parkway"," pkwy")\
    .replace(" lane"," ln")\
    .replace(" turnpike"," tpke")\
    .replace(" highway"," hwy")\
    .replace(" route "," rt ")\
    .replace(" rte "," rt ")\
    .replace(" avenue"," ave")\
    .replace(" freeway"," fwy")\
    .replace(" court"," ct")\
    .replace(" expressway"," expy")\
    .replace(" mount "," mt ")\
    .replace(" trail"," trl")\
    .replace("lyndon b johnson","lbj")\
    .replace(" center"," ctr")\
    .replace(" centre"," ctr")\
    .replace(" circle"," cir")\
    .replace(" bypass"," byp")\
    .replace(" pike"," pk")\
    .replace(" saint"," st")\
    .replace(" terrace"," ter")\
    .replace(" point"," pt")\
    .replace(" station"," sta")\
    .replace(" causeway"," cswy")\
    .replace(" crossing"," xing")\
    .replace(" gateway"," gtwy")\
    .replace(" creek"," ck")\
    .replace(" village"," vlg")\
    .replace(" first"," 1st")\
    .replace(" second"," 2nd")\
    .replace(" third"," 3rd")\
    .replace(" fourth"," 4th")\
    .replace(" fifth"," 5th")\
    .replace(" sixth"," 6th")\
    .replace(" seventh"," 7th")\
    .replace(" eighth"," 8th")\
    .replace(" ninth"," 9th")\
    .replace(" tenth"," 10th")\
    .replace("one","1")\
    .replace("two","2")\
    .replace("three","3")\
    .replace("four","4")\
    .replace("five","5")\
    .replace("six","6")\
    .replace("seven","7")\
    .replace("eight","8")\
    .replace("nine","9")\
    .replace("state road","sr")\
    .replace("county road","cr")\
    .replace(" hiway "," hwy ")\
    .replace("farm to market","fm")\
    .replace(" northwest"," nw")\
    .replace(" northeast"," ne")\
    .replace(" southwest"," sw")\
    .replace(" southeast"," se")\
    .replace(" s w "," sw ")\
    .replace(" n w "," nw ")\
    .replace(" s e "," se ")\
    .replace(" n e "," ne ")\
    .replace(" east "," e ")\
    .replace(" west "," w ")\
    .replace(" north "," n ")\
    .replace(" south "," s ")\
    .replace(" u s "," us ")\
    .replace(" u.s. "," us ")\
    .replace("straße ","str ")\
    .replace("strasse ","str ")\
    .replace("str. ","str ")\
    .replace(" suite "," ste ")\
    .replace("suite ","ste ")\
    .replace("ste #","ste ")\
    .replace("#","ste ")\
    .replace("building","bldg")\
    .replace("floor","flr")\
    .replace(" unit"," ste")\
    .replace("unit ","ste ")\
    .replace("=","'")\
    .replace("  "," ")\
    .replace("  "," ")\
    .replace("  "," ")\
    .replace("  "," ")\
    .replace("  "," ")\
    .replace("  "," ")

        address = re.sub('[^A-Za-z0-9\s]+', '', address)
        address = re.sub( '\s+', ' ', address)
    except AttributeError:
        address = ""
    return address
    
    
    
#This function cleans the city
def cleanCity(city):
    try:
        city = city.strip().lower()
        city = city.replace("&"," and ").replace("saint ","st ")\
    .replace("fort ","ft ")\
    .replace("saint ","st ")\
    .replace(".","")\
    .replace("-"," ")\
    .replace("north ","n ")\
    .replace("south ","s ")\
    .replace("east ","e ")\
    .replace("west ","w ")\
    .replace("mount ","mt ")\
    .replace("spring","spg")\
    .replace("height","ht")\


        city = re.sub('[^A-Za-z0-9\s]+', '', city)
        city = re.sub( '\s+', ' ', city)
    except AttributeError:
        city = ""
    return city

def noNames(df):        
    
  #  df['Name Score'] = df.apply(lambda row: cleanName(row['Location Name']), axis=1) 

    
    for index,row in df.iterrows():        
    #If name is blank, fills in last part of URL
        if row['Listing Name'] == None and row['Listing URL']!=None:
            df.loc[index,'Cleaned Listing Name'] = cleanName(row['Listing URL'].split('/')[-1])
            df.loc[index,'No Name'] = 'URL for name'
        elif row['Listing Name'] == None:
            df.loc[index,'No Name'] = 'No Name'


#df['Cleaned Listing Name']=df.apply(lambda x: x['Cleaned Listing Name'].replace\
#    (cleanName(x['Location City']),'') if cleanName(x['Location City']) in x['Cleaned Listing Name'] else\
#    x['Cleaned Listing Name'], axis=1)

    for index,row in df.iterrows():    
#    df.apply()
        #Removes City name if in Listing name
        if cleanName(row['Location City']) in row['Cleaned Listing Name']:
            df.loc[index,'Cleaned Listing Name']=row['Cleaned Listing Name']\
            .replace(cleanName(row['Location City']),'')

    df['Cleaned Location Name'] = df['Location Name'].apply(cleanName) 

    df['Cleaned Location Name'] = df['Location Name'].apply(cleanName) 
    df['Cleaned Listing Name'] = df['Listing Name'].apply(cleanName)

    
#This function compares the names in the file
def compareName(df, IndustryType, bid):
    df['Cleaned Location Name'] = df['Location Name'].apply(cleanName) 
    df['Cleaned Listing Name'] = df['Listing Name'].apply(cleanName)
    df['No Name'] = df['Listing Name'].isnull()
#df.apply(lambda row: np.isnan(row['Listing Name']), axis=1)

#    df.isna(df2['one'])
    # df.apply(noNames)
   
    df['Cleaned Listing Name']=df.apply(lambda x: x['Cleaned Listing Name'].replace\
    (cleanName(x['Location City']),'') if cleanName(x['Location City']) in x['Cleaned Listing Name'] else\
    x['Cleaned Listing Name'], axis=1)
    
    
    averagenamescore = []
    average=0
    
    #Populates businessNames with Account Name and alt Name Policies
    inputName = ''
    global businessNames
    
    businessNames = []
    businessNames.append(cleanName(getBusName(bid)))
    altNames = getAltName(bid)
    
    for name in altNames:
        businessNames.append(cleanName(name))

    global namesComplete
    global businessNameMatch
    
    print 'See popup!'
    #calls Tkinter input window for more business names. Waits for it to complete
    app.namesWindow(businessNames)
    
    app.wait_window(app.nameW)
   
    
    
#start of comparisons, broken out by industry
    #Normal Industry  
    if IndustryType == "0":
        print "Normal Naming"
        if businessNameMatch == 1:
            
            df['businessPartial']=0
            df['businessTokenSet'] =0
            df['businessTokenSort'] =0

            for bName in businessNames:
                    df['businessPartial']  = df.apply(lambda row: \
                        max(row['businessPartial'], fuzz.partial_ratio(bName, row['Cleaned Listing Name'])),axis=1)
                    df['businessTokenSet'] = df.apply(lambda row: \
                        max(row['businessTokenSet'], fuzz.token_set_ratio(bName, row['Cleaned Listing Name'])),axis=1)
                    df['businessTokenSort'] = df.apply(lambda row:\
                        max(row['businessTokenSort'], fuzz.token_sort_ratio(bName, row['Cleaned Listing Name'])),axis=1)
                  
                    
            df['Token Set'] = df.apply(lambda row: fuzz.token_set_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1) 
    
            df['Partial Score'] = df.apply(lambda row: fuzz.partial_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1) 
    
            df['Token sort'] = df.apply(lambda row: fuzz.token_sort_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)         
                    
            df['Name Score'] = df[['businessPartial','businessTokenSet','businessTokenSort',"Token Set", "Partial Score", "Token sort"]].max(axis=1)
            
            
            df=df.drop(['businessPartial','businessTokenSet','businessTokenSort',"Token Set", "Partial Score", "Token sort"],axis=1)         
#            ntpr = fuzz.partial_ratio(row['Cleaned Location Name'], row['Cleaned Listing Name'])
#            ntksr = fuzz.token_set_ratio(row['Cleaned Location Name'], row['Cleaned Listing Name'])
#            ntsr = fuzz.token_sort_ratio(row['Cleaned Location Name'], row['Cleaned Listing Name'])
#            
#                businessPartial = 0
#                businessTokenSet = 0
#                businessTokenSort = 0
#                #Check listing name against business names
#                for bName in businessNames:
#                    businessPartial  = max(businessPartial, fuzz.partial_ratio(bName, row['Cleaned Listing Name']))
#                    businessTokenSet = max(businessTokenSet, fuzz.token_set_ratio(bName, row['Cleaned Listing Name']))
#                    businessTokenSort = max(businessTokenSort, fuzz.token_sort_ratio(bName, row['Cleaned Listing Name']))
#                #Check listing name against location name
#                ntpr = fuzz.partial_ratio(row['Cleaned Location Name'], row['Cleaned Listing Name'])
#                ntksr = fuzz.token_set_ratio(row['Cleaned Location Name'], row['Cleaned Listing Name'])
#                ntsr = fuzz.token_sort_ratio(row['Cleaned Location Name'], row['Cleaned Listing Name'])
#                
#                #returns Max of Business Match or Location Name Match
#                total = max(businessPartial, businessTokenSet, businessTokenSort,ntpr,ntksr,ntsr)
#                averagenamescore.append(total)
#            df['Name Score'] = averagenamescore        
            
        else:
            df['Token Set'] = df.apply(lambda row: fuzz.token_set_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1) 
    
            df['Partial Score'] = df.apply(lambda row: fuzz.partial_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1) 
    
            df['Token sort'] = df.apply(lambda row: fuzz.token_sort_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1) 
    
            #returns Max of Business Match or Location Name Match
            df['Name Score Mean'] = df[["Token Set", "Partial Score", "Token sort"]].mean(axis=1)
    
            df['Name Score'] = df[["Token Set", "Partial Score", "Token sort"]].max(axis=1)

            df=df.drop(["Token Set", "Partial Score", "Token sort"],axis=1)         
       

    #Industry Hotel
    elif IndustryType == "2":
        print "Hotel Naming"
        BadHotel = ["optimeyes", "bakery","grill", "bar", "starbucks", "electric", "wedding", "gym",\
                     "pool", "restaurant", "bistro", "academy", "cafe", "salon",\
                     "5ten20", "lab", "rental", "car", "body", "fitness", "swim", "hertz", "steak",\
                     "sip", "zone", "alarm", "limestone", "catering", "room", "âme", "massage", "trivium",\
                     "broughton's", "broscheks", "vmug", "parking", "kenvigs", "martinis", "martini's"\
                     "beauty","formaggio","gallery", "motors",\
                     "sports", "formaggiosacramento", "rgs", "brasserie",\
                     "office", "vaso", "oceana", "vmware"\
                     "trivium", "fyve", "steakhouse", "ame", "wellness", "pay", "presentation"\
                     "presentations", "presentation", "visual",\
                     "tent", "eno", "copper", "coffee", "leisue","charter", "me", "ticketmaster",\
                     "swampers", "journeys", "friend", "orchards", "mandara", "camp"]     
#                     
        HotelBrands = ["ac hotels", "aloft", "america's best", \
        "americas best value", "ascend", "autograph", "baymont", "best western",\
        "cambria", "canadas best value", "candlewood", "clarion", "comfort inn",\
        "comfort suites", "country hearth", "courtyard", "crowne plaza", "curio",\
        "days inn", "doubletree", "econo lodge", "econolodge", "edition", "element",\
        "embassy", "even", "fairfield inn", "four points", "garden inn", "gaylord",\
        "hampton inn", "hilton", "holiday inn", "homewood", "howard johnson", "hyatt",\
        "indigo", "intercontinental", "jameson", "jw", "la quinta", "le meridien",\
        "le méridien", "lexington", "luxury collection", "mainstay", "marriott",\
        "microtel", "motel 6", "palace inn", "premier inn", "quality inn",\
        "quality suites", "ramada", "red roof", "renaissance", "residence", "ritz",\
        "rodeway", "sheraton", "signature Inn", "sleep inn", "springhill", "st regis",\
        "st. regis", "starwood", "staybridge", "studio 6", "super 8", "towneplace",\
        "value hotel", "value inn", "w hotel", "westin", "wingate", "wyndham"]        

        businessRatio = 0
        businessPartialRatio = 0
        
        df['Not Hotel'] = df['Cleaned Listing Name'].apply(lambda x: 1 if any(item in x.split() for item in BadHotel) else 0) 
        df['Other Hotel Match'] = df['Cleaned Listing Name'].apply(lambda x: 1 if any(item in x for item in HotelBrands) else 0)

        df['Name Score'] = df.apply(lambda row: \
         fuzz.token_set_ratio(row['Cleaned Location Name'], row['Cleaned Listing Name']) if ['Other Hotel Match'] == 1 else\
            fuzz.token_set_ratio(row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)         
        
        
    #Industry Healthcare Professional matching
    elif IndustryType == "3":
        print "HC Prof Naming"
        
        
#        nickNameExcelFile=pd.ExcelFile("~\Documents\Changing-the-World\nicknames.xlsx", keep_default_na = False)
#        nicknameList = nickNameExcelFile.parse('nicknames')
#        nicknameList = nicknameList.fillna("yext123")
#        nicknameList = nicknameList.applymap(lambda x : x.lower())
#        nicknameList = nicknameList.values.tolist()
#        
#        for index,row in df.itterows():
#        
#            locationNickNames=None   
#            locationName = (" " + str(row['Provider Name']) + " ").lower()
#            locationNickNames=set([tuple(group) for group in nicknameList for nick in group if nick in locationName.split(' ', 1)[0]])
#            
#            namescore=fuzz.token_set_ratio(row['Provider Name'], row['Cleaned Listing Name'])
#            for name in locationNickNames:
#                namescore=max(namescore,fuzz.token_set_ratio(name+" "+row['Provider Name'].split(' ', 1)[1], row['Cleaned Listing Name']))
#     
#            averagenamescore.append(namescore)
#            
#        df['Name Score'] = averagenamescore
#            
            
            
        df['Name Score'] = df.apply(lambda row: \
               fuzz.token_set_ratio(row['Provider Name'], row['Cleaned Listing Name']), axis=1)    
    #Industry Healthcare Facility matching

#    if IndustryType=="4":
#  return
    #Agent Names matching
    elif IndustryType == "5":
        print "Agent Naming"
        
#        nickNameExcelFile=pd.ExcelFile("~\Documents\Changing-the-World\nicknames.xlsx", keep_default_na = False)
#        nicknameList = nickNameExcelFile.parse('nicknames')
#        nicknameList = nicknameList.fillna("yext123")
#        nicknameList = nicknameList.applymap(lambda x : x.lower())
#        nicknameList = nicknameList.values.tolist()
#        
        
        if businessNameMatch == 1:
            
            df['businessPartial']=0
            df['businessTokenSet'] =0
            df['businessTokenSort'] =0

            for bName in businessNames:
                    df['businessPartial']  = df.apply(lambda row: \
                        max(row['businessPartial'], fuzz.partial_ratio(bName, row['Cleaned Listing Name'])),axis=1)
                    df['businessTokenSet'] = df.apply(lambda row: \
                        max(row['businessTokenSet'], fuzz.token_set_ratio(bName, row['Cleaned Listing Name'])),axis=1)
                    df['businessTokenSort'] = df.apply(lambda row:\
                        max(row['businessTokenSort'], fuzz.token_sort_ratio(bName, row['Cleaned Listing Name'])),axis=1)
                  
                    
            df['Token Set'] = df.apply(lambda row: fuzz.token_set_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1) 
    
            df['Partial Score'] = df.apply(lambda row: fuzz.partial_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1) 
    
            df['Token sort'] = df.apply(lambda row: fuzz.token_sort_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)         
                    
            df['Name Score'] = df[['businessPartial','businessTokenSet','businessTokenSort',"Token Set", "Partial Score", "Token sort"]].max(axis=1)
            
            
            df=df.drop(['businessPartial','businessTokenSet','businessTokenSort',"Token Set", "Partial Score", "Token sort"],axis=1)         

            
        else:
            df['Token Set'] = df.apply(lambda row: fuzz.token_set_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1) 
    
            df['Partial Score'] = df.apply(lambda row: fuzz.partial_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1) 
    
            df['Token sort'] = df.apply(lambda row: fuzz.token_sort_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1) 
    
            #returns Max of Business Match or Location Name Match
            df['Name Score Mean'] = df[["Token Set", "Partial Score", "Token sort"]].mean(axis=1)
    
            df['Name Score'] = df[["Token Set", "Partial Score", "Token sort"]].max(axis=1)

            df=df.drop(["Token Set", "Partial Score", "Token sort"],axis=1)         
        

    #Auto Name Matching
#    elif IndustryType == "6":
#        return        
    #Industry International    
    else:       
        print "Other Naming"
        if businessNameMatch == 1:
            
            df['businessPartial']=0
            df['businessTokenSet'] =0
            df['businessTokenSort'] =0

            for bName in businessNames:
                    df['businessPartial']  = df.apply(lambda row: \
                        max(row['businessPartial'], fuzz.partial_ratio(bName, row['Cleaned Listing Name'])),axis=1)
                    df['businessTokenSet'] = df.apply(lambda row: \
                        max(row['businessTokenSet'], fuzz.token_set_ratio(bName, row['Cleaned Listing Name'])),axis=1)
                    df['businessTokenSort'] = df.apply(lambda row:\
                        max(row['businessTokenSort'], fuzz.token_sort_ratio(bName, row['Cleaned Listing Name'])),axis=1)
                  
                    
            df['Token Set'] = df.apply(lambda row: fuzz.token_set_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1) 
    
            df['Partial Score'] = df.apply(lambda row: fuzz.partial_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1) 
    
            df['Token sort'] = df.apply(lambda row: fuzz.token_sort_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)         
                    
            df['Name Score'] = df[['businessPartial','businessTokenSet','businessTokenSort',"Token Set", "Partial Score", "Token sort"]].max(axis=1)
            
            
            df=df.drop(['businessPartial','businessTokenSet','businessTokenSort',"Token Set", "Partial Score", "Token sort"],axis=1)         
   
            
        else:
            df['Token Set'] = df.apply(lambda row: fuzz.token_set_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1) 
    
            df['Partial Score'] = df.apply(lambda row: fuzz.partial_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1) 
    
            df['Token sort'] = df.apply(lambda row: fuzz.token_sort_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1) 
    
            #returns Max of Business Match or Location Name Match
            df['Name Score Mean'] = df[["Token Set", "Partial Score", "Token sort"]].mean(axis=1)
    
            df['Name Score'] = df[["Token Set", "Partial Score", "Token sort"]].max(axis=1)

            df=df.drop(["Token Set", "Partial Score", "Token sort"],axis=1)         
            
        
#This function compares the countries in the file
def compareStateCountry(df):
    statematch = []
    for index, row in df.iterrows(): 
        # Drop different countries
        if (row['Input Location Country'] != row['Duplicate Location Country']):
            df.drop(index,inplace = True) 
        else:
            # Flag international locations
            if row['Input Location Country'] != "US":
                statematch.append("International")
            else:
                # Drop different states within US
                if (row['Location State'] != row['Listing State']):
                    df.drop(index,inplace = True)
                else:
                    statematch.append("US")
    df['US?'] = statematch

#This function compares the locationIds in the file
def compareId(df):
    IDmatch = []
    for index, row in df.iterrows(): 
        if (row['Input Location Id'] == row['Duplicate Location Id']):
            df.drop(index,inplace = True)
        else:
            IDmatch.append("Unique Pair")                
    df['Location ID Match'] = IDmatch

#This function removes matches that are not 1
def userMatch(df):
    print "nothing"
#    df.apply(lambda row: df.drop(index,inplace=True) if row['Match \n1 = yes, 0 = no'] != 1)


def compareStatus(df):    
    df['Claimed Score'] = df.apply(lambda x: 100 if x['Advertiser/Claimed'] == "Claimed" else 0, axis = 1)

#This function compares the phones in the file                
def comparePhone(df):
    
    
    df['Location Phone']=df.apply(lambda x: '%.12g' % x['Location Phone'] if isinstance(x['Location Phone'], float) else str(x['Location Phone']) ,axis=1)
    df['Location Local Phone']=df.apply(lambda x: '%.12g' % x['Location Local Phone'] if isinstance(x['Location Local Phone'], float) else str(x['Location Local Phone']),axis=1)
    df['Listing Phone']=df.apply(lambda x: '%.12g' % x['Listing Phone'] if isinstance(x['Listing Phone'], float) else str(x['Listing Phone']),axis=1)
    
    
    
    try:
        df['Phone Match'] = df.apply(lambda x: True \
            if (x['Location Phone'] == x['Listing Phone'] and \
                  (pd.isnull(x['Location Phone']) or  pd.isnull(x['Listing Phone']) or x['Listing Phone']==None or x['Listing Phone']=="nan" )==False)\
            else (True if (x['Location Local Phone'] == x['Listing Phone'] \
                           and pd.isnull(x['Location Local Phone'])==False \
                        and x['Listing Phone']!="nan" and  pd.isnull(x['Listing Phone'])==False and x['Listing Phone']!=None) else False), axis = 1)
        
        df['Phone Score'] = df.apply(lambda x: 100 if x['Phone Match'] else 0, axis = 1)
    except:
        df['Phone Match'] = 'x'   
     

#This function compares the addresses in the file                
def compareAddress(df,IndustryType):
    #International
    df['No Address'] = df['Listing Address'].isnull()        
    if IndustryType == '6':
        #Combine Address 1, Address 2
        df['Cleaned Input Address'] = df['Location Address'].apply(cleanAddress)\
                                    +' '+df['Location Address 2'].apply(cleanAddress) 
        df['Cleaned Listing Address'] = df['Listing Address'].apply(cleanAddress)\
                                    +' '+df['Listing Address 2'].apply(cleanAddress)
        #removes extra space where necessary
        df['Cleaned Input Address'] = df['Cleaned Input Address'].apply(cleanAddress)
        df['Cleaned Listing Address'] = df['Cleaned Listing Address'].apply(cleanAddress)
        averageaddressscore = []


        df['asr'] = df.apply(lambda row: fuzz.ratio(row['Cleaned Input Address'], row['Cleaned Listing Address']),axis=1)
        df['addressSortRatio'] = df.apply(lambda row:  fuzz.token_sort_ratio(row['Cleaned Input Address'], row['Cleaned Listing Address']),axis=1)
        df['apsr'] = df.apply(lambda row:  fuzz.partial_token_sort_ratio(row['Cleaned Input Address'], row['Cleaned Listing Address']),axis=1)
        
        df['Address Score'] = df[['asr','addressSortRatio','apsr']].max(axis=1)
        
        df=df.drop(['asr','addressSortRatio','apsr'],axis=1)
#        
#        for index, row in df.iterrows():                             
#                #Finds best match between normal ratio, sorted ratio
#                df['asr'] = df.apply(lamfuzz.ratio(row['Cleaned Input Address'], row['Cleaned Listing Address'])
#                addressSortRatio = fuzz.token_sort_ratio(row['Cleaned Input Address'], row['Cleaned Listing Address'])
#                apsr = fuzz.partial_token_sort_ratio(row['Cleaned Input Address'], row['Cleaned Listing Address'])
#
#                averageaddressscore.append(max(asr,addressSortRatio,apsr))
#        df['Address Score'] = averageaddressscore
        return
    #All other industries    
    else:        

        df['Cleaned Input Address'] = df['Location Address'].apply(cleanAddress) 
        df['Cleaned Listing Address'] = df['Listing Address'].apply(cleanAddress)
        
        #This is for the future when we have to parse out 
#        df['Input Street Number Score'] = df.apply(lambda row: \
#                                            [int(s) for s in row['Cleaned Input Address'].split() if s.isdigit()][:1], axis =1)
#        df['Listing Street Number Score'] = df.apply(lambda row: \
#                                            [int(s) for s in row['Cleaned Listing Address'].split() if s.isdigit()][:1], axis =1)
#
        
#        df['Token Set Ad'] = df.apply(lambda row: fuzz.token_set_ratio\
#            (row['Cleaned Input Address'], row['Cleaned Listing Address']), axis = 1) 
#        df['Partial Ad'] = df.apply(lambda row: fuzz.partial_ratio\
#            (row['Cleaned Input Address'], row['Cleaned Listing Address']), axis = 1) 
#        df['Sort Ad'] = df.apply(lambda row: fuzz.token_sort_ratio\
#            (row['Cleaned Input Address'], row['Cleaned Listing Address']), axis = 1) 
        df['Address Score'] = df.apply(lambda row: fuzz.token_set_ratio\
            (row['Cleaned Input Address'], row['Cleaned Listing Address']), axis = 1) 
        
        
#This function compares the cities in the file                
def compareCity(df):
    df['Cleaned Input City'] = df['Location City'].apply(cleanCity) 
    df['Cleaned Listing City'] = df['Listing City'].apply(cleanCity)
    averagecityscore = []

    df['csr'] = df.apply(lambda row: fuzz.ratio(row['Cleaned Input City'], row['Cleaned Listing City']),axis=1)
    df['ctpr'] =  df.apply(lambda row:fuzz.partial_ratio(row['Cleaned Input City'], row['Cleaned Listing City']),axis=1)
    df['City Score'] =df[['csr','ctpr']].mean(axis=1)
#    
#    for index, row in df.iterrows(): 
#        csr = fuzz.ratio(row['Cleaned Input City'], row['Cleaned Listing City'])
#        ctpr = fuzz.partial_ratio(row['Cleaned Input City'], row['Cleaned Listing City'])
#        
#        average = np.mean([csr,ctpr])
#        averagecityscore.append(average)   
#    df['City Score'] = averagecityscore

#This function compares the Country in the file                        
def compareCountry(df):
     df['Country Match'] = df.apply(lambda x: '0' if x['Input Location Country'] != x['Duplicate Location Country']\
                                     else '1', axis = 1)

#This function compares the Zips in the file                   
def compareZip(df):
     df['Zip Match'] = df.apply(lambda x: True if x['Location Zip'] == x['Listing Zip']\
                                 else False, axis = 1)
     
#This function compares the NPIs in the file
def compareNPI(df):
     df['NPI Match'] = df.apply(lambda x: True if x['Location NPI'] == x['Listing NPI'] else False, axis = 1)

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
    locationLat, locationLong, listingLat, listingLong =\
                map(radians, [locationLat, locationLong, listingLat, listingLong])

    # haversine formula
    dlon = listingLong - locationLong
    dlat = listingLat - locationLat
    a = sin(dlat/2)**2 + cos(locationLat) * cos(listingLat) * sin(dlon/2)**2
    c = 2 * asin(sqrt(a))
    r = 6371 # Radius of earth in kilometers. Use 3956 for miles
    return c * r * 1000

    
def calculateDoctorMatch(df):
    
    commonDoctorWords = ['md','pa','dr','do','np','phys','lpn','rn','dds','cnm','mph','phd','gp','dpm']
   # doctorSpecialty=pd.read_csv("~\Documents\Changing-the-World\SpecialtyDoctorMatching.csv")
    excelFile = pd.ExcelFile("~\Documents\entops\AutoMatcher\SpecialtyDoctorMatching.xlsx", keep_default_na = False)
    doctorSpecialty = excelFile.parse('Specialty')
    doctorSpecialty = doctorSpecialty.fillna("yext123")
    doctorSpecialty = doctorSpecialty.applymap(lambda x : x.lower())
    doctorSpecialty = doctorSpecialty.values.tolist()
    
    specialties=[]

    for index, row in df.iterrows(): 
        
        locationNameSpecialties=None
        listingNameSpecialties=None
        locationName = (" " + str(row['Cleaned Location Name']) + " ").lower()

        listingName = (" " + str(row['Cleaned Listing Name']) + " ").lower()
    
        if any([x for x in commonDoctorWords if x in listingName]): 
            specialties.append( "No Match - Doctor")
  #      elif any([x for x in punctuationWords if x in listingName]): specialties.append( "Check-Doctor")
        else:
            if any([x for x in [ 'Gift Shop', 'Cafe'] if x in listingName]): 
                specialties.append( "No Match - Excluded")
            else:
                locationNameSpecialties = set([tuple(group) for group in doctorSpecialty for specialty in group if specialty in locationName])
                listingNameSpecialties = set([tuple(group) for group in doctorSpecialty for specialty in group if specialty in listingName])
                

                if locationNameSpecialties:
                   if not listingNameSpecialties: 
                       specialties.append( "Check-Generic")
                   elif [locationNameSpecialties & listingNameSpecialties][0]: 
                       specialties.append( "Match-Specialty")
                   else: 
                       specialties.append( "No Match-Specialty")
            
                else:
                    if listingNameSpecialties: 
                        specialties.append( "Match-Specialty")
                    elif not listingNameSpecialties: 
                        specialties.append( "Check-Generic")
                    
               #this else seems out of place, but not sure where this goes.     
                    else:
                        specialties.append( "Check")

    df['Specialty Match']=specialties


    
#Potential Speed Optimizer of Passing in Rows instead of dataframes.    
#def compareEverything(df,IndustryType, bid):
#    averageaddressscore = []
#    phonescore = []
#    cleanedlocadd = []
#    cleanedlistingadd = []
#    print "working"
#    if IndustryType == '3':
#         NPIscore = []   
#    #    averagenamescore = []
#    for index, row in df.iterrows(): 

def compareData(df, IndustryType, bid):
    
    #compareId(df)
    print 'Comparing names...'
    compareName(df,IndustryType, bid)
    print 'Comparing phones...'
    comparePhone(df)
    compareStatus(df)
    #compareCountry(df)
    print 'Comparing zips...'
    compareZip(df)
    if IndustryType == '3':
        print 'Comparing NPIs...'
        compareNPI(df)

    if IndustryType == '4':
        print 'Checking specialties...'
        calculateDoctorMatch(df)
    print 'Comparing addresses...'
    compareAddress(df,IndustryType)
    #compareStateCountry(df)
    df['Distance (M)'] = df.apply(lambda row: calculateDistance(row), axis=1) 
    print 'Suggesting matches...'
    calculateTotalScore(df)
    suggestedmatch(df, IndustryType)

#This function provides a suggested match based on certain name/address thresholds
def suggestedmatch(df, IndustryType):  
    robotmatch = []
    
    #Creates new columns that are easier to read in code, and for users when outputed - Mostly boolean variables
    df['Address Match'] = df.apply(lambda x: True if x['Address Score'] >= 70 else False, axis=1)
    df['Name Match'] = df.apply(lambda x: 1 if x['Name Score'] >= 70 else (2 if 70 > x['Name Score'] >= 60 else 0), axis=1)
    df['Geocode Match'] = df.apply(lambda x: True if x['Distance (M)']<=200 else False, axis=1)
    df['Phone or Address Match']=df.apply(lambda x:  x['Phone Match'] or x['Address Match'],axis=1)
    
    liveSync = 'No Match - Live Sync'
    matchText = 'Match Suggested'
    noName = 'No Match - Name'
    noMatch = 'No Match'
    noAddress = 'No Match - Phone/Address'
    checkMissing = 'Check - No Name/Address'
    check = 'Check Name'
    noSpecialty = 'No Match - Specialty'
    checkSpecialty = 'Check Doctor/Specialty'
    npimatch = 'Match Suggested - NPI'
    
    #Normal Type
    if IndustryType == '0':        
        #Applies Match rules based on new columns.
        print "Normal Matching"
        df['Robot Suggestion'] = df.apply(lambda x: checkMissing if (x['No Name'] or x['No Address'])\
            else matchText if x['Name Match'] == 1 and (x['Phone Match'] or x['Address Match'] \
                or x['Geocode Match']) else (check if x['Name Match'] == 2 and (x['Phone Match'] or x['Address Match'] \
                or x['Geocode Match'])\
                else (noName if x['Name Match']==0 else (noAddress if not x['Address Match'] else 'uh oh'))) , axis=1)
        df['Match \n1 = yes, 0 = no'] = ""
        
    #Hotel matches each other
    #If hotel match another and phone and address         
    #Hotel Type
    #Need to Add Name Score taking the max of the brand
    elif IndustryType == '2':      
        print "Hotel Matching..."        
        df['Robot Suggestion'] = df.apply(lambda x: liveSync if x['Live Sync'] == 1 \
            else checkMissing if (x['No Name'] or x['No Address'])\
                else noName if x['Not Hotel'] == 1\
                    else noAddress if x['Address Match'] == False\
                        else noName if x['Name Match'] == 0 \
                            else check if x['Name Match'] == 2 and (x['Address Match'] == True or x['Phone Match'] == True)\
                                else matchText if x['Name Match'] == 1 and x['Address Match'] == True\
                                    and x['Phone Match'] == True and x['Other Hotel Match'] == 1\
                                        else matchText if x['Name Match'] == 1 \
                                        and x['Address Match'] == True or x['Phone Match'] == True\
                                                else noMatch ,axis = 1) 
        df['Match \n1 = yes, 0 = no'] = ""
        return
    #Healthcare Professional
    elif IndustryType == '3': 
        print "HC Prof Matching"
        df['Name Match'] = df.apply(lambda x: 1 if x['Name Score'] >= 76 else (2 if 76 > x['Name Score'] >= 66 else 0), axis=1)
       
        df['Robot Suggestion'] = df.apply(lambda x: npimatch if x['NPI Match'] == 1 \
            else checkMissing if (x['No Name'] or x['No Address'])
            else matchText if x['Name Match'] == 1 and (x['Phone Match'] or x['Address Match'] or x['Geocode Match']) \
            else (check if x['Name Match'] == 2 and (x['Phone Match'] or x['Address Match'] or x['Geocode Match'])
            else (noName if x['Name Match']==0 \
            else (noAddress if not x['Address Match'] else 'tbd'))) , axis=1)
        
#==============================================================================
#         for index, row in df.iterrows(): 
#             if row ['NPI Match'] :
#                 robotmatch.append("Match Suggested - NPI")
#             elif row['Phone Match'] == '1':
#                 if 66 < row['Name Score'] < 76 or row['Cleaned Listing Name'] is None:
#                     robotmatch.append("Check") 
#                 elif 76 <= row['Name Score']:
#                     robotmatch.append("Match Suggested") 
#                 else: 
#                     if row['No Name'] == 'URL for name':
#                         robotmatch.append('Check - URL name')
#                     else:
#                         robotmatch.append("No Match - Name")                         
#             elif row['Address Score'] < 70:
#                 if row['Distance (M)']<200:
#                     if 66 < row['Name Score'] < 76 or row['Cleaned Listing Name']is None:
#                         robotmatch.append("Check") 
#                     elif 76 <= row['Name Score']:
#                         robotmatch.append("Match Suggested - Geocode") 
#                     else: 
#                         if row['No Name'] == 'URL for name':
#                             robotmatch.append('Check - URL name')
#                         else:
#                             robotmatch.append("No Match - Name")
#                 else:
#                     robotmatch.append("No Match - Address")
#             else:    
#                 if 66 < row['Name Score'] < 76 or row['Cleaned Listing Name']is None:
#                     robotmatch.append("Check") 
#                 elif 76 <= row['Name Score']:
#                     robotmatch.append("Match Suggested") 
#                 else: 
#                     if row['No Name'] == 'URL for name':
#                         robotmatch.append('Check - URL name')
#                     else:
#                         robotmatch.append("No Match - Name")                         
#         
#         df['Robot Suggestion'] = robotmatch
#==============================================================================
        
        df['Name Match'] = df.apply(lambda x: True if x['Name Match'] == 1 \
                    else ('Check' if x['Name Match'] == 2 else False), axis=1)
        df['Match \n1 = yes, 0 = no'] = ""

 
#        for index, row in df.iterrows(): 
#            if row ['NPI Match'] :
#                robotmatch.append("Match - NPI")
#            elif row['Phone Match'] == '1':
#                if 66 < row['Name Score'] < 76 or row['Cleaned Listing Name'] is None:
#                    robotmatch.append("Check") 
#                elif 76 <= row['Name Score']:
#                    robotmatch.append("Match Suggested") 
#                else: 
#                    if row['No Name'] == 'URL for name':
#                        robotmatch.append('Check - URL name')
#                    else:
#                        robotmatch.append("No Match - Name")                         
#            elif row['Address Score'] < 70:
#               if row['Distance (M)']<200:
#                    if 66 < row['Name Score'] < 76 or row['Cleaned Listing Name']is None:
#                        robotmatch.append("Check") 
#                    elif 76 <= row['Name Score']:
#                        robotmatch.append("Match Suggested - Geocode") 
#                    else: 
#                        if row['No Name'] == 'URL for name':
#                            robotmatch.append('Check - URL name')
#                        else:
#                            robotmatch.append("No Match - Name")
#                else:
#                    robotmatch.append("No Match - Address")
#            else:    
#                if 66 < row['Name Score'] < 76 or row['Cleaned Listing Name']is None:
#                    robotmatch.append("Check") 
#                elif 76 <= row['Name Score']:
#                    robotmatch.append("Match Suggested") 
#                else: 
#                    if row['No Name'] == 'URL for name':
#                        robotmatch.append('Check - URL name')
#                    else:
#                        robotmatch.append("No Match - Name")                         
#        
#        df['Robot Suggestion'] = robotmatch
#       df['Match \n1 = yes, 0 = no'] = ""
#        return                    
    #Healthcare Facilities    
    elif IndustryType == '4':
        print "HC Facility Matching"
        df['Robot Suggestion'] = df.apply(lambda x: noSpecialty if x['Specialty Match'][0:2]=='No' \
                    else checkMissing if (x['No Name'] or x['No Address']) 
                    else((matchText if x['Name Match'] == 1 and (x['Phone Match'] or x['Address Match'] or x['Geocode Match'])\
                          else (check if x['Name Match']==2 \
                          else (noName if x['Name Match']==0 \
                          else (noAddress if not x['Address Match'] else 'uh oh') )))\
                          if x['Specialty Match'][0:5]=='Match' else \
                          ('Check Name and Specialty' if x['Name Match']==2 \
                          else (noName if x['Name Match']==0 \
                          else (noAddress if not x['Address Match'] else 'uh oh') ))),axis=1)
        
        df['Name Match'] = df.apply(lambda x: True if x['Name Match'] == 1 \
                    else ('Check' if x['Name Match'] == 2 else False), axis=1)
        df['Match \n1 = yes, 0 = no'] = ""
        return                        
    
    #International
    elif IndustryType == '6':
        print "Intl Matching"
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
                elif row['Country'] == 'GB':
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

        df['Name Match'] = df.apply(lambda x: True if x['Name Match'] == 1 \
                    else ('Check' if x['Name Match'] == 2 else False), axis=1)
        df['Match \n1 = yes, 0 = no'] = ""        
        return                        
    
     #All other industries
    else:
        print "Other Matching"
        #Applies Match rules based on new columns.
        df['Robot Suggestion'] = df.apply(lambda x: checkMissing if (x['No Name'] or x['No Address']) 
            else matchText if x['Name Match'] == 1 and (x['Phone Match'] or x['Address Match'] \
                or x['Geocode Match']) else (check if x['Name Match']==2 \
                else (noName if x['Name Match']==0 else (noAddress if not x['Address Match'] else 'uh oh'))) , axis = 1)


    df['Name Match'] = df.apply(lambda x: True if x['Name Match'] == 1 \
                    else ('Check' if x['Name Match'] == 2 else False), axis=1)
    df['Match \n1 = yes, 0 = no'] = ""
    return                        

def calculateTotalScore(df):
    #GET ALL THE SCORE THEN GIVE THEM WEIGHTING THEN CREATE A NEW TOTAL SCORE COLUMN    

    df['Total Score'] = df.apply(lambda row: float(row['Name Score'])*.6 + float(row['Address Score'])*.3\
                        + float(row['Phone Score'])*.1 + float(row['Claimed Score']), axis = 1)

#Calculates Match vs Anti-Match. Will need to take in FB page ID
def calculatePLStatus1(row, templateType, keepMatchNPL, BrandPageID, BrandPageVanity):
    #Antimatch any auto-anti-matched
    if row['Match Status'] == 1: return "Anti-Match"

    #Antimatch any Brand Page on ID or vanity
    for x in BrandPageID:
        if unicode(x) in row['Listing URL']: return "Anti-Match"
    for x in BrandPageID:
        if unicode(x) in row['External ID']: return "Anti-Match"    
    for x in BrandPageVanity:
        if row['Listing URL'] == "https://www.facebook.com/" + unicode(x): return "Anti-Match"
    for x in BrandPageVanity:
        if row['Listing URL'] == "https://www.facebook.com/" + unicode(x) + "/": return "Anti-Match"

    #If Suppression Template: 0 score = Anti-Match, All other = Suppress
    if templateType == 2:
        if row['Score'] == 0: return "Anti-Match"
        else: return "Suppress"

    return "Sync"


    
def ExternalID_De_Dupe(df):    
    #If Listing ID is matched to more than one location 
    df = df.sort_values(['Match','Listing ID','Total Score'], ascending = [True, True, True] )
#    print df['Match']
    df = df.reset_index(drop=True)
    
    for index,row in df.iterrows():      
        if index < df.shape[0]-1:
            if row['Match'] == 1 and df.iloc[index+1]['Match'] == 1:
                if row['Listing ID'] == df.iloc[index+1]['Listing ID']:
#                    row['Match'] = 0
                    df.set_value(index,'Match',0)
#    print df['Match']
    return df

        
#Reads CSV or XLSX files        
def readFile(xlsFile):
        
        print 'Reading file...'
        if xlsFile[-4:] == 'xlsx':
#            print "still reading"
#            wb = xlrd.open_workbook(xlsFile, on_demand=True,encoding_override="utf-8")
#            sNames = wb.sheet_names()        
#            wsTitle = "none"
##            for name in sNames:
##                 wsTitle = name
#      #Takes first sheet name, not last     
#            wsTitle=sNames[0
            x1 = pd.ExcelFile(xlsFile)
            df = x1.parse(0)

        elif xlsFile[-3:] == 'csv':
            df = pd.read_csv(xlsFile, encoding ='utf-8')
        else:
            raise Exception('What kind of file did you give me, bro?')
            sys.exit()
       
        #Finds Business ID
        bid = getBusIDfromLoc(df.loc[0,'Location ID'])
#        return df
        return df,bid

#Reads CSV or XLSX files        
def readMatchedFile(xlsFile):
        
        print 'Reading file...'
        if xlsFile[-4:] == 'xlsx':
#            wb = xlrd.open_workbook(xlsFile, on_demand=True,encoding_override="utf-8")
#            sNames = wb.sheet_names()        
#            wsTitle = "none"
##            for name in sNames:
##                 wsTitle = name
#      #Takes first sheet name, not last     
#            wsTitle=sNames[0]
            x1 = pd.ExcelFile(xlsFile)
            df = x1.parse(0)
            
        elif xlsFile[-3:] == 'csv':
            df = pd.read_csv(xlsFile, encoding ='utf-8')
        else:
            raise Exception('Please provide a csv or xlsx file.')
        return df
#        return df,bid
    

        
    
 #New main runtime function - broken out as previous functions now handled in Tkinter   
def main(df,IndustryType, bid):    
    row = 0 
    if IndustryType == '3':
    #Gets Providers First and Last name. Saves to column 'Provider Name'
        DoctorNameDF = getProviderName(df)
        df=df.merge(DoctorNameDF,on='Location ID', how='left')

    lastcol=df.shape[1]
    row=df.shape[0]

#Compares all, suggests matches    
    compareData(df,IndustryType, bid)
    
    print 'Creating matching questions...'
 #Completes Matching Question sheet   
    matchingNameQs = matchingQuestions(df)     
        
    FilepathMatch =  os.path.expanduser("~\Documents\Python Scripts\\"+ getBusName(bid)+\
    " AutoMatcher Output "+ str(date.today().strftime("%Y-%m-%d")) + " " + str(time.strftime("%H.%M.%S")) +".xlsx")

    columns=df.columns.tolist()
    reorder=['Location Name', 'Listing Name', \
    'Name Score', 'Location Address', 'Listing Address', 'Address Score', 'Distance (M)', \
    'Name Match','Phone Match', 'Address Match', 'Phone or Address Match', 'Geocode Match', 'Robot Suggestion','Match \n1 = yes, 0 = no']
    for col in reorder:    
        columns.append(columns.pop(columns.index(col)))
    df=df[columns]


    df=df.sort_values(by='Name Score')   
    df=df.sort_values(by='Robot Suggestion')
    print '/nWriting file...'
    writer = pd.ExcelWriter(FilepathMatch, engine='xlsxwriter',options={'strings_to_urls': False})
    df.to_excel(writer,sheet_name="Result", index=False,  encoding='utf8')
    matchingNameQs.to_excel(writer,sheet_name="Matching Questions", index=False)
    workbook  = writer.book
    worksheet = writer.sheets["Result"]
    worksheet.set_zoom(80)
    matchingSheet=writer.sheets['Matching Questions']

    print 'Formatting file...'
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
    formatGreen= workbook.add_format({'bg_color': '#c6efce', 'border' : 1, 'border_color': '#c0c0c0'}) #
   
    namescorecol = df.columns.get_loc("Name Score")
    addressscorecol = df.columns.get_loc("Address Score")
    namecol = df.columns.get_loc("Location Name")
    addresscol =  df.columns.get_loc("Location Address")
    robotcol =  df.columns.get_loc("Robot Suggestion")
    lastgencol = df.columns.get_loc("Match \n1 = yes, 0 = no")
    Lat =  df.columns.get_loc("Listing Latitude")
    LastPostDate=  df.columns.get_loc("Listing Latitude")
    nameMatchCol=df.columns.get_loc("Name Match")
    geocodeMatchCol=df.columns.get_loc("Geocode Match")

    
    worksheet.set_row(0, 45)

    #worksheet.set_column(Lat , LastPostDate, None, None, {'hidden': 1})
    worksheet.set_column(0, namecol-1, None, None, {'hidden': 1})
    #worksheet.set_column(lastcol+2, lastcol+3, None, None, {'hidden': 1})
    worksheet.set_column(namecol, namecol+1, 45)
    worksheet.set_column(namescorecol, namescorecol, 7.5)
    worksheet.set_column(addressscorecol, addressscorecol, 7.5)
    worksheet.set_column(addresscol, addresscol+1, 27)
    worksheet.set_column(robotcol, robotcol, 23.11)
    worksheet.set_column(robotcol+1, robotcol+1, 14.22)
    
    worksheet.autofilter(0,0,0,lastgencol)

    # Format Match columns
    worksheet.conditional_format(0, namecol, 0, lastgencol, {'type':'text',
                                'criteria': 'containing',
                                'value':    "Name",
                                'format':   formatOrange})
    worksheet.conditional_format(0, namecol, 0, lastgencol, {'type':'text',
                                'criteria': 'containing',
                                'value':    "Address",
                                'format':   formatPurp})
    # Format Score columns
    worksheet.conditional_format(0, namecol, 0, lastgencol, {'type':'text',
                                'criteria': 'containing',
                                'value':    "Score",
                                'format':   headerformat })
    
    worksheet.conditional_format(1, robotcol, row, lastgencol, {'type':'text',
                                        'criteria': 'begins with',
                                        'value':    "Match",
                                        'format':   formatBlue})
    worksheet.conditional_format(1, robotcol-1, row, lastgencol, {'type':'text',
                                        'criteria': 'begins with',
                                        'value':    "Check",
                                        'format':   formatYellow})
    worksheet.conditional_format(1, lastcol, row, lastgencol, {'type':'text',
                                        'criteria': 'begins with',
                                        'value':    "Unique",
                                        'format':   formatBlue})

    #Not match coloring
    worksheet.conditional_format(1, robotcol, row, lastgencol, {'type':'text',
                                        'criteria': 'begins with',
                                        'value':    "No Match",
                                        'format':   formatRed})
    worksheet.conditional_format(1, lastcol, row, lastgencol, {'type':'text',
                                        'criteria': 'begins with',
                                        'value':    "Same Location",
                                        'format':   formatRed})
    
    worksheet.conditional_format(1, nameMatchCol, row, geocodeMatchCol, {'type':'text',
                                        'criteria': 'begins with',
                                        'value':    "TRUE",
                                        'format':   formatGreen})
    worksheet.conditional_format(1, nameMatchCol, row, geocodeMatchCol, {'type':'text',
                                        'criteria': 'begins with',
                                        'value':    "FALSE",
                                        'format':   formatRed})
    worksheet.conditional_format(1, nameMatchCol, row, geocodeMatchCol, {'type':'text',
                                        'criteria': 'begins with',
                                        'value':    "Good",
                                        'format':   formatGreen})
    worksheet.conditional_format(1, nameMatchCol, row, geocodeMatchCol, {'type':'text',
                                        'criteria': 'begins with',
                                        'value':    "Bad",
                                        'format':   formatRed})
    worksheet.conditional_format(1, nameMatchCol, row, geocodeMatchCol, {'type':'text',
                                        'criteria': 'begins with',
                                        'value':    "Check",
                                        'format':   formatYellow})
    worksheet.freeze_panes(1, 0)
    
    matchingSheet.set_column(0,1,45)
    
    print 'Saving...'    
    try:
        writer.save()

        os.startfile(FilepathMatch)
#        os.system('start excel.exe "'+os.path.expanduser(FilepathMatch)+'"' % (sys.path[0], ))
        app.AllDone(df,1,"\nDone! Results have been wrizzled to your Excel file. 1love <3"+"\nMatching Template here:\n"+ FilepathMatch )
        print "\nDone! Results have been wrizzled to your Excel file. 1love <3"
        print "\nMatching Template here:"
        print FilepathMatch
   #Opens explorer window to path of output    
        subprocess.Popen(r'explorer /select,'+os.path.expanduser(FilepathMatch))
        
    except IOError:
        app.AllDone(df,1,"\nIOError: Make sure your Excel file is closed before re-running the script.")
        print "\nIOError: Make sure your Excel file is closed before re-running the script."          

#Pulls all location and listing match data
def sqlPull(bid,folderID,labelID,ReportType,IndustryType):
    print 'pulling data'
    #Pull Location info and Listing IDs
    #IF LISTINGS:
    if ReportType == 0:
        SQL_QueryMatches = open(os.path.expanduser("~/Documents/entops/AutoMatcher/SQL Data Pull/1. Pull Matches.sql")).read()
    #IF SUPPRESSION:
    elif ReportType==1 :
        SQL_QueryMatches = open(os.path.expanduser("~/Documents/entops/AutoMatcher/SQL Data Pull/1. Suppression - Pull Matches.sql")).read()    
    else:
        sys.exit()
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
    SQL_QueryListings = open(os.path.expanduser("~/Documents/entops/AutoMatcher/SQL Data Pull/2. Pull Listings Data.sql")).read()
    SQL_QueryListings=SQL_QueryListings.splitlines()
    
    for index, line in enumerate(SQL_QueryListings):
        SQL_QueryListings[index]=line.replace('@ListingIDs', ','.join(map(str, ListingIDs)) )
    if ReportType==1 :
        for index, line in enumerate(SQL_QueryListings):
            SQL_QueryListings[index]=line.replace('--left join warehouse.listings_additional_google_fields lagf on lagf.listing_id=wl.id', \
            'left join warehouse.listings_additional_google_fields lagf on lagf.listing_id=wl.id')            
        for index, line in enumerate(SQL_QueryListings):
            SQL_QueryListings[index]=line.replace('--AND (lagf.googlePlaceId not in (select tl.externalid from tags_listings tl join tags_listings_unavailable_reasons tlur on tlur.location_id = tl.location_id join alpha.tags_unavailable_reasons tur on tur.id=tlur.tagsunavailablereason_id where tlur.partner_id=715 and tur.showasWarning is false) or lagf.googlePlaceId is null)', \
            'AND (lagf.googlePlaceId not in (select tl.externalid from tags_listings tl join tags_listings_unavailable_reasons tlur on tlur.location_id = tl.location_id join alpha.tags_unavailable_reasons tur on tur.id=tlur.tagsunavailablereason_id where tlur.partner_id=715 and tur.showasWarning is false) or lagf.googlePlaceId is null)')            
        
            
            
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
    
    
    if ReportType==1:
        liveSyncSQL="SELECT tl.location_id as \"Location ID\", tl.partner_id as \"Publisher ID\", tl.externalId AS 'Sync External ID', tl.url AS 'Sync URL' from alpha.tags_listings tl "+ \
        "WHERE tl.tagsListingStatus_id=6 and tl.location_id in ("+','.join(SQL_DataMatches['Location ID'].to_csv(path=None,sep=',',index=False).split('\n'))[:-1]+");"
    
        Yext_Prod_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
        syncIDsDF = pd.read_sql(liveSyncSQL, con=Yext_Prod_DB)    
    
    
        
        
    #Combine results
    df=SQL_DataListings.merge(SQL_DataMatches,on='Listing ID',how='outer')
    if ReportType==1:
        df=df.merge(syncIDsDF,on=['Location ID','Publisher ID'],how='left')
        
    if IndustryType==3:
        NPI_SQL=open(os.path.expanduser("~/Documents/entops/AutoMatcher/SQL Data Pull/3. Pull NPI Data.sql")).read()
        NPI_SQL=NPI_SQL.splitlines()
        for index, line in enumerate(NPI_SQL):
            NPI_SQL[index]=line.replace('@bizid', str(bid))
        NPI_SQL = ' '.join(NPI_SQL)
        Yext_Prod_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
        NPI_Data = pd.read_sql(NPI_SQL, con=Yext_Prod_DB)       
        
        df=df.merge(NPI_Data,on=['Location ID'],how='left')
        
    return df

#Gets Alt Name policies at business level
def getAltName(bid):    
    SQL_AltNameQuery = open(os.path.expanduser("~/Documents/entops/AutoMatcher/SQL Data Pull/Alt Name policies.sql")).read()

    SQL_AltNameQuery=SQL_AltNameQuery.replace('@bizid', str(bid))
    try:    
        Yext_SMS_DB = MySQLdb.connect(host="127.0.0.1", port=5007, db="alpha")
        AltNames = pd.read_sql(SQL_AltNameQuery, con=Yext_SMS_DB)['policyString']
    except:
        AltNames = ""
    BusNames=[]
    for name in AltNames:
        BusNames.append(name)
    
    return BusNames      
    
#Gets Business Account name    
def getBusName(bid):    

    SQL_BusQuery='Select name from alpha.businesses where id='+str(bid)+';'
    try:    
        Yext_OPS_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
        busName = pd.read_sql(SQL_BusQuery, con=Yext_OPS_DB)['name'][0]
    except:    
        busName = "None"
    return busName  
    
#Gets Business ID from location ID
def getBusIDfromLoc(locationID):
    #locationID=locationID.replace("\'","")
    SQL_BusIDQuery='SELECT business_id from alpha.locations where id='+str(locationID).replace("\'","")+';'
    try:    
        Yext_OPS_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
        busID = pd.read_sql(SQL_BusIDQuery, con=Yext_OPS_DB)['business_id'][0]
    except:
        print "Connect to SDM!"
#        busID = "Dog"
        sys.exit()
    
    return busID    

   #Pulls First Name, and Last Name for Healthcare Professionals 
def getProviderName(df):
    print 'getting doctor names'
    SQL_DoctorNameQuery = open(os.path.expanduser("~/Documents/entops/AutoMatcher/SQL Data Pull/Doctor Name.sql")).read()    
    locationIDs = df['Location ID']
            
    SQL_DoctorNameQuery=SQL_DoctorNameQuery.replace('@locationIDs', ','.join(map(str, locationIDs)))    
    Yext_OPS_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
    DoctorNamesIDs = pd.read_sql(SQL_DoctorNameQuery, con=Yext_OPS_DB)

    return DoctorNamesIDs        
    
 #Finds Listing names that could be good to match to based on prevalence    
#==============================================================================
# def get_similarity(df, ind, col):
#     mapped = list(map(lambda x: fuzz.token_sort_ratio(x, df[col].loc[ind]), df[col]))
#     cond = (np.array(mapped) >= 70)
#     label = df[col][cond].iloc[0]
# 
#     return label
#==============================================================================
      
def matchingQuestions(df):
    
    Qdf=df[df['Robot Suggestion'].isin(['No Match - Name','Check Name'])]
    pivot=pd.pivot_table(Qdf,values='Link ID',index='Listing Name',aggfunc='count')

    listingNames=pd.DataFrame(pivot)
   
    ts=time.time()
    
    if listingNames.shape[0]>1 and listingNames.shape[0]<5000:
        listingNames.columns=['Count']    
        listingNames=listingNames.sort_values(by=['Count'],ascending=False)
        listingNames=listingNames.reset_index() 
        listingNames['Listing Name']=listingNames['Listing Name'].apply(cleanName)

        listingNames.columns = listingNames.columns.str.replace('\s+', '_')
        
        #mergedListNames=listingNames.groupby(lambda x: get_similarity(listingNames, x, 'Listing_Name'))['Count'].sum()

        alias = {l : r for l, r in itertools.product(listingNames.Listing_Name, listingNames.Listing_Name) if l < r and fuzz.token_sort_ratio(l, r) >= 70}
        series=listingNames.Count.groupby(listingNames.Listing_Name.replace(alias)).sum()
        
        mergedListNames=series.to_frame()
       
        mergedListNames=mergedListNames.reset_index()
        mergedListNames.columns=mergedListNames.columns=['Listing Name','Count']
        mergedListNames=mergedListNames.sort_values(by=['Count'],ascending=False)                                   
        mergedListNames= mergedListNames[mergedListNames['Count'] >= 5]
        mergedListNames = mergedListNames.replace(np.nan, '', regex=True)
        mergedListNames=mergedListNames[mergedListNames['Listing Name']!='']
      
        te=time.time()

        print te-ts
        
        return mergedListNames
    
    elif  not listingNames.empty:
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


def writeUploadFile(df):
    filePath =  os.path.expanduser("~\Documents\Python Scripts\\"+ \
                                   getBusName(getBusIDfromLoc(df.loc[0,'locationId']))+\
                                   " Upload Linkages "+ str(date.today().strftime("%Y-%m-%d")) \
                                    + " " + str(time.strftime("%H.%M.%S")) +".csv")

    print 'writing file'
    #writer = pd.ExcelWriter(filePath, engine='xlsxwriter')
    df.to_csv(filePath,sheet_name="Linkages", encoding='utf-8',index=False)
    return "\nUpload Linkages available: "+ filePath

def GoogleIDs(bid):
    
    #Gets IDs from GMB
    inputFile = open(os.path.expanduser("~/Documents/entops/Temporary Matching Template Workaround/GoogleIDs1.sql")).read()
    splitQueries= inputFile.split(";")
    SQL_Query715=splitQueries[1]
    SQL_Query715=SQL_Query715.splitlines()
    
    
    for index, line in enumerate(SQL_Query715):
        if line=="WHERE al.business_id = @BUSINESS_ID":
            SQL_Query715[index]=line.replace('@BUSINESS_ID', str(bid))
    SQLQuery715_3 = [x for x in SQL_Query715 if x.startswith("--") is False]
    SQLQuery715_4 = []
    for x in SQLQuery715_3:
        if '--' in x:
            SQLQuery715_4.append(x[0:x.index('-')])
        else:
            SQLQuery715_4.append(x)
        
        #Convert Query into string from list
    FinalQuery715 = ' '.join(SQLQuery715_4)
    
    Yext_Prod_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
    SQL_Data_715 = pd.read_sql(FinalQuery715, con=Yext_Prod_DB)
        
    
    
    #Gets listing IDs for Google Reviews
    inputFile = open(os.path.expanduser("~/Documents/entops/Temporary Matching Template Workaround/GoogleIDs2.sql")).read()
    splitQueries= inputFile.split(";")
    SQL_Query713=splitQueries[1]
    SQL_Query713=SQL_Query713.splitlines()
    
    for index, line in enumerate(SQL_Query713):
        if line=="WHERE al.business_id = @BUSINESS_ID":
                SQL_Query713[index]=line.replace('@BUSINESS_ID', str(bid))
                
                
    SQLQuery713_3 = [x for x in SQL_Query713 if x.startswith("--") is False]
    SQLQuery713_4 = []
    for x in SQLQuery713_3:
        if '--' in x:
            SQLQuery713_4.append(x[0:x.index('-')])
        else:
            SQLQuery713_4.append(x)
        
        #Convert Query into string from list
    FinalQuery713 = ' '.join(SQLQuery713_4)
    
    Yext_Prod_DB = MySQLdb.connect(host="127.0.0.1", port=5009, db="alpha")
    SQL_Data_713 = pd.read_sql(FinalQuery713, con=Yext_Prod_DB)    
    
    
    listingIDs=SQL_Data_713.iloc[:,0:1]
    listingIDs.columns=['ID']       
    LiDs = listingIDs["ID"].tolist()
    LiDs = ','.join(str(e) for e in LiDs)
    
    #Gets external IDs for Google Reviews listing IDs
    inputFile = open(os.path.expanduser("~/Documents/entops/Temporary Matching Template Workaround/GoogleIDs3.sql")).read()
    splitQueries= inputFile.split(";")
    SQL_Query_ext=splitQueries[0]
    SQL_Query_ext=SQL_Query_ext.splitlines()
    
    for index, line in enumerate(SQL_Query_ext):
        if line=="where wl.id in (123456)":
            SQL_Query_ext[index]=line.replace('123456', LiDs)
                
                
    SQL_Query_ext_3 = [x for x in SQL_Query_ext if x.startswith("--") is False]
    SQL_Query_ext_4 = []
    for x in SQL_Query_ext_3:
        if '--' in x:
            SQL_Query_ext_4.append(x[0:x.index('-')])
        else:
            SQL_Query_ext_4.append(x)
        
        #Convert Query into string from list
    FinalQuery_ext = ' '.join(SQL_Query_ext_4)
    
    Yext_Prod_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
    SQL_Data_ext = pd.read_sql(FinalQuery_ext, con=Yext_Prod_DB)    
    
    
    #combines lists
    col1 = SQL_Data_715.iloc[:,0:1]
    col2 = SQL_Data_ext.iloc[:,0:1]
    
    col1.columns=["Live Google External IDs"]
    col2.columns=["Live Google External IDs"]       
    
    SQL_Data = pd.concat([col1, col2], ignore_index=True)
    SQL_Data = SQL_Data.drop_duplicates()
    
    SQL_Data["Live Google External IDs"] = "'" + SQL_Data["Live Google External IDs"].astype(str)
        
        # SQL_Data.to_excel()
    newSQL_Data = SQL_Data.set_index('Live Google External IDs')
    
    #print(newSQL_Data)
    Yext_Prod_DB.close()
    
    return newSQL_Data
    
class MatchingInput(Tkinter.Frame):

    def __init__(self, master):

        
        root.protocol("WM_DELETE_WINDOW", self._delete_window)
        Tkinter.Frame.__init__(self, master, padx=10, pady=10)
#        self.grid()
        master.title("AutoMatcher Setup")

        master.minsize(width=500, height=300)
        
       # style = Style()
       # style.theme_use('classic')
        
#First screen - needs to explain what's going on, get input on where in process they are        
#        for r in range(6):
#            self.master.rowconfigure(r, weight=1)    
#        for c in range(5):
#            self.master.columnconfigure(c, weight=1)
#                    
#        Frame1 = Frame(master)
#        Frame1.grid(row = 0, column = 0, rowspan = 3, columnspan = 3, sticky = W+E+N+S) 
        
        self.IntroLabel = Label(master,text="   Welcome to the AutoMatcher! This will suggest matches"\
                                +" based on inputs as well as create an upload file"+
                                "\n\n   To start, select the first option. "\
                                +"\n   Review all checks and fill in a match status."
                                +"\n   Afterwards, run the program again, and select the second option."\
                                +"\n\n   Make sure you connect to SDM and the J:\ Drive!").grid(row=0,column=0, columnspan=2,pady=(0,20))
        self.processChoice = IntVar()
        self.processChoice.set(-1)
        self.StartMatch = Radiobutton(master,text="Pull data/enter file and suggest matches",\
                                      variable = self.processChoice,value = 0).grid(row = 1,column = 0)
        self.ChecksChoice = Radiobutton(master,text="Create upload file based on reviewed matches",\
                                        variable=self.processChoice,value=1).grid(row = 1,column = 1)
        
        self.Next=Button(master,text = "Next",command = lambda: [self.initialSettingsWindow() \
                                     if self.processChoice.get() == 0 else (self.inputChecks() \
                                      if self.processChoice.get() == 1 else self.processChoice.set(-1))]).grid(row = 2,column = 0,pady=(30,0),sticky=E)

        self.Quit=Button(master,text="Quit",command = lambda: [root.destroy()]).grid(row = 3,column = 0,sticky=E, pady=25)
        
#If the user has manually checked the matches, this will take those in, determine Match statuses, and produce upload document        
    def inputChecks(self):
        self.master.withdraw()
        self.UploadSetup=Toplevel()
        self.UploadSetup.protocol("WM_DELETE_WINDOW", self._delete_window)
        
        self.ReportType=IntVar()
        self.ReportType.set(-1)
        self.ReportLabel = Label(self.UploadSetup,text="Select report type:").grid(row=4,column=0,pady=(30,0),sticky=W)
        
        self.Listings=Radiobutton(self.UploadSetup, text="Listings", variable=self.ReportType,value=0).grid(row=5,column=0,sticky=W)
        self.Suppression=Radiobutton(self.UploadSetup, text="Suppression", variable=self.ReportType,value=1).grid(row=6,column=0, sticky=W)
    
        self.nextButton = Button(self.UploadSetup, text="Next", command=lambda: [self.readCheckedFile() \
                            if (self.ReportType.get() >-1) else self.ReportType.set(self.ReportType.get())])\
                                                        .grid(row=7,column=0,sticky=W,pady=(35,0))
        
    def readCheckedFile (self):         
        self.UploadSetup.destroy()
#Takes in completed matches file with checks filled out
        checkedFile = tkFileDialog.askopenfilename(initialdir = "/",title =\
                                 "Select completed matching file with Check column filled out",\
                                 defaultextension="*.xlsx;*.xls", \
                                 filetypes=( ("Excel files", "*.xlsx;*.xls"), ("CSV", "*.csv"),('All files','*.*') ))        
        checkedDF = readMatchedFile(checkedFile)

        
#Checks to see if all rows asking for a check have manual review        
        allChecksComplete = True
        
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
            checkedDF['Match'] = checkedDF.apply(lambda x: 1 if 'Match Suggested' in x['Robot Suggestion'] else 0, axis=1)
            checkedDF['Match'] = checkedDF.apply(lambda x: 1 if x['Match \n1 = yes, 0 = no'] == 1 else \
                                                (0 if x['Match \n1 = yes, 0 = no'] == 0 else x['Match']), axis=1)
           
            
            busID=getBusIDfromLoc(checkedDF.loc[0,'Location ID'])
            busName=getBusName(busID)
            completeCopyFilePath =  os.path.expanduser("J:\zAutomatcherData\CheckedFiles\\"+ \
                                   os.getenv('username')+ " - " + busName+" Data "+ str(date.today().strftime("%Y-%m-%d")) \
                                    + " " + str(time.strftime("%H.%M.%S")) +".csv")

            #writer = pd.ExcelWriter(filePath, engine='xlsxwriter')
            checkedDF.to_csv(completeCopyFilePath,sheet_name="Data", encoding='utf-8',index=False)
            
            
            
            print 'checking live sync'
            
            googleLiveIDs=GoogleIDs(busID)
            googleLiveIDs= googleLiveIDs.reset_index()
            checkedDF['Google Live Sync']=checkedDF['External ID'].isin(googleLiveIDs['Live Google External IDs'])
            checkedDF['Live Sync']=checkedDF.apply (lambda x: 1 if x['Google Live Sync'] else x['Live Sync'],axis=1)
                                
            checkedDF['Match'] = checkedDF.apply(lambda x: 0 if x['Live Sync'] == 1 else x['Match'], axis=1)
            checkedDF['Match'] = checkedDF.apply(lambda x: 0 if x['Live Suppress'] == 1 else x['Match'], axis=1)

                
            print "External ID deduping\n"

            checkedDF = ExternalID_De_Dupe(checkedDF)
            
            

            
#            
            checkedDF['override'] = checkedDF.apply(lambda x: 'Match' if x['Match'] == 1 else 'Antimatch',axis=1)
            
            if self.ReportType.get()==0:
                checkedDF['PL Status'] = checkedDF.apply(lambda x: 'Sync' if x['override'] == 'Match' else 'NoPowerListing',axis=1)
            elif self.ReportType.get()==1:
                checkedDF['PL Status'] = checkedDF.apply(lambda x: 'Suppress' if x['override'] == 'Match' else 'NoPowerListing',axis=1)    
            
            #Total Score at the top
            checkedDF = checkedDF.sort_values(['Location ID','Publisher ID', 'Match', 'Total Score']\
                                                              , ascending=[True, True,False, False])
            checkedDF = checkedDF.reset_index(drop=True)   
         
            #If Listings Type
            if self.ReportType.get() == 0:    
                for index,row in checkedDF.iterrows():
                    if index != 0:
                    #If Listings, Diagnostic. or Facebook Template: if previous listing is a Match of the same Loc/Pub ID combination, then NPL
                        if row['Location ID'] == checkedDF.iloc[index-1]['Location ID'] and row['Publisher ID']\
                             == checkedDF.iloc[index-1]['Publisher ID'] and row['override'] == 'Match':\
                            checkedDF.set_value(index,"PL Status","NoPowerListing")
                        else: 
                            pass
                        

            #EXTERNAL ID DEDUPE
            print "Creating Upload Overrides\n"
            
            #checkedDF=calculateTotalScore(checkedDF)
            uploadDF = checkedDF[['Publisher ID','Location ID','Listing ID','override','PL Status']]
            uploadDF.columns=[ "partnerId","locationId", "listingId", "override","PL Status"]
            
            try:
                uploadDF['listingId'] = uploadDF.apply(lambda x: x['listingId'].replace('\'',''),axis=1)            
            except:
                pass                

#            If Suppression Type
            if self.ReportType.get() == 1:    
                print "Creating Suppression Report"
                
                
                suppRepDF = checkedDF[checkedDF['PL Status'] == 'Suppress']

                ################Create a Suppression Approval File 
                pivot = pd.pivot_table(suppRepDF, values='PL Status',index='Publisher',aggfunc='count')
#                print pivot
                publisherNames = pd.DataFrame(pivot)
                publisherNames.columns=['Count of Duplicates']

                publisherNames = publisherNames.sort_index(ascending=True)
                print publisherNames

#                publisherNames = publisherNames.reset_index() 
                businessName=busName
                
                
#                Total_format = workbook.add_format({'bold': True})
#                
#                worksheet.set_row (0,None,Total_format)
#
#
#                Total_format.set_pattern(1)
#                Total_format.set_bg_color('#D9E1F2')
#                
#                
#                worksheet.set_row(len(publisherNames)+3,None, Total_format)
                
               
                
               

                #Needs logic if certain things exist
                claimedFBDF = checkedDF[(checkedDF['PL Status'] == 'Suppress')\
                                        & (checkedDF['Publisher'] == 'Facebook') & (checkedDF['Advertiser/Claimed'] == 'Claimed')]
                print "mixing mixing"
                if claimedFBDF.shape[0] > 0:
                
                    
                    filePath =  os.path.expanduser("~\Documents\Python Scripts\\" + businessName+\
                                               " Suppression Approval File "+ str(date.today().strftime("%Y-%m-%d")) + " " + str(time.strftime("%H.%M.%S")) +".xlsx")
                else:
                     filePath =  os.path.expanduser("~\Documents\Python Scripts\\" + businessName+\
                                               " Suppression Summary File "+ str(date.today().strftime("%Y-%m-%d")) + " " + str(time.strftime("%H.%M.%S")) +".xlsx")
        
                suppwriter = pd.ExcelWriter(filePath, engine='xlsxwriter')
                publisherNames.to_excel(suppwriter,sheet_name="Summary", index=True,  encoding='utf8', startrow=2)
                workbook  = suppwriter.book
                worksheet = suppwriter.sheets['Summary']  
               # format1= workbook.add_format({
              #                         'bold': False,})  
              #  
                worksheet.set_column(0,0,26)
                worksheet.set_column(1,1,17)
                worksheet.write(0, 0, businessName+ " - Suppression Approval Summary - "+ str(date.today().strftime("%Y-%m-%d")))

                worksheet.write(len(publisherNames)+3, 0, "Grand Total")
                worksheet.write(len(publisherNames)+3, 1, publisherNames['Count of Duplicates'].sum())
                worksheet.set_zoom(80)    
                
                
                if claimedFBDF.shape[0] > 0:
                    claimedFBDF = claimedFBDF[['Store ID', 'Location ID',\
                                               'Location Name', 'Location Address', 'Location Address 2',\
                                               'Location City', 'Location State', 'Location Zip', 'Location Phone',\
                                                'Sync External ID', 'Sync URL', 'Listing Name', \
                                               'Listing Address','Listing Address 2', 'Listing City', 'Listing State', \
                                               'Listing Zip', 'Listing Phone', \
                                               'Listing URL','Listing ID','External ID',\
                                                'Last Post Date']]
                    claimedFBDF['Yes/No']=""
                    claimedFBDF['Reason']=""
                    
                    
               
                                                
    #                                            nocols = ['Location URL','Yes/No', 'Reason']

                    claimedFBDF.to_excel(suppwriter,sheet_name="Facebook Claimed Pages", index=False,startrow=0 )
                    worksheet = suppwriter.sheets['Facebook Claimed Pages']                
                    
#                    worksheet.set_column(0,26,20)
#                    worksheet.set_zoom(80)
#                    
#                    format = workbook.add_format()
#                    format.set_center_across()
#                    
#                    for col in range(0,29):
#                        worksheet.write_blank(0, col, '', format)
#                        
#                    worksheet.write(0, 0, 'Location Information in Yext', format)
#                    worksheet.write(0, 13, 'Listing Information on Publishers', format)
#                    worksheet.write(0, 27, 'Approval', format)
#                    worksheet.write(0, 0, 'Location Information in Yext', format)
#                    

                    
                    
                suppwriter.save()
        
           
                xlApp= win32com.client.Dispatch("Excel.Application")
                SuppFile = xlApp.Workbooks.Open(filePath)
                toolkit=os.path.expanduser("~\Documents\entops\Templates and Macros\EntOpsMacroToolkit.xlam")
                a=xlApp.Application.Run('\''+toolkit+"\'!AutoMatcherSuppressionReport.Run_Suppression_Report")

                
                reportDF = checkedDF[['Store ID', 'Location ID',\
                                           'Location Name', 'Location Address', 'Location Address 2',\
                                           'Location City', 'Location State', 'Location Zip', 'Location Phone',\
                                            'Sync External ID', 'Sync URL', 'Listing Name', \
                                           'Listing Address','Listing Address 2', 'Listing City', 'Listing State', \
                                           'Listing Zip', 'Listing Phone', \
                                           'Listing URL','Listing ID','External ID',\
                                            'Last Post Date','Publisher ID', 'Publisher']]
                
                
                fullFilePath =  os.path.expanduser("~\Documents\Python Scripts\\" + businessName+\
                                               " Full Suppression Details "+ str(date.today().strftime("%Y-%m-%d")) + " " + str(time.strftime("%H.%M.%S")) +".xlsx")
        
                reportWriter = pd.ExcelWriter(fullFilePath, engine='xlsxwriter')
                publisherNames.to_excel(reportWriter,sheet_name="Summary", index=True,  encoding='utf8', startrow=2)
                workbook  = reportWriter.book
                worksheet = reportWriter.sheets['Summary']  
               # format1= workbook.add_format({
              #                         'bold': False,})  
              #  
                worksheet.set_column(0,0,26)
                worksheet.set_column(1,1,17)
                worksheet.write(0, 0, businessName+ " - Suppression Approval Summary - "+ str(date.today().strftime("%Y-%m-%d")))

                worksheet.write(len(publisherNames)+3, 0, "Grand Total")
                worksheet.write(len(publisherNames)+3, 1, publisherNames['Count of Duplicates'].sum())
                worksheet.set_zoom(80)                            
#                                            nocols = ['Location URL','Yes/No', 'Reason']

                reportDF.to_excel(reportWriter,sheet_name="Facebook Claimed Pages", index=False,startrow=0 )
                worksheet1 = reportWriter.sheets['Facebook Claimed Pages'] 
                reportWriter.save()
               
                xlApp1= win32com.client.Dispatch("Excel.Application")
                reportFile = xlApp1.Workbooks.Open(fullFilePath)
                toolkit=os.path.expanduser("~\Documents\entops\Templates and Macros\EntOpsMacroToolkit.xlam")
                a=xlApp.Application.Run('\''+toolkit+"\'!AutoMatcherSuppressionReport.Run_Suppression_Report")
            
            print "Writing to File"
            self.AllDone(uploadDF, 2, writeUploadFile(uploadDF))                
            print t0
            t1=time.time()
            print t1
            print t1-t0
       #remove this once we do something here     
               
            #Gets user input for how to set up matcher           
    def initialSettingsWindow(self):
        
        self.master.withdraw()
        self.settingWindow = Toplevel()
        self.settingWindow .protocol("WM_DELETE_WINDOW", self._delete_window)
        self.IndustryType=IntVar()
        self.IndustryType.set(-1)
        self.IndustryLabel = Label(self.settingWindow,text="Select Industry Type").grid(row = 0, column = 0, columnspan=2, pady=(10,10), sticky=W)

        self.Normal=Radiobutton(self.settingWindow, text="Normal", variable=self.IndustryType,value=0).grid(row=1,column=0,sticky=W)
        self.Auto=Radiobutton(self.settingWindow, text="Auto", variable=self.IndustryType,value=1).grid(row=2,column=0,sticky=W)
        self.Hotel=Radiobutton(self.settingWindow, text="Hotel", variable=self.IndustryType,value=2).grid(row=2,column=1,sticky=W)
        self.Doctor=Radiobutton(self.settingWindow, text="Healthcare Doctor", variable=self.IndustryType,value=3).grid(row=2,column=2,sticky=W)
        self.Facility=Radiobutton(self.settingWindow, text="Healthcare Facility", variable=self.IndustryType,value=4).grid(row=3,column=2,sticky=W)
        self.Agent=Radiobutton(self.settingWindow, text="Agent", variable=self.IndustryType,value=5).grid(row=3,column=1,sticky=W)
        self.International=Radiobutton(self.settingWindow, text="International", variable=self.IndustryType,value=6).grid(row=3,column=0,sticky=W)
      
        #Report Type Designation
        self.ReportType=IntVar()
        self.ReportType.set(-1)
        self.ReportLabel = Label(self.settingWindow,text="Select report type:").grid(row=4,column=2,pady=(30,0),padx=(15,0),sticky=W)
        
        self.Listings=Radiobutton(self.settingWindow, text="Listings", variable=self.ReportType,value=0).grid(row=5,column=2,padx=(15,0),sticky=W)
        self.Suppression=Radiobutton(self.settingWindow, text="Suppression", variable=self.ReportType,value=1).grid(row=6,column=2, padx=(15,0),sticky=W)
        #self.FB=Radiobutton(self.settingWindow, text="FB", variable=self.ReportType,value=2).grid(row=1,column=2)
        #self.Google=Radiobutton(self.settingWindow, text="Google", variable=self.ReportType,value=x).grid(row=1,column=3)
        
        
        #self.quitButton.pack()
        self.dataInput=IntVar()
        self.dataInput.set(0)
        self.inputType=Label(self.settingWindow,text="Select data input type:").grid(row=4,column=0, columnspan=2,pady=(30,0), sticky=W)

        self.SQL=Radiobutton(self.settingWindow, text="Pull Data from SQL", variable=self.dataInput,value=2).grid(row=5,column=0, columnspan=2, sticky=W)
        self.file=Radiobutton(self.settingWindow, text="Input File", variable=self.dataInput,value=1).grid(row=6,column=0, sticky=W)
        

        
        self.nextButton = Button(self.settingWindow, text="Next", command=lambda: [self.detailsWindow() \
                            if (self.IndustryType.get() >-1 and self.ReportType.get() >-1 and self.dataInput.get() >-1)\
                                                         else self.ReportType.set(self.ReportType.get())])\
                                                        .grid(row=7,column=0,sticky=W,pady=(35,0))
     
                                                                                                    
#        self.inputType.grid(row=2,column=0)
#        self.SQL.grid(row=3,column=1, sticky=W)
#        self.file.grid(row=4,column=1, sticky=W)
#        self.nextButton.grid(row=5,column=0,sticky=W,pady=25)

 #Gets File path or business, folder, and label ids to pull from SQL       
    def detailsWindow(self):
        self.settingWindow.destroy()
        #File pull
        if self.dataInput.get() == 1:
            fname = tkFileDialog.askopenfilename(initialdir = "/",title = "Select file",filetypes=(("CSV", "*.csv"), ("Excel files", "*.xlsx;*.xls") ))
            df, bid = readFile(fname)
            self.bid=bid
            main(df,str(self.IndustryType.get()),bid)
        #SQL
        elif self.dataInput.get() == 2:
            self.detailsW=Tkinter.Toplevel(self)
            self.detailsW .protocol("WM_DELETE_WINDOW", self._delete_window)            
            vcmd = self.master.register(self.validate)

            self.InputExplanation=Label(self.detailsW,text=\
            "Please enter account info for data to pull. Business ID is required. \nYou can leave Folder ID and Label ID blank or put 0 if not needed\n").grid(row=5,column=0, columnspan=2)

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
            self.bid=self.bizID.get()
            self.detailsW.destroy()
            if  self.folderID.get():
                folderID=self.folderID.get()
            else: 
                folderID = 0
            if self.labelID.get():
                labelID=self.labelID.get()
            else:
                labelID = 0
            df = sqlPull(self.bizID.get(),folderID,labelID, self.ReportType.get(),self.IndustryType.get())
            main(df,str(self.IndustryType.get()),self.bizID.get())
    
#button function to ammend businessNames list            
    def AddMore(self):
        global businessNames
        global businessNameMatch
        
        if self.NewName.get() != "":
            for i in self.NewName.get().split(","):
                businessNames.append(cleanName(i)) 
        
        businessNameMatch=self.busNameMatch.get()
       # businessNames.append(cleanName(NewName))
        #AddNameBox.delete(0, END)
        self.nameW.destroy()
       # self.namesWindow(businessNames)
 
  #Prints out current business names, asks for additional names  
    def namesWindow(self,busNames):
        global businessNames
        global businessNameMatch
        
        self.complete=BooleanVar()
        self.complete=False
        self.nameW=Toplevel()
        self.nameW.protocol("WM_DELETE_WINDOW", self._delete_window) 
        self.nameW.minsize(width=300, height=200)
        self.NewName=StringVar()
        
        self.busNameMatch=IntVar()
        self.busNameMatch.set(-1)
        self.busMatchLabel=Label(self.nameW,text="Do you want to match to generic business names?").grid(row=0,column=0, sticky=W, columnspan=2)
        self.yesMatch=Radiobutton(self.nameW,text="Yes, match to business names", variable=self.busNameMatch , value=1).grid(row=1,column=0,pady=(0,20),sticky=W)
        self.noMatch=Radiobutton(self.nameW,text="No, do not match to business names", variable=self.busNameMatch ,value=0).grid(row=1,column=1,pady=(0,20),sticky=E)
        self.CurrentNameLabel=Label(self.nameW,text="Current Business Names:\n"+ ", ".join([str(i) for i in businessNames]) ).grid(row=10,column=0, columnspan=2)
        self.AddMoreLabel=Label(self.nameW,text="If needed, enter comma separated list of additional business names to match to:").grid(row=11,column=0, columnspan=2)
        self.AddName=Entry(self.nameW,textvariable=self.NewName).grid(row=12,column=0, columnspan=2)
        self.AddMoreButton=Button(self.nameW,text="Next",command=lambda: [self.AddMore() if self.busNameMatch.get() >-1 else self.busNameMatch.set(-1)]).grid(row=13,column=0, columnspan=2 ,pady=(20,0))
#        self.Done=Button(self.nameW,text="No more names needed",command= lambda: [self.AddMore() if self.busNameMatch.get() >-1 else self.busNameMatch.set(-1)]).grid(row=13,column=1, sticky=W)  
  
    
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
    def AllDone(self,df,part,msg):
        global t0
        t1=time.time()
        print "start: "+datetime.datetime.fromtimestamp(t0).strftime('%Y-%m-%d %H:%M:%S')
        print "end: "+ datetime.datetime.fromtimestamp(t1).strftime('%Y-%m-%d %H:%M:%S')
        print str(t1-t0)+" seconds"
        
        
        try:
            if part==1:
                currentStats=pd.read_csv("J:\zAutomatcherData\Part1Stats.csv", encoding ='utf-8')
              #  stats=open("J:\zAutomatcherData\Part1Stats.csv",'a')
                checks=len(df[df['Robot Suggestion'].str.contains('Check')])
                #stats.write(','.join((os.getenv('username'),str(self.bizID.get()),str(getBusName(\
              #     self.bizID.get())),str(self.IndustryType.get()),str(datetime.datetime.fromtimestamp(t1).strftime('%Y-%m-%d %H:%M:%S')),str(t1-t0),\
              #          str(df.shape[0]),str(checks))))
              #  stats.close()
                currentStats.loc[len(currentStats)]=[os.getenv('username'),self.bid,getBusName(\
                    self.bid),self.IndustryType.get(),self.ReportType.get(), self.dataInput.get(), datetime.datetime.fromtimestamp(t1).strftime('%Y-%m-%d %H:%M:%S'),t1-t0,\
                        df.shape[0],checks]
                
                currentStats.to_csv("J:\zAutomatcherData\Part1Stats.csv", encoding='utf-8',index=False)
        except:
            pass

        self.completeWindow=Toplevel()  
        self.completeWindow.protocol("WM_DELETE_WINDOW", self._delete_window) 
        self.completeMsg=Label(self.completeWindow,text="time to complete: "+str(t1-t0)+"  "+msg).grid(row=1,column=1)
        self.DoneButton=Button(self.completeWindow,text="Exit", command= lambda:[root.destroy()]).grid(row=2,column=1)
    def _delete_window(self):
        
        try:
            root.destroy()
        except:
            pass
        


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
try:
    app.mainloop()
except:
    pass