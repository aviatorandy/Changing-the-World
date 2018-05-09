# -*- coding: utf-8 -*-
"""
Created on Wed Oct 25 14:23:11 2017

@author: achang, plasser, storres
"""

import os
import re
from math import radians, cos, sin, asin, sqrt, isnan
from Tkinter import *
from ttk import *
import tkFileDialog
import Tkinter
import subprocess
import time
#from time import strftime
import datetime
from datetime import date
import sys
import warnings
import itertools
import collections
import csv
import operator
#from collections import OrderedDict
import ast
import win32com.client
import pandas as pd
import numpy as np
from fuzzywuzzy import fuzz
import MySQLdb
#import math
#import json



warnings.simplefilter("ignore", UserWarning)
warnings.filterwarnings('ignore', category=MySQLdb.Warning)


class NameDenormalizer(object):
    def __init__(self, filename='names.csv'):
        filename = filename or 'names.csv'
        lookup = collections.defaultdict(list)
        with open(filename) as f:
            reader = csv.reader(f)
            for line in reader:
                matches = set(line)
                for match in matches:
                    lookup[match].append(matches)
        self.lookup = lookup

    def __getitem__(self, name):
        name = name.lower()
        if name not in self.lookup:
            raise KeyError(name)
        names = reduce(operator.or_, self.lookup[name])
        if name in names:
            names.remove(name)
        return names

    def get(self, name, default=None):
        try:
            return self[name]
        except KeyError:
            return default


class NameDenormalizerWithOriginal(object):
    def __init__(self, filename='names.csv'):
        filename = filename or 'names.csv'
        lookup = collections.defaultdict(list)
        with open(filename) as f:
            reader = csv.reader(f)
            for line in reader:
                matches = set(line)
                for match in matches:
                    lookup[match].append(matches)
        self.lookup = lookup

    def __getitem__(self, name):
        name = name.lower()
        if name not in self.lookup:
            raise KeyError(name)
        names = reduce(operator.or_, self.lookup[name])
       # if name in names:
          #  names.remove(name)
        return names

    def get(self, name, default=None):
        try:
            return self[name]
        except KeyError:
            return default


#This function cleans the names
def cleanName(name):
    try:
        name = name.strip().lower()
        name = name.replace("&", " and ").replace("professional corporation", "")
        name = name.replace(" pc ", " ").replace(" lp ", " ").replace(" llc ", " ")
        name = name.replace("incorporated", "").replace(" inc.", " ").replace(" inc ", " ")
        name = re.sub('[^A-Za-z0-9\s]+', ' ', name)
        name = re.sub('\s+', ' ', name)
    except AttributeError:
        name = ""
    return name

#This function cleans the addresses
def cleanAddress(address):
    try:
        address = address.strip().lower()
        address = address.replace("&", " and ").replace(".", "")
        address = address.replace(" avenue ", " ave ").replace(" boulevard ", " blvd ").replace(" bypass ", " byp ")
        address = address.replace(" circle ", " cir ").replace(" drive ", " dr ").replace(" expressway ", " expy ")
        address = address.replace(" highway ", " hwy ").replace(" parkway ", " pkwy ").replace(" road ", " rd ")
        address = address.replace(" street ", " st ").replace(" turnpike ", " tpke ")\
    .replace(" road", " rd")\
    .replace(" street", " st")\
    .replace(" place", " pl")\
    .replace(" drive", " dr")\
    .replace(" i-h ", " i")\
    .replace(" ih ", " i")\
    .replace(" interstate ", " i")\
    .replace(" boulevard", " blvd")\
    .replace(" parkway", " pkwy")\
    .replace(" lane", " ln")\
    .replace(" turnpike", " tpke")\
    .replace(" highway", " hwy")\
    .replace(" route ", " rt ")\
    .replace(" rte ", " rt ")\
    .replace(" avenue", " ave")\
    .replace(" freeway", " fwy")\
    .replace(" court", " ct")\
    .replace(" expressway", " expy")\
    .replace(" mount ", " mt ")\
    .replace(" trail", " trl")\
    .replace("lyndon b johnson", "lbj")\
    .replace(" center", " ctr")\
    .replace(" centre", " ctr")\
    .replace(" circle", " cir")\
    .replace(" bypass", " byp")\
    .replace(" pike", " pk")\
    .replace(" saint", " st")\
    .replace(" terrace", " ter")\
    .replace(" point", " pt")\
    .replace(" station", " sta")\
    .replace(" causeway", " cswy")\
    .replace(" crossing", " xing")\
    .replace(" gateway", " gtwy")\
    .replace(" creek", " ck")\
    .replace(" village", " vlg")\
    .replace(" first", " 1st")\
    .replace(" second", " 2nd")\
    .replace(" third", " 3rd")\
    .replace(" fourth", " 4th")\
    .replace(" fifth", " 5th")\
    .replace(" sixth", " 6th")\
    .replace(" seventh", " 7th")\
    .replace(" eighth", " 8th")\
    .replace(" ninth", " 9th")\
    .replace(" tenth", " 10th")\
    .replace("one", "1")\
    .replace("two", "2")\
    .replace("three", "3")\
    .replace("four", "4")\
    .replace("five", "5")\
    .replace("six", "6")\
    .replace("seven", "7")\
    .replace("eight", "8")\
    .replace("nine", "9")\
    .replace("state road", "sr")\
    .replace("county road", "cr")\
    .replace(" hiway ", " hwy ")\
    .replace("farm to market", "fm")\
    .replace(" northwest", " nw")\
    .replace(" northeast", " ne")\
    .replace(" southwest", " sw")\
    .replace(" southeast", " se")\
    .replace(" s w ", " sw ")\
    .replace(" n w ", " nw ")\
    .replace(" s e ", " se ")\
    .replace(" n e ", " ne ")\
    .replace(" east ", " e ")\
    .replace(" west ", " w ")\
    .replace(" north ", " n ")\
    .replace(" south ", " s ")\
    .replace(" u s ", " us ")\
    .replace("straße ", "str ")\
    .replace("strasse ", "str ")\
    .replace("str. ", "str ")\
    .replace(" suite ", " ste ")\
    .replace("suite ", "ste ")\
    .replace("ste #", "ste ")\
    .replace("#", "ste ")\
    .replace("building", "bldg")\
    .replace("floor", "flr")\
    .replace(" unit", " ste")\
    .replace("unit ", "ste ")\
    .replace("=", "'")\
    .replace("  ", " ")\
    .replace("  ", " ")\
    .replace("  ", " ")\
    .replace("  ", " ")\
    .replace("  ", " ")\
    .replace("  ", " ")

        address = re.sub('[^A-Za-z0-9\s]+', ' ', address)
        address = re.sub('\s+', ' ', address)
    except AttributeError:
        address = ""
    return address



#This function cleans the city
def cleanCity(city):
    try:
        city = city.strip().lower()
        city = city.replace("&", " and ")\
    .replace("saint ", "st ")\
    .replace("fort ", "ft ")\
    .replace(".", "")\
    .replace("-", " ")\
    .replace("north ", "n ")\
    .replace("south ", "s ")\
    .replace("east ", "e ")\
    .replace("west ", "w ")\
    .replace("mount ", "mt ")\
    .replace("spring", "spg")\
    .replace("height", "ht")\

        city = re.sub('[^A-Za-z0-9\s]+', ' ', city)
        city = re.sub('\s+', ' ', city)
    except AttributeError:
        city = ""
    return city

#==============================================================================
#      #If name is blank, fills in last part of URL
#     for index,row in df.iterrows():
#
#         if row['Listing Name'] == None and row['Listing URL']!=None:
#             df.loc[index,'Cleaned Listing Name'] = cleanName(row['Listing URL'].split('/')[-1])
#             df.loc[index,'No Name'] = 'URL for name'
#         elif row['Listing Name'] == None:
#             df.loc[index,'No Name'] = 'No Name'
#
#
#==============================================================================

#This function compares the names in the file
def compareName(df, IndustryType, bid):
    
    df['Cleaned Location Name'] = df['Location Name'].apply(cleanName)
    df['Cleaned Listing Name'] = df['Listing Name'].apply(cleanName)
    df['No Name'] = df['Listing Name'].isnull()


    #Removes City name rom Listing name, if present
    df['Cleaned Listing Name'] = df.apply(lambda x: x['Cleaned Listing Name'].replace\
    (cleanName(x['Location City']), '') if cleanName(x['Location City']) in x['Cleaned Listing Name'] else\
    x['Cleaned Listing Name'], axis=1)


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

    if app.WordsIgnore:    
        [x.lower() for x in app.WordsIgnore]
        df['Must Ignore'] = df['Cleaned Listing Name'].apply(lambda x: any(item in x for item in app.WordsIgnore))
         
        for words in app.WordsIgnore:
            df['Cleaned Location Name'] = df['Cleaned Location Name'].apply(lambda x: x.replace(words, "").strip())          
            df['Cleaned Listing Name'] = df['Cleaned Listing Name'].apply(lambda x: x.replace(words, "").strip())

    #Should this be here, or only for some industries?
    df['ListPeopleNames'] = np.empty((len(df), 0)).tolist()
    df['LocPeopleNames'] = np.empty((len(df), 0)).tolist()
    df['ListPeopleNamesNotInLoc'] = np.empty((len(df), 0)).tolist()

    for index, row in df.iterrows():

        for word in row['Cleaned Listing Name'].split():
            otherListNames = firstNames.get(word.rstrip())
            if otherListNames:
                for name in otherListNames:
                    if name != '':
                        df.loc[index, 'ListPeopleNames'].append(name)
            otherListNames = None
        for word in row['Cleaned Location Name'].split():
            otherLocNames = firstNames.get(word.rstrip())
            if otherLocNames:
                for name in otherLocNames:
                    if name != '':
                        df.loc[index, 'LocPeopleNames'].append(name)
            otherLocNames = None

        df.at[index, 'ListPeopleNamesNotInLoc'] = list(set(df.loc[index, 'ListPeopleNames']).difference(set(df.loc[index, 'LocPeopleNames'])))

    df['ExtraPeopleNamesInListing'] = df['ListPeopleNamesNotInLoc'].apply(lambda x: len(x) > 0)


#start of comparisons, broken out by industry
    #Normal Industry
    if IndustryType == "0":
        print "Normal Naming"
        if businessNameMatch == 1:

            df['businessPartial'] = 0
            df['businessTokenSet'] = 0
            df['businessTokenSort'] = 0

            #compares business names vs listing names on different comparison methods. Takes highest score
            for bName in app.WordsAlt:
                df['businessPartial'] = df.apply(lambda row: \
                    max(row['businessPartial'], fuzz.partial_ratio(bName, row['Cleaned Listing Name'])), axis=1)
                df['businessTokenSet'] = df.apply(lambda row: \
                    max(row['businessTokenSet'], fuzz.token_set_ratio(bName, row['Cleaned Listing Name'])), axis=1)
                df['businessTokenSort'] = df.apply(lambda row:\
                    max(row['businessTokenSort'], fuzz.token_sort_ratio(bName, row['Cleaned Listing Name'])), axis=1)
                
            #Compares location name to listing name on different methods. Takes highest score
            df['Token Set'] = df.apply(lambda row: fuzz.token_set_ratio\
                    (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)
        
            df['Partial Score'] = df.apply(lambda row: fuzz.partial_ratio\
                (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)
        
            df['Token sort'] = df.apply(lambda row: fuzz.token_sort_ratio\
                (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)
        
            df['Name Score'] = df[['businessPartial', 'businessTokenSet', 'businessTokenSort', "Token Set", "Partial Score", "Token sort"]].max(axis=1)
        
            if app.WordsExclude:
                df['Words Exclude'] = df['Cleaned Listing Name'].apply(lambda x: any(item in x for item in app.WordsExclude))
            else:
                df['Words Exclude'] = ""

            if app.WordsAlt:
                df['Words Alt'] = df['Cleaned Listing Name'].apply(lambda x: any(item in x for item in app.WordsAlt))
            else:
                df['Words Alt'] = ""

            if app.WordsMust:
                df['Words Must'] = df['Cleaned Listing Name'].apply(lambda x: any(item in x for item in app.WordsMust))
            else:
                df['Words Must'] = ""


            #Removes extra columns
            df = df.drop(['businessPartial', 'businessTokenSet', 'businessTokenSort', 'Token Set', 'Partial Score', 'Token sort'], axis=1)

        else:

        #Compares location name to listing name on different methods. Takes highest score
            df['Token Set'] = df.apply(lambda row: fuzz.token_set_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Partial Score'] = df.apply(lambda row: fuzz.partial_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Token sort'] = df.apply(lambda row: fuzz.token_sort_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Name Score Mean'] = df[["Token Set", "Partial Score", "Token sort"]].mean(axis=1)

            df['Name Score'] = df[["Token Set", "Partial Score", "Token sort"]].max(axis=1)
            
            if app.WordsExclude:
                df['Words Exclude'] = df['Cleaned Listing Name'].apply(lambda x: any(item in x for item in app.WordsExclude))
            else:
                df['Words Exclude'] = ""

            if app.WordsAlt:
                df['Words Alt'] = df['Cleaned Listing Name'].apply(lambda x: any(item in x for item in app.WordsAlt))
            else:
                df['Words Alt'] = ""

            if app.WordsMust:
                df['Words Must'] = df['Cleaned Listing Name'].apply(lambda x: any(item in x for item in app.WordsMust))
            else:
                df['Words Must'] = ""
            
            df = df.drop(["Token Set", "Partial Score", "Token sort"], axis=1)

    #Auto Name Matching
    elif IndustryType == "1":
        print "Auto Naming"

        CarCheck = ["svc", "service", "parts", "part", "body", "collision", \
                    "clln", "body shop", "used", "usado", "pre-owned"]

        #Match to Sales and Services and New and Used
        CarBrands = ['gm', 'general motors', 'gmc', 'buick', 'cadillac', 'chevrolet', 'chevy', 'pontiac', \
        'oldsmobile', 'olds mobile', 'pntc', 'geo', 'hummer', 'saturn', \
        'corvette', 'vette', 'corvet', 'acura', 'alfa', 'aston martin', \
        'audi', 'enterprise rent', 'hertz', 'bentley', 'bmw', 'bugatti', 'chrysler', \
        'dodge', 'ferrari', 'fiat', 'ford', 'honda', 'hyundai', 'infiniti', 'isuzu', \
        'isuzu', 'jaguar', 'jeep', 'kia', 'lamborghini', 'lexus', 'lincoln', 'lotus', \
        'maserati', 'maybach', 'mazda', 'mercedes', 'mini cooper', 'mitsubishi', 'nissan', \
        'plymouth', 'porsche', 'rover', 'royce', 'saab', 'saturn', 'scion', 'subaru', \
        'suzuki', 'suzuki', 'toyota', 'volkswagen', 'vw', 'volvo']

        #If anything in CarBrands is not the Brand it is bad

        #This identifies the actual car brand name within the string
        df['Brand Name'] = df['Cleaned Location Name'].apply(lambda name: next((brand for brand in CarBrands if brand in name.split()), None))

        df['Removed Brand'] = df.apply(lambda x: x['Cleaned Listing Name'] if x['Brand Name'] == None\
                else x['Cleaned Listing Name'].replace(x['Brand Name'], "") if x['Brand Name'] in x['Cleaned Listing Name'] \
                                    else "", axis=1)

        df['Other Brand Check'] = df.apply(lambda x: False if x['Brand Name'] is None \
                                                else True if any(brand in x['Removed Brand'] for brand in CarBrands)\
                                                        else False, axis=1)

        #This puts certain listings into Checks based off of commonly excluded Auto words
        df['Car Excl Word Check'] = df['Cleaned Listing Name'].apply(lambda x: any(item in x for item in CarCheck))


        df['Name Score'] = df.apply(lambda row: \
         fuzz.token_set_ratio(row['Cleaned Location Name'], row['Cleaned Listing Name']) if pd.isnull(row['Brand Name']) else\
            fuzz.token_set_ratio(row['Brand Name'], row['Cleaned Listing Name']), axis=1)
        return


    #Industry Hotel
    elif IndustryType == "2":
        print "Hotel Naming"
        BadHotel = ["optimeyes", "bakery", "grill", "bar", "starbucks", "electric", "wedding", "gym",\
                     "pool", "restaurant", "bistro", "academy", "cafe", "salon",\
                     "5ten20", "lab", "rental", "car", "body", "fitness", "swim", "hertz", "steak",\
                     "sip", "zone", "alarm", "limestone", "catering", "room", "âme", "massage", "trivium",\
                     "broughton's", "broscheks", "vmug", "parking", "kenvigs", "martinis", "martini's"\
                     "beauty", "formaggio", "gallery", "motors",\
                     "sports", "formaggiosacramento", "rgs", "brasserie",\
                     "office", "vaso", "oceana", "vmware"\
                     "trivium", "fyve", "steakhouse", "ame", "wellness", "pay", "presentation"\
                     "presentations", "presentation", "visual",\
                     "tent", "eno", "copper", "coffee", "leisure", "charter", "me", "ticketmaster",\
                     "swampers", "journeys", "friend", "orchards", "mandara", "camp", "broughtons"]

        HotelBrands = ["ac hotel", "aloft", "america's best", "moxy", "tribute"\
        "americas best value", "ascend", "autograph", "baymont", "best western",\
        "cambria", "canadas best value", "country inn", "park inn", "cassa", "candlewood", "clarion", "conrad", "comfort inn", "elyton hotel"\
        "comfort suites", "delta", "country hearth", "courtyard", "crowne plaza", "curio",\
        "casa monica", "days inn", "grand residence", "doubletree", "econo lodge", "econolodge", "edition", "element",\
        "embassy suites", "even", "fairfield inn", "four points", "garden inn", "gaylord", "renaissance"\
        "hampton inn", "hilton garden inn", "holiday inn", "homewood suites", "howard johnson", "hyatt",\
        "indigo", "intercontinental", "jameson", "jw marriott", "la quinta", "le meridien",\
        "le méridien", "lexington", "luxury collection", "protea", "mainstay", "marriott executive", \
        "microtel", "motel 6", "palace inn", "rezidor", "radisson" "premier inn", "quality inn",\
        "quality suites", "ramada", "red roof", "renaissance", "residence inn", "bulgari", "ritz carlton", "protea"\
        "rodeway", "sheraton", "signature Inn", "sleep inn", "springhill", "st regis",\
        "st. regis", "starwood", "staybridge", "studio 6", "super 8", "towneplace",\
        "value hotel", "value inn", "w hotel", "westin", "wingate", "wyndham", "marriott"]

        #If anything in Bad Hotel is in Listing name, is not a hotel

        #This identifies the actual car brand name within the string

        df['Not Hotel'] = df['Cleaned Listing Name'].apply(lambda x: 1 if any(item in x.split() for item in BadHotel) else 0)
        df['Hotel Brand Name'] = df['Cleaned Location Name'].apply(lambda name: next((brand for brand in HotelBrands if brand in name), None))

        #If hotel brand exists in listing name,is a hotel.
        df['Other Hotel Match'] = df['Cleaned Listing Name'].apply(lambda x: 1 if any(item in x for item in HotelBrands) else 0)

        df['Name Score'] = df.apply(lambda row: \
         fuzz.token_set_ratio(row['Cleaned Location Name'], row['Cleaned Listing Name']) if pd.isnull(row['Hotel Brand Name']) else\
            fuzz.token_set_ratio(row['Hotel Brand Name'], row['Cleaned Listing Name']), axis=1)

    #Industry Healthcare Professional matching
    elif IndustryType == "3":
        print "HC Prof Naming"

#        nickNameExcelFile=pd.ExcelFile("~\Documents\Changing-the-World\nicknames.xlsx", keep_default_na = False)
#        nicknameList = nickNameExcelFile.parse('nicknames')
#        nicknameList = nicknameList.fillna("yext123")
#        nicknameList = nicknameList.applymap(lambda x : x.lower())
#        nicknameList = nicknameList.values.tolist()#

    #Get name score from Provider name v listing name
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
        df['altListNames'] = np.empty((len(df), 0)).tolist()
        df['altLocNames'] = np.empty((len(df), 0)).tolist()

        for index, row in df.iterrows():
            row['altListNames'].append(row['Cleaned Listing Name'])
            row['altLocNames'].append(row['Cleaned Location Name'])

            for word in row['Cleaned Listing Name'].split():
                otherListNames = nickNames.get(word)
                if otherListNames:
                    for name in otherListNames:
                        df.loc[index, 'altListNames'].append(row['Cleaned Listing Name'].replace(word, name))
                otherListNames = None
            for word in row['Cleaned Location Name'].split():
                otherLocNames = nickNames.get(word)
                if otherLocNames:
                    for name in otherLocNames:
                        df.loc[index, 'altLocNames'].append(row['Cleaned Location Name'].replace(word, name))
                otherLocNames = None

        if businessNameMatch == 1:
            df['businessPartial'] = 0
            df['businessTokenSet'] = 0
            df['businessTokenSort'] = 0

        #compares business names vs listing names on different comparison methods. Takes highest score
            for bName in app.wordsAlt:
                df['businessPartial'] = df.apply(lambda row: \
                    max(row['businessPartial'], fuzz.partial_ratio(bName, row['Cleaned Listing Name'])), axis=1)
                df['businessTokenSet'] = df.apply(lambda row: \
                    max(row['businessTokenSet'], fuzz.token_set_ratio(bName, row['Cleaned Listing Name'])), axis=1)
                df['businessTokenSort'] = df.apply(lambda row:\
                    max(row['businessTokenSort'], fuzz.token_sort_ratio(bName, row['Cleaned Listing Name'])), axis=1)

            df['Token Set'] = df.apply(lambda row: fuzz.token_set_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Partial Score'] = df.apply(lambda row: fuzz.partial_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Token sort'] = df.apply(lambda row: fuzz.token_sort_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Alt Token Set'] = ""
            df['Alt Partial Score'] = ""
            df['Alt Token Sort'] = ""
            for index, row in df.iterrows():
                for altLocName in row['altLocNames']:
                    for altListName in row['altListNames']:
                        df.loc[index, 'Alt Token Set'] = max(df.loc[index, 'Alt Token Set'], fuzz.token_set_ratio(altLocName, altListName))
                        df.loc[index, 'Alt Partial Score'] = max(df.loc[index, 'Alt Partial Score'], fuzz.token_set_ratio(altLocName, altListName))
                        df.loc[index, 'Alt Token Sort'] = max(df.loc[index, 'Alt Token Sort'], fuzz.token_set_ratio(altLocName, altListName))

            df['Name Score'] = df[['businessPartial', 'businessTokenSet', 'businessTokenSort',\
            "Token Set", "Partial Score", "Token sort", 'Alt Token Set', 'Alt Partial Score', 'Alt Token Sort']].max(axis=1)


            df = df.drop(['businessPartial', 'businessTokenSet', 'businessTokenSort', "Token Set", "Partial Score", "Token sort"], axis=1)
        else:
    #Compares location name to listing name on different methods. Takes highest score

            df['Token Set'] = df.apply(lambda row: fuzz.token_set_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Partial Score'] = df.apply(lambda row: fuzz.partial_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Token sort'] = df.apply(lambda row: fuzz.token_sort_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Token Set'] = df.apply(lambda row: fuzz.token_set_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Partial Score'] = df.apply(lambda row: fuzz.partial_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Token sort'] = df.apply(lambda row: fuzz.token_sort_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Alt Token Set'] = 0
            df['Alt Partial Score'] = 0
            df['Alt Token Sort'] = 0
            for index, row in df.iterrows():
                for altLocName in row['altLocNames']:
                    for altListName in row['altListNames']:
                        df.loc[index, 'Alt Token Set'] = max(df.loc[index, 'Alt Token Set'], fuzz.token_set_ratio(altLocName, altListName))
                        df.loc[index, 'Alt Partial Score'] = max(df.loc[index, 'Alt Partial Score'], fuzz.token_set_ratio(altLocName, altListName))
                        df.loc[index, 'Alt Token Sort'] = max(df.loc[index, 'Alt Token Sort'], fuzz.token_set_ratio(altLocName, altListName))

            #returns Max of Business Match or Location Name Match
            df['Name Score Mean'] = df[["Token Set", "Partial Score", "Token sort", 'Alt Token Set', 'Alt Partial Score', 'Alt Token Sort']].mean(axis=1)

            df['Name Score'] = df[["Token Set", "Partial Score", "Token sort", 'Alt Token Set', 'Alt Partial Score', 'Alt Token Sort']].max(axis=1)

            df = df.drop(["Token Set", "Partial Score", "Token sort"], axis=1)

    #Industry International/Healthcare Facility
    else:
        print "Other Naming"
        print businessNameMatch 
        if businessNameMatch == 1:
            df['businessPartial'] = 0
            df['businessTokenSet'] = 0
            df['businessTokenSort'] = 0
        #Compares location name to listing name on different methods. Takes highest score

            for bName in app.wordsAlt:
                df['businessPartial'] = df.apply(lambda row: \
                    max(row['businessPartial'], fuzz.partial_ratio(bName, row['Cleaned Listing Name'])), axis=1)
                df['businessTokenSet'] = df.apply(lambda row: \
                    max(row['businessTokenSet'], fuzz.token_set_ratio(bName, row['Cleaned Listing Name'])), axis=1)
                df['businessTokenSort'] = df.apply(lambda row:\
                    max(row['businessTokenSort'], fuzz.token_sort_ratio(bName, row['Cleaned Listing Name'])), axis=1)

        #Compares location name to listing name on different methods. Takes highest score

            df['Token Set'] = df.apply(lambda row: fuzz.token_set_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Partial Score'] = df.apply(lambda row: fuzz.partial_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Token sort'] = df.apply(lambda row: fuzz.token_sort_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Name Score'] = df[['businessPartial', 'businessTokenSet', 'businessTokenSort', "Token Set", "Partial Score", "Token sort"]].max(axis=1)

            if app.WordsExclude:
                df['Words Exclude'] = df['Cleaned Listing Name'].apply(lambda x: any(item in x for item in app.WordsExclude))
            else:
                df['Words Exclude'] = ""

            if app.WordsAlt:
                df['Words Alt'] = df['Cleaned Listing Name'].apply(lambda x: any(item in x for item in app.WordsAlt))
            else:
                df['Words Alt'] = ""

            if app.WordsMust:
                df['Words Must'] = df['Cleaned Listing Name'].apply(lambda x: any(item in x for item in app.WordsMust))
            else:
                df['Words Must'] = ""

            df = df.drop(['businessPartial', 'businessTokenSet', 'businessTokenSort', "Token Set", "Partial Score", "Token sort"], axis=1)
        

        else:
        #Compares  location name to listing name on different methods. Takes highest score

            df['Token Set'] = df.apply(lambda row: fuzz.token_set_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Partial Score'] = df.apply(lambda row: fuzz.partial_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            df['Token sort'] = df.apply(lambda row: fuzz.token_sort_ratio\
            (row['Cleaned Location Name'], row['Cleaned Listing Name']), axis=1)

            #returns Max of Business Match or Location Name Match
            df['Name Score Mean'] = df[["Token Set", "Partial Score", "Token sort"]].mean(axis=1)

            df['Name Score'] = df[["Token Set", "Partial Score", "Token sort"]].max(axis=1)

            if app.WordsExclude:
                df['Words Exclude'] = df['Cleaned Listing Name'].apply(lambda x: any(item in x for item in app.WordsExclude))
            else:
                df['Words Exclude'] = ""

            if app.WordsAlt:
                df['Words Alt'] = df['Cleaned Listing Name'].apply(lambda x: any(item in x for item in app.WordsAlt))
            else:
                df['Words Alt'] = ""

            if app.WordsMust:
                df['Words Must'] = df['Cleaned Listing Name'].apply(lambda x: any(item in x for item in app.WordsMust))
            else:
                df['Words Must'] = ""
                
            df = df.drop(["Token Set", "Partial Score", "Token sort"], axis=1)
            
#If claimed, give score
def compareStatus(df):
    df['Claimed Score'] = df.apply(lambda x: 100 if x['Advertiser/Claimed'] == "Claimed" else 0, axis=1)

#This function compares the phones in the file
def comparePhone(df):
    #Handles different variable types of phones, cleans up so can be used.
    df['Location Phone'] = df.apply(lambda x: '%.12g' % x['Location Phone'] if isinstance(x['Location Phone'], float) else str(x['Location Phone']), axis=1)
    df['Location Local Phone'] = df.apply(lambda x: '%.12g' % x['Location Local Phone'] if isinstance(x['Location Local Phone'], float) else str(x['Location Local Phone']), axis=1)
    df['Listing Phone'] = df.apply(lambda x: '%.12g' % x['Listing Phone'] if isinstance(x['Listing Phone'], float) else str(x['Listing Phone']), axis=1)


    #Finds if phones are equal
    try:
        df['Phone Match'] = df.apply(lambda x: True \
            if (x['Location Phone'] == x['Listing Phone'] and \
                  (pd.isnull(x['Location Phone']) or  pd.isnull(x['Listing Phone']) or x['Listing Phone'] == None or x['Listing Phone'] == "nan") == False)\
            else (True if (x['Location Local Phone'] == x['Listing Phone'] \
                           and pd.isnull(x['Location Local Phone']) == False \
                        and x['Listing Phone'] != "nan" and  pd.isnull(x['Listing Phone']) == False and x['Listing Phone'] != None) else False), axis=1)

        df['Phone Score'] = df.apply(lambda x: 100 if x['Phone Match'] else 0, axis=1)
    except:
        df['Phone Match'] = 'x'


#This function compares the addresses in the file
def compareAddress(df, IndustryType):
    df['No Address'] = df['Listing Address'].isnull()

    #International
    if IndustryType == '6':
        #Combine Address 1, Address 2
        df['Cleaned Input Address'] = df['Location Address'].apply(cleanAddress)\
                                    +' '+df['Location Address 2'].apply(cleanAddress)
        df['Cleaned Listing Address'] = df['Listing Address'].apply(cleanAddress)\
                                    +' '+df['Listing Address 2'].apply(cleanAddress)
        #removes extra space where necessary
        df['Cleaned Input Address'] = df['Cleaned Input Address'].apply(cleanAddress)
        df['Cleaned Listing Address'] = df['Cleaned Listing Address'].apply(cleanAddress)


        #Compares addresses on different comparison methods. Takes max.
        df['ar'] = df.apply(lambda row: fuzz.ratio(row['Cleaned Input Address'], row['Cleaned Listing Address']), axis=1)
        df['aSortr'] = df.apply(lambda row: fuzz.token_sort_ratio(row['Cleaned Input Address'], row['Cleaned Listing Address']), axis=1)
        df['aSetr'] = df.apply(lambda row: fuzz.token_set_ratio(row['Cleaned Input Address'], row['Cleaned Listing Address']), axis=1)

        df['Address Score'] = df[['ar', 'aSortr', 'aSetr']].max(axis=1)
      #Deletes extra columns
        df = df.drop(['ar', 'aSortr', 'aSetr'], axis=1)

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

        df['Address Score'] = df.apply(lambda row: fuzz.token_set_ratio\
            (row['Cleaned Input Address'], row['Cleaned Listing Address']), axis=1)


#This function compares the cities in the file
def compareCity(df):
    df['Cleaned Input City'] = df['Location City'].apply(cleanCity)
    df['Cleaned Listing City'] = df['Listing City'].apply(cleanCity)

    #Compares cities on different methods. Takes mean
    df['csr'] = df.apply(lambda row: fuzz.ratio(row['Cleaned Input City'], row['Cleaned Listing City']), axis=1)
    df['ctpr'] = df.apply(lambda row: fuzz.partial_ratio(row['Cleaned Input City'], row['Cleaned Listing City']), axis=1)
    df['City Score'] = df[['csr', 'ctpr']].mean(axis=1)

#This function compares the Zips in the file
def compareZip(df):
    df['Zip Match'] = df.apply(lambda x: True if x['Location Zip'] == x['Listing Zip']\
                                 else False, axis=1)

#This function compares the NPIs in the file
def compareNPI(df):
    df['NPI Match'] = df.apply(lambda x: True if x['Location NPI'] == x['Listing NPI'] else False, axis=1)

#copied from old template. Calculates metric distance between geocodes
def calculateDistance(row):
    try:
        locationLat = float(row['Location Latitude'])
    except:
        locationLat = 0.0
    try:
        locationLong = float(row['Location Longitude'])
    except:
        locationLong = 0.0
    try:
        listingLat = float(row['Listing Latitude'])
    except:
        listingLat = 90.0
    try:
        listingLong = float(row['Listing Longitude'])
    except:
        listingLong = 90.0


    if locationLat == 0.0 or listingLat == 0.0:
        return "n/a"

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


    #When running facilities, checks for specialty match, or excludes doctor words
def calculateDoctorMatch(df):

    commonDoctorWords = ['md', 'pa', 'dr', 'do', 'np', 'phys', 'lpn', 'rn', 'dds', 'cnm', 'mph', 'phd', 'od', 'gp', 'dpm', 'gift', 'cafe', 'café']
#Reads doctor specialty doc
    excelFile = pd.ExcelFile("~\Documents\Changing-the-World\SpecialtyDoctorMatching.xlsx", keep_default_na=False)
    doctorSpecialty = excelFile.parse('Specialty')
    doctorSpecialty = doctorSpecialty.fillna("yext123")
    doctorSpecialty = doctorSpecialty.applymap(lambda x: x.lower())
    doctorSpecialty = doctorSpecialty.values.tolist()

    specialties = []

    df['Bad Facility'] = df['Cleaned Listing Name'].apply(lambda x: 1 if any(word in x.split() for word in commonDoctorWords) else 0)

    for index, row in df.iterrows():

        locationNameSpecialties = None
        listingNameSpecialties = None
        locationName = (" " + str(row['Cleaned Location Name']) + " ").lower()

        listingName = (" " + str(row['Cleaned Listing Name']) + " ").lower()
    
        
   #If matches to physician titles, antimatch
        if row['Bad Facility'] == 1:
            specialties.append("No Match - Bad Facility")
        else:
            locationNameSpecialties = set([tuple(group) for group in doctorSpecialty for specialty in group if specialty in locationName])
            listingNameSpecialties = set([tuple(group) for group in doctorSpecialty for specialty in group if specialty in listingName])

#Compares specialty, with synonyms from location to listing. if same, match, if different, antimatch
            if locationNameSpecialties:
                if not listingNameSpecialties:
                    specialties.append("Check-Generic")
                elif [locationNameSpecialties & listingNameSpecialties][0]:
                    specialties.append("Match-Specialty")
                else:
                    specialties.append("No Match-Specialty")
            else:
                if listingNameSpecialties:
                    specialties.append("No Match-Specialty")
                elif not listingNameSpecialties:
                    specialties.append("Check-Generic")

           #this else seems out of place, but not sure where this goes.
                else:
                    specialties.append("Check")

    df['Specialty Match'] = specialties

    df['Biz Name Match'] = df['Cleaned Listing Name'].apply(lambda x: True if any(item in x for item in businessNames) else False)

#Main function to run through comparisons
def compareData(df, IndustryType, bid):

    print 'Comparing names...'
    compareName(df, IndustryType, bid)
    print 'Comparing phones...'
    comparePhone(df)
    #compares claimed pages
    compareStatus(df)
    print 'Comparing zips...'
    compareZip(df)
    if IndustryType == '3':
        print 'Comparing NPIs...'
        compareNPI(df)

    if IndustryType == '4':
        print 'Checking specialties...'
        calculateDoctorMatch(df)
    print 'Comparing addresses...'
    compareAddress(df, IndustryType)
    df['Distance (M)'] = df.apply(lambda row: calculateDistance(row), axis=1)
    print 'Suggesting matches...'
    calculateTotalScore(df)
    suggestmatch(df, IndustryType)

#This function provides a suggested match based on certain name/address thresholds
def suggestmatch(df, IndustryType):
    robotmatch = []
    global reportType

    #if FB then create new column

    #Creates new columns that are easier to read in code, and for users when outputed - Mostly boolean variables
    df['Address Match'] = df.apply(lambda x: True if x['Address Score'] >= 70 else False, axis=1)
    df['Name Match'] = df.apply(lambda x: 1 if x['Name Score'] >= 70 else (2 if 70 > x['Name Score'] >= 60 else 0), axis=1)
    df['Geocode Match'] = df.apply(lambda x: True if x['Distance (M)'] <= 200 else False, axis=1)
    df['Phone or Address Match'] = df.apply(lambda x: x['Phone Match'] or x['Address Match'], axis=1)
    if 'Sync URL' in df.columns:
        df['FB Error'] = df.apply(lambda x: x['Listing URL'] == x['Sync URL'] and x['Advertiser/Claimed'] == 'Claimed'\
        and x['Live Sync'] == 0 and x['Live Suppress'] == 0, axis= 1)
    else:
        df['FB Error'] == False

    liveSync = 'No Match - Live Sync'
    liveSuppress = 'No Match - Live Suppress'
    matchText = 'Match Suggested'
    noName = 'No Match - Name'
    noMustMatch = 'No Match - No Must Have'
    noNameExc = 'No Match - Excluded'
    noMatch = 'No Match'
    checkBrand = "Check - No Loc Brand"
    noAddress = 'No Match - Phone/Address'
    checkMissing = 'Check No Name/Address'
    checkMissingName = 'Check No Name'
    noMatchFBDupeError = 'No Match - FB Dupe Error'
    checkAuto = 'Check - Excl Auto Words'
    check = 'Check Name'
    noSpecialty = 'No Match - Specialty'
    checkSpecialty = 'Check Doctor/Specialty'
    #npimatch = 'Match Suggested - NPI'
    clusternpimatch = 'Match Suggested - Cluster NPI'
    clusternpimismatch = 'Check Name - Cluster Pub'
    diffBrand = 'No Match - Diff Brand'
    userMatch = 'User Match'
    noFBMatch = 'User Manual No Match'
    global bpgid
    if reportType != 2:
        bpgid = []
    #if external ID External ID is in the array then we would the following
    #if user match is there then we would say user match
    # if type facebook then

    df['External ID'] = df['External ID'].map(lambda x: x.lstrip('\''))


    #Normal Type
    if IndustryType == '0':
        #Applies Match rules based on new columns.
        print "Normal Matching"
        df['Robot Suggestion'] = df.apply(lambda x: liveSync if x['Live Sync'] == 1 \
            else liveSuppress if x['Live Suppress'] == 1 \
            else noMatchFBDupeError if x['FB Error']\
                else noNameExc if x['Words Exclude']\
                    else noMustMatch if x['Words Must'] == False \
                        else noFBMatch if x['External ID'] in bpgid \
                            else userMatch if x['User Match'] == 1 \
                                    else checkMissing if (x['No Name'])\
                                        else noName if x['Name Match'] == 0\
                                            else check if x['Name Match'] == 2 and (x['Address Match'] == True or x['Phone Match'] == True)\
                                                else matchText if x['Name Match'] == 1\
                                                    and x['Address Match'] == True or x['Phone Match'] == True\
                                                        else checkMissing if x['No Address']\
                                                            else noAddress if x['Address Match'] == False\
                                                                else noMatch, axis=1)
        df['Match \n1 = yes, 0 = no'] = ""

    #Auto Type
    elif IndustryType == '1':
        print "Auto Matching..."
        df['Robot Suggestion'] = df.apply(lambda x: liveSync if x['Live Sync'] == 1 \
            else liveSuppress if x['Live Suppress'] == 1 \
            else noMatchFBDupeError if x['FB Error']\
                else noFBMatch if x['External ID'] in bpgid \
                    else userMatch if x['User Match'] == 1 \
                        else checkMissing if (x['No Name'])\
                            else checkAuto if x['Car Excl Word Check']\
                                else checkBrand if x['Brand Name'] == None\
                                    else diffBrand if x['Other Brand Check'] == True\
                                        else noName if x['Name Match'] == 0 \
                                            else check if x['Name Match'] == 2 and (x['Address Match'] == True or x['Phone Match'] == True)\
                                                else matchText if x['Name Match'] == 1 \
                                                    and x['Address Match'] == True or x['Phone Match'] == True\
                                                        else checkMissing if x['No Address']\
                                                            else noAddress if x['Address Match'] == False\
                                                                else noMatch, axis=1)

        df['Name Match'] = df.apply(lambda x: True if x['Name Match'] == 1 \
                            else ('Check' if x['Name Match'] == 2 else False), axis=1)

        df['Match \n1 = yes, 0 = no'] = ""

    #Hotel Type
    #Need to Add Name Score taking the max of the brand
    elif IndustryType == '2':
        print "Hotel Matching..."
        df['Robot Suggestion'] = df.apply(lambda x: liveSync if x['Live Sync'] == 1 \
            else liveSuppress if x['Live Suppress'] == 1 \
                else noMatchFBDupeError if x['FB Error']\
                   else noFBMatch if x['External ID'] in bpgid \
                        else userMatch if x['User Match'] == 1 \
                            else checkMissing if (x['No Name'] or x['No Address'])\
                                else noName if x['Not Hotel'] == 1\
                                    else noAddress if x['Address Match'] == False\
                                        else noName if x['Name Match'] == 0 \
                                            else check if x['Name Match'] == 2 and (x['Address Match'] == True or x['Phone Match'] == True)\
                                                else matchText if x['Name Match'] == 1 and x['Address Match'] == True\
                                                    and x['Phone Match'] == True and x['Other Hotel Match'] == 1\
                                                        else matchText if x['Name Match'] == 1 \
                                                        and x['Address Match'] == True or x['Phone Match'] == True\
                                                                else noMatch, axis=1)

        df['Name Match'] = df.apply(lambda x: True if x['Name Match'] == 1 \
        else ('Check' if x['Name Match'] == 2 else False), axis=1)

        df['Match \n1 = yes, 0 = no'] = ""

    #Healthcare Professional
    elif IndustryType == '3':
        print "HC Prof Matching"
        clusterpubs = [737, 735, 736, 776, 741]

        df['Name Match'] = df.apply(lambda x: 1 if x['Name Score'] >= 76 else (2 if 76 > x['Name Score'] >= 66 else 0), axis=1)

        df['Robot Suggestion'] = df.apply(lambda x: liveSync if x['Live Sync'] == 1 \
            else liveSuppress if x['Live Suppress'] == 1 \
                else noMatchFBDupeError if x['FB Error']\
                    else noFBMatch if x['External ID'] in bpgid \
                        else userMatch if x['User Match'] == 1 \
                            else clusternpimatch if x['Publisher ID'] in clusterpubs and x['NPI Match'] == 1 \
                                else clusternpimismatch if x['NPI Match'] != 1 and x['Publisher ID'] in clusterpubs \
                                    else matchText if x['Name Match'] == 1 and (x['Phone Match'] or x['Address Match'] or x['Geocode Match']) \
                                        else (check if x['Name Match'] == 2 and (x['Phone Match'] or x['Address Match'] or x['Geocode Match']) \
                                            else (noName if x['Name Match'] == 0 \
                                                  else (noAddress if not x['Address Match'] else 'tbd'))), axis=1)

        df['Match \n1 = yes, 0 = no'] = ""

        return

    #Healthcare Facilities
    elif IndustryType == '4':
        print "HC Facility Matching"
        
        df['Name Match'] = df.apply(lambda x: 1 if x['Name Score'] >= 90 else (2 if 90 > x['Name Score'] >= 76 else 0), axis=1)

        df['Robot Suggestion'] = df.apply(lambda x: liveSync if x['Live Sync'] == 1 \
            else liveSuppress if x['Live Suppress'] == 1 \
                else noMatchFBDupeError if x['FB Error']\
                    else noNameExc if x['Words Exclude']\
                        else noMustMatch if x['Words Must'] == False \
                            else noFBMatch if x['External ID'] in bpgid \
                                else userMatch if x['User Match'] == 1 \
                                    else noSpecialty if x['Specialty Match'][0:2] == 'No' \
                                        else checkMissing if (x['No Name'] or x['No Address'])\
                                            else ((matchText if (x['Name Match'] == 1 and (x['Phone Match'] or x['Address Match'] or x['Geocode Match'])) or \
                                                  ((x['Biz Name Match'] == True and x['Specialty Match'][0:5] == 'Match') and (x['Phone Match'] or x['Address Match'] or x['Geocode Match'])) \
                                                  else (check if x['Name Match'] == 2 \
                                                  else (noName if x['Name Match'] == 0 \
                                                  else (noAddress if not x['Address Match'] else 'uh oh'))))\
                                                  if x['Specialty Match'][0:5] == 'Match' else \
                                                  ('Check Name and Specialty' if x['Name Match'] == 2 \
                                                  else (noName if x['Name Match'] == 0 \
                                                  else (noAddress if not x['Address Match'] else 'uh oh')))), axis=1)

        df['Name Match'] = df.apply(lambda x: True if x['Name Match'] == 1 \
                    else ('Check' if x['Name Match'] == 2 else False), axis=1)
        df['Match \n1 = yes, 0 = no'] = ""

    #International
    elif IndustryType == '6':
        print "Intl Matching"
        df['Robot Suggestion'] = df.apply(lambda x: liveSync if x['Live Sync'] == 1 \
            else liveSuppress if x['Live Suppress'] == 1 \
                else noMatchFBDupeError if x['FB Error']\
                    else noFBMatch if x['External ID'] in bpgid \
                        else userMatch if x['User Match'] == 1 \
                            else checkMissing if (x['No Name'])\
                                else noName if x['Name Match'] == 0 \
                                    else check if x['Name Match'] == 2 and (x['Address Match'] == True or x['Phone Match'] == True)\
                                        else matchText if x['Name Match'] == 1 \
                                            and x['Address Match'] == True or x['Phone Match'] == True\
                                                else checkMissing if x['No Address']\
                                                    else noAddress if x['Address Match'] == False\
                                                        else noMatch, axis=1)

        df['Name Match'] = df.apply(lambda x: True if x['Name Match'] == 1 \
                    else ('Check' if x['Name Match'] == 2 else False), axis=1)
        df['Match \n1 = yes, 0 = no'] = ""

     #All other industries
    else:
        print "Other Matching"
        #Applies Match rules based on new columns.
        df['Robot Suggestion'] = df.apply(lambda x: liveSync if x['Live Sync'] == 1 \
            else liveSuppress if x['Live Suppress'] == 1 \
                else noMatchFBDupeError if x['FB Error']\
                    else noFBMatch if x['External ID'] in bpgid \
                        else userMatch if x['User Match'] == 1 \
                            else checkMissing if (x['No Name'])\
                                else noName if x['Name Match'] == 0\
                                    else check if x['Name Match'] == 2 and (x['Address Match'] == True or x['Phone Match'] == True)\
                                        else matchText if x['Name Match'] == 1\
                                            and x['Address Match'] == True or x['Phone Match'] == True\
                                                else checkMissing if x['No Address']\
                                                    else noAddress if x['Address Match'] == False\
                                                    else noMatch, axis=1)
    
        #we might want to move where this lives.        

    df['Robot Suggestion'] = df.apply(lambda x: 'No Match - People Name' if x['ExtraPeopleNamesInListing'] else x['Robot Suggestion'], axis=1)            

    df['Name Match'] = df.apply(lambda x: True if x['Name Match'] == 1 \
                    else ('Check' if x['Name Match'] == 2 else False), axis=1)
    df['Match \n1 = yes, 0 = no'] = ""



def calculateTotalScore(df):
    global reportType


    #GET ALL THE SCORE THEN GIVE THEM WEIGHTING THEN CREATE A NEW TOTAL SCORE COLUMN
    if reportType == 2:
        df['Total Score'] = df.apply(lambda row: float(row['Name Score'])*.6 + float(row['Address Score'])*.3\
                                + float(row['Phone Score'])*.1 + float(row['Claimed Score']) + float(1337*row['User Match']), axis=1)
    else:
        df['Total Score'] = df.apply(lambda row: float(row['Name Score'])*.6 + float(row['Address Score'])*.3\
                                + float(row['Phone Score'])*.1 + float(row['Claimed Score']), axis=1)


    #FB Brand page anti match - NEED TO DO

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


       #If Listing ID is matched to more than one location
def ExternalID_De_Dupe(df):
    df = df.sort_values(['Match', 'Listing ID', 'Total Score'], ascending=[True, True, True])
#    print df['Match']
    df = df.reset_index(drop=True)

    for index, row in df.iterrows():
        if index < df.shape[0]-1:
            if row['Match'] == 1 and df.iloc[index+1]['Match'] == 1:
                if row['Listing ID'] == df.iloc[index+1]['Listing ID']:
#                    row['Match'] = 0
                    df.set_value(index, 'Match', 0)
#    print df['Match']
    return df


#Reads CSV or XLSX files
def readFile(xlsFile):

    print 'Reading file...'
    if xlsFile[-4:] == 'xlsx':
        x1 = pd.ExcelFile(xlsFile)
        df = x1.parse(0)

    elif xlsFile[-3:] == 'csv':
        df = pd.read_csv(xlsFile, encoding='utf-8')
    else:
        raise Exception('What kind of file did you give me, bro?')
        sys.exit()

    #Finds Business ID
    bid = getBusIDfromLoc(df.loc[0, 'Location ID'])
    return df, bid

#Reads CSV or XLSX files
def readMatchedFile(xlsFile):

    print 'Reading file...'
    if xlsFile[-4:] == 'xlsx':
        x1 = pd.ExcelFile(xlsFile)
        df = x1.parse(0)

    elif xlsFile[-3:] == 'csv':
        df = pd.read_csv(xlsFile, encoding='utf-8')
    else:
        raise Exception('Please provide a csv or xlsx file.')
    return df

#main runtime function
def main(df, IndustryType, bid):
    row = 0
    if IndustryType == '3':
    #Gets Providers First and Last name. Saves to column 'Provider Name'
        DoctorNameDF = getProviderName(df)
        df = df.merge(DoctorNameDF, on='Location ID', how='left')

    lastcol = df.shape[1]
    row = df.shape[0]

#Compares all, suggests matches
    compareData(df, IndustryType, bid)

    print 'Creating matching questions...'
 #Completes Matching Question sheet
    matchingNameQs = matchingQuestions(df)

#Gets save file path
    FilepathMatch = os.path.expanduser("~\Documents\Python Scripts\\"+ re.sub(r'[\\/*?:"<>|]', "", getBusName(bid)) +\
    " AutoMatcher Output "+ str(date.today().strftime("%Y-%m-%d")) + " " + str(time.strftime("%H.%M.%S")) +".xlsx")

#Reorders relevant columns to end
    columns = df.columns.tolist()
    reorder = ['Location Name', 'Listing Name', \
    'Name Score', 'Location Address', 'Listing Address', 'Address Score', 'Distance (M)', \
    'Name Match', 'Phone Match', 'Address Match', 'Phone or Address Match', 'Geocode Match', 'Robot Suggestion', 'Match \n1 = yes, 0 = no']
    for col in reorder:
        columns.append(columns.pop(columns.index(col)))
    df = df[columns]

#Sorts df
    df = df.sort_values(by=['Robot Suggestion', 'Name Score'], ascending=[True, False])

    print 'Writing file...'
    writer = pd.ExcelWriter(FilepathMatch, engine='xlsxwriter', options={'strings_to_urls': False})
#    writer = pd.ExcelWriter(FilepathMatch, engine='xlsxwriter', options={'strings_to_urls': False, 'constant_memory': False})
    df.to_excel(writer, sheet_name="Result", index=False, encoding='utf8')
    matchingNameQs.to_excel(writer, sheet_name="Matching Questions", index=False)
    workbook = writer.book
    worksheet = writer.sheets["Result"]
    worksheet.set_zoom(80)
    matchingSheet = writer.sheets['Matching Questions']

    print 'Formatting file...'
    headerformat = workbook.add_format({
        'bold': True,
        'text_wrap': True})

    df['User Match'] = df['User Match'].apply(lambda x: True if x == 1 else False)
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, headerformat)

    # Define column formatting
    formatBlue = workbook.add_format({'bg_color': '#77e8da', 'border' : 1, 'border_color': '#c0c0c0'}) #
    formatRed = workbook.add_format({'bg_color': '#f47676', 'border' : 1, 'border_color': '#c0c0c0'}) #
    formatYellow = workbook.add_format({'bg_color': '#f5cb70', 'border' : 1, 'border_color': '#c0c0c0'}) #
    formatPurp = workbook.add_format({'bg_color': '#d9b3ff', 'border' : 1, 'border_color': '#c0c0c0'}) #
    formatOrange = workbook.add_format({'bg_color': '#ffaa80', 'border' : 1, 'border_color': '#c0c0c0'}) #
    formatGreen = workbook.add_format({'bg_color': '#c6efce', 'border' : 1, 'border_color': '#c0c0c0'}) #

    namescorecol = df.columns.get_loc("Name Score")
    addressscorecol = df.columns.get_loc("Address Score")
    namecol = df.columns.get_loc("Location Name")
    addresscol = df.columns.get_loc("Location Address")
    robotcol = df.columns.get_loc("Robot Suggestion")
    lastgencol = df.columns.get_loc("Match \n1 = yes, 0 = no")
    Lat = df.columns.get_loc("Listing Latitude")
    LastPostDate = df.columns.get_loc("Listing Latitude")
    nameMatchCol = df.columns.get_loc("Name Match")
    geocodeMatchCol = df.columns.get_loc("Geocode Match")

    worksheet.set_row(0, 45)

    #worksheet.set_column(Lat , LastPostDate, None, None, {'hidden': 1})
    worksheet.set_column(0, namecol-1, None, None, {'hidden': 1})
#    worksheet.set_column(0, namecol-1, None, None, {'collapsed': 1})
    #worksheet.set_column(lastcol+2, lastcol+3, None, None, {'hidden': 1})
    worksheet.set_column(namecol, namecol+1, 45)
    worksheet.set_column(namescorecol, namescorecol, 7.5)
    worksheet.set_column(addressscorecol, addressscorecol, 7.5)
    worksheet.set_column(addresscol, addresscol+1, 27)
    worksheet.set_column(robotcol, robotcol, 23.11)
    worksheet.set_column(robotcol+1, robotcol+1, 14.22)

    worksheet.autofilter(0, 0, 0, lastgencol)

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
                                                             'format':   headerformat})

    worksheet.conditional_format(1, robotcol, row, lastgencol, {'type':'text',
                                                                'criteria': 'begins with',
                                                                'value':    "Match",
                                                                'format':   formatBlue})

    worksheet.conditional_format(1, robotcol, row, lastgencol, {'type':'text',
                                                                'criteria': 'begins with',
                                                                'value':    "User Match",
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

    matchingSheet.set_column(0, 1, 45)

    print 'Saving...'
    try:
        writer.save()


        os.startfile(FilepathMatch)
#        os.system('start excel.exe "'+os.path.expanduser(FilepathMatch)+'"' % (sys.path[0], ))
        app.AllDone(df, 1, "\nDone! Results have been wrizzled to your Excel file. 1love <3"+"\nMatching Template here:\n"+ FilepathMatch)
        print "\nDone! Results have been wrizzled to your Excel file. 1love <3"
        print "\nMatching Template here:"
        print FilepathMatch
   #Opens explorer window to path of output
        subprocess.Popen(r'explorer /select,'+os.path.expanduser(FilepathMatch))

    except IOError:
        app.AllDone(df, 1, "\nIOError: Make sure your Excel file is closed before re-running the script.")
        print "\nIOError: Make sure your Excel file is closed before re-running the script."

#Pulls all location and listing match data
def sqlPull(bid, folderID, labelID, ReportType, IndustryType):
    print 'pulling data'
    #Pull Location info and Listing IDs
    #IF LISTINGS:
    if ReportType == 0 or ReportType == 2:
        SQL_QueryMatches = open(os.path.expanduser("~/Documents/Changing-the-World/SQL Data Pull/1. Pull Matches.sql")).read()
    #IF SUPPRESSION:
    elif ReportType == 1:
        SQL_QueryMatches = open(os.path.expanduser("~/Documents/Changing-the-World/SQL Data Pull/1. Suppression - Pull Matches.sql")).read()
    else:
        sys.exit()
    SQL_QueryMatches = SQL_QueryMatches.splitlines()

 #Subs out variables for Account ID numbers
    for index, line in enumerate(SQL_QueryMatches):
        SQL_QueryMatches[index] = line.replace('@bizid', str(bid))

    if folderID != 0:
        for index, line in enumerate(SQL_QueryMatches):
            SQL_QueryMatches[index] = line.replace('--left join alpha.location_tree_nodes ltn on ltn.id=l.treeNode_id',\
                    'left join alpha.location_tree_nodes ltn on ltn.id=l.treeNode_id')\
                    .replace('--and l.treeNode_id=@folderid', 'and l.treeNode_id=@folderid').replace('@folderid', str(folderID))

    if labelID != 0:
        for index, line in enumerate(SQL_QueryMatches):
            SQL_QueryMatches[index] = line.replace('--left join alpha.location_labels ll on ll.location_id=l.id', \
            'left join alpha.location_labels ll on ll.location_id=l.id')
        for index, line in enumerate(SQL_QueryMatches):
            SQL_QueryMatches[index] = line.replace('--and ll.label_id=@labelid', 'and ll.label_id=@labelid')
        for index, line in enumerate(SQL_QueryMatches):
            SQL_QueryMatches[index] = line.replace('@labelid', str(labelID))


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
    ListingIDs = SQL_DataMatches['Listing ID']
    ListingIDs = ListingIDs.map(lambda x: x.lstrip('\''))

    #Gets Listing Data from Listing IDs
    SQL_QueryListings = open(os.path.expanduser("~/Documents/Changing-the-World/SQL Data Pull/2. Pull Listings Data.sql")).read()
    SQL_QueryListings = SQL_QueryListings.splitlines()

    for index, line in enumerate(SQL_QueryListings):
        SQL_QueryListings[index] = line.replace('@ListingIDs', ','.join(map(str, ListingIDs)))
    if ReportType == 1:
        for index, line in enumerate(SQL_QueryListings):
            SQL_QueryListings[index] = line.replace('--left join warehouse.listings_additional_google_fields lagf on lagf.listing_id=wl.id', \
            'left join warehouse.listings_additional_google_fields lagf on lagf.listing_id=wl.id')
        for index, line in enumerate(SQL_QueryListings):
            SQL_QueryListings[index] = line.replace('--AND (lagf.googlePlaceId not in (select tl.externalid from tags_listings tl join tags_listings_unavailable_reasons tlur on tlur.location_id = tl.location_id join alpha.tags_unavailable_reasons tur on tur.id=tlur.tagsunavailablereason_id where tlur.partner_id=715 and tur.showasWarning is false) or lagf.googlePlaceId is null)', \
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

 #if suppression
    if ReportType == 1:
        liveSyncSQL = "SELECT tl.location_id as \"Location ID\", tl.partner_id as \"Publisher ID\", tl.externalId AS 'Sync External ID', tl.url AS 'Sync URL' from alpha.tags_listings tl "+ \
        "WHERE tl.tagsListingStatus_id=6 and tl.location_id in ("+','.join(SQL_DataMatches['Location ID'].to_csv(path=None, sep=',', index=False).split('\n'))[:-1]+");"

        Yext_Prod_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
        syncIDsDF = pd.read_sql(liveSyncSQL, con=Yext_Prod_DB)

    #Combine results
    df = SQL_DataListings.merge(SQL_DataMatches, on='Listing ID', how='outer')
    if ReportType == 1:
        df = df.merge(syncIDsDF, on=['Location ID', 'Publisher ID'], how='left')

 #If doctors, pull npi
    if IndustryType == 3:
        NPI_SQL = open(os.path.expanduser("~/Documents/Changing-the-World/SQL Data Pull/3. Pull NPI Data.sql")).read()
        NPI_SQL = NPI_SQL.splitlines()
        for index, line in enumerate(NPI_SQL):
            NPI_SQL[index] = line.replace('@bizid', str(bid))
        NPI_SQL = ' '.join(NPI_SQL)
        Yext_Prod_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
        NPI_Data = pd.read_sql(NPI_SQL, con=Yext_Prod_DB)

        df = df.merge(NPI_Data, on=['Location ID'], how='left')

    return df


def pullFBManualMatches(ListingIDs):
    ListingIDs = ListingIDs.replace('\'', '')
    FB_SQL = open(os.path.expanduser("~/Documents/Changing-the-World/SQL Data Pull/Facebook - User Match Listing IDs.sql")).read()
    FB_SQL = FB_SQL.splitlines()
    for index, line in enumerate(FB_SQL):
        FB_SQL[index] = line.replace('@listingid', str(ListingIDs))
    FB_SQL = ' '.join(FB_SQL)
    Yext_Mat_DB = MySQLdb.connect(host="127.0.0.1", port=5009, db="alpha")
    FB_Data = pd.read_sql(FB_SQL, con=Yext_Mat_DB)
    FB_Data['Listing ID'] = FB_Data['Listing ID'].astype(str)
    return FB_Data
   # df = df.merge(NPI_Data, on=['Location ID'], how='left')

#    return df
#Gets Alt Name policies at business level
def getAltName(bid):
    SQL_AltNameQuery = open(os.path.expanduser("~/Documents/Changing-the-World/SQL Data Pull/Alt Name policies.sql")).read()

    SQL_AltNameQuery = SQL_AltNameQuery.replace('@bizid', str(bid))
    try:
        Yext_SMS_DB = MySQLdb.connect(host="127.0.0.1", port=5007, db="alpha")
        AltNames = pd.read_sql(SQL_AltNameQuery, con=Yext_SMS_DB)['policyString']
    except:
        AltNames = ""
    BusNames = []
    for name in AltNames:
        BusNames.append(name)

    return BusNames

#Gets Business Account name
def getBusName(bid):

    SQL_BusQuery = 'Select name from alpha.businesses where id='+str(bid)+';'
    try:
        Yext_OPS_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
        busName = pd.read_sql(SQL_BusQuery, con=Yext_OPS_DB)['name'][0]
    except:
        busName = "None"
    return busName

#Gets Business ID from location ID
def getBusIDfromLoc(locationID):
    #locationID=locationID.replace("\'","")
    SQL_BusIDQuery = 'SELECT business_id from alpha.locations where id='+str(locationID).replace("\'", "")+';'
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
    SQL_DoctorNameQuery = open(os.path.expanduser("~/Documents/Changing-the-World/SQL Data Pull/Doctor Name.sql")).read()
    locationIDs = df['Location ID']

    SQL_DoctorNameQuery = SQL_DoctorNameQuery.replace('@locationIDs', ','.join(map(str, locationIDs)))
    Yext_OPS_DB = MySQLdb.connect(host="127.0.0.1", port=5020, db="alpha")
    DoctorNamesIDs = pd.read_sql(SQL_DoctorNameQuery, con=Yext_OPS_DB)

    return DoctorNamesIDs

#Identifies common names that are in checks or No Match - Name
def matchingQuestions(df):

    Qdf = df[df['Robot Suggestion'].isin(['No Match - Name', 'Check Name'])]
    pivot = pd.pivot_table(Qdf, values='Link ID', index='Listing Name', aggfunc='count')
    listingNames = pd.DataFrame(pivot)
    # ts=time.time()

    if listingNames.shape[0] > 1 and listingNames.shape[1] > 0 and listingNames.shape[0] < 5000:
        listingNames.columns = ['Count']
        listingNames = listingNames.sort_values(by=['Count'], ascending=False)
        listingNames = listingNames.reset_index()
        listingNames['Listing Name'] = listingNames['Listing Name'].apply(cleanName)

        listingNames.columns = listingNames.columns.str.replace('\s+', '_')

        alias = {l : r for l, r in itertools.product(listingNames.Listing_Name, listingNames.Listing_Name) if l < r and fuzz.token_sort_ratio(l, r) >= 70}

        if len(alias) > 0:
            series = listingNames.Count.groupby(listingNames.Listing_Name.replace(alias)).sum()

            mergedListNames = series.to_frame()

            mergedListNames = mergedListNames.reset_index()
            mergedListNames.columns = mergedListNames.columns = ['Listing Name', 'Count']
            mergedListNames = mergedListNames.sort_values(by=['Count'], ascending=False)
            mergedListNames = mergedListNames[mergedListNames['Count'] >= 5]
            mergedListNames = mergedListNames.replace(np.nan, '', regex=True)
            mergedListNames = mergedListNames[mergedListNames['Listing Name'] != '']
        elif not listingNames.empty:
            listingNames = listingNames.sort_values(by=['Count'], ascending=False)
            listingNames = listingNames.reset_index()
            listingNames['Listing_Name'] = listingNames['Listing_Name'].apply(cleanName)
            listingNames = listingNames[listingNames['Count'] > 5]
            return listingNames

        else:
            return pd.DataFrame([{'Listing Name' : 'None', 'Count' : '0'}])

  #      te = time.time()

      #  print te-ts

        return mergedListNames

    elif not listingNames.empty:
        listingNames.columns = ['Count']
        listingNames = listingNames.sort_values(by=['Count'], ascending=False)
        listingNames = listingNames.reset_index()
        listingNames['Listing Name'] = listingNames['Listing Name'].apply(cleanName)
        listingNames = listingNames[listingNames['Count'] > 5]
        return listingNames

    else:
        return pd.DataFrame([{'Listing Name' : 'None', 'Count' : '0'}])


#saves down upload linkages file
def writeUploadFile(df):
    filePath = os.path.expanduser("~\Documents\Python Scripts\\"+ \
                                   getBusName(getBusIDfromLoc(df.loc[0, 'locationId']))+\
                                   " Upload Linkages "+ str(date.today().strftime("%Y-%m-%d")) \
                                    + " " + str(time.strftime("%H.%M.%S")) +".csv")

    print 'writing file'
    #writer = pd.ExcelWriter(filePath, engine='xlsxwriter')
    df.to_csv(filePath, sheet_name="Linkages", encoding='utf-8', index=False)
    return "\nUpload Linkages available: "+ filePath

#For suppression, finds live google syncs, to antimatch
def GoogleIDs(bid):

    #Gets IDs from GMB
    inputFile = open(os.path.expanduser("~/Documents/entops/Temporary Matching Template Workaround/GoogleIDs1.sql")).read()
    splitQueries = inputFile.split(";")
    SQL_Query715 = splitQueries[1]
    SQL_Query715 = SQL_Query715.splitlines()


    for index, line in enumerate(SQL_Query715):
        if line == "WHERE al.business_id = @BUSINESS_ID":
            SQL_Query715[index] = line.replace('@BUSINESS_ID', str(bid))
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
    splitQueries = inputFile.split(";")
    SQL_Query713 = splitQueries[1]
    SQL_Query713 = SQL_Query713.splitlines()

    for index, line in enumerate(SQL_Query713):
        if line == "WHERE al.business_id = @BUSINESS_ID":
            SQL_Query713[index] = line.replace('@BUSINESS_ID', str(bid))


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


    listingIDs = SQL_Data_713.iloc[:, 0:1]
    listingIDs.columns = ['ID']
    LiDs = listingIDs["ID"].tolist()
    LiDs = ','.join(str(e) for e in LiDs)

    #Gets external IDs for Google Reviews listing IDs
    inputFile = open(os.path.expanduser("~/Documents/entops/Temporary Matching Template Workaround/GoogleIDs3.sql")).read()
    splitQueries = inputFile.split(";")
    SQL_Query_ext = splitQueries[0]
    SQL_Query_ext = SQL_Query_ext.splitlines()

    for index, line in enumerate(SQL_Query_ext):
        if line == "where wl.id in (123456)":
            SQL_Query_ext[index] = line.replace('123456', LiDs)


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
    col1 = SQL_Data_715.iloc[:, 0:1]
    col2 = SQL_Data_ext.iloc[:, 0:1]

    col1.columns = ["Live Google External IDs"]
    col2.columns = ["Live Google External IDs"]

    SQL_Data = pd.concat([col1, col2], ignore_index=True)
    SQL_Data = SQL_Data.drop_duplicates()

    SQL_Data["Live Google External IDs"] = "'" + SQL_Data["Live Google External IDs"].astype(str)

        # SQL_Data.to_excel()
    newSQL_Data = SQL_Data.set_index('Live Google External IDs')

    #print(newSQL_Data)
    Yext_Prod_DB.close()

    return newSQL_Data


#GUI Tkinter section!
class MatchingInput(Tkinter.Frame):

    #Initial window
    def __init__(self, master):


        root.protocol("WM_DELETE_WINDOW", self._delete_window)
        Tkinter.Frame.__init__(self, master, padx=10, pady=10)
        master.title("AutoMatcher Setup")

        master.minsize(width=500, height=300)
        root.lift()


        self.IntroLabel = Label(master, text="   Welcome to the AutoMatcher! This will suggest matches"\
                                +" based on inputs as well as create an upload file"+
                                "\n\n   To start, select the first option. "\
                                +"\n   Review all checks and fill in a match status."
                                +"\n   Afterwards, run the program again, and select the second option."\
                                +"\n\n   Make sure you connect to SDM and the J:\ Drive!").grid(row=0, column=0, columnspan=2, pady=(0, 20))
        self.processChoice = IntVar()
        self.processChoice.set(-1)
        self.StartMatch = Radiobutton(master, text="Pull data/enter file and suggest matches",\
                                      variable=self.processChoice, value=0).grid(row=1, column=0)
        self.ChecksChoice = Radiobutton(master, text="Create upload file based on reviewed matches",\
                                        variable=self.processChoice, value=1).grid(row=1, column=1)

        self.Next = Button(master, text="Next", command=lambda: [self.initialSettingsWindow() \
                                     if self.processChoice.get() == 0 else (self.inputChecks() \
                                      if self.processChoice.get() == 1 else self.processChoice.set(-1))]).grid(row=2, column=0, pady=(30, 0), sticky=E)

        self.Quit = Button(master, text="Quit", command=lambda: [root.destroy()]).grid(row=3, column=0, sticky=E, pady=25)

#If the user has manually checked the matches, this will take those in, determine Match statuses, and produce upload document
    def inputChecks(self):
        self.master.withdraw()
        self.UploadSetup = Toplevel()
        self.UploadSetup.protocol("WM_DELETE_WINDOW", self._delete_window)

        self.ReportType = IntVar()
        self.ReportType.set(-1)
        self.ReportLabel = Label(self.UploadSetup, text="Select report type:").grid(row=4, column=0, pady=(30, 0), sticky=W)

        self.Listings = Radiobutton(self.UploadSetup, text="Listings", variable=self.ReportType, value=0).grid(row=5, column=0, sticky=W)
        self.Suppression = Radiobutton(self.UploadSetup, text="Suppression", variable=self.ReportType, value=1).grid(row=6, column=0, sticky=W)

        self.nextButton = Button(self.UploadSetup, text="Next", command=lambda: [self.readCheckedFile() \
                            if (self.ReportType.get() > -1) else self.ReportType.set(self.ReportType.get())])\
                                                        .grid(row=7, column=0, sticky=W, pady=(35, 0))
#Reads checked data file
    def readCheckedFile(self):
        self.UploadSetup.destroy()
#Takes in completed matches file with checks filled out
        checkedFile = tkFileDialog.askopenfilename(initialdir="/", title=\
                                 "Select completed matching file with Check column filled out",\
                                 defaultextension="*.xlsx;*.xls", \
                                 filetypes=(("Excel files", "*.xlsx;*.xls"), ("CSV", "*.csv"), ('All files', '*.*')))
        checkedDF = readMatchedFile(checkedFile)

#Checks to see if all rows asking for a check have manual review
        allChecksComplete = True

        for index, row in checkedDF.iterrows():
            if  ('Check' in row['Robot Suggestion'] and (isnan(row['Match \n1 = yes, 0 = no']))):
                allChecksComplete = False

#Exits if manual review incomplete
        if not allChecksComplete:
            self.errorBox = Toplevel()
            self.errorMsg = Label(self.errorBox, text="Please complete all checks first. Bye.").pack()
            self.okButton = Button(self.errorBox, text="OK", command=lambda: [root.destroy()]).pack()
#If complete, determines matches, creates upload
        else:
            checkedDF['Match'] = checkedDF.apply(lambda x: 1 if 'Match Suggested' in x['Robot Suggestion'] else 0, axis=1)
            checkedDF['Match'] = checkedDF.apply(lambda x: 1 if x['Match \n1 = yes, 0 = no'] == 1 else \
                                                (0 if x['Match \n1 = yes, 0 = no'] == 0 else x['Match']), axis=1)

            busID = getBusIDfromLoc(checkedDF.loc[0, 'Location ID'])
            busName = getBusName(busID)

            print 'checking live sync'
            if self.ReportType.get() == 1:
                googleLiveIDs = GoogleIDs(busID)
                googleLiveIDs = googleLiveIDs.reset_index()
                checkedDF['Google Live Sync'] = checkedDF['External ID'].isin(googleLiveIDs['Live Google External IDs'])
                checkedDF['Live Sync'] = checkedDF.apply(lambda x: 1 if x['Google Live Sync'] else x['Live Sync'], axis=1)

            checkedDF['Match'] = checkedDF.apply(lambda x: 0 if x['Live Sync'] == 1 else x['Match'], axis=1)
            checkedDF['Match'] = checkedDF.apply(lambda x: 0 if x['Live Suppress'] == 1 else x['Match'], axis=1)


            print "External ID deduping\n"

            checkedDF = ExternalID_De_Dupe(checkedDF)


            #writer = pd.ExcelWriter(filePath, engine='xlsxwriter')
            #
            checkedDF['override'] = checkedDF.apply(lambda x: 'Match' if x['Match'] == 1 else 'Antimatch', axis=1)

            if self.ReportType.get() == 0:
                checkedDF['PL Status'] = checkedDF.apply(lambda x: 'Sync' if x['override'] == 'Match' else 'NoPowerListing', axis=1)
            elif self.ReportType.get() == 1:
                checkedDF['PL Status'] = checkedDF.apply(lambda x: 'Suppress' if x['override'] == 'Match' else 'NoPowerListing', axis=1)

            #Total Score at the top
            checkedDF = checkedDF.sort_values(['Location ID', 'Publisher ID', 'Match', 'Total Score']\
                                                              , ascending=[True, True, False, False])
            checkedDF = checkedDF.reset_index(drop=True)

            #If Listings Type
            if self.ReportType.get() == 0:
                for index, row in checkedDF.iterrows():
                    if index != 0:
                    #If Listings, Diagnostic. or Facebook Template: if previous listing is a Match of the same Loc/Pub ID combination, then NPL
                        if row['Location ID'] == checkedDF.iloc[index-1]['Location ID'] and row['Publisher ID']\
                             == checkedDF.iloc[index-1]['Publisher ID'] and row['override'] == 'Match':\
                            checkedDF.set_value(index, "PL Status", "NoPowerListing")
                        else:
                            pass

            completeCopyFilePath = os.path.expanduser("J:\zAutomatcherData\CheckedFiles\\"+ \
                                               os.getenv('username')+ " - " + busName+" Data "+ str(date.today().strftime("%Y-%m-%d")) \
                                                + " " + str(time.strftime("%H.%M.%S")) +".csv")
            checkedDF.to_csv(completeCopyFilePath, sheet_name="Data", encoding='utf-8', index=False)


            #EXTERNAL ID DEDUPE
            print "Creating Upload Overrides\n"

            #checkedDF=calculateTotalScore(checkedDF)
            uploadDF = checkedDF[['Publisher ID', 'Location ID', 'Listing ID', 'override', 'PL Status']]
            uploadDF.columns = ["partnerId", "locationId", "listingId", "override", "PL Status"]

            try:
                uploadDF['listingId'] = uploadDF.apply(lambda x: x['listingId'].replace('\'', ''), axis=1)
            except:
                pass

#            If Suppression Type
            if self.ReportType.get() == 1:
                print "Creating Suppression Report"


                suppRepDF = checkedDF[checkedDF['PL Status'] == 'Suppress']

                ################Create a Suppression Approval File
                pivot = pd.pivot_table(suppRepDF, values='PL Status', index='Publisher', aggfunc='count')
#                print pivot
                publisherNames = pd.DataFrame(pivot)
                publisherNames.columns = ['Count of Duplicates']

                publisherNames = publisherNames.sort_index(ascending=True)
                print publisherNames

#                publisherNames = publisherNames.reset_index()
                businessName = busName

                #Needs logic if certain things exist
                claimedFBDF = checkedDF[(checkedDF['PL Status'] == 'Suppress')\
                                        & (checkedDF['Publisher'] == 'Facebook') & (checkedDF['Advertiser/Claimed'] == 'Claimed')]
                print "mixing mixing"
                if claimedFBDF.shape[0] > 0:
                    print "there are claimed FB pages! be sure to review them!"

                    filePath = os.path.expanduser("~\Documents\Python Scripts\\" + businessName+\
                                               " Suppression Approval File "+ str(date.today().strftime("%Y-%m-%d")) + " " + str(time.strftime("%H.%M.%S")) +".xlsx")
                else:
                    filePath = os.path.expanduser("~\Documents\Python Scripts\\" + businessName+\
                                               " Suppression Summary File "+ str(date.today().strftime("%Y-%m-%d")) + " " + str(time.strftime("%H.%M.%S")) +".xlsx")

                suppwriter = pd.ExcelWriter(filePath, engine='xlsxwriter')
                publisherNames.to_excel(suppwriter, sheet_name="Summary", index=True, encoding='utf8', startrow=2)
                workbook = suppwriter.book
                worksheet = suppwriter.sheets['Summary']
                worksheet.set_column(0, 0, 26)
                worksheet.set_column(1, 1, 17)
                worksheet.write(0, 0, businessName+ " - Suppression Approval Summary - "+ str(date.today().strftime("%Y-%m-%d")))

                worksheet.write(len(publisherNames)+3, 0, "Grand Total")
                worksheet.write(len(publisherNames)+3, 1, publisherNames['Count of Duplicates'].sum())
                worksheet.set_zoom(80)


                if claimedFBDF.shape[0] > 0:
                    claimedFBDF = claimedFBDF[['Store ID', 'Location ID',\
                                               'Location Name', 'Location Address', 'Location Address 2',\
                                               'Location City', 'Location State', 'Location Zip', 'Location Phone',\
                                                'Sync External ID', 'Sync URL', 'Listing Name', \
                                               'Listing Address', 'Listing Address 2', 'Listing City', 'Listing State', \
                                               'Listing Zip', 'Listing Phone', \
                                               'Listing URL', 'Listing ID', 'External ID',\
                                                'Last Post Date']]
                    claimedFBDF['Yes/No'] = ""
                    claimedFBDF['Reason'] = ""

                    claimedFBDF.to_excel(suppwriter, sheet_name="Facebook Claimed Pages", index=False, startrow=0)
                    worksheet = suppwriter.sheets['Facebook Claimed Pages']

                suppwriter.save()


                xlApp = win32com.client.Dispatch("Excel.Application")
                SuppFile = xlApp.Workbooks.Open(filePath)
                toolkit = os.path.expanduser("~\Documents\entops\Templates and Macros\EntOpsMacroToolkit.xlam")
                a = xlApp.Application.Run('\''+toolkit+"\'!AutoMatcherSuppressionReport.Run_Suppression_Report")

                reportDF = checkedDF[checkedDF['PL Status'] == 'Suppress']
                reportDF = reportDF[['Store ID', 'Location ID',\
                                           'Location Name', 'Location Address', 'Location Address 2',\
                                           'Location City', 'Location State', 'Location Zip', 'Location Phone',\
                                            'Sync External ID', 'Sync URL', 'Listing Name', \
                                           'Listing Address', 'Listing Address 2', 'Listing City', 'Listing State', \
                                           'Listing Zip', 'Listing Phone', \
                                           'Listing URL', 'Listing ID', 'External ID',\
                                            'Last Post Date', 'Publisher ID', 'Publisher']]


                fullFilePath = os.path.expanduser("~\Documents\Python Scripts\\" + businessName+\
                                               " Full Suppression Details "+ str(date.today().strftime("%Y-%m-%d")) + " " + str(time.strftime("%H.%M.%S")) +".xlsx")

                reportWriter = pd.ExcelWriter(fullFilePath, engine='xlsxwriter')
                publisherNames.to_excel(reportWriter, sheet_name="Summary", index=True, encoding='utf8', startrow=2)
                workbook = reportWriter.book
                worksheet = reportWriter.sheets['Summary']
               # format1= workbook.add_format({
              #                         'bold': False,})
              #
                worksheet.set_column(0, 0, 26)
                worksheet.set_column(1, 1, 17)
                worksheet.write(0, 0, businessName+ " - Suppression Approval Summary - "+ str(date.today().strftime("%Y-%m-%d")))

                worksheet.write(len(publisherNames)+3, 0, "Grand Total")
                worksheet.write(len(publisherNames)+3, 1, publisherNames['Count of Duplicates'].sum())
                worksheet.set_zoom(80)
#                                            nocols = ['Location URL','Yes/No', 'Reason']

                reportDF.to_excel(reportWriter, sheet_name="Facebook Claimed Pages", index=False, startrow=0)
                worksheet1 = reportWriter.sheets['Facebook Claimed Pages']
                reportWriter.save()

                xlApp1 = win32com.client.Dispatch("Excel.Application")
                reportFile = xlApp1.Workbooks.Open(fullFilePath)
                toolkit = os.path.expanduser("~\Documents\entops\Templates and Macros\EntOpsMacroToolkit.xlam")
                a = xlApp.Application.Run('\''+toolkit+"\'!AutoMatcherSuppressionReport.Run_Suppression_Report")

            print "Writing to File"
            self.AllDone(uploadDF, 2, writeUploadFile(uploadDF))
            print t0
            t1 = time.time()
            print t1
            print t1-t0
       #remove this once we do something here

            #Gets user input for how to set up matcher
    def initialSettingsWindow(self):

        self.master.withdraw()
        self.settingWindow = Toplevel()
        self.settingWindow.protocol("WM_DELETE_WINDOW", self._delete_window)
        self.IndustryType = IntVar()
        self.IndustryType.set(-1)
        self.IndustryLabel = Label(self.settingWindow, text="Select Industry Type").grid(row=0, column=0, columnspan=2, pady=(10, 10), sticky=W)

        self.Normal = Radiobutton(self.settingWindow, text="Normal", variable=self.IndustryType, value=0).grid(row=1, column=0, sticky=W)
        self.Auto = Radiobutton(self.settingWindow, text="Auto", variable=self.IndustryType, value=1).grid(row=2, column=0, sticky=W)
        self.Hotel = Radiobutton(self.settingWindow, text="Hotel", variable=self.IndustryType, value=2).grid(row=2, column=1, sticky=W)
        self.Doctor = Radiobutton(self.settingWindow, text="Healthcare Doctor", variable=self.IndustryType, value=3).grid(row=2, column=2, sticky=W)
        self.Facility = Radiobutton(self.settingWindow, text="Healthcare Facility", variable=self.IndustryType, value=4).grid(row=3, column=2, sticky=W)
        self.Agent = Radiobutton(self.settingWindow, text="Agent", variable=self.IndustryType, value=5).grid(row=3, column=1, sticky=W)
        self.International = Radiobutton(self.settingWindow, text="International", variable=self.IndustryType, value=6).grid(row=3, column=0, sticky=W)

        #Report Type Designation
        self.ReportType = IntVar()
        self.ReportType.set(-1)
        self.ReportLabel = Label(self.settingWindow, text="Select report type:").grid(row=4, column=2, pady=(30, 0), padx=(15, 0), sticky=W)

        self.Listings = Radiobutton(self.settingWindow, text="Listings", variable=self.ReportType, value=0).grid(row=5, column=2, padx=(15, 0), sticky=W)
        self.Suppression = Radiobutton(self.settingWindow, text="Suppression", variable=self.ReportType, value=1).grid(row=6, column=2, padx=(15, 0), sticky=W)
        self.fb = Radiobutton(self.settingWindow, text="Facebook Listings", variable=self.ReportType, value=2).grid(row=7, column=2, padx=(15, 0), sticky=W)
        #self.FB = Radiobutton(self.settingWindow, text="FB", variable=self.ReportType, value=2).grid(row=1, column=2)
        #self.Google = Radiobutton(self.settingWindow, text="Google", variable=self.ReportType, value=x).grid(row=1, column=3)


        #self.quitButton.pack()
        self.dataInput = IntVar()
        self.dataInput.set(0)
        self.inputType = Label(self.settingWindow, text="Select data input type:").grid(row=4, column=0, columnspan=2, pady=(30, 0), sticky=W)

        self.SQL = Radiobutton(self.settingWindow, text="Pull Data from SQL", variable=self.dataInput, value=2).grid(row=5, column=0, columnspan=2, sticky=W)
        self.file = Radiobutton(self.settingWindow, text="Input File", variable=self.dataInput, value=1).grid(row=6, column=0, sticky=W)



        self.nextButton = Button(self.settingWindow, text="Next", command=lambda: [self.detailsWindow() \
                            if (self.IndustryType.get() > -1 and self.ReportType.get() > -1 and self.dataInput.get() > -1)\
                                                         else self.ReportType.set(self.ReportType.get())])\
                                                        .grid(row=8, column=0, sticky=W, pady=(35, 0))


#        self.inputType.grid(row=2,column=0)
#        self.SQL.grid(row=3,column=1, sticky=W)
#        self.file.grid(row=4,column=1, sticky=W)
#        self.nextButton.grid(row=5,column=0,sticky=W,pady=25)

 #Gets File path or business, folder, and label ids to pull from SQL
    def detailsWindow(self):
        self.settingWindow.destroy()
        global reportType
        reportType = self.ReportType.get()
        #File pull
        if self.dataInput.get() == 1:
            fname = tkFileDialog.askopenfilename(initialdir="/", title="Select file", filetypes=(("CSV", "*.csv"), ("Excel files", "*.xlsx;*.xls")))
            df, bid = readFile(fname)
            self.bid = bid
            df['Listing ID'] = df['Listing ID'].map(lambda x: x.lstrip('\''))
            if self.ReportType.get() == 2:
                df = df.merge(pullFBManualMatches(','.join(map(str, df['Listing ID']))), on=['Listing ID'], how='left')
                df['User Match'] = df['User Match'].apply(lambda x: 1 if x == 1 else 0)
            else:
                df['User Match'] = False
            main(df, str(self.IndustryType.get()), bid)
        #SQL
        elif self.dataInput.get() == 2:
            self.detailsW = Tkinter.Toplevel(self)
            self.detailsW.protocol("WM_DELETE_WINDOW", self._delete_window)
            vcmd = self.master.register(self.validate)

            self.InputExplanation = Label(self.detailsW, text=\
            "Please enter account info for data to pull. Business ID is required. \nYou can leave Folder ID and Label ID blank or put 0 if not needed\n").grid(row=5, column=0, columnspan=2)

            self.busIDLabel = Label(self.detailsW, text="Enter Business ID").grid(row=6, column=0, sticky=W)
            self.folderIDLabel = Label(self.detailsW, text="Enter Folder ID").grid(row=7, column=0, sticky=W)
            self.labelIDLabel = Label(self.detailsW, text="Enter Label ID").grid(row=8, column=0, sticky=W)
            self.folderIDLabel = Label(self.detailsW, text="Enter Folder ID").grid(row=7, column=0, sticky=W)
            self.labelIDLabel = Label(self.detailsW, text="Enter Label ID").grid(row=8, column=0, sticky=W)

            if self.ReportType.get() == 2:
                self.FBLabel = Label(self.detailsW, text="Enter Facebook Brand Page ID").grid(row=9, column=0, sticky=W)


            self.bizID = StringVar()
            self.folderID = StringVar()
            self.labelID = StringVar()
            self.fbBrandID = StringVar()
            self.busIDentry = Entry(self.detailsW, validate="key", validatecommand=(vcmd, '%P'), textvariable=self.bizID).grid(row=6, column=1, sticky=W)
            self.folderIDentry = Entry(self.detailsW, validate="key", validatecommand=(vcmd, '%P'), textvariable=self.folderID).grid(row=7, column=1, sticky=W)
            self.labelIDentry = Entry(self.detailsW, validate="key", validatecommand=(vcmd, '%P'), textvariable=self.labelID).grid(row=8, column=1, sticky=W)

            if self.ReportType.get() == 2:
                self.FBEntry = Entry(self.detailsW, validate="key", validatecommand=(vcmd, '%P'), textvariable=self.fbBrandID).grid(row=9, column=1, sticky=W)


            self.pullButton = Button(self.detailsW, text="Pull Data", command=self.pullSQLRun).grid(row=10, column=1, sticky=W)



#Runs SQl Pull
    def pullSQLRun(self):

        if self.bizID.get():
            self.bid = self.bizID.get()
            self.detailsW.destroy()
            if  self.folderID.get():
                folderID = self.folderID.get()
            else:
                folderID = 0
            if self.labelID.get():
                labelID = self.labelID.get()
            else:
                labelID = 0
            df = sqlPull(self.bizID.get(), folderID, labelID, self.ReportType.get(), self.IndustryType.get())
            df['Listing ID'] = df['Listing ID'].map(lambda x: x.lstrip('\''))
            if self.ReportType.get() == 2:
                global bpgid
                bpgid = self.fbBrandID.get()
                df = df.merge(pullFBManualMatches(','.join(map(str, df['Listing ID']))), on=['Listing ID'], how='left')
            else:
                df['User Match'] = False
            main(df, str(self.IndustryType.get()), self.bizID.get())

#button function to ammend businessNames list
    def AddMore(self):
        global businessNames
        global businessNameMatch

#        if self.NewName.get() != "":
#            for i in self.NewName.get().split(","):
#                businessNames.append(cleanName(i))

        self.nameW.destroy()
        self.WordsMust = []
        self.WordsAlt = []
        self.WordsIgnore = []
        self.WordsExclude = []

        for x, v in self.varN.iteritems():
            self.Words[x] = self.varN[x].get()

        for i in range(len(self.moreWords)):
            self.Words[self.moreWords[i].get()] = self.MoreVarN[i].get()

        for key, value in self.varN.iteritems():
            if value.get() == 0:
                self.WordsMust.append(cleanName(key))
            elif value.get() == 1:
                self.WordsAlt.append(cleanName(key))
            elif value.get() == 2:
                self.WordsIgnore.append(cleanName(key))
            elif value.get() == 3:
                self.WordsExclude.append(cleanName(key))

        print self.WordsMust
        print self.WordsAlt
        print self.WordsIgnore
        print self.WordsExclude

        self.PreviousWords.loc[self.indexVal, 'Account_ID'] = self.bid
        self.PreviousWords.loc[self.indexVal, 'Words'] = str(self.Words)
        try:
            self.PreviousWords.to_csv("J:\zAutomatcherData\Words.csv", index=False)
        except:
            pass
        if len(self.WordsAlt) > 0:        
            businessNameMatch = 1
        else:
            businessNameMatch = 0


  #Prints out current business names, asks for additional names
    def namesWindow(self, busNames):
        global businessNames
        global businessNameMatch
        self.all_entries = []
        self.Words = {}
        self.moreWords = {}
        self.moreWordValues = []
        self.count = 0




        for name in busNames:
            self.Words[name] = 1

        try:
            self.PreviousWords = pd.read_csv("J:\zAutomatcherData\Words.csv")
            self.bid = int(self.bid)
            self.PreviousWords['Account_ID'].astype(np.int64)
            self.PreviousWords['Words'] = self.PreviousWords['Words'].map(ast.literal_eval)
            self.indexVal = -1
            self.indexVal = int(self.PreviousWords[self.PreviousWords['Account_ID'] == int(self.bid)].index.values)

        except:
            pass

        if self.indexVal > -1:
            self.Words.update(self.PreviousWords.at[self.indexVal, 'Words'])
        else:
            self.indexVal = self.PreviousWords.shape[0]


        self.nameW = Toplevel()
        self.nameW.protocol("WM_DELETE_WINDOW", self._delete_window)
        self.nameW.minsize(width=300, height=200)


        self.NameIntro = Label(self.nameW, text='Select how each word should be used in matching.'+\
                             '\nOptions:\n\nMust Have: All listings must have the word to be matched.'+\
                             '\n\nAlt Words: Good words to match to, but not mandatory.'+\
                             '\n\nIgnore: Common words that should be ignored from both Location  and Listing Name. This word will be disregarded in matching'+\
                             '\n\nExclude: If these words exist in Listing Name, it will be anti-matched\n')\
                             .grid(row=0, column=0, sticky=W, columnspan=5)

        self.mustHaveLabel = Label(self.nameW, text='Must Have').grid(row=1, column=1, sticky=W)
        self.altWordLabel = Label(self.nameW, text='Alt Word').grid(row=1, column=2, sticky=W)
        self.ignoreLabel = Label(self.nameW, text='Ignore').grid(row=1, column=3, sticky=W)
        self.excludeLabel = Label(self.nameW, text='Exclude').grid(row=1, column=4, sticky=W)
        self.noneLabel = Label(self.nameW, text='None').grid(row=1, column=5, sticky=W)

        self.varN = dict()
        i = 1
        for key, value in self.Words.iteritems():

            self.varN[key] = IntVar()
            self.varN[key].set(value)
            self.WordLabel = Label(self.nameW, text=key).grid(row=i+1, column=0, sticky=W)
            self.mustHave = Radiobutton(self.nameW, variable=self.varN[key], value=0).grid(row=i+1, column=1, sticky=W)
            self.altWord = Radiobutton(self.nameW, variable=self.varN[key], value=1).grid(row=i+1, column=2, sticky=W)
            self.ignore = Radiobutton(self.nameW, variable=self.varN[key], value=2).grid(row=i+1, column=3, sticky=W)
            self.excludeWord = Radiobutton(self.nameW, variable=self.varN[key], value=3).grid(row=i+1, column=4, sticky=W)
            self.NoneWord = Radiobutton(self.nameW, variable=self.varN[key], value=4).grid(row=i+1, column=5, sticky=W)
            i += 1


        self.lastRow = len(self.Words)+2

        self.lastRow2 = self.lastRow
        self.MoreVarN = dict()


        if self.ReportType.get() == 2:
            vcmd = self.master.register(self.validate)
            self.fbBrandID = StringVar()
            self.FBLabel = Label(self.nameW, text="Enter Facebook Brand Page ID")
            self.FBEntry = Entry(self.nameW, validate="key", validatecommand=(vcmd, '%P'), textvariable=self.fbBrandID)
            self.FBLabel.grid(row=self.lastRow2+1, column=0, sticky=W)
            self.FBEntry.grid(row=self.lastRow2+1, column=1, sticky=W)
            global bpgid
            bpgid = self.fbBrandID.get()


        self.another = Button(self.nameW, text="Add Another", command=lambda: self.addBox())
        self.another.grid(row=self.lastRow2+2+self.count, column=0, columnspan=2, pady=(20, 0))

        self.AddMoreButton = Button(self.nameW, text="Done", command=lambda: self.AddMore())
        self.AddMoreButton.grid(row=self.lastRow2+2+self.count, column=1, columnspan=2, pady=(20, 0))





    def addBox(self):

        self.moreWords[len(self.all_entries)] = StringVar()

        self.MoreVarN[len(self.all_entries)] = IntVar()
        self.MoreVarN[len(self.all_entries)].set(4)
        ent = Entry(self.nameW, textvariable=self.moreWords[len(self.all_entries)]).grid(row=self.lastRow+len(self.all_entries), column=0, sticky=W)
        self.MoreMustHave = Radiobutton(self.nameW, variable=self.MoreVarN[len(self.all_entries)], value=0).grid(row=self.lastRow+len(self.all_entries), column=1, sticky=W)
        self.MoreAltWord = Radiobutton(self.nameW, variable=self.MoreVarN[len(self.all_entries)], value=1).grid(row=self.lastRow+len(self.all_entries), column=2, sticky=W)
        self.MoreIgnore = Radiobutton(self.nameW, variable=self.MoreVarN[len(self.all_entries)], value=2).grid(row=self.lastRow+len(self.all_entries), column=3, sticky=W)
        self.MoreExcludeWord = Radiobutton(self.nameW, variable=self.MoreVarN[len(self.all_entries)], value=3).grid(row=self.lastRow+len(self.all_entries), column=4, sticky=W)
        self.MoreNone = Radiobutton(self.nameW, variable=self.MoreVarN[len(self.all_entries)], value=4).grid(row=self.lastRow+len(self.all_entries), column=5, sticky=W)

        self.all_entries.append(ent)

        self.lastRow2 = self.lastRow+len(self.all_entries)
        self.count += 1

        self.moveButtons()

    def moveButtons(self):
        self.another.grid(row=self.lastRow+(self.count)+2, column=0, columnspan=2, pady=(20, 0))
        self.AddMoreButton.grid(row=self.lastRow+(self.count)+2, column=1, columnspan=2, pady=(20, 0))
        if self.ReportType.get() == 2:
            self.FBLabel.grid(row=self.lastRow2+1, column=0, sticky=W)
            self.FBEntry.grid(row=self.lastRow2+1, column=1, sticky=W)

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
    def AllDone(self, df, part, msg):
        global t0
        t1 = time.time()
        print "start: "+datetime.datetime.fromtimestamp(t0).strftime('%Y-%m-%d %H:%M:%S')
        print "end: "+ datetime.datetime.fromtimestamp(t1).strftime('%Y-%m-%d %H:%M:%S')
        print str(t1-t0)+" seconds"



        if part == 1:
            try:
                currentStats = pd.read_csv("J:\zAutomatcherData\Part1Stats.csv", encoding='utf-8')
              #  stats=open("J:\zAutomatcherData\Part1Stats.csv",'a')
                checks = len(df[df['Robot Suggestion'].str.contains('Check')])
                #stats.write(','.join((os.getenv('username'),str(self.bizID.get()),str(getBusName(\
              #     self.bizID.get())),str(self.IndustryType.get()),str(datetime.datetime.fromtimestamp(t1).strftime('%Y-%m-%d %H:%M:%S')),str(t1-t0),\
              #          str(df.shape[0]),str(checks))))
              #  stats.close()
                currentStats.loc[len(currentStats)] = [os.getenv('username'), self.bid, getBusName(\
                    self.bid), self.IndustryType.get(), self.ReportType.get(), self.dataInput.get(), datetime.datetime.fromtimestamp(t1).strftime('%Y-%m-%d %H:%M:%S'), t1-t0,\
                        df.shape[0], checks]

                currentStats.to_csv("J:\zAutomatcherData\Part1Stats.csv", encoding='utf-8', index=False)
            except:
                pass

        self.completeWindow = Toplevel()
        self.completeWindow.protocol("WM_DELETE_WINDOW", self._delete_window)
        self.completeMsg = Label(self.completeWindow, text="time to complete: "+str(t1-t0)+"  "+msg).grid(row=1, column=1)
        self.DoneButton = Button(self.completeWindow, text="Exit", command=lambda: [root.destroy()]).grid(row=2, column=1)
    def _delete_window(self):

        try:
            root.destroy()
        except:
            pass

#Defining this globally helps being able to call it within the Tkinter class
nickNames = NameDenormalizer()
firstNames = NameDenormalizerWithOriginal()
global bpgid
global reportType
global df
df = pd.DataFrame
global checkedDF
global root
checkedDF = pd.DataFrame
global t0
t0 = time.time()
#starts Tkinter
root = Tkinter.Tk()
app = MatchingInput(root)
try:
    app.mainloop()
except:
    pass
