#! /usr/bin/env python3
# pip3 install gspread oauth2client pandas bs4 --user

from __future__ import print_function
import os
import bs4
import time
import gspread
import smtplib  
import email.utils
import numpy as np
import pandas as pd
import pprint as pp
from httplib2 import Http
from apiclient import discovery
from email.mime.text import MIMEText
from googleapiclient.discovery import build
from oauth2client import file, client, tools
from email.mime.multipart import MIMEMultipart
from oauth2client.service_account import ServiceAccountCredentials
from jinja2 import Environment, PackageLoader, select_autoescape, FileSystemLoader


####--------------------- GENERAL VARS ----------------------####
SHEETDATA = 'BII-EQ (Beta)'
####---------------------------------------------------------####



####--------------------- CREATE ACCESS ---------------------####
# oauth permission requests
SCOPES = [
	'https://spreadsheets.google.com/feeds',
	'https://www.googleapis.com/auth/drive'
]

# stores access and refresh tokens, and is auto-created when the authorization flow completes for the first time.
store = file.Storage('storage.json')
    
# create objects for making api calls
creds = store.get()
if not creds or creds.invalid:
	flow = client.flow_from_clientsecrets('client_secret_eq.json', SCOPES)
	creds = tools.run_flow(flow, store)
    
# create service endpoints to api's
HTTP = creds.authorize(Http())
DRIVE = discovery.build('drive', 'v3', http=HTTP)
#SPREADSHEET = discovery.build('sheets', 'v4', http=HTTP)
"""
# get sheet id
print('** Fetching BII-EQ Sheet Data')
SHEET = DRIVE.files().list(q="name='%s'" % SHEETDATA).execute().get('files')[0]
sheetID = SHEET['id']

# get sheet
orders = SPREADSHEET.spreadsheets().values().get(range='Responses',
		spreadsheetId = sheetID).execute().get('values')
pp.pprint(type(orders))
"""
CREDENTIALS = ServiceAccountCredentials.from_json_keyfile_name('service_BII-EQ.json', SCOPES)

SPREADSHEET = gspread.authorize(CREDENTIALS)

# If new sheet, ensure it is shared with gspread email in 'service_BII-EQ.json' file
SHEET = SPREADSHEET.open(SHEETDATA)

WORKSHEET = SHEET.get_worksheet(0)

####---------------------------------------------------------####



####--------------------- COMPUTE SCORES --------------------####
def eqScore(v, sa, sr, em, mo, sk):
	print('inside eqScore')
	# Self Awareness
	sa_sc = [sa['SA1'], sa['SA2'], sa['SA3']]
	sa_sc = ['SA1', 'SA2', 'SA3'
	[x for x in mylist if len(x)==3]]
	sa_sf = [sa['SA4'], sa['SA5']]
	sa_se = [sa['SA6'], sa['SA7'], sa['SA8']]

    # Self Regulation
	sr_si = [sr['SR1'], sr['SR2']]
	sr_sc = [sr['SR3'], sr['SR4'], sr['SR5']]
	sr_se = [sr['SR6'], sr['SR7'], sr['SR8']]

	# Empathy
	em_el = [em['EMP1'], em['EMP2'], em['EMP3'], em['EMP4'], em['EMP5'], em['EMP6'], em['EMP7']]
	em_ec = [em['EMP8'], em['EMP9'], em['EMP10']]
	em_ea = [em['EMP11'], em['EMP12'], em['EMP13'], em['EMP14'], em['EMP15']]

	# Motivation 
	mo_im = [mo['MOT1'], mo['MOT2'], mo['MOT3'], mo['MOT4'], mo['MOT5'], mo['MOT6']]
	mo_em = [mo['MOT7'], mo['MOT8'], mo['MOT9'], mo['MOT10'], mo['MOT11']]

	# Social Skills
	sk_sw = [sk['SK1'], sk['SK2'], sk['SK3'], sk['SK4'], sk['SK5']]
	sk_sc = [sk['SK6'], sk['SK7'], sk['SK8']]

	### GET AVERAGES ###
	print('** Processing self_awareness')
	self_awareness = computeAvg([computeAvg(sa_sc), computeAvg(sa_sf), computeAvg(sa_se)])
	print('** Processing self_regulation')
	self_regulation = computeAvg([computeAvg(sr_si), computeAvg(sr_sc), computeAvg(sr_se)])
	print('** Processing empathy')
	empathy = computeAvg([computeAvg(em_el), computeAvg(em_ec), computeAvg(em_ea)])
	print('** Processing motivation')
	motivation = computeAvg([computeAvg(mo_im), computeAvg(mo_em)])
	print('** Processing social_skill')
	social_skill = computeAvg([computeAvg(sk_sw), computeAvg(sk_sc)])
	print('** Processing eq_score')
	eq_score = computeAvg([self_awareness, self_regulation, empathy, motivation, social_skill])
	print(' - Averages Complete')

	scores = {
		'sa': self_awareness, 
		'sr': self_regulation,
		'em': empathy,
		'mo': motivation,
		'sk': social_skill,
		'eq': eq_score
		}

	return(scores)

def indexValue(v, val):


def computeAvg(ls):
	"""
	function takes a list and returns a single float
	"""
	return(np.mean(ls))
####---------------------------------------------------------####



####------------------------ CONTROL ------------------------####
# First time, don't execute process. 
values_list = WORKSHEET.get_all_values()
to_add = []
labels = values_list[0]

# take entire list and append each row as tupple to new list
for v in values_list[1:]:
    row = tuple(v)
    to_add.append(row)

all_records = pd.DataFrame.from_records(to_add, columns=labels)

# store locally, and creat time stamp for control purposes 
all_records.to_csv("participant_results.csv")
# ensure each time stamp is unique
entry_times = set()
# extract timestamp
dates = pd.read_csv("participant_results.csv")['Submitted At']
for date in dates:
	entry_times.add(date)
    

# Execute process for new rows of data - Repeat forever, every 5 seconds
while True:
    values_list = WORKSHEET.get_all_values()
    to_add = []
    labels = values_list[0] 

    # Initialize indices
    invitation_code_index = labels.index("INVCODE") 
    self_awareness = {
    	'SA1' : labels.index("SA1"), 'SA2' : labels.index("SA2"),
    	'SA3' : labels.index("SA3"), 'SA4' : labels.index("SA4"),
    	'SA5' : labels.index("SA5"), 'SA6' : labels.index("SA6"),
    	'SA7' : labels.index("SA7"), 'SA8' : labels.index("SA8") 
    }
    self_regulation = {
    	'SR1' : labels.index("SR1"), 'SR2' : labels.index("SR2"),
    	'SR3' : labels.index("SR3"), 'SR4' : labels.index("SR4"),
    	'SR5' : labels.index("SR5"), 'SR6' : labels.index("SR6"),
    	'SR7' : labels.index("SR7"), 'SR8' : labels.index("SR8")
    }
    empathy = {
		'EMP1' : labels.index("EMP1"), 'EMP2' : labels.index("EMP2"),
    	'EMP3' : labels.index("EMP3"), 'EMP4' : labels.index("EMP4"),
    	'EMP5' : labels.index("EMP5"), 'EMP6' : labels.index("EMP6"),
    	'EMP7' : labels.index("EMP7"), 'EMP8' : labels.index("EMP8"),
    	'EMP9' : labels.index("EMP9"), 'EMP10' : labels.index("EMP10"),
    	'EMP11' : labels.index("EMP11"), 'EMP12' : labels.index("EMP12"),
    	'EMP13' : labels.index("EMP13"), 'EMP14' : labels.index("EMP14"),
    	'EMP15' : labels.index("EMP15")
    }
    motivation = {
		'MOT1' : labels.index("MOT1"), 'MOT2' : labels.index("MOT2"),
    	'MOT3' : labels.index("MOT3"), 'MOT4' : labels.index("MOT4"),
    	'MOT5' : labels.index("MOT5"), 'MOT6' : labels.index("MOT6"),
    	'MOT7' : labels.index("MOT7"), 'MOT8' : labels.index("MOT8"),
    	'MOT9' : labels.index("MOT9"), 'MOT10' : labels.index("MOT10"),
    	'MOT11' : labels.index("MOT11")
    }
    social_skills = {
    	'SK1' : labels.index("SK1"), 'SK2' : labels.index("SK2"),
    	'SK3' : labels.index("SK3"), 'SK4' : labels.index("SK4"),
    	'SK5' : labels.index("SK5"), 'SK6' : labels.index("SK6"),
    	'SK7' : labels.index("SK7"), 'SK8' : labels.index("SK8") 
    }
    education_index = labels.index("EDUCATION")
    study_index = labels.index("STUDY")
    stage_index = labels.index("STAGE")
    affiliation_index = labels.index("AFFILIATION")
    responsibility_index = labels.index("RESPONSIBILITY")
    area_index = labels.index("AREA")
    employees_index = labels.index("EMPLOYEES")
    gender_index = labels.index("GENDER")
    age_index = labels.index("AGE")
    country_index = labels.index("COUNTRY")
    postal_index = labels.index("POSTAL")
    email_index = labels.index("EMAIL") 
    date_index = labels.index("Submitted At")
    token_index = labels.index("Token")

    for v in values_list[1:]:
        print('inside for loop')
        scores = eqScore(v, self_awareness, self_regulation, empathy, motivation, social_skills])
        updated_row = v + [scores]
        row = tuple(updated_row)
        to_add.append(row)
      	# ensures loop does not run continuously every 5 seconds
        if v[date_index] not in entry_times:
            print("Received new EQ Email Request at "  + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
            entry_times.add(v[date_index])
            #surveys = v[survey_index].split(",")
            #for survey in surveys:
            #    survey = survey.lstrip()
            #    colors = get_color(survey)
            #    urls = get_url(survey, invcode)
            #    time.sleep(1)
            #    send_survey_email(v[email_index], v[name_index], v[institution_index], v[name_index], survey, invcode, urls)
                

            #send_instructor_email(v[email_index], v[name_index], surveys, invcode)


    all_records = pd.DataFrame.from_records(to_add, columns=labels)
    all_records.to_csv("participant_results.csv")
    time.sleep(5)
####---------------------------------------------------------####