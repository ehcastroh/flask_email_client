#! /usr/bin/env python3
# pip3 install -U google-api-python-client
# pip3 install oauth2client pandas bs4 --user

from __future__ import print_function
from apiclient import discovery
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
from httplib2 import Http
from oauth2client import file, client, tools

import uuid
import gspread


####--------------------- PARTNER INFO ----------------------####
#LOGO = 'bii_radial_plot.png'
TEMPLATE = 'Automatic BII Report (Template)'
#PARTNER_NAME = "Girls in Tech (Austin)"
#DATE = '1/7/2019'
####---------------------------------------------------------####



####--------------------- CREATE ACCESS ---------------------####
# Create UUID for future use
#gen_uuid = lambda : str(uuid.uuid4())  # get random UUID string

# oauth permission requests
SCOPES = [
	'https://spreadsheets.google.com/feeds',
	'https://www.googleapis.com/auth/drive', 
	'https://www.googleapis.com/auth/presentations'
]

# if needed, the presentation ID is found in url. e.g. https://docs.google.com/presentation/d/1MxTCRkYl3SN-qQvnKk-iMIyOG97Fe6d5tZ0C7aS4dVI/edit
PRESENTATION_ID = '1MxTCRkYl3SN-qQvnKk-iMIyOG97Fe6d5tZ0C7aS4dVI'

# stores access and refresh tokens, and is auto-created when the authorization flow completes for the first time.
store = file.Storage('storage.json')
    
# create objects for making api calls
creds = store.get()
if not creds or creds.invalid:
	flow = client.flow_from_clientsecrets('client_secret_drive.json', SCOPES)
	creds = tools.run_flow(flow, store)
    
# create service endpoints to api's
HTTP = creds.authorize(Http())
DRIVE = discovery.build('drive', 'v3', http=HTTP)
SLIDES = discovery.build('slides', 'v1', http=HTTP)

# find template file, copy it, and save file id for referencing
rsp = DRIVE.files().list(q="name='%s'" % TEMPLATE).execute().get('files')[0]
DATA = {'name': 'BII Report - ' + PARTNER_NAME}
print('** Copying template %r as %r' % (rsp['name'], DATA['name']))
DECK_ID = DRIVE.files().copy(body=DATA, fileId=rsp['id']).execute().get('id')
####---------------------------------------------------------####




####----------------- MODIFY PRESENTATION -------------------####
# find page elements (rectangle placeholders) to be replaced with images
print('** Grabbing image placeholders')
slide = SLIDES.presentations().get(presentationId=DECK_ID, 
	fields='slides').execute().get('slides', [])[0]
obj = None

# Explore slide elements (nested ordering: closest is outputed last)
"""for obj in slide ['pageElements']:
	print('NEW OBJECT: ')
	print("  " + obj['objectId'])
"""

# stop when desired shape found
for obj in slide['pageElements']:
	if obj['objectId'] == 'g4c31c59c3e_0_3':
		print(' - placeholder found')
		break
else:
	print(' - No placeholders found')


# find image file(s) in google drive and get secure URL to download it
print('** Searching for image file')
rsp = DRIVE.files().list(q="name='%s'" % LOGO).execute().get('files')[0]
print(' - Found image %r' % rsp['name'])
 
# create request (not executed) and extract image file's URI
img_url = '%s&access_token=%s' % (
	DRIVE.files().get_media(fileId=rsp['id']).uri, creds.access_token)
print(img_url)

# place requests to API's
print('** Replacing placeholders for text and images')
reqs = [
	# replace text
	{'replaceAllText': {
		'containsText': {'text': '{{PARTNER_NAME}}'}, 
		'replaceText':PARTNER_NAME
	}},
	{'replaceAllText':{
		'containsText': {'text': '{{DATE}}'},
		'replaceText': DATE
	}},
	# get, resize image, and place it on placeholder
	{'createImage': {
		'url': img_url,
		'elementProperties': {
			'pageObjectId': slide['objectId'],
			'size': obj['size'],
			'transform': obj['transform'],
		}
	}},
	# delete placeholder
	{'deleteObject': {'objectId': obj['objectId']}},
]

# execute requests and notify user
SLIDES.presentations().batchUpdate(body={'requests':reqs},
		presentationId=DECK_ID, fields='').execute()
print('DONE')
####---------------------------------------------------------####
