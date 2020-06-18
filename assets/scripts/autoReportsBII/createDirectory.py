#! /usr/bin/env python3
# pip3 install -U google-api-python-client
# pip3 install oauth2client pandas bs4 --user

from __future__ import print_function
from apiclient import discovery
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
from httplib2 import Http
from oauth2client import file, client, tools


PARENT_DIR_ID = '1v4I9F_aR3QKlsz8rSvUuctgsbE-W5xgM'

####--------------------- PARTNER INFO ----------------------####
PARTNER_NAME = "Girls in Tech (Austin)"
DATE = '1/7/2019'
BLOB = ''
FOLDER = ''
DOCUMENT = ''
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

# stores access and refresh tokens, and is auto-created when the authorization flow completes for the first time.
store = file.Storage('storage.json')
CLIENT_SECRET = 'client_secret_drive.json'
    
# create objects for making api calls
creds = store.get()
if not creds or creds.invalid:
	flow = client.flow_from_clientsecrets(CLIENT_SECRET, SCOPES)
	creds = tools.run_flow(flow, store)
    
# create service endpoints to api's
HTTP = creds.authorize(Http())
# build service object
SERVICE = build('drive', 'v3', http=HTTP)
####---------------------------------------------------------####



####-------------------- CREATE DIRECTORY -------------------####
# view objects in drive
"""for f in FILES:
	f_name = f['title']
	f_type = f['mimeType'] 
	print('File Name: ' + f_name, ' is of Type ' + f_type)
	### want: application/vnd.google-apps.folder
"""
file_name = PARTNER_NAME + ' -- ' + DATE

file_metadata = {
    'name': file_name,
    'mimeType': 'application/vnd.google-apps.folder'
}
print('** Creating new contributing-partner folder')

# make request
FILE_SERVICE = SERVICE.files().create(body=file_metadata,
                                    fields='id').execute()
file_id = FILE_SERVICE.get('id')
print (' - Folder ID: %s' % file_id)

####---------------------------------------------------------####



####-------------------- MOVE DIRECTORY -------------------####
print("** Moving partner folder to 'BII Reports' parent directory")

# Retrieve the existing parents to remove
file = SERVICE.files().get(fileId=file_id,
                                 fields='parents').execute()
previous_parents = ",".join(file.get('parents'))
# Move the file to the new folder
file = SERVICE.files().update(fileId=file_id,
                                    addParents=PARENT_DIR_ID,
                                    removeParents=previous_parents,
                                    fields='id, parents').execute()
####---------------------------------------------------------####
print('DONE')