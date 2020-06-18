#! /usr/bin/env python3
# pip3 install gspread oauth2client pandas bs4 --user
from __future__ import print_function

import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from math import pi
from httplib2 import Http
from apiclient import discovery
from googleapiclient.discovery import build
from oauth2client import file, client, tools
from oauth2client.service_account import ServiceAccountCredentials
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import os
import pandas as pd
import time
import bs4
import smtplib  
import email.utils
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from jinja2 import Environment, PackageLoader, select_autoescape, FileSystemLoader

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('BII-Individual-074ad670ce39.json', scope)

gc = gspread.authorize(credentials)

# If new sheet, ensure it is shared with gspread member under BII-Individual Project (GCP)
spreadsheet = gc.open("Automatic Survey Generator (ASG)")

worksheet = spreadsheet.get_worksheet(0)

def send_receipt_email(to_address, name, surveys, invCode):
    file_loader = FileSystemLoader('./templates/')
    env = Environment(loader=file_loader)
    template = env.get_template('instructorReceipt.html')
    body_html = template.render(name=name, invCode = invCode, surveys = surveys)

    recipient = to_address
    admin_subject = "Summary of Your BII Invitation Code"
    name = "Berkeley Innovation Index"
    bcc = "admin@innovation-engineering.net"
    sender = "admin@innovation-engineering.net"
    
    # SES Set up
    username_smtp = "AKIAI5NYYP7NMF6M2QJA"
    password_smtp = "AqZHmg8OhOAcUu7AHJw/F+zx35znRV0zX8b7sS+4yroZ"
    # SMTP Endpoint
    host = "email-smtp.us-west-2.amazonaws.com"
    port = 587
    
    msg = MIMEMultipart('alternative')
    msg['Subject'] = admin_subject
    msg['From'] = email.utils.formataddr((name, sender))
    msg['To'] = recipient
    msg['Bcc'] = bcc
    part1 = MIMEText(body_html, 'html')
    msg.attach(part1)

    try:  
        server = smtplib.SMTP(host, port)
        server.ehlo()
        server.starttls()
        #stmplib docs recommend calling ehlo() before & after starttls()
        server.ehlo()
        server.login(username_smtp, password_smtp)
        server.send_message(msg)
        server.close()
    # Display an error message if something goes wrong.
    except Exception as e:
        response = "Failed to send instructor email"
        print(response)
        print(e)
    else:
        response = "Sent instructor email"
        print (response)

def send_ticket_email(to_address, instructor, institution, name, survey, invcode, urls):

    file_loader = FileSystemLoader('./templates/')
    env = Environment(loader=file_loader)
    template = env.get_template('ticket.html')
    body_html = template.render(instructor=instructor, institution = institution, invCode = invcode, survey = survey, typeform_ulr = urls)

    recipient = to_address
    particip_subject = survey + " Invitation Code."
    name = "Berkeley Innovation Index"
    bcc = "admin@innovation-engineering.net"
    sender = "admin@innovation-engineering.net"
    
    # SES Set Up
    username_smtp = "AKIAI5NYYP7NMF6M2QJA"
    password_smtp = "AqZHmg8OhOAcUu7AHJw/F+zx35znRV0zX8b7sS+4yroZ"
    # SMTP Endopoint
    host = "email-smtp.us-west-2.amazonaws.com"
    # SMTP Default Port (per AWS)
    port = 587
    
    msg = MIMEMultipart('alternative')
    msg['Subject'] = particip_subject
    msg['From'] = email.utils.formataddr((name, sender))
    msg['To'] = recipient
    msg['Bcc'] = bcc
    part1 = MIMEText(body_html, 'html')
    msg.attach(part1)

    try:  
        server = smtplib.SMTP(host, port)
        server.ehlo()
        server.starttls()
        #stmplib docs recommend calling ehlo() before & after starttls()
        server.ehlo()
        server.login(username_smtp, password_smtp)
        server.send_message(msg)
        server.close()
    # Display an error message if something goes wrong.
    except Exception as e:
        response = "Failed to send " + survey + " email"
        print(response)
        print(e)
    else:
        response = "Sent " + survey + " email"
        print(response)
    

def get_url(survey,invCode):
    url = ""
    if survey == "Ind - BII's Individual Innovation Index Survey": 
        url = str("https://scet-bii.typeform.com/to/wP7IGZ")
    ##elif survey == "I+O - Organizational Part of Comp Survey": 
    ##    url = str("https://scet-bii.typeform.com/to/C1CLlI?invcode="+invCode)
    ##elif survey == "Comp - Composite Survey": 
    ##    url = str("https://scet-bii.typeform.com/to/W9Hhi6?invcode="+invCode)
    ##elif survey == "BMVS - Berkeley Method Venture Score": 
    #3    url = str("https://scet-bii.typeform.com/to/XB2hKM")
    elif survey == "EQ - BII's Emotional Quotient Survey":
        url = str("https://scet-bii.typeform.com/to/CcRDBc?invcode="+invCode)
    ##elif survey == "Grit - BII's Grit Scale":
    ##    url = str("https://scet-bii.typeform.com/to/DgqSvO")
    else: 
        url = str("https://berkeleyinnovationindex.org")    
    return(url)

def generate_inv_code(name, institution, token, survey):
    invcode = ""
    # Add name
    shortened_name = "".join(name.split(" "))
    if len(shortened_name) < 4:
        invcode += shortened_name
    else:
        invcode += shortened_name[:4]
    # Add institution
    shortened_institution = "".join(institution.split(" "))
    if len(shortened_institution) < 4:
        invcode += shortened_institution
    else:
        invcode += shortened_institution[:4]
    # Add suvery
    shortened_survey = "".join(survey.split(" "))
    if len(shortened_survey) < 4:
        invcode += shortened_survey
    else:
        invcode += shortened_survey[:4]
    # Add token
    invcode += token[:4]
    return(invcode)


# First time, don't send out emails at all. 
values_list = worksheet.get_all_values()
to_add = []
labels = values_list[0]
for v in values_list[1:]:
    row = tuple(v)
    to_add.append(row)
all_records = pd.DataFrame.from_records(to_add, columns=labels)
all_records.to_csv("instructor_results.csv")
entry_times = set()
dates = pd.read_csv("instructor_results.csv")['Submitted At']
for date in dates:
    entry_times.add(date)


# Repeat forever, every 5 seconds
while True:
    values_list = worksheet.get_all_values()
    to_add = []
    labels = values_list[0] + ["Invite Code"]

    # Initialize indices
    name_index = labels.index("Please enter your full name.")
    institution_index = labels.index("Please enter the name of your institution.")
    token_index = labels.index("Token")
    date_index = labels.index("Submitted At")
    survey_index = labels.index("Select the survey you would like to conduct.")
    email_index = labels.index("What's the best email address for you?")

    for v in values_list[1:]:
        invcode = generate_inv_code(v[name_index], v[institution_index], v[token_index], v[survey_index])
        updated_row = v + [invcode]
        row = tuple(updated_row)
        to_add.append(row)
        if v[date_index] not in entry_times:
            print("Received new instructor interest form at "  + time.strftime('%Y-%m-%d %H:%M:%S', time.localtime()))
            entry_times.add(v[date_index])
            surveys = v[survey_index].split(",")
            for survey in surveys:
                survey = survey.lstrip()
                #colors = get_color(survey)
                urls = get_url(survey, invcode)
                time.sleep(1)
                send_ticket_email(v[email_index], v[name_index], v[institution_index], v[name_index], survey, invcode, urls)
                

            send_receipt_email(v[email_index], v[name_index], surveys, invcode)


    all_records = pd.DataFrame.from_records(to_add, columns=labels)
    all_records.to_csv("instructor_results.csv")
    time.sleep(5)
