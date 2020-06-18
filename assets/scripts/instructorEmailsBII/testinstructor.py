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


def bii_radial_plot():
    """
            Creates radial plot for user scores.
            @param user_vals is dictionary of user scores (float)
            @param avg_global_vals is dictionary of average scores for sheet
            @return plot_url is url string for templating

    """

        ### ------------------ Create Plot  ------------------ ###

        # ------- PART 0: Get Data
    """ df = pd.DataFrame({
            'group': ['user_vals','avg_global_vals'],
            'Trust': [user_vals.tru,avg_global_vals.tru],
            'Resilience': [user_vals.res, avg_global_vals.res],
            'Diversity': [user_vals.div, avg_global_vals.div],
            'Belief': [user_vals.bel, avg_global_vals.bel],
            'Collaboration': [user_vals.col, avg_global_vals.col],
            'Perfection': [user_vals.per, avg_global_vals.per],
            'Comfort Zone': [user_vals.cz, avg_global_vals.cz],
            'Innovation Zone': [user_vals.iz, avg_global_vals.iz]
        })"""
         
    df = pd.DataFrame({
        'group': ['global','user'],
        'Trust': [8,7.2],
        'Resilience': [6.3, 6],
        'Diversity': [7.4,4],
        'Belief': [7,5],
        'Collaboration': [7.1,5],
        'Perfection': [4.8,7.5],
        'Comfort Zone': [6.9,8],
        'Innovation Zone': [8.8,9.2]
        }) 


    # ------- PART 1: Create background
    print("** Creating User Radial Plot")
    # set global style
    sns.set(style='white', context='paper', rc={"grid.linewidth":.2})

    # number of variable
    categories=list(df)[1:]
    N = len(categories)
     
    # What will be the angle of each axis in the plot? (we divide the plot / number of variable)
    angles = [n / float(N) * 2 * pi for n in range(N)]
    angles += angles[:1]

    # control tapering of radial slice
    world_width = .225* np.pi
    user_width = .20 * np.pi
     
    # Initialise figure and axes for radial plot
    f, ax = plt.subplots(dpi=300, subplot_kw=dict(projection='polar'))


    # If you want the first axis to be on top:
    ax.set_theta_offset(pi / 2) # default location for radial scale
    ax.set_theta_direction(-1)  # increase in clockwise direction
    ax.set_rorigin(-.2)         # donut hole
    #ax.axis('off')

     
    # Draw one axe per variable 
    plt.xticks(ticks=angles, labels=categories, verticalalignment='center',
              fontsize='medium', fontfamily='sans', rotation='vertical',  color='#4E5E6E')

    # Make sure text does not overlap plot
    for x, label in zip(ax.get_xticks(), ax.get_xticklabels()):
        if np.sin(x) > 0.1:
            label.set_horizontalalignment('left')
        if np.sin(x) < -0.1:
            label.set_horizontalalignment('right')
        if np.sin(x) > 0.5:
            label.set_verticalalignment('bottom')
        if np.sin(x) > 0.5:
            label.set_verticalalignment('top')
     
    # Draw ylabels
    ax.set_rlabel_position(48)
    plt.yticks([2.5,5,7.5,10,11.3], ["2.5","5.0","7.5", "10.0"], color='#4E5E6E', size=5)
    plt.ylim(0,11.5)



    # ------- PART 2: Add plots
    # colors
    world_palette = sns.dark_palette(color='#F77F03', n_colors=N, reverse=True) #orange: #F77F03; blue-gray: #4E5E6E
    user_palette = sns.dark_palette(color='#E26130', n_colors=N,reverse=True)


     
    # user plot
    values=df.loc[0].drop('group').values.flatten().tolist()
    values += values[:1]
    # alternatively, use world_palette
    ax.bar(angles, values, align='center',width=world_width, color = '#4E5E6E', linewidth=0, alpha=.25, label="Global Average")

    # global average plot
    values=df.loc[1].drop('group').values.flatten().tolist()
    values += values[:1]
    user_bars=ax.bar(angles, values, align='center', color = '#E26130',width=user_width, linewidth=0.5, label="Me", alpha=.80);

    # Add legend
    plt.legend(loc='upper right', shadow=True, fontsize='medium',bbox_to_anchor=(1.32,0.03),borderaxespad=0.,frameon=False);

     # ensure margins capture legend
    plt.subplots_adjust(hspace=0.65, wspace=1.50)
    print('   - png saved to disk')
    f.savefig('bii_radial_plot.png')
    ### -------------------------------------------------- ###



    ### ------------------ Img to Drive  ------------------ ###
    LOGO = 'bii_radial_plot.png'

    # oauth permission requests
    SCOPES = ['https://www.googleapis.com/auth/drive']

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

    # file name, convert to drive type
    FILES = (
        (LOGO, None),
        )

    # check for image 'template'
    rsp = DRIVE.files().list(q="name='%s'" % LOGO).execute().get('files')[0]
        
    # iterate over all contents of FILES (for pushing more than one filetype)
    for filename, mimeType in FILES:
        print('** Searching for %r' % rsp['name'])
        if not rsp:
            print('   - File %r does not exist' % rsp['name'])
            print('  -- creating file')
            metadata={'name':filename}
            if mimeType:
                metadata['mimeType'] = mimeType
            create = DRIVE.files().create(body=metadata, media_body=filename).execute()
            if create:
                print(' --- Uploaded "%s" (%s) to admin google drive.' % (filename, create['mimeType']))
        else:
            print('   - File %r found' % rsp['name'])
            print('  -- replacing drive file with local')
            metadata={'name':filename}
            if mimeType:
                metadata['mimeType'] = mimeType
            update=DRIVE.files().update(fileId=rsp['id'],body=metadata, media_body=filename).execute()
            if update:
                print(' --- Uploaded %r to admin google drive.' % rsp['name'])
    ### --------------------------------------------------- ###



    ### ------------------ Img from Drive to url  ------------------ ###
    # find image file(s) in google drive and get secure URL to download it
    print('** Converting image file to url')

    # create request (not executed) and extract image file's URI
    image_url = '%s&access_token=%s' % (
        DRIVE.files().get_media(fileId=rsp['id']).uri, creds.access_token)
    print('DONE')
    return(image_url)
### ------------------------------------------------------------ ###

plot = bii_radial_plot()

def send_ticket_email(to_address, instructor, institution, name, survey, invcode, urls):

    file_loader = FileSystemLoader('./templates/')
    env = Environment(loader=file_loader)
    template = env.get_template('testticket.html')
    body_html = template.render(instructor=instructor, institution = institution, invCode = invcode, survey = survey, typeform_ulr = urls, image_url=plot )

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
