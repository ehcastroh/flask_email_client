# pip3 install -U google-api-python-client
# pip3 install oauth2client  schedule pandas bs4 --user
from __future__ import print_function

import os
import bs4
import time
import urllib
import logging
import gspread
import smtplib  
import schedule
import email.utils

import numpy as np
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

from math import pi
from httplib2 import Http
from datetime import datetime
from apiclient import discovery
from email.mime.text import MIMEText
from IPython.core.display import HTML
from googleapiclient.discovery import build
from oauth2client import file, client, tools
from email.mime.multipart import MIMEMultipart
from oauth2client.service_account import ServiceAccountCredentials
from jinja2 import Environment, PackageLoader, select_autoescape, FileSystemLoader


# Set up logging
logging.basicConfig(filename='log.txt', format='%(asctime)s %(message)s', datefmt='%m/%d/%Y %I:%M:%S %p')

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('BII-Individual-074ad670ce39.json', scope)

gc = gspread.authorize(credentials)
# If new sheet, ensure it is shared with gspread member under BII-Individual Project (GCP)
# Reference: https://github.com/burnash/gspread/issues/226
spreadsheet = gc.open("BII-Individual_V4.3 (active 8-22-18)")
worksheet = spreadsheet.get_worksheet(0)


def calculate_scores(raw_scores):
    """
    Returns a dictionary of calculated scores for score, trust, res, div, bel, perf, col, iz. 
    Given a numerical array RAW_SCORES of length 30 ordered:
    qt1,qt2,qt3,qt4,qt5,qf1,qf2,qf3,qf4,qd1,qd2,qd3,qd4,qb1,qb2,qb3,qb4,qb5,qc1,qc2,qc3,qc4,qp1,qp2,qp3,qp4,cz0,sdr,cf
    computes the following scores and returns a dictionary that maps the following categories to their corresponding score values"
        - score (total score)
        - trust (trust)
        - res (resilience)
        - div (diversity)
        - bel (strength)
        - perf (resource)
        - col (collaboration)
        - iz (iz) 
    """
    print("** computing scores")
    (qt1,qt2,qt3,qt4,qt5,qf1,qf2,qf3,qf4,qd1,qd2,qd3,qd4,qb1,qb2,qb3,qb4,qb5,qc1,qc2,qc3,qc4,qp1,qp2,qp3,qp4,cz0,sdr,cf) = raw_scores

    qt2 = 6-qt2
    qt4 = 6-qt4
    qt5 = 6-qt5
    qp1 = 6-qp1
    qp3 = 6-qp3

	# TRUST
    wt1 = 0.2+0.2/3
    wt2 = 0.1
    wt3 = 0.1
    wt4 = 0.2+0.2/3
    wt5 = 0.2+0.2/3   

    trust = (wt1*qt1 + wt2*qt2 + wt3*qt3 + wt4*qt4 + wt5*qt5)*9.0/4-5.0/4


	# RESILIENCE / FAILURE
    wf1 = 0.25/2
    wf2 = 0.25+0.25/2
    wf3 = 0.25+0.25/2
    wf4 = 0.25/2
    
    res = (wf1*qf1 + wf2*qf2 + wf3*qf3 + wf4*qf4)*9.0/4-5.0/4


	#DIVERSITY
    wd1 = 0.25/2
    wd2 = 0.25+0.25/2
    wd3 = 0.25+0.25/2 
    wd4 = 0.25/2
    
    div = (wd1*qd1+wd2*qd2+wd3*qd3+wd4*qd4)*9.0/4-5.0/4


	#BELIEF
    wb1 = 0.1
    wb2 = 0.2+0.2/3
    wb3 = 0.2+0.2/3
    wb4 = 0.1
    wb5 = 0.2+0.2/3
    
    bel = (wb1*qb1+wb2*qb2+wb3*qb3+wb4*qb4+wb5*qb5)*9.0/4-5.0/4


	#COLLABORATION
    wc1 = 0.25/2
    wc2 = 0.25/2
    wc3 = 0.25/2
    wc4 = 0.25+3*0.25/2
    
    col = (wc1*qc1 + wc2*qc2 + wc3*qc3 + wc4*qc4)*9.0/4-5.0/4


	#PERFECTION
    wp1 = 0.25/2
    wp2 = 0.25/2
    wp3 = 0.25+3*0.25/2
    wp4 = 0.25/2

    perf = (wp1*qp1 + wp2*qp2 + wp3*qp3 + wp4*qp4)*9.0/4 - 5.0/4

	#Comfort Zone and IZ
    cz = (cz0-1)*3.0+1
    sdr = sdr*9.0/4-5.0/4
    iz = 0.8*cz+.2*sdr

	#Total INNOVATION MINDSET SCORE
    score = (iz + perf + col + bel + div + res + trust)/7.0
    trust, res, div, bel, perf, col, cz, iz, sdr, score = round(trust, 2), round(res, 2), round(div, 2), round(bel, 2), round(perf, 2), round(col, 2), round(cz, 2), round(iz, 2), round(sdr, 2), round(score, 2)
    results = {"trust": trust, "res": res, "div": div, "bel": bel, "perf": perf, "col": col, "cz": cz, "iz": iz, "sdr": sdr, "score": score}
    print(" - scores done")
    return (results)
    ### --------------------------------------------------- ###


def getMAandR(cz_0, sdr_0):
    """
    Given CZ_0 and SDR_0 (2 int values), returns a dictionary for CZ and list for SDR.
    The dictionary for CZ will include 2 keys with corresponding strings:
        - leaning: description of mindset leaning
        - preference: potential recommendations due to leaning
    The list for SDR will include 1 or more strings on potential recommendations to increase overall effectiveness.
    """

    CZ, SDR = {}, {}
    if cz_0 == 4:
        CZ = { "leaning" : "MINDSET LEANS STRONGLY towards INNOVATION.",
        "preference" : "If you have interest in operational innovation, you should pre-analyze situations and focus more on risk mitigation."}
    elif cz_0 == 3:
        CZ = { "leaning" : "MINDSET covers both operations and innovation, but LEANS towards INNOVATION.",
        "preference" : "If you have interest in operational innovation and precision, you should pre-analyze situations and focus more on risk mitigation."}
    elif cz_0 == 2:
        CZ = { "leaning" : "MINDSET covers both operations and innovation but LEANS towards OPERATIONAL INNOVATION.",
        "preference" : "If you are interested in innovation or entrepreneurship, you should try to grow by increasing your comfort even when you are in areas where you are not knowledgeable. Also look into techniques that reduce fears."}
    elif cz_0 == 1:
        CZ = { "leaning" : "MINDSET LEANS STRONGLY towards OPERATIONAL INNOVATION and PRECISION.",
        "preference" : "If you are interested in innovation or entrepreneurship, you should increase your comfort when you are in uncertain situations. Also look into techniques that reduce fears."}

    # SDR Evaluation Logic
    if sdr_0 > 3:
        SDR = ["You are already an effective innovator. To continue to increase your overall effectiveness,use your talents to develop people that can then develop others."]
    if sdr_0 == 3:
        SDR = ["To increase your overall effectiveness, you should seek to increase your follow-through in your commitments.", "Your statements should be valued more like promises: Do what you say, and say what you mean."]
    if sdr_0 < 3:
        SDR = ["To increase your overall effectiveness, you should seek to increase your follow-through in your commitments.", "Your statements should be valued more like promises: Do what you say, and say what you mean.", "You may want to keep lists so that you are more aware of your obligations."]
    return(CZ, SDR)
    ### --------------------------------------------------- ###


def create_bii_radial_plot(scores, token):
    """
    Creates radial plot for user scores.
        @param user_vals is dictionary of user scores (float)
        @param avg_global_vals is dictionary of average scores for sheet
        @return plot_url is url string for templating

    """

    # average global scores are manually computed, 
    # need to come up with a solution to automatically track and update global averages.
    print(" - creating radial plot")
    avg_scores = {
        'trust': round(5.355890, 2),
        'res'  : round(8.105351, 2),
        'div'  : round(7.720042, 2),
        'col'  : round(7.049783, 2),
        'bel'  : round(7.666142, 2), 
        'perf' : round(5.520001, 2),
        'iz'   : round(4.221178, 2)
    }


    # ------- PART 0: Get Data
    df = pd.DataFrame({
        'group'          : ['avg_global_vals','user_vals'],
        'Trust'          : [avg_scores["trust"],scores["trust"]],
        'Resilience'     : [avg_scores["res"],scores["res"]],
        'Diversity'      : [avg_scores["div"],scores["div"]],
        'Collaboration'  : [avg_scores["col"],scores["col"]],
        'Belief'         : [avg_scores["bel"],scores["bel"]],    
        'Perfection'     : [avg_scores["perf"],scores["perf"]],
        'Innovation Zone': [avg_scores["iz"],scores["iz"]]
        })


    # ------- PART 1: Create background
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
    plt.legend(loc='upper right', shadow=True, fontsize='medium',bbox_to_anchor=(1.34,-0.03),borderaxespad=0.,frameon=False);

     # ensure margins capture legend
    plt.subplots_adjust(hspace=0.65, wspace=1.50)
    PLOT = 'bii_radial_plot'+token[-5:]+'.jpg'
    f.savefig(PLOT)


    #------------------ PART 3: Img to url
    try:
        import argparse
        flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
    except ImportError:
        flags = None    
    
    # oauth permission requests
    SCOPES = ['https://www.googleapis.com/auth/drive']

    # stores access and refresh tokens, and is auto-created when the authorization flow completes for the first time.
    store = file.Storage('storage.json')
        
    # create objects for making api calls
    creds = store.get()
    if not creds or creds.invalid:
        flow = client.flow_from_clientsecrets('client_secret_drive.json', scope=SCOPES)
        creds = tools.run_flow(flow, store, flags)\
            if flags else tools.run(flow,store)

    # create service endpoints to api's
    HTTP = creds.authorize(Http())
    DRIVE = build('drive', 'v3', http=HTTP)

    # file name, convert to drive type
    FILES = (
        (PLOT, 'image/jpeg'),
        )
        
    # iterate over all contents of FILES (for pushing more than one filetype)
    for filename, mimeType in FILES:
        metadata={'name':filename}
        if mimeType:
            metadata['mimeType'] = mimeType
        res = DRIVE.files().create(body=metadata, media_body=filename).execute()
        if res:
            print(' - uploaded "%s" (%s)' % (filename, res['mimeType']))

        # check for image uploaded
        rsp = DRIVE.files().list(q="name='%s'" % filename).execute().get('files')[0]
        #------------------ Img from Drive to url
        if rsp:
            print(" - converting plot to url")
            #get url
            image_url = '%s&access_token=%s' % (DRIVE.files().get_media(fileId=rsp['id']).uri, creds.access_token)
            #print(image_url)
            print(" - clearing temporary remote image")
            DRIVE.files().delete(fileId=rsp['id']) # remove file so that it does not take up space
        return(image_url)


### --------------------------------------------------- ###


def send_email(to_address, scores, token):
    """
    Sends individual survey email TO_ADDRESS based on SCORES. 
    Templates the HTML file sent using Jinja2 and uses AWS SES to send email. 
    """
    print("** creating email object")

    mindset, recommendations = getMAandR(scores['cz'], scores['sdr']) 
    plots = create_bii_radial_plot(scores, token)
    #plots = get_bii_radial_plot(tokenized)
    #print(plots)
    plots = '<img src="'+plots+'" alt=""  width="90%" align="center" style="border: 5; display: inline;margin-top:0px;margin-bottom:0px;">'
    #print(plots)

    # Jinja2 templating
    file_loader = FileSystemLoader('../Response_Participant/Templates/')
    env = Environment(loader=file_loader)
    template = env.get_template('biiUserResponseEmail.html')

    
    body_html = template.render(
        score=scores["score"], 
        trust = scores["trust"], 
        resilience = scores["res"], 
        diversity = scores["div"], 
        strength = scores["bel"], 
        resource = scores["perf"], 
        collaboration = scores["col"], 
        iz = scores["iz"], 
        mindset = mindset, 
        recommendations = recommendations, 
        image_url=plots
        )

    recipient = to_address
    email_subject = 'Your Personal Innovation Mindset Evaluation Results'
    admin = "admin@innovation-engineering.net"
    sender_name = "Berkeley Innovation Index"; 
    bcc = "admin@innovation-engineering.net"
    sender_address = "admin@innovation-engineering.net"
    
    # SES Set up
    username_smtp = "AKIAI5NYYP7NMF6M2QJA"
    password_smtp = "AqZHmg8OhOAcUu7AHJw/F+zx35znRV0zX8b7sS+4yroZ"
    # SMTP Endpoint
    host = "email-smtp.us-west-2.amazonaws.com"
    port = 587
    
    msg = MIMEMultipart('alternative')
    msg['Subject'] = email_subject
    msg['From'] = email.utils.formataddr((sender_name, sender_address))
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
        logging.exception("Failed to send individual email to %s", recipient)
    else:
        print(" - email done")
        print("** Sent individual email to %s", recipient)
        #logging.info("** Sent individual email to %s", recipient)
    ### --------------------------------------------------- ###


"""def calculate_global_avg(spreadsheet):
    df = spreadsheet.get_worksheet(1)
    df = df.sample(n = 500, random_state = None)        # randomly select a unique subset for faster computation

    # clean df
    df = [col.lower() for col in df.columns]            # column names to lowercase
    df = remove_test(df)                                # remove rows that contain developer tests
    df = df.drop_duplicates()                           # remove duplicate rows
    df = df.drop(columns=['invcode', 'email'])          # remove unwanted columns
    df = df[df['trust'].apply(lambda x: x.isnumeric())] # remove rows that contain nonnumeric elements

    # compute column averages using auxillary function 
    trust = round_avg(df, 'trust')
    res = round_avg(df, 'resilience')
    div = round_avg(df, 'diversity')
    bel = round_avg(df, 'belief')
    perf = round_avg(df, 'perfection')
    col = round_avg(df, 'collaboration')
    cz =  round_avg(df, 'comfort zone')
    iz = round_avg(df, 'innovation zone')
    score = round_avg(df, 'score')
    avg_results = {"trust": trust, "res": res, "div": div, "bel": bel, "perf": perf, "col": col, "cz": cz, "iz": iz, "sdr": sdr, "score": score}
    return(avg_results)"""

def round_avg(df, col_name):
    avg=df.loc[:, col_name].mean()
    return(round(avg, 2))

def remove_test(df):
    """
    @param df is dataframe to be processed
    @return cleaned dataframe
    Function removes whitespace, sets elements to lowercase, and remove invalid invitation codes
    """
    # set to lowercase, remove leading/trailing whitespace, and replace middle whitespace with -
    df.invitation_code = df.invitation_code.str.lower().str.strip().str.replace(' ', '-')
    # remove test (developer testing rows), keep rows where invitation_code is null
    df = df[(~df.invitation_code.str.contains('test', na=False))]    
    return(df)
    ### --------------------------------------------------- ###



###------------------ PART 4: Main Control ------------------ ###
# First time, don't send out emails at all. 
values_list = worksheet.get_all_values()
headers = values_list[0]
all_records = pd.DataFrame(values_list[1:], columns=headers)
all_records.to_csv("bii_individual.csv")

# sent_tokens = set()
# tokens = pd.read_csv("bii_individual.csv")['Token']
# for token in tokens:
#     sent_tokens.add(token)
# print("Added " +  str(len(tokens)) + " to tokens list.")
token_index = values_list[0].index("Token")
email_index = values_list[0].index("EMAIL")
curr_index = len(values_list)

# Repeat forever, every 3 seconds
while True:
    # values_list = worksheet.get_all_values()
    # for row in values_list[1:]:
    #     token = row[token_index]
    #     if token not in sent_tokens:
    #         sent_tokens.add(token)
    #         raw_scores = [int(value) for value in row[1:30]]
    #         scores = calculate_scores(raw_scores)
    #         print(scores)
    #         send_email(row[email_index], scores)
    # headers = values_list[0]
    # all_records = pd.DataFrame(values_list[1:], columns=headers)
    # all_records.to_csv("bii_individual.csv")

    try:
        time.sleep(1) 
        row = worksheet.row_values(curr_index+1)
        print("Received new survey response from %s", str(row[email_index]))
        #logging.info("Received new survey response from %s", str(row[email_index]))
        curr_index += 1
        raw_scores = [int(value) for value in row[1:30]]
        scores = calculate_scores(raw_scores)
        send_email(row[email_index], scores, row[token_index])

        # Write updated records to disk
        headers = values_list[0]
        all_records = pd.DataFrame(values_list[1:], columns=headers)
        all_records.to_csv("bii_individual.csv")
        print("DONE")
    except Exception as e:
        pass
        
    time.sleep(2)
###---------------------------------------------------------- ###