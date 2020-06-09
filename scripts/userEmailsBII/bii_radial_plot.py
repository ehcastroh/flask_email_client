
#! /usr/bin/env python3
# pip3 install -U google-api-python-client
# pip3 install oauth2client pandas bs4 --user
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

#def bii_radial_plot(user_vals, avg_global_vals):

def bii_radial_plot():
	"""
			Creates radial plot for user scores.
			@param user_vals is dictionary of user scores (float)
			@param avg_global_vals is dictionary of average scores for sheet
			@return plot_url is url string for templating

	"""

		### ------------------ Create Plot  ------------------ ###

		# ------- PART 0: Get Data
	"""	df = pd.DataFrame({
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


#return(results)