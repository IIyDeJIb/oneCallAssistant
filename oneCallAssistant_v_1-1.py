#! py -3
# Onecall locate ticket assistant
# Copy the ticket text and then run the code. The code generates
# a pastable excel row for our log and a csv file with the
# bounding box which can be loaded into google earth to check for conflicts.
#
# Peyruz Gasimov, Sep 2019

# Version history:
# 1-1 - Now a KML file is generated alongside the csv file for the bounding box.
# The KML file contains the bounding box polygon


from tkinter import messagebox as tkMessageBox

import pandas as pd
import pyperclip as py
import regex as re

ticket = py.paste()

# Define regexes
rxBb = re.compile(
		r'Latitude:\s+(\d+.\d+)\s+Longitude:\s+(-\d+.\d+)\s+Second Latitude:\s+('
		r'\d+.\d+)\s+Second Longitude:\s+(-\d+.\d+)')
rxType = re.compile(r'Type:\s+(.*)\s+Date')
rxStr = re.compile(r'Street:\s+(.*)\r')
rxInt = re.compile(r'Intersection:\s+(.*)\r')
rxNOF = re.compile(r'Nature of Work:\s+(.*)\r')
rxCity = re.compile(r'City:\s+(.*)\r')
rxWDF = re.compile(r'Work Done For:\s+(.*)\r')
rxNum = re.compile(r'Ticket Number:\s+(\d*)\s+Old')
rxDate = re.compile(r'Date:\s+(.*)\r')

# Extract ticket info
ticketInfo = {}
bbFields = rxBb.search(ticket)

try:
	ticketInfo['Type'] = rxType.search(ticket).group(1)
except AttributeError as e:
	tkMessageBox.showerror('Onecall Assistant',
						   'No ticket data found in the clipboard. Make sure you '
						   'have copy the ticket text before running the Onecall '
						   'Assitant.')
	quit()

ticketInfo['Street'] = rxStr.search(ticket).group(1)
ticketInfo['Intersection'] = rxInt.search(ticket).group(1)
ticketInfo['Nature of Work'] = rxNOF.search(ticket).group(1)
ticketInfo['City'] = rxCity.search(ticket).group(1)
ticketInfo['Work Done For'] = rxWDF.search(ticket).group(1)
ticketInfo['Ticket Number'] = rxNum.search(ticket).group(1)
ticketInfo['Date'] = rxDate.search(ticket).group(1)

# Write the ticket info into a copyable string which can be readily pasted into Excel
toPaste = ticketInfo['Date'] + '\t' + \
		  ticketInfo['Ticket Number'] + '\t' + \
		  ticketInfo['Type'] + '\t' + \
		  ticketInfo['Nature of Work'] + '\t' + \
		  ticketInfo['Work Done For'] + '\t' + \
		  ticketInfo['City'] + '\t' + \
		  ticketInfo['Street'] + '\t' + \
		  ticketInfo['Intersection']

py.copy(toPaste)

# Construct bounding box dataframe
bbDF = pd.DataFrame({
		'latitude' : [bbFields.group(1), bbFields.group(3),
					  bbFields.group(3), bbFields.group(1)],
		'longitude': [bbFields.group(2), bbFields.group(2),
					  bbFields.group(4), bbFields.group(4)]
		})

# Generate KML file for the bounding region
import simplekml

kml = simplekml.Kml()
bBox = kml.newpolygon(name='Ticket # ' + str(ticketInfo['Ticket Number']) +
						   'bBox',
					  outerboundaryis=[(bbDF.loc[ii, 'longitude'],
										bbDF.loc[ii, 'latitude'])
									   for ii in range(bbDF.shape[0])] + [
											  (bbDF.loc[0,
														'longitude'],
											   bbDF.loc[0,
														'latitude'])])
bBox.style.linestyle.color = 'ff00ffff'
bBox.style.polystyle.color = '5a00ffff'

import os

# Default saving path
kmlPath = os.path.join(r'S:', 'One call', 'bBoxes',
					   'bBox_ticket_' + ticketInfo['Ticket Number'] + '.kml')

try:
	kml.save(kmlPath)
except FileNotFoundError:
	# Save to alternative path
	kmlPath = os.path.join(r'C:', os.sep, 'Users',
						   'peyruz.gasimov', 'Work', 'Python', 'Utilities',
						    'bBoxes', 'bBox_ticket_' + ticketInfo['Ticket Number']
						   + '.kml')
	kml.save(kmlPath)

os.system('"'+kmlPath+'"' + '& exit')

# Write bounding box to csv
# bbDF.to_csv(path_or_buf=kmlPath)

tkMessageBox.showinfo('Onecall Assistant', 'Success')
