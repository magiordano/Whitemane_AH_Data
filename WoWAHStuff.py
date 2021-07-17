#!/usr/bin/env python
# coding: utf-8

# In[72]:


from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from datetime import datetime
import requests
import json
from statistics import mean
import config


#Query the Blizzard Auction House API

response = requests.get(config.url + config.api_key)
#Convert the API Response into a JSON object
newJson = json.loads(response.content)


# In[10]:


#See https://developers.google.com/sheets/api/guides/values for documentation
#Google Sheets API scope for credentials
scope = ['https://www.googleapis.com/auth/spreadsheets']
#Google Sheed ID
spreadsheet_id = '1zlK8BCuDNMvA0Dth142yxghplG7dNA3_VkgLSaeA8X8'
#Value input option can be RAW or USER_ENTERED
value_input_option = 'RAW'

#From the  Google documentation, creates the credentials if none exist already
def makecreds():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', scope)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', scope)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

#From the Google documentation, creates the 'service' variable
def makeservice(creds):
    service = build('sheets', 'v4', credentials=creds)
    return service

#Based on the documentation, this function takes the service variable, a spreadsheet ID,
# a spreadsheet range in A1 format, a Value Input Option, and the Values and updates a sheet
def update(service, spreadId, spreadRange, valueInOp, values):
    body = {
        'values': values
    }
    result = service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=spreadRange, 
        valueInputOption=value_input_option, body=body).execute()
    print('{0} cells updated.'.format(result.get('updatedCells')))   


# ### Update Date

# In[73]:


#Our sheet has the update time in the top left corner, cells A1 and A2
#Set the range in A1 format
rangeDate = 'A1:A2'

#Set the values using the datetime.now() function and turning it into a string
value = [['Last updated:'],[datetime.now().strftime('%d %b, %Y %H:%M:%S')]]

#Run the update function to update time on the sheet
update(makeservice(makecreds()), spreadsheet_id, rangeDate, value_input_option, value)


# ### Fel Iron Ore

# In[76]:


#Define 4 variables per AH item:
# An empty array to store buyout price over quantity
# A range, for the cell which should be updated in the sheet
# An item ID based on TBC Wowhead (eg. https://tbc.wowhead.com/item=23424/fel-iron-ore)
# A max value, hand picked for eliminating outlier AH postings. (eg. 20000 for 2g00s00c)
felIronOre = []
rangeFIO = 'C2'
felOreId = '23424'
felOreMax = 10000

# Iterate over the JSON from the AH API query, for each posting that matches the item ID we are looking for
# Each matching posting, divide the buyout price by the quantity to get a buyout price per item
# Store those values in an array
# Take the mean of those values, divided by 10000 to covert to WoW Gold and round to 2 decimals
# Pass into the update function defined above to update our sheet at the appropriate cell for this item
for i in newJson['auctions']:
    if str(i["item"]) == "{'id': " + felOreId + "}":
        if i['buyout']/i['quantity'] <= felOreMax:
             felIronOre.append(i['buyout']/i['quantity'])
value =[[round(mean(felIronOre)/10000, 2)]]
update(makeservice(makecreds()), spreadsheet_id, rangeFIO, value_input_option, value)


# ### Adamantite Ore

# In[77]:


adamantiteOre = []
rangeAO = 'C3'
adamOreId = '23425'
adamOreMax = 18000

for i in newJson['auctions']:
    if str(i["item"]) == "{'id': " + adamOreId + "}":
        if i['buyout']/i['quantity'] <= adamOreMax:
             adamantiteOre.append(i['buyout']/i['quantity'])
value =[[round(mean(adamantiteOre)/10000, 2)]]
update(makeservice(makecreds()), spreadsheet_id, rangeAO, value_input_option, value)


# ### Blood Garnet (C4)

# In[80]:


bloodGarnet = []
rangeBG = 'C4'
garnetId = '23077'
garnetMax = 30000

for i in newJson['auctions']:
    if str(i["item"]) == "{'id': " + garnetId + "}":
        if i['buyout']/i['quantity'] <= garnetMax:
             bloodGarnet.append(i['buyout']/i['quantity'])
value = [[round(mean(bloodGarnet)/10000, 2)*.95]]
update(makeservice(makecreds()), spreadsheet_id, rangeBG, value_input_option, value)
            


# ### Deep Peridot

# In[79]:


deepPeridot = []
rangeDP = 'C5'
peridotId = '23079'
peridotMax = 15000

for i in newJson['auctions']:
    if str(i["item"]) == "{'id': " + peridotId + "}":
        if i['buyout']/i['quantity'] <= peridotMax:
             deepPeridot.append(i['buyout']/i['quantity'])
value =[[round(mean(deepPeridot)/10000, 2)*.95]]
update(makeservice(makecreds()), spreadsheet_id, rangeDP, value_input_option, value)               


# ### Flame Spessarite

# In[81]:


flameSpessarite = []
rangeFS = 'C6'
spessariteId = '21929'
spessariteMax = 9000

for i in newJson['auctions']:
    if str(i["item"]) == "{'id': " + spessariteId + "}":
        if i['buyout']/i['quantity'] <= spessariteMax:
             flameSpessarite.append(i['buyout']/i['quantity'])
value =[[round(mean(flameSpessarite)/10000, 2)*.95]]
update(makeservice(makecreds()), spreadsheet_id, rangeFS, value_input_option, value)  


# ### Golden Draenite

# In[82]:


goldenDraenite = []
rangeGD = 'C7'
gdraeniteId = '23112'
gdraeniteMax = 17500

for i in newJson['auctions']:
    if str(i["item"]) == "{'id': " + gdraeniteId + "}":
        if i['buyout']/i['quantity'] <= gdraeniteMax:
             goldenDraenite.append(i['buyout']/i['quantity'])
value =[[round(mean(goldenDraenite)/10000, 2)*.95]]
update(makeservice(makecreds()), spreadsheet_id, rangeGD, value_input_option, value)   


# ### Shadow Draenite

# In[83]:


shadowDraenite = []
rangeSD = 'C8'
sdraeniteId = '23107'
sdraeniteMax = 10000

for i in newJson['auctions']:
    if str(i["item"]) == "{'id': " + sdraeniteId + "}":
        if i['buyout']/i['quantity'] <= sdraeniteMax:
             shadowDraenite.append(i['buyout']/i['quantity'])
value =[[round(mean(shadowDraenite)/10000, 2)*.95]]
update(makeservice(makecreds()), spreadsheet_id, rangeSD, value_input_option, value)


# ### Azure Moonstone

# In[84]:


azureMoonstone = []
rangeAM = 'C9'
moonstoneId = '23117'
moonstoneMax = 10000

for i in newJson['auctions']:
    if str(i["item"]) == "{'id': " + moonstoneId + "}":
        if i['buyout']/i['quantity'] <= moonstoneMax:
             azureMoonstone.append(i['buyout']/i['quantity'])
value =[[round(mean(azureMoonstone)/10000, 2)*.95]]
update(makeservice(makecreds()), spreadsheet_id, rangeAM, value_input_option, value)


# ### Noble Topaz

# In[85]:


nobleTopaz = []
rangeNT = 'C10'
nobleTopazId = '23439'
nobleTopazMax = 500000

for i in newJson['auctions']:
    if str(i["item"]) == "{'id': " + nobleTopazId + "}":
        if i['buyout']/i['quantity'] <= nobleTopazMax:
             nobleTopaz.append(i['buyout']/i['quantity'])
value =[[round(mean(nobleTopaz)/10000, 2)*.95]]
update(makeservice(makecreds()), spreadsheet_id, rangeNT, value_input_option, value) 


# ### Dawnstone

# In[86]:


dawnstone = []
rangeD = 'C11'
dawnstoneId = '23440'
dawnstoneMax = 500000

for i in newJson['auctions']:
    if str(i["item"]) == "{'id': " + dawnstoneId + "}":
        if i['buyout']/i['quantity'] <= dawnstoneMax:
             dawnstone.append(i['buyout']/i['quantity'])
value =[[round(mean(dawnstone)/10000, 2)*.95]]
update(makeservice(makecreds()), spreadsheet_id, rangeD, value_input_option, value) 


# ### Living Ruby

# In[87]:


livingRuby = []
rangeLR = 'C12'
livingRubyId = '23436'
livingRubyMax = 1500000

for i in newJson['auctions']:
    if str(i["item"]) == "{'id': " + livingRubyId + "}":
        if i['buyout']/i['quantity'] <= livingRubyMax:
             livingRuby.append(i['buyout']/i['quantity'])
value =[[round(mean(livingRuby)/10000, 2)*.95]]
update(makeservice(makecreds()), spreadsheet_id, rangeLR, value_input_option, value) 


# ### Nightseye

# In[88]:


nightseye = []
rangeNE = 'C13'
nightseyeId = '23441'
nightseyeMax = 200000

for i in newJson['auctions']:
    if str(i["item"]) == "{'id': " + nightseyeId + "}":
        if i['buyout']/i['quantity'] <= nightseyeMax:
             nightseye.append(i['buyout']/i['quantity'])
value =[[round(mean(nightseye)/10000, 2)*.95]]
update(makeservice(makecreds()), spreadsheet_id, rangeNE, value_input_option, value)


# ### Star of Elune

# In[89]:


starOfElune = []
rangeSE = 'C14'
eluneId = '23438'
eluneMax = 250000

for i in newJson['auctions']:
    if str(i["item"]) == "{'id': " + eluneId + "}":
        if i['buyout']/i['quantity'] <= eluneMax:
             starOfElune.append(i['buyout']/i['quantity'])
value =[[round(mean(starOfElune)/10000, 2)*.95]]
update(makeservice(makecreds()), spreadsheet_id, rangeSE, value_input_option, value)


# ### Talasite

# In[90]:


talasite = []
rangeT = 'C15'
talasiteId = '23437'
talasiteMax = 50000

for i in newJson['auctions']:
    if str(i["item"]) == "{'id': " + talasiteId + "}":
        if i['buyout']/i['quantity'] <= talasiteMax:
             talasite.append(i['buyout']/i['quantity'])
value =[[round(mean(talasite)/10000, 2)*.95]]
update(makeservice(makecreds()), spreadsheet_id, rangeT, value_input_option, value)
#print("min price per item#" + talasiteId + ": " + str(round(min(talasite) / 10000, 2)))
#print("avg price per item#" + talasiteId + ": " + str(round(mean(talasite) / 10000, 2)))

