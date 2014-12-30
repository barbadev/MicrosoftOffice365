# o365.py
# ----------------------------------------------
# Test Office 365 platform API
# Source : http://blogs.msdn.com/b/exchangedev/archive/2014/06/19/zapier-s-office-365-api-journey.aspx
# ----------------------------------------------
# 30/01/2014 : Philippe GAUVIN

import requests # install http://docs.python-requests.org/

EMAIL = raw_input("Login : ") 
PASSWORD =  raw_input("Password : ")

# Get a list of all calendars.
calendars_url = 'https://outlook.office365.com/EWS/OData/Me/Calendars'

print('Calendars:')
response = requests.get(calendars_url, auth=(EMAIL, PASSWORD))
for calendar in response.json()['value']:
  print(calendar['Name'])
print('')

# Get a list of all events.
events_url = 'https://outlook.office365.com/EWS/OData/Me/Events'

print('Events:')
response = requests.get(events_url, auth=(EMAIL, PASSWORD))
for event in response.json()['value']:
  print(event['Subject'])
print('')

# Get a list of all contacts.
contacts_url = 'https://outlook.office365.com/EWS/OData/Me/Contacts'

print('Contacts:')
response = requests.get(contacts_url, auth=(EMAIL, PASSWORD))
for contact in response.json()['value']:
  print(contact['EmailAddresses'])
print('')

# Get a list of the 10 most recently created contacts.
contacts_url = 'https://outlook.office365.com/EWS/OData/Me/Contacts'
params = {'$orderby': 'DateTimeCreated desc', '$top': '10'}

print('10 Most Recent Contacts:')
response = requests.get(contacts_url, params=params, auth=(EMAIL, PASSWORD))
for contact in response.json()['value']:
  print(contact['EmailAddresses'])
print('')  

# Get a list of the 25 most recently received messages.
messages_url = 'https://outlook.office365.com/EWS/OData/Me/Messages'
params = {'$orderby': 'DateTimeReceived desc', '$top': '25', '$select': 'Subject,From'}

print('25 Most Recent Messages:')
response = requests.get(messages_url, params=params, auth=(EMAIL, PASSWORD))
for message in response.json()['value']:
  print("%s : %s " %(message['From']['EmailAddress']['Address'],message['Subject']))
print('')