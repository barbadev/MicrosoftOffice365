# sendmail.py
# ----------------------------------------------
# Send an email from Office 365 platform
# Just to test
# ----------------------------------------------
# 30/01/2014 : Philippe GAUVIN

import requests
import json

EMAIL = raw_input("Login : ") # or your custom domain
PASSWORD =  raw_input("Password : ") # the same one you use to log in with

sendmail_url = 'https://outlook.office365.com/EWS/OData/me/sendmail'
headers = {'content-type': 'application/json'}
message = {"Message":{"Subject":"Test Python", "Body":{"ContentType":"Text", "Content":"Test depuis Python"}, "ToRecipients":[{"EmailAddress":{"Address" : "email@acme.com"}}]}, "SaveToSentItems":"false"}

print json.dumps(message)
r = requests.post(sendmail_url, data=json.dumps(message), headers=headers, auth=(EMAIL, PASSWORD))
print(r.text)