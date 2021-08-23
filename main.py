from pprint import pprint

import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials


CREDENTIALS_FILE = 'credentials.json'
GetSheets_id = '1Bu8Hh2jJTXhxtTEGsrePacX4LoQ90kvGXl9tMQ0EO6c'

credentials = ServiceAccountCredentials.from_json_keyfile_name(
	CREDENTIALS_FILE,
	['https://www.googleapis.com/auth/spreadsheets',
	'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth)

sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetId=GetSheets_id,
                            range="K2:AB4000").execute()
values = result.get('values', [])
branches = {}
for row in values:
	branches[(row[0])] = {}
for row in values:
	branches[(row[0])].update({row[1]:{row[-1]:[row[2:-2]]}})

# 	print(row)
pprint(branches)