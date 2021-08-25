from pprint import pprint

import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials


def get_list_files(drive):
    mimeType = 'application/vnd.google-apps.spreadsheet'
    response = drive.files().list(q=f"mimeType='{mimeType}'",
                                          spaces='drive',
                                          fields='nextPageToken, files(id, name)',
                                          pageToken=None).execute()
    # print(response)


def create_sheet(service, spreadsheet_id, room):
    new_sheet = service.spreadsheets().batchUpdate(
        spreadsheetId = spreadsheet_id,
        body = 
    {
      "requests": [
        {
          "addSheet": {
            "properties": {
              "title": f"{room}",
              "gridProperties": {
                "rowCount": 20,
                "columnCount": 10
              }
            }
          }
        }
      ]
    }).execute()


def format_sheet(service, spreadsheetId, sheetId):
    results = service.spreadsheets().batchUpdate(spreadsheetId = spreadsheetId, body = {
      "requests": [



    {'mergeCells': {
            'mergeType': 'MERGE_COLUMNS',
            'range': {
                'endColumnIndex': 1,
                'endRowIndex': 13,
                'sheetId': sheetId,
                'startColumnIndex': 0,
                'startRowIndex': 2
            }
        }},




            # Задать ширину столбца A: 100 пикселей
            {
              "updateDimensionProperties": {
                "range": {
                  "sheetId": sheetId,
                  "dimension": "COLUMNS",  # Задаем ширину колонки
                  "startIndex": 0, # Нумерация начинается с нуля
                  "endIndex": 1 # Со столбца номер startIndex по endIndex - 1 (endIndex не входит!)
                },
                "properties": {
                  "pixelSize": 120 # Ширина в пикселях
                },
                "fields": "pixelSize" # Указываем, что нужно использовать параметр pixelSize  
              }
            },

            # Задать ширину столбцов B : 100 пикселей
            {
              "updateDimensionProperties": {
                "range": {
                  "sheetId": sheetId,
                  "dimension": "COLUMNS",
                  "startIndex": 1,
                  "endIndex": 2
                },
                "properties": {
                  "pixelSize": 100
                },
                "fields": "pixelSize"
              }
            },


            # Задать ширину столбца D: 200 пикселей
            {
              "updateDimensionProperties": {
                "range": {
                  "sheetId": sheetId,
                  "dimension": "COLUMNS",
                  "startIndex": 3,
                  "endIndex": 8
                },
                "properties": {
                  "pixelSize": 140
                },
                "fields": "pixelSize"
              }
            },





    {
          "repeatCell": 
          {
            "cell": 
            {
              "userEnteredFormat": 
              {
                "horizontalAlignment": 'CENTER',
                "verticalAlignment": 'MIDDLE',
                
                "textFormat":
                 {
                   "bold": True,
                   "fontSize": 36
                 }
              }
            },
            "range": 
            {
              "sheetId": sheetId,
              "startRowIndex": 2,
              "endRowIndex": 16,
              "startColumnIndex": 0,
              "endColumnIndex": 1
            },
            "fields": "userEnteredFormat"
          }
        },






    {
          "repeatCell": 
          {
            "cell": 
            {
              "userEnteredFormat": 
              {
                "horizontalAlignment": 'CENTER',
                "verticalAlignment": 'MIDDLE',
                
                "textFormat":
                 {
                   "bold": True,
                   "fontSize": 10
                 }
              }
            },
            "range": 
            {
              "sheetId": sheetId,
              "startRowIndex": 1,
              "endRowIndex": 2,
              "startColumnIndex": 0,
              "endColumnIndex": 9
            },
            "fields": "userEnteredFormat"
          }
        },




    {'updateBorders': {'range': {'sheetId': sheetId,
                             'startRowIndex': 1,
                             'endRowIndex': 13,
                             'startColumnIndex': 0,
                             'endColumnIndex': 9},
                   'bottom': {  
                   # Задаем стиль для верхней границы
                              'style': 'SOLID', # Сплошная линия
                              'width': 1,       # Шириной 1 пиксель
                              'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}}, # Черный цвет
                   'top': { 
                   # Задаем стиль для нижней границы
                              'style': 'SOLID',
                              'width': 1,
                              'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}},
                   'left': { # Задаем стиль для левой границы
                              'style': 'SOLID',
                              'width': 1,
                              'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}},
                   'right': { 
                   # Задаем стиль для правой границы
                              'style': 'SOLID',
                              'width': 1,
                              'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}},
                   'innerHorizontal': { 
                   # Задаем стиль для внутренних горизонтальных линий
                              'style': 'SOLID',
                              'width': 1,
                              'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}},
                   'innerVertical': { 
                   # Задаем стиль для внутренних вертикальных линий
                              'style': 'SOLID',
                              'width': 1,
                              'color': {'red': 0, 'green': 0, 'blue': 0, 'alpha': 1}}
                              
                              }},


          ]
        }).execute()



def update_sheet(service, spreadsheetId, room, sheetId):
    results = service.spreadsheets().values().batchUpdate(spreadsheetId = spreadsheetId, body = {
    "valueInputOption": "USER_ENTERED", # Данные воспринимаются, как вводимые пользователем (считается значение формул)
    "data": [
        {"range": f"{room}!A2:L20",
         "majorDimension": "ROWS",     # Сначала заполнять строки, затем столбцы
         "values": [
                    ["Кабинет", "Время", "Понедельник","Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"], # Заполняем первую строку
                    [f'{room.split()[-1]}', "10:00 - 11:00" ],
                    ['', "11:00 - 12:00" ],
                    ['', "12:00 - 13:00" ],
                    ['', "13:00 - 14:00" ],
                    ['', "14:00 - 15:00" ],  # Заполняем вторую строку
                    ['', "15:00 - 16:00" ],
                    ['', "16:00 - 17:00" ],
                    ['', "17:00 - 18:00" ],
                    ['', "18:00 - 19:00" ],
                    ['', "19:00 - 20:00" ],
                    ['', "20:00 - 21:00" ],

                   ]}
    ]
}).execute()
    format_sheet(service, spreadsheetId, sheetId)


def get_short_name(full_name):
    full_name_split = full_name.split()
    short_name = f'{full_name_split[0]}'
    short_name += f' {full_name_split[1][0]}.'
    short_name += f'{full_name_split[1][0]}.'
    return short_name

def data_conversion(data, branch, room):
    week = [str(i) for i in range(7)]
    time = [str(i) for i in range(10, 22)]
    n_room = room.split('№')[-1]
    data_branch = []
    for d in data:
        if branch in d:
            data_branch.append(d[branch])

    data_branch_room = []
    for b in data_branch:
        if n_room in b:
            data_branch_room.append(b[n_room])

    d_time = []
    for t in time:
        d_time.append({t:[]})
    print(d_time)
    for r in data_branch_room:
        r = r[0]
        name = list(r.keys())[0]
        short_name = get_short_name(name)
        short_name:r[name][0]



    #     r = r[0]
    #     name = list(r.keys())[0]
    #     short_name = get_short_name(name)
    #     row = r[name]
    #     # print(row)

    #     coll = []

    #     for t in time:
    #         flag = True
    #         for r in row:
    #             if t == r.split(':')[0]:
    #                coll.append(short_name)
    #                flag = False 
    #             continue
    #         if flag:
    #             coll.append('') 
    #     week.append(coll)
    #     # print(coll)
    
    # pprint(week)





def update_sheet_data(service, spreadsheetId, room, data):
    colls = data_conversion(data, branch = 'Свободы 18', room = room)
    results = service.spreadsheets().values().batchUpdate(spreadsheetId = spreadsheetId, body = {
    "valueInputOption": "USER_ENTERED", # Данные воспринимаются, как вводимые пользователем (считается значение формул)
    "data": [
        {"range": f"{room}!C2:J20",
         "majorDimension": "COLUMNS",     # Сначала заполнять строки, затем столбцы
         "values": [
                    colls
                   ]}
    ]
}).execute()





def main(drive, service, files, data):
    name = 'Svobody 18'
    spreadsheet_id = '1QUZCgmCjPmz04LFXBSCE8ELm2jsyFqoWPGoU7BDSdec'
    room = 'Кабинет №4' # TODO 
    is_exist = False

    spreadsheet = service.spreadsheets().get(spreadsheetId = spreadsheet_id).execute()
    sheetList = spreadsheet.get('sheets')
    for sheet in sheetList:
        # print(sheet['properties']['title'])
        if room == sheet['properties']['title']:
            is_exist = True
            sheetId = sheet['properties']['sheetId']

    if not is_exist:
        create_sheet(service, spreadsheet_id, room)
    update_sheet(service, spreadsheet_id, room, sheetId)

    update_sheet_data(service, spreadsheet_id, room, data)




# def get_data(service):
#     sheet = service.spreadsheets()
#     result = sheet.values().get(spreadsheetId=GetSheets_id,
#                                 range="K2:AB4000").execute()
#     values = result.get('values', [])
#     branches = {}
#     for row in values:
#         branches[(row[0])] = {}
#     for row in values:
#         branches[(row[0])].update({row[1]:{row[-1]:[row[2:-2]]}})
#     return branches


def get_data(service):
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=GetSheets_id,
                                range="K2:AB4000").execute()
    values = result.get('values', [])
    branches = []
    for row in values:
        branch = {row[1]:[{row[-1]:row[2:-2]}]}
        branches.append({row[0]:branch})

    return branches

CREDENTIALS_FILE = 'key.json'
GetSheets_id = '1Bu8Hh2jJTXhxtTEGsrePacX4LoQ90kvGXl9tMQ0EO6c'

credentials = ServiceAccountCredentials.from_json_keyfile_name(
    CREDENTIALS_FILE,
    ['https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth)
drive = apiclient.discovery.build('drive', 'v3', credentials=credentials)


files = [
        {
        'id': '1cQUGcPtVwdJbd2ntUSs8n2vyLNnqpl_IonVGX4Xlx64', 
        'name': 'Stroginsky 7k3'
    }, 
        {
        'id': '1usO6Gdaa9Pv1XewiZvddrbMaId4kHDNYLQl53yBcwgY', 
        'name': 'Stroginsky 17k2'
    }, 
        {
        'id': '1QUZCgmCjPmz04LFXBSCE8ELm2jsyFqoWPGoU7BDSdec', 
        'name': 'Свободы 18'
    },
]


# get_data(service)

data = get_data(service)

main(drive, service, files, data)
