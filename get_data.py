from libraries import *
def get_data():
    #autoryzacja google
    print("#################################################################")
    print("#-----------------> User and Devaices Resources <---------------#")
    print("#---> Copyright © 2022 Mateusz Zarzycki All Rights Reserved <---#")
    print("#################################################################")

    scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("auth/creds.json", scope)
    client = gspread.authorize(creds)

    global sheet
    sheet = client.open("sample-db")

    global data
    global data_sheet_name_main
    data_sheet_name_main = sheet.get_worksheet(0)
    data = data_sheet_name_main.get_all_records()

get_data()