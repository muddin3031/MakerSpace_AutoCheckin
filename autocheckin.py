from __future__ import print_function
from datetime import datetime, date, time, timedelta
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
import numpy as np
import glob
import copy

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/calendar.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1VXXQeA8BMPeN57D8VsjM8fvGg0Eg_vLkpHCDX3TNJYc'
SAMPLE_RANGE_NAME = 'A1:AW'

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    sheetsservice = build('sheets', 'v4', credentials=creds)
    calendarservice = build('calendar', 'v3', credentials=creds)


    # Call the Sheets API
    sheet = sheetsservice.spreadsheets()

    print('Getting Training Data')
    sheetresult = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    db_values = sheetresult.get('values', [])

    file = open("output.txt","w")

    # if not db_values:
    #     file.write("No Data \n")
    # else:
    #     for i in db_values:
    #         for j in i:
    #             file.write(j+'\n')

    file.close()


    # Call the Calendar API
    now = datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time

    tmstart = (datetime.combine(datetime.utcnow().date(), time(0, 0)) + timedelta(1)).isoformat() + 'Z'
    tmend = (datetime.combine(datetime.utcnow(), time(23, 59)) + timedelta(1)).isoformat() + 'Z'

    print('Getting Tomorrow Events')
    events_result = calendarservice.events().list(calendarId='primary', timeMin=tmstart, timeMax=tmend, singleEvents=True, orderBy='startTime').execute()
    events = events_result.get('items', [])

    file = open("output.txt","a")
    if not events:
        file.write("No Calendar Events \n")
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        file.write(event['description']+"\n")


    file.close()
    print("done")
    print(db_values[0])
    cancelled1 = sort_training_data(db_values)
    for i in range(len(cancelled1)):
        print(cancelled1[i])
    cancelled2 = sort_reservation_data(db_values)
    for i in range(len(cancelled2)):
        print(cancelled2[i])



def sort_reservation_data(db_values):
    cancel = []
    calendar = open('output.txt', 'r')
    machines = ["Ultimaker 3D Printer Reservation", "Fortus 3D Printer Reservation", "Epilog Fusion Reservation", "Epilog Mini Reservation",
                 "Mojo 3D Printer Reservation", "Elite 3D Printer Reservation","Othermill Reservation","Cricut Reservation","Soldering Station Reservation","Sewing Machine Reservation"]
    for line in calendar:
        if line == "No Calendar Events":
            break;
        if any(x in line for x in machines):
            line = line.replace('\n', '')
            reservation_type = line[14:]

            for line2 in calendar:
                line2 = line2.replace('\n', '')
                if "NetID" in line2:
                    netID = line2[30:]
                if "Net ID" in line2:
                    netID = line2[31:]
                if "YCBM link ref:" in line2:
                    break;

            if not check_reservation(db_values, netID, reservation_type):
                string = ("Cancel ", netID, " ", reservation_type)
                cancel.append(string)
    calendar.close()
    return cancel




def check_reservation(db_values, netID, reservation_type):
    machines = ["Ultimaker 3D Printer Reservation", "Fortus 3D Printer Reservation", "Epilog Fusion Reservation",
                "Epilog Mini Reservation",
                "Mojo 3D Printer Reservation", "Elite 3D Printer Reservation", "Othermill Reservation"]
    ultimaker = 1
    mojo = 2
    elite = 3
    fortus = 4
    ep_mini = 10
    ep_fusion = 11
    othermill = 7
    o_safety = 36
    check = False
    if any(x in reservation_type for x in machines):
        for i in range(len(db_values)):
            for j in range(len(db_values[i])):
                if db_values[i][j]==netID:
                    if "Ultimaker 3D Printer Reservation" == reservation_type:
                        if db_values[i][ultimaker] != 'NULL' and db_values[i][ultimaker] != '':
                            check = True
                    if  "Fortus 3D Printer Reservation"== reservation_type:
                        if db_values[i][fortus] != 'NULL' and db_values[i][fortus] != '':
                            check = True
                    if "Elite 3D Printer Reservation" == reservation_type:
                        if db_values[i][elite] != 'NULL' and db_values[i][elite] != '':
                            check = True
                    if "Mojo 3D Printer Reservation" == reservation_type:
                        if db_values[i][mojo] != 'NULL' and db_values[i][mojo] != '':
                            check = True
                    if "Epilog Fusion Reservation" == reservation_type:
                        if db_values[i][ep_fusion] != 'NULL' and db_values[i][ep_fusion] != '':
                            check = True
                    if "Epilog Mini Reservation" == reservation_type:
                        if db_values[i][ep_mini] != 'NULL' and db_values[i][ep_mini] != '':
                            check = True
                    if "Othermill Reservation" == reservation_type:
                        if db_values[i][othermill] != 'NULL' and db_values[i][othermill] != '':
                            check = True
    else:
        for i in range(len(db_values)):
            for j in range(len(db_values[i])):
                if db_values[i][j]==netID:
                    if (db_values[i][o_safety] != 'NULL' and db_values[i][o_safety] != '' ) or (db_values[i][ultimaker] != 'NULL' and db_values[i][ultimaker] != ''):
                        check = True
    return check















def sort_training_data(db_values):
    cancel=[]
    calendar = open('output.txt','r')
    trainings=["Ultimaker Training","Laser Cutter Training","Vinyl Cutter Training","Soldering Training","MakerGarage Training"]
    for line in calendar:
        if line=="No Calendar Events":
            break;
        if any(x in line for x in trainings):
            line = line.replace('\n', '')
            training_type = line[14:]
            first_name=None
            last_name=None
            netID=None
            n_number=None
            barcode = None
            for line2 in calendar:
                line2 = line2.replace('\n', '')
                if "First Name" in line2:
                    first_name = line2[12:]
                if "First name" in line2:
                    first_name = line2[12:]
                if "Last Name" in line2:
                    last_name = line2[11:]
                if "Last name" in line2:
                    last_name = line2[11:]
                if "NetID" in line2:
                    netID = line2[30:]
                if "Net ID" in line2:
                    netID=line2[31:]
                if "N-Number" in line2:
                    n_number = line2[35:]
                if "Barcode" in line2:
                    barcode = line2[23:]
                if "YCBM link ref:" in line2:
                    break;


            if first_name==None or last_name==None or netID==None or n_number==None or barcode==None:
                string = ("Please do a manual check for: ",netID," for training ",training_type)
                cancel.append(string)
            elif not check_training(db_values,first_name,last_name,netID,n_number,barcode,training_type):
                string = ("Cancel ",netID," ",training_type)
                cancel.append(string)
    calendar.close()
    return cancel















def check_training(db_values,first_name,last_name,netID,n_number,barcode,training_type):

    ultimaker = ""
    safety = ""
    soldering = ""
    vinyl_cutter = ""
    makergarage = ""
    lasercutter= ""
    for name in glob.glob('Training/*.csv'):
        if "Laser Cutter" in name:
            lasercutter=name
        elif "MakerGarage" in name:
            makergarage = name
        elif "Safety" in name:
            safety = name
        elif "Soldering" in name:
            soldering = name
        elif "Ultimaker" in name:
            ultimaker = name
        elif "Vinyl" in name:
            vinyl_cutter = name
    if training_type=="Ultimaker Training":
        df1 = pd.read_csv(ultimaker)
        df2 =pd.read_csv(safety)
        check = check_quizScoresandDB(db_values,df1,df2,first_name,last_name,netID,n_number,barcode,training_type)
    if training_type=="Laser Cutter Training":
        df1 =  pd.read_csv(lasercutter)
        df2 = pd.DataFrame({'A' : []})
        check = check_quizScoresandDB(db_values,df1,df2,first_name,last_name,netID,n_number,barcode,training_type)
    if training_type=="Vinyl Cutter Training":
        df1 =  pd.read_csv(vinyl_cutter)
        df2 = pd.DataFrame({'A' : []})
        check = check_quizScoresandDB(db_values,df1,df2,first_name,last_name,netID,n_number,barcode,training_type)
    if training_type=="Soldering Training":
        df1 =  pd.read_csv(soldering)
        df2 = pd.DataFrame({'A' : []})
        check = check_quizScoresandDB(db_values,df1,df2,first_name,last_name,netID,n_number,barcode,training_type)
    if training_type=="MakerGarage Training":
        df1 =  pd.read_csv(makergarage)
        df2 = pd.DataFrame({'A': []})
        check = check_quizScoresandDB(db_values,df1,df2,first_name,last_name,netID,n_number,barcode,training_type)
    return check








def check_quizScoresandDB(db_values,d1,d2,first_name,last_name,netID,n_number,barcode,training_type):
    check1 = False
    check2 = False
    if d2.empty:#if d2 is empty it means they are not doing an ultimaker training
        for i in range(len(d1['User ID'])):
            if netID == d1['User ID'][i]:
                if d1['Score (%)'][i] >= 75:

                    check1 = True
                else:
                    check1 = False
        check2 = database_Check(db_values,netID)
        if check1 and check2:
            update_database(db_values,netID,training_type)
    else:
        for i in range(len(d1['User ID'])):#This is for O-Ultimaker check
            if netID == d1['User ID'][i]:
                if d1['Score (%)'][i] >= 75:
                    check1 = True
                    add_database(db_values,first_name,last_name,netID,n_number,barcode,'O-Ultimaker')
                    #add_database #If they watched vid and had successful score add them to database
                else:
                    check1 = False
        for i in range(len(d2['User ID'])):#This is for O-Safety check
            if netID == d2['User ID'][i]:
                if d2['Score (%)'][i] >= 75:
                    check2 = True
                    add_database(db_values,first_name, last_name, netID, n_number, barcode,'O-Safety')
                    #add_database #If they watched vid and had successful score add them to database
                else:
                    check2 = False
    return check1 and check2




def database_Check(db_values,netID):
    ultimaker = 1
    check = False
    for i in range(len(db_values)):
        for j in range(len(db_values[i])):
            if db_values[i][j] == netID:
                if db_values[i][ultimaker] != 'NULL' and db_values[i][ultimaker] != '':
                    check = True
    return check





def update_database(db_values,netID,training_type):
    netid_index=db_values[0].index("net_id")
    o_laser=db_values[0].index("O-Laser")
    o_soldering=db_values[0].index("O-Soldering")
    o_vinyl=db_values[0].index("O-Vinyl")
    o_makergarage=db_values[0].index("O-Makergarage")
    for i in range(len(db_values)):
        for j in range(len(db_values[i])):
            if db_values[i][netid_index]==netID:
                if training_type=="Laser Cutter Training" and (db_values[i][o_laser] != 'NULL' and db_values[i][o_laser] != ''):
                    db_values[i][o_laser]=str(datetime.date(datetime.now()))
                if training_type=="Soldering Training" and (db_values[i][o_soldering] != 'NULL' and db_values[i][o_soldering] != ''):
                    db_values[i][o_soldering]=str(datetime.date(datetime.now()))
                if training_type=="Vinyl Cutter Training" and (db_values[i][o_vinyl] != 'NULL' and db_values[i][o_vinyl] != ''):
                    db_values[i][o_vinyl]=str(datetime.date(datetime.now()))
                if training_type=="MakerGarage Training" and (db_values[i][o_makergarage] != 'NULL' and db_values[i][o_makergarage] != ''):
                    db_values[i][o_makergarage]=str(datetime.date(datetime.now()))


def add_database(db_values,first_name, last_name, netID, n_number, barcode, training_type):
    netid_index=db_values[0].index("net_id")
    firstName_index=db_values[0].index("first_name")
    lastName_index=db_values[0].index("last_name")
    nNumber_index=db_values[0].index("n_number")
    barcode_index=db_values[0].index("barcode")
    o_ultimaker=db_values[0].index("O-Ultimaker")
    o_safety=db_values[0].index("O-Safety")
    found = False
    for i in range(len(db_values)):
        if db_values[i][netID]==netID:
            found = True
            if training_type == "O-Ultimaker" and (db_values[i][o_ultimaker] != 'NULL' and db_values[i][o_ultimaker] != ''):
                db_values[i][o_ultimaker] = str(datetime.date(datetime.now()))
            if training_type == "O-Safety" and (db_values[i][o_safety] != 'NULL' and db_values[i][o_safety] != ''):
                db_values[i][o_safety] = str(datetime.date(datetime.now()))
    if not found:
        new_lst=[]
        for i in range(len(db_values[0])):
            new_lst.append("")
        new_lst[netid_index]=netID
        new_lst[firstName_index]=first_name
        new_lst[lastName_index]=last_name
        new_lst[nNumber_index]=n_number[1:]
        new_lst[barcode_index]=barcode
        if training_type=="O-Safety":
            new_lst[o_safety]=str(datetime.date(datetime.now()))
        if training_type=="O-Ultimaker":
            new_lst[o_ultimaker]=str(datetime.date(datetime.now()))
        db_values.append(copy.deepcopy(new_lst))




























if __name__ == '__main__':
    main()
