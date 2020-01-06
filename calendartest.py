import openpyxl
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from httplib2 import Http 

class Person:
    #Initializes the person and collects data from the list. Add times, if needed.
    def __init__(self, personData):
        self.name = personData[0]
        self.firstTime = personData[1]
        self.secondTime = personData[2]
        self.email = personData[3]
        self.place = personData[4]
        self.description = personData[5]
        self.invitor = personData[6]

    def rtn_firstTime(self):
        return self.firstTime

    def rtn_secondTime(self):
        return self.secondTime

def readDetails():
    #Reads your .xlsx file one row at the time.
    wb = openpyxl.load_workbook("filename.xlsx")
    ws = wb["sheetname"]

    persondata = []
    service = cldinit() 
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            persondata.append(cell.value)
        
        name = Person(persondata)
        parsi(name, service)
        persondata = []

def maintimeconv(time):
    #Converts time to Google API requirements. In this case e.g. Jan 1, 12:00pm - 12:30pm.
    timetrim = time.split("-")
    alku = timetrim[0].rstrip()
    loppu = timetrim[0][:7].rstrip() + timetrim[1].rstrip()
    formatti = "%Y%b %d, %I:%M%p"
    alku = datetime.datetime.strptime("2019" + alku, formatti)
    loppu = datetime.datetime.strptime("2019"+ loppu, formatti)
    return alku, loppu

def parsi(name, service):
    #Sends the right amount of events for each person. Add more events, if necessary.
    if name.rtn_firstTime(): 
        alku, loppu = stgtimeconv(name.rtn_firstTime())
        alku = str(alku.date())+ "T" + str(alku.time()) + "+02:00"
        loppu = str(loppu.date())+ "T" + str(loppu.time()) + "+02:00"
        print(alku, " # ", loppu, "Stage", name.name)
        calendarevent("Event", "Central Park " + name.place, name.description, alku, loppu, name.email, service)

    if name.rtn_studio():
        alku, loppu = maintimeconv(name.rtn_studio())
        alku = str(alku.date())+ "T" + str(alku.time()) + "+02:00"
        loppu = str(loppu.date())+ "T" + str(loppu.time()) + "+02:00"
        print(alku, " # ", loppu, "Studio", name.name)
        calendarevent("Second event", "Best Coffee Place", name.description, alku, loppu, name.email, service)


def cldinit():
    SCOPES = ['https://www.googleapis.com/auth/calendar.events']
    creds = None

    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)   

    service = build('calendar', 'v3', credentials=creds)

    return service

def calendarevent(event, place, description, beginning, ending, person, service):
    event = {
        'summary' : event,
        'location' : place,
        'description' : description,
        'start' : {
            'dateTime' : beginning,
            'timeZone' : 'Europe/Helsinki'
        },
        'end' : {
            'dateTime' : ending,
            'timeZone' : 'Europe/Helsinki'
        },
        'attendees' : [
            {'email' : person}
        ],
        'reminders' : {
            'useDefault' : False,
            'overrides' : {
                'method' : 'popup', 'minutes' : 10
            },
        },
        'organiser' : {
            'email' : 'youremailaddress@emails.com',
        },
        'creator' : {
            'email' : 'youremailaddress@emails.com',
        },
    }     

    event = service.events().insert(calendarId='primary', sendNotifications=True, body=event).execute()
    print("Event", event, "created for:", person)
    print(" ")


def main():
    readDetails()

main()