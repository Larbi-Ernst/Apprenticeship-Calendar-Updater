import pyexcel
from httplib2 import Http
from oauth2client import file, client, tools
import asyncio
from pyexcel_xlsx import save_data, get_data
from googleapiclient.discovery import build
import googleapiclient
from collections import OrderedDict
import json
import datetime
import tkinter as Tkinter

def json_converter(json_date) -> str:
        if isinstance(json_date, datetime.datetime):
            return json_date.__str__()

def leap_year_check(year):
    
    if year/4 != int(year/4) :
        return False
    elif year/100 != int(year/100):
        return True
    elif year/400 != int(year/400):
        return False
    else:
        return True

def month_name_converter(month_name: str) -> int:
    month_name = month_name.lower().replace(" ","")
    month_num = {
        "jan":1,
        "january":1,
        "feb":2,
        "february":2,
        "mar":3,
        "march":3,
        "apr":4,
        "april":4,
        "may":5,
        "jun":6,
        "june":6,
        "july":7,
        "jul":7,
        "aug":8,
        "august":8,
        "sept":9,
        "september":9,
        "october":10,
        "oct":10,
        "nov":11,
        "november":11,
        "dec":12,
        "december":12}

    return month_num[month_name]
        
def month_num_to_day(month_num: int,year: int) -> int:
    
    month_length = {
        1:31,
        2:28,
        3:31,
        4:30,
        5:31,
        6:30,
        7:31,
        8:31,
        9:30,
        10:31,
        11:30,
        12:31}

    if leap_year_check(year):
        month_length[2] = 29
        
    return month_length[month_num]

def week_day_from_date(date:dict) -> int:
    
    Year_Last_Digits = int(str(date["year"])[2:4])
    Year_Code = (Year_Last_Digits + (Year_Last_Digits//4))%7
    
    Month_Equivalent = [0,3,3,6,1,4,6,2,5,0,3,5] # this is the conversion key of codes used for day of the week algorthimn
    Month_Code = {Month_Num:Month_Equivalent[Month_Num-1] for Month_Num in range(1, 13)}
    
    Century_Code = { #this is the conversion key for the century of the date
        2000:6}
    
    Total = (Century_Code[2000] + Year_Code + date["day"] + Month_Code[date["month"]])
    
    if leap_year_check(date["year"]):
        Total += -1

    print("day =",Total%7)
    
def week_day_span(length):
    print(length)


class Interface(Tkinter.Frame):
    
    def __init__(self, master = None):
        Tkinter.Frame.__init__(self, master) 
        #self.create_widgets()
        self.create_list_box()
            
    def create_list_box(self):
        self.listbox_label = Tkinter.Label(self, text = "select files")
        self.listbox_label.pack()
        
        self.listbox = Tkinter.Listbox(self)
        self.listbox.pack()
        self.listbox.insert("end",'YES')
        
    def create_widgets(self):
        self.quit_button = Tkinter.Button(self, text = 'Quit', command = self.quit)
        

class Manager(object):
    
    def __init__(self):
        self.format = "calender"
        self.schedule = []
        self.pages = []
        self.calendarID = ""
   
    def deschedule(self) -> "config":
        for scheduled_sheet in self.schedule:
            print(1)
            yield self.config(scheduled_sheet)
    
    def config(self,sheet_path: str) -> "table_extract":
        print(2)
        self.path = sheet_path
  
        self.sheet_content = get_data(self.path)
        
        self.json_data = json.loads(json.dumps(self.sheet_content, default = json_converter))

        #add a way of gathering all of the page names V
        
        self.page_selection_list = [property_name for property_name in self.json_data]
        self.page_selection_iter = iter(self.page_selection_list)
        self.page_content = {}
        for property_name in self.json_data:

            page = next(self.page_selection_iter)
            self.page_content[page] = [property_value if property_value != [] else None for property_value in self.json_data[property_name]]

            print(self.pages)
            if page in self.pages:
        
                self.table_extract(self.page_content[page], False)
        
    def table_extract(self, page_data: list,Condition) -> "calender_update":

        multiple = False
        known_categories = ["year","dates","month","weekend","location"]

        for line_content in page_data:
            if isinstance(line_content,list):
                for item in line_content:
                    if str(item).lower() in known_categories:
                        categories = line_content

        for line_content in page_data:

            if line_content != None:       
        
                if not isinstance(line_content[0],float):
                        
                    page_data.remove(line_content)
              
            else:
                page_data.remove(line_content)
                
        extra_dates_start = page_data.index(categories,2)
        college_dates = page_data[0:extra_dates_start]
        extra_dates = page_data[extra_dates_start:len(page_data)]
        
        for item in extra_dates:

            print(item)
            if item[0] != "":
                date = item[0]
                
            else:
                item[0] = date
         
        final_dates = [date for date in college_dates+extra_dates if type(date[0]) == float]
        print(final_dates)
        self.json_create(final_dates)
        print("YES")
        
    
    def json_create(self, final_event_dates: list):
        
        for event in final_event_dates:
            date_suffixes = ["th", "rd", "nd", "st"]

            num_strings = [str(num) for num in range(0,10)]
            
            event_days = [day if any(character in num_strings for character in day) else "" for day in event[2].split("-")]
            
            for day in event_days:
                
                for date_suffix in date_suffixes:
                    
                    if " " in day:
                        event_days[event_days.index(day)] = day.replace(" ","")
                        day = day.replace(" ","")
                        
                    if date_suffix in day.lower():
                        event_days[event_days.index(day)] = day.replace(date_suffix,"")
                        
            if event_days[0] != "":
                
                if len(event_days) == 1:
                    event_days*=2
                    
                month = month_name_converter(event[1])

                try:
                    int(event_days[1])
                    other_month = month
                    
                except ValueError:
                    for key,_ in conversion.items():
                        
                        if key in event_days[1]:
                            event_days[1] = event_days[1].replace(key, "")
                            other_month = month_name_converter(key)

                self.event_split({"start day":event_days[0],"end day":event_days[1],"start month":month,"end month":other_month,"year":event[0]})
                
                event_item = {
                    
                    'header':'Content-Type: application/json',
                    'summary': event[3],
                    'start':  {
                        'date': datetime.date((int(event[0])), month, int(event_days[0])).__str__()
                        },
                    'end':  {
                        'date': datetime.date((int(event[0])), other_month, int(event_days[1])).__str__()
                        },
                    'location': "Ada. National College for Digital Skills, Broad Lane, London N15 4AG, UK"
                }   
              
        #self.calendar_update(event_item)
    def event_split(self, event: dict):
        print(event)
    def calendar_update(self, json_item: dict):

        SCOPES = 'https://www.googleapis.com/auth/calendar'
        store = file.Storage('storage.json')
        credentials = store.get()
        
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
            creds = tools.run_flow(flow, store)
            
        API_KEY = "AIzaSyDHN04fCmdUH5Tw3uvkgBiq-AJKM8RIiUY"
        google_calendar = build('calendar', 'v3', developerKey = API_KEY, http=creds.authorize(Http()))

        google_calendar.events().insert(calendarId = self.calendarID, body = json_item).execute()

        print(complete)

        
    def update(self):
        
        while len(self.schedule) > 0:
            print(0)
            self.current_event = self.deschedule()
            next(self.current_event)
            self.schedule.remove(self.schedule[0])
            
            
A = Manager()

with open("C:/Users/Student/Documents/Apprenticeship Script/Calendar Automation Start/InputData.txt", "r") as INPUT_CONTENT:
    INPUT_CONTENT = list(INPUT_CONTENT)
    File_name = INPUT_CONTENT[0][6:].replace("\n", "")
    A.pages = INPUT_CONTENT[1][7:].replace("\n", "").split(",")
    print(A.pages)
    A.CalendarID = INPUT_CONTENT[2][11:].replace("\n", "")
    
#A.schedule.append(File_name)
#A.update()

#App = Interface()
#App.master.title('Calendar Manager')
#App.mainloop()
