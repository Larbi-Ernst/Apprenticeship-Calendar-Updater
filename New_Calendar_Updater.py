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
            yield self.config(scheduled_sheet)
    
    def config(self,sheet_path: str) -> "table_extract":

        self.path = sheet_path
  
        self.sheet_content = get_data(self.path)
        
        self.json_data = json.loads(json.dumps(self.sheet_content, default = self.json_converter))

        #add a way of gathering all of the page names V
        
        self.page_selection = iter(["Page1", "Page2", "Page3", "Page4", "_", "_", "_", "_", "_", "_"])

        for property_name in self.json_data:

            page = next(self.page_selection)
            setattr(self, name,[property_value if property_value != [] else None for property_value in self.json_data[property_name]])

            if page in self.pages:
                self.table_extract(eval(f"self.{name}"), False)

    
    def json_converter(json_date, _) -> str:
        if isinstance(json_date, datetime.datetime):
            return json_date.__str__()
        
    def table_extract(self, page_data: list,Condition) -> "calender_update":

        multiple = False
        
        for line_content in page_data:
            
            if line_content != None:
           
                if len(line_content) != 4:
                    data_set.remove(line_content)
            else:
                data_set.remove(line_content)

        extra_dates_start = data_set.index(['Year','Month','Dates'],2)
        college_dates = data_set[0:extra_date_start]
        extra_dates = data_set[extra_date_start:len(data_set)]
        
        for item in extra_dates:
            
            if item[0] != "":
                date = item[0]
                
            else:
                item[0] = date
         
        final_Dates = [dates for dates in college_dates+extra_dates if type(dates[0]) == float]
        self.calendar_update(final_dates)

    def month_name_converter(self, name_of_month: str) -> str:
        name_of_month = name_of_month.split(" ")[0]
        self.conversion = {
            "Jan":"01",
            "January":"01",
            "Feb":"02",
            "March":"03",
            "April":"04",
            "May":"05",
            "June":"06",
            "July":"07",
            "August":"08",
            "Sept":"09",
            "September":"09",
            "Oct":"10",
            "October":"10",
            "Nov":"11",
            "November":"11",
            "Dec":"12",
            "December":"12"}

        return self.conversion[name_of_month]
    
    def calendar_update(self, final_event_dates: list):

        SCOPES = 'https://www.googleapis.com/auth/calendar'
        store = file.Storage('storage.json')
        credentials = store.get()
        
        if not credentials or credentials.invalid:
            flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
            creds = tools.run_flow(flow, store)
            
        API_KEY = "AIzaSyDHN04fCmdUH5Tw3uvkgBiq-AJKM8RIiUY"
        google_calendar = build('calendar', 'v3', developerKey = API_KEY, http=creds.authorize(Http()))
        
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
                    
                month = self.month_name_converter(event[1])

                try:
                    int(event_days[1])
                    other_month = month
                    
                except ValueError:
                    for key,_ in self.conversion.items():
                        
                        if key in event_days[1]:
                            event_days[1] = event_days[1].replace(key, "")
                            other_month = self.month_name_converter(key)
                
                event_item = {
                    
                    'header':'Content-Type: application/json',

                    'summary': event[3],

                    'start':  {

                        'date': datetime.date((int(event[0])), int(month), int(event_days[0])).__str__()

                        },

                    'end':  {

                        'date': datetime.date((int(event[0])), int(other_month), int(event_days[1])).__str__()

                        },
                    
                    'location': "Ada. National College for Digital Skills, Broad Lane, London N15 4AG, UK"
                }   
                 
                google_calendar.events().insert(calendarId = self.calendarID, body = event_item).execute()
        
        
        print("COMPLETE")        
        
    def update(self):
        
        while len(self.schedule) > 0:  
            self.current_event = self.deschedule()
            next(self.current_event)
            self.schedule.remove(self.schedule[0])
      
            
A = Manager()

with open("C:/Users/Student/Documents/Apprenticeship Script/Calendar Automation Start/InputData.txt", "r") as INPUT_CONTENT:
    INPUT_CONTENT = list(INPUT_CONTENT)
    File_name = INPUT_CONTENT[0].replace("File: ", "").replace("\n", "")
    A.Pages = [item for item in INPUT_CONTENT[1].replace("Pages: ", "").replace("\n", "").split()]
    A.CalendarID = INPUT_CONTENT[2].replace("CalendarID: ", "").replace("\n", "")
    
#A.schedule.append(File_name)
#A.update()

App = Interface()
App.master.title('Calendar Manager')
App.mainloop()
