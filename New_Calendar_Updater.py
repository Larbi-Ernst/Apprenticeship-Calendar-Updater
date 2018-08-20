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

class Manager(object):

    def __init__(self):
        self.format = "calender"
        self.schedule = []
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
        self.Pages = []
        self.CalendarID = ""
   
    def deschedule(self) -> "config":
        
        for path_data in self.schedule: 
            yield self.config(path_data)
    
    def config(self,path_data: str) -> "table_extract":

        self.path = path_data
  
        self.data = get_data(self.path)

        #HERE
        
        self.json_data = json.loads(json.dumps(self.data,default=self.json_converter))
        self.names = iter(["Page1","Page2","Page3","Page4","_","_","_","_","_","_"])

        for item in self.json_data:

            name = next(self.names)
            setattr(self,name,[item if item != [] else None for item in self.json_data[item]])

            if name in self.Pages:
                self.table_extract(eval(f"self.{name}"),False)
                
        #cohort_1
        
    def json_converter(value,x) -> str:
        
        if isinstance(value,datetime.datetime):
            return value.__str__()
        
    def table_extract(self, data_set: list,Condition) -> "calender_update":


        multiple = False
        
        for item in data_set:
            
            if item != None:
           
                if len(item) != 4:
                    data_set.remove(item)
            else:
                data_set.remove(item)

        Extra_Date_Start = data_set.index(['Year','Month','Dates'],2)
        College_Dates = data_set[0:Extra_Date_Start]
        Extra_Dates = data_set[Extra_Date_Start:len(data_set)]
        
        for item in Extra_Dates:
            if item[0] != "":
                date = item[0]
            else:
                item[0] = date
         
        Final_Dates = [item for item in College_Dates+Extra_Dates if type(item[0]) == float]
        self.calendar_update(Final_Dates)

    def month_name_converter(self, name: str) -> str:
        name = name.split(" ")[0]

        return self.conversion[name]
    
    def calendar_update(self, final_data_set: list):

        SCOPES = 'https://www.googleapis.com/auth/calendar'
        store = file.Storage('storage.json')
        creds = store.get()
        if not creds or creds.invalid:
            flow = client.flow_from_clientsecrets('client_secret.json', SCOPES)
            creds = tools.run_flow(flow, store)
        print("YES")
        API_KEY = "AIzaSyDHN04fCmdUH5Tw3uvkgBiq-AJKM8RIiUY"
        GCAL = build('calendar', 'v3', developerKey = API_KEY,http=creds.authorize(Http()))
        
        for event in final_data_set:


            date_suffixes = ["th","rd","nd","st"]

            print(event[2])
            A = [item if any(i in [str(x) for x in range(0,10)] for i in item) else "" for item in event[2].split("-")]
            
            for item in A:
                for date_suffix in date_suffixes:
                    if " " in item.lower():
                        A[A.index(item)] = item.replace(" ","")
                        item = item.replace(" ","")
                        
                    if date_suffix in item.lower():
                        A[A.index(item)] = item.replace(date_suffix,"")
                    
            print(A)
                        
                
            if A[0] != "":
                if len(A) == 1:
                    A*=2
                    
                Month = self.month_name_converter(event[1])

                try:
                    int(A[1])
                    Other_Month = Month
                except:
                    for key,_ in self.conversion.items():
                        if key in A[1]:
                            A[1] = A[1].replace(key,"")
                            Other_Month = self.month_name_converter(key)
                print(Other_Month)
                print(A)
                    
                
                event_item = {
                    'header':'Content-Type: application/json',
                    'summary': event[3],
                    'start':  {
                        'date': datetime.date((int(event[0])),int(Month),int(A[0])).__str__()
                        },
                    'end':  {
                        'date': datetime.date((int(event[0])),int(Other_Month),int(A[1])).__str__()
                        },
                    'location': "Ada. National College for Digital Skills, Broad Lane, London N15 4AG, UK"
                }   
                 
                GCAL.events().insert(calendarId = self.CalendarID,body= event_item).execute()
        
        
        print("COMPLETE")        
        
    def update(self):
        while len(self.schedule) > 0:
            self.current_event = self.deschedule()
            next(self.current_event)
            self.schedule.remove(self.schedule[0])
      
            
A = Manager()
with open("C:/Users/Student/Documents/Apprenticeship Script/Calendar Automation Start/InputData.txt","r") as INPUT_CONTENT:
    INPUT_CONTENT = list(INPUT_CONTENT)
    File_name = INPUT_CONTENT[0].replace("File: ","").replace("\n","")
    A.Pages = [item for item in INPUT_CONTENT[1].replace("Pages: ","").replace("\n","").split()]
    A.CalendarID = INPUT_CONTENT[2].replace("CalendarID: ","").replace("\n","")
A.schedule.append(File_name)
A.update()
