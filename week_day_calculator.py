import math

def leap_year_check(year:int) -> bool:
    
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
        
def month_num_to_day(month_num: int, year: int) -> int:
    
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
    
def week_day_from_date(date:dict) -> dict:
    
    Year_Last_Digits = int(str(date["year"])[2:4])
    Year_Code = (Year_Last_Digits + (Year_Last_Digits//4))%7
    
    Month_Equivalent = [0,3,3,6,1,4,6,2,5,0,3,5] # this is the conversion key of codes used for day of the week algorthimn
    Month_Code = {Month_Num:Month_Equivalent[Month_Num-1] for Month_Num in range(1, 13)}
    
    Century_Code = { #this is the conversion key for the century of the date
        2000:6}
    
    Total = (Century_Code[2000] + Year_Code + date["date"] + Month_Code[date["month"]])
    
    if leap_year_check(date["year"]):
        Total += -1

    return {"day": Total%7, "date": date["date"]}

def week_span_calculator(event: dict) -> list:

    end_month = int(event["end month"])
    end_day = int(event["end day"])
    start_day = int(event["start day"])
    start_month = int(event["start month"])
    
    day_count = end_day - start_day 
    month_turn_dates = [start_day]
        
    for month_count in range(start_month,end_month):
        month_length = month_num_to_day(month_count,int(event["year"]))
        day_count+= month_length
            
        month_turn_dates+=[month_length,1]
            
    month_turn_dates.append(end_day)

    month_span = []

    for i in range(int(len(month_turn_dates)/2)):
 
        list_range = list(range(month_turn_dates[i],month_turn_dates[i+1]+1))

        if list_range != []:
            
            month_span += [list_range]

        
    return month_span
    
def week_day_from_span(event: dict) -> dict:

    month_span = week_span_calculator(event)
    #week span -> 55 days, etc.
    #start_week_date -> first day date 30th, etc.
    #start_week_day -> day number, 1 = monday
    month_counter = event["start month"]

    event_year_span = []
    
    prev_day = 0
    for month in month_span:
        new_week = []
        for day in month:
            if prev_day > day:
                month_counter += 1
            retrieved_day_data = week_day_from_date({"year": event["year"],"month": month_counter, "date": day})

       
            if retrieved_day_data["day"] % 6 != 0 and retrieved_day_data["day"] % 7 != 0:
                new_week.append(day)
        
               
            prev_day = day
            
        event_year_span.append(new_week)

    final_big_list = []
    final_list = []

    prev_value = event_year_span[0][0]-1
    
    for item in event_year_span:
 
        for number in item:

            

            month = event["start month"]+event_year_span.index(item)
            month_length = month_num_to_day(month, event["year"])


            if abs(number - prev_value) == 1:
                final_list.append({"day":number,"month":month})
            elif prev_value == month_length:
                final_list.append({"day":prev_value,"month":month})
            else:
                final_big_list.append(final_list)
                final_list = [{"day":number,"month":month}]

      
            
                
            prev_value = number
            
    final_big_list.append(final_list)

        
    
            

    return final_big_list 
    
