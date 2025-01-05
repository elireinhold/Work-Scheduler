import pandas as pd
from ortools.sat.python import cp_model

#Hours of operation: 7am-8pm Mon-Thurs, 7am-10pm Fri, 11am-8pm Sat-Sun
#Shifts: 7am-2pm and 2pm-close Mon-Friday, 11am-8pm Sat-Sun *5 people per shift*
#Need to error check valid availabilty times
#When inputing availability use military times
#Example format for availability: Mon 7-20, Thurs 15-20, Fri 7-22, Sat 11-20, Sun 11-20

class ReadFile:
    def __init__(self, file_path):
        self.file_path = file_path
    
    def read_file(self):
        return pd.read_excel(self.file_path).to_dict('records')
    
class WriteFile:
    def __init__(self, file_path, data):
        self.file_path = file_path
        self.data = data
    
    def write_file(self):
        return pd.DataFrame.from_dict(self.data, orient='index').transpose().to_excel('schedule.xlsx', index=False)

def parseAvailabilty(availability):

    days = ["Mon", "Tues", "Wed", "Thurs", "Fri", "Sat", "Sun"]
    slots = []

    #availability string format: Mon 9-12, Tues 1-5
    shifts = availability.split(',')
    for shift in shifts:
        day, hours = shift.strip().split(' ')
        start, end = hours.split('-')
        start = int(start)
        end = int(end)

        #Converts hour block into numbered slots
        #Each slot corresponds to a specific hour and day of the week (0-82)
        for hour in range(start, end):
            if day == "Fri":
                slot = 52 + (hour - 7)
                slots.append(slot)
            elif day == "Sat" or day == "Sun":
                #Calculates slot for weekend
                slot = 67 + (days.index(day)-5) * 9 + (hour - 11)
                slots.append(slot)
            
            else:
                #Calculates slot for weekday (13 hrs in 1 day starting at 7am)
                slot = days.index(day) * 13 + (hour - 7)
                slots.append(slot)

    return slots

def generateSchedule(data):
    model = cp_model.CpModel()

    #Varibles
    num_employees = len(data)
    max_slots = 85
    shifts = {}
    for i in range(num_employees):
        for s in range(max_slots):
            #1 if Employee i is assigned to slot s, 0 if not
            shifts[(i, s)] = model.NewBoolVar(f"shift_e{i}_s{s}")

    #Constraints
    #Employee Availability
    for i, employee in enumerate(data):
        available_slots = parseAvailabilty(employee['Availability'])
        for s in range(max_slots):
            #Checks if employee can not work the slot
            if s not in available_slots:
                model.Add(shifts[(i,s)] == 0)

    #Max of 5 employees working at once
    for s in range(max_slots):
        model.Add(sum(shifts[(i,s)] for i in range(num_employees)) <= 5)

    #Max of desired hours
    for i, employee in enumerate(data):
        model.Add(sum(shifts[(i,s)] for s in range(max_slots)) <= employee['Desired Hours'])

    #Needs to fill whole shift

    #Define objective function
    model.Maximize(sum(shifts[(i, s)] for i in range(num_employees) for s in range(max_slots)))

    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    #If the status is optimal or feasible the slot is appended to the slot of employee
    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        schedule = {}
        for i, employee in enumerate(data):
            assigned_slots = []
            for s in range(max_slots):
                if solver.Value(shifts[(i, s)]) == 1:
                    assigned_slots.append(s)
            schedule[employee['Name']] = assigned_slots
        return schedule
    else:
        return None
    
def reformatSchedule(schedule):
    days = ["Mon", "Tues", "Wed", "Thurs", "Fri", "Sat", "Sun"]
    newSchedule = {}

    for name, slots in schedule.items():
        newSchedule[name] = []
        for slot in slots:
            #Mon-Thurs
            if slot < 52:
                day = days[slot//13]
                hour = (slot % 13)+7
                if hour < 20 and hour > 7:
                    if hour > 12:
                        hour = hour - 12
                        newSchedule[name].append(f"{day} {hour}pm-{hour+1}pm")
                    elif hour == 12:
                        newSchedule[name].append(f"{day} {hour}pm-1pm")
                    elif hour == 11:
                        newSchedule[name].append(f"{day} {hour}am-12pm")
                    else:
                        newSchedule[name].append(f"{day} {hour}am-{hour+1}am")
            
            #Fri
            elif slot < 67:
                day = "Fri"
                hour = (slot -52) + 7
                if hour < 22 and hour > 7:
                    if hour > 12:
                        hour = hour - 12
                        newSchedule[name].append(f"{day} {hour}pm-{hour+1}pm")
                    elif hour == 12:
                        newSchedule[name].append(f"{day} {hour}pm-1pm")
                    elif hour == 11:
                        newSchedule[name].append(f"{day} {hour}am-12pm")
                    else:
                        newSchedule[name].append(f"{day} {hour}am-{hour+1}am")

            #Sat-Sun
            else:
                day = days[(slot-67) // 9 + 5]
                hour = (slot - 67 ) % 9 + 11
                if hour < 20 and hour > 11:
                    if hour > 12:
                        hour = hour - 12
                        newSchedule[name].append(f"{day} {hour}pm-{hour+1}pm")
                    elif hour == 12:
                        newSchedule[name].append(f"{day} {hour}pm-1pm")
                    elif hour == 11:
                        newSchedule[name].append(f"{day} {hour}am-12pm")
                    else:
                        newSchedule[name].append(f"{day} {hour}am-{hour+1}am")

    return newSchedule

if __name__ == '__main__':
    sheet1 = ReadFile("C:\\Users\\elire\\Downloads\\Scheduler Project\\employee_data.xlsx")
    data1 = sheet1.read_file()

    schedule = generateSchedule(data1)

    #Prints the 1 hour slots in the terminal in a readable format
    if schedule:
        schedule = reformatSchedule(schedule)
        for name,shifts in schedule.items():
            print(f"{name}:")
            for shift in shifts:
                print(f"  - {shift}")
        
        write_file = WriteFile("C:\\Users\\elire\\Downloads\\Scheduler Project\\schedule.xlsx", schedule)
        write_file.write_file()