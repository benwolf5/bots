# -*- coding: utf-8 -*-
"""
Created on Wed Feb 12 14:42:46 2020

@author: benwo
"""

#Rota
import openpyxl as xl
import os 
import datetime
import random

class cellarman_getter():
    def __init__(self):
        self.cellarman = dict(zip(cellarman, [0] * len(cellarman)))
        self.due_busy = dict(zip(cellarman, [0] * len(cellarman)))
        self.bar_staff = list(bar_staff)
        self.door_staff = list(door_staff)
        self.cellar_recent = []
        self.bar_recent = []
        self.door_recent = []
        
        
    def get_potential(self, utility, group):
        listOfPotential = []
        for key, value in group.items():
            if value == utility:
                listOfPotential.append(key)
        return listOfPotential
        
    def get_cellarman(self, shift_type):
        exit_loop = True
        while exit_loop == True:
            minUtility = min(self.cellarman.items(), key=lambda x: x[1])[1]
            listOfPotential = self.get_potential(minUtility, self.cellarman)
            name = listOfPotential.pop(random.randrange(0,len(listOfPotential)))
            print(name)
            print(self.cellarman)
            print(self.due_busy)
            print(self.cellar_recent)
            if shift_type == 0: # normal
                if name not in self.cellar_recent:
                    self.cellarman[name] += 1
                    self.cellar_recent.append(name)
                    exit_loop = False
            else:
                offset = 0
                minUtility = min(self.due_busy.items(), key=lambda x: x[1])[1] + offset
                listOfPotential_busy = self.get_potential(minUtility, self.due_busy)
                name = listOfPotential_busy.pop(random.randrange(0,len(listOfPotential_busy)))
                if name not in self.cellar_recent:
                    self.due_busy[name] += 1
                    self.cellar_recent.append(name)
                    exit_loop = False
                else:
                    if len(self.bar_recent) == 1:
                        offset += 1
        if len(self.cellar_recent) == 10:
            self.cellar_recent = []
        return name

    def get_barstaff(self):
        exit_loop = True
        while exit_loop == True:
            name = self.bar_staff.pop(random.randrange(0,len(self.bar_staff)))
            if name in self.bar_recent:
                self.bar_staff.append(name)
            else:
                self.bar_recent.append(name)
                exit_loop = False                    
        if len(self.bar_staff) == 0:
            self.bar_staff = list(bar_staff)
        if len(self.bar_recent) == 30:
            self.bar_recent = []
        return name
    
    def get_doorstaff(self):
        exit_loop = True
        while exit_loop == True:
            name = self.door_staff.pop(random.randrange(0,len(self.door_staff)))
            if name in self.door_recent:
                self.door_staff.append(name)
            else:
                self.door_recent.append(name)
                exit_loop = False                    
        if len(self.door_staff) == 0:
            self.door_staff = list(door_staff)
        if len(self.door_recent) == 10:
            self.door_recent = []
        return name
        

wb = xl.Workbook()
sheet = wb.active
sheet.title = "Easter Term Rota"
number_weeks = 10
cellarman=["Wolfy","Kavi","Pat","Emilibobs","SGG","Jwal","Izzy","Abi","Sam","Stonk","Adam","Liberty","Sandys","Dallas","Charlotte","Watson","Cayford","Jimbob","Redfern","Nikhil"]
week_list = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
row_titles = [" ", " ", "Cellarman", "Staff", "Staff", " ", "Quad Cellarman", "Quad Staff", " ", "Door Staff", "Door Staff", " "]
start_date = datetime.date(2020,4,26)
deltaday = datetime.timedelta(days=1)
number_staff = 58
number_doorstaff = 43

#read bar staff from txt file
bar_staff = []
file_staff = open("barstaff.txt","r")
file_staff1 = file_staff.readlines()
for i in range(number_staff):
    bar_staff.append(file_staff1[4*i])

#read door staff from txt file
door_staff = []
file_door = open("doorstaff.txt","r")
file_door1 = file_door.readlines()
for i in range(number_doorstaff):
    door_staff.append(file_door1[4*i])

cg = cellarman_getter()

#make it look pretty
for i in range(3,122):
    sheet.row_dimensions[i].height = 30
for i in range(8):
    sheet.column_dimensions[xl.cell.cell.get_column_letter(i+1)].width = 20
for i in range(number_weeks):
    for day in range(7):
        current_cell1 = sheet.cell(row = (12*i)+3,column = day+2)
        current_cell1.fill = xl.styles.PatternFill(fgColor="85C1E9", fill_type = "solid")
        current_cell2 = sheet.cell(row = (12*i)+4,column = day+2)
        current_cell2.fill = xl.styles.PatternFill(fgColor="EC7063", fill_type = "solid")

# Days of week loop 
#Week
for i in range(number_weeks):
    #days
    for day in range(7):
        current_cell = sheet.cell(row = (12*i)+3,column = day+2)
        current_cell.value = week_list[day]

#Row heading loop
#Days of Week
for i in range(number_weeks):
    #heading
    for j in range(12):
        current_cell = sheet.cell(row=3+j+(12*i),column=1)
        current_cell.value = row_titles[j]
        current_cell.fill = xl.styles.PatternFill(fgColor="2ECC71",fill_type="solid")
     
#Date loop
current_date = start_date
for i in range(number_weeks):
    for day in range (7):
        current_cell = sheet.cell(row = (12*i)+4, column = day + 2)
        current_cell.value = str(current_date)
        current_date = current_date + deltaday
        
#Bar Staff 
for i in range(number_weeks):
    for day in range(7):
        for k in range(2):
            current_cell = sheet.cell(row = (12*i)+6+k, column = day + 2)
            if i == 0 and day == 0:
                current_cell.value = "Barcom"
            current_cell.value = cg.get_barstaff()
            if (day%7 == 3) or (day%7 == 5) or (day%7 == 6):
                current_cell3 = sheet.cell(row = (12*i)+10, column = day + 2)
                current_cell3.value = cg.get_barstaff()


#Door Staff
for i in range(number_weeks):
    for day in range(7):
        if (day%7 == 3) or (day%7 == 5) or (day%7 == 6):
            for k in range(2):
                current_cell = sheet.cell(row = (12*i)+12+k, column = day + 2)
                current_cell.value = cg.get_doorstaff()
                
#Cellarman 
current_date = start_date
for i in range(number_weeks):
    for day in range(7):
        current_cell = sheet.cell(row = (12*i)+5, column = day + 2)
        if i == 0 and day == 0:
            current_cell.value = "Barcom"
        else:
            current_cell.value = cg.get_cellarman(0)
        if (day%7 == 3) or (day%7 == 5) or (day%7 == 6):
            while True:
                potential_cellarman = cg.get_cellarman(1)
                if potential_cellarman != current_cell.value: #stops quad and main being same
                    current_cell = sheet.cell(row = (12*i)+9, column = day + 2)
                    current_cell.value = potential_cellarman
                    break
        current_date = current_date + deltaday
                
                    

wb.save(r"C:\Users\benwo\OneDrive\Documents\Chad's Year 3\Bar Manager\rota.xlsx")
