"""
This script belongs to the Phi Kappa Psi Cal Gamma chapter. 

Use it to generate random door and bar shifts. 

@author Vladislav Karchevsky
@author Ryan Flynn
@author Reese Levine
"""

import random
import xlsxwriter
import os
#import easygui as eg 

# How many shifts?
NUM_SLOTS = 7
DATE = raw_input("What is the date of this event?\n")
#DATE = eg.enterbox("What is the date of the event?", "Date")

def pickRandomBros(broList, numSample):
    # Note that this function affects the original list
    numSample = min(numSample, len(broList))
    randomSet = random.sample(set(broList), numSample)
    for bro in randomSet:
        broList.remove(bro)
    return randomSet

# ------------------------ start of Erik'd controbution -----------------------

def pick_bros(class_list, num):
    """
    Choose brothers randomly after ordering them by year in school
    :param class_list: list of (name, year) tuples for members of class
    :param num: number of people to be selected out of class
    :return: list of names of people selected out of class
    """
    years = [[name for (name, year) in class_list if year == i] for i in xrange(1,5)]
    names = list()
    for year in years:
        if year:
            sample_num = min(num, len(year))
            if sample_num:
                names += random.sample(year, sample_num)
                num -= sample_num
        if not num:
            break
    print names
    return names

# ----------------------- End of erik's controbution -----------------------------

doorShift1 = []
doorShift2 = []
doorShift3 = []
doorShift4 = []
barShift1 = []
barShift2 = []
barShift3 = []
barShift4 = []


# Create brotherhood lists - tuples of name and year (I probably messed up someone's year but idgaf)
epsilon = [(p, 4) for p in ["Mitchell Pok",
           "Paul Levchenko",
           "Conor Stanton",
           "Rehman Minhas"]]
# this just works because everyone in epsilon is a senior

zeta = [("Taylor Ferguson", 3),
        ("Evan Mason", 3),
        ("Eric Gabrielli", 3),
        ("Chris Farmer", 3),
        ("Zachary Hawtof", 3),
        ("Ian Mason", 3),
        ("Elliot Surovell", 3),
        ("Anurag Reddy", 3),
        ("Ryan Flynn", 3),
        ("Sam Rausser", 3),
        ("Rikesh Patel", 4),
        ("Jack Hendershott", 3),
        ("Mark Traganza", 3),
        ("Evin Wieser", 3),
        ("Matt Buckley", 3)]

eta = [("Kyle Joyner", 3),
       ("Richard Mercer", 3),
       ("Andrew Soncrant", 3),
       ("Joey Papador", 4),
       ("Christian Collins", 4),
       ("Anand Dharia", 3),
       ("Francisco Torres", 4),
       ("Mustapha Khokhar", 1)]

theta = [("Aman Khan", 2),
         ("Andrew Ahmadi",2),
         ("Aneesh Prasad",2),
         ("Ben Kurschner",2),
         ("Christos Gkolias",2),
         ("Elliot Dunn", 4),
         ("Harrison Agrusa",2),
         ("Jack Sweeney",2),
         ("Jason Blore",2),
         ("Jeremy Mack",2),
         ("Joe Labrum",2),
         ("Keeton Ross",2),
         ("Lawrence Dong", 4),
         ("Matt Nisenboym", 4),
         ("Mitchell Stieg",2),
         ("Nabil Farooqi",2),
         ("Nathan Aminpour",2),
         ("Reese Levine",2),
         ("Ricky Philipossian",2),
         ("Riley Pok",2),
         ("Rokhan Khan",2),
         ("Sahand Saberi",2),
         ("Thomas Zorrilla",2),
         ("Will Morrow", 2)]

iota = [("Alex Clark", 2),
        ("Kenny Dang", 4),
        ("Brent Freed", 3),
        ("Jacob Gill", 2),
        ("Darius Kay", 2),
        ("David Kret", 3),
        ("Will Lopez", 3),
        ("Dhruv Malik", 2),
        ("Ian Moon", 2),
        ("Francisco Peralta",2),
        ("Brandt Sheets", 2),
        ("Andrew Ting", 3),
        ("Evan Wilson", 2)]

kappa = [("Anthony Fortney", 3),
         ("Steven Lin",2),
         ("Ford Noble", 1),
         ("Ryan Leyba", 1),
         ("Robert Mcilhatton", 1),
         ("Jonathan Tuttle", 1),
         ("Morris Ravis", 1),
         ("Ben Lalezari", 1),
         ("Drew Hanson", 1),
         ("Josh Bradley-Bevan", 1),
         ("Steven Beelar", 1),
         ("Gabriel Bogner", 1),
         ("Dylan Dreyer", 1),
         ("Luke Thomas", 1),
         ("Konstantinos Tzartzas", 1),
         ("Nate Parke", 1),
         ("Dan Lee", 1),
         ("Max Seltzer", 1),
         ("Andy Frey", 1),
         ("Nathan Kelleher", 1),
         ("Arnav Chaturvedi", 1),
         ("Sam Giacometti", 1),
         ("Sam Bauman", 1)]

# Aggregate brothers gone from social event
permaAbsent = ["Zachary Hawtof", "Ryan Flynn", "Evin Wieser", "Christian Collins", "Erik Bartlett"]

# -------- Reese's addition for special events with more positions -----------
specialEvent = raw_input("Is this a special event? ")

if specialEvent.strip().lower() == "yes":
         runSpecialScheduler = True
else:
        runSpecialScheduler = False

# -------- Ryan's addition to make excluding absent brothers easier -----------

anyAbsent = raw_input("Will any brothers be absent from this event? ")
#absentMsg = "Will any brothers be absent from this event?"
#absentTtl = "Any Absent?"
#anyAbsent = eg.enterbox(absentMsg, absentTtl)
if anyAbsent.strip().lower() == "yes":
    runAbsentSurvey = True
    print("Please enter the names of each absent brother, one per line. Once " +
          "you have entered all the names, end with a new line containing a " +
          "single period. \n")
else:
    runAbsentSurvey = False

absentTonight = []

while runAbsentSurvey:
    name = raw_input()
    #msg = ("Please enter the name of the absent brother. "+
    #    "If this is the last absent brother, end with a period.")
    #ttl = "Absent Brothers"
    #name = eg.enterbox(msg, ttl)
    if name[0] == ".":
        last = True
    else:
        last = False
    if last:
        runAbsentSurvey = False
    else:
        absentTonight += [name]

print("Creating the schedule now.")

# ----------------------- End of Ryan's First Addition ------------------------

absent = permaAbsent + absentTonight

# Create a subset of brothers that can do work during social event
eligibleBros = list(set(epsilon + zeta + eta + theta + iota + kappa))

# -------------------------- Ryan's next addition -----------------------------
# takes into account class in the likelihood of selection
# also writes to excel file

# List of brothers who are good at door
'''brothers_good_at_door = ["Curtis Siegfried",
                         "Taylor Ferguson",
                         "Evan Mason",
                         "Chris Farmer",
                         "Zachary Hawtof",
                         "Ian Mason",
                         "Elliot Surovell",
                         "Ryan Flynn",
                         "Jack Hendershott",
                         "Han Li",
                         "Matthew Buckley",
                         "Erik Bartlett",
                         "Kyle Joyner",
                         "Richard Mercer",
                         "Andrew Soncrant",
                         "Joey Papador",
                         "Christian Collins",
                         "Anand Dharia",
                         "Francisco Torres",
                         "Nick Alaverdyan",
                         ]'''

# List of those available by class who are also not in brothers_good_at_door
available_kappa = [kappa_mem[0] for kappa_mem in kappa if kappa_mem[0] not in absent]
#index for pledges since they're selected no matter what
available_iota = [iota_mem for iota_mem in iota if iota_mem not in absent]
available_theta = [theta_mem for theta_mem in theta if theta_mem not in absent]
available_eta = [eta_mem for eta_mem in eta if eta_mem not in absent]
available_zeta = [zeta_mem for zeta_mem in zeta if zeta_mem not in absent]
available_epsilon = [epsilon_mem for epsilon_mem in epsilon if epsilon_mem not in absent]
all_people = [available_iota, available_theta, available_eta, available_zeta, available_epsilon]

# Get phone numbers for all brothers not absent
f = open('PhiPsi-Contact-list.csv')
lines = [line for line in f]
f.close()
lines = lines[1:]
file_lasts = {}
for line in lines:
    line = line.split(',')
    name = line[1].lower()
    number = line[3].replace('-', '')
    file_lasts[name] = number.strip()

# Picks pledges out of available_pledges now so that they will not be given 
# multiple shifts later
doorShift2 = pickRandomBros(available_kappa, NUM_SLOTS)
roofShift = pickRandomBros(available_kappa, 4)
if runSpecialScheduler:
    doorShift4 = pickRandomBros(available_kappa, NUM_SLOTS)

final_lst = []

# Similar to PickRandomBros, but not destructive of broList
def pickRandomBrosFromClass(broList, numSample):
    numSample = min(numSample, len(broList))
    randomLst = random.sample(broList, numSample)
    return randomLst

# Places a number of bros from each class in the final list.  Higher
# classes get less bros (per population) placed in this list
slots_to_fill = NUM_SLOTS * 6 if runSpecialScheduler else NUM_SLOTS * 3
for class_list in all_people:
    if len(final_lst) < slots_to_fill:
        final_lst += pick_bros(class_list, slots_to_fill - len(final_lst))
    else:
        break



doorShift1 = pickRandomBros(final_lst, NUM_SLOTS)
barShift1 = pickRandomBros(final_lst, NUM_SLOTS)
barShift2 = pickRandomBros(final_lst, NUM_SLOTS)
if runSpecialScheduler:
        doorShift3 = pickRandomBros(final_lst, NUM_SLOTS)
        barShift3 = pickRandomBros(final_lst, NUM_SLOTS)
        barShift4 = pickRandomBros(final_lst, NUM_SLOTS)
# ---------------- Writes shift assignments to an excel file ------------------

filename =  "PhiPsiShifts.xlsx"
workbook = xlsxwriter.Workbook(filename)
worksheet = workbook.add_worksheet()
worksheet.set_landscape()

title = workbook.add_format()
title.set_font_size(30)
title.set_bold()
title.set_align('center')
names = workbook.add_format()
names.set_font_size(12)
header = workbook.add_format()
header.set_font_size(14)
header.set_bold()
header.set_bottom(1)
header.set_left(1)
time = workbook.add_format()
time.set_font_size(14)
time.set_bold()
time.set_top(1)
time.set_right(1)
names1 = workbook.add_format()
names1.set_font_size(12)
names1.set_border(1)
names2 = workbook.add_format()
names2.set_font_size(12)
names2.set_top(1)
names3 = workbook.add_format()
names3.set_font_size(12)
names3.set_left(1)
names4 = workbook.add_format()
names4.set_font_size(12)
names4.set_left(1)
names4.set_top(1)
empty = workbook.add_format()
empty.set_border(1)
empty.set_bg_color('gray')
empty1 = workbook.add_format()
empty1.set_top(1)
empty1.set_right(1)
empty1.set_left(1)
empty1.set_bg_color('gray')


worksheet.set_column(0, 0, 1)
worksheet.set_column(1, 1, 15)
worksheet.set_column(2, 2, 20)
worksheet.set_column(3, 3, 20)
worksheet.set_column(4, 4, 20)
worksheet.set_column(5, 5, 20)
worksheet.set_column(6, 6, 20)
if runSpecialScheduler:
    worksheet.set_column(7, 7, 20)
    worksheet.set_column(8, 8, 20)
    worksheet.set_column(9, 9, 20)
    worksheet.set_column(10, 10, 20)

TITLE = '&C&30&"Calibri,Bold"Phi Psi Door and Bar Shift ' + DATE
worksheet.set_header(TITLE)

worksheet.write(3, 1, "10:00-10:30", time)
worksheet.write(4, 1, "10:30-11:00", time)
worksheet.write(5, 1, "11:00-11:30", time)
worksheet.write(6, 1, "11:30-12:00", time)
worksheet.write(7, 1, "12:00-12:30", time)
worksheet.write(8, 1, "12:30-1:00", time)
worksheet.write(9, 1, "1:00-1:30", time)

if runSpecialScheduler:
    worksheet.write(2, 2, "Back Door 1", header)
else:
    worksheet.write(2, 2, "Door 1", header)
row = 3
for bro in doorShift1:
    if row == 9:
        worksheet.write(row, 2, bro, names3)
    else:
        worksheet.write(row, 2, bro, names1)
    row += 1

if runSpecialScheduler:
    worksheet.write(2, 3, "Back Door 2", header)
else:
    worksheet.write(2, 3, "Door 2", header)
row = 3
for bro in doorShift2:
    if row == 9:
        worksheet.write(row, 3, bro, names3)
    else:
        worksheet.write(row, 3, bro, names1)
    row += 1

worksheet.write(2, 4, "Roof Shift", header)
row = 5
for i in range(3, 10):
    if i < 5:
        worksheet.write(i, 4, "", empty)
    elif i >= 5 and i < 9:
        worksheet.write(i, 4, roofShift[i-5], names1)
    else:
        worksheet.write(i, 4, "", empty1)

worksheet.write(2, 5, "Downstairs Bar 1", header)
row = 3
for bro in barShift1:
    if row == 9:
        worksheet.write(row, 5, bro, names3)
    else:
        worksheet.write(row, 5, bro, names1)
    row += 1

worksheet.write(2, 6, "Downstairs Bar 2", header)
row = 3
for bro in barShift2:
    if row == 9:
        worksheet.write(row, 6, bro, names4)
    else:
        worksheet.write(row, 6, bro, names2)
    row += 1

if runSpecialScheduler:
    worksheet.write(2, 7, "Upstairs Bar 1", header)
    row = 3
    for bro in barShift3:
        if row == 9:
                worksheet.write(row, 7, bro, names3)
        else:
                worksheet.write(row, 7, bro, names1)
        row += 1

    worksheet.write(2, 8, "Upstairs Bar 2", header)
    row = 3
    for bro in barShift4:
        if row == 9:
                worksheet.write(row, 8, bro, names3)
        else:
                worksheet.write(row, 8, bro, names1)
        row += 1

    worksheet.write(2, 9, "Courtyard 1", header)
    row = 3
    for bro in doorShift3:
        if row == 9:
                worksheet.write(row, 9, bro, names3)
        else:
                worksheet.write(row, 9, bro, names1)
        row += 1

    worksheet.write(2, 10, "Courtyard 2", header)
    row = 3
    for bro in doorShift4:
        if row == 9:
                worksheet.write(row, 10, bro, names4)
        else:
                worksheet.write(row, 10, bro, names2)
        row += 1

workbook.close()

os.system("open " + filename)

# ------------------------ end of Ryan's contribution -------------------------




