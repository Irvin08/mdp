#Irvin Carrillo 6/18/2020
#This program uses the Test Queue excel file to find the most recently crossed
#chambers to MDP Test and makes a list of the chambers along with relevant info.

import pandas as pd
from datetime import datetime, timedelta

class Port:
    def __init__ (self, portName, occupied, chamber, chamberPO, chamberType):
        self.portName = portName
        self.occupied = occupied
        self.chamber = chamber
        self.chamberPO = chamberPO
        self.chamberType = chamberType

    def print(self):
        print(self.portName + ' ' + self.chamber + ' ' +self.chamberPO)
               
d = datetime.today().strftime('%#Y  %#m-%#d')
d2 = (datetime.today() - timedelta(1)).strftime('%#Y  %#m-%#d')
try:
    crossoverChecklist = r'\\amat.com\Folders\Austin\Global-Ops\AMO\CPI_TestWorkCntr\SUPERVISOR PASSDOWN\( DTF Checklists for Systems )\( TEST QUEUE - Checklist Forms - Pics )\TEST QUEUE ' + d + ' Day_.xlsx'
except FileNotFoundError:
    pass

try:
    crossoverChecklist = r'\\amat.com\Folders\Austin\Global-Ops\AMO\CPI_TestWorkCntr\SUPERVISOR PASSDOWN\( DTF Checklists for Systems )\( TEST QUEUE - Checklist Forms - Pics )\TEST QUEUE ' + d + ' Day.xlsx'
except FileNotFoundError:
    crossoverChecklist = r'\\amat.com\Folders\Austin\Global-Ops\AMO\CPI_TestWorkCntr\SUPERVISOR PASSDOWN\( DTF Checklists for Systems )\( TEST QUEUE - Checklist Forms - Pics )\TEST QUEUE ' + d2 + ' Night.xlsx'
    
#From test queue excel file, make a dataframe that includes location, slot#, chamber PO#, and chamber type
data = pd.read_excel(crossoverChecklist, sheet_name= 'QUEUE', usecols = 'E:H', dtype=str, skiprows = 3)
df = pd.DataFrame(data)

mdpTest = []
floor = ['1A', '1B', '1C', '1D', '1E', '1F',
         '2A', '2B', '2C', '2D', '2E', '2F',
         '3A', '3B', '3C', '3D', '3E', '3F',
         '4A', '4B', '4C', '4D', '4E', '4F',
         '5A', '5B', '5C', '5D', '5E', '5F',
         '6A', '6B', '6C', '6D', '6E', '6F',
         '7A', '7B', '7C', '7D', '7E', '7F',
         '8A', '8B', '8C', '8D', '8E', '8F',
         '9A', '9B', '9C', '9D', '9E', '9F',
         '10A', '10B', '10C', '10D', '10E', '10F']

#Initialize empty ports
for i in range(60):
    mdpTest.append(Port(floor[i], False, 'XXXXX-X', 'XXXXX', 'XXXXX'))


#Sarting from the end, look for chambers in queue that were moved to a test port and add them to mdpTest list
#Note lastInQ variable is the length of the dataframe - 5 because we skipped the first 3 rows and
#the excel file has 2 empty rows at the end which means the total size of the sheet is 5 more than our dataframe.
count = 0
lastInQ = len(df.index) - 5
while count < 60:
    loc = df.at[lastInQ, 'Location']
    if loc in floor:
        for p in mdpTest:
            if loc == p.portName and not p.occupied:
                p.occupied = True
                p.chamber = df.at[lastInQ, 'Slot /Sys - Ch# ']
                p.chamberPO = df.at[lastInQ, 'Build PO#']
                p.chamberType = df.at[lastInQ, 'CH Type']
                break
        count+=1
        lastInQ-=1
    else:
        lastInQ-=1
        
for p in mdpTest:
    p.print()
