########################CHAMBER CLASS TO CONTAIN ALL CHAMBER RLEVANT INFO###########################################
import pandas as pd

class Chamber:
    def __init__ (self, system, po):
        self.system = system
        self.po = po

    def __str__(self):
        return 'System #:{}, PO #: {}'.format(self.system,self.po)

    def print(self):
        print(self.system + " " + self.po)

file = r'\\amat.com\Folders\Austin\Global-Ops\AMO\CPI_TestWorkCntr\TECH FOLDERS\ Irvin Carrillo\PO#.xlsx'
data = pd.read_excel(file, dtype=str)
df = pd.DataFrame(data, columns= ['SYSTEM', 'CH PO'])
chambers = []

for x in range(0,120,2):
    if "EMPTY" in str(df.at[(x), 'SYSTEM']):
        system_num = None
        po_num = None
    else:
        system_num = df.at[x, 'SYSTEM']
        po_num = df.at[x, 'CH PO']
    chambers.append(Chamber(system_num, po_num))

for ch in chambers:
    print(ch)
