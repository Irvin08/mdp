##Changes needed:
##Check for most recent leads passdown
##add mayan test percent checker
##expand on chamber class
##refactor to use chamber class when possibe
##Look into SAP connection
##Look into EQRK status

import pandas as pd
from datetime import datetime, timedelta
from PIL import Image, ImageDraw
import tkinter as tk
from tkinter import *
from PIL import ImageTk
import openpyxl
from openpyxl import load_workbook
import webbrowser

class Chamber:
    def __init__ (self, system, po):
        self.system = system
        self.po = po

    def __str__(self):
        return 'System #:{}, PO #: {}'.format(self.system,self.po)

    def print(self):
        print(self.system + " " + self.po)


def getPOs(file, chambers):
    data = pd.read_excel(file, dtype=str)
    df = pd.DataFrame(data, columns= ['SYSTEM', 'CH PO'])
    for x in range(0,120,2):
        if "EMPTY" in str(df.at[(x), 'SYSTEM']):
            system_num = None
            po_num = None
        else:
            system_num = df.at[x, 'SYSTEM']
            po_num = df.at[x, 'CH PO']
        chambers.append(Chamber(system_num, po_num))


def getPriorityColors(file, cells):
    wb = load_workbook(file, data_only = True)
    sh = wb['Lead Passdown']
    for x in range (61):
        color_in_hex = sh["A" + str(x+1)].fill.start_color.index
        if color_in_hex != 0:
            try:
                rgb = tuple(int(color_in_hex[i:i+2], 16) for i in ( 2, 4, 6))
                cells.append(rgb)
            except TypeError:
                rgb = (255, 255, 255)
                cells.append(rgb)
        else:
            cells.append((255, 255, 255))
    del cells[0]


def findChamberLocations(bay):
    global chamber_locations
    global df
    for x in range (6):
        if "EMPTY" in df.at[x + (6 * (bay - 1)),'System #']:
            chamber_locations[x] = 0
        else:
            chamber_locations[x] = 1

    return chamber_locations


def create_buttons(root, chamber_image, chamber_locations, active_buttons):
    if chamber_locations[0] == 1:
        portA_button = tk.Button(root, image=chamber_image, command=(lambda: create_window_Generic(0)), anchor = "w")
        portA_button_window = canvas.create_window(890,75, anchor= "nw", window = portA_button)
        active_buttons.append(portA_button)

    if chamber_locations[1] == 1:
        portB_button = tk.Button(root, image=chamber_image, command=(lambda: create_window_Generic(1)), anchor = "w")
        portB_button_window = canvas.create_window(890,340, anchor= "nw", window = portB_button)
        active_buttons.append(portB_button)

    if chamber_locations[2] == 1:
        portC_button = tk.Button(root, image=chamber_image, command=(lambda: create_window_Generic(2)), anchor = "w")
        portC_button_window = canvas.create_window(890,605, anchor= "nw", window = portC_button)
        active_buttons.append(portC_button)

    if chamber_locations[5] == 1:
        portF_button = tk.Button(root, image=chamber_image, command=(lambda: create_window_Generic(5)), anchor = "w")
        portF_button_window = canvas.create_window(100,75, anchor= "nw", window = portF_button)
        active_buttons.append(portF_button)

    if chamber_locations[4] == 1:
        portE_button = tk.Button(root, image=chamber_image, command=(lambda: create_window_Generic(4)), anchor = "w")
        portE_button_window = canvas.create_window(100,340, anchor= "nw", window = portE_button)
        active_buttons.append(portE_button)

    if chamber_locations[3] == 1:
        portD_button = tk.Button(root, image=chamber_image, command=(lambda: create_window_Generic(3)), anchor = "w")
        portD_button_window = canvas.create_window(100,605, anchor= "nw", window = portD_button)
        active_buttons.append(portD_button)


def create_window_Generic(x):
    window =tk.Toplevel(root)
    window.geometry('+%d+%d' % (690, 100))
    if cells[x + (6 * (bay_num - 1))] == (0, 0, 0):
        cells[x + (6 * (bay_num - 1))] = (255, 255, 255)
    PriorityLabel = tk.Label(window, background = ("#%02x%02x%02x" % cells[x + (6 * (bay_num - 1))]))
    PriorityLabel.pack()
    Portlabel = tk.Label(window, text = "Port: " + bay_num_str + ports[x])
    Portlabel.pack()
    Systemlabel = tk.Label(window, text = "System #: " + df.at[x + (6 * (bay_num - 1)),'System #'])
    Systemlabel.pack()
    Statuslabel = tk.Label(window, text = "Status of chamber: " + df.at[x + (6 * (bay_num - 1)),'Status Of Chamber'])
    Statuslabel.pack()
    Passdownlabel = tk.Label(window, text = "Passdown issues: " + str(df.at[x + (6 * (bay_num - 1)),'Passdown Issues']))
    Passdownlabel.pack()
    StartDatelabel = tk.Label(window, text = "Start Date: " + str(df.at[x + (6 * (bay_num - 1)),'START Date']))
    StartDatelabel.pack()
    PortDayslabel = tk.Label(window, text = "Port days: " + str(df.at[x + (6 * (bay_num - 1)),'Port Days']))
    PortDayslabel.pack()
    system = df.at[x + (6 * (bay_num - 1)),'System #']
    qn_button = tk.Button(window, text = "view qns", command = (lambda: openqn(system.split('-', 1)[0])))
    qn_button.pack()
    tlc_button = tk.Button(window, text = "TLC", command = (lambda: opentlc(chambers[x + (6 * (bay_num - 1))].po)))
    tlc_button.pack()
    quit_buttonGeneric = tk.Button(window, text = "quit", command = window.destroy)
    quit_buttonGeneric.pack(side = "left")
    window.focus_set()                                                        
    window.grab_set()


def change_bay(root, chamber_image, new_bay, entry, active_buttons):
    global chamber_locations, bay_num_str, bay_num
    entry.delete('0', 'end')
    try:
        bay_num = int(new_bay)
        bay_num_str = new_bay
    except ValueError:
        return
    delete_buttons()
##    bay_num_str = new_bay
##    bay_num = int(bay_num_str)
    chamber_locations = findChamberLocations(bay_num)
    create_buttons(root, chamber_image, chamber_locations, active_buttons)


def delete_buttons():
    global active_buttons, chamber_locations
    x = 0
    for x in active_buttons:
        x.destroy()
    active_buttons.clear()


def openqn(system):
    webbrowser.get(chrome).open_new_tab("http://dca-wb-263/QM/QM/ViewQN?SlotNum=" + str(system) + "&ProdOrder=")


def opentlc(po):
    webbrowser.get(chrome).open_new_tab("http://ioms/MFG/ModuleStatus?PO=" + po + "#!/laborcosting")


#def updateLastTestRan(bay_num_str, port):
    # TODO: implement 


#def getEQRKStatus(system, chamber):
    # TODO: implement
    
############################################################ MAIN PROGRAM BEGINS ########################################################################
d = (datetime.today() - timedelta(1)).strftime('%m-%d-%y')
print(d)
POfile = r'\\amat.com\Folders\Austin\Global-Ops\AMO\CPI_TestWorkCntr\TECH FOLDERS\ Irvin Carrillo\PO#.xlsx'
file = r'\\amat.com\Folders\Austin\Global-Ops\AMO\CPI_TestWorkCntr\SUPERVISOR PASSDOWN\LEADS Passdown\LEADS PASSDOWN ' + d + 'Nite.xlsx'
chrome = "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s"
chamber_image_file = r'\\amat.com\Folders\Austin\Global-Ops\AMO\CPI_TestWorkCntr\TECH FOLDERS\ Irvin Carrillo\chamber.png'
bay_image_file = r'\\amat.com\Folders\Austin\Global-Ops\AMO\CPI_TestWorkCntr\TECH FOLDERS\ Irvin Carrillo\baydrawing.png'
chambers = []
cells = []
active_buttons = []
chamber_locations = [1,1,1,1,1,1]
ports = ["A","B","C","D","E","F"]
getPOs(POfile, chambers)
getPriorityColors(file, cells)

while True:
    bay_num_str = input('Enter bay: ')
    try:
        bay_num = int(bay_num_str)
    except ValueError:
       print('That\'s not a number!')
    else:
        if 1 <= int(bay_num_str) <= 10:
            break
        else:
            print("That\'s not a real bay, try again.")

data = pd.read_excel(file)
df = pd.DataFrame(data, columns= ['Bay ','System #', 'Status Of Chamber', 'Passdown Issues','START Date','Port Days'])
chamber_locations = findChamberLocations(bay_num)
print(chamber_locations)
data = pd.read_excel(file)
df = pd.DataFrame(data, columns= ['Bay ','System #', 'Status Of Chamber', 'Passdown Issues','START Date','Port Days'])
root = tk.Tk()
root.title("Bay Status")
w = 1154
h = 881
ws = root.winfo_screenwidth() # width of the screen
hs = root.winfo_screenheight() # height of the screen
x = (ws/2) - (w/2)
y = (hs/2) - (h/2)
root.geometry('%dx%d+%d+%d' % (w, h, x, y-40))
canvas = Canvas(root, width=1154, height=881)
bay_image = ImageTk.PhotoImage(file = bay_image_file)
new_bay_entry = tk.Entry(root)
new_bay_entry.pack()
refresh_button = tk.Button(root, text = "change bay", command = (lambda: change_bay(root, chamber_image, new_bay_entry.get(), new_bay_entry, active_buttons)))#locations
refresh_button.pack()
canvas.create_image(575,450, image = bay_image)
canvas.pack()
chamber_image = PhotoImage(file = chamber_image_file)

create_buttons(root, chamber_image, chamber_locations, active_buttons)
    

root.mainloop()

