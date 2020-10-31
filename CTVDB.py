import tkinter as tk
from tkinter import ttk
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import os
import re
import sqlite3
from sqlite3 import Error
import json

ctvdbfile = r'\\amat.com\Folders\Austin\Global-Ops\AMO\CPI_TestWorkCntr\TECH FOLDERS\ Irvin Carrillo\Bay8CTV\Bay8CTV.db'
keys= ["A", "B", "C", "D", "E", "F"]

#needed for driver
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def createDriver(view):
    options = Options()
    options.headless = view
    options.add_argument("--window-size=1920,1200")
    DRIVER_PATH = r"./driver/chromedriver.exe"
    return webdriver.Chrome(options=options, executable_path=resource_path(DRIVER_PATH))

def resetDB(conn, keys):
    global bayDict
    genericPO = "No PO Set"
    genericPort = [genericPO]
    genericPortJSON = json.dumps(genericPort)
    bayDict.clear()
    conn.executescript('''DROP TABLE IF EXISTS CTVMATS;
            CREATE TABLE IF NOT EXISTS CTVMATS(
            PORT TEXT PRIMARY KEY     NOT NULL,
            MATS TEXT
            );''')
    for key in keys:
        updateDB(conn, (key, genericPortJSON))
    print("Database reset")

#need to reget dbtodict
def resetAndRefresh(conn, keys, nb, bayDictJSON, frames):
    print("reseting")
    print(bayDict)
    resetDB(conn, keys)
    print(bayDict)
    updateNotebook(nb, json.loads(bayDictJSON), frames)
    
#Print database contents
def get_posts():
    with conn:
        cursor.execute("SELECT * FROM CTVMATS")
        print(cursor.fetchall())

        
#Add or replace row in database
def updateDB(conn, port):
    sql = "REPLACE INTO CTVMATS(PORT, MATS) VALUES (?, ?)"
    cursor = conn.cursor()
    cursor.execute(sql, ((port)))
    conn.commit()
    print("update to " + port[0] + " successful")

#Create dictionary from DB
def dbToDict(cursor, bayDict):
    cursor.execute("SELECT * FROM CTVMATS ORDER BY PORT")
    rows = cursor.fetchall()
    count = 0
    for r in rows:
        for _ in r:
            try:
                x = json.loads(_)
                bayDict[keys[count]] = x
            except:
                pass
        count = count + 1
    dictJSON = json.dumps(bayDict, indent=4, sort_keys=True)
    return dictJSON

#This will contain all relevant details about each material
def createDict(pn, desc, needed, scanned, vf=False):
    x = {
        "PN": pn,
        "DESC": desc,
        "NEEDED": needed,
        "SCANNED": scanned,
        "VERIFIED": vf
    }
    return dict(x)

def getCTVMaterials(po):
    driver = createDriver(True)
    driver.get("http://dca-app-833/LTSWeb/LTSJOBVIEW/AuditViewDownload?id=" + po + "&Offset=300")
    iframe = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, 'body > div > div.content-wrapper > div > iframe')))
    driver.switch_to.frame(iframe)
    table = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[3]/span/div/table/tbody/tr[5]/td[3]/div/div[1]/div/table/tbody/tr/td/table/tbody/tr/td/table/tbody/tr[16]/td[3]/table/tbody')))
    rows = table.find_elements_by_tag_name('tr')
    CTVMaterials = []
    CTVMaterials.append(int(po))
    for r in rows[2:]:
        rowElements = r.find_elements_by_tag_name('td')
        details = []
        for i in [0,2,3,5]:
            details.append(rowElements[i].text)
        CTVMaterials.append(createDict(*details))
    driver.quit()
    return CTVMaterials


def updatePO(poEntry, port):
    global conn
    global cursor
    global bayDict
    global nb
    global frames
    po = poEntry.get()
    print(po)
    if(po.isdigit() and (len(po) == 7)):
        newData = getCTVMaterials(po)
        newDataJSON = json.dumps(newData)
        bayDict[port] = newData
        updateDB(conn, (port, newDataJSON))
        bayDictJSON = dbToDict(cursor, bayDict)
        updateNotebook(nb, json.loads(bayDictJSON), frames)
    else:
        poEntry.delete(0, 'end')
        

def updateNotebook(nb, ctvStatus, frames):
    for f in frames:
        for i in f.winfo_children():
            if type(i) == tk.Label:
                i.destroy()
    i = 0
    rowIdx = 2
    for key in ctvStatus:
        rawText = ""
        for mat in ctvStatus[key]:
            try:
                rawText =mat["PN"] + " - " + mat["DESC"] + "  " + str(int(float(mat["SCANNED"])))  + "/" + str(int(float((mat["NEEDED"])))) + " scanned"
                tk.Label(frames[i], font="Helvetica 14", text=rawText, justify='left', bg=('#04EE0B' if (int(float(mat["SCANNED"])) == int(float(mat["NEEDED"]))) else 'red')).grid(row=rowIdx, column=0,sticky='ew')
                ##May add manual verificaton if it feels useful
##                tk.Checkbutton(frames[i], text = 'Manually Verified', variable = tk.IntVar()).grid(row=rowIdx, column = 1)
                rowIdx = rowIdx + 1
            except TypeError:
                nb.tab(i, text = "     " + key + " - " + str(mat) + "     ")
                pass
        i= i+1
    

#Create main GUI element
def populateNotebook(nb, ctvStatus, frames):
    i = 0
    rowIdx = 2
    for key in keys:
        rawText = ""
        for mat in ctvStatus[key]:
            try:
                rawText =mat["PN"] + " - " + mat["DESC"] + "  " + str(int(float(mat["SCANNED"])))  + "/" + str(int(float((mat["NEEDED"])))) + " scanned"
                tk.Label(frames[i], font="Helvetica 14", text=rawText, justify='left', bg=('#04EE0B' if (int(float(mat["SCANNED"])) == int(float(mat["NEEDED"]))) else 'red')).grid(row=rowIdx, column=0,sticky='ew')
##                tk.Checkbutton(frames[i], text = 'Manually Verified', variable = tk.IntVar()).grid(row=rowIdx, column = 1)
                rowIdx = rowIdx + 1
            except TypeError:
                nb.tab(i, text = "     " + key + " - " + str(mat) + "     ")
                pass
        updatePOEntry = tk.Entry(frames[i])
        updatePOEntry.grid(row=0, column =0, sticky='nsew')
        updateCTVMaterialsButton = tk.Button(frames[i], text="Update CTV Materials", command= (lambda x = updatePOEntry, y = key: updatePO(x, y)))
        updateCTVMaterialsButton.grid(row=1,column=0,sticky='nsew')
        commitToDBButton = tk.Button(frames[i], text="Push to DB")
        commitToDBButton.grid(row=1, column=1, sticky='nsew')
        i=i+1



#Connect to database, create table if it does not exist and if
#table existsed, pull data into dictionary
conn = sqlite3.connect(ctvdbfile)
cursor = conn.cursor()
print("connected")
bayDict = dict()
conn.execute('''CREATE TABLE IF NOT EXISTS CTVMATS(
        PORT TEXT PRIMARY KEY     NOT NULL,
        MATS TEXT
        );''')
#resetDB(conn, keys)
print("Table exists")
bayDictJSON = dbToDict(cursor, bayDict)
##########################################################-----MAIN------###############################################################################
root = tk.Tk()
root.title("CTV Stuff")
w = 1300
h = 950
ws = root.winfo_screenwidth() # width of the screen
hs = root.winfo_screenheight() # height of the screen
x = (ws/2) - (w/2)
y = (hs/2) - (h/2)
root.geometry('%dx%d+%d+%d' % (w, h, x, y-40))
nb = ttk.Notebook(root, height = 895, width = 1300)
f1 = ttk.Frame(nb)
f2 = ttk.Frame(nb)
f3 = ttk.Frame(nb)
f4 = ttk.Frame(nb)
f5 = ttk.Frame(nb)
f6 = ttk.Frame(nb)
frames = [f1, f2, f3, f4, f5, f6]
for i in range(6):
    nb.add(frames[i], text = "EMPTY")
nb.pack()
resetButton = tk.Button(root, text="RESET DATABASE", command=(lambda:resetAndRefresh(conn, keys, nb, bayDictJSON, frames)))
resetButton.pack()
populateNotebook(nb, json.loads(bayDictJSON), frames)
root.mainloop()


#####################################################33OLD TEST STUFF############################################################################
#Replace this with material list getter, example data
##aPO = 12345
##bPO = 23456
#genericPO = 99999
##porta = [aPO, createDict("0000-00000", "some part", "1", "1", True), createDict("1111-11111", "some other part", "5", "0", False)]
##portb =[bPO, createDict("2222-22222", "yet another part", "3", "3", True)]
#genericPort = [genericPO]
##portajson = json.dumps(porta)
##portbjson = json.dumps(portb)
#genericPortJSON = json.dumps(genericPort)
##bay9 = {
##        "9A":[],
##        "9B":[]
##    }
##bayDict = dict(bay9)
#bayDict = dict()



#get_posts()
##updateDB(conn, ("9A", portajson))
##updateDB(conn, ("9B", portbjson))
#updateDB(conn, ("9A", genericPortJSON))
##updateDB(conn, ("9B", genericPortJSON))
##updateDB(conn, ("9C", genericPortJSON))
##updateDB(conn, ("9D", genericPortJSON))
##updateDB(conn, ("9E", genericPortJSON))
##updateDB(conn, ("9F", genericPortJSON))
##updateDB(conn, ("9G", genericPortJSON))
##updateDB(conn, ("9H", genericPortJSON))


##Create new table each time for debugging
##conn.executescript('''DROP TABLE IF EXISTS CTVMATS;
##        CREATE TABLE IF NOT EXISTS CTVMATS(
##        PORT TEXT PRIMARY KEY     NOT NULL,
##        MATS TEXT
##        );''')
