#Irvin Carrillo 8/13/2020
#Tool to get QN, Inspection Lot, and ESW data
#for a list of PO's. Meant to be used
#by DTF that is looking for chambers to cross to MDP test.
#Outputs format requested by DTF

import tkinter as tk
from selenium import webdriver
import pandas as pd
from selenium.webdriver.chrome.options import Options
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import os
import inspect
import re
import time
#Classes
###############################################################################################################################
class QN:
    def __init__ (self, QNNum, Type, Desc, isOpen):
        self.QNNum = QNNum
        self.Type = Type
        self.Desc = Desc
        self.isOpen = isOpen
        
    def isOpen(self):
        return self.isOpen

    def print(self):
        global result
        result = result + (self.Type + " " + self.QNNum + " - " + self.Desc.upper() + ("\n" if self.isOpen else " | CLOSED\n"))


class InspLot:
    def __init__ (self, LotNum, Desc, Status, isOpen):
        self.LotNum = LotNum
        self.Desc = Desc
        self.Status = Status
        self.isOpen = isOpen

    def isOpen(self):
        return self.isOpen#Irvin Carrillo 8/13/2020
#Tool to get QN, Inspection Lot, and ESW data
#for a list of PO's. Meant to be used
#by DTF that is looking for chambers to cross to MDP test.
#Outputs format requested by DTF

import tkinter as tk
from selenium import webdriver
import pandas as pd
from selenium.webdriver.chrome.options import Options
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import os
import inspect
import re
import time
#Classes
###############################################################################################################################
class QN:
    def __init__ (self, QNNum, Type, Desc, isOpen):
        self.QNNum = QNNum
        self.Type = Type
        self.Desc = Desc
        self.isOpen = isOpen
        
    def isOpen(self):
        return self.isOpen

    def print(self):
        global result
        result = result + (self.Type + " " + self.QNNum + " - " + self.Desc.upper() + ("\n" if self.isOpen else " | CLOSED\n"))


class InspLot:
    def __init__ (self, LotNum, Desc, Status, isOpen):
        self.LotNum = LotNum
        self.Desc = Desc
        self.Status = Status
        self.isOpen = isOpen

    def isOpen(self):
        return self.isOpen

    def print(self):
        global result
        result = result + "Insp Lot " + self.LotNum + " - " + self.Desc + ("\n" if self.isOpen else (" | " + self.Status + "\n"))


class ESW:
    def __init__ (self, ESWNum, Desc, Status, isOpen):
        self.ESWNum = ESWNum
        self.Desc = Desc
        self.Status = Status
        self.isOpen = isOpen

    def isOpen(self):
        return self.isOpen()

    def print(self):
        global result
        result = result + "ESW " + self.ESWNum + " - " + self.Desc + ("\n" if self.isOpen else (" | " + self.Status + "\n"))

#Failed attempt to have two expected conditions at the same time
##class wait_for_all(object):
##    def __init__(self, methods):
##        self.methods = methods
##
##    def __call__(self, driver):
##        try:
##            for method in self.methods:
##                if not method(driver):
##                    return False
##            return True
##        except:
##            return False

##    methods = []
##    methods.append(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="inspection"]/div[2]'), "No data available"))
##    methods.append(EC.presence_of_element_located((By.XPATH, '//*[@id="inspection"]/div[1]/div/div/div[2]/div/table/tbody/tr[1]/td[1]/a')))
##    method = wait_for_all(methods)
##    WebDriverWait(driver, 15).until(method)


#Functions
###############################################################################################################################
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def entryToPo(entry):
    POs = []
    POs = re.split("[^0-9]", entry)
    while "" in POs:
        POs.remove("")
    return POs

def printOpen(items):
    global result
    allClosed = True
    for i in items:
        if i.isOpen:
            i.print()
            allClosed = False
    if allClosed:
        result = result + ("There are no open QN's\n" if isinstance(items[0], QN) else\
                            ("There are no open ESWs\n" if isinstance(items[0], ESW) else "There are no open inspection lots\n"))


def printAll(items):
    global result
    for i in items:
        i.print()

        
def getQNs(driver, po, ic):
    global result
    driver.get("http://dca-wb-263/QM/QM/ViewQN?SlotNum=&ProdOrder=" + po)
##    result = result + "\n################################################################################################\n"                                                                                             #\n"
##    result = result + po + "\n"
    try:
        e = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "ui-grid-row")))
        rows = driver.find_elements_by_class_name("ui-grid-row")
        for r in rows:
            qnnum = r.find_element_by_css_selector('a.ng-binding').text
            qnType = r.find_element_by_class_name('ui-grid-coluiGrid-000K').text
            shortText = r.find_element_by_class_name('ui-grid-coluiGrid-000L').text
            status = not r.find_element_by_xpath('.//input[@type="checkbox"]').is_selected()
            qns.append(QN(qnnum, qnType, shortText, status))
        printAll(qns) if ic else printOpen(qns)
        qns.clear()
    except:
        result = result + "No QN data found\n"


def getESWs(driver, po, ic):
    global result
    driver.get("http://ioms/MFG/ModuleStatus?PO=" + po + "#!/ESWs")



    timeout = False
    while(not timeout):
        try:
            text = driver.find_element_by_xpath('//*[@id="ESWs"]/div[2]').text
            if text == "ESW Data is unavailable":
                #print(text)
                timeout = True
                break
        except:
            time.sleep(1)
            pass
        try:
            text = driver.find_element_by_xpath('//*[@id="ESWs"]/div[1]/div/div/div[2]/div/table/tbody/tr')
##            print("found table")
##            print(text)
            break
        except:
            time.sleep(1)
            pass
        #print("nothing found")
    
##  WIP for faster performance when there is no data to retrieve
##    noData = False
##    try:
##        WebDriverWait(driver, 10).until(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="ESWs"]/div[2]'), "ESW Data is unavailable"))
##        noData = True
##    except:
##        pass
##    try:
##        ready = WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ESWs"]/div[1]/div/div/div[2]/div/table/tbody/tr[1]/td[1]/a')))
##    except:
##        ready = False
##    if ready:
####    try:
####        check = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ESWs"]/div[1]/div/div/div[2]/div/table/tbody/tr[1]/td[1]/a')))
    if not timeout:
        body = driver.find_element_by_xpath('//*[@id="ESWs"]/div[1]/div/div/div[2]/div/table/tbody')
        rows = body.find_elements_by_tag_name("tr")
        for r in rows:
            eswNum = r.find_element_by_css_selector('a.ng-binding').text
            eswDesc = r.find_element_by_xpath('.//td[2]').text
            eswStatus = r.find_element_by_xpath('.//td[4]/button').text
            if eswStatus == "Click to Sign":
                eswStatus = "Not Signed"
                eswOpen = True
            else:
                eswOpen = False
            esws.append(ESW(eswNum, eswDesc, eswStatus, eswOpen))
        printAll(esws) if ic else printOpen(esws)
        esws.clear()
    else:
    ##except:
        result = result + "No ESW data found\n"   



def getInspLotsAndBuild(driver, po, ic):
    global result
    driver.get("http://ioms/MFG/ModuleStatus?PO=" + po + "#!/inspection")

    timeout = False
    while(not timeout):
        try:
            text = driver.find_element_by_xpath('//*[@id="inspection"]/div[2]').text
            if text == "No data available":
                #print(text)
                timeout = True
                break
        except:
            time.sleep(1)
            pass
        try:
            text = driver.find_element_by_xpath('//*[@id="inspection"]/div[1]/div/div/div[2]/div/table/tbody/tr')
##            print("found table")
##            print(text)
            break
        except:
            time.sleep(1)
            pass
        #print("noting found")
##    noData = False
##    try:
##        WebDriverWait(driver, 10).until(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="inspection"]/div[2]'), "No data available"))
##        noData = True
##    except:
##        pass
##    try:
##        ready = WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.XPATH, '//*[@id="inspection"]/div[1]/div/div/div[2]/div/table/tbody/tr[1]/td[1]/a')))
##    except:
##        ready = False
##    buildPercent = driver.find_element_by_xpath('//*[@id="Container"]/div[1]/div[3]/div[1]/div/div[3]/div/div/div[1]/span[2]').text
##    if ready:
####    timeout = False
####    try:
####        check = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//*[@id="inspection"]/div[1]/div/div/div[2]/div/table/tbody/tr[1]/td[1]/a')))
####    except:
####        timeout = True
    buildPercent = driver.find_element_by_xpath('//*[@id="Container"]/div[1]/div[3]/div[1]/div/div[3]/div/div/div[1]/span[2]').text
    if not timeout:
        body = driver.find_element_by_xpath('//*[@id="inspection"]/div[1]/div/div/div[2]/div/table/tbody')
        rows = body.find_elements_by_tag_name("tr")
        for r in rows:
            lotNum = r.find_element_by_css_selector('a.ng-binding').text
            lotDesc = r.find_element_by_xpath('.//td[4]').text
            lotStatus = r.find_element_by_xpath('.//td[6]').text
            if lotStatus == "Open":
                lotOpen = True
            else:
                lotOpen = False
            lots.append(InspLot(lotNum, lotDesc, lotStatus, lotOpen))
        printAll(lots) if ic else printOpen(lots)
        lots.clear()
        result = result + buildPercent + "\n"
    else:
        result = result + "No inspection lot data found\n"  + buildPercent + "\n"


def runUpdater(entry, qns, text, qnSelect, eswSelect, ilSelect, allSelect, includeClosed):
    global result
    result = ""
    text.configure(state='normal')
    text.delete(1.0, tk.END)
    r = getResult(entry, qns, qnSelect, eswSelect, ilSelect, allSelect, includeClosed)
    text.insert(tk.END, r)
    text.configure(state='disabled')


def getResult(entry, qns, qnSelect, eswSelect, ilSelect, allSelect, includeClosed):
    global result
    entryStr = entry.get()
    entry.delete('0', 'end')
    options = Options()
    options.headless = True # set to False to see chrome window while running
    options.add_argument("--window-size=1920,1200")
    DRIVER_PATH = r"./driver/chromedriver.exe"
    driver = webdriver.Chrome(options=options, executable_path=resource_path(DRIVER_PATH))
    result = result + ("Time started: " + datetime.now().strftime("%m/%d/%Y %H:%M:%S") + "\n")
    POs = entryToPo(entryStr)
    for po in POs:
        print(po)
    for po in POs:
        if allSelect:
            qnSelect = True
            eswSelect = True
            ilSelect = True
        result = result + "\n################################################################################################\n"                                                                                             #\n"
        result = result + po + "\n"
        if qnSelect:
            getQNs(driver, po, includeClosed)
        if eswSelect:
            getESWs(driver, po, includeClosed)
        if ilSelect:
            getInspLotsAndBuild(driver, po, includeClosed)
    result = result + ("\nTime finished: " + datetime.now().strftime("%m/%d/%Y %H:%M:%S") + "\n")
    driver.quit()
    #print("Time finished: " + datetime.now().strftime("%m/%d/%Y %H:%M:%S"))
    return result

######################################################################             MAIN                        ####################################################
qns = []
lots = []
esws = []
result = ""
root = tk.Tk()
root.title("Updater")
w = 1300
h = 950
ws = root.winfo_screenwidth() # width of the screen
hs = root.winfo_screenheight() # height of the screen
x = (ws/2) - (w/2)
y = (hs/2) - (h/2)
root.geometry('%dx%d+%d+%d' % (w, h, x, y-40))
#creditLabel = tk.Label(root, text = "v1.0 Irvin Carrillo").grid(row = 2, column = 0, sticky = "sw")
POLabel = tk.Label(root, text = "Copy PO's here:").grid(row = 0, column = 0, sticky = "e")
POEntry = tk.Entry(root)
POEntry.grid(row = 0, column = 1, sticky = "nsew")
text = tk.Text(root, width = 160, height = 55)
text.grid(row = 1, column = 0, columnspan = 8, sticky = "nsew")
qnSelect = tk.IntVar()
tk.Checkbutton(root, text="Include QNs", variable=qnSelect).grid(row = 0, column = 3, sticky = "nsew")
eswSelect = tk.IntVar()
tk.Checkbutton(root, text="Include ESWs", variable=eswSelect).grid(row = 0, column = 4, sticky = "nsew")
ilSelect = tk.IntVar()
tk.Checkbutton(root, text="Include Insp Lots & Build %", variable=ilSelect).grid(row = 0, column = 5, sticky = "nsew")
allSelect = tk.IntVar()
tk.Checkbutton(root, text="Include all data", variable=allSelect).grid(row = 0, column = 6, sticky = "nsew")
includeClosed = tk.IntVar()
tk.Checkbutton(root, text="Include closed", variable=includeClosed).grid(row = 0, column = 7, sticky = "nsew")
QNCheckButton = tk.Button(root, text = "Get Update", command = (lambda: runUpdater(POEntry, qns, text, qnSelect.get(), eswSelect.get(), ilSelect.get(), allSelect.get(), includeClosed.get()))).grid(row = 0, column = 2, sticky = "nsew")
creditLabel = tk.Label(root, text = "IC  v1.4").grid(row = 2, column =7 , sticky = "se")
root.mainloop()



    def print(self):
        global result
        result = result + "Insp Lot " + self.LotNum + " - " + self.Desc + ("\n" if self.isOpen else (" | " + self.Status + "\n"))


class ESW:
    def __init__ (self, ESWNum, Desc, Status, isOpen):
        self.ESWNum = ESWNum
        self.Desc = Desc
        self.Status = Status
        self.isOpen = isOpen

    def isOpen(self):
        return self.isOpen()

    def print(self):
        global result
        result = result + "ESW " + self.ESWNum + " - " + self.Desc + ("\n" if self.isOpen else (" | " + self.Status + "\n"))

#Failed attempt to have two expected conditions at the same time
##class wait_for_all(object):
##    def __init__(self, methods):
##        self.methods = methods
##
##    def __call__(self, driver):
##        try:
##            for method in self.methods:
##                if not method(driver):
##                    return False
##            return True
##        except:
##            return False

##    methods = []
##    methods.append(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="inspection"]/div[2]'), "No data available"))
##    methods.append(EC.presence_of_element_located((By.XPATH, '//*[@id="inspection"]/div[1]/div/div/div[2]/div/table/tbody/tr[1]/td[1]/a')))
##    method = wait_for_all(methods)
##    WebDriverWait(driver, 15).until(method)


#Functions
###############################################################################################################################
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def entryToPo(entry):
    POs = []
    POs = re.split("[^0-9]", entry)
    while "" in POs:
        POs.remove("")
    return POs

def printOpen(items):
    global result
    allClosed = True
    for i in items:
        if i.isOpen:
            i.print()
            allClosed = False
    if allClosed:
        result = result + ("There are no open QN's\n" if isinstance(items[0], QN) else\
                            ("There are no open ESWs\n" if isinstance(items[0], ESW) else "There are no open inspection lots\n"))


def printAll(items):
    global result
    for i in items:
        i.print()

        
def getQNs(driver, po, ic):
    global result
    driver.get("http://dca-wb-263/QM/QM/ViewQN?SlotNum=&ProdOrder=" + po)
    result = result + "\n################################################################################################\n"                                                                                             #\n"
    result = result + po + "\n"
    try:
        e = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "ui-grid-row")))
        rows = driver.find_elements_by_class_name("ui-grid-row")
        for r in rows:
            qnnum = r.find_element_by_css_selector('a.ng-binding').text
            qnType = r.find_element_by_class_name('ui-grid-coluiGrid-000K').text
            shortText = r.find_element_by_class_name('ui-grid-coluiGrid-000L').text
            status = not r.find_element_by_xpath('.//input[@type="checkbox"]').is_selected()
            qns.append(QN(qnnum, qnType, shortText, status))
        printAll(qns) if ic else printOpen(qns)
        qns.clear()
    except:
        result = result + "No QN data found\n"


def getESWs(driver, po, ic):
    global result
    driver.get("http://ioms/MFG/ModuleStatus?PO=" + po + "#!/ESWs")
##  WIP for faster performance when there is no data to retrieve
##    noData = False
##    try:
##        WebDriverWait(driver, 10).until(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="ESWs"]/div[2]'), "ESW Data is unavailable"))
##        noData = True
##    except:
##        pass
##    try:
##        ready = WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ESWs"]/div[1]/div/div/div[2]/div/table/tbody/tr[1]/td[1]/a')))
##    except:
##        ready = False
##    if ready:
    try:
        check = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//*[@id="ESWs"]/div[1]/div/div/div[2]/div/table/tbody/tr[1]/td[1]/a')))
        body = driver.find_element_by_xpath('//*[@id="ESWs"]/div[1]/div/div/div[2]/div/table/tbody')
        rows = body.find_elements_by_tag_name("tr")
        for r in rows:
            eswNum = r.find_element_by_css_selector('a.ng-binding').text
            eswDesc = r.find_element_by_xpath('.//td[2]').text
            eswStatus = r.find_element_by_xpath('.//td[4]/button').text
            if eswStatus == "Click to Sign":
                eswStatus = "Not Signed"
                eswOpen = True
            else:
                eswOpen = False
            esws.append(ESW(eswNum, eswDesc, eswStatus, eswOpen))
        printAll(esws) if ic else printOpen(esws)
        esws.clear()
    #else:
    except:
        result = result + "No ESW data found\n"   



def getInspLotsAndBuild(driver, po, ic):
    global result
    driver.get("http://ioms/MFG/ModuleStatus?PO=" + po + "#!/inspection")
##    noData = False
##    try:
##        WebDriverWait(driver, 10).until(EC.text_to_be_present_in_element((By.XPATH, '//*[@id="inspection"]/div[2]'), "No data available"))
##        noData = True
##    except:
##        pass
##    try:
##        ready = WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.XPATH, '//*[@id="inspection"]/div[1]/div/div/div[2]/div/table/tbody/tr[1]/td[1]/a')))
##    except:
##        ready = False
##    buildPercent = driver.find_element_by_xpath('//*[@id="Container"]/div[1]/div[3]/div[1]/div/div[3]/div/div/div[1]/span[2]').text
##    if ready:
    timeout = False
    try:
        check = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.XPATH, '//*[@id="inspection"]/div[1]/div/div/div[2]/div/table/tbody/tr[1]/td[1]/a')))
    except:
        timeout = True
    buildPercent = driver.find_element_by_xpath('//*[@id="Container"]/div[1]/div[3]/div[1]/div/div[3]/div/div/div[1]/span[2]').text
    if not timeout:
        body = driver.find_element_by_xpath('//*[@id="inspection"]/div[1]/div/div/div[2]/div/table/tbody')
        rows = body.find_elements_by_tag_name("tr")
        for r in rows:
            lotNum = r.find_element_by_css_selector('a.ng-binding').text
            lotDesc = r.find_element_by_xpath('.//td[4]').text
            lotStatus = r.find_element_by_xpath('.//td[6]').text
            if lotStatus == "Open":
                lotOpen = True
            else:
                lotOpen = False
            lots.append(InspLot(lotNum, lotDesc, lotStatus, lotOpen))
        printAll(lots) if ic else printOpen(lots)
        lots.clear()
        result = result + buildPercent + "\n"
    else:
        result = result + "No inspection lot data found\n"  + buildPercent + "\n"


def runUpdater(entry, qns, text, allSelect, includeClosed):
    global result
    result = ""
    text.configure(state='normal')
    text.delete(1.0, tk.END)
    r = getResult(entry, qns, allSelect, includeClosed)
    text.insert(tk.END, r)
    text.configure(state='disabled')


def getResult(entry, qns, allSelect, includeClosed):
    global result
    entryStr = entry.get()
    entry.delete('0', 'end')
    options = Options()
    options.headless = True # set to False to see chrome window while running
    options.add_argument("--window-size=1920,1200")
    DRIVER_PATH = r"./driver/chromedriver.exe"
    driver = webdriver.Chrome(options=options, executable_path=resource_path(DRIVER_PATH))
    result = result + ("Time started: " + datetime.now().strftime("%m/%d/%Y %H:%M:%S") + "\n")
    POs = entryToPo(entryStr)
    for po in POs:
        print(po)
    for po in POs:
        getQNs(driver, po, includeClosed)
        if allSelect:
            getESWs(driver, po, includeClosed)
            getInspLotsAndBuild(driver, po, includeClosed)
    result = result + ("\nTime finished: " + datetime.now().strftime("%m/%d/%Y %H:%M:%S") + "\n")
    driver.quit()
    #print("Time finished: " + datetime.now().strftime("%m/%d/%Y %H:%M:%S"))
    return result

######################################################################             MAIN                        ####################################################
qns = []
lots = []
esws = []
result = ""
root = tk.Tk()
root.title("Updater")
w = 1300
h = 950
ws = root.winfo_screenwidth() # width of the screen
hs = root.winfo_screenheight() # height of the screen
x = (ws/2) - (w/2)
y = (hs/2) - (h/2)
root.geometry('%dx%d+%d+%d' % (w, h, x, y-40))
#creditLabel = tk.Label(root, text = "v1.0 Irvin Carrillo").grid(row = 2, column = 0, sticky = "sw")
POLabel = tk.Label(root, text = "Copy PO's here:").grid(row = 0, sticky = "e")
POEntry = tk.Entry(root)
POEntry.grid(row = 0, column = 1, sticky = "nsew")
text = tk.Text(root, width = 160, height = 55)
text.grid(row = 1, column = 0, columnspan = 5, sticky = "nsew")
select = tk.IntVar()
tk.Checkbutton(root, text="Include all data", variable=select).grid(row = 0, column = 3, sticky = "nsew")
includeClosed = tk.IntVar()
tk.Checkbutton(root, text="Include closed", variable=includeClosed).grid(row = 0, column = 4, sticky = "nsew")
QNCheckButton = tk.Button(root, text = "Get Update", command = (lambda: runUpdater(POEntry, qns, text, select.get(), includeClosed.get()))).grid(row = 0, column = 2, sticky = "nsew")
creditLabel = tk.Label(root, text = "By Irvin Carrillo  v1.3").grid(row = 2, column =4 , sticky = "se")
root.mainloop()
