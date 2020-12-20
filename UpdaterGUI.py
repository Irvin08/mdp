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
import json
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


class wait_for_text_to_appear(object):
    def __init__(self, locator,):
        self.locator = locator

    def __call__(self, driver):
        try:
            element_text = EC._find_element(driver, self.locator).text
            return len(element_text) > 1
        except StaleElementReferenceException:
            return False

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

def entryToSystem(entry):
    SysAndPO = []
    systems = []
    SysAndPO = entry.split()
    print(SysAndPO)
    for s in SysAndPO:
        if "-" in s:
           systems.append(s.split("-")[0])
    print(systems)
    return systems

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

##def getRCTD(driver, system):
##    global result
##    driver.get("http://ioms/Engg/SystemSearch?SystemNo="+ system)
##    e = WebDriverWait(driver, 120).until(wait_for_text_to_appear((By.CSS_SELECTOR, "#lblprojRCTD")))
##    x = driver.find_element_by_css_selector("#headertbl > div:nth-child(7) > div > div:nth-child(2)")
##    dates = x.find_elements_by_class_name("list-group-item")
##    for _ in dates:
##        print(_.text)
##        result = result + _.text + "\n"

def getRCTD(driver, toolId):
    global result
    RCTD_P = ""
    RCTD_R = ""
    SHIP_P = ""
    SHIP_R = ""
    driver.get("http://dca-wb-281/PROMPT/api/SystemViewAPI/GetToolGridDetails?ToolId="+ toolId)
    try:
        RCTD_P = driver.find_element_by_css_selector("#folder7 > div.opened > div:nth-child(5) > span:nth-child(2)").text
        RCTD_R = driver.find_element_by_css_selector("#folder7 > div.opened > div:nth-child(6) > span:nth-child(2)").text
    except:
        pass
    try:
        SHIP_P = driver.find_element_by_css_selector("#folder8 > div.opened > div:nth-child(5) > span:nth-child(2)").text
        SHIP_R = driver.find_element_by_css_selector("#folder8 > div.opened > div:nth-child(6) > span:nth-child(2)").text
    except:
            pass
    result = result + ((SHIP_R[0:10] if (len(SHIP_R) > 1) else SHIP_P[0:10]) + "\t" + (RCTD_R[0:10] if (len(RCTD_R) > 1) else RCTD_P[0:10]) + "\n")
    print((SHIP_R[0:10] if (len(SHIP_R) > 1) else SHIP_P[0:10]) + "\t" + (RCTD_R[0:10] if (len(RCTD_R) > 1) else RCTD_P[0:10]))



#Get QNs and ILs from PROMT JSON
def getQNsAndInspLotsJSON(driver, po, ic):
    global result
    try:
        driver.get("http://webmprd:3333/rest/AMAT_iOMS_SSG/api/getQN_LOT_DataFromSAP?ORD_NUMBER=" + po)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, "body > pre")))
        x = driver.find_element_by_css_selector("body > pre").text
        y = json.loads(x)
        #print(json.dumps(y, indent=4, sort_keys=True))
        #Start QNs
        if y["SAP_QN_DATA_RESPONSE"]["RESPONSE"]["QL_COUNT"] != "0":
            for q in y["SAP_QN_DATA_RESPONSE"]["RESPONSE"]["QN_DATA"]["QLNOTE"]:
                status = True if (q["ZZIMMFIX"] != "true") else False
                qns.append(QN(q["QMNUM"], q["QMART"], q["QMTXT"], status))
            printAll(qns) if ic else printOpen(qns)
            qns.clear()
        else:
            result = result + ("No QN data found\n")
            
        #Start Insp Lots
        if y["SAP_QN_DATA_RESPONSE"]["RESPONSE"]["INSLOT_COUNT"] != "0":
            for i in y["SAP_QN_DATA_RESPONSE"]["RESPONSE"]["INSPECTION_DATA"]["INSLOT"]:
                status = True if (i["VBEWERTUNG"] != "A") else False
                lots.append(InspLot(i["PRUEFLOS"], i["KTEXT"], i["VBEWERTUNG"], status))
            printAll(lots) if ic else printOpen(lots)
            lots.clear()
        else:
            result = result + ("No inspection lot data found" + "\n")
        #print("PO " + po + " done")
    except:
        result = result + "ERROR on PO: " + po

#Get QNs from PROMPT
def getQNs(driver, po, ic):
    global result
    driver.get("http://dca-wb-281/QM/QM/ViewQN?SlotNum=&ProdOrder=" + po)
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

#Get InspLots from PROMPT
def getInspLots(driver, po, ic):
    global result
    qnCount = 0
    lots = []
    driver.get("http://dca-wb-281/QM/QM/ViewQN?SlotNum=&ProdOrder=" + po)
    try:
        e = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "ui-grid-row")))
        rows = driver.find_elements_by_class_name("ui-grid-row")
        for r in rows:
            qnCount = qnCount + 1
    except:
        print("error on qn page")
    driver.find_element_by_xpath("/html/body/div[1]/div[1]/section/section/ul/li[2]/a").click()
    time.sleep(3)
    try:
        e = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CLASS_NAME, "ui-grid-row")))
        rows = driver.find_elements_by_class_name("ui-grid-row")
        if len(rows) != qnCount:
            for r in rows:
                try:
                    lotDesc = r.find_element_by_class_name('ui-grid-coluiGrid-001G').text#was b
                    #print(lotDesc)
                    lotNum = r.find_element_by_css_selector('a.ng-binding').text
                    lotStatus = r.find_element_by_class_name('ui-grid-coluiGrid-001K').text#wasf
                    if lotStatus == "A":
                        lotOpen = False
                    else:
                        lotOpen = True
                    lots.append(InspLot(lotNum, lotDesc, lotStatus, lotOpen))
                except:
                    continue
        else:
            result = result + "No inspection lot data found\n"
    except:
        result = result + "No inspection lot data found\n"
    if lots:
        printAll(lots) if ic else printOpen(lots)
        lots.clear()



#Get QNs and InspLots from PROMPT
def getQNsAndInspLots(driver, po, ic):
    global result
    qns = []
    lots = []
    qnCount = 0
    driver.get("http://dca-wb-281/QM/QM/ViewQN?SlotNum=&ProdOrder=" + po)
    try:
        e = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.CLASS_NAME, "ui-grid-row")))
        rows = driver.find_elements_by_class_name("ui-grid-row")
        for r in rows:
            qnnum = r.find_element_by_css_selector('a.ng-binding').text
            qnType = r.find_element_by_class_name('ui-grid-coluiGrid-000K').text
            shortText = r.find_element_by_class_name('ui-grid-coluiGrid-000L').text
            status = not r.find_element_by_xpath('.//input[@type="checkbox"]').is_selected()
            qns.append(QN(qnnum, qnType, shortText, status))
            qnCount = qnCount + 1
        printAll(qns) if ic else printOpen(qns)
        qns.clear()
    except:
        result = result + "No QN data found\n"
    driver.find_element_by_xpath("/html/body/div[1]/div[1]/section/section/ul/li[2]/a").click()
    time.sleep(3)
    try:
        e = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.CLASS_NAME, "ui-grid-row")))
        rows = driver.find_elements_by_class_name("ui-grid-row")
        if len(rows) != qnCount:
            for r in rows:
                try:
                    lotDesc = r.find_element_by_class_name('ui-grid-coluiGrid-001G').text
                    lotNum = r.find_element_by_css_selector('a.ng-binding').text
                    lotStatus = r.find_element_by_class_name('ui-grid-coluiGrid-001K').text
                    if lotStatus == "A":
                        lotOpen = False
                    else:
                        lotOpen = True
                    lots.append(InspLot(lotNum, lotDesc, lotStatus, lotOpen))
                except:
                    continue
        else:
            result = result + "No inspection lot data found\n"
    except:
        result = result + "No inspection lot data found\n"
    if qns:
        printAll(qns) if ic else printOpen(qns)
        qns.clear()
    if lots:
        printAll(lots) if ic else printOpen(lots)
        qns.clear()




#Get ESWs from iOMS
def getESWs(driver, po, ic):
    global result
    driver.get("http://ioms/MFG/ModuleStatus?PO=" + po + "#!/ESWs")
    timeout = False
    while(not timeout):
        try:
            text = driver.find_element_by_xpath('//*[@id="ESWs"]/div[2]').text
            if text == "ESW Data is unavailable":
                timeout = True
                break
        except:
            time.sleep(1)
            pass
        try:
            text = driver.find_element_by_xpath('//*[@id="ESWs"]/div[1]/div/div/div[2]/div/table/tbody/tr')
            break
        except:
            time.sleep(1)
            pass
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
        result = result + "No ESW data found\n"


#Get ESWs and Build % from iOMS
def getESWsAndBuild(driver, po, ic):
    global result
    driver.get("http://ioms/MFG/ModuleStatus?PO=" + po + "#!/ESWs")
    timeout = False
    while(not timeout):
        try:
            text = driver.find_element_by_xpath('//*[@id="ESWs"]/div[2]').text
            if text == "ESW Data is unavailable":
                timeout = True
                break
        except:
            time.sleep(1)
            pass
        try:
            text = driver.find_element_by_xpath('//*[@id="ESWs"]/div[1]/div/div/div[2]/div/table/tbody/tr')
            break
        except:
            time.sleep(1)
            pass
    buildPercent = driver.find_element_by_xpath('//*[@id="Container"]/div[1]/div[3]/div[1]/div/div[3]/div/div/div[1]/span[2]').text
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
        result = result + buildPercent + "\n"
    else:
        result = result + "No ESW data found\n" + buildPercent + "\n"


#Get InspLots and Build % from iOMS
def getInspLotsAndBuild(driver, po, ic):
    global result
    driver.get("http://ioms/MFG/ModuleStatus?PO=" + po + "#!/inspection")
    timeout = False
    while(not timeout):
        try:
            text = driver.find_element_by_xpath('//*[@id="inspection"]/div[2]').text
            if text == "No data available":
                timeout = True
                break
        except:
            time.sleep(1)
            pass
        try:
            text = driver.find_element_by_xpath('//*[@id="inspection"]/div[1]/div/div/div[2]/div/table/tbody/tr')
            break
        except:
            time.sleep(1)
            pass
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


def runUpdater(entry, qns, text, qnSelect, ilSelect, eswSelect, allSelect, includeClosed, rctdSelect):
    global result
    result = ""
    text.configure(state='normal')
    text.delete(1.0, tk.END)
    r = getResult(entry, qns, qnSelect, ilSelect, eswSelect, allSelect, includeClosed, rctdSelect)
    text.insert(tk.END, r)
    text.configure(state='disabled')
    #print("All done")


def getResult(entry, qns, qnSelect, ilSelect, eswSelect, allSelect, includeClosed, rctdSelect):
    global result
    entryStr = entry.get()
    entry.delete('0', 'end')
    if (qnSelect or ilSelect or eswSelect or allSelect or includeClosed) and rctdSelect:
        return "RCTD/SHIP date update must be done by itself and requires Tool ID not PO."
    options = Options()
    options.headless = True # set to False to see chrome window while running
    options.add_argument("--window-size=1920,1200")
    DRIVER_PATH = r"./driver/chromedriver.exe"
    driver = webdriver.Chrome(options=options, executable_path=resource_path(DRIVER_PATH))
    result = result + ("Time started: " + datetime.now().strftime("%m/%d/%Y %H:%M:%S") + "\n")
    if rctdSelect:
        #Systems = entryToSystem(entryStr)
        print("SHIP" + "\t\t" + "RCTD" + "\n")
        result = result + ("SHIP" + "\t" + "   RCTD" + "\n\n")
        #Systems = ["129660","122279","141509"]
        Systems = entryToPo(entryStr)
        for s in Systems:
            if len(s) != 6:
                result = result + "Tool ID " + s + " not recognized, please verify this is a valid Tool ID.\n"
                continue
            getRCTD(driver, s)
        result = result + ("\nTime finished: " + datetime.now().strftime("%m/%d/%Y %H:%M:%S") + "\n")
        driver.quit()
        return result
    POs = entryToPo(entryStr)
    if allSelect:
        qnSelect = True
        eswSelect = True
    for po in POs:
        if len(po) != 7:
            result = result + "PO " + po + " not recognized, please verify this is a valid PO.\n"
            continue
        result = result + "\n################################################################################################\n"                                                                                             #\n"
        result = result + po + "\n"
        if qnSelect:
            getQNsAndInspLotsJSON(driver, po, includeClosed)
        if eswSelect:
            getESWsAndBuild(driver, po, includeClosed)
    result = result + ("\nTime finished: " + datetime.now().strftime("%m/%d/%Y %H:%M:%S") + "\n")
    driver.quit()
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
POLabel = tk.Label(root, text = "Paste input here:").grid(row = 0, column = 0, sticky = "e")
POEntry = tk.Entry(root)
POEntry.grid(row = 0, column = 1, sticky = "nsew")
text = tk.Text(root, width = 160, height = 55)
text.grid(row = 1, column = 0, columnspan = 8, sticky = "nsew")
qnSelect = tk.IntVar()
tk.Checkbutton(root, text="Include QNs and Insp Lots", variable=qnSelect).grid(row = 0, column = 3, sticky = "nsew")
ilSelect = tk.IntVar()
#tk.Checkbutton(root, text="Include Insp Lots", variable=ilSelect).grid(row = 0, column = 4, sticky = "nsew")
eswSelect = tk.IntVar()
tk.Checkbutton(root, text="Include ESWs and Build %", variable=eswSelect).grid(row = 0, column = 4, sticky = "nsew")
allSelect = tk.IntVar()
tk.Checkbutton(root, text="Include all data", variable=allSelect).grid(row = 0, column = 5, sticky = "nsew")
includeClosed = tk.IntVar()
tk.Checkbutton(root, text="Include closed", variable=includeClosed).grid(row = 0, column = 6, sticky = "nsew")
rctdSelect = tk.IntVar()
tk.Checkbutton(root, text="Get RCTD/SHIP dates", variable=rctdSelect).grid(row = 0, column = 7, sticky = "nsew")
QNCheckButton = tk.Button(root, text = "Get Update", command = (lambda: runUpdater(POEntry, qns, text, qnSelect.get(), ilSelect.get(), eswSelect.get(), allSelect.get(), includeClosed.get(),\
                                                                                   rctdSelect.get()))).grid(row = 0, column = 2, sticky = "nsew")
##rctdChekButton = tk.Button(root, text = "Check RCTD", command = (lambda: runRCTD(POEntry)))
creditLabel = tk.Label(root, text = "IC  v1.8").grid(row = 2, column = 7 , sticky = "se")
root.mainloop()
