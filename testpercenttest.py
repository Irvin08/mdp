#import time

while True:
    try:
        bay = input("Enter bay: ")
        bay_num = int(bay)
    except ValueError:
        print("That's not a number dude")
    else:
        if 1 <= bay_num <= 10:
            break
        else:
            print("Only bays 1-10 are available.")
while True:
        port = input("Enter port letter: ")
        if len(port) > 1:
            print("Only ports A-F are available.")
        elif "a"  <= port <= "f" or "A" <= port <= "F":
            break
        else:
            print("Only ports A-F are available.")

mlog = r'\\port' + bay + port + r'\Temp\automation\logs\Mayan.log'
elog = r'\\port' + bay + port + r'\Temp\automation\logs\MayanExecutive.log'

#May need to modify this for centura chambers
configfile = r'\\port'+ bay + port + r'\Amat\EnduraCGA\Data\Config.en'

class Chamber:
    def _init_(self, kind, position):
        self.kind = kind
        self.position = position

    def info(self):
        return self.kind + " " + self.position

    def print(self):
        print(self.kind + " " + self.position)

#Read chamber description lines to find current chamber
config = open(configfile, 'r')
lines = []
with open(configfile) as config:
    head = [next(config) for x in range(14)]

C = Chamber()

for x in head[5:]:
    if "Absent" in x:
        continue
    else:
        C.kind = x[(x.find("=") + 1):x.find("	")]
        y = x.find("@_") + 2
        C.position = x[y:(y + 3)]
        C.print()
lines.clear()

#Stopwatch to see how long the reversing of logs takes
#starttime = time.time()

#Read logs and reverse them since latest test data is at end
log = open(mlog, 'r')
lines = log.readlines()
log.close()
reversed_lines = reversed(lines)
reversed_log = []
for i in range(0,200):#reversed_lines:
    reversed_log.append(next(reversed_lines))#reversed_log.append(i)

xlog = open(elog, 'r')
xlines = xlog.readlines()
xlog.close()
reversed_xlines = reversed(xlines)
reversed_xlog = []
for i in range(0,200):# reversed_xlines:
    reversed_xlog.append(next(reversed_xlines))#reversed_xlog.append(i)

#print("--- %s seconds ---" % (time.time() - starttime))

#Find latest test run/passed on both logs and print out most recent
for i in reversed_log:
    if "Run Test" in i:
        last_test = "Last test ran: '" + i[(i.find("Run Test")+10):-1] + "' at " + i[0:20]
        logdate = i[0:9]
        print(logdate)
        break

for i in reversed_xlog:
        if "Test Passed" in i:
            last_xtest = "Last test ran: " + i[i.find("'"):-1] + " at " + i[0:20]
            elogdate = i[0:9]
            print(elogdate)
            break

if logdate > elogdate:
    print(last_test)
else:
    print(last_xtest)


