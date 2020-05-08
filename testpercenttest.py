####################################################################################################################
############## BIG BUG NEED TO CHECK DATES ON LAST TEST RAN NOT ON LAST LINE########################################
#####################################################################################################################

while True:
    try:
        bay_num = int(input("Enter bay: "))
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

mlog = r'\\port8' + port + r'\Temp\automation\logs\Mayan.log'
elog = r'\\port8' + port + r'\Temp\automation\logs\MayanExecutive.log'
######################################################################

configfile = r'\\port8' + port + r'\Amat\EnduraCGA\Data\Config.en'

class Chamber:
    def _init_(self, kind, position):
        self.kind = kind
        self.position = position

    def info(self):
        return self.kind + " " + self.position

    def print(self):
        print(self.kind + " " + self.position)
#####################################################################

with open(mlog, 'rb') as fh:
    firstlog = next(fh).decode()
    fh.seek(-1024, 2)
    lastlog = fh.readlines()[-1].decode()
    logdate = lastlog[0:9]
    print(logdate)

with open(elog, 'rb') as fh:
    firstelog = next(fh).decode()
    fh.seek(-1024, 2)
    lastelog = fh.readlines()[-1].decode()
    elogdate = lastelog[0:9]
    print(elogdate)

if logdate > elogdate:
    print("mlog is more recent")
    log = open(mlog, 'r')
    usingelog = False
else:
    print("elog is more recent")
    log = open(elog, 'r')
    usingelog = True

lines = log.readlines()
log.close()
reversed_lines = reversed(lines)
reversed_log = []
reversed_elog = []
#########################################################

config = open(configfile, 'r')
lines = []
with open(configfile) as config:
    head = [next(config) for x in range(14)]

C = Chamber()

for x in head[5:]:
    if "Absent" in x:
        continue
        #print("no chamber configured")
    else:
        C.kind = x[(x.find("=") + 1):x.find("	")]
        y = x.find("@_") + 2
        C.position = x[y:(y + 3)]
        C.print()

########################################################
for i in reversed_lines:
    reversed_log.append(i)

if usingelog == True:
    for i in reversed_log:
        if "Test Passed" in i:
            print("last test ran: " + i[i.find("'"):-1] + " at " + i[0:20]) #implement find ' and start there
            break

else:
    for i in reversed_log:
        if "Run Test" in i:
            print("Last test ran: '" + i[(i.find("Run Test")+10):-1] + "' at " + i[0:20])
            break

print("found last test attempted")
