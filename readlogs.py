#Python code used to generate access.log file statistics 
#in Excel with IronSpread

import datetime

#base class to keep the # of unique IPs in a time interval
#with size of interval = timedelta
class Stat:
    stats = []
    sumCount = 0
    ycoord = 2
    timedelta = None
    curDate = None
    totCount = 0
    
    #adds a unique entry to the current time interval
    def addEntry(self, ip, date):
        duplicate = False
        self.totCount += 1
        for el in self.stats:
            if el[0] == ip:
                duplicate = True

        if not duplicate:
            self.stats += [(ip, date)]

    #prints statistics for all intervals in a date range
    #between self.curDate and dateLimit (exclusive)
    #and sets curDay to dayLimit
    def printStatsPeriod(self, dateLimit):
        while not self.sameDate(dateLimit):
            self.printCurStats()
            self.curDate += self.timedelta
            self.totCount = 0
            self.stats = []

#derived class for hourly statistics
#define timedelta = 1 hour and defines specific Excel output formattin
class HourStat(Stat):
    def __init__(self, initDate):
        self.curDate = initDate
        self.timedelta = datetime.timedelta(hours=1)
    
    def printCurStats(self):
        avg = 0.0
        if len(self.stats) > 0:
            avg = self.totCount/float(len(self.stats))
        #defined by IronSpread when run in Excel
        Cell("Hourly", self.ycoord, 1).value = self.curDate.strftime("%Y/%b/%d")
        Cell("Hourly", self.ycoord, 2).value = self.curDate.strftime("%H")
        Cell("Hourly", self.ycoord, 3).value = len(self.stats)
        Cell("Hourly", self.ycoord, 4).value = avg
        self.ycoord += 1

    def sameDate(self, date):
        if self.curDate == None:
            return False
        return self.curDate.year == date.year and \
            self.curDate.month == date.month and \
            self.curDate.day == date.day and \
            self.curDate.hour == date.hour
        
#class for daily statistics
class DayStat(Stat):
    def __init__(self, initDate):
        self.curDate = initDate
        self.timedelta = datetime.timedelta(days=1)

    def printCurStats(self):
        avg = 0.0
        if len(self.stats) > 0:
            avg = self.totCount/float(len(self.stats))
        #defined by IronSpread when run in Excel
        Cell("Daily", self.ycoord, 1).value = self.curDate.strftime("%Y/%b/%d")
        Cell("Daily", self.ycoord, 2).value = len(self.stats)
        self.sumCount += len(self.stats)
        Cell("Daily", self.ycoord, 3).value = self.sumCount
        Cell("Daily", self.ycoord, 4).value = avg
        self.ycoord += 1


    def sameDate(self, date):
        if self.curDate == None:
            return False
        return self.curDate.year == date.year and \
            self.curDate.month == date.month and \
            self.curDate.day == date.day

#weekly statistics
class WeekStat(Stat):
    def __init__(self, initDate):
        self.curDate = initDate
        self.timedelta = datetime.timedelta(days = 7)

    def sameDate(self, date):
        if self.curDate == None:
            return False
        return self.curDate.isocalendar()[1] == date.isocalendar()[1]

    def printCurStats(self):
        avg = 0.0
        if len(self.stats) > 0:
            avg = self.totCount/float(len(self.stats))
        Cell("Weekly", self.ycoord, 1).value = self.curDate.strftime("%Y/%b/%d")
        Cell("Weekly", self.ycoord, 2).value = len(self.stats)
        self.sumCount += len(self.stats)
        Cell("Weekly", self.ycoord, 3).value = self.sumCount
        Cell("Weekly", self.ycoord, 4).value = avg
        self.ycoord += 1


#actual execution starts here
logFile = open("access.log", 'r')
if logFile == None:
    print "failed to open the log file"
    exit(1)

#list of IPs to blacklist
blacklist = [ ]

#headers for each spreadsheet/column
Cell("Hourly", 1,1).value = Cell("Daily", 1,1).value = Cell("Weekly", 1,1).value = "Date"
Cell("Hourly", 1, 3).value = Cell("Daily", 1,2).value = Cell("Weekly", 1,2).value = "#"
Cell("Hourly", 1, 4).value = Cell("Daily", 1,3).value = Cell("Weekly", 1,3).value = "sum"
Cell("Hourly", 1, 5).value = Cell("Daily", 1,4).value = Cell("Weekly", 1,4).value = "ratio"
Cell("Hourly", 1, 2).value = "H"

#make each of them bold
for x in range(1, 6):
    Cell("Hourly", 1, x).font.bold = Cell("Daily", 1, x).font.bold = Cell("Weekly", 1, x).font.bold = True


count = 0
#will be initiated to hourly, daily, weekly on the first log entry
stats = [None, None, None]
for line in logFile:
    if not "FileName" in line:
        continue
    
    words = line.split()
    logIp = words[0]
    
    if logIp in blacklist:
        continue

    logDateTime = words[3][1:-1]
    logDate = datetime.datetime.strptime(logDateTime, "%d/%b/%Y:%H:%M:%S")

    if stats[0] == None:
        stats[0] = DayStat(logDate)
    if stats[1] == None:
        stats[1] = WeekStat(logDate)
    if stats[2] == None:
        stats[2] = HourStat(logDate)


    #print current intervals and move to the latest one
    #if needed
    for istat in range(0, 3):
        if not stats[istat].sameDate(logDate):
            stats[istat].printStatsPeriod(logDate)
        stats[istat].addEntry(logIp, logDate)

#print last final day
for istat in range(0, 3):
    stats[istat].printCurStats()

#automatically adjust width of all columns in the entire spreadsheet
autofit()
logFile.close()    

