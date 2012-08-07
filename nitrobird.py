#twitter client for Excel; feel free to modify to your needs
import tweepy
import sys
import time
import threading

REFRESH_RATE = 120 #refresh rate in seconds

#provides functionality for twitter & Excel
class NitroBird():
    consumerKey = None
    consumerSecret = None
    accessToken = None
    accessTokenSecret = None
    tokenPin = None
    auth = None
    authUrl = None
    api = None
    prevSheet = None
    prevX = None
    prevY = None

    #reads app keys from the settings tab
    def __init__(self):
        self.consumerKey = Cell("settings", 12, 2).value
        self.consumerSecret = Cell("settings", 13, 2).value

        self.accessToken = str(Cell("settings", 8, 2).value)
        self.accessTokenSecret = str(Cell("settings", 9, 2).value)

        self.tokenPin = str(Cell("settings", 5, 2).value)

        self.auth = tweepy.OAuthHandler(self.consumerKey, self.consumerSecret)

    #obtains authentication URL
    def getAuthUrl(self):
        ret = False
        try:
            self.auth_url = self.auth.get_authorization_url()
            Cell("settings", 2, 2).value = self.auth_url
            Cell("settings", 1, 2).value = "Authorization URL issued"
            ret = True
        except:
            Cell("settings", 1, 2).value = "Sorry, failed to issue the URL. Make sure you're connected to the Internet and run the script again."
            ret = False
        return ret

    #keeps waiting until a valid PIN number is entered (or access keys) on the settings tab
    def doAuth(self):
        if len(str(self.accessToken)) < 5 or len(str(self.accessTokenSecret)) < 5:
            self.getAuthUrl()
            done = False
            while not done:
                while self.tokenPin == None or len(self.tokenPin) < 5:
                    self.tokenPin = str(Cell("settings", 5, 2).value)
                    time.sleep(0.1)
            #got the pin, try to get access token
                try:
                    self.auth.get_access_token(self.tokenPin)
                    self.accessToken = Cell("settings", 8, 2).value = str(self.auth.access_token.key)
                    self.accessTokenSecret = Cell("settings", 9, 2).value = str(self.auth.access_token.secret)
                    done = True
                except:
                    self.tokenPin = None
                    done = False
                    Cell("settings", 1, 2).value = "Incorrect PIN number, please try again"
                    Cell("settings", 5, 2).value = ""
                    time.sleep(0.1)
        else:
            self.auth.set_access_token(self.accessToken, self.accessTokenSecret)

        #auth success
        Cell("settings", 1, 2).value = "Everything's OK, enjoy your Twitter in Excel!"
        Cell("settings", 1, 2).font.color = Cell.COLORS.GREEN
        self.api = tweepy.API(self.auth)

    def updateEverything(self):
        self.updateProfile(self.api.me(), "profile")
        self.updateHome()

    def clearColumn(self, sheet, startY, x):
        i = 0
        while Cell(sheet, startY + i, x).value != "" and Cell(sheet, startY + i, x).value != None:
            Cell(sheet, startY + i, x).value = ""
            i += 1

    def clearUser(self):
        return self.clearProfile("user")

    def clearProfile(self, page = "profile"):
        print "clearing " + str(page)
        Cell(page, 2, 2).value = Cell(page, 3, 2).value = ""
        Cell(page, 2, 3).value = Cell(page, 3,3).value = ""
        Cell(page, 4, 2).value = Cell(page, 5, 2).value = ""
        Cell(page, 5, 4).value = ""
        Cell(page, 4, 4).value = ""
        #clear friends, followers, tweets
        self.clearColumn(page, 7, 1)
        self.clearColumn(page, 7, 2)
        self.clearColumn(page, 7, 3)

    #updates the profile sheet
    def updateProfile(self, u, page = "profile"):
        try:
            Cell(page, 2, 2).value = u.screen_name
            Cell(page, 3, 2).value = u.name
            Cell(page, 2, 3).value = u.description
            Cell(page, 3, 3).value = u.url
            Cell(page, 4, 2).value = u.followers_count
            Cell(page, 5, 2).value = u.friends_count

            i = 0
            for f in u.friends():
                Cell(page, 7 + i, 1).value = f.screen_name
                i += 1
            i = 0
            for f in u.followers():
                Cell(page, 7 + i, 2).value = f.screen_name
                i += 1
            i = 0
            for p in range(5):
                for t in self.api.user_timeline(u.screen_name, page = p):
                    Cell(page, 7 + i, 3).value = str(t.text.encode('ascii', 'ignore')).replace("\n", " ")            
                    i += 1
        except:
            Cell(page, 7, 3).value = "Sorry, you're not authorized to see this!"

    def clearHome(self):
        Cell("home", 1, 1).value = ""
        self.clearColumn("home", 3, 1)
        self.clearColumn("home", 3, 2)

    def postTweet(self):
        msg = Cell("profile", 4, 4).value
        if msg == None or msg == "" or len(msg) < 1:
            return
        try:
            self.api.update_status(msg)
            Cell("profile", 5, 4).value = "Tweeted!"
        except:
            Cell("profile", 5, 4).value = "Sorry, failed to tweet!"
        

    #updates home, which contain all your & friends' tweets ("timeline")
    def updateHome(self):
        self.clearHome()
        Cell("home", 1, 1).value = self.api.me().screen_name
        Cell("home", 2, 1).value = Cell("home", 2, 2).value = ""
        i = 3
        for p in range(5):
            for t in self.api.home_timeline(page = p):
                Cell("home", i, 1).value = t.author.screen_name
                Cell("home", i, 2).value = str(t.text.encode('ascii', 'ignore')).replace("\n", " ")
                i += 1
        
    #updates another user's profile
    def updateUser(self, screen_name):
        self.clearProfile("user")
        print "updating " + str(screen_name)
        u = None
        try:
            u = self.api.get_user(screen_name)
        except: 
            Cell("user", 2, 2).value = "None"
            return
        self.updateProfile(u, "user")
        iron.setActiveCell(1,1)

    def clearSearch(self):
        self.clearColumn("search", 3, 2)

    def search(self):
        self.clearSearch()
        query = Cell("search", 1, 2).value
        print "query = " + str(query)
        results = self.api.search(query)
        i = 3
        for t in results:
            Cell("search", i, 1).value = t.from_user
            Cell("search", i, 2).value = str(t.text.encode('ascii', 'ignore'))
            i += 1
            
    #waits for clicks on certain sheets/cells
    def processEvent(self):
        ret = False
        sheet = iron.getActiveWorksheet()
        (y, x) = iron.getActiveCell()
        if sheet != None and y >= 1 and x >= 1:
            if (sheet, y, x) != (self.prevSheet, self.prevY, self.prevX):                    
                if ((sheet == "profile" or sheet == "user") and \
                        (x <= 2 and y >= 7)) or \
                        ((sheet == "home" or sheet == "search") and x == 1 and y >= 3):
                    s = Cell(sheet, y, x).value
                    if len(str(s)) > 0:
                        print "updating"
                        iron.setActiveCell(1,1) #previous sheet
                        iron.setActiveWorksheet("user")
                        iron.setActiveCell(1,1)
                        self.updateUser(s)
                        ret = True
                        (self.prevSheet, self.prevY, self.prevX) = (sheet, y, x)                        
                #do search
                if (sheet == "search" and y == 1 and x == 3): 
                    print "in search"
                    iron.setActiveCell(1,1)
                    self.search()
                    (self.prevSheet, self.prevY, self.prevX) = (sheet, y, x)
                    ret = True
                if sheet == "profile" and y == 5 and x == 3:
                    iron.setActiveCell(1,1)
                    self.postTweet()
                    iron.setActiveCell(1,1)
                    self.updateHome()
                    self.updateProfile(self.api.me(), "profile")
                    iron.setActiveCell(1,1)                    
                    (self.prevSheet, self.prevY, self.prevX) = (sheet, y, x)
                if sheet == "home" and y == 1 and x == 11:
                    self.updateHome()
                    (self.prevSheet, self.prevY, self.prevX) = (sheet, y, x)
        
        return ret

#script execution starts here
Cell("settings", 1, 2).font.color = Cell.COLORS.RED
client = NitroBird()
print "nitrobird"
if client.doAuth() == False:
    sys.exit(1)
client.clearProfile()
client.clearHome()
client.clearSearch()
client.clearUser()
client.updateEverything()

#event loop
cnt = 0
while True:
    r = client.processEvent()
    cnt += 1
    if cnt > REFRESH_RATE * 10: #every roughly 2mins update
        cnt -= REFRESH_RATE * 10
        client.updateHome()
        print "auto refresh done"
    if r == False:
        time.sleep(0.1)


