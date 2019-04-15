#This script pulls the most current Real Time Settlement Point Price CSV file
#from the ERCOT Mass Information System every 15 minutes

import webbrowser
import time
import bs4
import requests

while True: #Loops forever until manually interrupted
    try:
        #Capturing the reports page in a variable
        res = requests.get('http://mis.ercot.com/misapp/GetReports.do?reportTypeId=12301&reportTitle=Settlement%20Point%20Prices%20at%20Resource%20Nodes,%20Hubs%20and%20Load%20Zones&showHTMLView=&mimicKey')
        SPPSoup = bs4.BeautifulSoup(res.text, "lxml")   #Parsing the HTML from reports page
        List_of_DocumentIDs = SPPSoup.select('a')   #Pulls the URL links for active reports from the HTML
        Top_Document = List_of_DocumentIDs[0].get('href')   #Grabs and labels the top active report link
    except:
        print("ERCOT MIS Site Failure " + time.ctime())
        time.sleep(300)

    try:
        Last_Save = open("lastSPP.txt", "r") #Looks for lastSPP.txt
        MRD = Last_Save.read() #If available, reads lastSPP.txt and commits to MRD variable
        Last_Save.close()
    except: #If lastSPP.txt does not exist, run the following.
        print("No previous save file exists! Creating new file.")
        Last_Save = open("lastSPP.txt", "w") #Creates lastSPP.txt
        Last_Save.write(Top_Document)   #Adds the Top_Document to the lastSPP file
        Last_Save.close()
        MRD = Top_Document  #Sets MRD variable to the Top_Document string
        pass

    #Find the most recent downloaded SPP
    while (int(Top_Document[58:]) != int(MRD[58:])):  #Check if Top_Document is same as MRD variable
        Check = 0 #Set a check value to 0
        TDL = List_of_DocumentIDs[Check].get('href')    #Creates TDL (Temporary Download) variable
        while (int(TDL[58:]) != int(MRD[58:])):   #Loop until TDL varable finds MRD variable
            Check = Check + 2   #Change check value by 2 (csv files are separated by 2 on page)
            TDL = List_of_DocumentIDs[Check].get('href')    #Resets TDL variable to next URL
        Check = Check - 2   #Returns check value to one instance after MRD variable
        TDL = List_of_DocumentIDs[Check].get('href')    #Sets TDL variable to one instance after MRD variable
        MRD = TDL   #Temporary DL variable becomes Most Recent Download
        print(TDL[58:] + " " + time.ctime())
        webbrowser.open('http://mis.ercot.com' + (str(TDL)))    #Downloads next SPP
        Last_Save = open("lastSPP.txt", "w")
        Last_Save.write(MRD)    #Update lastSPP with new MRD variable
        Last_Save.close()
        time.sleep(1)  #Pause 1 second to ensure download starts

    time.sleep(900)    #Wait 900 seconds to restart loop
