#This script pulls the most current Real Time Settlement Point Price CSV file
#from the ERCOT Mass Information System every 15 minutes

import webbrowser
import time
import requests
import csv
from zipfile import ZipFile as zf
from io import TextIOWrapper
from os import remove

from bs4 import BeautifulSoup
import xlwings as xw

downloads_folder = "path/to/downloads_folder"
save_folder = "path/to/save_folder"
ercot_site_address = "http://mis.ercot.com/misapp/GetReports.do?reportTypeId=12301&reportTitle=Settlement%20Point%20Prices%20at%20Resource%20Nodes,%20Hubs%20and%20Load%20Zones&showHTMLView=&mimicKey"

def create_new_daily_xlsx():
    global excel_file

    with zf(downloads_folder + zip_filename) as current_zip:
        with current_zip.open(csv_filename, 'r') as f:
            readCSV = csv.reader(TextIOWrapper(f), delimiter=',')
            delivery_date = []
            settlement_points = []
            next(readCSV)
            for row in readCSV:
                delivery_date.append(row[0])
                settlement_points.append(row[3])

    excel_file = xw.Book()
    excel_file.sheets[0].range("D1").value = settlement_points
    excel_file.sheets[0].range("A2:A97").value = delivery_date[3]

    temp_counter = 2
    hour = 1
    for i in range(24):
        minute_interval_counter = temp_counter
        minute_interval = 1
        for minute in range(4):
            excel_file.sheets[0].range("C" + str(minute_interval_counter)).value = str(minute_interval)
            minute_interval_counter += 1
            minute_interval += 1
        excel_file.sheets[0].range("B" + str(temp_counter) + ":B" + str(temp_counter + 3)).value = str(hour)
        temp_counter += 4
        hour += 1

    excel_file.save(save_folder + csv_compendium_filename)

def add_price_data():
    global price_data_loc
    global price_data_compendium_interval

    with zf(downloads_folder + zip_filename) as current_zip:
        with current_zip.open(csv_filename, 'r') as f:
            readCSV = csv.reader(TextIOWrapper(f), delimiter=',')
            price_data = []
            next(readCSV)
            for row in readCSV:
                price_data.append(row[5])
            excel_file.sheets[0].range(price_data_loc).value = price_data
            price_data_compendium_interval += 1
            price_data_loc = "D" + str(price_data_compendium_interval)
        excel_file.save()
    remove(downloads_folder + zip_filename)

def update_csv_data():
    global zip_filename
    global csv_filename
    global csv_compendium_filename
    global price_data_loc
    global price_data_compendium_interval
    global excel_file

    list_of_csvs = ercot_soup.findAll('td', attrs={'class': 'labelOptional_ind'})
    zip_filename = list_of_csvs[document_iterable].text

    verify_download()

    csv_filename = zf(downloads_folder + zip_filename, 'r').namelist()[0]

    if zip_filename[71:75] == "0015":
        csv_compendium_filename = zip_filename[62:70]+".xlsx"
        create_new_daily_xlsx()
        print("File " + csv_compendium_filename + " created.")

        price_data_compendium_interval = 2
        price_data_loc = "D" + str(price_data_compendium_interval)
        add_price_data()

    elif zip_filename[71:75] == "0000":
        add_price_data()
        excel_file.close()

    else:
        try:
            add_price_data()
        except:
            try:
                csv_compendium_filename = zip_filename[62:70]+".xlsx"
                excel_file = xw.Book(save_folder + csv_compendium_filename)

                price_data_compendium_interval = find_hour_interval() + 1
                price_data_loc = "D" + str(price_data_compendium_interval)
                add_price_data()
            except Exception as e:
                find_most_recent_0000()

def verify_download():
    while True:
        try:
            with zf(downloads_folder + zip_filename) as current_zip:
                break
        except:
            time.sleep(1)
            continue

def find_hour_interval():
    hour_intervals = ["0000", "0015", "0030", "0045", "0100", "0115", "0130",
                    "0145", "0200", "0215", "0230", "0245", "0300", "0315",
                    "0330", "0345", "0400", "0415", "0430", "0445", "0500",
                    "0515", "0530", "0545", "0600", "0615", "0630", "0645",
                    "0700", "0715", "0730", "0745", "0800", "0815", "0830",
                    "0845", "0900", "0915", "0930", "0945", "1000", "1015",
                    "1030", "1045", "1100", "1115", "1130", "1145", "1200",
                    "1215", "1230", "1245", "1300", "1315", "1330", "1345",
                    "1400", "1415", "1430", "1445", "1500", "1515", "1530",
                    "1545", "1600", "1615", "1630", "1645", "1700", "1715",
                    "1730", "1745", "1800", "1815", "1830", "1845", "1900",
                    "1915", "1930", "1945", "2000", "2015", "2030", "2045",
                    "2100", "2115", "2130", "2145", "2200", "2215", "2230",
                    "2245", "2300", "2315", "2330", "2345"
                    ]

    position = hour_intervals.index(zip_filename[71:75])

    return position

def find_most_recent_0000():
    global most_recent_download

    most_recent_download = list_of_documents[document_iterable + (find_hour_interval() * 2)].get('href')

while True: #Loops forever until manually interrupted
    try:
        #Capturing the reports page in a variable
        res = requests.get(ercot_site_address)
        ercot_soup = BeautifulSoup(res.text, "lxml")   #Parsing the HTML from reports page
        list_of_documents = ercot_soup.select('a')   #Pulls the URL links for active reports from the HTML
        top_document = list_of_documents[0].get('href')   #Grabs and labels the top active report link
    except Exception as e:
        print("ERCOT MIS Site Failure " + time.ctime())
        time.sleep(300)

    try:
        last_save = open("testSPP.txt", "r") #Looks for testSPP.txt
        most_recent_download = last_save.read() #If available, reads testSPP.txt and commits to most_recent_download variable
        last_save.close()
    except: #If testSPP.txt does not exist, run the following.
        print("No previous save file exists! Creating new file.")
        last_save = open("testSPP.txt", "w") #Creates testSPP.txt
        last_save.write(top_document)   #Adds the top_Document to the testSPP file
        last_save.close()
        most_recent_download = top_document  #Sets most_recent_download variable to the top_Document string
        pass

    #Find the most recent downloaded SPP
    while (int(top_document[58:]) != int(most_recent_download[58:])):  #document_iterable if top_Document is same as most_recent_download variable
        document_iterable = 0 #Set a document_iterable value to 0
        temp_download = list_of_documents[document_iterable].get('href')    #Creates temp_download variable

        while (int(temp_download[58:]) != int(most_recent_download[58:])):   #Loop until temp_download varable finds most_recent_download variable
            document_iterable = document_iterable + 2   #Change document_iterable value by 2 (csv files are separated by 2 on page)
            temp_download = list_of_documents[document_iterable].get('href')    #Resets temp_download variable to next URL

        document_iterable = document_iterable - 2   #Returns document_iterable value to one instance after most_recent_download variable
        temp_download = list_of_documents[document_iterable].get('href')    #Sets temp_download variable to one instance after most_recent_download variable

        most_recent_download = temp_download   #Temporary DL variable becomes Most Recent Download
        print(temp_download[58:] + " " + time.ctime(), end='\r')
        webbrowser.open('http://mis.ercot.com' + (str(temp_download)))    #Downloads next SPP

        update_csv_data()

        last_save = open("testSPP.txt", "w")
        last_save.write(most_recent_download)    #Update testSPP with new most_recent_download variable
        last_save.close()

    time.sleep(900)    #Wait 900 seconds to restart loop
