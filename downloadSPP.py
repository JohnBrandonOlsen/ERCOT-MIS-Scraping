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

def generate_documents_list(site_address):
    try:
        res = requests.get(site_address)
        ercot_soup = BeautifulSoup(res.text, "lxml")
        list_of_documents = ercot_soup.findAll('td', attrs={'class': 'labelOptional_ind'})
        list_of_links = ercot_soup.select('a')

        return list_of_documents, list_of_links

    except:
        print("\nERCOT MIS Site Failure at: " + time.ctime())
        time.sleep(300)

def initial_most_recent_download(top_document):
    try:
        last_save = open("testSPP.txt", "r")
        most_recent_download = last_save.read()
        last_save.close()
    except:
        print("No previous save file exists! Creating new file.")
        last_save = open("testSPP.txt", "w")
        last_save.write(top_document)
        last_save.close()
        most_recent_download = top_document
        pass

    return most_recent_download

def find_document_iterable(most_recent_download, list_of_documents):
    document_iterable = 0
    temp_download = list_of_documents[document_iterable].text[62:70] + list_of_documents[document_iterable].text[71:75]

    while int(temp_download) != int(most_recent_download):
        document_iterable = document_iterable + 2   #Because the page listsl csv and xml documents, needs to cycle two at a time
        temp_download = list_of_documents[document_iterable].text[62:70] + list_of_documents[document_iterable].text[71:75]
    if document_iterable != 0:
        document_iterable -= 2
    else:
        pass

    return document_iterable

def download_csv_zip(most_recent_download, list_of_documents, list_of_links):
    current_document_position = find_document_iterable(most_recent_download, list_of_documents)
    temp_download = list_of_links[current_document_position].get('href')    #Sets temp_download variable to one instance after most_recent_download variable
    print("Downloading " + list_of_documents[current_document_position].text[64:] + " at: " + time.ctime(), end='\r')
    webbrowser.open('http://mis.ercot.com' + str(temp_download))    #Downloads next SPP

    most_recent_download = list_of_documents[current_document_position].text[62:70] + list_of_documents[current_document_position].text[71:75]
    update_last_save(most_recent_download)

    return(most_recent_download, current_document_position)


def update_last_save(most_recent_download):
        last_save = open("testSPP.txt", "w")
        last_save.write(most_recent_download)    #Update testSPP with new most_recent_download variable
        last_save.close()

def verify_download(zip_filename):
    while True:
        try:
            with zf(downloads_folder + zip_filename) as current_zip:
                current_zip.close()
                break
        except:
            time.sleep(1)
            continue

def create_new_daily_xlsx(zip_filename, csv_filename, csv_compendium_filename):
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

def add_price_data(zip_filename, csv_filename, price_data_loc):
    with zf(downloads_folder + zip_filename) as current_zip:
        with current_zip.open(csv_filename, 'r') as f:
            readCSV = csv.reader(TextIOWrapper(f), delimiter=',')
            price_data = []
            next(readCSV)
            for row in readCSV:
                price_data.append(row[5])
            excel_file.sheets[0].range(price_data_loc).value = price_data
        excel_file.save()
    remove(downloads_folder + zip_filename)

def update_csv_data(zip_filename, csv_filename):
    if zip_filename[71:75] == "0015":
        csv_compendium_filename = zip_filename[62:70]+".xlsx"
        create_new_daily_xlsx(zip_filename, csv_filename, csv_compendium_filename)
        print("\nFile " + csv_compendium_filename + " created.")
        price_data_loc = "D2"
        add_price_data(zip_filename, csv_filename, price_data_loc)

    elif zip_filename[71:75] == "0000":
        price_data_compendium_interval = "D97"
        add_price_data(zip_filename, csv_filename, price_data_compendium_interval)
        excel_file.close()

    else:
        price_data_loc = "D" + str(find_hour_interval(zip_filename) + 2)
        add_price_data(zip_filename, csv_filename, price_data_loc)

def find_hour_interval(zip):
    hour_intervals = ["0015", "0030", "0045", "0100", "0115", "0130",
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
                    "2245", "2300", "2315", "2330", "2345", "0000"
                    ]

    position = hour_intervals.index(zip[71:75])

    return position

def find_most_recent_0000(most_recent_download, list_of_documents, current_document_position, zip_filename):
    current_document_position = current_document_position + ((find_hour_interval(zip_filename) + 1) * 2)
    most_recent_download = list_of_documents[current_document_position].text[62:70] + list_of_documents[current_document_position].text[71:75]
    return most_recent_download

def main():
    while True:
        ercot_lists = generate_documents_list(ercot_site_address)
        list_of_documents = ercot_lists[0]
        list_of_links = ercot_lists[1]
        top_document = list_of_documents[0].text[62:70] + list_of_documents[0].text[71:75]
        most_recent_download = initial_most_recent_download(top_document)

        csv_zip_details = download_csv_zip(most_recent_download, list_of_documents, list_of_links)

        while int(top_document) != int(most_recent_download):
            most_recent_download = csv_zip_details[0]
            current_document_position = csv_zip_details[1]

            zip_filename = list_of_documents[current_document_position].text
            verify_download(zip_filename)
            csv_filename = zf(downloads_folder + zip_filename, 'r').namelist()[0]

            try:
                update_csv_data(zip_filename, csv_filename)

            except:
                try:
                    global excel_file
                    excel_file = xw.Book(save_folder + zip_filename[62:70] + ".xlsx")
                    update_csv_data(zip_filename, csv_filename)
                except:
                    most_recent_download = find_most_recent_0000(most_recent_download, list_of_documents, current_document_position, zip_filename)
                    remove(downloads_folder + zip_filename)
            if most_recent_download != top_document:
                csv_zip_details = download_csv_zip(most_recent_download, list_of_documents, list_of_links)
            else:
                pass

        time.sleep(900)

main()
