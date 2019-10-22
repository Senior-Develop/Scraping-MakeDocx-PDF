# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import os
import csv
import time
import datetime
import pdfkit
import pyodbc
import pypandoc


output = "LetterInfo.csv"

try:
    os.remove(output)
except OSError:
    pass

line = ["FNAME", "LNAME", "ADDRESS", "ADDRESS2", "CITY", "STATE", "ZIP"]
with open(output, 'w', newline='') as file1:
    writer = csv.writer(file1, delimiter=',')
    writer.writerow(line)


def get_driver():

    options = Options()
    options.add_experimental_option("excludeSwitches",
                                    ["ignore-certificate-errors", "safebrowsing-disable-download-protection",
                                     "safebrowsing-disable-auto-update", "disable-client-side-phishing-detection"])

    options.add_argument('--disable-infobars')
    options.add_argument('--disable-extensions')
    options.add_argument('--profile-directory=Default')
    options.add_argument("--incognito")
    options.add_argument("--disable-plugins-discovery")
    prefs = {'profile.default_content_setting_values.automatic_downloads': 1}
    options.add_experimental_option("prefs", prefs)
    #options.add_argument("--headless")
    driver = webdriver.Chrome('chromedriver', options=options)
    return driver





def main():

    BOOK = "NO.txt"
    if os.path.isfile(BOOK):
        with open(BOOK, 'r') as filehandle:
            BOOK_NO = filehandle.readline()
            if not BOOK_NO:
                BOOK_NO = input("Please  Input BOOKING NO :")
    else:
        BOOK_NO = input("Please  Input BOOKING NO :")
    driver = get_driver()
    try:
        print(datetime.datetime.now())

        url = "https://apps.co.lubbock.tx.us/jailrosters/activejail.aspx"
        driver.get(url)

        table = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, 'gridaj')))
        BOOKING = table.find_element_by_tag_name("td")
        BOOKING.click()
        table = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, 'gridaj')))
        BOOKING = table.find_element_by_tag_name("td")
        BOOKING.click()
        Next_check = 0
        BOOK_check = True
        MAX_check = 0

        while BOOK_check == True:
            table = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, 'gridaj')))
            BOOK_ARR = []
            trs = table.find_elements_by_tag_name("tr")
            for i, tr in enumerate(trs):
                if i != 0:
                    BOOkN = tr.find_element_by_tag_name("td")
                    BOOK_ARR.append(BOOkN.text)
            if MAX_check == 0:
                BOOK_MAX = BOOK_ARR[0]
                MAX_check += 1
            sheets = table.find_elements_by_tag_name("input")
            sheet_len = len(sheets)
            idx = 0
            while idx < sheet_len:
                person = []
                table = WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, 'gridaj')))
                if int(BOOK_ARR[idx]) > int(BOOK_NO):
                    sheets = table.find_elements_by_tag_name("input")
                    main_window_handle = None
                    while not main_window_handle:
                        main_window_handle = driver.current_window_handle
                    sheets[idx].click()
                    time.sleep(1)
                    signin_window_handle = None
                    while not signin_window_handle:
                        for handle in driver.window_handles:
                            if handle != main_window_handle:
                                signin_window_handle = handle
                                break
                    driver.switch_to.window(signin_window_handle)
                    time.sleep(1)
                    driver.switch_to.frame(driver.find_element_by_tag_name("frame"))
                    time.sleep(1)
                    tables = driver.find_elements_by_tag_name("table")
                    if len(tables) != 10:
                        addr = driver.find_element_by_id("addr")
                        address = addr.text
                        if ("HOMELESS" not in address) and (address != "") :
                            person.append(BOOK_ARR[idx])
                            Name_Label = driver.find_element_by_id("Label1")
                            Name = Name_Label.text
                            LName = Name[:Name.index(",")]
                            Names = Name[Name.index(",") + 2 :]
                            Name_a = Names.split(" ")
                            FName = Name_a[0]
                            person.append(FName)
                            person.append(LName)

                            if "#" in address:
                                person.append(address[:address.index("#")])
                                person.append(address[address.index("#") :])
                            elif "APT" in address:
                                person.append(address[:address.index("APT")])
                                person.append(address[address.index("APT"):])
                            elif "SUITE" in address:
                                person.append(address[:address.index("SUITE")])
                                person.append(address[address.index("SUITE"):])
                            else:
                                person.append(address)
                                person.append(" ")

                            Citys = driver.find_element_by_id("citystzip")
                            City_Zip = Citys.text
                            citys_arr = City_Zip.split(" ")
                            for idk , city in enumerate(citys_arr):
                                if idk == 0:
                                    person.append(city[:-1])
                                else:
                                    person.append(city)
                            FIRST_NAME = FName.title()
                            LAST_NAME = LName.title()
                            person.append(FIRST_NAME)
                            person.append(LAST_NAME)
                            person.append(BOOK_ARR[idx] + ".pdf")

                            dir_path = os.getcwd()
                            mdb_path = dir_path + "\\person.mdb"
                            connect_db = 'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ= ' + mdb_path + ';'
                            conn = pyodbc.connect(connect_db)
                            cursor = conn.cursor()
                            sql = "INSERT INTO [person] ([BOOK_NO],[Fname],[Lname],[BASEADDRESS],[APT],[BASECITY],[BASESTATE],[BASEZIPCODE],[FIRST_NAME],[LAST_NAME],[PDF_NAME]) VALUES (?,?,?,?,?,?,?,?,?,?,?)"
                            cursor.execute(sql,person)
                            conn.commit()
                            directory = dir_path + "\\" + BOOK_ARR[idx]
                            if not os.path.exists(directory):
                                os.makedirs(directory)
                            page_html = driver.page_source
                            page_result = page_html.replace("../lsoimages","http://apps.co.lubbock.tx.us/lsoimages")
                            export_pdf = directory + "\\" + BOOK_ARR[idx] + ".pdf"
                            pdfkit.from_string(page_result, export_pdf)

                            with open(output, 'a', newline='') as file1:
                                writer = csv.writer(file1, delimiter=',')
                                writer.writerow(person)
                    else:
                        BOOK_MAX = BOOK_ARR[idx]
                    time.sleep(1)

                    driver.switch_to.default_content()
                    time.sleep(1)
                    driver.switch_to.window(main_window_handle)
                else:
                    BOOK_check = False
                    break
                idx += 1
            if Next_check == 0:
                Next = driver.find_element_by_xpath("//*[@id='gridaj']/tbody/tr[12]/td/a")
                Next_check += 1
            else:
                Next = driver.find_element_by_xpath("//*[@id='gridaj']/tbody/tr[12]/td/a[2]")
            Next.click()
            if int(BOOK_ARR[0]) >= int(BOOK_MAX):
                with open(BOOK, 'w') as the_file:
                    BOOK_MAX = BOOK_ARR[0]
                    the_file.write(BOOK_MAX)

    except:
        print("loading page error")
        pass
    driver.quit()
    print("---------processing end-----------")
    print(datetime.datetime.now())
    time.sleep(60 * 30)

if __name__ == "__main__":
    while True:
        main()
