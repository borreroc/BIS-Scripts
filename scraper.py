import csv
import re
import sys    
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import requests
from bs4 import BeautifulSoup, Comment

def get_elements (soup):
    #find the table by traversing back up nested tags, this can probably be done better
    table = soup.find(text = 'Job Title').parent.parent.parent.parent.parent.parent.parent.parent.parent

    #get us the rows
    rows = table.find_all("td")
    my_string = ''
    row_list = []
    my_list = []

    #get the text of the rows and remove empty lines, replace \n and commas, and acknowledge null entries. take each line and add it as an element to row_list
    j = 0
    while j < len(rows):

        refined_list = re.split(r'\s{2,}', rows[j].get_text())
        if (refined_list[0] != ''):
            this = refined_list[0].replace('\n', '')
            this = this.replace(',', ' -')
            row_list.append(this.replace('\xa0', 'n/a'))
        j = j + 1

    return row_list


def group_and_print (row_list):
    my_list = []

    #add 11 elements to my_list, then write them as a row to the csv (skip headers). clear my_list and start with the next 11 elements
    k = 11
    while k < len(row_list):
        j = 0
        while j < 11:
            my_list.append(row_list[k])
            j = j + 1
            k = k + 1
        print (my_list)
        global_list.append(my_list)
        print ("\n\n\n")
        my_list = []

def recursive ():
    #get table of info
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    row_list = get_elements(soup = soup)
    group_and_print(row_list = row_list)

    try:
        if (driver.find_element_by_id("functionalityTableDelegate_getNextRowSet")):
            driver.find_element_by_id("functionalityTableDelegate_getNextRowSet").click()
            recursive()
    except NoSuchElementException:
        return 0
        

#open chrome and go to rms
driver = webdriver.Chrome()
driver.get("https://jobs.ucsc.edu/userfiles/jsp/shared/frameset/Frameset.jsp?time=1348080164562")

#let page load
driver.implicitly_wait(10)
driver.switch_to.frame("contentFrame")

rms_user = input("Enter RMS username:")
rms_pw = input ("Enter RMS password:")

#enter username and password
username = driver.find_element_by_id("userName")
username.clear()
username.send_keys(rms_user)
password = driver.find_element_by_name("password")
password.clear()
password.send_keys(rms_pw)
driver.find_element_by_name("submit").click()

#open output file and add headers
with open('output.csv','w', newline = '') as outfile:
    writer = csv.writer(outfile, dialect='excel')
    my_list = ['Job Title','Job Number', 'Rec. Type', 'Div/Org', 'CHM', 'Rec. Status', 'Status Date', 'IRD', 'Total Apps.', 'Incomplete Job Offers', 'Hires']
    writer.writerow(my_list)

global_list = []

driver.switch_to.frame("contentFrame")
recursive()
#print(global_list)

for val in global_list:
        with open('output.csv','a', newline = '') as outfile:
            writer = csv.writer(outfile, dialect='excel')
            writer.writerow(val)
        
print("Done")

