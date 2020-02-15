import csv
import re
import sys    
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import requests
from bs4 import BeautifulSoup, Comment

def get_more (page_num, start_list, head_list, idx):
    #get all the links to view rec

    offset = idx - 1
    idx = idx + offset
         
    length = driver.find_elements_by_xpath("(//a[@class = 'commandLinkSmall'])")

    #we want only every other link, so access by index
    view_rec = "(//a[@class = 'commandLinkSmall']) [" + str(idx)  + "]"
    driver.find_element_by_xpath(view_rec).click()


    #rec details is always in 1 of the the first 3 tabs, so we only check those
    j = 1
    while j < 3:
        my_string = "(//a[@class = 'tabUnselectedText']) [" + str(j) + "]"
        links = driver.find_element_by_xpath(my_string)
        if (links.text.find("Rec Details") >= 0):
            links.click()

        j = j + 1

    #grab the page
    driver.implicitly_wait(100)
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    plz = ''
    rec_list = []

    table_info = soup.findAll("td", {"class": "tableInLtShade"})
    table_headers = soup.findAll("td", {"class": "tableInDeepShade"})


    j = 1
    #the last two cells dont contain any specific info to the listing
    info_list = []
    header_list = []
    
    while (j < len(table_info) - 2):

        #get headers on first pass, they dont change
        if (idx == 1):
            refined_headers = re.split(r'\s{2,}', table_headers[j].get_text())
            header_list.append(refined_headers[1])

        #remove newlines and carriage returns
        refined_info = re.split(r'\s{2,}', table_info[j].get_text())
        long_pieces = len(refined_info)
        new_ref = ""

        if (long_pieces > 3 and refined_info[1] != ''):
            k = 1
            while k < long_pieces - 1:
                new_ref = new_ref + refined_info[k] + "\r"
                k = k + 1
            info_list.append (new_ref)
        else:
            info_list.append (refined_info[1])
        
        j = j + 1
        #jump past if we encounter any of these blank cells
        if (j == 8):
            j = j + 1
        elif (j == 15):
            j = j + 1
        elif (j == 27):
            j = j + 1
        elif (j == 34):
            j = j + 1
        elif (j == 43):
            j = j + 1
        elif (j == 49):
            j = j + 1

    #print(info_list)
    if (idx == 1 and page_num == 1):
        head_list = head_list + header_list
        
        with open('lel.csv','w', newline = '') as outfile:
            writer = csv.writer(outfile, dialect='excel')
            writer.writerow(head_list)

        print(head_list)   

    info_list = start_list + info_list
    fixed_list = []
    
    k = 0
    while k < len(info_list):
        fix = info_list[k].replace('\u2010', '-')
        fix = fix.replace('\u25cf', '-')
        fix = fix.replace('\u037e', ';')
        
        fixed_list.append(fix)
        k = k + 1

    print(fixed_list)
    with open('lel.csv','a', newline = '') as outfile:
            writer = csv.writer(outfile, dialect='excel')
            writer.writerow(fixed_list)
    #print(header_list)
    


    driver.find_element_by_xpath("(//input[@value = 'Cancel'])").click()
    driver.find_element_by_xpath("(//input[@value = 'Confirm Cancel'])").click()

    #if we were on a page other than page 1, go back to that page
    if (page_num > 1):
        page_input = driver.find_element_by_name("ftPageNumber")
        page_input.clear()
        page_input.send_keys(str(page_num))
        driver.find_element_by_xpath("(//input[@value = 'Go'])").click()
      


def get_elements (soup, page_num, head_list, set_idx):
    #find the table by traversing back up nested tags
    table = soup.find(text = 'Job Title').parent.parent.parent.parent.parent.parent.parent.parent.parent

    #get us the rows
    rows = table.find_all("td")
    my_string = ''
    row_list = []
    my_list = []

    #get the text of the rows and remove empty lines, replace commas, and acknowledge null entries. take each line and add it as an element to row_list
    j = 0
    while j < len(rows):

        refined_list = re.split(r'\s{2,}', rows[j].get_text())
        if (refined_list[0] != ''):
            this = refined_list[0].replace('\n', '')
            this = this.replace(',', ' -')
            row_list.append(this.replace('\xa0', 'n/a'))
        j = j + 1
    
    #11 attributes from front page
    k = 11 * set_idx
    if k < len(row_list):
        j = 0
        while j < 11:
            my_list.append(row_list[k])
            j = j + 1
            k = k + 1
        print (len(row_list))

    return my_list

def recursive (page_num, head_list):

    soup = BeautifulSoup(driver.page_source, 'html.parser')
    set_idx = 1
    while set_idx < 26:
        start_list = get_elements(soup, page_num, head_list, set_idx)
        get_more(page_num, start_list, head_list, set_idx)
        set_idx = set_idx + 1


    page_num = page_num + 1

    try:
        if (driver.find_element_by_id("functionalityTableDelegate_getNextRowSet")):
            driver.find_element_by_id("functionalityTableDelegate_getNextRowSet").click()
            recursive(page_num, head_list)
    except NoSuchElementException:
        return 0
        

#open chrome and go to rms
driver = webdriver.Chrome()
driver.get("https://jobs.ucsc.edu/userfiles/jsp/shared/frameset/Frameset.jsp?time=1348080164562")

#let page load
driver.implicitly_wait(10)
driver.switch_to.frame("contentFrame")

#enter username and password
username = driver.find_element_by_id("userName")
username.clear()
username.send_keys("elancast")
password = driver.find_element_by_name("password")
password.clear()
password.send_keys("1224Eel1")
driver.find_element_by_name("submit").click()

head_list = ['Job Title','Job Number', 'Rec. Type', 'Div/Org', 'CHM', 'Rec. Status', 'Status Date', 'IRD', 'Total Apps.', 'Incomplete Job Offers', 'Hires']
global_list = []
page_num = 1

driver.switch_to.frame("contentFrame")
recursive(page_num, head_list)
