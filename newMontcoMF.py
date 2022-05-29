#!/usr/bin/env python3
import pyperclip, webbrowser, selenium
import pyautogui, docx, os
import sys, time, warnings
import re, openpyxl, docx2txt
from selenium.webdriver import Chrome
from selenium.webdriver.common.keys import Keys

date1 = input("Enter the starting day (Thursday) in mm/dd/yyyy format: ")
date2 = input("Enter the ending day (Wednesday, or yesterday) in mm/dd/yyyy format: ")

browser = Chrome('/Users/garcgabe/Downloads/chromedriver')

# Open Montco database
browser.get('https://courtsapp.montcopa.org/psi/v/search/case')
# Date 1
searchDateElem = browser.find_element_by_css_selector('#DateCommencedFrom')
searchDateElem.send_keys(date1)
# Date 2
searchDate2Elem = browser.find_element_by_css_selector('#DateCommencedTo')
searchDate2Elem.send_keys(date2)
searchDate2Elem.submit()
### Computer Information for screen usage
d = docx.Document()
d.save('/Users/garcgabe/Downloads/NewAutomation/MontcoMF.docx')
path = '/Users/garcgabe/Downloads/NewAutomation/MontcoMF.docx'
mac = 'command'
windows = 'ctrl'
choice = 1
if choice == 1:
    key = mac
if choice == 0:
    key = windows
# horizontal values of screen for each point
x_value_print_window = 993
x_value_save_as_pdf = 950
x_value_print = 1072
# vertical values of screen for each point
y_value_print_window = 205
y_value_save_as_pdf = 222
y_value_print = 745

browser.back()
# Search MF in Montco
clickElem = browser.find_element_by_css_selector('#s2id_CaseType > a')
clickElem.click()
searchElem = browser.find_element_by_css_selector('#s2id_autogen1_search')
searchElem.send_keys('Complaint In Mortgage Foreclosure')
searchElem.send_keys(u'\ue007')
# Submitting
submitIt = browser.find_element_by_css_selector('#page-content-wrapper > div > div:nth-child(2) > div > form > div > div:nth-child(8) > div:nth-child(14) > button.btn.fa.fa-search')
submitIt.click()
############################################################################################################################
###### onto cases



## first page is separate from others ##



## BASE CASE, FIRST case
caseClick = browser.find_element_by_css_selector('#gridViewResults > table > tbody > tr:nth-child(1) > td.noprint > a')
caseClick.click()
pyautogui.click(700,700)
time.sleep(2)
pyautogui.hotkey(key, 'a')
time.sleep(0.5)
pyautogui.hotkey(key, 'c')
time.sleep(0.5)
d.add_paragraph(str(pyperclip.paste()))
d.save(path)
# saving as PDF
# on windows: pyautogui.hotkey('ctrl', 'p')
pyautogui.hotkey(key, 'p')
pyautogui.moveTo(x_value_print_window, y_value_print_window, 2)
pyautogui.click()
pyautogui.press('down') # CHANGE # PRESSES DEPENDING ON COMPUTER
pyautogui.click(x_value_save_as_pdf, y_value_save_as_pdf)
pyautogui.moveTo(x_value_print, y_value_print, 2) # subject to change (((in terminal, python3, then import pyauto, then pyautogui.displayMousePosition()
pyautogui.click()
time.sleep(3)
# save as pdf should be open
pyautogui.press('1')
time.sleep(0.5)
pyautogui.press('enter')
browser.back()

for i in range(2,21):
    caseClick = browser.find_element_by_css_selector('#gridViewResults > table > tbody > tr:nth-child(' + str(i) + ') > td.noprint > a')
    caseClick.click()
    pyautogui.click(700,700)
    time.sleep(2)
    pyautogui.hotkey(key, 'a')
    time.sleep(0.5)
    pyautogui.hotkey(key, 'c')
    time.sleep(0.5)
    d.add_paragraph(str(pyperclip.paste()))
    d.save(path)
    # saving as PDF
    # on windows: pyautogui.hotkey('ctrl', 'p')
    pyautogui.hotkey(key, 'p')
    time.sleep(0.8)
    pyautogui.moveTo(x_value_print, y_value_print, 1) # subject to change (((in terminal, python3, then import pyauto, then pyautogui.displayMousePosition()
    pyautogui.click()
    time.sleep(2.5)
    # save as pdf should be open
    numbers = [int(x) for x in str(i)]
    if (len(numbers)>1):
        pyautogui.press(str(numbers[0]))
        time.sleep(0.2)
        pyautogui.press(str(numbers[1]))
    else:
        pyautogui.press(str(numbers[0]))
    time.sleep(0.5)
    pyautogui.press('enter')
    browser.back()
nextClick = browser.find_element_by_css_selector('#gridViewResults > div:nth-child(3) > a:nth-child(2)')
nextClick.click()
## on to page 2
y = 0
z = 0
for page in range(3,10): ## 10 can be increased to scale
    z = page        ## establish next page's selector
    if (z == 3):    ## first case ==> z never equals 3, it goes from 2 to 4, 4 is for page 3
        z += 1
    y += 20   # starts at 20, add 20 each page
    for i in range(1,21):
        x = y + i
        caseClick = browser.find_element_by_css_selector('#gridViewResults > table > tbody > tr:nth-child(' + str(i) + ') > td.noprint > a')
        caseClick.click()
        pyautogui.click(700,700)
        time.sleep(1.5)
        pyautogui.hotkey(key, 'a')
        time.sleep(0.5)
        pyautogui.hotkey(key, 'c')
        time.sleep(0.5)
        d.add_paragraph(str(pyperclip.paste()))
        d.save(path)
        # saving as PDF
        # on windows: pyautogui.hotkey('ctrl', 'p')
        pyautogui.hotkey(key, 'p')
        time.sleep(0.8)
        pyautogui.moveTo(x_value_print, y_value_print, 1) # subject to change (((in terminal, python3, then import pyauto, then pyautogui.displayMousePosition()
        pyautogui.click()
        time.sleep(2.5)
        # save as pdf should be open
        numbers = [int(x) for x in str(x)]
        if (len(numbers)>2):
            pyautogui.press(str(numbers[0]))
            time.sleep(0.2)
            pyautogui.press(str(numbers[1]))
            time.sleep(0.2)
            pyautogui.press(str(numbers[2]))
        elif (len(numbers)>1):
            pyautogui.press(str(numbers[0]))
            time.sleep(0.2)
            pyautogui.press(str(numbers[1]))
        else:
            pyautogui.press(str(numbers[0]))
        time.sleep(0.5)
        pyautogui.press('enter')
        browser.back()
    nextClick = browser.find_element_by_css_selector('#gridViewResults > div:nth-child(3) > a:nth-child(' + str(z) + ')')
    nextClick.click()
