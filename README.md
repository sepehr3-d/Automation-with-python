# Automation-with-python
How to web scraping with excel as a DB

import xlrd
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
from openpyxl import load_workbook

# load google chrome driver and enter url
driver = webdriver.Chrome('/path/to/chromedriver')
driver.implicitly_wait(1)
driver.get("https://abadis.ir")
vacab = driver.find_element_by_id("tb_Word")

# load excel file from your computer
workbook = xlrd.open_workbook(r'C:\Users\sepehr\PycharmProjects\pythonProjects\pythonProject\Vocabulary.xlsx')
sheet = workbook.sheet_by_name("Sheet")

# i = row number of excel sheet
i = 588

#read row from excel file
for curr_row in range(1, i):
	Vocabulary = sheet.cell_value(curr_row, 0)

# search your file from a site
vacab.clear()
vacab.send_keys(Vocabulary)
vacab.send_keys(Keys.RETURN)
time.sleep(0.1)
url = driver.current_url

#search items that you want from webpage
try:
	meaning_FA = driver.find_element_by_xpath('//*[@id="Means"]/div[1]/div[1]').text
except:
	pass
try:
	meaning_EN = driver.find_element_by_xpath('//*[@id="EnToFaAnalyze"]/div[2]').text
except:
	pass
try:
	parts_of_speech = driver.find_element_by_xpath('//*[@id="EnToFaAnalyze"]/div[1]').text
except:
	pass
try:
	example = driver.find_element_by_xpath('//*[@id="EnToFaAnalyze"]/div[3]').text
except:
	pass


# load your file again for saving items in th sheet
wb = load_workbook(filename=r'C:\Users\sepehr\PycharmProjects\pythonProjects\pythonProject\Vocabulary.xlsx')
ws = wb.worksheets[0]
ws_tables = []
try:
	ws.cell(row=i, column=2).value = meaning_FA
except:
	pass
try:
	ws.cell(row=i, column=3).value = meaning_EN
except:
	pass
try:
	ws.cell(row=i, column=4).value = parts_of_speech
except:
	pass
try:
	ws.cell(row=i, column=5).value = example
except:
	pass
try:
	ws.cell(row=i, column=6).value = url
except:
	pass
driver.close()

wb.save(filename=r'C:\Users\sepehr\PycharmProjects\pythonProjects\pythonProject\Vocabulary.xlsx')

