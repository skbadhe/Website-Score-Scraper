from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import xlrd
from xlutils import copy

excelpath = r"C:\temp\temp.xlsx"
workbookread= xlrd.open_workbook(r"C:\temp\temp.xls",formatting_info=True)
worksheet = workbookread.sheet_by_index(0)
listrollNumber = []
listmothername = []
for i in range(1,75):
    listrollNumber.append(worksheet.cell_value(i,1))
    listmothername.append(worksheet.cell_value(i,4))

workbookwrite = copy.copy(workbookread)
wb = workbookwrite.get_sheet(0)

driver = webdriver.Chrome('C:\Program Files (x86)\Google\Chrome\Application\chromedriver.exe')

for i in range(0,74):
    roll_no=listrollNumber[i]
    mothername=listmothername[i]

    driver.get("http://mahresult.nic.in/hscresmar20/hscresmar20.htm")
    element = driver.find_element_by_name("regno")
    element.send_keys(roll_no)
    element = driver.find_element_by_name("mname")
    element.send_keys(mothername)
    element.send_keys(Keys.RETURN)

    English = driver.find_element_by_xpath('/html/body/div/table[1]/tbody/tr[1]/td[3]').text
    Marathi = driver.find_element_by_xpath('/html/body/div/table[1]/tbody/tr[2]/td[3]').text
    Geograpghy= driver.find_element_by_xpath('/html/body/div/table[1]/tbody/tr[3]/td[3]').text
    Maths= driver.find_element_by_xpath('/html/body/div/table[1]/tbody/tr[4]/td[3]').text
    Physics= driver.find_element_by_xpath('/html/body/div/table[1]/tbody/tr[5]/td[3]').text
    Chemistry = driver.find_element_by_xpath('/html/body/div/table[1]/tbody/tr[6]/td[3]').text
    Environment= driver.find_element_by_xpath('/html/body/div/table[1]/tbody/tr[7]/td[3]').text

    wb.write(i+1, 5, int(English))
    wb.write(i+1, 6, int(Marathi))
    wb.write(i+1, 7, int(Geograpghy))
    wb.write(i+1, 8, int(Maths))
    wb.write(i+1, 9, int(Physics))
    wb.write(i+1, 10, int(Chemistry))
    wb.write(i+1, 11, int(Environment))

workbookwrite.save('save.xls')
driver.close()