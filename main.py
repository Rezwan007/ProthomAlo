import time
import xlwt

from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium import webdriver

workbook = xlwt.Workbook()

sheet = workbook.add_sheet("Data")

style= xlwt.easyxf('font: bold 1, color red;')

browser = webdriver.Chrome("C:/Users/Rezwan/Documents/Python Scripts/chromedriver.exe")
browser.get("https://www.prothomalo.com/")

print(browser.title)

sportClick = browser.find_element_by_xpath("/html/body/div[1]/header/div/div/div/div/div[2]/nav/ul/li[9]/div/a")
sportClick.click()

time.sleep(20)
try:
    selectItem = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[1]/div[8]/div/div[1]/ul/li[1]/a"))
    )
    print(selectItem.text)
    selectItem.click()
    time.sleep(5)
    sheet.write(0,0, *selectItem, style)

    selectItem2 = WebDriverWait(browser, 10).until(
        EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[1]/div[8]/div/div[2]/div/div/div[1]/div/div[1]/div[1]/div[1]/a[1]/h2"))
    )
    print(selectItem2.text)
    selectItem2.click()
    time.sleep(20)
    sheet.write(1, 0, *selectItem2, style)

    selectItem3 = WebDriverWait(browser, 20).until(
        EC.presence_of_element_located(
            (By.XPATH, "/html/body/div[1]/div[8]/div/div[1]/div/div/div/div[2]/div[1]/div/div[4]/div"))
    )
    print(selectItem3.text)
    time.sleep(2)
    sheet.write(2, 0, *selectItem3, style)
except:
    print("Something wrong going on! Remove the Add.")
    time.sleep(2)
    browser.quit()

workbook.save("data.xls")
print("All task done!")
time.sleep(20)
browser.quit()
