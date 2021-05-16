from selenium import webdriver
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
import pandas as pd


id = input("e-Arsiv Portal Kullanıcı Kodu: ")
password = input("Sifre: ")

#start browser 
driver = webdriver.Chrome()
url = "https://earsivportal.efatura.gov.tr/intragiris.html"
driver.get(url)
driver.maximize_window()

#login by id and password
driver.find_element_by_xpath('//*[@id="userid"]').send_keys(id)
driver.find_element_by_xpath('//*[@id="password"]').send_keys(password)
driver.find_element_by_xpath('//*[@id="formdiv"]/div[3]/div/button').click()
time.sleep(10)

#open the page where invoice is being maden out 
select = Select(driver.find_element_by_id('gen__1006'))
select.select_by_index(1)
time.sleep(2)
driver.find_element_by_class_name('cstree-closed').click()
driver.find_element_by_class_name('cstree-open').find_elements_by_css_selector("ul>li")[0].click()
time.sleep(2)
driver.find_element_by_tag_name("body").send_keys(Keys.PAGE_DOWN)
time.sleep(2)



#pull data from excel
file = pd.read_excel("dosya.xlsx", sheet_name="SATIS").values.tolist()

i=0
#write the data on page
for name, quantity, price in file:

    driver.find_element_by_xpath('//*[@id="gen__1088"]/div/div[1]').click() # create new row

    table = driver.find_element_by_xpath('//*[@id="gen__1069-b"]')
    row = table.find_elements_by_css_selector("tr")[i]
    product_name = row.find_elements_by_css_selector("td")[3].find_element_by_tag_name("input").send_keys(int(name))
    product_quantity = row.find_elements_by_css_selector("td")[4].find_element_by_tag_name("input").send_keys(int(quantity))
    unit = row.find_elements_by_css_selector("td")[5].find_element_by_tag_name("select")
    Select(unit).select_by_value("C62")
    product_price = row.find_elements_by_css_selector("td")[6].find_element_by_tag_name("input")
    product_price.send_keys(Keys.CONTROL + "a").send_keys(Keys.DELETE)
    product_price.send_keys(str(price).replace(".",","))
    kdv = row.find_elements_by_css_selector("td")[10].find_element_by_tag_name("select")
    Select(kdv).select_by_value("8")
    i += 1 
    time.sleep(1)
    
   