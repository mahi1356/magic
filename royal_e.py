
#this is the first draft 
from selenium import webdriver 
from selenium.webdriver.common.keys import Keys
import time 
from selenium.webdriver.support.ui import Select


driver = webdriver.Firefox()
driver.get('http://www.royalmail.com/price-finder')

#click next button
next_a = driver.find_element_by_xpath("//div[@id='where-choice-panel-']/div/div[2]/div/div/div")
next_a.click()

time.sleep(1)
#select parcel
d= driver.find_element_by_xpath("//li[3]/div/div[9]")
d.click()

time.sleep(1)
#select first item using Select class Selenium provides  
select_parcel =Select(driver.find_element_by_xpath("//select[@id='weightselect']"))
options = select_parcel.options
# print(options) vs. print(len(options)) what is options?? 
for index in range(0, len(options) - 1):
    select_parcel.select_by_index(index)
    print("this one:"+ str(index))
time.sleep(1)
next_b = driver.find_element_by_xpath("//div[3]/div/div") 
next_b.click()
time.sleep(1)
#select med parcel 
select_med_parcel =driver.find_element_by_xpath("//li[@id='size-choice-tab-1']/div/div[9]/div[2]/div")
select_med_parcel.click()
time.sleep(1)
#select price
select_price_list =driver.find_element_by_xpath("//div[@id='size-choice-panel-']/div/div[2]/div/div/div")
select_price_list.click()
    
time.sleep(10)
driver.close()