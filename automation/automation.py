import time
import urllib.request

import xlrd
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.keys import Keys


def downloadImage(url):

    file_loc = "/users/saurabhbansal/desktop/Product_down.xlsx"
    arr = {}
    try:
        file = xlrd.open_workbook(file_loc,"r")
        sheet = file.sheet_by_index(0)
        for i in range(sheet.nrows):
            arr[i] = sheet.cell_value(i,0)
    except FileNotFoundError:
        print("Check the Source Path")

    try:
        for j in range(1,(len(arr))):

            driver = webdriver.Chrome(executable_path="/Users/saurabhbansal/PycharmProjects/automation/venv/bin/chromedriver")
            print(driver)
            driver.get(url)
            productName = arr[j]
            id_box = driver.find_element_by_name("search")
            id_box.send_keys(arr[j])
            time.sleep(5)
            id_box.send_keys(Keys.ENTER)
            time.sleep(5)
            img_box = driver.find_element_by_xpath("//div[@class='image']")
            (img_box.click())
            time.sleep(20)

            containers = driver.find_elements_by_xpath('.//div[@class="outer_container style-3  "]')

            for contain in containers:

                j = 1
                zoom_wrap = contain.find_element_by_xpath('.//div[@class="cloud-zoom-wrap"]')
                zoom = zoom_wrap.find_element_by_xpath('.//a[@id="zoom1"]')
                imageprop = zoom.find_element_by_xpath('.//img[@itemprop="image"]')
                base_url = (imageprop.get_attribute("src"))
                f_name = productName + "-" + str(j) + ".jpg"
                urllib.request.urlretrieve(base_url, f_name)
                time.sleep(10)
                zoom_button = contain.find_element_by_xpath('.//a[@id="zoom-btn"]')
                #zoom_button.click()
                if(len(contain.find_elements_by_xpath('.//div[@class="image-additional"]')) > 0):

                    image_additional = contain.find_element_by_xpath('.//div[@class="image-additional"]')
                    ul_class = image_additional.find_element_by_xpath('.//ul[@class="image_carousel owl-carousel owl-theme"]')
                    owl_outer = ul_class.find_element_by_xpath('.//div[@class="owl-wrapper-outer"]')
                    owl_wrapper = owl_outer.find_element_by_xpath('.//div[@class="owl-wrapper"]')
                    owl_item = owl_wrapper.find_elements_by_xpath('.//div[@class="owl-item"]')
                    i = 0
                    time.sleep(10)
                    print(len(owl_item))
                    while i < len(owl_item):
                        try:
                            zoom_gallery = owl_item[i].find_element_by_xpath('.//a[@class="cloud-zoom-gallery colorbox cboxElement"]')
                            link = (zoom_gallery.get_attribute("href"))
                            j += 1
                            file_name = (productName + "-" +str(j) +'.jpg')
                            urllib.request.urlretrieve(link, file_name)
                        except NoSuchElementException:
                            break
                        i = i + 1
                        time.sleep(10)
                else:
                    print("Element length is smaller than 1")
            driver.close()
    except (NoSuchElementException, IndexError, KeyboardInterrupt):
        print("Element Not Found, Kindly, refer the inspector or check your connection")
        driver.close()


downloadImage("https://www.armenliving.com")