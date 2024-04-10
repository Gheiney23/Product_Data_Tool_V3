import pandas as pd
import time
import re
import PySimpleGUI as sg
from urllib.request import urlretrieve
from openpyxl import load_workbook
from selenium import webdriver as wb
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from openpyxl.worksheet.properties import WorksheetProperties as wp
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.service import Service

conv_df2 = {'Sku': [], 'Image_2_Name': [], 'Image_2_Failed': []}
conv_df3 = {'Sku': [], 'Image_3_Name': [], 'Image_3_Failed': []}
conv_df4 = {'Sku': [], 'Image_4_Name': [], 'Image_4_Failed': []}
conv_df5 = {'Sku': [], 'Image_5_Name': [], 'Image_5_Failed': []}
conv_df6 = {'Sku': [], 'Image_6_Name': [], 'Image_6_Failed': []}


def Build_tool(sku_list, uid_1_list, uid_2_list):
    # Setting up the webdriver for Selenium
    service = Service()
    options = wb.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument("--disable-notifications")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = wb.Chrome(service=service, options=options)

    # opener = urllib.request.build_opener()
    # opener.addheaders = [('User-Agent', 'MyApp/1.0')]
    # urllib.request.install_opener(opener)

    data_dict = {'Uniqueid': [],
                 'Sku': [],
                 'Category': [],
                 'Product_Title': [],
                 'Marketing_Copy': [],
                 'Img_url1': [],
                 'Img_url2': [],
                 'Img_url3': [],
                 'Img_url4': [],
                 'Img_url5': [],
                 'Img_url6': [],
                 'Img_url7': [],
                 'Img_url8': [],
                 'Img_url9': [],
                 'Bullet1': [],
                 'Bullet2': [],
                 'Bullet3': [],
                 'Bullet4': [],
                 'Bullet5': [],
                 'Bullet6': [],
                 'Bullet7': [],
                 'Bullet8': [],
                 'Bullet9': [],
                 'PDF_1': [],
                 'PDF_2': [],
                 'PDF_3': [],
                 'PDF_4': [],
                 'Skus_Not_Found': []}

    # build_path = 'https://www.build.com/showroom'
    # driver.get(build_path)
    # time.sleep(45)
    # driver.execute_script("window.open('');")

    for (uid_1, uid_2, sku) in zip(uid_1_list, uid_2_list, sku_list):
        try:
            path = 'https://www.build.com/pfister-hhl-089tb/s{}?uid={}'.format(uid_2, uid_1)
            driver.get(path)
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "root")))

            # time.sleep(1)
            driver.execute_script("window.scrollTo(0, 300)")

            # Adding sku to dictionary
            data_dict['Sku'].append(sku)
            data_dict['Uniqueid'].append(uid_1)

            # Extracting the category
            try:

                category = driver.find_element(By.XPATH,
                                               "//*[@id='main-content']/div/section[1]/div[1]/nav/ol/li[4]/a/span").text
                data_dict['Category'].append(category)
            except:
                data_dict['Category'].append('NULL')

            # Extracting product title
            try:
                desc = driver.find_element(By.XPATH, "//h1[@class='ma0 fw6 lh-title di f5 f3-ns']").text
                data_dict['Product_Title'].append(desc)
            except:
                data_dict['Product_Title'].append('NULL')

            #  Extracting marketing copy
            try:
                marketing_copy = driver.find_element(By.XPATH,
                                                     "/html/body/div[2]/div/main/div[1]/section[4]/div[1]/section/div[1]/section/div[2]/div[1]/div[1]/p[1]").text
                data_dict['Marketing_Copy'].append(marketing_copy)
            except:
                data_dict['Marketing_Copy'].append('NULL')

            # Extracting the image src
            src_1 = driver.find_element(By.XPATH, "//*[contains(@class,'w-auto self-center undefined')]").get_attribute(
                'src')
            # time.sleep(1)
            data_dict['Img_url1'].append(src_1)

            try:
                element = driver.find_element(By.XPATH, "//div[@aria-label='thumb slide 1']")
                actions = ActionChains(driver)
                actions.move_to_element(element).perform()
                element.click()
                time.sleep(1)
                src_2 = driver.find_element(By.XPATH, "//div[starts-with(@aria-label, '2 /')]").find_element(
                    By.TAG_NAME, 'img').get_attribute('src')
                data_dict['Img_url2'].append(src_2)
            except:
                data_dict['Img_url2'].append('NULL')

            try:
                element = driver.find_element(By.XPATH, "//div[@aria-label='thumb slide 2']")
                actions = ActionChains(driver)
                actions.move_to_element(element).perform()
                element.click()
                time.sleep(1)
                src_3 = driver.find_element(By.XPATH, "//div[starts-with(@aria-label, '3 /')]").find_element(
                    By.TAG_NAME, 'img').get_attribute('src')
                data_dict['Img_url3'].append(src_3)
            except:
                data_dict['Img_url3'].append('NULL')

            try:
                element = driver.find_element(By.XPATH, "//div[@aria-label='thumb slide 3']")
                actions = ActionChains(driver)
                actions.move_to_element(element).perform()
                element.click()
                time.sleep(1)
                src_4 = driver.find_element(By.XPATH, "//div[starts-with(@aria-label, '4 /')]").find_element(
                    By.TAG_NAME, 'img').get_attribute('src')
                data_dict['Img_url4'].append(src_4)
            except:
                data_dict['Img_url4'].append('NULL')

            try:
                element = driver.find_element(By.XPATH, "//div[@aria-label='thumb slide 4']")
                actions = ActionChains(driver)
                actions.move_to_element(element).perform()
                element.click()
                time.sleep(1)
                src_5 = driver.find_element(By.XPATH, "//div[starts-with(@aria-label, '5 /')]").find_element(
                    By.TAG_NAME, 'img').get_attribute('src')
                data_dict['Img_url5'].append(src_5)
            except:
                data_dict['Img_url5'].append('NULL')

            try:
                element = driver.find_element(By.XPATH, "//div[@aria-label='thumb slide 5']")
                actions = ActionChains(driver)
                actions.move_to_element(element).perform()
                element.click()
                time.sleep(1)
                src_6 = driver.find_element(By.XPATH, "//div[starts-with(@aria-label, '6 /')]").find_element(
                    By.TAG_NAME, 'img').get_attribute('src')
                data_dict['Img_url6'].append(src_6)
            except:
                data_dict['Img_url6'].append('NULL')

            try:
                element = driver.find_element(By.XPATH, "//div[@aria-label='thumb slide 6']")
                actions = ActionChains(driver)
                actions.move_to_element(element).perform()
                element.click()
                time.sleep(1)
                src_7 = driver.find_element(By.XPATH, "//div[starts-with(@aria-label, '7 /')]").find_element(
                    By.TAG_NAME, 'img').get_attribute('src')
                data_dict['Img_url7'].append(src_7)
            except:
                data_dict['Img_url7'].append('NULL')

            try:
                element = driver.find_element(By.XPATH, "//div[@aria-label='thumb slide 7']")
                actions = ActionChains(driver)
                actions.move_to_element(element).perform()
                element.click()
                time.sleep(1)
                src_8 = driver.find_element(By.XPATH, "//div[starts-with(@aria-label, '8 /')]").find_element(
                    By.TAG_NAME, 'img').get_attribute('src')
                data_dict['Img_url8'].append(src_8)
            except:
                data_dict['Img_url8'].append('NULL')

            try:
                element = driver.find_element(By.XPATH, "//div[@aria-label='thumb slide 8']")
                actions = ActionChains(driver)
                actions.move_to_element(element).perform()
                element.click()
                time.sleep(1)
                src_9 = driver.find_element(By.XPATH, "//div[starts-with(@aria-label, '9 /')]").find_element(
                    By.TAG_NAME, 'img').get_attribute('src')
                data_dict['Img_url9'].append(src_9)
            except:
                data_dict['Img_url9'].append('NULL')

            # Extracting bullet points
            element = driver.find_element(By.XPATH, "//div[@class='lh-copy H_oFW']")
            actions = ActionChains(driver)
            actions.move_to_element(element).perform()
            # driver.execute_script("window.scrollTo(0, 800)")
            time.sleep(3)

            # Extracting the webelement then transforming it to text with no special characters
            li_elements = element.find_elements(By.TAG_NAME, 'li')
            b_list = []

            for li in li_elements:
                li_text = li.text
                bullet = re.sub("[^A-Za-z0-9 -\/]", "", li_text)
                bullet = bullet.replace('"', "-in")
                b_list.append(bullet)

            # Loading all bullet points found into the data_dict
            try:
                data_dict['Bullet1'].append(b_list[0])
            except:
                data_dict['Bullet1'].append('NULL')

            try:
                data_dict['Bullet2'].append(b_list[1])
            except:
                data_dict['Bullet2'].append('NULL')

            try:
                data_dict['Bullet3'].append(b_list[2])
            except:
                data_dict['Bullet3'].append('NULL')

            try:
                data_dict['Bullet4'].append(b_list[3])
            except:
                data_dict['Bullet4'].append('NULL')

            try:
                data_dict['Bullet5'].append(b_list[4])
            except:
                data_dict['Bullet5'].append('NULL')

            try:
                data_dict['Bullet6'].append(b_list[5])
            except:
                data_dict['Bullet6'].append('NULL')

            try:
                data_dict['Bullet7'].append(b_list[6])
            except:
                data_dict['Bullet7'].append('NULL')

            try:
                data_dict['Bullet8'].append(b_list[7])
            except:
                data_dict['Bullet8'].append('NULL')

            try:
                data_dict['Bullet9'].append(b_list[8])
            except:
                data_dict['Bullet9'].append('NULL')

            # Extracting pdfs
            driver.execute_script("window.scrollTo(0, 600)")
            time.sleep(1)

            hrefs = driver.find_elements(By.XPATH,
                                         "//a[@class='f-inherit fw-inherit link theme-primary  pb3 f7 db underline-hover']")
            href_list = []

            for href in hrefs:
                href = href.get_attribute('href')
                href_list.append(href)

            try:
                data_dict['PDF_1'].append(href_list[0])
            except:
                data_dict['PDF_1'].append('NULL')

            try:
                data_dict['PDF_2'].append(href_list[1])
            except:
                data_dict['PDF_2'].append('NULL')

            try:
                data_dict['PDF_3'].append(href_list[2])
            except:
                data_dict['PDF_3'].append('NULL')

            try:
                data_dict['PDF_4'].append(href_list[3])
            except:
                data_dict['PDF_4'].append('NULL')

        except:      
            data_dict['Skus_Not_Found'].append(sku)
            data_dict['Img_url1'].append('NULL')
            data_dict['Img_url2'].append('NULL')
            data_dict['Img_url3'].append('NULL')
            data_dict['Img_url4'].append('NULL')
            data_dict['Img_url5'].append('NULL')
            data_dict['Img_url6'].append('NULL')
            data_dict['Img_url7'].append('NULL')
            data_dict['Img_url8'].append('NULL')
            data_dict['Img_url9'].append('NULL')
            data_dict['Bullet1'].append('NULL')
            data_dict['Bullet3'].append('NULL')
            data_dict['Bullet4'].append('NULL')
            data_dict['Bullet5'].append('NULL')
            data_dict['Bullet6'].append('NULL')
            data_dict['Bullet7'].append('NULL')
            data_dict['Bullet2'].append('NULL')
            data_dict['Bullet8'].append('NULL')
            data_dict['Bullet9'].append('NULL')
            data_dict['PDF_1'].append('NULL')
            data_dict['PDF_2'].append('NULL')
            data_dict['PDF_3'].append('NULL')
            data_dict['PDF_4'].append('NULL')

    # quitting the driver and manipulation the dictionary into a dataframe
    driver.quit()

    df = pd.DataFrame.from_dict(data_dict, orient='index')
    df = df.transpose()

    # Writing the dataframe to an excel worksheet
    df.to_excel('Build_Data.xlsx', sheet_name='Build_Data')


def Ferg_Tool(sku_list, uniqueid_list):
    # Setting up the webdriver for Selenium
    service = Service()
    options = wb.ChromeOptions()
    options.add_argument('--start-maximized')
    options.add_argument("--disable-notifications")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = wb.Chrome(service=service, options=options)

    # src_dict = {'Sku': [], 'Img_url1': [], 'Img_url2': [], 'Img_url3': [], 'Img_url4': [], 'Img_url5': [], 'Img_url6': [], 'Img_url7': [], 'Img_url8': [], 'Img_url9': [], 'Skus_Not_Found': []}
    data_dict = {'Uniqueid': [],
                 'Sku': [],
                 'Img_url1': [],
                 'Img_url2': [],
                 'Img_url3': [],
                 'Img_url4': [],
                 'Img_url5': [],
                 'Img_url6': [],
                 'Img_url7': [],
                 'Img_url8': [],
                 'Img_url9': [],
                 'Product_Title': [],
                 'Category': [],
                 'PDF_1': [],
                 'PDF_2': [],
                 'PDF_3': [],
                 'PDF_4': [],
                 'Bullet1': [],
                 'Bullet2': [],
                 'Bullet3': [],
                 'Bullet4': [],
                 'Bullet5': [],
                 'Bullet6': [],
                 'Bullet7': [],
                 'Bullet8': [],
                 'Bullet9': [],
                 'Sku_Not_Found': []}

    for uid, sku in zip(uniqueid_list,sku_list):
        try:
            path = 'https://www.ferguson.com/'
            driver.get(path)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.NAME, "q")))
            time.sleep(1)
            driver.find_element(By.NAME, "q").click()
            time.sleep(1)
            driver.find_element(By.NAME, "q").send_keys(sku)
            driver.find_element(By.NAME, "q").send_keys(Keys.RETURN)
            # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@class,'hero__img')]")))
            time.sleep(5)

            # Adding sku to dictionary
            data_dict['Sku'].append(sku)
            data_dict['Uniqueid'].append(uid)

            # Extracting the product title
            try:
                title = driver.find_element(By.XPATH,
                                            "/html/body/div[1]/div[3]/div[2]/div[3]/div[2]/div[1]/div[2]/h1").text
                data_dict['Product_Title'].append(title)
            except:
                data_dict['Product_Title'].append('NULL')

            driver.execute_script("window.scrollTo(0, 200)")

            # Extracting the image src
            try:
                src_1 = driver.find_element(By.XPATH,
                                            "//img[@class='img-fluid js-zoom-image  c-product-details__img lazyloaded']").get_attribute(
                    'src')
                data_dict['Img_url1'].append(src_1)
            except:
                data_dict['Img_url1'].append('NULL')

            try:
                element = driver.find_element(By.XPATH, "//div[@aria-label='slide 2']")
                # actions = ActionChains(driver)
                # actions.move_to_element(element).perform()
                element.click()
                time.sleep(3)
                src_2 = driver.find_element(By.XPATH,
                                            "//img[@class='img-fluid js-zoom-image  c-product-details__img lazyloaded']").get_attribute(
                    'src')
                data_dict['Img_url2'].append(src_2)
            except:
                data_dict['Img_url2'].append('NULL')

            try:
                element = driver.find_element(By.XPATH, "//div[@data-index='2']")
                # actions = ActionChains(driver)
                # actions.move_to_element(element).perform()
                element.click()
                time.sleep(3)
                src_3 = driver.find_element(By.XPATH,
                                            "//img[@class='img-fluid js-zoom-image  c-product-details__img lazyloaded']").get_attribute(
                    'src')
                data_dict['Img_url3'].append(src_3)
            except:
                data_dict['Img_url3'].append('NULL')

            try:
                element = driver.find_element(By.XPATH, "//div[@data-index='3']")
                # actions = ActionChains(driver)
                # actions.move_to_element(element).perform()
                element.click()
                time.sleep(3)
                src_4 = driver.find_element(By.XPATH,
                                            "//img[@class='img-fluid js-zoom-image  c-product-details__img lazyloaded']").get_attribute(
                    'src')
                data_dict['Img_url4'].append(src_4)
            except:
                data_dict['Img_url4'].append('NULL')

            try:
                element = driver.find_element(By.XPATH, "//div[@data-index='4']")
                # actions = ActionChains(driver)
                # actions.move_to_element(element).perform()
                element.click()
                time.sleep(3)
                src_5 = driver.find_element(By.XPATH,
                                            "//img[@class='img-fluid js-zoom-image  c-product-details__img lazyloaded']").get_attribute(
                    'src')
                data_dict['Img_url5'].append(src_5)
            except:
                data_dict['Img_url5'].append('NULL')

            try:
                element = driver.find_element(By.XPATH, "//div[@data-index='5']")
                # actions = ActionChains(driver)
                # actions.move_to_element(element).perform()
                element.click()
                time.sleep(3)
                src_6 = driver.find_element(By.XPATH,
                                            "//img[@class='img-fluid js-zoom-image  c-product-details__img lazyloaded']").get_attribute(
                    'src')
                data_dict['Img_url6'].append(src_6)
            except:
                data_dict['Img_url6'].append('NULL')

            try:
                element = driver.find_element(By.XPATH, "//div[@data-index='6']")
                # actions = ActionChains(driver)
                # actions.move_to_element(element).perform()
                element.click()
                time.sleep(3)
                src_7 = driver.find_element(By.XPATH,
                                            "//img[@class='img-fluid js-zoom-image  c-product-details__img lazyloaded']").get_attribute(
                    'src')
                data_dict['Img_url7'].append(src_7)
            except:
                data_dict['Img_url7'].append('NULL')

            try:
                element = driver.find_element(By.XPATH, "//div[@data-index='7']")
                # actions = ActionChains(driver)
                # actions.move_to_element(element).perform()
                element.click()
                time.sleep(3)
                src_8 = driver.find_element(By.XPATH,
                                            "//img[@class='img-fluid js-zoom-image  c-product-details__img lazyloaded']").get_attribute(
                    'src')
                data_dict['Img_url8'].append(src_8)
            except:
                data_dict['Img_url8'].append('NULL')

            try:
                element = driver.find_element(By.XPATH, "//div[@data-index='8']")
                # actions = ActionChains(driver)
                # actions.move_to_element(element).perform()
                element.click()
                time.sleep(3)
                src_9 = driver.find_element(By.XPATH,
                                            "//img[@class='img-fluid js-zoom-image  c-product-details__img lazyloaded']").get_attribute(
                    'src')
                data_dict['Img_url9'].append(src_9)
            except:
                data_dict['Img_url9'].append('NULL')

            # Extracting the category
            try:
                category = driver.find_element(By.XPATH,
                                               "/html/body/div[1]/div[3]/div/div[1]/div/div/div/div/ol/li[5]/a").text
                data_dict['Category'].append(category)
            except:
                data_dict['Category'].append('NULL')

            # Extracting pdfs
            driver.execute_script("window.scrollTo(0, 1000)")
            time.sleep(1)

            hrefs = driver.find_elements(By.XPATH, "//div[@class='col-auto my-auto']/a")
            # a_elements = driver.find_elements(By.TAG_NAME, 'a')
            # actions = ActionChains(driver)
            # actions.move_to_element(hrefs).perform()

            href_list = []

            for href in hrefs:
                href = href.get_attribute('href')
                href_list.append(href)

            try:
                data_dict['PDF_1'].append(href_list[0])
            except:
                data_dict['PDF_1'].append('NULL')

            try:
                data_dict['PDF_2'].append(href_list[1])
            except:
                data_dict['PDF_2'].append('NULL')

            try:
                data_dict['PDF_3'].append(href_list[2])
            except:
                data_dict['PDF_3'].append('NULL')

            try:
                data_dict['PDF_4'].append(href_list[3])
            except:
                data_dict['PDF_4'].append('NULL')

            # Extracting bullet points
            try:
                ul_element = driver.find_element(By.XPATH,
                                                 "/html/body/div[1]/div[3]/div[2]/div[4]/div[2]/div[1]/div[2]/div/ul")
                actions = ActionChains(driver)
                actions.move_to_element(ul_element).perform()
                # driver.execute_script("window.scrollTo(0, 800)")
                time.sleep(3)

                li_elements = ul_element.find_elements(By.TAG_NAME, 'li')
                b_list = []

                for li in li_elements:
                    li_text = li.text
                    bullet = re.sub("[^A-Za-z0-9 -\/]", "", li_text)
                    bullet = bullet.replace('"', "-in")
                    b_list.append(bullet)

                # Loading all bullet points found into the data_dict
                try:
                    data_dict['Bullet1'].append(b_list[0])
                except:
                    data_dict['Bullet1'].append('NULL')

                try:
                    data_dict['Bullet2'].append(b_list[1])
                except:
                    data_dict['Bullet2'].append('NULL')

                try:
                    data_dict['Bullet3'].append(b_list[2])
                except:
                    data_dict['Bullet3'].append('NULL')

                try:
                    data_dict['Bullet4'].append(b_list[3])
                except:
                    data_dict['Bullet4'].append('NULL')

                try:
                    data_dict['Bullet5'].append(b_list[4])
                except:
                    data_dict['Bullet5'].append('NULL')

                try:
                    data_dict['Bullet6'].append(b_list[5])
                except:
                    data_dict['Bullet6'].append('NULL')

                try:
                    data_dict['Bullet7'].append(b_list[6])
                except:
                    data_dict['Bullet7'].append('NULL')

                try:
                    data_dict['Bullet8'].append(b_list[7])
                except:
                    data_dict['Bullet8'].append('NULL')

                try:
                    data_dict['Bullet9'].append(b_list[8])
                except:
                    data_dict['Bullet9'].append('NULL')

            except:
                data_dict['Bullet1'].append('NULL')
                data_dict['Bullet2'].append('NULL')
                data_dict['Bullet3'].append('NULL')
                data_dict['Bullet4'].append('NULL')
                data_dict['Bullet5'].append('NULL')
                data_dict['Bullet6'].append('NULL')
                data_dict['Bullet7'].append('NULL')
                data_dict['Bullet8'].append('NULL')
                data_dict['Bullet9'].append('NULL')

        except:
            driver.find_element(By.NAME, "q").clear()
            data_dict['Uniqueid'].append('NULL')
            data_dict['PDF_1'].append('NULL')
            data_dict['PDF_2'].append('NULL')
            data_dict['PDF_3'].append('NULL')
            data_dict['PDF_4'].append('NULL')
            data_dict['Bullet1'].append('NULL')
            data_dict['Bullet2'].append('NULL')
            data_dict['Bullet3'].append('NULL')
            data_dict['Bullet4'].append('NULL')
            data_dict['Bullet5'].append('NULL')
            data_dict['Bullet6'].append('NULL')
            data_dict['Bullet7'].append('NULL')
            data_dict['Bullet8'].append('NULL')
            data_dict['Bullet9'].append('NULL')
            data_dict['Sku_Not_Found'].append(sku)

    driver.quit()

    df = pd.DataFrame.from_dict(data_dict, orient='index')
    df = df.transpose()

    df.to_excel('Ferg_Site_Data.xlsx', sheet_name='Ferg_Site_Data')


def converter_tool(mfg_list_primary,
                   mfg_list_2,
                   mfg_list_3,
                   mfg_list_4,
                   mfg_list_5,
                   mfg_list_6,
                   Primary_url_list,
                   img_2_url_list,
                   img_3_url_list,
                   img_4_url_list,
                   img_5_url_list,
                   img_6_url_list,
                   folder_name):
    # Creating a folder variable for output
    output_directory = '{}'.format(folder_name)

    # Looping through the dictionary and creating jpgs from the urls and loading the file names into a list
    primary_df = {'Sku': [], 'Primary_File_name': [], 'Primary_Failed': []}

    for (mfg, url) in zip(mfg_list_primary, Primary_url_list):
        try:
            primary_file_name = mfg + '_Primary.jpg'
            urlretrieve(url, output_directory + f"\{primary_file_name}")
            primary_df['Primary_File_name'].append(primary_file_name)
            primary_df['Sku'].append(mfg)
            primary_df['Primary_Failed'].append('NULL')

        except:
            primary_df['Primary_Failed'].append(mfg)

    primary_df = pd.DataFrame.from_dict(primary_df).fillna('NULL')

    def dataframe2():
        global conv_df2

        for (mfg, url) in zip(mfg_list_2, img_2_url_list):
            try:
                file_name2 = mfg + '_img2.jpg'
                urlretrieve(url, output_directory + f"\{file_name2}")
                conv_df2['Image_2_Name'].append(file_name2)
                conv_df2['Sku'].append(mfg)
                conv_df2['Image_2_Failed'].append('NULL')

            except:
                conv_df2['Image_2_Failed'].append(mfg)

        conv_df2 = pd.DataFrame.from_dict(conv_df2).fillna('NULL')

    if any(mfg_list_2):
        dataframe2()
    else:
        global conv_df2
        conv_df2 = pd.DataFrame.from_dict(conv_df2).fillna('NULL')

    def dataframe3():
        global conv_df3

        for (mfg, url) in zip(mfg_list_3, img_3_url_list):
            try:
                file_name3 = mfg + '_img3.jpg'
                urlretrieve(url, output_directory + f"\{file_name3}")
                conv_df3['Image_3_Name'].append(file_name3)
                conv_df3['Sku'].append(mfg)
                conv_df3['Image_3_Failed'].append('NULL')

            except:
                conv_df3['Image_3_Failed'].append(mfg)

        conv_df3 = pd.DataFrame.from_dict(conv_df3).fillna('NULL')

    if any(mfg_list_3):
        dataframe3()
    else:
        global conv_df3
        conv_df3 = pd.DataFrame.from_dict(conv_df3).fillna('NULL')

    def dataframe4():
        global conv_df4

        for (mfg, url) in zip(mfg_list_4, img_4_url_list):
            try:
                file_name4 = mfg + '_img4.jpg'
                urlretrieve(url, output_directory + f"\{file_name4}")
                conv_df4['Image_4_Name'].append(file_name4)
                conv_df4['Sku'].append(mfg)
                conv_df4['Image_4_Failed'].append('NULL')

            except:
                conv_df4['Image_4_Failed'].append(mfg)

        conv_df4 = pd.DataFrame.from_dict(conv_df4).fillna('NULL')

    if any(mfg_list_4):
        dataframe4()
    else:
        global conv_df4
        conv_df4 = pd.DataFrame.from_dict(conv_df4).fillna('NULL')

    def dataframe5():
        global conv_df5

        for (mfg, url) in zip(mfg_list_5, img_5_url_list):
            try:
                file_name5 = mfg + '_img5.jpg'
                urlretrieve(url, output_directory + f"\{file_name5}")
                conv_df5['Image_5_Name'].append(file_name5)
                conv_df5['Sku'].append(mfg)
                conv_df5['Image_5_Failed'].append('NULL')

            except:
                conv_df5['Image_5_Failed'].append(mfg)

    if any(mfg_list_5):
        dataframe5()
    else:
        global conv_df5
        conv_df5 = pd.DataFrame.from_dict(conv_df5).fillna('NULL')

    conv_df5 = pd.DataFrame.from_dict(conv_df5).fillna('NULL')

    def dataframe6():
        global conv_df6

        for (mfg, url) in zip(mfg_list_6, img_6_url_list):
            try:
                file_name6 = mfg + '_img6.jpg'
                urlretrieve(url, output_directory + f"\{file_name6}")
                conv_df6['Image_6_Name'].append(file_name6)
                conv_df6['Sku'].append(mfg)
                conv_df6['Image_6_Failed'].append('NULL')

            except:
                conv_df6['Image_6_Failed'].append(mfg)

        conv_df6 = pd.DataFrame.from_dict(conv_df6).fillna('NULL')

    if any(mfg_list_6):
        dataframe6()
    else:
        global conv_df6
        conv_df6 = pd.DataFrame.from_dict(conv_df6).fillna('NULL')

    file_df = pd.merge(pd.merge(pd.merge(pd.merge(pd.merge(primary_df, conv_df2,
                                                           how='left',
                                                           left_on='Sku',
                                                           right_on='Sku'),
                                                  conv_df3,
                                                  how='left',
                                                  left_on='Sku',
                                                  right_on='Sku'),
                                         conv_df4,
                                         how='left',
                                         left_on='Sku',
                                         right_on='Sku'),
                                conv_df5,
                                how='left',
                                left_on='Sku',
                                right_on='Sku'),
                       conv_df6,
                       how='left',
                       left_on='Sku',
                       right_on='Sku')

    Primary_Failed_col = file_df.pop('Primary_Failed')
    Image_2_Failed_col = file_df.pop('Image_2_Failed')
    Image_3_Failed_col = file_df.pop('Image_3_Failed')
    Image_4_Failed_col = file_df.pop('Image_4_Failed')
    Image_5_Failed_col = file_df.pop('Image_5_Failed')
    Image_6_Failed_col = file_df.pop('Image_6_Failed')

    file_df.insert(7, 'Primary_Failed', Primary_Failed_col)
    file_df.insert(8, 'Image_2_Failed', Image_2_Failed_col)
    file_df.insert(9, 'Image_3_Failed', Image_3_Failed_col)
    file_df.insert(10, 'Image_4_Failed', Image_4_Failed_col)
    file_df.insert(11, 'Image_5_Failed', Image_5_Failed_col)
    file_df.insert(12, 'Image_6_Failed', Image_6_Failed_col)

    file_df.fillna('NULL', inplace=True)

    #  Writing the dataframe to an excel worksheet
    file_df.to_excel('Image_Data.xlsx', sheet_name='File_Name_Data')


def make_main_window():
    # Theme of windows
    sg.theme('Dark Grey 13')

    # Creating window layouts
    main_layout = [[sg.Text("Team Product Tool")],
                   [sg.Text("Choose which tool you want.")],
                   [sg.Button("Build Tool"), sg.Button("Ferg Tool"), sg.Button("Image Converter Tool"),
                    sg.Button("Exit")]]

    return sg.Window('Main Window', main_layout)


def make_build_window():
    # Theme of windows
    sg.theme('Dark Grey 13')

    img_layout = [[sg.Text("Build Scraper Tool")],
                  [sg.Text('Please enter Sku(MFG Number) list.'), sg.InputText(key='-SKU-', pad=(0, 0))],
                  [sg.Text('Please enter Build Unique ID list.'), sg.InputText(key='-UID-', pad=(0, 0))],
                  [sg.Text('Please enter Build Family ID list.'), sg.InputText(key='-FID-', pad=(0, 0))],
                  [sg.Button("Run"), sg.Button("Exit")]]

    image_window = sg.Window('Build Scraper Window', img_layout, modal=True)

    while True:

        event, values = image_window.read()

        if event in (sg.WIN_CLOSED, "Exit"):
            break

        sku_list = values['-SKU-'].split('\n')
        uid_1_list = values['-UID-'].split('\n')
        uid_2_list = values['-FID-'].split('\n')
        # file_name = values['-E_NAME-'].rstrip()

        if event == 'Run':

            try:
                Build_tool(sku_list, uid_1_list, uid_2_list)
                sg.popup("Run Complete!")
            except Exception as e:
                sg.popup("Something went wrong. Please make sure everything was entered correctly.", e)

    image_window.close()


def make_ferg_window():
    # Theme of windows
    sg.theme('Dark Grey 13')

    img_layout = [[sg.Text("Ferg Scraper Tool")],
                  [sg.Text('Please enter Alt 1 list.'), sg.InputText(key='-ALT1-', pad=(0, 0))],
                  [sg.Text('Please enter Uniqueid list.'), sg.InputText(key='-UID-', pad=(0, 0))],
                  [sg.Button("Run"), sg.Button("Exit")]]

    image_window = sg.Window('Ferg Scraper Window', img_layout, modal=True)

    while True:

        event, values = image_window.read()

        if event in (sg.WIN_CLOSED, "Exit"):
            break

        sku_list = values['-SKU-'].split('\n')
        uniqueid_list = values['-UID-'].split('\n')

        if event == 'Run':

            try:
                Ferg_Tool(sku_list, uniqueid_list)
                sg.popup("Run Complete!")
            except Exception as e:
                sg.popup("Something went wrong. Please make sure everything was entered correctly.", e)

    image_window.close()


def make_converter_window():
    # Theme of windows
    sg.theme('Dark Grey 13')
    # WORK HERE
    converter_layout = [[sg.Text("Image Converter Tool")],
                        # [sg.Text("Enter Skus and URLs for Primary images:")],
                        [sg.Text('Please enter Primary Sku(MFG Number) list.'), sg.InputText(key='-SKU-', pad=(0, 0))],
                        [sg.Text('Please enter Primary image URL list.'), sg.InputText(key='-URL-', pad=(0, 0))],
                        # [sg.Text("Enter Skus and URLs for 2nd image:")],
                        [sg.Text('Please enter Sku(MFG Number) list for 2nd image.'),
                         sg.InputText(key='-SKU2-', pad=(0, 0))],
                        [sg.Text('Please enter image URL list for 2nd image.'), sg.InputText(key='-URL2-', pad=(0, 0))],
                        # [sg.Text("Enter Skus and URLs for 3rd image:")],
                        [sg.Text('Please enter Sku(MFG Number) list for 3rd image.'),
                         sg.InputText(key='-SKU3-', pad=(0, 0))],
                        [sg.Text('Please enter image URL list for 3rd image.'), sg.InputText(key='-URL3-', pad=(0, 0))],
                        # [sg.Text("Enter Skus and URLs for 4th image:")],
                        [sg.Text('Please enter Sku(MFG Number) list for 4th image.'),
                         sg.InputText(key='-SKU4-', pad=(0, 0))],
                        [sg.Text('Please enter image URL list for 4th image.'), sg.InputText(key='-URL4-', pad=(0, 0))],
                        [sg.Text('Please enter Sku(MFG Number) list for 5th image.'),
                         sg.InputText(key='-SKU5-', pad=(0, 0))],
                        [sg.Text('Please enter image URL list for 5th image.'), sg.InputText(key='-URL5-', pad=(0, 0))],
                        [sg.Text('Please enter Sku(MFG Number) list for 6th image.'),
                         sg.InputText(key='-SKU6-', pad=(0, 0))],
                        [sg.Text('Please enter image URL list for 6th image.'), sg.InputText(key='-URL6-', pad=(0, 0))],
                        # [sg.Text('Please enter the absolute path of Excel file to use.'), sg.InputText(key='-E_NAME-')],
                        [sg.Text('Please enter the absolute path of folder to download images to.'),
                         sg.InputText(key='-F_NAME-')],
                        [sg.Button("Run"), sg.Button("Exit")]]

    convert_window = sg.Window('Image Converter Window', converter_layout, modal=True)

    while True:

        event, values = convert_window.read()

        if event in (sg.WIN_CLOSED, "Exit"):
            break

        mfg_list_primary = values['-SKU-'].split('\n')
        mfg_list_2 = values['-SKU2-'].split('\n')
        mfg_list_3 = values['-SKU3-'].split('\n')
        mfg_list_4 = values['-SKU4-'].split('\n')
        mfg_list_5 = values['-SKU5-'].split('\n')
        mfg_list_6 = values['-SKU6-'].split('\n')
        Primary_list = values['-URL-'].split('\n')
        img_2_url_list = values['-URL2-'].split('\n')
        img_3_url_list = values['-URL3-'].split('\n')
        img_4_url_list = values['-URL4-'].split('\n')
        img_5_url_list = values['-URL5-'].split('\n')
        img_6_url_list = values['-URL6-'].split('\n')
        # excel_file_name = r'{}'.format(values['-E_NAME-'].rstrip())
        folder_name = r'{}'.format(values['-F_NAME-'].rstrip())

        if event == 'Run':

            try:
                # converter_tool(mfg_list_primary, mfg_list_2, mfg_list_3, mfg_list_4, Primary_list, img_2_list, img_3_list, img_4_list, folder_name)
                converter_tool(mfg_list_primary,
                               mfg_list_2,
                               mfg_list_3,
                               mfg_list_4,
                               mfg_list_5,
                               mfg_list_6,
                               Primary_list,
                               img_2_url_list,
                               img_3_url_list,
                               img_4_url_list,
                               img_5_url_list,
                               img_6_url_list,
                               folder_name)

                sg.popup("Run Complete!")
            except Exception as e:
                sg.popup("Something went wrong. Please make sure everything was entered correctly.", e)

    convert_window.close()


def main():
    # Theme of windows
    sg.theme('Dark Grey 13')

    # Creating window layouts
    main_layout = [[sg.Text("Team Product Tool")],
                   [sg.Text("Make sure you are on the VPN!", text_color='red', font=('Arial Bold', 10))],
                   [sg.Text("Choose which tool you want.")],
                   [sg.Button("Build Tool"), sg.Button("Ferg Tool"), sg.Button("Image Converter Tool"),
                    sg.Button("Exit")]]

    main_window = sg.Window('Main Window', main_layout)

    # Event Loop
    while True:
        event, values = main_window.read()

        # End program if conditions met
        if event in (sg.WIN_CLOSED, "Exit"):
            break

        # Runs the Image scraper tool window and tool
        elif event == 'Build Tool':
            make_build_window()

        # Runs the Bullet scraper tool window and tool
        elif event == 'Ferg Tool':
            make_ferg_window()

        # Runs the Image URL Converter tool window and tool
        elif event == 'Image Converter Tool':
            make_converter_window()

    main_window.close()


# Run the program
if __name__ == "__main__":
    main()
