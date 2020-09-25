from selenium import webdriver
import time
import pyautogui
from pywinauto.keyboard import send_keys
import pyperclip
from tqdm import trange, tqdm, tqdm_notebook
import pandas as pd
from datetime import datetime
import csv
import numpy as np
import os
import socket
import random


def copy_input(xpath, input):
    pyperclip.copy(input)
    driver.find_element_by_xpath(xpath).click()
    ActionChains(driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
    time.sleep(1)

df = pd.read_excel('./coupang_input.xlsx', header=None, sheet_name = 'Sheet1',nrows = 10)
dg = pd.read_excel('./coupang_keyword.xlsx', header=None, sheet_name = 'Sheet1',nrows = 10)


keyword_df = pd.read_excel('./coupang_keyword.xlsx', sheet_name = 'Sheet1')
prodnum_df = pd.read_excel('./coupang_input.xlsx', sheet_name = 'Sheet1')

log_df = pd.DataFrame(columns=['작업시간','작업키워드','상품번호','페이지','순위','아이피주소'])

index = 0
while True:
    group_list = list(set(keyword_df.그룹))
    for group in group_list:
        keyword_list = list(keyword_df[keyword_df['그룹'] == group]['키워드'])
        item_name = str(prodnum_df[prodnum_df['그룹'] == group]['상품 번호'].values[0])
        traffic_name = prodnum_df[prodnum_df['그룹'] == group]['반복횟수'].values[0]
        
        for i in range(traffic_name):
            keyword_name = random.choice(keyword_list)
            time.sleep(3)
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_argument("--incognito")
            driver = webdriver.Chrome(chrome_options=chrome_options)
            driver.implicitly_wait(5)
            driver.get("https://m.coupang.com/nm/")
            time.sleep(1)

            try:
                driver.find_element_by_css_selector('#fullBanner > div > div > a.close-banner').click() 
            except:
                pass

            time.sleep(1)
            driver.find_element_by_css_selector('#searchBtn').click()
            time.sleep(1)
            driver.find_element_by_css_selector('#q').send_keys(keyword_name + '\n')
            time.sleep(1)

            break_check = False
            item_list = driver.find_elements_by_css_selector('#productList > li')
            for rank, item in enumerate(item_list):
                if -1 != item.find_element_by_css_selector('a').get_attribute('href').find(item_name):
                    page = driver.find_element_by_css_selector('span.page.selected').text
                    time.sleep(2)
                    item.find_element_by_css_selector('a').click()
                    break_check = True
                    break
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            if break_check == False:
                driver.find_elements_by_css_selector('div.page-wrapper > span')[2].click()
                item_list = driver.find_elements_by_css_selector('#productList > li')
                for rank, item in enumerate(item_list):
                    if -1 != item.find_element_by_css_selector('a').get_attribute('href').find(item_name):
                        page = driver.find_element_by_css_selector('span.page.selected').text
                        time.sleep(2)
                        item.find_element_by_css_selector('a').click()
                        break_check = True
                        break
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            if break_check == False:
                driver.find_elements_by_css_selector('div.page-wrapper > span')[3].click()
                item_list = driver.find_elements_by_css_selector('#productList > li')
                for rank, item in enumerate(item_list):
                    if -1 != item.find_element_by_css_selector('a').get_attribute('href').find(item_name):
                        page = driver.find_element_by_css_selector('span.page.selected').text
                        time.sleep(2)
                        item.find_element_by_css_selector('a').click()
                        break_check = True
                        break
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(2)
            if break_check == False:
                while True:
                    try:
                        driver.find_elements_by_css_selector('div.page-wrapper > span')[4].click()
                        item_list = driver.find_elements_by_css_selector('#productList > li')
                        for rank, item in enumerate(item_list):
                            if -1 != item.find_element_by_css_selector('a').get_attribute('href').find(item_name):
                                page = driver.find_element_by_css_selector('span.page.selected').text
                                time.sleep(2)
                                item.find_element_by_css_selector('a').click()
                                break_check = True
                                break
                        if break_check:
                            break
                        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
                        time.sleep(2)
                    except:
                        break
            if break_check == False:
                page = '검색되지 않습니다.'
                rank = '검색되지 않습니다.'
            temp_list = []
            temp_list.append(str(datetime.now())[:19])
            temp_list.append(keyword_name)
            temp_list.append(item_name)
            temp_list.append(page)
            try:
                temp_list.append(rank+1)
            except:
                temp_list.append(rank)
            temp_list.append(socket.gethostbyname(socket.getfqdn()))

            log_df.loc[index] = temp_list
            index += 1
            log_df.to_csv('coupang_result.csv',index=False,encoding='utf-8-sig')
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(9)
            driver.quit()
            time.sleep(2)
            pyautogui.hotkey('Alt','P')
            pyautogui.hotkey('Alt','P')
            time.sleep(2)
            