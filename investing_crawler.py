#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   douban_crawl.py
@Time    :   2021/10/11 13:55:59
@Author  :   Pan Yifan
@Version :   1.0
@Desc    :   None
1. 
'''

#%% 0. 工具库
import datetime
import gc
import time
import pandas as pd
from dateutil import parser
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import urllib

#%% 1. 参数设定
# 1.1. 网址设置
url_investing = 'https://www.investing.com/news/stock-market-news/'

# 1.2. 时间
today = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
today_dt = datetime.datetime.now()


#%% 2. 函数设置
def judge_alpha(string):
    pool = list(range(len(string)))[::-1]
    count = pool[0]
    for i in pool:
        if string[i].isalpha():
            count = i
            break
    return string[:count] + string[count]


#%% 3.
title = []
date = []
chrome_options = Options()
chrome_options.add_argument('--disable-dev-shm-usage')
chrome_options.add_argument('--no-sandbox')
driver = webdriver.Chrome('C:\\chromedriver.exe', options=chrome_options)
for i in range(8885, 11000):
    df = {'title': title, 'date': date}
    df = pd.DataFrame(df)
    df.to_excel('D:\\{}.xlsx'.format(today))
    print('------Saved-----------------')
    print(i + 1)
    if (i + 1) % 20 == 0:
        driver.implicitly_wait(3)
        driver.delete_all_cookies()
        driver.close()
        time.sleep(2)
        try:
            driver = webdriver.Chrome('C:\\chromedriver.exe', options=chrome_options)
        except:
            driver = webdriver.Chrome('C:\\chromedriver.exe', options=chrome_options)
    driver.implicitly_wait(3)
    try:
        driver.get(url_investing + str(i + 1))
        driver.implicitly_wait(5)
        for m in driver.find_elements_by_css_selector('.largeTitle'):
            for j in m.find_elements_by_css_selector('.textDiv > a'):
                title.append(j.get_attribute('title'))
            for j in m.find_elements_by_css_selector('.textDiv'):
                for k in j.find_elements_by_css_selector('.articleDetails'):
                    title_time = k.text.split(' - ')[-1][:12]
                    if 'minute' in title_time:
                        time_num = title_time.split(' ')[0]
                        date.append(today_dt - datetime.timedelta(minutes=int(time_num)))
                    elif 'hour' in title_time:
                        time_num = title_time.split(' ')[0]
                        date.append(today_dt - datetime.timedelta(hours=int(time_num)))
                    elif title_time == 'Just Now':
                        date.append(today_dt)
                    else:
                        date.append(parser.parse(title_time))
            del j, k, title_time
            break
        del m
        print(len(title) == len(date), len(title), len(date))
        gc.collect()
    except:
        time.sleep(2)
        driver = webdriver.Chrome('C:\\chromedriver.exe', options=chrome_options)
        time.sleep(2)
        driver.get(url_investing + str(i + 1))
        time.sleep(2)
        for i in driver.find_elements_by_css_selector('.largeTitle'):
            for j in i.find_elements_by_css_selector('.textDiv > a'):
                title.append(j.get_attribute('title'))
            for j in i.find_elements_by_css_selector('.textDiv'):
                for k in j.find_elements_by_css_selector('.articleDetails'):
                    title_time = k.text.split(' - ')[-1][:12]
                    if 'minute' in title_time:
                        time_num = title_time.split(' ')[0]
                        date.append(today_dt - datetime.timedelta(minutes=int(time_num)))
                    elif 'hour' in title_time:
                        time_num = title_time.split(' ')[0]
                        date.append(today_dt - datetime.timedelta(hours=int(time_num)))
                    elif title_time == 'Just Now':
                        date.append(today_dt)
                    else:
                        date.append(parser.parse(title_time))
            del j, k, title_time
            break
        del i
        print(len(title) == len(date), len(title), len(date))
        gc.collect()
driver.close()
df = {'title': title, 'date': date}
df = pd.DataFrame(df)
df.to_excel('D:\\{}.xlsx'.format(today))