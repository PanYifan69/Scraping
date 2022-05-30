import os
import time
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select

from selenium.webdriver.firefox.options import Options as FirefoxOptions

firefox_options = FirefoxOptions()
# firefox_options.add_argument("--headless")
driver = webdriver.Firefox(executable_path="C:\\geckodriver.exe", options=firefox_options)
url = "http://www.czce.com.cn/cn/jysj/ccpm/H770304index_1.htm"
driver.get(url)

chrome_options = Options()
#chrome_options.add_argument("--headless") # does it open web browser or not
#chrome_options.binary_location = '/usr/bin/chromium-browser'

# url = "http://www.czce.com.cn/cn/jysj/ccpm/H770304index_1.htm"
# driver = webdriver.Chrome(os.getcwd() +"\\chromedriver.exe",   options = firefox_options)
# driver.get(url)

#scripts = driver.find_elements_by_tag_name('script')
# iframe = driver.find_element_by_xpath('//*[@id="senfe16"]')
# driver.switch_to.frame(iframe)
# div = driver.find_elements_by_tag_name('div')

commodity = "FG"


# if True, we are in the right year and just need to select the month
def find_right_year(year):
    for k in range(3):
        for kk in range(10):
            time.sleep(4)
            try:
                date_navigator = driver.find_element_by_xpath('//*[@id="rendez-vous-open"]/div[2]/div/nav/button[2]')
                break
            except:
                pass
        if date_navigator.text.find("-") != -1:
            if int(year) < 2020:
                date_navigator_left = driver.find_element_by_xpath('//*[@id="rendez-vous-open"]/div[2]/div/nav/button[1]')
                date_navigator_left.click()
                all_years = driver.find_element_by_xpath('//*[@id="rendez-vous-open"]/div[2]/div/section/div')
                all_years = all_years.find_elements_by_tag_name("button")
                all_years[int(year) - 2010].send_keys(Keys.ENTER)
            elif int(year) >= 2020:
                all_years = driver.find_element_by_xpath('//*[@id="rendez-vous-open"]/div[2]/div/section/div')
                all_years = all_years.find_elements_by_tag_name("button")
                all_years[int(year) - 2020].send_keys(Keys.ENTER)
            break

        date_navigator.click()  #send_keys(Keys.ENTER)


def scrape_daily(starting_date):
    #starting_date = "2020.03.03"
    year = starting_date.split(".")[0]
    month = str(int(starting_date.split(".")[1]))
    day = str(int(starting_date.split(".")[2]))
    find_right_year(year)

    all_months = driver.find_element_by_xpath('//*[@id="rendez-vous-open"]/div[2]/div/section/div')
    all_months = all_months.find_elements_by_tag_name("button")
    all_months[int(month) - 1].send_keys(Keys.ENTER)
    time.sleep(4)
    all_days = driver.find_element_by_xpath('//*[@id="rendez-vous-open"]/div[2]/div/section/div')
    all_days = all_days.find_elements_by_tag_name("button")
    all_days[int(day) - 1].send_keys(Keys.ENTER)
    select_date = driver.find_elements_by_class_name("seach")
    if len(select_date) == 1:
        select_date[0].send_keys(Keys.ENTER)
        time.sleep(4)
        right_div = driver.find_element_by_id("datebox2")
        iframe = right_div.find_element_by_tag_name("iframe")
        driver.switch_to.frame(iframe)
        tables = driver.find_elements_by_tag_name("table")
        done = False
        for k in range(len(tables)):
            # if we the right commodity, then we have to stop after because we are only interested in the summary
            if len(tables[k].text) > 0:
                table_list = tables[k].text.split('\n')
                positions_long = {}
                positions_short = {}
                volumes = {}
                contracts_volume = {}
                for i in range(len(table_list)):
                    if table_list[i].find(commodity) != -1:
                        if table_list[i].find('品种') != -1:
                            for j in range(i + 2, min(len(table_list), i + 300)):
                                tmp = table_list[j].split(" ")
                                if tmp[0] == '合计':
                                    break
                                broker_volume = tmp[1]
                                if tmp[2] == "-":
                                    volume = 0
                                else:
                                    volume = int(tmp[2].replace(",", ""))
                                volumes[broker_volume] = volume
                                broker_long = tmp[4]
                                long = int(tmp[5].replace(",", ""))
                                positions_long[broker_long] = long
                                broker_short = tmp[7]
                                short = int(tmp[8].replace(",", ""))
                                positions_short[broker_short] = short

                                positions_net = {}

                                for k, v in positions_long.items():
                                    if k not in positions_net.keys():
                                        positions_net[k] = v
                                    else:
                                        positions_net[k] += v

                                for k, v in positions_short.items():
                                    if k not in positions_net.keys():
                                        positions_net[k] = -v
                                    else:
                                        positions_net[k] -= v

                                positions_net = pd.DataFrame.from_dict(positions_net, orient="index")

                        else:
                            volume_sum_contract = 0
                            contract = table_list[i].split("：")[1].split(" ")[0]
                            for j in range(i + 2, min(len(table_list), i + 300)):
                                tmp = table_list[j].split(" ")
                                if tmp[0] == '合计':
                                    break
                                if tmp[2] == "-":
                                    volume = 0
                                else:
                                    volume = int(tmp[2].replace(",", ""))
                                volume_sum_contract += volume
                            contracts_volume[contract] = volume_sum_contract

                positions_net = positions_net.sort_values(by=0, ascending=False)
                sorted_volumes = []
                for broker in positions_net.index:
                    try:
                        sorted_volumes.append(volumes[broker])
                    except:
                        sorted_volumes.append(0)
                positions_net["Volume"] = sorted_volumes
                positions_net.columns = ["Net position", "Volume"]
                contracts_volume = pd.DataFrame.from_dict(contracts_volume, orient="index")
                contracts_volume = contracts_volume.sort_values(by=0, ascending=False)
                contracts_ = list(contracts_volume.index) + [np.nan for ii in range(len(positions_net) - len(contracts_volume))]
                contracts_volume_ = list(contracts_volume.iloc[:, 0]) + [np.nan for ii in range(len(positions_net) - len(contracts_volume))]
                positions_net["Contracts"] = contracts_
                positions_net["Contract volume"] = contracts_volume_
                positions_net.to_excel(os.path.dirname(__file__) + '\\' + commodity + "_" + starting_date + ".xlsx", sheet_name=starting_date)
                done = True

            if done == True:
                break

    else:
        print("no select date button")

    driver.close()


# Scrape today
# today = datetime.today()
# today = str(today.year) + "." + str(today.month) + "." + str(today.day)
# scrape_daily(today)

#
# Scrape random date - uncomment line below and set the date you want
scrape_daily("2022.5.5")
# scrape_daily("2021.01.21")
# scrape_daily("2021.01.22")
# scrape_daily("2021.01.25")
# scrape_daily("2021.01.29")
# scrape_daily("2015.01.16")
# scrape_daily("2017.06.06")