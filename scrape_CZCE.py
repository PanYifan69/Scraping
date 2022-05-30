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
wait_time = 3
# driver.refresh()
time.sleep(wait_time)
holidays = pd.read_excel(os.path.dirname(__file__) + '\\Holidays.xlsx')
holidays = holidays.values[1:]

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


starting_date = "2021.1.4"
ending_date = "2021.1.4"
writer=pd.ExcelWriter(os.path.dirname(__file__) + '\\' + "Daily CZCE Top Brokers for product " + commodity + " " + starting_date + \
                    " - " + ending_date + ".xlsx") #enter file name for different product
actual_date = starting_date
act_year = int(actual_date.split(".")[0])
act_month = int(actual_date.split(".")[1])
act_day = int(actual_date.split(".")[2])
date = datetime(act_year, act_month, act_day)
ending_date_year = int(ending_date.split(".")[0])
ending_date_month = int(ending_date.split(".")[1])
ending_date_day = int(ending_date.split(".")[2])
ending_datetime = datetime(ending_date_year, ending_date_month, ending_date_day)
brokers = []
while (date <= ending_datetime):
    if (not date.weekday() + 1 in range(1, 6)) or (date in holidays):
        date = date + timedelta(days=1)
        continue
    print(date)
    # print("Starting date must not be a holiday or weekend")
    starting_date = date.strftime("%Y.%m.%d")
    year = starting_date.split(".")[0]
    month = str(int(starting_date.split(".")[1]))
    day = str(int(starting_date.split(".")[2]))
    starting_day = int(day) - 1
    last_available_date = time.time()
    find_right_year(year)
    time.sleep(wait_time / 2)
    all_brokers = []

    # if len(select_date) == 1:
    while (1 == 1):
        try:
            find_right_year(year)
            time.sleep(wait_time / 2)
            all_brokers = []
            all_months = driver.find_element_by_xpath('//*[@id="rendez-vous-open"]/div[2]/div/section/div')
            all_months = all_months.find_elements_by_tag_name("button")
            all_months[int(month) - 1].send_keys(Keys.ENTER)
            time.sleep(2)
            # print(datetime(act_year, act_month, act_day))
            # if time.time() - last_available_date > 60:
            #     break
            all_days = driver.find_element_by_xpath('//*[@id="rendez-vous-open"]/div[2]/div/section/div')
            all_days = all_days.find_elements_by_tag_name("button")
            all_days[int(day) - 1].send_keys(Keys.ENTER)
            time.sleep(2)
            select_comd = Select(driver.find_element_by_xpath('//*[@id="formrl"]/div[2]/select'))
            select_comd.select_by_value(commodity)
            time.sleep(2)
            select_date = driver.find_elements_by_class_name("seach")
            select_date[0].send_keys(Keys.ENTER)
            time.sleep(2)
            temp = driver.find_element_by_xpath('/html/body/div/div/div[1]/span[1]/a')
            break
        except:
            driver.refresh()
            time.sleep(2)
            pass
    right_div = driver.find_element_by_id("datebox2")
    iframe = right_div.find_element_by_tag_name("iframe")
    driver.switch_to.frame(iframe)
    tables = driver.find_elements_by_tag_name("table")
    # print(len(tables))
    done = False
    for k in range(len(tables)):
        # print(k)
        print(len(tables[k].text))
        if len(tables[k].text) > 0:
            table_list = tables[k].text.split('\n')
            positions_long = {}
            positions_short = {}
            volumes = {}
            contracts_volume = {}
            commodity_indices = [0, 0]
            for i in range(len(table_list)):
                if commodity_indices[-1] - commodity_indices[-2] > 35:
                    if len(commodity_indices) > 4:
                        break
                if table_list[i].find(commodity) != -1:
                    commodity_indices.append(i)
                    if table_list[i].find('品种') != -1:
                        actual_date = table_list[i].split("：")[2]
                        print(actual_date)
                        last_available_date = time.time()
                        act_year = int(actual_date.split("-")[0])
                        act_month = int(actual_date.split("-")[1])
                        act_day = int(actual_date.split("-")[2])
                        for j in range(i + 2, 300):
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
                            all_brokers.append(broker_long)
                            all_brokers.append(broker_short)
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
                        done = True
                        contracts_volume[contract] = volume_sum_contract

            if done == True:
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
                # print(positions_net)
                positions_net.to_excel(writer, sheet_name=actual_date.replace("-", "."))
                brokers.append(pd.DataFrame(positions_net.index))

        if done == True:
            break
    driver.switch_to.default_content()

    # else:
    #     print("no select date button")
    date_navigator_right = driver.find_element_by_xpath('//*[@id="rendez-vous-open"]/div[2]/div/nav/button[3]')
    date_navigator_right.send_keys(Keys.ENTER)
    time.sleep(1)
    starting_day = 0
    date = date + timedelta(days=1)
brokers = pd.concat(brokers, axis=0).drop_duplicates().reset_index(drop=True)
brokers.columns = ['broker']
# brokers = pd.DataFrame(brokers)
# print(brokers)
brokers.to_excel(writer, sheet_name='all brokers')
writer.save()
