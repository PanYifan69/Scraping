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
from zipfile import ZipFile

#### WEBSCRAPING PART

commodity = 'j'
max_num = 'all'
download_directory = os.path.dirname(__file__) + '\\temp\\'


def string_f(s):
    if s < 10:
        return "0" + str(s)
    else:
        return str(s)


def scrape_day(commodity, year_var, month_var, day_var, max_num):
    chrome_options = Options()
    # chrome_options.add_argument("--headless")
    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': download_directory}
    chrome_options.add_experimental_option('prefs', prefs)
    driver = webdriver.Chrome("C:\\chromedriver.exe", chrome_options=chrome_options)
    url = "http://www.dce.com.cn/dalianshangpin/xqsj/tjsj26/rtj/rcjccpm/index.html"
    driver.get(url)

    # chrome_options = Options()
    # #chrome_options.add_argument("--headless") # does it open web browser or not
    # #chrome_options.binary_location = '/usr/bin/chromium-browser'

    # url = "http://www.dce.com.cn/dalianshangpin/xqsj/tjsj26/rtj/rcjccpm/index.html"
    # driver = webdriver.Chrome(os.getcwd() +"\\chromedriver.exe",   chrome_options=chrome_options)
    # driver.get(url)

    # odpri z f12 inspect. potem pa v iframe je neka nova html stran.. copy c path in dobiš ta iframe find element by xpath.
    # Odpri z copy by xpth - treba se v inspectu pomikat po značkah

    # Switch to inner frame
    for ii in range(10):
        time.sleep(10)
        print(ii)
        try:
            iframe = driver.find_element_by_xpath("//*[@id='12651']/div[2]/div/iframe")
            driver.switch_to.frame(iframe)  # tu se premakneš v nov iframe v katerem so te tabele
            break
        except:
            pass

    # all_tables = {}
    days_dict = {}

    start = time.time()

    wait = 1

    kok_praznih = 0
    we_are_in_future = False
    for year in range(year_var, year_var + 1):
        # Select Year
        selectYear = Select(
            driver.find_element_by_xpath("/html/body//*[@id='memberDealPosiQuotesForm']//*[@id='calender']//*[@id='control']/select[1]"))
        selectYear.select_by_value(str(year))
        for month in range(month_var - 1, month_var):
            if we_are_in_future == True:
                break
            # Select Month (Starting from 0, Jan == 0)

            max_row = 8
            if month == 1 and year == 2015:  # february has only 4 rows not 5 rows
                max_row = 7
            if day_var == 31 and datetime(year_var, month_var, day_var).weekday() == 0:  # If the 31st is a monday there are 9 rows
                max_row = 9
            for row in range(3, max_row):
                # day is saved in a matrix and first week starts in third row
                if we_are_in_future == True:
                    break
                for col in range(2, 7):
                    selectYear = Select(
                        driver.find_element_by_xpath("/html/body//*[@id='memberDealPosiQuotesForm']//*[@id='calender']//*[@id='control']/select[1]"))
                    selectYear.select_by_value(str(year))
                    selectMonth = Select(
                        driver.find_element_by_xpath("/html/body//*[@id='memberDealPosiQuotesForm']//*[@id='calender']//*[@id='control']/select[2]"))
                    selectMonth.select_by_value(str(month))
                    selectDay = driver.find_element_by_xpath("/html/body//*[@id='memberDealPosiQuotesForm']//*[@id='calender']/table/tbody/tr[" +
                                                             str(row) + "]/td[" + str(col) + "]")
                    day = selectDay.text
                    if day != "":
                        if int(day) != day_var:
                            continue
                        date = datetime(year, month + 1, int(day))
                        print(date)
                        if datetime.today() < date:
                            we_are_in_future = True
                            break
                        cont_dict = {}
                        selectDay.click()
                        print("Wait for site to load")
                        time.sleep(5)

                        ## This while loop is to wait for site to load
                        while 5 < 6:
                            try:
                                driver.execute_script("setVariety('{}')".format(commodity))
                                break
                            except:
                                print("not yet loaded")
                                pass

                        print("Wait for site to load")
                        time.sleep(5)

                        for pp in range(10):
                            try:
                                selectContracts = driver.find_element_by_xpath("//*[@id='memberDealPosiQuotesForm']/div/div[1]/div[4]/div/ul")
                                break
                            except:
                                time.sleep(3)
                                print("not yet loaded 2")
                                pass
                        try:
                            href = driver.find_element_by_xpath('//*[@id="memberDealPosiQuotesForm"]/div/div[2]/ul[1]/li[4]/a')
                            href.click()
                            driver.switch_to.alert.accept()
                        except:
                            continue
                        download_done = False
                        number_of_tries = 1
                        while download_done == False:
                            time.sleep(5)
                            try:
                                number_of_tries += 1
                                file_name = str(year_var) + string_f(int(month_var) + 1) + string_f(int(day_var)) + "_DCE_DPL.zip"
                                with ZipFile(download_directory + file_name) as zfile:
                                    download_done = True
                            except:
                                if number_of_tries > 5:
                                    break
                                continue
    #                     while 5 < 6:
    #                         try:
    #                             selectContracts = driver.find_element_by_xpath("//*[@id='memberDealPosiQuotesForm']/div/div[1]/div[4]/div/ul")
    #                             break
    #                         except:
    #                             print("not yet loaded")
    #                             pass

    #                     for contract in selectContracts.text.split():
    #                         brokers_dict = {}
    #                         while 5 < 6:
    #                             try:
    #                                 driver.execute_script("setContract_id('" + contract + "')")
    #                                 break
    #                             except:
    #                                 print("not yet loaded")
    #                                 pass

    #                         print("using contract: ", contract)

    #                         print("Wait for site to load")
    #                         time.sleep(15)
    #                         try:
    #                             table = driver.find_element_by_xpath(
    #                                 "/html/body//*[@id='memberDealPosiQuotesForm']/div/div[2]//*[@id='printData']/div/table[2]")
    #                         except:
    #                             print("no table")
    #                             continue

    #                         # all_tables[(date,contract)] = table.text

    #                         lines = table.text.split('\n')

    #                         for i in range(1, len(lines) - 1):
    #                             line = lines[i]
    #                             val = line.split()
    #                             if len(val) != 12:
    #                                 kok_praznih += 1
    #                                 continue
    #                             # print(val)
    #                             if (val[1] in brokers_dict.keys()):
    #                                 position_dict = brokers_dict[val[1]]
    #                             else:
    #                                 position_dict = {}
    #                             position_dict["volume"] = float(val[2].replace(",", ""))
    #                             brokers_dict[val[1]] = position_dict.copy()

    #                             if (val[5] in brokers_dict.keys()):
    #                                 position_dict = brokers_dict[val[5]]
    #                             else:
    #                                 position_dict = {}
    #                             position_dict["long"] = float(val[6].replace(",", ""))
    #                             brokers_dict[val[5]] = position_dict.copy()

    #                             if (val[9] in brokers_dict.keys()):
    #                                 position_dict = brokers_dict[val[9]]
    #                             else:
    #                                 position_dict = {}
    #                             position_dict["short"] = float(val[10].replace(",", ""))
    #                             brokers_dict[val[9]] = position_dict.copy()
    #                         cont_dict[contract] = brokers_dict.copy()
    #                     days_dict[date] = cont_dict.copy()
    #                 else:
    #                     print("empty day")

    # print("It took: ", (time.time() - start) / 60, " minutes")
    # print(kok_praznih)

    # driver.close()

    # ## The data scraped has the following dictionarial structure: dict_days[datetime]:
    # ##                                                                    cont_dict[i1901]:
    # ##                                                                       brokers_dict[ijangwo]:
    # ##                                                                            position_dict[short]: 100

    # #### Put data together ###

    # daily_lists = {}
    # all_brokers = set()

    # for i in range(len(list(days_dict.keys()))):
    #     date = list(days_dict.keys())[i]
    #     contracts_on_date = list(days_dict[date].keys())  # a list of all active contracts on that date
    #     net_positions_on_date = {}  # a dict with brokers as keys - used to net up all long and short positions over all contracts on this day
    #     volumes_on_date = {}  # volume on that day
    #     volume_on_contracts = {}  # volume in each contract
    #     for contract in contracts_on_date:
    #         top_brokers_on_contract = days_dict[date][contract].keys()
    #         for broker in top_brokers_on_contract:
    #             all_brokers.add(broker)
    #             if 'short' in days_dict[date][contract][broker].keys():
    #                 # try If the dictionary has key for this broker , else do except
    #                 try:
    #                     net_positions_on_date[broker] -= days_dict[date][contract][broker]['short']
    #                 except:
    #                     net_positions_on_date[broker] = -days_dict[date][contract][broker]['short']
    #             if 'long' in days_dict[date][contract][broker].keys():
    #                 # try If the dictionary has key for this broker , else do except
    #                 try:
    #                     net_positions_on_date[broker] += days_dict[date][contract][broker]['long']
    #                 except:
    #                     net_positions_on_date[broker] = days_dict[date][contract][broker]['long']
    #             if 'volume' in days_dict[date][contract][broker].keys():
    #                 # try If the dictionary has key for this broker , else do except
    #                 try:
    #                     volumes_on_date[broker] += days_dict[date][contract][broker]['volume']
    #                 except:
    #                     volumes_on_date[broker] = days_dict[date][contract][broker]['volume']
    #                     # try if this contract is already in dict else do except
    #                 try:
    #                     volume_on_contracts[contract] += days_dict[date][contract][broker]['volume']
    #                 except:
    #                     volume_on_contracts[contract] = days_dict[date][contract][broker]['volume']

    #     # add volumes for each broker that has information about that
    #     for broker in list(net_positions_on_date.keys()):
    #         try:
    #             net_positions_on_date[broker] = [net_positions_on_date[broker], volumes_on_date[broker]]
    #         except:
    #             net_positions_on_date[broker] = [net_positions_on_date[broker], 0]

    #     df = pd.DataFrame.from_dict(net_positions_on_date, orient="index", columns=["net_position", "volume"])
    #     df = df.sort_values(by=["net_position"], ascending=False)

    #     # Find dominant contract
    #     contracts_by_volume = pd.DataFrame.from_dict(volume_on_contracts, orient="index", columns=["volume"])
    #     contracts_by_volume = contracts_by_volume.sort_values(by=["volume"], ascending=False)
    #     dominant_contracts = np.array(["     " for ii in range(len(net_positions_on_date))])
    #     dominant_contracts[0:len(list(contracts_by_volume.index))] = list(contracts_by_volume.index)
    #     df["dominant contract by volume"] = dominant_contracts

    #     daily_lists[date] = df
    # ## Write to excel
    # writer = pd.ExcelWriter(
    #     os.path.abspath(os.path.dirname(__file__)) + "\\Top Brokers on " + str(year_var) + "." + str(month_var) + "." + str(day_var) + "-" +
    #     commodity + '-' + str(max_num) + ".xlsx")
    # # pd.Series(list(all_brokers),name="broker").to_excel(writer,sheet_name="all brokers")
    # for day, daily_list in daily_lists.items():
    #     daily_list.to_excel(writer, sheet_name=str(day)[0:10].replace("-", "."))
    # writer.save()


#
today = datetime.today()

# scrape_day(today.year,today.month,today.day)
if __name__ == "__main__":
    scrape_day(commodity, 2012, 2, 28, max_num)
# scrape_day(2021, 1, 4)
# scrape_day(2020, 8, 31)
# scrape_day(2020, 5, 12)
