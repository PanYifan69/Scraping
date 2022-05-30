import os
import time
from datetime import datetime, timedelta
from matplotlib.pyplot import waitforbuttonpress
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from zipfile import ZipFile
import io

#### WEBSCRAPING PART

all_top_brokers = True  # If false, then just top 20 brokers will be taken into account, otherwise all of them
download_directory = os.path.dirname(__file__) + '\\temp\\'  # Change it accordingly to where the files are being downloaded from your browser
commodity = "j"
max_num = 1
end_year = 2012


def is_integer(s):
    try:
        int(s)
        return True
    except ValueError:
        return False


def string_f(s):
    if s < 10:
        return "0" + str(s)
    else:
        return str(s)


def scrape_DCE(commodity, max_num, end_year):
    wait = 4
    chrome_options = Options()
    # chrome_options.add_argument("--headless") # does it open web browser or not
    prefs = {'profile.default_content_settings.popups': 0, 'download.default_directory': download_directory}
    chrome_options.add_experimental_option('prefs', prefs)
    # chrome_options.binary_location = '/usr/bin/chromium-browser'

    url = "http://www.dce.com.cn/dalianshangpin/xqsj/tjsj26/rtj/rcjccpm/index.html"
    driver = webdriver.Chrome("C:\\chromedriver.exe", options=chrome_options)
    driver.get(url)
    # time.sleep(wait)

    # Switch to inner frame
    iframe = driver.find_element_by_xpath("//*[@id='12651']/div[2]/div/iframe")
    driver.switch_to.frame(iframe)

    # all_tables = {}

    days_dict = {}

    start = time.time()

    kok_praznih = 0
    we_are_in_future = False
    for year in range(2012, end_year + 1):
        # Select Year
        for month in range(0, 1):
            if we_are_in_future == True:
                break
            # Select Month (Starting from 0, Jan == 0
            max_row = 8
            if month == 1 and year == 2015:  # february has only 4 rows not 5 rows
                max_row = 7
            for row in range(3, 4):  # )max_row): # day is saved in a matrix and first week starts in third row
                if we_are_in_future == True:
                    break
                for col in range(2, 6):  # 7):
                    # driver.refresh()
                    time.sleep(wait)
                    print(row, col, year, month)
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
                        date = datetime(year, month + 1, int(day))
                        print(date)
                        if datetime.today() < date:
                            we_are_in_future = True
                            break
                        cont_dict = {}
                        selectDay.click()
                        ##print("Wait for site to load")
                        time.sleep(wait)  # wait time for load can change this

                        ## This while loop is to wait for site to load
                        for pp in range(10):
                            try:
                                driver.execute_script("setVariety('{}')".format(commodity))
                                break
                            except:
                                time.sleep(3)
                                print("not yet loaded 1")
                                pass

                        print("Wait for site to load 1")
                        time.sleep(wait)
                        for pp in range(10):
                            try:
                                selectContracts = driver.find_element_by_xpath("//*[@id='memberDealPosiQuotesForm']/div/div[1]/div[4]/div/ul")
                                break
                            except:
                                time.sleep(3)
                                print("not yet loaded 2")
                                pass
                        time.sleep(0.5)
                        try:
                            href = driver.find_element_by_xpath('//*[@id="memberDealPosiQuotesForm"]/div/div[2]/ul[1]/li[4]/a')
                            href.click()
                            driver.switch_to.alert.accept()
                        except:
                            continue
                        download_done = False
                        number_of_tries = 1
                        while download_done == False:
                            time.sleep(3)
                            number_of_tries += 1
                            try:
                                encoding = 'utf-8'
                                file_name = str(year) + string_f(int(month) + 1) + string_f(int(day)) + "_DCE_DPL.zip"
                                with ZipFile(download_directory + file_name) as zfile:
                                    download_done = True
                                    cnt_volume = []
                                    if commodity != 'all':
                                        for name in zfile.namelist():
                                            # print(name)
                                            if name.find(commodity) != -1:
                                                subname = name[name.find(commodity):name.find(commodity) + 6]
                                                if len(commodity) == 2:
                                                    contract = subname
                                                else:
                                                    if subname[-1] != "_":
                                                        continue
                                                    else:
                                                        contract = subname[:5]
                                                print(contract)
                                                with zfile.open(name) as readfile:
                                                    lines = []
                                                    for line in io.TextIOWrapper(readfile, encoding):
                                                        lines.append(line.split("\t"))

                                                brokers_dict = {}
                                                table_nr = 0  # First table is volume, second table is long, third table is short
                                                for line in lines:
                                                    if line[0] == '名次':
                                                        table_nr += 1

                                                    if line[0] == '总计' and table_nr == 1:
                                                        cnt_volume.append(float(line[4].replace(',', '')))

                                                    if is_integer(line[0]) == True:
                                                        if all_top_brokers == False and int(line[0]) > 20:
                                                            continue
                                                        else:
                                                            broker_name = line[2]
                                                            value = line[3]
                                                            if table_nr == 1:
                                                                if (broker_name in brokers_dict.keys()):
                                                                    position_dict = brokers_dict[broker_name]
                                                                else:
                                                                    position_dict = {}
                                                                position_dict["volume"] = float(value.replace(",", ""))
                                                                brokers_dict[broker_name] = position_dict.copy()
                                                            elif table_nr == 2:
                                                                if (broker_name in brokers_dict.keys()):
                                                                    position_dict = brokers_dict[broker_name]
                                                                else:
                                                                    position_dict = {}
                                                                position_dict["long"] = float(value.replace(",", ""))
                                                                brokers_dict[broker_name] = position_dict.copy()
                                                            elif table_nr == 3:
                                                                if (broker_name in brokers_dict.keys()):
                                                                    position_dict = brokers_dict[broker_name]
                                                                else:
                                                                    position_dict = {}
                                                                position_dict["short"] = float(value.replace(",", ""))
                                                                brokers_dict[broker_name] = position_dict.copy()
                                            cont_dict[contract] = brokers_dict.copy()
                                        df_volume = pd.DataFrame({'contract': list(cont_dict.keys()), 'volume': cnt_volume})
                                        df_volume.sort_values('volume', ascending=False, ignore_index=True, inplace=True)
                                        if max_num != 'all':
                                            main_cont = df_volume['contract'].iloc[0]
                                            cont_dict2 = cont_dict
                                            cont_dict = {}
                                            cont_dict[main_cont] = cont_dict2[main_cont]

                            except:
                                if number_of_tries > 5:
                                    break
                                continue

                        days_dict[date] = cont_dict.copy()
                    else:
                        print("empty day")

                    driver.refresh()
    print("It took: ", (time.time() - start) / 60, " minutes")
    print(kok_praznih)

    driver.close()

    ## The data scraped has the following dictionarial structure: dict_days[datetime]:
    ##                                                                    cont_dict[i1901]:
    ##                                                                       brokers_dict[ijangwo]:
    ##                                                                            position_dict[short]: 100

    #### Put data together ###
    if commodity != 'all':
        daily_lists = {}
        all_brokers = set()

        for i in range(1, len(list(days_dict.keys()))):
            date = list(days_dict.keys())[i]
            contracts_on_date = list(days_dict[date].keys())  # a list of all active contracts on that date
            net_positions_on_date = {}  # a dict with brokers as keys - used to net up all long and short positions over all contracts on this day
            volumes_on_date = {}  # volume on that day
            volume_on_contracts = {}  # volume in each contract
            for contract in contracts_on_date:
                top_brokers_on_contract = days_dict[date][contract].keys()
                for broker in top_brokers_on_contract:
                    all_brokers.add(broker)
                    if 'short' in days_dict[date][contract][broker].keys():
                        # try If the dictionary has key for this broker , else do except
                        try:
                            net_positions_on_date[broker] -= days_dict[date][contract][broker]['short']
                        except:
                            net_positions_on_date[broker] = -days_dict[date][contract][broker]['short']
                    if 'long' in days_dict[date][contract][broker].keys():
                        # try If the dictionary has key for this broker , else do except
                        try:
                            net_positions_on_date[broker] += days_dict[date][contract][broker]['long']
                        except:
                            net_positions_on_date[broker] = days_dict[date][contract][broker]['long']
                    if 'volume' in days_dict[date][contract][broker].keys():
                        # try If the dictionary has key for this broker , else do except
                        try:
                            volumes_on_date[broker] += days_dict[date][contract][broker]['volume']
                        except:
                            volumes_on_date[broker] = days_dict[date][contract][broker]['volume']
                            # try if this contract is already in dict else do except
                        try:
                            volume_on_contracts[contract] += days_dict[date][contract][broker]['volume']
                        except:
                            volume_on_contracts[contract] = days_dict[date][contract][broker]['volume']

            # add volumes for each broker that has information about that
            for broker in list(net_positions_on_date.keys()):
                try:
                    net_positions_on_date[broker] = [net_positions_on_date[broker], volumes_on_date[broker]]
                except:
                    net_positions_on_date[broker] = [net_positions_on_date[broker], 0]

            df = pd.DataFrame.from_dict(net_positions_on_date, orient="index", columns=["net_position", "volume"])
            df = df.sort_values(by=["net_position"], ascending=False)

            # Find dominant contract
            contracts_by_volume = pd.DataFrame.from_dict(volume_on_contracts, orient="index", columns=["volume"])
            contracts_by_volume = contracts_by_volume.sort_values(by=["volume"], ascending=False)
            # dominant_contracts = np.array(["     " for ii in range(len(net_positions_on_date))])
            # dominant_contracts[0:len(list(contracts_by_volume.index))]=list(contracts_by_volume.index)
            dominant_contracts = []
            volume_dom_contracts = []
            for pp in range(len(df)):
                if pp < len(list(contracts_by_volume.index)):
                    dominant_contracts.append(list(contracts_by_volume.index)[pp])
                    volume_dom_contracts.append(contracts_by_volume["volume"].iloc[pp])
                else:
                    dominant_contracts.append(np.nan)
                    volume_dom_contracts.append(np.nan)
            df["dominant contract"] = dominant_contracts
            df["dominant contract volume"] = volume_dom_contracts

            daily_lists[date] = df

        ## Write to excel
        with pd.ExcelWriter(os.path.dirname(__file__) + "\\data\\" + commodity +
                            " open interest {1}_{0}.xlsx".format(max_num, end_year)) as writer:  # enter file name for different product
            pd.Series(list(all_brokers), name="broker").to_excel(writer, sheet_name="all brokers")
            days = list(daily_lists.keys())
            days.reverse()
            for day in days:
                daily_list = daily_lists[day]
                if not daily_list.empty:
                    daily_list.to_excel(writer, sheet_name=str(day)[0:10].replace("-", "."))

        def delete_downloaded_files():
            for el in os.listdir(download_directory):
                if el.find("_DCE_DPL") != -1:
                    os.remove(download_directory + el)

        delete_downloaded_files()


if __name__ == "__main__":
    scrape_DCE(commodity, max_num, end_year)