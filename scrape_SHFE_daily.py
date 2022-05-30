import sys
import os
import time
from datetime import datetime, timedelta
import pandas as pd
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select

### Use this script to scrape any day you want - just call function scrape_day(year,month,day)

#### WEBSCRAPING PART

commodity = "hc"
max_num = 1


def scrape_day(commodity, year_var, month_var, day_var):
    def table_into_contracts(table):
        # this function returns how many contracts there are in the table and how many lines each contract has
        # table is given as a list and each element is a string that needs to be split by space
        conts = {}
        for i in range(len(table)):
            # print(i)
            val = table[i].split()
            if table[i].find(commodity) != -1:  # new contract starts
                cont = table[i][6:12]
                conts[cont] = ()  # each cont in dict will have as value a list of lines(mostly between 1 and 20 but can be less)
            #print(conts,len(conts.keys()),len(val))
            if len(val) == 12 and len(conts.keys()) > 0 and val[0].isdigit():
                tmp = conts[cont] + (val, )
                #print(conts[cont],tmp)
                conts[cont] = tmp
            elif len(val) == 9 and len(conts.keys()) > 0 and val[0].isdigit():
                val[1:1] = ['0', '0', '0']
                tmp = conts[cont] + (val, )
                #print(conts[cont],tmp)
                conts[cont] = tmp
            if val[0] == '合计':
                if len(val) != 9:
                    val[1:1] = ['0', '0']
                conts[cont] += (val, )
                # print(val)
        # print(conts)
        return (conts)

    chrome_options = Options()
    # chrome_options.add_argument("--headless")  # does it open web browser or not
    #chrome_options.add_argument('--proxy-server=%s' % proxies)
    #chrome_options.binary_location = '/usr/bin/chromium-browser'

    url = "http://www.shfe.com.cn/statements/dataview.html?paramid=delaymarket_rb"
    driver = webdriver.Chrome("C:\\chromedriver.exe", chrome_options=chrome_options)
    driver.get(url)

    # odpri z f12 inspect. potem pa v iframe je neka nova html stran.. copy c path in dobiš ta iframe find element by xpath.
    # Odpri z copy by xpth - treba se v inspectu pomikat po značkah

    sleep_seconds = 2
    start = time.time()

    Ranks = driver.find_element_by_xpath("""//*[@id="pm"]""")
    Ranks.click()
    time.sleep(sleep_seconds)

    clicked_on_hrc = False

    if clicked_on_hrc == False:
        try:
            commodity_el = driver.find_element_by_xpath('//*[@id="li_' + commodity + '"]/a/span')
            commodity_el.click()
            clicked_on_hrc = True
        except:
            pass

    days_dict = {}

    we_are_in_future = False

    test = []

    for year in range(year_var, year_var + 1):

        while 5 < 6:  # wait for page to load
            try:
                selectYear = Select(driver.find_element_by_xpath("""//*[@id="calendar"]/div/div/div/select[1]"""))
                break
            except:
                pass
        selectYear.select_by_value(str(year))
        for month in range(month_var - 1, month_var):
            if we_are_in_future == True:
                break
            while 5 < 6:
                try:
                    selectMonth = Select(driver.find_element_by_xpath("""//*[@id="calendar"]/div/div/div/select[2]"""))
                    break
                except:
                    pass
            selectMonth.select_by_index(month)
            # row od 1 to 7 col 2 do 7
            for row in range(1, 7):
                if we_are_in_future == True:
                    break
                for col in range(2, 7):

                    #  just wait for the page to load
                    #time.sleep(5)
                    try:
                        Day_element = driver.find_element_by_xpath("""//*[@id="calendar"]/div/table/tbody/tr[""" + str(row) + """]/td[""" + str(col) +
                                                                   """]/a""")
                        day = Day_element.text
                        if int(day) != day_var:
                            continue
                        date = datetime(year, month + 1, int(day))
                        print(date)
                        if datetime.today() < date:
                            we_are_in_future = True
                            break
                        Day_element.click()

                        time.sleep(sleep_seconds)

                        # if today is a holiday we have not clicked on the hrc product in the beginning so need to do it now

                        for try_k in range(100):
                            try:
                                time.sleep(3)
                                commodity_el = driver.find_element_by_xpath('//*[@id="li_' + commodity + '"]/a/span')
                                print("Found table")
                                break
                            except:
                                print("Did not find it")

                        time.sleep(sleep_seconds)
                        commodity_el.click()
                        time.sleep(sleep_seconds)

                        Table_element = driver.find_element_by_xpath("""//*[@id="divtable"]""")
                        table = Table_element.text.split("\n")
                        # print(table)
                        split_table = table_into_contracts(table)
                        # print(split_table)
                        cont_dict = {}
                        cnt_volume = []
                        for j in range(len(list(split_table.keys()))):
                            try:
                                cnt_volume.append(list(split_table.values())[j][-1][1])
                            except:
                                cnt_volume.append('0')
                        # print(cnt_volume)
                        for i in range(len(cnt_volume)):
                            if cnt_volume[i] == '合计':
                                cnt_volume[i] = 0
                        cnt_volume = list(map(int, cnt_volume))
                        df_volume = pd.DataFrame({'contract': list(split_table.keys()), 'volume': cnt_volume})
                        df_volume.sort_values('volume', ascending=False, ignore_index=True, inplace=True)
                        print(df_volume)

                        for j in range(len(list(split_table.keys()))):
                            cont = list(split_table.keys())[j]
                            if max_num != 'all':
                                # print(cont)
                                # print(df_volume['contract'].iloc[max_num - 1])
                                # print(split_table(df_volume[df_volume['rank'] == max_num]['contract'].iloc[0]))
                                if cont != df_volume['contract'].iloc[max_num - 1]:
                                    # print(1)
                                    continue

                            brokers_dict = {}
                            lines = list(split_table.values())[j]  # get all the lines(between 1 and 20)

                            for val in lines:  #go over top 20(or whatever it is) for this contract
                                if not val[0].isdigit():
                                    continue

                                if (val[1] in brokers_dict.keys()):
                                    position_dict = brokers_dict[val[1]]
                                else:
                                    position_dict = {}
                                position_dict["volume"] = float(val[2].replace(",", ""))
                                brokers_dict[val[1]] = position_dict.copy()

                                if (val[5] in brokers_dict.keys()):
                                    position_dict = brokers_dict[val[5]]
                                else:
                                    position_dict = {}
                                position_dict["long"] = float(val[6].replace(",", ""))
                                brokers_dict[val[5]] = position_dict.copy()

                                if (val[9] in brokers_dict.keys()):
                                    position_dict = brokers_dict[val[9]]
                                else:
                                    position_dict = {}
                                position_dict["short"] = float(val[10].replace(",", ""))
                                brokers_dict[val[9]] = position_dict.copy()
                            cont_dict[cont] = brokers_dict.copy()
                            print(cont, "succesful with lines #", len(lines) - 1)
                        days_dict[date] = cont_dict.copy()
                        print(cont_dict)

                    except:
                        print("no such day", row, col)
                        pass

    print("It took: ", (time.time() - start) / 60, " minutes")

    driver.close()

    # The data scraped has the following dictionarial structure: dict_days[datetime]:
    #                                                                    cont_dict[i1901]:
    #                                                                       brokers_dict[ijangwo]:
    #                                                                            position_dict[short]: 100

    #### Put data together ###
    # print(days_dict)
    daily_lists = {}
    all_brokers = set()

    for i in range(len(list(days_dict.keys())) - 1, -1, -1):  # Reverse order
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
        dominant_contracts = np.array(["      " for ii in range(len(net_positions_on_date))])
        dominant_contracts[0:len(list(contracts_by_volume.index))] = list(contracts_by_volume.index)
        df["dominant contract by volume"] = dominant_contracts

        daily_lists[date] = df

    ## Write to excel
    writer = pd.ExcelWriter(
        os.path.abspath(os.path.dirname(__file__)) + "\\others\\Top Brokers on " + str(year_var) + "." + str(month_var) + "." + str(day_var) + "-" +
        commodity + '-' + str(max_num) + ".xlsx")
    # pd.Series(list(all_brokers),name="broker").to_excel(writer,sheet_name="all brokers")
    for day, daily_list in daily_lists.items():
        daily_list.to_excel(writer, sheet_name=str(day)[0:10].replace("-", "."))
    writer.save()

    # holidays = pd.read_excel("Holidays.xlsx")
    # missing_dates = []
    # for kk in range(len(list(days_dict.values()))):
    #     if list(days_dict.values())[kk] == {}:
    #         #print(list(days_dict.keys())[kk])
    #         if list(days_dict.keys())[kk] not in list(holidays.iloc[:, 0]):
    #             missing_dates.append(list(days_dict.keys())[kk])


# today = datetime.today()

# scrape_day(today.year,today.month,today.day)
#
if __name__ == "__main__":
    # scrape_day('sn', 2018, 1, 11)
    scrape_day('rb', 2013, 8, 1)
# scrape_day(2022,1,13)
# scrape_day(2020,9,22)

#
#missing_dates = pd.read_excel("missing dates.xlsx")
#
#
#for el in missing_dates.iloc[:,0]:
#    tmp = el
#    scrape_day(tmp.year,tmp.month,tmp.day)
#
