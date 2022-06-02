#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@File    :   123.py
@Time    :   2022/03/06 16:18:33
@Author  :   Pan Yifan
@Version :   1.0
@Desc    :   None
1. 
'''

#%% 0. 工具库
import os
import gc
import pandas as pd

#%% 1. 参数设置
# 1.1. 文件夹位置
file_list = os.path.realpath(__file__).split('\\')[:-1]
file_path = '\\'.join(file_list)

# print(file_path)


#%% 2. 函数设置
def xlsx2all(file_path=file_path):
    xlsx_all = []
    for i in os.listdir(file_path):
        if i.endswith('xlsx') and i != '000.xlsx':
            data = pd.read_excel(file_path + '\\' + i, index_col=[0], engine="openpyxl")
            xlsx_all.append(data)
            del data
            gc.collect()
    xlsx_all = pd.concat(xlsx_all, axis=0).reset_index(drop=True)
    xlsx_all['date'] = pd.to_datetime(xlsx_all['date'], format='%Y-%m-%d').dt.date
    xlsx_all = xlsx_all.sort_values(by='date', ascending=True).drop_duplicates().reset_index(drop=True)
    xlsx_all.to_excel(file_path + '\\000.xlsx')
    print(xlsx_all)


#%% 3. 函数运行
if __name__ == "__main__":
    xlsx2all()