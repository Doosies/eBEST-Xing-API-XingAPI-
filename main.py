import win32com.client
import pythoncom
import os
import sys
import inspect

import sqlite3

import pandas as pd
from pandas import DataFrame, Series, Panel

import matplotlib
import matplotlib.pyplot as plt

import XingAPI

API = XingAPI.XingAPI()

account_path = pd.read_csv('C:\\Users\\SongMinhyung\\PycharmProjects\\pythonProject\\private\\info.csv')
API.login(account_path)

week_data = API.t1514_업종기간별추이(업종코드='001', 구분1='', 구분2='1', CTS일자='', 조회건수='100', 비중구분='')
week_data.to_csv('C:\\Users\\SongMinhyung\\PycharmProjects\\pythonProject\\output.csv', index=False, mode='w',
           encoding='utf-8-sig')

# accounts = API.getAccount()
# test_data = API.CSPAQ12200_예수금상세현황요청_주문가능금액_총평가조회(레코드갯수='', 관리지점번호='', 계좌번호=accounts[0],비밀번호=2809,잔고생성구분=0)
# print(test_data)
