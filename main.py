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
