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


# 서버에서 해당 이벤트를 발생시키는 함수
class XAQueryEvents(object):
    상태 = False

    def __init__(self):
        self.parent = None

    def set_params(self, parent):
        self.parent = parent

    def OnReceiveData(self, szTrCode):
        # print("OnReceiveData : %s" % szTrCode)
        if self.parent != None:
            self.parent.OnReceiveData(szTrCode)

        XAQueryEvents.상태 = True

    def OnReceiveMessage(self, systemError, messageCode, message):
        # print("OnReceiveMessage : ", systemError, messageCode, message)
        pass

# 서버에서 데이터가 올 때 까지 대기하는 함수
def Waiting():
    while XAQueryEvents.상태 == False:
        pythoncom.PumpWaitingMessages()

# 다른 데이터를 받는 함수들에서 공통적으로 초기화 해줘야 할 부분을 추상클래스로 빼버림
class DataParent:
    def __init__(self,kind):
        self.RESDIR = 'C:\\eBEST\\xingAPI\\Res\\'
        self.MYNAME = kind
        self.RESFILE = self.RESDIR + self.MYNAME + ".res"
        self.INBLOCK = "%sInBlock" % self.MYNAME
        self.OUTBLOCK = "%sOutBlock" % self.MYNAME
        self.OUTBLOCK1 = "%sOutBlock1" % self.MYNAME

        self.query = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
        self.query.ResFileName = self.RESFILE
        self.query.set_params(parent=self)


        self.result = []

    def Request(self):
        """데이터 요청
        성공시 OnReceiveData() 실행
        """
        pass

    def OnReceiveData(self):
        """데이터 수신
        데이터가 수신되면 실행되는 함수
        """
        pass

    def GetResult(self):
        """결과값을 리턴
        """
        pass

# 업종 기간별 추이
class T1514(DataParent):
    '''
    업종기간별추이!!!
    '''

    def __init__(self):
        super().__init__('t1514')

    def Request(self, 업종코드, 구분1, 구분2, CTS일자, 조회건수, 비중구분):
        self.query.SetFieldData(self.INBLOCK, "upcode", 0, 업종코드)
        self.query.SetFieldData(self.INBLOCK, "gubun1", 0, 구분1)
        self.query.SetFieldData(self.INBLOCK, "gubun2", 0, 구분2)
        self.query.SetFieldData(self.INBLOCK, "cts_date", 0, CTS일자)
        self.query.SetFieldData(self.INBLOCK, "cnt", 0, 조회건수)
        self.query.SetFieldData(self.INBLOCK, "rate_gbn", 0, 비중구분)
        self.query.Request(0)
        Waiting()
        
    def OnReceiveData(self,szTrCode):
        nCount = self.query.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            일자 = self.query.GetFieldData(self.OUTBLOCK1, "date", i).strip()
            지수 = float(self.query.GetFieldData(self.OUTBLOCK1, "jisu", i).strip())
            전일대비구분 = self.query.GetFieldData(self.OUTBLOCK1, "sign", i).strip()
            전일대비 = float(self.query.GetFieldData(self.OUTBLOCK1, "change", i).strip())
            등락율 = float(self.query.GetFieldData(self.OUTBLOCK1, "diff", i).strip())
            거래량 = int(self.query.GetFieldData(self.OUTBLOCK1, "volume", i).strip())
            거래증가율 = float(self.query.GetFieldData(self.OUTBLOCK1, "diff_vol", i).strip())
            거래대금1 = int(self.query.GetFieldData(self.OUTBLOCK1, "value1", i).strip())
            상승 = int(self.query.GetFieldData(self.OUTBLOCK1, "high", i).strip())
            보합 = int(self.query.GetFieldData(self.OUTBLOCK1, "unchg", i).strip())
            하락 = int(self.query.GetFieldData(self.OUTBLOCK1, "low", i).strip())
            상승종목비율 = float(self.query.GetFieldData(self.OUTBLOCK1, "uprate", i).strip())
            외인순매수 = int(self.query.GetFieldData(self.OUTBLOCK1, "frgsvolume", i).strip())
            시가 = float(self.query.GetFieldData(self.OUTBLOCK1, "openjisu", i).strip())
            고가 = float(self.query.GetFieldData(self.OUTBLOCK1, "highjisu", i).strip())
            저가 = float(self.query.GetFieldData(self.OUTBLOCK1, "lowjisu", i).strip())
            거래대금2 = int(self.query.GetFieldData(self.OUTBLOCK1, "value2", i).strip())
            상한 = int(self.query.GetFieldData(self.OUTBLOCK1, "up", i).strip())
            하한 = int(self.query.GetFieldData(self.OUTBLOCK1, "down", i).strip())
            종목수 = int(self.query.GetFieldData(self.OUTBLOCK1, "totjo", i).strip())
            기관순매수 = int(self.query.GetFieldData(self.OUTBLOCK1, "orgsvolume", i).strip())
            업종코드 = self.query.GetFieldData(self.OUTBLOCK1, "upcode", i).strip()
            거래비중 = float(self.query.GetFieldData(self.OUTBLOCK1, "rate", i).strip())
            업종배당수익률 = float(self.query.GetFieldData(self.OUTBLOCK1, "divrate", i).strip())

            lst = [일자, 지수, 전일대비구분, 전일대비, 등락율, 거래량, 거래증가율, 거래대금1, 상승, 보합, 하락, 상승종목비율,
                   외인순매수, 시가, 고가, 저가, 거래대금2, 상한, 하한, 종목수, 기관순매수, 업종코드, 거래비중, 업종배당수익률]

            self.result.append(lst)

        XAQueryEvents.상태 = False

    def GetResult(self):
        columns = ['일자', '지수', '전일대비구분', '전일대비', '등락율', '거래량', '거래증가율', '거래대금1', '상승', '보합', '하락', '상승종목비율', '외인순매수',
                   '시가', '고가', '저가', '거래대금2', '상한', '하한', '종목수', '기관순매수', '업종코드', '거래비중', '업종배당수익률']

        return DataFrame(data=self.result, columns=columns)

