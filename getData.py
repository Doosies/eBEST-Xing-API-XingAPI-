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
        print("yeah")
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
        pass

    def OnReceiveData(self):
        pass

    def GetResult(self):
        pass

# 업종 기간별 추이
class T1514(DataParent):
    '''
    업종기간별추이
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
        Waiting()
        columns = ['일자', '지수', '전일대비구분', '전일대비', '등락율', '거래량', '거래증가율', '거래대금1', '상승', '보합', '하락', '상승종목비율', '외인순매수',
                   '시가', '고가', '저가', '거래대금2', '상한', '하한', '종목수', '기관순매수', '업종코드', '거래비중', '업종배당수익률']

        return DataFrame(data=self.result, columns=columns)

# 예수금상세현황요청, 주문가능금액, 총평가 조회
class CSPAQ12200(DataParent):
    '''
    예수금상세현황요청, 주문가능금액, 총평가조회
    '''
    def __init__(self):
        super().__init__('CSPAQ12200')

    def Request(self, 레코드갯수, 관리지점번호, 계좌번호, 비밀번호, 잔고생성구분):
        self.query.SetFieldData(self.INBLOCK, "RecCn", 0, 레코드갯수)
        self.query.SetFieldData(self.INBLOCK, "MgmtBrnNo", 0, 관리지점번호)
        self.query.SetFieldData(self.INBLOCK, "AcntNo", 0, 계좌번호)
        self.query.SetFieldData(self.INBLOCK, "Pwd", 0, 비밀번호)
        self.query.SetFieldData(self.INBLOCK, "BalCreTp", 0, 잔고생성구분)
        self.query.Request(0)

    def OnReceiveData(self):
        nCount = self.query.GetBlockCount(self.OUTBLOCK1)
        for i in range(nCount):
            레코드갯수 = self.query.GetFieldData(self.OUTBLOCK2,"RecCnt",i).strip()
            지점명 = self.query.GetFieldData(self.OUTBLOCK2,"BrnNm",i).strip()
            계좌명 = self.query.GetFieldData(self.OUTBLOCK2,"AcntNm",i).strip()
            현금주문가능금액 = self.query.GetFieldData(self.OUTBLOCK2,"MnyOrdAbleAmt",i).strip()
            출금가능금액 = self.query.GetFieldData(self.OUTBLOCK2,"MnyoutAbleAmt",i).strip()
            거래소금액 = self.query.GetFieldData(self.OUTBLOCK2,"SeOrdAbleAmt",i).strip()
            코스닥금액 = self.query.GetFieldData(self.OUTBLOCK2,"KdqOrdAbleAmt",i).strip()
            잔고평가금액 = self.query.GetFieldData(self.OUTBLOCK2,"BalEvalAmt",i).strip()
            미수금액 = self.query.GetFieldData(self.OUTBLOCK2,"RcvblAmt",i).strip()
            예탁자산총액 = self.query.GetFieldData(self.OUTBLOCK2,"DpsastTotamt",i).strip()
            손익율 = self.query.GetFieldData(self.OUTBLOCK2,"PnlRat",i).strip()
            투자원금 = self.query.GetFieldData(self.OUTBLOCK2,"InvstOrgAmt",i).strip()
            투자손익금액 = self.query.GetFieldData(self.OUTBLOCK2,"InvstPlAmt",i).strip()
            신용담보주문금액 = self.query.GetFieldData(self.OUTBLOCK2,"CrdtPldgOrdAmt",i).strip()
            예수금 = self.query.GetFieldData(self.OUTBLOCK2,"Dps",i).strip()
            대용금액 = self.query.GetFieldData(self.OUTBLOCK2,"SubstAmt",i).strip()
            D1예수금 = self.query.GetFieldData(self.OUTBLOCK2,"D1Dps",i).strip()
            D2예수금 = self.query.GetFieldData(self.OUTBLOCK2,"D2Dps",i).strip()
            현금미수금액 = self.query.GetFieldData(self.OUTBLOCK2,"MnyrclAmt",i).strip()
            증거금현금 = self.query.GetFieldData(self.OUTBLOCK2,"MgnMny",i).strip()
            증거금대용 = self.query.GetFieldData(self.OUTBLOCK2,"MgnSubst",i).strip()
            수표금액 = self.query.GetFieldData(self.OUTBLOCK2,"ChckAmt",i).strip()
            대용주문가능금액 = self.query.GetFieldData(self.OUTBLOCK2,"SubstOrdAbleAmt",i).strip()
            증거금률100퍼센트주문가능금액 = self.query.GetFieldData(self.OUTBLOCK2,"MgnRat100pctOrdAbleAmt[",i).strip()
            증거금률35주문가능금액 = self.query.GetFieldData(self.OUTBLOCK2,"MgnRat35ordAbleAmt",i).strip()
            증거금률50주문가능금액 = self.query.GetFieldData(self.OUTBLOCK2,"MgnRat50ordAbleAmt",i).strip()
            전일매도정산금액 = self.query.GetFieldData(self.OUTBLOCK2,"PrdaySellAdjstAmt",i).strip()
            전일매수정산금액 = self.query.GetFieldData(self.OUTBLOCK2,"PrdayBuyAdjstAmt",i).strip()
            금일매도정산금액 = self.query.GetFieldData(self.OUTBLOCK2,"CrdaySellAdjstAmt",i).strip()
            금일매수정산금액 = self.query.GetFieldData(self.OUTBLOCK2,"CrdayBuyAdjstAmt",i).strip()
            D1연체변제소요금액 = self.query.GetFieldData(self.OUTBLOCK2,"D1ovdRepayRqrdAmt",i).strip()
            D2연체변제소요금액 = self.query.GetFieldData(self.OUTBLOCK2,"D2ovdRepayRqrdAmt",i).strip()
            D1추정인출가능금액 = self.query.GetFieldData(self.OUTBLOCK2,"D1PrsmptWthdwAbleAmt[",i).strip()
            D2추정인출가능금액 = self.query.GetFieldData(self.OUTBLOCK2,"D2PrsmptWthdwAbleAmt[",i).strip()
            예탁담보대출금액 = self.query.GetFieldData(self.OUTBLOCK2,"DpspdgLoanAmt",i).strip()
            신용설정보증금 = self.query.GetFieldData(self.OUTBLOCK2,"Imreq",i).strip()
            융자금액 = self.query.GetFieldData(self.OUTBLOCK2,"MloanAmt",i).strip()
            변경후담보비율 = self.query.GetFieldData(self.OUTBLOCK2,"ChgAfPldgRat",i).strip()
            원담보금액 = self.query.GetFieldData(self.OUTBLOCK2,"OrgPldgAmt",i).strip()
            부담보금액 = self.query.GetFieldData(self.OUTBLOCK2,"SubPldgAmt",i).strip()
            소요담보금액 = self.query.GetFieldData(self.OUTBLOCK2,"RqrdPldgAmt",i).strip()
            원담보부족금액 = self.query.GetFieldData(self.OUTBLOCK2,"OrgPdlckAmt",i).strip()
            담보부족금액 = self.query.GetFieldData(self.OUTBLOCK2,"PdlckAmt",i).strip()
            추가담보현금 = self.query.GetFieldData(self.OUTBLOCK2,"AddPldgMny",i).strip()
            D1주문가능금액 = self.query.GetFieldData(self.OUTBLOCK2,"D1OrdAbleAmt",i).strip()
            신용이자미납금액 = self.query.GetFieldData(self.OUTBLOCK2,"CrdtIntdltAmt",i).strip()
            기타대여금액 = self.query.GetFieldData(self.OUTBLOCK2,"EtclndAmt",i).strip()
            익일추정반대매매금액 = self.query.GetFieldData(self.OUTBLOCK2,"NtdayPrsmptCvrgAmt",i).strip()
            원담보합계금액 = self.query.GetFieldData(self.OUTBLOCK2,"OrgPldgSumAmt",i).strip()
            신용주문가능금액 = self.query.GetFieldData(self.OUTBLOCK2,"CrdtOrdAbleAmt",i).strip()
            부담보합계금액 = self.query.GetFieldData(self.OUTBLOCK2,"SubPldgSumAmt",i).strip()
            신용담보금현금 = self.query.GetFieldData(self.OUTBLOCK2,"CrdtPldgAmtMny",i).strip()
            신용담보대용금액 = self.query.GetFieldData(self.OUTBLOCK2,"CrdtPldgSubstAmt",i).strip()
            추가신용담보현금 = self.query.GetFieldData(self.OUTBLOCK2,"AddCrdtPldgMny",i).strip()
            신용담보재사용금액 = self.query.GetFieldData(self.OUTBLOCK2,"CrdtPldgRuseAmt",i).strip()
            추가신용담보대용 = self.query.GetFieldData(self.OUTBLOCK2,"AddCrdtPldgSubst",i).strip()
            매도대금담보대출금액 = self.query.GetFieldData(self.OUTBLOCK2,"CslLoanAmtdt1",i).strip()
            처분제한금액 = self.query.GetFieldData(self.OUTBLOCK2,"DpslRestrcAmt",i).strip()

            lst = [레코드갯수,지점명,계좌명,현금주문가능금액,출금가능금액,거래소금액,코스닥금액,잔고평가금액,미수금액,예탁자산총액,
            손익율,투자원금,투자손익금액,신용담보주문금액,예수금,대용금액,D1예수금,D2예수금,현금미수금액,증거금현금,증거금대용,
            수표금액,대용주문가능금액,증거금률100퍼센트주문가능금액,증거금률35%주문가능금액,증거금률50%주문가능금액,전일매도정산금액,
            전일매수정산금액,금일매도정산금액,금일매수정산금액,D1연체변제소요금액,D2연체변제소요금액,D1추정인출가능금액,D2추정인출가능금액,
            예탁담보대출금액,신용설정보증금,융자금액,변경후담보비율,원담보금액,부담보금액,소요담보금액,원담보부족금액,담보부족금액,추가담보현금,
            D1주문가능금액,신용이자미납금액,기타대여금액,익일추정반대매매금액,원담보합계금액,신용주문가능금액,부담보합계금액,신용담보금현금,
            신용담보대용금액,추가신용담보현금,신용담보재사용금액,추가신용담보대용,매도대금담보대출금액,처분제한금액]

            self.result.append(lst)

        XAQueryEvents.상태 = False

    def GetResult(self):
        Waiting()
        columns = ["레코드갯수","지점명","계좌명","현금주문가능금액","출금가능금액","거래소금액","코스닥금액","잔고평가금액","미수금액",
        "예탁자산총액","손익율","투자원금","투자손익금액","신용담보주문금액","예수금","대용금액","D1예수금","D2예수금","현금미수금액",
        "증거금현금","증거금대용","수표금액","대용주문가능금액","증거금률100퍼센트주문가능금액","증거금률35%주문가능금액","증거금률50%주문가능금액",
        "전일매도정산금액","전일매수정산금액","금일매도정산금액","금일매수정산금액","D1연체변제소요금액","D2연체변제소요금액","D1추정인출가능금액",
        "D2추정인출가능금액","예탁담보대출금액","신용설정보증금","융자금액","변경후담보비율","원담보금액","부담보금액","소요담보금액","원담보부족금액",
        "담보부족금액","추가담보현금","D1주문가능금액","신용이자미납금액","기타대여금액","익일추정반대매매금액","원담보합계금액","신용주문가능금액",
        "부담보합계금액","신용담보금현금","신용담보대용금액","추가신용담보현금","신용담보재사용금액","추가신용담보대용","매도대금담보대출금액","처분제한금액",]
        return DataFrame(data=self.result, columns=columns)