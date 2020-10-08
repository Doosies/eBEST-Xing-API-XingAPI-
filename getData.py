import win32com.client
import pythoncom
# import sqlite3
# import pandas as pd
# from pandas import DataFrame, Series, Panel
# import matplotlib
# import matplotlib.pyplot as plt

def waiting():
    while XAQueryEvents.상태 == False:
        pythoncom.PumpWaitingMessages()
    XAQueryEvents.상태 = False

# 서버에서 해당 이벤트를 발생시키는 함수
class XAQueryEvents:
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

# 다른 데이터를 받는 함수들에서 공통적으로 초기화 해줘야 할 부분을 추상클래스로 빼버림
class DataParent:
    def __init__(self,kind):
        self.RESDIR = 'C:\\eBEST\\xingAPI\\Res\\'
        self.MYNAME = kind
        self.RESFILE = self.RESDIR + self.MYNAME + ".res"
        self.INBLOCK = "%sInBlock" % self.MYNAME
        self.OUTBLOCK = "%sOutBlock" % self.MYNAME
        self.OUTBLOCK1 = "%sOutBlock1" % self.MYNAME
        self.OUTBLOCK2 = "%sOutBlock2" % self.MYNAME
        
        # query_events = XAQueryEvents
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
        
        # getData.XAQueryEvents.상태 = False

    def GetResult(self, 업종코드, 구분1, 구분2, CTS일자, 조회건수, 비중구분):
        self.Request(업종코드, 구분1, 구분2, CTS일자, 조회건수, 비중구분)
        waiting()
        columns = ['일자', '지수', '전일대비구분', '전일대비', '등락율', '거래량', '거래증가율', '거래대금1', '상승', '보합', '하락', '상승종목비율', '외인순매수',
                   '시가', '고가', '저가', '거래대금2', '상한', '하한', '종목수', '기관순매수', '업종코드', '거래비중', '업종배당수익률']

        return DataFrame(data=self.result, columns=columns)

class T0424_주식잔고2(DataParent):
    '''
    주식잔고2!!!
    '''
    def __init__(self): 
        super().__init__('t0424')

    def Request(self, 계좌번호, 비밀번호, 단가구분, 체결구분, 단일가구분, 제비용포함여부, CTS_종목번호):
        self.query.SetFieldData(self.INBLOCK, "accno", 0, 계좌번호)
        self.query.SetFieldData(self.INBLOCK, "passwd", 0, 비밀번호)
        self.query.SetFieldData(self.INBLOCK, "prcgb", 0, 단가구분)
        self.query.SetFieldData(self.INBLOCK, "chegb", 0, 체결구분)
        self.query.SetFieldData(self.INBLOCK, "dangb", 0, 단일가구분)
        self.query.SetFieldData(self.INBLOCK, "charge", 0, 제비용포함여부)
        self.query.SetFieldData(self.INBLOCK, "cts_expcode", 0, CTS_종목번호)
        self.query.Request(0)
        # Waiting()

    def OnReceiveData(self,szTrCode):
        nCount = self.query.GetBlockCount(self.OUTBLOCK)
        for i in range(nCount):
            추정순자산 = self.query.GetFieldData(self.OUTBLOCK,"sunamt",i).strip()
            실현손익 = self.query.GetFieldData(self.OUTBLOCK,"dtsunik",i).strip()
            매입금액 = self.query.GetFieldData(self.OUTBLOCK,"mamt",i).strip()
            추정D2예수금 = self.query.GetFieldData(self.OUTBLOCK,"sunamt1",i).strip()
            CTS_종목번호 = self.query.GetFieldData(self.OUTBLOCK,"cts_expcode",i).strip()
            평가금액 = self.query.GetFieldData(self.OUTBLOCK,"tappamt",i).strip()
            평가손익 = self.query.GetFieldData(self.OUTBLOCK,"tdtsunik",i).strip()

            lst = [추정순자산,실현손익,매입금액,추정D2예수금,CTS_종목번호,평가금액,평가손익]

            self.result.append(lst)

    def GetResult(self, 계좌번호, 비밀번호, 단가구분, 체결구분, 단일가구분, 제비용포함여부, CTS_종목번호):
        self.Request(계좌번호, 비밀번호, 단가구분, 체결구분, 단일가구분, 제비용포함여부, CTS_종목번호)
        waiting()
        columns = ["추정순자산","실현손익","매입금액","추정D2예수금","CTS_종목번호","평가금액","평가손익"]
        return DataFrame(data=self.result, columns=columns)

class T8412_주식차트N분(DataParent):
    '''
    주식차트 조회
    '''
    def __init__(self): 
        super().__init__('t8412')

    def Request(self, 단축코드, 단위, 요청건수, cts_date):
        # 만약 최초 요청일경우
        if cts_date == '':
            self.query.SetFieldData(self.INBLOCK, "shcode", 0, 단축코드)
            self.query.SetFieldData(self.INBLOCK, "ncnt", 0, 단위)
            self.query.SetFieldData(self.INBLOCK, "qrycnt", 0, 요청건수)
            self.query.SetFieldData(self.INBLOCK, "nday", 0, 0)
            self.query.SetFieldData(self.INBLOCK, "edate", 0, '99999999')
            self.query.SetFieldData(self.INBLOCK, "cts_date", 0, cts_date)
            # self.query.SetFieldData(self.INBLOCK, "cts_time", 0, 연속시간)
            self.query.SetFieldData(self.INBLOCK, "comp_yn", 0, 'Y')
            self.query.Request(0)
        # 2번째 이상의 요청일 경우
        else:
            self.SetFieldData(self.INBLOCK, "cts_date", 0, self.CTS_DATE)
            err_code = self.Request(True) # 연속조회인경우만 True

            if err_code < 0:
                print("error... {0}".format(err_code))


        # Waiting()

    def OnReceiveData(self,szTrCode):
        # 더 많은 결과값을 받아오기 위해 압축모듈을 사용하므로 압축해제를 받아온 블럭의 압축을 해제해줌
        nOrgSize = self.query.Decompress("t8412OutBlock1")
        if nOrgSize > 0:
            nCount = self.query.GetBlockCount(self.OUTBLOCK1)
            for i in range(nCount):
                날짜 = self.query.GetFieldData(self.OUTBLOCK1,"date",i).strip()
                시간 = self.query.GetFieldData(self.OUTBLOCK1,"time",i).strip()
                시가 = self.query.GetFieldData(self.OUTBLOCK1,"open",i).strip()
                고가 = self.query.GetFieldData(self.OUTBLOCK1,"high",i).strip()
                저가 = self.query.GetFieldData(self.OUTBLOCK1,"low",i).strip()
                종가 = self.query.GetFieldData(self.OUTBLOCK1,"close",i).strip()
                거래량 = self.query.GetFieldData(self.OUTBLOCK1,"jdiff_vol",i).strip()
                거래대금 = self.query.GetFieldData(self.OUTBLOCK1,"value",i).strip()
                수정구분 = self.query.GetFieldData(self.OUTBLOCK1,"jongchk",i).strip()
                수정비율 = self.query.GetFieldData(self.OUTBLOCK1,"rate",i).strip()
                종가등락구분 = self.query.GetFieldData(self.OUTBLOCK1,"sign",i).strip()

                lst = [날짜,시간,시가,고가,저가,종가,거래량,거래대금,수정구분,수정비율,종가등락구분]

                self.result.append(lst)
        self.CTS_DATE = self.query.GetFieldData(self.OUTBLOCK,"date",0).strip()

    def GetResult(self, 단축코드, 단위, 요청건수, cts_date):
        self.Request(단축코드, 단위, 요청건수, cts_date)
        waiting()
        print("ctsdata = " + self.CTS_DATE)
        columns = ["날짜","시간","시가","고가","저가","종가","거래량","거래대금","수정구분","수정비율","종가등락구분"]
        return DataFrame(data=self.result, columns=columns)