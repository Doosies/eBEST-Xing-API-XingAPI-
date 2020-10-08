import pandas as pd
from pandas import DataFrame, Series, Panel
from time import sleep
import matplotlib
import matplotlib.pyplot as plt
import win32com.client
import pythoncom

import getData
import Account






class XingAPI:

    def __init__(self):
        # stockAPI = getData
        self.loginAPI = Account.Account()

    def getAccount(self):
        return self.accounts

    def login(self, path):
        self.loginAPI.Login(path)
        self.accounts = self.loginAPI.getAccount()
        print(self.accounts)

    def logout(self):
        self.loginAPI.Logout()


    def t1514_업종기간별추이(self, 업종코드, 구분1, 구분2, CTS일자, 조회건수, 비중구분):
        """업종기간별추이 조회함수
        :param 업종코드: 업종 코드를 입력
        :param 구분1: 미사용 항목, 스페이스설정
        :param 구분2: 1:일, 2:주, 3:월
        :param CTS일자: 연속조회일 경우 이 값기준으로 조회(cont1일때)(이전 조회한 cts_date 값으로 설정), 처음 조회시 스페이스 설정
        :param 조회건수: 조회건수 입력
        :param 비중구분: 1:거래량비중, 2:거래대금비중
        """
        api = getData.T1514()
        # api.Request(업종코드, 구분1, 구분2, CTS일자, 조회건수, 비중구분)
        # getData.XAQueryEvents.상태 = False
        result = api.GetResult(업종코드, 구분1, 구분2, CTS일자, 조회건수, 비중구분)
        return result

    def CSPAQ12200_예수금조회(self):
        pass
    
    def t0424_주식잔고2(self, 계좌번호, 비밀번호, 단가구분, 체결구분, 단일가구분, 제비용포함여부, CTS_종목번호):
        """주식잔고2 조회함수
        :param 계좌번호: 계좌번호 입력
        :param 비밀번호: 비밀번호 입력(모의투자일경우 0000)
        :param 단가구분: 1:평균단가, 2:BEP단가
        :param 체결구분: 0:결제기준잔고, 2:체결기준(잔고가 0이 아닌 종목만 조회)
        :param 단일가구분: 0:정규장, 1:시간외단일가
        :param 제비용포함여부: 0:제비용미포함, 1:제비용포함
        :param CTS_종목번호: 처음조회시는 공백, 연속조회시는 이전 cts_expcode값으로 설정
        """
        api = getData.T0424_주식잔고2()
        api.Request(계좌번호, 비밀번호, 단가구분, 체결구분, 단일가구분, 제비용포함여부, CTS_종목번호)
        getData.XAQueryEvents.상태 = False
        result = api.GetResult()
        return result

    def t8412_주식차트N분(self, 단축코드, 분단위, 요청건수, cts_time):
        api = getData.T8412_주식차트N분()
        api.Request(단축코드, 분단위, 요청건수, cts_time)
        getData.XAQueryEvents.상태 = False
        result = api.GetResult()
        return result

if __name__ == "__main__":
    accountAPI = XingAPI()
    account_path = pd.read_csv('private\\info.csv')
    accountAPI.login(account_path)
    accounts = accountAPI.getAccount()

    week_data1 = XingAPI().t1514_업종기간별추이(업종코드='001', 구분1='', 구분2='1', CTS일자='', 조회건수='10', 비중구분='')
    print(week_data1)
    # week_data1.to_csv('output.csv', index=False, mode='w',encoding='utf-8-sig')
    sleep(1.0)
    week_data2 = XingAPI().t1514_업종기간별추이(업종코드='001', 구분1='', 구분2='1', CTS일자='', 조회건수='2', 비중구분='')
    print(week_data2)
    # week_data = XingAPI().t1514_업종기간별추이(업종코드='001', 구분1='', 구분2='1', CTS일자='', 조회건수='100', 비중구분='')
    # week_data.to_csv('C:\\Users\\SongMinhyung\\PycharmProjects\\pythonProject\\output.csv', index=False, mode='w',
    #         encoding='utf-8-sig')
    # test_data = API.t0424_주식잔고2(accounts[0], 0000, 1, 0, 0, 0, '')
    # print(test_data)
    # test_data = API.CSPAQ12200_예수금상세현황요청_주문가능금액_총평가조회(레코드갯수='', 관리지점번호='', 계좌번호=accounts[0],비밀번호=0000,잔고생성구분=0)
    # print(test_data)

    # test_data = XingAPI().t8412_주식차트N분(단축코드='005930', 분단위='5', 요청건수='10', cts_time='')
    # print(test_data)

    # sleep(1.0)
    # test_data2 = XingAPI().t8412_주식차트N분(단축코드='005930', 분단위='5', 요청건수='10', cts_time='')
    # print(test_data2)