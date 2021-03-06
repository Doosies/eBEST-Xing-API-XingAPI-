
# from pandas import DataFrame, Series, Panel
# from time import sleep
# import matplotlib
# import matplotlib.pyplot as plt
# import win32com.client
# import pythoncom
from time import sleep
import pandas as pd
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
        # print(self.accounts)

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
        result = getData.T1514().GetResult(업종코드, 구분1, 구분2, CTS일자, 조회건수, 비중구분)
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
        result = getData.T0424_주식잔고2().GetResult(계좌번호, 비밀번호, 단가구분, 체결구분, 단일가구분, 제비용포함여부, CTS_종목번호)
        return result

    def t8412_주식차트N분(self, 단축코드, 분단위, 요청건수, 연속조회, cts_date, cts_time):
        """주식잔고2 조회함수
        :param 단축코드: 주식종목의 단축코드 입력
        :param 분단위: 시간의 단위 입력 (1: 1분, 2: 2분, 3: 3분 ... n: n분)
        :param 요청건수: 한번 블럭을 받아올 때 2000건까지 가능
        :param cts_time: 최초 조회시 ''입력
        """
        result = getData.T8412_주식차트N분().GetResult(단축코드, 분단위, 요청건수, 연속조회, cts_date, cts_time)
        return result

    def t8436_주식종목조회(self, 구분):
        """주식잔고2 조회함수
        :param 구분: 0: 전체, 1: 코스피, 2: 코스닥
        """
        result = getData.T8436_주식종목조회().GetResult(구분)
        return result
    

if __name__ == "__main__":
    api = XingAPI()
    account_path = pd.read_csv('private\\info.csv')
    api.login(account_path)
    accounts = api.getAccount()

    # week_data1 = api.t1514_업종기간별추이(업종코드='001', 구분1='', 구분2='1', CTS일자='', 조회건수='10', 비중구분='')
    # print(week_data1)
    
    # test_data = api.t0424_주식잔고2(accounts[0], 0000, 1, 0, 0, 0, '')
    # print(test_data)

    test_data2 = api.t8412_주식차트N분(단축코드='005930', 분단위='5', 요청건수='2000', 연속조회 = False, cts_date='', cts_time='')
    test_data2.to_csv('output2.csv', index=False, mode='w',
           encoding='utf-8-sig')

    # sleep(0.9)
    # test_data2 = api.t8412_주식차트N분(단축코드='005930', 분단위='5', 요청건수='2000',  연속조회 = False, cts_date='', cts_time='')
    # test_data2.to_csv('output2.csv', index=False, mode='w',
    #        encoding='utf-8-sig')
    # print(test_data2)