import getData
import Account

import pandas as pd
from pandas import DataFrame, Series, Panel

import matplotlib
import matplotlib.pyplot as plt



stockAPI = getData
loginAPI = Account.Account()

class XingAPI:

    def getAccount(self):
        return self.accounts

    def login(self, path):
        loginAPI.Login(path)
        self.accounts = loginAPI.getAccount()

    def logout(self):
        loginAPI.Logout()


    def t1514_업종기간별추이(self, 업종코드, 구분1, 구분2, CTS일자, 조회건수, 비중구분):
        """업종기간별추이 조회함수
        :param 업종코드: 업종 코드를 입력
        :param 구분1: 미사용 항목, 스페이스설정
        :param 구분2: 일=1, 주=2, 월=3
        :param CTS일자: 연속조회일 경우 이 값기준으로 조회(cont1일때)(이전 조회한 cts_date 값으로 설정), 처음 조회시 스페이스 설정
        :param 조회건수: 조회건수 입력
        :param 비중구분: 거래량비중=1, 거래대금비중=2
        """
        api = stockAPI.T1514()
        api.Request(업종코드, 구분1, 구분2, CTS일자, 조회건수, 비중구분)
        result = api.GetResult()
        return result


if __name__ == "__main__":
    API = XingAPI()

    account_path = pd.read_csv('C:\\Users\\SongMinhyung\\PycharmProjects\\pythonProject\\private\\info.csv')
    API.login(account_path)

    week_data = API.t1514_업종기간별추이(업종코드='001', 구분1='', 구분2='1', CTS일자='', 조회건수='100', 비중구분='')
    week_data.to_csv('C:\\Users\\SongMinhyung\\PycharmProjects\\pythonProject\\output.csv', index=False, mode='w',
            encoding='utf-8-sig')

    # accounts = API.getAccount()
    # test_data = API.CSPAQ12200_예수금상세현황요청_주문가능금액_총평가조회(레코드갯수='', 관리지점번호='', 계좌번호=accounts[0],비밀번호=0000,잔고생성구분=0)
    # print(test_data)