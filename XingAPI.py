import getData
import Account

stockAPI = getData
loginAPI = Account.Account()

class XingAPI:

    def login(self, path):
        loginAPI.Login(path)

    def logout(self):
        loginAPI.Logout()

    def t1514_업종기간별추이(self, 업종코드, 구분1, 구분2, CTS일자, 조회건수, 비중구분):
        stockAPI.T1514().Request(업종코드, 구분1, 구분2, CTS일자, 조회건수, 비중구분)
        result = stockAPI.T1514().GetResult()
        return result


