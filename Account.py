import win32com.client
import pythoncom

class XASessionEvents:
    상태 = False

    def OnLogin(self, code, msg):
        print("OnLogin : ", code, msg)
        XASessionEvents.상태 = True

    def OnLogout(self):
        print('--------------------')
        pass

    def OnDisconnect(self):
        # print('=====================')
        pass

class Account:
    def __init__(self):
        self.session = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEvents)
        self.account = []

    def Login(self, account_path):
        id = account_path['0'][0]
        pwd = account_path['1'][0]
        cert = account_path['2'][0]

        port = 20001
        url = 'demo.ebestsec.co.kr'
        # session.SetMode("_XINGAPI7_", "TRUE")
        # 서버에 연결함.
        result = self.session.ConnectServer(url, port)

        # 서버에 연결이 안되면
        if not result:
            nErrCode = self.session.GetLastError()
            strErrMsg = self.session.GetErrorMessage(nErrCode)
            # 에러코드를 리턴시킴
            return (False, nErrCode, strErrMsg, None, self.session)

        # 연결이 된다면 로그인을 진행함.
        self.session.Login(id, pwd, cert, 0, False)



       

    def Logout(self):
        self.session.Logout()
        self.session.DisconnectServer()

    def getAccount(self):
        # 서버에서 데이터를 송신할 때까지 대기
        while XASessionEvents.상태 == False:
            pythoncom.PumpWaitingMessages()
            
        account_cnt = self.session.GetAccountListCount()
        # 계좌정보를 account에 넣음
        for i in range(account_cnt):
            self.account.append(self.session.GetAccountList(i))

        for number in self.account:
            print(number)

        return self.account
