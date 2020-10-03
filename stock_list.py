import win32com.client
import pythoncom

class stockListHandler:
    isGetData = False

    def OnReceiveData(self, code):
        stockListHandler.isGetData = True

