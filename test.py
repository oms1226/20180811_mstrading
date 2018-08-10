# -*-coding:utf-8 -*-
import sys
from PyQt5.QtWidgets import QMainWindow, QPushButton
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtWidgets import QApplication


class MyWindow(QMainWindow):
    def __init__(self):
        super(MyWindow, self).__init__()
        self.setWindowTitle("PyStock")
        self.setGeometry(300, 300, 300, 400)

        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.kiwoom.OnReceiveTrData.connect(self.OnReceiveTrData)

        btn1 = QPushButton("Log In", self)
        btn1.move(20, 20)
        btn1.clicked.connect(self.btn_clicked)

        btn2 = QPushButton("Get Info", self)
        btn2.move(20, 70)
        btn2.clicked.connect(self.btn_clicked2)

        btn3 = QPushButton("connect state", self)
        btn3.move(20, 120)
        btn3.clicked.connect(self.btn_clicked3)

    def OnReceiveTrData(self, sScrNo, sRQName, sTRCode, sRecordName, sPreNext, nDataLength, sErrorCode, sMessage, sSPlmMsg):
        if sRQName == "주식기본정보":
            cnt = self.kiwoom.dynamicCall('GetRepeatCnt(QString, QString)', sTRCode, sRQName)
            name = self.kiwoom.dynamicCall('GetCommData(QString, QString, int, QString)', sTRCode, sRQName, 0, "종목명")
            cur_price = self.kiwoom.dynamicCall('GetCommData(QString, QString, int, QString)', sTRCode, sRQName, 0, "현재가")
            print(name.strip())
            print(cur_price.strip())


    def btn_clicked(self):
        ret = self.kiwoom.dynamicCall("CommConnect()")

    def btn_clicked2(self):
        ret = self.kiwoom.dynamicCall('SetInputValue(QString, QString)', "종목코드", "000660")
        ret = self.kiwoom.dynamicCall('CommRqData(QString, QString, int, QString)', "주식기본정보", "OPT10001", 0, "0101")

    def btn_clicked3(self):
        if self.kiwoom.dynamicCall('GetConnectState()') == 0:
            print("Not connected")
        else:
            print("Connnected")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    myWindow.show()
    app.exec_()
