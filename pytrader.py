# coding:utf-8
"""
QtDesigner로 만든 UI와 해당 UI의 위젯에서 발생하는 이벤트를 컨트롤하는 클래스
"""

import sys, time
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QTableWidget, QTableWidgetItem
from PyQt5.QtCore import Qt, QTimer, QTime
from PyQt5 import uic
from Kiwoom import Kiwoom, ParameterTypeError, ParameterValueError, KiwoomProcessingError, KiwoomConnectError
from Kiwoom_sungwon import *

ui = uic.loadUiType("pytrader.ui")[0]


class MyWindow(QMainWindow, ui):
    def __init__(self):
        super().__init__()

        self.setupUi(self)
        self.show()

        self.kiwoom = Kiwoom_sungwon()
        self.kiwoom.commConnect()
        #self.kiwoom.log.setLevel('INFO')
        self.kiwoom.log.info("# MyWindow Init Start")

        self.server = self.kiwoom.getLoginInfo("GetServerGubun")

        if self.server is None or self.server != "1":
            self.serverGubun = "실제운영"
        else:
            self.serverGubun = "모의투자"

        self.codeList = self.kiwoom.getCodeList("0")

        # 메인 타이머
        self.timer = QTimer(self)
        self.timer.start(1000)
        self.timer.timeout.connect(self.timeout)

        # 잔고 및 보유종목 조회 타이머
        self.inquiryTimer = QTimer(self)
        self.inquiryTimer.start(1000 * 10)
        self.inquiryTimer.timeout.connect(self.timeout)

        self.setAccountComboBox()
        self.codeLineEdit.textChanged.connect(self.setCodeName)
        self.orderBtn.clicked.connect(self.sendOrder)
        self.inquiryBtn.clicked.connect(self.inquiryBalance)


        # 자동 주문
        self.isAutomaticOrder = self.AutomaticOrder_CheckBox.isChecked()
        self.AutomaticOrder_CheckBox.stateChanged.connect(self.AutomaticOrder_CheckBox_change)
        # 당일 매수한 종목 리스트
        self.todayBuyList = []
        self.stockList = dict()

        # 실시간 조건검색 시작
        self.realtimeConditonStart()

        # 자동 선정 종목 리스트 테이블 설정
        self.setAutomatedStocks()

        self.kiwoom.log.info("# MyWindow Init End")

    def timeout(self):
        """ 타임아웃 이벤트가 발생하면 호출되는 메서드 """

        # 어떤 타이머에 의해서 호출되었는지 확인
        sender = self.sender()
        self.kiwoom.log.info("# timeout : {}".format(sender))

        # 메인 타이머
        if id(sender) == id(self.timer):
            currentTime = QTime.currentTime().toString("hh:mm:ss")
            automaticOrderTime = QTime.currentTime().toString("hhmm")

            # 상태바 설정
            state = ""

            if self.kiwoom.getConnectState() == 1:

                state = self.serverGubun + " 서버 연결중"
            else:
                state = "서버 미연결"

            self.statusbar.showMessage("현재시간: " + currentTime + " | " + state)

            # 자동 주문 실행
            # 1100은 11시 00분을 의미합니다.
            if self.isAutomaticOrder and (900 <= int(automaticOrderTime) <= 2300):
                self.automaticOrder()

            # log
            if self.kiwoom.msg:
                self.logTextEdit.append(self.kiwoom.msg)
                self.kiwoom.msg = ""

        # 실시간 조회 타이머
        # 잔고 및 보유종목 조회 타이머
        else:
            if self.realtimeCheckBox.isChecked():
                self.inquiryBalance()

    def AutomaticOrder_CheckBox_change(self):
        self.isAutomaticOrder = self.AutomaticOrder_CheckBox.isChecked()

    def setCodeName(self):
        self.kiwoom.log.info("# setCodeName")
        """ 종목코드에 해당하는 한글명을 codeNameLineEdit에 설정한다. """

        code = self.codeLineEdit.text()

        if code in self.codeList:
            codeName = self.kiwoom.getMasterCodeName(code)
            self.codeNameLineEdit.setText(codeName)

    def setAccountComboBox(self):
        self.kiwoom.log.info("# setAccountComboBox")
        """ accountComboBox에 계좌번호를 설정한다. """

        try:
            cnt = int(self.kiwoom.getLoginInfo("ACCOUNT_CNT"))
            accountList = self.kiwoom.getLoginInfo("ACCNO").split(';')
            self.accountComboBox.addItems(accountList[0:cnt])
        except (KiwoomConnectError, ParameterTypeError, ParameterValueError) as e:
            self.showDialog('Critical', e)

    def sendOrder(self):
        self.kiwoom.log.info("# sendOrder")
        """ 키움서버로 주문정보를 전송한다. """

        orderTypeTable = {'신규매수': 1, '신규매도': 2, '매수취소': 3, '매도취소': 4}
        hogaTypeTable = {'지정가': "00", '시장가': "03"}

        account = self.accountComboBox.currentText()
        orderType = orderTypeTable[self.orderTypeComboBox.currentText()]
        code = self.codeLineEdit.text()
        hogaType = hogaTypeTable[self.hogaTypeComboBox.currentText()]
        qty = self.qtySpinBox.value()
        price = self.priceSpinBox.value()

        try:
            self.kiwoom.sendOrder("수동주문", "0101", account, orderType, code, qty, price, hogaType, "")

        except (ParameterTypeError, KiwoomProcessingError) as e:
            self.showDialog('Critical', e)

    def inquiryBalance(self):
        self.kiwoom.log.info("# inquiryBalance")
        """ 예수금상세현황과 계좌평가잔고내역을 요청후 테이블에 출력한다. """

        self.inquiryTimer.stop()

        try:
            # 예수금상세현황요청
            self.kiwoom.setInputValue("계좌번호", self.accountComboBox.currentText())
            self.kiwoom.setInputValue("비밀번호", "0000")
            self.kiwoom.commRqData("예수금상세현황요청", "opw00001", 0, "2000")

            # 계좌평가잔고내역요청 - opw00018 은 한번에 20개의 종목정보를 반환
            self.kiwoom.setInputValue("계좌번호", self.accountComboBox.currentText())
            self.kiwoom.setInputValue("비밀번호", "0000")
            self.kiwoom.commRqData("계좌평가잔고내역요청", "opw00018", 0, "2000")

            while self.kiwoom.inquiry == '2':
                time.sleep(0.2)
                self.kiwoom.setInputValue("계좌번호", self.accountComboBox.currentText())
                self.kiwoom.setInputValue("비밀번호", "0000")
                self.kiwoom.commRqData("계좌평가잔고내역요청", "opw00018", 2, "2000")

        except (ParameterTypeError, ParameterValueError, KiwoomProcessingError) as e:
            self.showDialog('Critical', e)

        # accountEvaluationTable 테이블에 정보 출력
        item = QTableWidgetItem(self.kiwoom.opw00001Data)  # d+2추정예수금
        item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
        self.accountEvaluationTable.setItem(0, 0, item)

        for i in range(1, 6):
            item = QTableWidgetItem(self.kiwoom.opw00018Data['accountEvaluation'][i - 1])
            item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
            self.accountEvaluationTable.setItem(0, i, item)

        self.accountEvaluationTable.resizeRowsToContents()

        # stocksTable 테이블에 정보 출력
        cnt = len(self.kiwoom.opw00018Data['stocks'])
        self.stocksTable.setRowCount(cnt)

        for i in range(cnt):
            row = self.kiwoom.opw00018Data['stocks'][i]

            for j in range(len(row)):
                item = QTableWidgetItem(row[j])
                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignRight)
                self.stocksTable.setItem(i, j, item)

        self.stocksTable.resizeRowsToContents()

        # stock list
        for (_, code, volume, _, _, _, profit_ratio) in self.kiwoom.opw00018Data['stocks']: # keyList = ["종목명", "종목번호", "보유수량", "매입가", "현재가", "평가손익", "수익률(%)"]
            self.stockList[code] = (profit_ratio, volume)

        # 데이터 초기화
        self.kiwoom.opwDataReset()

        # inquiryTimer 재시작
        self.inquiryTimer.start(1000 * 10)

    # 경고창
    def showDialog(self, grade, error):
        self.kiwoom.log.info("# showDialog")
        gradeTable = {'Information': 1, 'Warning': 2, 'Critical': 3, 'Question': 4}

        dialog = QMessageBox()
        dialog.setIcon(gradeTable[grade])
        dialog.setText(error.msg)
        dialog.setWindowTitle(grade)
        dialog.setStandardButtons(QMessageBox.Ok)
        dialog.exec_()

    def setAutomatedStocks(self):
        self.kiwoom.log.info("# setAutomatedStocks")
        fileList = ["buy_list.txt", "sell_list.txt"]
        automatedStocks = []

        try:
            for file in fileList:
                # utf-8로 작성된 파일을
                # cp949 환경에서 읽기위해서 encoding 지정
                with open(file, 'rt', encoding='utf-8') as f:
                    stocksList = f.readlines()
                    automatedStocks += stocksList
        except Exception as e:
            e.msg = "setAutomatedStocks() 에러"
            self.showDialog('Critical', e)
            return

        # 테이블 행수 설정
        cnt = len(automatedStocks)
        self.automatedStocksTable.setRowCount(cnt)

        # 테이블에 출력
        for i in range(cnt):
            stocks = automatedStocks[i].split(';')

            for j in range(len(stocks)):
                if j == 1:
                    name = self.kiwoom.getMasterCodeName(stocks[j].rstrip())
                    item = QTableWidgetItem(name)
                else:
                    item = QTableWidgetItem(stocks[j].rstrip())

                item.setTextAlignment(Qt.AlignVCenter | Qt.AlignCenter)
                self.automatedStocksTable.setItem(i, j, item)

        self.automatedStocksTable.resizeRowsToContents()

    def automaticOrder(self):
        self.kiwoom.log.info("# automaticOrder")
        fileList = ["buy_list.txt", "sell_list.txt"]
        hogaTypeTable = {'지정가': "00", '시장가': "03"}
        account = self.accountComboBox.currentText()

        automatedStocks = self.buyStrategy()
        automatedStocks += self.sellStrategy()

        cnt = len(automatedStocks)

        print("automatedStocks: ", automatedStocks)
        print("cnt: ", cnt)

        # 매매할 종목이 없으면 종료
        if cnt == 0:
            return

        # 주문하기
        buyResult = []
        sellResult = []

        for i in range(cnt):
            stocks = automatedStocks[i].split(';')

            code = stocks[1]
            hoga = stocks[2]
            qty = stocks[3]
            price = stocks[4]

            try:
                if stocks[5].rstrip() == '매수전':
                    self.kiwoom.sendOrder("자동매수주문", "0101", account, 1, code, int(qty), int(price), hogaTypeTable[hoga],
                                          "")

                    # 주문 접수시
                    if self.kiwoom.orderNo:
                        buyResult += automatedStocks[i].replace("매수전", "매수주문완료")
                        self.todayBuyList.append(code)
                        self.kiwoom.orderNo = ""
                    # 주문 미접수시
                    else:
                        buyResult += automatedStocks[i]

                # 참고: 해당 종목을 현재도 보유하고 있다고 가정함.
                elif stocks[5].rstrip() == '매도전':
                    self.kiwoom.sendOrder("자동매도주문", "0101", account, 2, code, int(qty), int(price), hogaTypeTable[hoga],
                                          "")

                    # 주문 접수시
                    if self.kiwoom.orderNo:
                        sellResult += automatedStocks[i].replace("매도전", "매도주문완료")
                        self.kiwoom.orderNo = ""
                    # 주문 미접수시
                    else:
                        sellResult += automatedStocks[i]

            except (ParameterTypeError, KiwoomProcessingError) as e:
                self.showDialog('Critical', e)

            time.sleep(0.2)

        # 잔고및 보유종목 디스플레이 갱신
        self.inquiryBalance()

        # 결과저장하기
        for file, result in zip(fileList, [buyResult, sellResult]):
            with open(file, 'wt', encoding='utf-8') as f:
                for data in result:
                    f.write(data)

        self.setAutomatedStocks()

    def realtimeConditonStart(self):
        self.kiwoom.log.info("# realtimeConditionStart")
        """ 실시간 조건검색 시작 메서드 """
        try:
            self.kiwoom.getConditionLoad()

            for index in self.kiwoom.condition.keys():
                self.kiwoom.sendCondition("0156", self.kiwoom.condition[index], index, 1)

                # 조건식 하나만 테스트
                break

        except Exception as e:
            print(e)

    def buyStrategy(self):
        self.kiwoom.log.info("# BuyStrategy")
        """ 매수전략을 이용하여 매수할 종목 선정 """

        try:
            stockList = self.kiwoom.realConditionCodeList[0:]

            if len(stockList) == 0:
                return []

            # 매수할 종목 리스트
            codeList = []

            for code in stockList:
                if code not in self.todayBuyList:
                    order = "매수;{};시장가;10;0;매수전\n".format(code)
                    codeList.append(order)
                    # TODO: 매수전략 작성

            return codeList

        except Exception as e:
            print(e)

    def sellStrategy(self):
        self.kiwoom.log.info("# sellStrategy")
        """ 매도전략을 이용하여 매도할 종목 선정 """

        # TODO: 매도전략 작성
        try:
            # ["종목명", "종목코드", "보유수량", "매입가", "현재가", "평가손익", "수익률(%)"]

            if len(self.stockList) == 0:
                return []

            # 매수할 종목 리스트
            codeList = []

            for (code, (profit_ratio, volume))  in self.stockList.items():
                code = code[-6:]
                if (float(profit_ratio) < -2) or (float(profit_ratio) > 2):
                    order = "매도;{};시장가;{};0;매도전\n".format(code, volume)
                    codeList.append(order)
                    self.kiwoom.setRealRemove("0156", code)
                    # self.todayBuyList.remove(code)
            return codeList
        except Exception as e:
            print(e)
        return []

    def DDD(self):
        stock_list = ['000020', '000030', '000040', '000050', '000060', '000070', '000075', '000080', '000087' ]
        for stock in stock_list:
            self.kiwoom.setInputValue("종목코드", stock)
            self.kiwoom.setInputValue("기준일자", '2017-09-29')
            self.kiwoom.setInputValue("수정주가구분", '0')
            self.kiwoom.commRqData("주식일봉차트조회요청", "OPT10081", 0, "0615")
            for cnt in self.kiwoom.data:
                print("""replace into opt10081
                          (code, cur_price, volume, volume_price, date,
                          open_price, high_price, low_price, modify_gubun, modify_ratio,
                          big_gubun, small_gubun, code_inform, modify_event, before_close) values
                              ({}, {}, {}, {}, {},
                               {}, {}, {}, {}, {},
                               {}, {}, {}, {}, {})
                              """.format(stock, cnt[1], cnt[2], cnt[3], cnt[4],
                                    cnt[5], cnt[6], cnt[7], cnt[8], cnt[9],
                                    cnt[10], cnt[11], cnt[12], cnt[13], cnt[14]))
            time.sleep(0.5)
            while self.kiwoom.inquiry == '2':
                self.kiwoom.setInputValue("종목코드", stock)
                self.kiwoom.setInputValue("기준일자", '2017-09-29')
                self.kiwoom.setInputValue("수정주가구분", '0')
                self.kiwoom.commRqData("주식일봉차트조회요청", "OPT10081", 2, "0615")
                for cnt in self.kiwoom.data:
                    print("""replace into opt10081
                              (code, cur_price, volume, volume_price, date,
                              open_price, high_price, low_price, modify_gubun, modify_ratio,
                              big_gubun, small_gubun, code_inform, modify_event, before_close) values
                                  ({}, {}, {}, {}, {},
                                   {}, {}, {}, {}, {},
                                   {}, {}, {}, {}, {})
                                  """.format(stock, cnt[1], cnt[2], cnt[3], cnt[4],
                                             cnt[5], cnt[6], cnt[7], cnt[8], cnt[9],
                                             cnt[10], cnt[11], cnt[12], cnt[13], cnt[14]))
                time.sleep(0.5)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = MyWindow()
    sys.exit(app.exec_())
