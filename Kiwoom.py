# -*- encoding:utf-8 -*-
"""
Kiwoom 기본 클래스
"""

import sys, time
import logging
import logging.config
from pandas import DataFrame
from PyQt5.QAxContainer import QAxWidget
from PyQt5.QtCore import QEventLoop
from PyQt5.QtWidgets import QApplication
from Code.ReturnCode import ReturnCode
from Code.FidList import FidList
from Code.RealType import RealType


class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1")

        # Loop 변수
        # 비동기 방식으로 동작되는 이벤트를 동기화(순서대로 동작) 시킬 때
        self.loginLoop = None
        self.requestLoop = None
        self.orderLoop = None
        self.conditionLoop = None

        # 서버구분
        self.server = None

        # 조건식
        self.condition = None

        # 조건검색에 의한 종목코드 리스트
        self.realConditionCodeList = []

        # 에러
        self.error = None

        # 주문번호
        self.orderNo = ""

        # 조회
        self.inquiry = 0

        # 서버에서 받은 메시지
        self.msg = ""

        # 예수금 d+2
        self.opw00001Data = 0

        # 보유종목 정보
        self.opw00018Data = {'accountEvaluation': [], 'stocks': []}

        # signal & slot
        self.OnEventConnect.connect(self.eventConnect)
        self.OnReceiveMsg.connect(self.receiveMsg)
        self.OnReceiveTrData.connect(self.receiveTrData)
        self.OnReceiveConditionVer.connect(self.receiveConditionVer)
        self.OnReceiveTrCondition.connect(self.receiveTrCondition)
        self.OnReceiveRealData.connect(self.receiveRealData)
        self.OnReceiveChejanData.connect(self.receiveChejanData)
        self.OnReceiveRealCondition.connect(self.receiveRealCondition)

        # void OnReceiveInvestRealData(QString sRealKey);
        # QObject::connect(object, SIGNAL(OnReceiveTrData(QString, QString, QString, QString, QString, int, QString, QString, QString)), receiver, SLOT(someSlot(QString, QString, QString, QString, QString, int, QString, QString, QString)));
        self.OnReceiveInvestRealData.connect(self.receiveInvestRealData)

        # void exception(int code, QString source, QString disc, QString help);
        # QObject::connect(object, SIGNAL(exception(int, QString, QString, QString)), receiver, SLOT(someSlot(int, QString, QString, QString)));

        # void propertyChanged(QString name);
        # QObject::connect(object, SIGNAL(propertyChanged(QString)), receiver, SLOT(someSlot(QString)));

        # void signal(QString name, int argc, void* argv);
        # QObject::connect(object, SIGNAL(signal(QString, int, void *)), receiver, SLOT(someSlot(QString, int, void *)));

        # 로깅용 설정파일
        logging.config.fileConfig('logging.conf')
        self.log = logging.getLogger('Kiwoom')

    ###############################################################
    # 로깅용 메서드 정의                                               #
    ###############################################################

    def logger(origin):
        def wrapper(*args, **kwargs):
            args[0].log.debug('{} args - {}, kwargs - {}'.format(origin.__name__, args, kwargs))
            return origin(*args, **kwargs)

        return wrapper

    ###############################################################
    # 이벤트 정의                                                    #
    ###############################################################

    def eventConnect(self, returnCode):
        """
        통신 연결 상태 변경시 이벤트

        returnCode가 0이면 로그인 성공
        그 외에는 ReturnCode 클래스 참조.

        :param returnCode: int
        """
        self.log.info("<<eventConnect>>")
        self.log.debug("returnCode : {}".format(returnCode))

        try:
            if returnCode == ReturnCode.OP_ERR_NONE:

                self.server = self.getLoginInfo("GetServerGubun", True)

                if (self.server is None) and self.server != "1":
                    self.msg += "실서버 연결 성공" + "\r\n\r\n"

                else:
                    self.msg += "모의투자서버 연결 성공" + "\r\n\r\n"

            else:
                self.msg += "연결 끊김: 원인 - " + ReturnCode.CAUSE[returnCode] + "\r\n\r\n"

        except Exception as error:
            self.log.error('eventConnect {}'.format(error))

        finally:
            # commConnect() 메서드에 의해 생성된 루프를 종료시킨다.
            # 로그인 후, 통신이 끊길 경우를 대비해서 예외처리함.
            try:
                self.loginLoop.exit()
            except AttributeError:
                pass

    def receiveMsg(self, screenNo, requestName, trCode, msg):
        """
        수신 메시지 이벤트

        서버로 어떤 요청을 했을 때(로그인, 주문, 조회 등), 그 요청에 대한 처리내용을 전달해준다.

        :param screenNo: string - 화면번호(4자리, 사용자 정의, 서버에 조회나 주문을 요청할 때 이 요청을 구별하기 위한 키값)
        :param requestName: string - TR 요청명(사용자 정의)
        :param trCode: string
        :param msg: string - 서버로 부터의 메시지
        """
        self.log.info("<<receiveMsg>>")
        self.log.debug("screenNo, requestName, trCode, msg : {}, {}, {}, {}".format(screenNo, requestName, trCode, msg))
        self.msg += requestName + ": " + msg + "\r\n\r\n"

    def receiveTrData(self, screenNo, requestName, trCode, recordName, inquiry,
                      deprecated1, deprecated2, deprecated3, deprecated4):
        """
        TR 수신 이벤트

        조회요청 응답을 받거나 조회데이터를 수신했을 때 호출됩니다.
        requestName과 trCode는 commRqData()메소드의 매개변수와 매핑되는 값 입니다.
        조회데이터는 이 이벤트 메서드 내부에서 getCommData() 메서드를 이용해서 얻을 수 있습니다.

        :param screenNo: string - 화면번호(4자리)
        :param requestName: string - TR 요청명(commRqData() 메소드 호출시 사용된 requestName)
        :param trCode: string
        :param recordName: string
        :param inquiry: string - 조회('0': 남은 데이터 없음, '2': 남은 데이터 있음)
        """
        self.log.info("<<receiveTrData>>")
        self.log.debug(
            "screenNo, requestName, trCode, recordName, inquiry, deprecated1, deprecated2, deprecated3, deprecated4 : {0:5} {1:10} {2:5} {3:3}".format(
                screenNo, requestName, trCode, recordName, inquiry, deprecated1, deprecated2, deprecated3, deprecated4))
        # 주문번호와 주문루프
        self.orderNo = self.getCommData(trCode, requestName, 0, "주문번호")

        try:
            self.orderLoop.exit()
        except AttributeError:
            pass

        self.inquiry = inquiry

        if requestName == "관심종목정보요청":
            self.data = self.getCommDataEx(trCode, "관심종목정보")
            print(type(self.data))
            print(self.data)

            """ getCommData
            cnt = self.getRepeatCnt(trCode, requestName)

            for i in range(cnt):
                data = self.getCommData(trCode, requestName, i, "종목명")
                print(data)
            """

        elif requestName == "주식틱차트조회요청":
            self.data = self.getCommDataEx(trCode, "주식틱차트조회")

        elif requestName == "주식일봉차트조회요청":
            self.data = self.getCommDataEx(trCode, "주식일봉차트조회")

        elif requestName == "예수금상세현황요청":
            deposit = self.getCommData(trCode, requestName, 0, "d+2추정예수금")
            deposit = self.changeFormat(deposit)
            self.opw00001Data = deposit

        elif requestName == "계좌평가잔고내역요청":
            # 계좌 평가 정보
            accountEvaluation = []
            keyList = ["총매입금액", "총평가금액", "총평가손익금액", "총수익률(%)", "추정예탁자산"]

            for key in keyList:
                value = self.getCommData(trCode, requestName, 0, key)

                # if key.startswith("총수익률"):
                #     value = self.changeFormat(value, 1)
                # else:
                #     value = self.changeFormat(value)

                accountEvaluation.append(value)

            self.opw00018Data['accountEvaluation'] = accountEvaluation

            # 보유 종목 정보
            cnt = self.getRepeatCnt(trCode, requestName)
            keyList = ["종목명", "종목번호", "보유수량", "매입가", "현재가", "평가손익", "수익률(%)"]

            for i in range(cnt):
                stock = []

                for key in keyList:
                    value = self.getCommData(trCode, requestName, i, key)

                    if key.startswith("수익률"):
                        value = self.changeFormat(value, 2)
                    elif key != "종목번호" and key != "종목명":
                        value = self.changeFormat(value)

                    stock.append(value)

                self.opw00018Data['stocks'].append(stock)

        try:
            self.requestLoop.exit()
        except AttributeError:
            pass

    def receiveConditionVer(self, receive, msg):
        """
        getConditionLoad() 메서드의 조건식 목록 요청에 대한 응답 이벤트

        :param receive: int - 응답결과(1: 성공, 나머지 실패)
        :param msg: string - 메세지
        """
        self.log.info("<<receiveConditionVer>>")
        self.log.debug("receive, msg : ({}, {})".format(receive, msg))

        try:
            if not receive:
                return

            self.condition = self.getConditionNameList()
            print("조건식 개수: ", len(self.condition))

            for key in self.condition.keys():
                print("조건식: ", key, ": ", self.condition[key])
                print("key type: ", type(key))

        except Exception as e:
            print(e)

        finally:
            self.conditionLoop.exit()

    def receiveTrCondition(self, screenNo, codes, conditionName, conditionIndex, inquiry):
        """
        (1회성, 실시간) 종목 조건검색 요청시 발생되는 이벤트

        :param screenNo: string
        :param codes: string - 종목코드 목록(각 종목은 세미콜론으로 구분됨)
        :param conditionName: string - 조건식 이름
        :param conditionIndex: int - 조건식 인덱스
        :param inquiry: int - 조회구분(0: 남은데이터 없음, 2: 남은데이터 있음)
        """
        self.log.info("<<receiveTrCondition>>")
        self.log.debug(
            "screenNo, codes, conditionName, conditionIndex, inquiry : ({}, {}, {}, {}, {})".format(screenNo, codes,
                                                                                                    conditionName,
                                                                                                    conditionIndex,
                                                                                                    inquiry))
        print("[receiveTrCondition]")

        try:
            if codes == "":
                return

            codeList = codes.split(';')
            del codeList[-1]

            print(codeList)
            print("종목개수: ", len(codeList))

            self.realConditionCodeList += codeList

        finally:
            self.conditionLoop.exit()

    def receiveRealData(self, code, realType, realData):
        """
        실시간 데이터 수신 이벤트

        실시간 데이터를 수신할 때 마다 호출되며,
        setRealReg() 메서드로 등록한 실시간 데이터도 이 이벤트 메서드에 전달됩니다.
        getCommRealData() 메서드를 이용해서 실시간 데이터를 얻을 수 있습니다.

        :param code: string - 종목코드
        :param realType: string - 실시간 타입(KOA의 실시간 목록 참조)
        :param realData: string - 실시간 데이터 전문
        """
        self.log.info("<<receiveRealData>>")
        self.log.debug("code, realType, realData : ({}, {}, {})".format(code, realType, realData))
        try:
            if realType not in RealType.REALTYPE:
                return

            data = []

            if code != "":
                data.append(code)
                codeOrNot = code
            else:
                codeOrNot = realType

            for fid in sorted(RealType.REALTYPE[realType].keys()):
                value = self.getCommRealData(codeOrNot, fid)
                data.append(value)

            # TODO: DB에 저장
            self.log.debug(data)

        except Exception as e:
            self.log.error('{}'.format(e))

    def receiveChejanData(self, gubun, itemCnt, fidList):
        """
        주문 접수/확인 수신시 이벤트

        주문요청후 주문접수, 체결통보, 잔고통보를 수신할 때 마다 호출됩니다.

        :param gubun: string - 체결구분('0': 주문접수/주문체결, '1': 잔고통보, '3': 특이신호)
        :param itemCnt: int - fid의 갯수
        :param fidList: string - fidList 구분은 ;(세미콜론) 이다.
        """
        self.log.info("<<receiveChejanData>>")
        self.log.debug("gubun, itemCnt, fidList : ({}, {}, {})".format(gubun, itemCnt, fidList))

        fids = fidList.split(';')
        print("[receiveChejanData]")
        print("gubun: ", gubun, "itemCnt: ", itemCnt, "fidList: ", fidList)
        print("========================================")
        print("[ 구분: ", self.getChejanData(913) if '913' in fids else '잔고통보', "]")
        for fid in fids:
            print(FidList.CHEJAN[int(fid)] if int(fid) in FidList.CHEJAN else fid, ": ", self.getChejanData(int(fid)))
        print("========================================")

    def receiveRealCondition(self, code, event, conditionName, conditionIndex):
        """
        실시간 종목 조건검색 요청시 발생되는 이벤트

        :param code: string - 종목코드
        :param event: string - 이벤트종류("I": 종목편입, "D": 종목이탈)
        :param conditionName: string - 조건식 이름
        :param conditionIndex: string - 조건식 인덱스(여기서만 인덱스가 string 타입으로 전달됨)
        """
        self.log.info("<<receiveRealCondition>>")
        self.log.debug(
            "code, event, conditionName, conditionIndex : ({}, {}, {}, {})".format(code, event, conditionName,
                                                                                   conditionIndex))

        print("[receiveRealCondition]")

        print("종목코드: ", code)
        print("이벤트: ", "종목편입" if event == "I" else "종목이탈")

        if event == "I":
            self.realConditionCodeList.append(code)

    def receiveInvestRealData(self, sRealKey):
        print("ON receive invest Real Data")
        self.log.info("<<receiveInvestRealData>>")
        self.log.debug("sRealKey : ({})".format(sRealKey))

    ###############################################################
    # 메서드 정의: 로그인 관련 메서드                                    #
    ###############################################################

    def commConnect(self):
        """
        로그인을 시도합니다.

        수동 로그인일 경우, 로그인창을 출력해서 로그인을 시도.
        자동 로그인일 경우, 로그인창 출력없이 로그인 시도.
        """
        self.log.info("[commConnect]")

        self.dynamicCall("CommConnect()")
        self.loginLoop = QEventLoop()
        self.loginLoop.exec_()

    def getConnectState(self):
        """
        현재 접속상태를 반환합니다.

        반환되는 접속상태는 아래와 같습니다.
        0: 미연결, 1: 연결

        :return: int
        """
        self.log.info("[getConnectState]")

        state = self.dynamicCall("GetConnectState()")
        return state

    def getLoginInfo(self, tag, isConnectState=False):
        """
        사용자의 tag에 해당하는 정보를 반환한다.

        tag에 올 수 있는 값은 아래와 같다.
        ACCOUNT_CNT: 전체 계좌의 개수를 반환한다.
        ACCNO: 전체 계좌 목록을 반환한다. 계좌별 구분은 ;(세미콜론) 이다.
        USER_ID: 사용자 ID를 반환한다.
        USER_NAME: 사용자명을 반환한다.
        GetServerGubun: 접속서버 구분을 반환합니다.("1": 모의투자, 그외(빈 문자열포함): 실서버)

        :param tag: string
        :param isConnectState: bool - 접속상태을 확인할 필요가 없는 경우 True로 설정.
        :return: string
        """
        self.log.info("[getLoginInfo]")
        self.log.debug("tag, isConnectState : ({}, {})".format(tag, isConnectState))

        if not isConnectState:
            if not self.getConnectState():
                raise KiwoomConnectError()

        if not isinstance(tag, str):
            raise ParameterTypeError()

        if tag not in ['ACCOUNT_CNT', 'ACCNO', 'USER_ID', 'USER_NAME', 'GetServerGubun']:
            raise ParameterValueError()

        if tag == "GetServerGubun":
            info = self.getServerGubun()
        else:
            cmd = 'GetLoginInfo("{}")'.format(tag)
            info = self.dynamicCall(cmd)

        return info

    def getServerGubun(self):
        """
        서버구분 정보를 반환한다.
        리턴값이 "1"이면 모의투자 서버이고, 그 외에는 실서버(빈 문자열포함).

        :return: string
        """
        self.log.info("[getServerGubun]")

        ret = self.dynamicCall("KOA_Functions(QString, QString)", "GetServerGubun", "")
        return ret

    #################################################################
    # 메서드 정의: 조회 관련 메서드                                        #
    # 시세조회, 관심종목 조회, 조건검색 등 이들의 합산 조회 횟수가 1초에 5회까지 허용 #
    #################################################################

    def setInputValue(self, key, value):
        """
        TR 전송에 필요한 값을 설정한다.

        :param key: string - TR에 명시된 input 이름
        :param value: string - key에 해당하는 값
        """
        self.log.info("[setInputValue]")
        self.log.debug("key, value : ({}, {})".format(key, value))

        if not (isinstance(key, str) and isinstance(value, str)):
            raise ParameterTypeError()

        self.dynamicCall("SetInputValue(QString, QString)", key, value)

    def commRqData(self, requestName, trCode, inquiry, screenNo):
        """
        키움서버에 TR 요청을 한다.

        조회요청메서드이며 빈번하게 조회요청시, 시세과부하 에러값 -200이 리턴된다.

        :param requestName: string - TR 요청명(사용자 정의)
        :param trCode: string
        :param inquiry: int - 조회(0: 조회, 2: 남은 데이터 이어서 요청)
        :param screenNo: string - 화면번호(4자리)
        """
        self.log.info("[commRqData]")
        self.log.debug(
            "requestName, trCode, inquiry, screenNo : ({}, {}, {}, {})".format(requestName, trCode, inquiry, screenNo))

        if not self.getConnectState():
            raise KiwoomConnectError()

        if not (isinstance(requestName, str)
                and isinstance(trCode, str)
                and isinstance(inquiry, int)
                and isinstance(screenNo, str)):
            raise ParameterTypeError()

        returnCode = self.dynamicCall("CommRqData(QString, QString, int, QString)", requestName, trCode, inquiry,
                                      screenNo)

        if returnCode != ReturnCode.OP_ERR_NONE:
            raise KiwoomProcessingError("commRqData(): " + ReturnCode.CAUSE[returnCode])

        # 루프 생성: receiveTrData() 메서드에서 루프를 종료시킨다.
        self.requestLoop = QEventLoop()
        self.requestLoop.exec_()

    def getCommData(self, trCode, requestName, index, key):
        """
        데이터 획득 메서드

        receiveTrData() 이벤트 메서드가 호출될 때, 그 안에서 조회데이터를 얻어오는 메서드입니다.

        :param trCode: string
        :param requestName: string - TR 요청명(commRqData() 메소드 호출시 사용된 requestName)
        :param index: int
        :param key: string - 수신 데이터에서 얻고자 하는 값의 키(출력항목이름)
        :return: string
        """
        self.log.info("[getCommData]")
        self.log.debug("trCode, requestName, index, key : ({}, {}, {}, {})".format(trCode, requestName, index, key))

        if not (isinstance(trCode, str)
                and isinstance(requestName, str)
                and isinstance(index, int)
                and isinstance(key, str)):
            raise ParameterTypeError()

        data = self.dynamicCall("GetCommData(QString, QString, int, QString)",
                                trCode, requestName, index, key)
        return data.strip()

    def getRepeatCnt(self, trCode, requestName):
        """
        서버로 부터 전달받은 데이터의 갯수를 리턴합니다.(멀티데이터의 갯수)

        receiveTrData() 이벤트 메서드가 호출될 때, 그 안에서 사용해야 합니다.

        키움 OpenApi+에서는 데이터를 싱글데이터와 멀티데이터로 구분합니다.
        싱글데이터란, 서버로 부터 전달받은 데이터 내에서, 중복되는 키(항목이름)가 하나도 없을 경우.
        예를들면, 데이터가 '종목코드', '종목명', '상장일', '상장주식수' 처럼 키(항목이름)가 중복되지 않는 경우를 말합니다.
        반면 멀티데이터란, 서버로 부터 전달받은 데이터 내에서, 일정 간격으로 키(항목이름)가 반복될 경우를 말합니다.
        예를들면, 10일간의 일봉데이터를 요청할 경우 '종목코드', '일자', '시가', '고가', '저가' 이러한 항목이 10번 반복되는 경우입니다.
        이러한 멀티데이터의 경우 반복 횟수(=데이터의 갯수)만큼, 루프를 돌면서 처리하기 위해 이 메서드를 이용하여 멀티데이터의 갯수를 얻을 수 있습니다.

        :param trCode: string
        :param requestName: string - TR 요청명(commRqData() 메소드 호출시 사용된 requestName)
        :return: int
        """
        self.log.info("[getRepeatCnt]")
        self.log.debug("trCode, requestName : ({}, {})".format(trCode, requestName))

        if not (isinstance(trCode, str)
                and isinstance(requestName, str)):
            raise ParameterTypeError()

        count = self.dynamicCall("GetRepeatCnt(QString, QString)", trCode, requestName)
        return count

    def getCommDataEx(self, trCode, multiDataName):
        """
        멀티데이터 획득 메서드

        receiveTrData() 이벤트 메서드가 호출될 때, 그 안에서 사용해야 합니다.

        :param trCode: string
        :param multiDataName: string - KOA에 명시된 멀티데이터명
        :return: list - 중첩리스트
        """
        self.log.info("[getCommDataEx]")
        self.log.debug("trCode, multiDataName : ({}, {})".format(trCode, multiDataName))

        if not (isinstance(trCode, str)
                and isinstance(multiDataName, str)):
            raise ParameterTypeError()

        data = self.dynamicCall("GetCommDataEx(QString, QString)", trCode, multiDataName)
        return data

    def commKwRqData(self, codes, inquiry, codeCount, requestName, screenNo, typeFlag=0):
        """
        복수종목조회 메서드(관심종목조회 메서드라고도 함).

        이 메서드는 setInputValue() 메서드를 이용하여, 사전에 필요한 값을 지정하지 않는다.
        단지, 메서드의 매개변수에서 직접 종목코드를 지정하여 호출하며,
        데이터 수신은 receiveTrData() 이벤트에서 아래 명시한 항목들을 1회 수신하며,
        이후 receiveRealData() 이벤트를 통해 실시간 데이터를 얻을 수 있다.

        복수종목조회 TR 코드는 OPTKWFID 이며, 요청 성공시 아래 항목들의 정보를 얻을 수 있다.

        종목코드, 종목명, 현재가, 기준가, 전일대비, 전일대비기호, 등락율, 거래량, 거래대금,
        체결량, 체결강도, 전일거래량대비, 매도호가, 매수호가, 매도1~5차호가, 매수1~5차호가,
        상한가, 하한가, 시가, 고가, 저가, 종가, 체결시간, 예상체결가, 예상체결량, 자본금,
        액면가, 시가총액, 주식수, 호가시간, 일자, 우선매도잔량, 우선매수잔량,우선매도건수,
        우선매수건수, 총매도잔량, 총매수잔량, 총매도건수, 총매수건수, 패리티, 기어링, 손익분기,
        잔본지지, ELW행사가, 전환비율, ELW만기일, 미결제약정, 미결제전일대비, 이론가,
        내재변동성, 델타, 감마, 쎄타, 베가, 로

        :param codes: string - 한번에 100종목까지 조회가능하며 종목코드사이에 세미콜론(;)으로 구분.
        :param inquiry: int - api 문서는 bool 타입이지만, int로 처리(0: 조회, 1: 남은 데이터 이어서 조회)
        :param codeCount: int - codes에 지정한 종목의 갯수.
        :param requestName: string
        :param screenNo: string
        :param typeFlag: int - 주식과 선물옵션 구분(0: 주식, 3: 선물옵션), 주의: 매개변수의 위치를 맨 뒤로 이동함.
        :return: list - 중첩 리스트 [[종목코드, 종목명 ... 종목 정보], [종목코드, 종목명 ... 종목 정보]]
        """
        self.log.info("[commKwRqData]")
        self.log.debug(
            "codes, inquiry, codeCount, requestName, screenNo, typeFlag : ({}, {}, {}, {}, {}, {})".format(codes,
                                                                                                           inquiry,
                                                                                                           codeCount,
                                                                                                           requestName,
                                                                                                           screenNo,
                                                                                                           typeFlag))

        if not self.getConnectState():
            raise KiwoomConnectError()

        if not (isinstance(codes, str)
                and isinstance(inquiry, int)
                and isinstance(codeCount, int)
                and isinstance(requestName, str)
                and isinstance(screenNo, str)
                and isinstance(typeFlag, int)):
            raise ParameterTypeError()

        returnCode = self.dynamicCall("CommKwRqData(QString, QBoolean, int, int, QString, QString)",
                                      codes, inquiry, codeCount, typeFlag, requestName, screenNo)

        if returnCode != ReturnCode.OP_ERR_NONE:
            raise KiwoomProcessingError("commKwRqData(): " + ReturnCode.CAUSE[returnCode])

        # 루프 생성: receiveTrData() 메서드에서 루프를 종료시킨다.
        self.requestLoop = QEventLoop()
        self.requestLoop.exec_()

    ###############################################################
    # 메서드 정의: 실시간 데이터 처리 관련 메서드                           #
    ###############################################################

    def disconnectRealData(self, screenNo):
        """
        해당 화면번호로 설정한 모든 실시간 데이터 요청을 제거합니다.

        receiveTrData,, 이벤트내에서 호출해주시기 바랍니다.
        화면을 종료할 때 반드시 이 메서드를 호출해야 합니다.

        :param screenNo: string
        """
        self.log.info("[disconnectRealData]")
        self.log.debug("screenNo : ({})".format(screenNo))

        if not self.getConnectState():
            raise KiwoomConnectError()

        if not isinstance(screenNo, str):
            raise ParameterTypeError()

        self.dynamicCall("DisconnectRealData(QString)", screenNo)

    def getCommRealData(self, code, fid):
        """
        실시간 데이터 획득 메서드

        이 메서드는 반드시 receiveRealData() 이벤트 메서드가 호출될 때, 그 안에서 사용해야 합니다.

        :param code: string - 종목코드
        :param fid: - 실시간 타입에 포함된 fid
        :return: string - fid에 해당하는 데이터
        """
        self.log.info("[getCommRealData]")
        self.log.debug("code, fid : ({}, {})".format(code, fid))

        if not (isinstance(code, str)
                and isinstance(fid, int)):
            raise ParameterTypeError()

        value = self.dynamicCall("GetCommRealData(QString, int)", code, fid)

        return value

    def setRealReg(self, screenNo, codes, fids, realRegType):
        """
        실시간 데이터 요청 메서드

        종목코드와 fid 리스트를 이용해서 실시간 데이터를 요청하는 메서드입니다.
        한번에 등록 가능한 종목과 fid 갯수는 100종목, 100개의 fid 입니다.
        실시간등록타입을 0으로 설정하면, 첫 실시간 데이터 요청을 의미하며
        실시간등록타입을 1로 설정하면, 추가등록을 의미합니다.

        실시간 데이터는 실시간 타입 단위로 receiveRealData() 이벤트로 전달되기 때문에,
        이 메서드에서 지정하지 않은 fid 일지라도, 실시간 타입에 포함되어 있다면, 데이터 수신이 가능하다.

        :param screenNo: string
        :param codes: string - 종목코드 리스트(종목코드;종목코드;...)
        :param fids: string - fid 리스트(fid;fid;...)
        :param realRegType: string - 실시간등록타입(0: 첫 등록, 1: 추가 등록)
        """
        self.log.info("[setRealReg]")
        self.log.debug(
            "screenNo, codes, fids, realRegType : ({}, {}, {}, {})".format(screenNo, codes, fids, realRegType))

        if not self.getConnectState():
            raise KiwoomConnectError()

        if not (isinstance(screenNo, str)
                and isinstance(codes, str)
                and isinstance(fids, str)
                and isinstance(realRegType, str)):
            raise ParameterTypeError()

        self.dynamicCall("SetRealReg(QString, QString, QString, QString)",
                         screenNo, codes, fids, realRegType)

    def setRealRemove(self, screenNo, code):
        """
        실시간 데이터 중지 메서드

        setRealReg() 메서드로 등록한 종목만, 이 메서드를 통해 실시간 데이터 받기를 중지 시킬 수 있습니다.

        :param screenNo: string - 화면번호 또는 ALL 키워드 사용가능
        :param code: string - 종목코드 또는 ALL 키워드 사용가능
        """
        self.log.info("[setRealRemove]")
        self.log.debug("screenNo, code : ({}, {})".format(screenNo, code))

        if not self.getConnectState():
            raise KiwoomConnectError()

        if not (isinstance(screenNo, str)
                and isinstance(code, str)):
            raise ParameterTypeError()

        self.dynamicCall("SetRealRemove(QString, QString)", screenNo, code)

    ###############################################################
    # 메서드 정의: 조건검색 관련 메서드
    ###############################################################

    def getConditionLoad(self):
        """ 조건식 목록 요청 메서드 """
        self.log.info("[getConditionLoad]")

        if not self.getConnectState():
            raise KiwoomConnectError()

        isLoad = self.dynamicCall("GetConditionLoad()")

        # 요청 실패시
        if not isLoad:
            raise KiwoomProcessingError("getConditionLoad(): 조건식 요청 실패")

        # receiveConditionVer() 이벤트 메서드에서 루프 종료
        self.conditionLoop = QEventLoop()
        self.conditionLoop.exec_()

    def getConditionNameList(self):
        """
        조건식 획득 메서드

        조건식을 딕셔너리 형태로 반환합니다.
        이 메서드는 반드시 receiveConditionVer() 이벤트 메서드안에서 사용해야 합니다.

        :return: dict - {인덱스:조건명, 인덱스:조건명, ...}
        """
        self.log.info("[getConditionNameList]")

        data = self.dynamicCall("GetConditionNameList()")

        if data == "":
            raise KiwoomProcessingError("getConditionNameList(): 사용자 조건식이 없습니다.")

        conditionList = data.split(';')
        del conditionList[-1]

        conditionDictionary = {}

        for condition in conditionList:
            key, value = condition.split('^')
            conditionDictionary[int(key)] = value

        return conditionDictionary

    def sendCondition(self, screenNo, conditionName, conditionIndex, isRealTime):
        """
        종목 조건검색 요청 메서드

        이 메서드로 얻고자 하는 것은 해당 조건에 맞는 종목코드이다.
        해당 종목에 대한 상세정보는 setRealReg() 메서드로 요청할 수 있다.
        요청이 실패하는 경우는, 해당 조건식이 없거나, 조건명과 인덱스가 맞지 않거나, 조회 횟수를 초과하는 경우 발생한다.

        조건검색에 대한 결과는
        1회성 조회의 경우, receiveTrCondition() 이벤트로 결과값이 전달되며
        실시간 조회의 경우, receiveTrCondition()과 receiveRealCondition() 이벤트로 결과값이 전달된다.

        :param screenNo: string
        :param conditionName: string - 조건식 이름
        :param conditionIndex: int - 조건식 인덱스
        :param isRealTime: int - 조건검색 조회구분(0: 1회성 조회, 1: 실시간 조회)
        """
        self.log.info("[sendCondition]")
        self.log.debug(
            "screenNo, conditionName, conditionIndex, isRealTime : ({}, {}, {}, {})".format(screenNo, conditionName,
                                                                                            conditionIndex, isRealTime))

        if not self.getConnectState():
            raise KiwoomConnectError()

        if not (isinstance(screenNo, str)
                and isinstance(conditionName, str)
                and isinstance(conditionIndex, int)
                and isinstance(isRealTime, int)):
            raise ParameterTypeError()

        isRequest = self.dynamicCall("SendCondition(QString, QString, int, int",
                                     screenNo, conditionName, conditionIndex, isRealTime)

        if not isRequest:
            raise KiwoomProcessingError("sendCondition(): 조건검색 요청 실패")

        # receiveTrCondition() 이벤트 메서드에서 루프 종료
        self.conditionLoop = QEventLoop()
        self.conditionLoop.exec_()

    def sendConditionStop(self, screenNo, conditionName, conditionIndex):
        """ 종목 조건검색 중지 메서드 """
        self.log.info("[sendConditionStop]")
        self.log.debug(
            "screenNo, conditionName, conditionIndex : ({}, {}, {})".format(screenNo, conditionName, conditionIndex))

        if not self.getConnectState():
            raise KiwoomConnectError()

        if not (isinstance(screenNo, str)
                and isinstance(conditionName, str)
                and isinstance(conditionIndex, int)):
            raise ParameterTypeError()

        self.dynamicCall("SendConditionStop(QString, QString, int)", screenNo, conditionName, conditionIndex)

    ###############################################################
    # 메서드 정의: 주문과 잔고처리 관련 메서드                              #
    # 1초에 5회까지 주문 허용                                          #
    ###############################################################

    def sendOrder(self, requestName, screenNo, accountNo, orderType, code, qty, price, hogaType, originOrderNo):

        """
        주식 주문 메서드

        sendOrder() 메소드 실행시,
        OnReceiveMsg, OnReceiveTrData, OnReceiveChejanData 이벤트가 발생한다.
        이 중, 주문에 대한 결과 데이터를 얻기 위해서는 OnReceiveChejanData 이벤트를 통해서 처리한다.
        OnReceiveTrData 이벤트를 통해서는 주문번호를 얻을 수 있는데, 주문후 이 이벤트에서 주문번호가 ''공백으로 전달되면,
        주문접수 실패를 의미한다.

        :param requestName: string - 주문 요청명(사용자 정의)
        :param screenNo: string - 화면번호(4자리)
        :param accountNo: string - 계좌번호(10자리)
        :param orderType: int - 주문유형(1: 신규매수, 2: 신규매도, 3: 매수취소, 4: 매도취소, 5: 매수정정, 6: 매도정정)
        :param code: string - 종목코드
        :param qty: int - 주문수량
        :param price: int - 주문단가
        :param hogaType: string - 거래구분(00: 지정가, 03: 시장가, 05: 조건부지정가, 06: 최유리지정가, 그외에는 api 문서참조)
        :param originOrderNo: string - 원주문번호(신규주문에는 공백, 정정및 취소주문시 원주문번호르 입력합니다.)
        """
        self.log.info("[sendOrder]")
        self.log.debug(
            "requestName, screenNo, accountNo, orderType, code, qty, price, hogaType, originOrderNo : ({}, {}, {}, {}, {}, {}, {}, {}, {})".format(
                requestName, screenNo, accountNo, orderType, code, qty, price, hogaType, originOrderNo))

        if not self.getConnectState():
            raise KiwoomConnectError()

        if not (isinstance(requestName, str)
                and isinstance(screenNo, str)
                and isinstance(accountNo, str)
                and isinstance(orderType, int)
                and isinstance(code, str)
                and isinstance(qty, int)
                and isinstance(price, int)
                and isinstance(hogaType, str)
                and isinstance(originOrderNo, str)):
            raise ParameterTypeError()

        returnCode = self.dynamicCall("SendOrder(QString, QString, QString, int, QString, int, int, QString, QString)",
                                      [requestName, screenNo, accountNo, orderType, code, qty, price, hogaType,
                                       originOrderNo])

        if returnCode != ReturnCode.OP_ERR_NONE:
            raise KiwoomProcessingError("sendOrder(): " + ReturnCode.CAUSE[returnCode])

        # receiveTrData() 에서 루프종료
        self.orderLoop = QEventLoop()
        self.orderLoop.exec_()

    def getChejanData(self, fid):
        """
        주문접수, 주문체결, 잔고정보를 얻어오는 메서드

        이 메서드는 receiveChejanData() 이벤트 메서드가 호출될 때 그 안에서 사용해야 합니다.

        :param fid: int
        :return: string
        """
        self.log.info("[getChejanData]")
        self.log.debug("fid : ({})".format(fid))

        if not isinstance(fid, int):
            raise ParameterTypeError()

        cmd = 'GetChejanData("{}")'.format(fid)
        data = self.dynamicCall(cmd)
        return data

    ###############################################################
    # 기타 메서드 정의                                                #
    ###############################################################

    def getCodeListByMarket(self, market):
        """
        시장 구분에 따른 종목코드의 목록을 List로 반환한다.

        market에 올 수 있는 값은 아래와 같다.
        '0': 장내, '3': ELW, '4': 뮤추얼펀드, '5': 신주인수권, '6': 리츠, '8': ETF, '9': 하이일드펀드, '10': 코스닥, '30': 제3시장

        :param market: string
        :return: List
        """
        self.log.info("[getCodeListByMarket]")
        self.log.debug("market : ({})".format(market))

        if not self.getConnectState():
            raise KiwoomConnectError()

        if not isinstance(market, str):
            raise ParameterTypeError()

        if market not in ['0', '3', '4', '5', '6', '8', '9', '10', '30']:
            raise ParameterValueError()

        cmd = 'GetCodeListByMarket("{}")'.format(market)
        codeList = self.dynamicCall(cmd)
        return codeList[:-1].split(';')

    def getCodeList(self, *market):
        """
        여러 시장의 종목코드를 List 형태로 반환하는 헬퍼 메서드.

        :param market: Tuple - 여러 개의 문자열을 매개변수로 받아 Tuple로 처리한다.
        :return: List
        """
        self.log.info("[getCodeList]")
        self.log.debug("*market, market : ({}, {})".format(*market, market))

        codeList = []

        for m in market:
            tmpList = self.getCodeListByMarket(m)
            codeList += tmpList

        return codeList

    def getMasterCodeName(self, code):
        """
        종목코드의 한글명을 반환한다.

        :param code: string - 종목코드
        :return: string - 종목코드의 한글명
        """
        self.log.info("[getMasterCodeName]")
        self.log.debug("code : ({})".format(code))

        if not self.getConnectState():
            raise KiwoomConnectError()

        if not isinstance(code, str):
            raise ParameterTypeError()

        cmd = 'GetMasterCodeName("{}")'.format(code)
        name = self.dynamicCall(cmd)
        return name

    def changeFormat(self, data, percent=0):
        self.log.info("[changeFormat]")
        self.log.debug("data, percent : ({})".format(data, percent))

        if percent == 0:
            d = int(data)
            formatData = '{:-,d}'.format(d)

        elif percent == 1:
            f = int(data) / 100
            formatData = '{:-,.2f}'.format(f)

        elif percent == 2:
            f = float(data)
            formatData = '{:-,.2f}'.format(f)

        return formatData

    def opwDataReset(self):
        """ 잔고 및 보유종목 데이터 초기화 """
        self.log.info("[opwDataReset]")

        self.opw00001Data = 0
        self.opw00018Data = {'accountEvaluation': [], 'stocks': []}


class ParameterTypeError(Exception):
    """ 파라미터 타입이 일치하지 않을 경우 발생하는 예외 """

    def __init__(self, msg="파라미터 타입이 일치하지 않습니다."):
        self.msg = msg

    def __str__(self):
        return self.msg


class ParameterValueError(Exception):
    """ 파라미터로 사용할 수 없는 값을 사용할 경우 발생하는 예외 """

    def __init__(self, msg="파라미터로 사용할 수 없는 값 입니다."):
        self.msg = msg

    def __str__(self):
        return self.msg


class KiwoomProcessingError(Exception):
    """ 키움에서 처리실패에 관련된 리턴코드를 받았을 경우 발생하는 예외 """

    def __init__(self, msg="처리 실패"):
        self.msg = msg

    def __str__(self):
        return self.msg

    def __repr__(self):
        return self.msg


class KiwoomConnectError(Exception):
    """ 키움서버에 로그인 상태가 아닐 경우 발생하는 예외 """

    def __init__(self, msg="로그인 여부를 확인하십시오"):
        self.msg = msg

    def __str__(self):
        return self.msg


if __name__ == "__main__":
    """ 조건검색 테스트 코드 """

    app = QApplication(sys.argv)

    try:
        kiwoom = Kiwoom()
        kiwoom.commConnect()

        server = kiwoom.getServerGubun()
        print("server: ", server)
        print("type: ", type(server))

        if len(server) == 0 or server != "1":
            print("실서버 입니다.")

        else:
            print("모의투자 서버입니다.")
    except Exception as e:
        print(e)

    sys.exit(app.exec_())
