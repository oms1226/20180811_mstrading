# coding:utf-8
'''
주식 틱차트 조회 요청
'''

class OPT10079:
    def receiveTrData(self, screenNo, requestName, trCode, recordName, inquiry, deprecated1, deprecated2, deprecated3, deprecated4):
        # getCommDataEx로 한번에 받아오는 방법
        data = self.getCommDataEx(trCode, "주식틱차트조회")

        colName = ['현재가', '거래량', '체결시간', '시가', '고가', '저가', '수정주가구분', '수정비율', '대업종구분', '소업종구분', '종목정보', '수정주가이벤트', '전일종가']

        self.data = DataFrame(data, columns=colName)

        print(type(self.data))
        print(self.data.head(5))

        """ commGetData
        cnt = self.getRepeatCnt(trCode, requestName)

        for i in range(cnt):
            date = self.commGetData(trCode, "", requestName, i, "일자")
            open = self.commGetData(trCode, "", requestName, i, "시가")
            high = self.commGetData(trCode, "", requestName, i, "고가")
            low = self.commGetData(trCode, "", requestName, i, "저가")
            close = self.commGetData(trCode, "", requestName, i, "현재가")
            print(date, ": ", open, ' ', high, ' ', low, ' ', close)
        """