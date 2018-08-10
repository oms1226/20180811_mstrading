# -*- encoding:utf8 -*-
class OPW00018():
    def receiveTrData(self, screenNo, requestName, trCode, recordName, inquiry,
                      deprecated1, deprecated2, deprecated3, deprecated4):
        # 계좌 평가 정보
        accountEvaluation = []
        keyList = ["총매입금액", "총평가금액", "총평가손익금액", "총수익률(%)", "추정예탁자산"]

        for key in keyList:
            value = self.getCommData(trCode, requestName, 0, key)

            if key.startswith("총수익률"):
                value = self.changeFormat(value, 1)
            else:
                value = self.changeFormat(value)

            accountEvaluation.append(value)

        self.opw00018Data['accountEvaluation'] = accountEvaluation

        # 보유 종목 정보
        cnt = self.getRepeatCnt(trCode, requestName)
        keyList = ["종목명", "보유수량", "매입가", "현재가", "평가손익", "수익률(%)"]

        for i in range(cnt):
            stock = []

            for key in keyList:
                value = self.getCommData(trCode, requestName, i, key)

                if key.startswith("수익률"):
                    value = self.changeFormat(value, 2)
                elif key != "종목명":
                    value = self.changeFormat(value)

                stock.append(value)

            self.opw00018Data['stocks'].append(stock)
