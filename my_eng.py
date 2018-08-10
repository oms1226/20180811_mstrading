# coding=utf-8
"""
작성중인..호출 프로그램
"""

import time, sys
from Kiwoom_sungwon import *

import MySQLdb
from PyQt5.QtWidgets import QApplication

ACTION = ["주식틱차트조회요청"]
TODAY = '20170626'
TARGET = ["코스피,코스닥"]
if __name__ == '__main__':
    print("START", sys.argv)
    app = QApplication(sys.argv)

    try:
        kiwoom = Kiwoom_sungwon()
        kiwoom.commConnect()
        # 코스피 종목리스트 가져오기
        if '코스피' in TARGET:
            code_list = kiwoom.getCodeList("0")
        elif '코스닥' in TARGET:
            code_list = kiwoom.getCodeList("10")
        elif '코스피,코스닥' in TARGET:
            code_list = kiwoom.getCodeList("0", "10")
        else:
            sys.exit(app.exec_())

        #code_list = ['900050', '000020', '000030', '000040', '000050']
        print(code_list)
        print(len(code_list))
        # DB 접속
        conn = MySQLdb.connect(host='localhost', user='pyadmin', password='password', db='pystock', charset='utf8',
                               port=3390)
        curs = conn.cursor()
        if "주식기본정보요청" in ACTION:
            for code in code_list:
                print(code)
                kiwoom.setInputValue("종목코드", code)
                kiwoom.commRqData("주식기본정보요청", "OPT10001", 0, "0001")
                for cnt in kiwoom.data:
                    curs.execute("""replace into opt10079
                                  (code, cur_price, volume, date, open,
                                  high_price, low_price, modify_gubun, modify_ratio, big_gubun,
                                  small_gubun, stock_inform, modify_event, before_close) values
                                  (%s, %s, %s, %s, %s,
                                   %s, %s, %s, %s, %s,
                                   %s, %s, %s, %s)
                                  """, (code, cnt[0], cnt[1], cnt[2], cnt[3],
                                        cnt[4], cnt[5], cnt[6], cnt[7], cnt[8],
                                        cnt[9], cnt[10], cnt[11], cnt[12]))
                time.sleep(0.5)
        if "업종별투자자순매수요청" in ACTION:  # opt10051
            pass
        if "업종별주가요청" in ACTION:  # OPT20002
            pass
        if "주식틱차트조회요청" in ACTION:
            for code in code_list:
                print(code)
                curs.execute("select * from opt10079 where code=%s and date=%s", (code, TODAY))
                if curs.fetchall():
                    continue
                kiwoom.setInputValue("종목코드", code)
                kiwoom.setInputValue("틱범위", '1')
                kiwoom.setInputValue("수정주가구분", '0')
                kiwoom.commRqData("주식틱차트조회요청", "OPT10079", 0, "1000")
                for cnt in kiwoom.data:
                    curs.execute("""replace into opt10079
                                  (code, cur_price, volume, date, open_price,
                                  high_price, low_price, modify_gubun, modify_ratio, big_gubun,
                                  small_gubun, stock_inform, modify_event, before_close) values
                                  (%s, %s, %s, %s, %s,
                                   %s, %s, %s, %s, %s,
                                   %s, %s, %s, %s)
                                  """, (code, cnt[0], cnt[1], cnt[2], cnt[3],
                                        cnt[4], cnt[5], cnt[6], cnt[7], cnt[8],
                                        cnt[9], cnt[10], cnt[11], cnt[12]))
                time.sleep(0.1)
                conn.commit()

                while kiwoom.inquiry == '2':
                    kiwoom.setInputValue("종목코드", code)
                    kiwoom.setInputValue("틱범위", '1')
                    kiwoom.setInputValue("수정주가구분", '0')
                    kiwoom.commRqData("주식틱차트조회요청", "OPT10079", 0, "1000")
                    for cnt in kiwoom.data:
                        curs.execute("""replace into opt10079
                                      (code, cur_price, volume, date, open_price,
                                      high_price, low_price, modify_gubun, modify_ratio, big_gubun,
                                      small_gubun, stock_inform, modify_event, before_close) values
                                      (%s, %s, %s, %s, %s,
                                       %s, %s, %s, %s, %s,
                                       %s, %s, %s, %s)
                                      """, (code, cnt[0], cnt[1], cnt[2], cnt[3],
                                            cnt[4], cnt[5], cnt[6], cnt[7], cnt[8],
                                            cnt[9], cnt[10], cnt[11], cnt[12]))
                    time.sleep(0.1)
                    conn.commit()
        if "주식일봉차트조회요청" in ACTION:
            # opt10081
            for code in code_list:
                print(code)
                curs.execute("select date from opt10081 where code=%s and date=%s", (code, TODAY))
                if curs.fetchall():
                    continue
                kiwoom.setInputValue("종목코드", code)
                kiwoom.setInputValue("기준일자", TODAY)
                kiwoom.setInputValue("수정주가구분", '0')
                kiwoom.commRqData("주식일봉차트조회요청", "OPT10081", 0, "1000")
                for cnt in kiwoom.data:
                    curs.execute("""replace into opt10081
                              (code, cur_price, volume, volume_price, date,
                              open, high, low, modify_gubun, modify_ratio,
                              big_gubun, small_gubun, code_inform, modify_event, before_close) values
                                  (%s, %s, %s, %s, %s,
                                   %s, %s, %s, %s, %s,
                                   %s, %s, %s, %s, %s)
                                  """, (code, cnt[0], cnt[1], cnt[2], cnt[3],
                                        cnt[4], cnt[5], cnt[6], cnt[7], cnt[8],
                                        cnt[9], cnt[10], cnt[11], cnt[12]))
                time.sleep(0.5)
                while kiwoom.inquiry == '2':
                    kiwoom.setInputValue("종목코드", code)
                    kiwoom.setInputValue("기준일자", TODAY)
                    kiwoom.setInputValue("수정주가구분", '0')
                    kiwoom.commRqData("주식일봉차트조회요청", "OPT10081", 2, "1000")
                    for cnt in kiwoom.data:
                        curs.execute("""replace into opt10081
                                  (code, cur_price, volume, volume_price, date,
                                  open, high, low, modify_gubun, modify_ratio,
                                  big_gubun, small_gubun, code_inform, modify_event, before_close) values
                                      (%s, %s, %s, %s, %s,
                                       %s, %s, %s, %s, %s,
                                       %s, %s, %s, %s, %s)
                                      """, (code, cnt[0], cnt[1], cnt[2], cnt[3],
                                            cnt[4], cnt[5], cnt[6], cnt[7], cnt[8],
                                            cnt[9], cnt[10], cnt[11], cnt[12]))
                    time.sleep(0.5)
                conn.commit()

    except Exception as e:
        print(e)

    print("END")

    sys.exit(app.exec_())

