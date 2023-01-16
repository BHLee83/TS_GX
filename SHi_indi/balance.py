from PyQt5.QAxContainer import *



class Balance():
    def __init__(self, instInterface) -> None:
        self.instInterface = instInterface
        self.strCurrentAcnt = ''

        # 일반 TR OCX
        self.IndiTR = QAxWidget("GIEXPERTCONTROL.GiExpertControlCtrl.1")
        self.IndiReal = QAxWidget("GIEXPERTCONTROL.GiExpertControlCtrl.1")
        self.IndiTR.ReceiveData.connect(self.ReceiveData)
        self.IndiReal.ReceiveRTData.connect(self.ReceiveRTData)
        self.IndiTR.ReceiveSysMsg.connect(self.ReceiveSysMsg) # 일반 TR에 대한 응답을 받는 함수를 연결

        # TR ID를 저장해놓고 처리할 딕셔너리 생성
        self.rqidD = {}


    def rqBalance(self, rqDate:str, acnt_num:str, pwd:str, product:str='13'):
        "선물 잔고 조회"
        ret = self.IndiTR.dynamicCall("SetQueryName(QString)", "SABF581Q2")
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 0, rqDate) # 조회일자 (YYYYMMDD)
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 1, acnt_num)   # 계좌번호
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 2, product)    # 계좌상품 (해외선물(13))
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 3, pwd)    # 비밀번호
        rqid = self.IndiTR.dynamicCall("RequestData()")  # 데이터 요청
        self.rqidD[rqid] = "SABF581Q2"

    
    def startBalanceRT(self, acnt_num):
        '선물 잔고 실시간'
        self.stopBalanceRT()    # 기존 실시간 해제
        self.strCurrentAcnt = acnt_num
        ret = self.IndiReal.dynamicCall("RequestRTReg(QString, QString)", "f3", self.strCurrentAcnt)


    def stopBalanceRT(self, acnt_num):
        if self.strCurrentAcnt != '':
            self.IndiReal.dynamicCall("UnRequestRTReg(QString, QString)", "f3", acnt_num)

    
    def ReceiveData(self, rqid):
        TRName = self.rqidD[rqid]

        # if TRName == "SABC967Q1":
        #     lstDATA = []
        #     nCnt = self.IndiTR.dynamicCall("GetMultiRowCount()")
        #     if nCnt != 0:
        #         for i in range(0, nCnt):
        #             DATA = {}
        #             DATA['종목코드'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 0)
        #             DATA['종목명'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 1)
        #             DATA['매매구분'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 2)  # 01:매도, 02:매수, 03:금전신탁
        #             DATA['잔고수량'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 3)
        #             DATA['청산가능수량'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 4)
        #             DATA['미체결수량'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 5)
        #             DATA['평균단가'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 6)  # 소수점2자리
        #             DATA['현재가'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 7)  # 소수점2자리
        #             DATA['전일대비'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 8)  # 소수점2자리
        #             DATA['전일대비율'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 9)  # 소수점2자리
        #             DATA['매입금액'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 10)
        #             DATA['평가금액'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 11)
        #             DATA['평가손익'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 12)
        #             DATA['손익율'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 13)   # 소수점2자리
        #             DATA['전일대비손익'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 14)
        #             DATA['손익가감액'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 15)
        #             DATA['델타'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 16)   # 소수점6자리
        #             DATA['감마'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 17)   # 소수점6자리
        #             DATA['매매손익'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 18)
        #             DATA['수수료'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 19)
        #             DATA['계약당승수'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 20)
        #             DATA['전일종가'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 21)
        #             DATA['선물옵션구분'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 22)
        #             DATA['베가'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 23) # 소수점6자리
        #             DATA['세타'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 24) # 소수점6자리
        #             DATA['기초자산'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 25)
        #             DATA['원체결단가'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 26)
        #         self.instInterface.chkStop2(lstDATA)

        #     self.instInterface.setTwBalanceInfoUI(lstDATA, False)

        if TRName == "SABF581Q2":
            lstDATA = []
            nCnt = self.IndiTR.dynamicCall("GetMultiRowCount()")
            if nCnt != 0:
                for i in range(0, nCnt):
                    DATA = {}
                    DATA['종목코드'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 0)
                    DATA['매매구분'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 1)  # 01:매도, 02:매수, 03:금전신탁
                    DATA['미결제수량'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 2)
                    DATA['청산가능수량'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 3)
                    DATA['평균약정가'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 4)  # 소수점2자리
                    DATA['현재가'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 5)  # 소수점2자리
                    DATA['통화코드'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 6)
                    DATA['평가손익'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 7)
                    DATA['잔고금액'] = self.IndiTR.dynamicCall("GetMultiData(int, int)", i, 8)   # 소수점2자리
                self.instInterface.chkStop2(lstDATA)

            self.instInterface.setTwBalanceInfoUI(lstDATA, False)
    

    def ReceiveRTData(self, RealType):
        # if RealType == 'AE':
        #     DATA = {}
        #     DATA['계좌번호'] = self.IndiTR.dynamicCall("GetSingleData(int)", 0)
        #     DATA['상품구분'] = self.IndiTR.dynamicCall("GetSingleData(int)", 1)
        #     DATA['종목코드'] = self.IndiTR.dynamicCall("GetSingleData(int)", 2)
        #     DATA['종목명'] = self.IndiTR.dynamicCall("GetSingleData(int)", 3)
        #     DATA['계좌명'] = self.IndiTR.dynamicCall("GetSingleData(int)", 4)
        #     DATA['매수매도구분'] = self.IndiTR.dynamicCall("GetSingleData(int)", 5) # 01: 매도, 02: 매수
        #     DATA['당일잔고'] = self.IndiTR.dynamicCall("GetSingleData(int)", 6)
        #     DATA['평균단가'] = self.IndiTR.dynamicCall("GetSingleData(int)", 7)
        #     DATA['미체결수량'] = self.IndiTR.dynamicCall("GetSingleData(int)", 8)
        #     DATA['청산가능수량'] = self.IndiTR.dynamicCall("GetSingleData(int)", 9)
        #     DATA['총매입금액'] = self.IndiTR.dynamicCall("GetSingleData(int)", 10)
        #     DATA['평가금액'] = self.IndiTR.dynamicCall("GetSingleData(int)", 11)
        #     DATA['평가손익'] = self.IndiTR.dynamicCall("GetSingleData(int)", 12)
        #     DATA['손익률'] = self.IndiTR.dynamicCall("GetSingleData(int)", 13)
        #     DATA['수수료'] = self.IndiTR.dynamicCall("GetSingleData(int)", 14)
        #     DATA['세금'] = self.IndiTR.dynamicCall("GetSingleData(int)", 15)
        #     DATA['현재가'] = self.IndiTR.dynamicCall("GetSingleData(int)", 17)
        #     DATA['거래승수'] = self.IndiTR.dynamicCall("GetSingleData(int)", 23)
        #     DATA['종목구분'] = self.IndiTR.dynamicCall("GetSingleData(int)", 24)    # 1: 선물, 2: 지수옵션, 3: 주식옵션
        #     DATA['종목매매손익'] = self.IndiTR.dynamicCall("GetSingleData(int)", 25)
        #     DATA['선물총매매손익'] = self.IndiTR.dynamicCall("GetSingleData(int)", 26)
        #     DATA['옵션총매매손익'] = self.IndiTR.dynamicCall("GetSingleData(int)", 27)
        #     DATA['기초자산명'] = self.IndiTR.dynamicCall("GetSingleData(int)", 28)
        #     DATA['이동평균단가'] = self.IndiTR.dynamicCall("GetSingleData(int)", 29)
        #     DATA['원체결단가'] = self.IndiTR.dynamicCall("GetSingleData(int)", 30)
        #     DATA['이동평균 매매손익'] = self.IndiTR.dynamicCall("GetSingleData(int)", 31)

        #     if int(DATA['청산가능수량']) > 0:
        #         self.instInterface.setTwBalanceInfoUI(DATA, True)
        #         self.instInterface.chkPL(DATA)

        if RealType == 'f3':
            DATA = {}
            DATA['처리구분'] = self.IndiTR.dynamicCall("GetSingleData(int)", 0) # 무시
            DATA['계좌번호'] = self.IndiTR.dynamicCall("GetSingleData(int)", 1)
            DATA['계좌명'] = self.IndiTR.dynamicCall("GetSingleData(int)", 2)
            DATA['종목코드'] = self.IndiTR.dynamicCall("GetSingleData(int)", 3)
            DATA['종목명'] = self.IndiTR.dynamicCall("GetSingleData(int)", 4)
            DATA['매매구분'] = self.IndiTR.dynamicCall("GetSingleData(int)", 5) # 'B' 매수, 'S' 매도
            DATA['잔고수량'] = self.IndiTR.dynamicCall("GetSingleData(int)", 6)
            DATA['단가'] = self.IndiTR.dynamicCall("GetSingleData(int)", 7)
            DATA['청산가능'] = self.IndiTR.dynamicCall("GetSingleData(int)", 8)
            DATA['미체결수량'] = self.IndiTR.dynamicCall("GetSingleData(int)", 9)
            DATA['현재가'] = self.IndiTR.dynamicCall("GetSingleData(int)", 10)
            DATA['전일대비'] = self.IndiTR.dynamicCall("GetSingleData(int)", 11)    # 0으로 내려감
            DATA['전일대비율'] = self.IndiTR.dynamicCall("GetSingleData(int)", 12)
            DATA['평가금액'] = self.IndiTR.dynamicCall("GetSingleData(int)", 13)
            DATA['평가손익'] = self.IndiTR.dynamicCall("GetSingleData(int)", 14)
            DATA['손익율'] = self.IndiTR.dynamicCall("GetSingleData(int)", 15)
            DATA['매입금액'] = self.IndiTR.dynamicCall("GetSingleData(int)", 16)
            DATA['승수'] = self.IndiTR.dynamicCall("GetSingleData(int)", 17)    # 계약사이즈 * 승수
            DATA['종목구분'] = self.IndiTR.dynamicCall("GetSingleData(int)", 18)    # F:선물, O:옵션
            DATA['통화코드'] = self.IndiTR.dynamicCall("GetSingleData(int)", 19)    # USD, AUD, JPY, EUR ......
            DATA['지점코드'] = self.IndiTR.dynamicCall("GetSingleData(int)", 20)
            DATA['지점명'] = self.IndiTR.dynamicCall("GetSingleData(int)", 21)
            DATA['관리자'] = self.IndiTR.dynamicCall("GetSingleData(int)", 22)
            DATA['관리자명'] = self.IndiTR.dynamicCall("GetSingleData(int)", 23)

            if int(DATA['청산가능']) > 0:
                self.instInterface.setTwBalanceInfoUI(DATA, True)
                self.instInterface.chkPL(DATA)


    # 시스템 메시지를 받은 경우 출력
    def ReceiveSysMsg(self, MsgID):
        self.instInterface.setSysMsgOnStatusBar(MsgID, __file__)