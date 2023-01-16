from PyQt5.QAxContainer import *
import numpy as np
# from System.strategyRT import StrategyRT


class PriceRT():
    def __init__(self, instInterface):
        self.instInterface = instInterface
        
        # 인디의 TR을 처리할 변수를 생성합니다.
        self.IndiTR = QAxWidget("GIEXPERTCONTROL.GiExpertControlCtrl.1")
        self.IndiReal = QAxWidget("GIEXPERTCONTROL.GiExpertControlCtrl.1")

        # Indi API event
        self.IndiTR.ReceiveData.connect(self.ReceiveData)
        self.IndiTR.ReceiveSysMsg.connect(self.ReceiveSysMsg)
        self.IndiReal.ReceiveRTData.connect(self.ReceiveRTData)

        self.rqidD = {} # TR 관리를 위해 사전 변수를 하나 생성

        # 데이터
        self.currentCode = None

        PriceInfodt = np.dtype([('종목명', 'S40'), ('상한가', 'f'), ('하한가', 'f'), ('전일종가', 'f'),
                        ('종목코드', 'S10'), ('시간', 'S6'), ('시가', 'f'), ('고가', 'f'), ('저가', 'f'), ('현재가', 'f'), 
                        ('체결수량', 'u4'), ('누적체결수량', 'u4'), ('매도1호가', 'f'), ('매도1호가수량', 'u4'), ('매수1호가', 'f'), ('매수1호가수량', 'u4')])
        self.PriceInfo = np.empty([1], dtype=PriceInfodt)

    def stopTR(self):
        # self.IndiReal.dynamicCall("UnRequestRTRegAll()")
        if self.currentCode != None:
            self.IndiReal.dynamicCall("UnRequestRTReg(QString, QString)", "fb", self.currentCode)
            self.IndiReal.dynamicCall("UnRequestRTReg(QString, QString)", "fc", self.currentCode)
            self.IndiReal.dynamicCall("UnRequestRTReg(QString, QString)", "fh", self.currentCode)

    def startTR(self):
        # 기존종목 실시간 해제
        self.stopTR()

        # 현재 종목코드 정보
        self.currentCode = self.instInterface.wndIndi.cbProductCode.currentText().upper()

        # 종목 기본정보 조회
        ret = self.IndiTR.dynamicCall("SetQueryName(QString)", "fb")
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 0, self.currentCode)
        rqid = self.IndiTR.dynamicCall("RequestData()")
        self.rqidD[rqid] = "fb"

        # 종목 현재가 조회
        ret = self.IndiTR.dynamicCall("SetQueryName(QString)", "fc")
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 0, self.currentCode)
        rqid = self.IndiTR.dynamicCall("RequestData()")
        self.rqidD[rqid] = "fc"

        # 종목 호가 조회
        ret = self.IndiTR.dynamicCall("SetQueryName(QString)", "fh")
        ret = self.IndiTR.dynamicCall("SetSingleData(int, QString)", 0, self.currentCode)
        rqid = self.IndiTR.dynamicCall("RequestData()")
        self.rqidD[rqid] = "fh"
        
        # 실시간 등록
        # ret = self.IndiReal.dynamicCall("RequestRTReg(QString, QString)", "fb", self.currentCode)   # 종목 기본정보
        # ret = self.IndiReal.dynamicCall("RequestRTReg(QString, QString)", "fc", self.currentCode)   # 현재가
        # ret = self.IndiReal.dynamicCall("RequestRTReg(QString, QString)", "fh", self.currentCode)   # 호가

    # 요청한 TR로 부터 데이터 수신
    def ReceiveData(self, rqid):
        # TR을 날릴때의 ID를 통해 TR이름 가져옴
        TRName = self.rqidD[rqid]
        
        # 종목 기본정보 수신
        if TRName == "fb":
            self.PriceInfo[0]['종목명'] = self.IndiTR.dynamicCall("GetSingleData(int)", 7)
            self.PriceInfo[0]['상한가'] = self.IndiTR.dynamicCall("GetSingleData(int)", 26)
            self.PriceInfo[0]['하한가'] = self.IndiTR.dynamicCall("GetSingleData(int)", 27)
            self.PriceInfo[0]['전일종가'] = self.IndiTR.dynamicCall("GetSingleData(int)", 29)
            # 실시간 등록
            ret = self.IndiReal.dynamicCall("RequestRTReg(QString, QString)", "fb", self.currentCode)
        
        # 종목 현재가 수신
        elif TRName == "fc":
            self.PriceInfo[0]['종목코드'] = self.IndiTR.dynamicCall("GetSingleData(int)", 0)
            self.PriceInfo[0]['시간'] = self.IndiTR.dynamicCall("GetSingleData(int)", 3)
            self.PriceInfo[0]['현재가'] = self.IndiTR.dynamicCall("GetSingleData(int)", 6)
            self.PriceInfo[0]['누적체결수량'] = self.IndiTR.dynamicCall("GetSingleData(int)", 11)
            self.PriceInfo[0]['체결수량'] = self.IndiTR.dynamicCall("GetSingleData(int)", 10)
            self.PriceInfo[0]['시가'] = self.IndiTR.dynamicCall("GetSingleData(int)", 12)
            self.PriceInfo[0]['고가'] = self.IndiTR.dynamicCall("GetSingleData(int)", 13)
            self.PriceInfo[0]['저가'] = self.IndiTR.dynamicCall("GetSingleData(int)", 14)

            # Data transfer
            self.instInterface.executeStrategy(self.PriceInfo[0])
            self.instInterface.setTwProductInfoUI(self.PriceInfo[0])

            # 실시간 등록
            ret = self.IndiReal.dynamicCall("RequestRTReg(QString, QString)", "fc", self.currentCode)

        # 종목 호가 수신
        elif TRName == "fh":
            try:
                self.PriceInfo[0]['매도1호가'] = self.IndiTR.dynamicCall("GetSingleData(int)", 22)
                self.PriceInfo[0]['매수1호가'] = self.IndiTR.dynamicCall("GetSingleData(int)", 5)
                self.PriceInfo[0]['매도1호가수량'] = self.IndiTR.dynamicCall("GetSingleData(int)", 23)
                self.PriceInfo[0]['매수1호가수량'] = self.IndiTR.dynamicCall("GetSingleData(int)", 6)
            except:
                self.instInterface.setSysMsgOnStatusBar("현재 호가 데이터가 수신되지 않습니다", __file__)
            # 실시간 등록
            ret = self.IndiReal.dynamicCall("RequestRTReg(QString, QString)", "fh", self.currentCode)

        self.rqidD.__delitem__(rqid)


    def ReceiveRTData(self, RealType):
        # 종목 기본정보 실시간 수신
        if RealType == "fb":
            self.PriceInfo[0]['종목명'] = self.IndiReal.dynamicCall("GetSingleData(int)", 2)
            self.PriceInfo[0]['상한가'] = self.IndiReal.dynamicCall("GetSingleData(int)", 26)
            self.PriceInfo[0]['하한가'] = self.IndiReal.dynamicCall("GetSingleData(int)", 27)
            self.PriceInfo[0]['전일종가'] = self.IndiReal.dynamicCall("GetSingleData(int)", 29)

        # 종목 현재가 실시간 수신
        elif RealType == "fc":
            self.PriceInfo[0]['종목코드'] = self.IndiReal.dynamicCall("GetSingleData(int)", 0)
            self.PriceInfo[0]['시간'] = self.IndiReal.dynamicCall("GetSingleData(int)", 3)
            self.PriceInfo[0]['현재가'] = self.IndiReal.dynamicCall("GetSingleData(int)", 6)
            self.PriceInfo[0]['누적체결수량'] = self.IndiReal.dynamicCall("GetSingleData(int)", 11)
            self.PriceInfo[0]['체결수량'] = self.IndiReal.dynamicCall("GetSingleData(int)", 10)
            self.PriceInfo[0]['시가'] = self.IndiReal.dynamicCall("GetSingleData(int)", 12)
            self.PriceInfo[0]['고가'] = self.IndiReal.dynamicCall("GetSingleData(int)", 13)
            self.PriceInfo[0]['저가'] = self.IndiReal.dynamicCall("GetSingleData(int)", 14)

            # Data transfer
            self.instInterface.executeStrategy(self.PriceInfo[0])
            self.instInterface.setTwProductInfoUI(self.PriceInfo[0])
            
        # 종목 호가 실시간 수신
        elif RealType == "fh":
            self.PriceInfo[0]['매도1호가'] = self.IndiReal.dynamicCall("GetSingleData(int)", 22)
            self.PriceInfo[0]['매수1호가'] = self.IndiReal.dynamicCall("GetSingleData(int)", 5)
            self.PriceInfo[0]['매도1호가수량'] = self.IndiReal.dynamicCall("GetSingleData(int)", 23)
            self.PriceInfo[0]['매수1호가수량'] = self.IndiReal.dynamicCall("GetSingleData(int)", 6)

        # print(self.PriceInfo[0])


    # 시스템 메시지를 받은 경우 출력
    def ReceiveSysMsg(self, MsgID):
        self.instInterface.setSysMsgOnStatusBar(MsgID, __file__)