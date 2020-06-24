from PyQt5.QAxContainer import *
from PyQt5.QtCore import *
from config.errorCode import  *
from PyQt5.QtTest import *

class Kiwoom(QAxWidget):
    def __init__(self):
        super().__init__()
        print("Kiwoom() class start.")

        ###### event loop 를 실행시키기 위한 변수 모음
        self.login_event_loop= QEventLoop() # 로그인 요청용 이벤트 루프
        self.detail_account_info_event_loop = QEventLoop()
        self.calculator_event_loop = QEventLoop()
        #########################################

        #### 계좌 관련 변수
        self.accoun_num=None
        self.deposit = 0 #예수금
        self.use_money = 0 #실제투자에 사용할 금 액
        self.use_money_percent = 0.5 # 예수금에서 실제 사용할 비율
        self.output_deposit = 0 # 출력가능금액

        #### 변수모음
        self.account_stock_dict={}
        self.not_account_stock_dict = {}

        #### 요청스크린 번호
        self.screen_my_info = "2000"# 계좌 관련한 스크린번호
        self.screen_calculation_stock = "4000" #계산용 스크린 번호

        ###### 초기 셋팅함수들 바로 실행
        self.get_ocx_instance() # OCX 방식을 파이썬에 사용할 수 있게 반환해 주는 함수 실행
        self.event_slots() # 키움과 연결하기 위한 시그널/슬록모음
        self.signal_login_commConnect() # 로그인 요청 함수 포함
        self.get_account_info() #계좌번호 가져오기
        self.detail_account_info() #예수금 요청 시그널 포함
        self.detail_account_mystock() #계좌평가 잔고내역요청
        self.not_concluded_account() #미체결 요청
        #QTimer.singleShot(5000, self.not_concluded_account)
        self.calculator_fnc() #종목분석용, 임시용으로 실행

    def get_ocx_instance(self):
        self.setControl("KHOPENAPI.KHOpenAPICtrl.1") #래지스트리에 저장된 API모듈 불러오기

    def event_slots(self):
        self.OnEventConnect.connect(self.login_slot) # 로그인 관련 이벤트
        self.OnReceiveTrData.connect(self.trdata_slot)

    def signal_login_commConnect(self):
        self.dynamicCall("CommConnect()")  # 로그인 요청 시그널
        self.login_event_loop.exec_()  # 이벤트 루프 실행

    def login_slot(self,err_code):
        print(errors(err_code))
        # 로그인 처리가 완료됐으면 이벤트 루프를 종료한다.
        self.login_event_loop.exit()

    def get_account_info(self):
        account_list = self.dynamicCall("GetLoginInfo(QString)","ACCNO")
        account_num = account_list.split(";")[0] #a;b;c -> [a,b,c]
        self.account_num=account_num
        print('계좌번호 : %s' %account_num) # 8136787011

    def detail_account_info(self, sPrevNext="0"):
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, Qstring)", "예수금상세현황요청","opw00001",sPrevNext,self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

    def detail_account_mystock(self, sPrevNext="0"):
        print("계좌평가잔고내역요청")
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호", "0000")
        self.dynamicCall("SetInputValue(QString, QString)", "비밀번호입력매체구분", "00")
        self.dynamicCall("SetInputValue(QString, QString)", "조회구분", "1")
        self.dynamicCall("CommRqData(QString, QString, int, Qstring)", "계좌평가잔고내역요청", "opw00018", sPrevNext, self.screen_my_info)

        self.detail_account_info_event_loop.exec_()

    def not_concluded_account(self, sPrevNext = "0"):
        print("미체결")
        self.dynamicCall("SetInputValue(QString, QString)", "계좌번호", self.account_num)
        self.dynamicCall("SetInputValue(QString, QString)", "체결구분", "1")
        self.dynamicCall("SetInputValue(QString, QString)", "매매구분", "0")
        self.dynamicCall("CommRqData(QString, QString, int, QString)", "실시간미체결요청","opt10075",sPrevNext,self.screen_my_info)

        self.detail_account_info_event_loop.exec_()


    def trdata_slot(self, sScrNo, sRQName, sTrCode, sRecordName, sPrevNext):
        '''
        tr요청을 받는 구역
        :param sScrNo: 스크린번호
        :param sRQName: 내가 요청했을때 지은 이름
        :param sTrCode: 요청id / trCode
        :param sRecordName: 사용안함
        :param sPrevNext: 다음 페이지가 있는지
        :return:
        '''
        if sRQName =='예수금상세현황요청':
            deposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "예수금")
            print("예수금 : %s" %self.deposit)
            self.deposit = int(deposit)
            use_money = float(self.deposit)*self.use_money_percent
            self.use_money = int(use_money)
            self.use_money = self.use_money/4

            output_deposit = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "출금가능금액")
            self.output_deposit = int(output_deposit)
            print("출금가능금액 : %s" %self.output_deposit)
            self.stop_screen_cancel(self.screen_my_info)
            self.detail_account_info_event_loop.exit()

        elif sRQName =='계좌평가잔고내역요청':
            total_buy_money = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총매입금액")
            total_buy_money_result = int(total_buy_money)
            print("총 매입금액 %s" %total_buy_money_result)
            total_profit_loss_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, 0, "총수익률(%)")
            total_profit_loss_rate_result = float(total_profit_loss_rate)
            print("총 수익률 %s" %total_profit_loss_rate_result)

            rows = self.dynamicCall("GetRepeatCnt(QString, Qstring)",sTrCode,sRQName)
            cnt = 0
            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString, QString, int, QString)",sTrCode,sRQName, i, "종목번호")
                code = code.strip()[1:]
                code_nm = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "종목명")
                stock_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                  "보유수량")
                buy_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i, "매입가")
                learn_rate = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                              "수익률(%)")
                current_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                 "현재가")
                total_chegual_price = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName,
                                                       i, "매입금액")
                possible_quantity = self.dynamicCall("GetCommData(QString, QString, int, QString)", sTrCode, sRQName, i,
                                                     "매매가능수량")
                print("종목번호: %s - 종목명: %s - 보유수량: %s - 매입가:%s - 수익률: %s - 현재가: %s" % (code, code_nm, stock_quantity, buy_price, learn_rate, current_price))

                if code in self.account_stock_dict:
                    pass
                else:
                    self.account_stock_dict[code] ={}

                code_nm = code_nm.strip()
                stock_quantity = int(stock_quantity.strip())
                buy_price = int(buy_price.strip())
                learn_rate = float(learn_rate.strip())
                current_price = int(current_price.strip())
                total_chegual_price = int(total_chegual_price.strip())
                possible_quantity = int(possible_quantity.strip())

                self.account_stock_dict[code].update({"종목명": code_nm})
                self.account_stock_dict[code].update({"보유수량": stock_quantity})
                self.account_stock_dict[code].update({"매입가": buy_price})
                self.account_stock_dict[code].update({"수익률(%)": learn_rate})
                self.account_stock_dict[code].update({"현재가": current_price})
                self.account_stock_dict[code].update({"매입금액": total_chegual_price})
                self.account_stock_dict[code].update({"매매가능수량": possible_quantity})

                cnt+=1

            print("계좌에 가지고 있는 종목 %s" %cnt)
            print("계좌 보유종목 count %s" %cnt)

            if sPrevNext=="2":
                self.detail_account_mystock(sPrevNext="2")
            else:
                self.detail_account_info_event_loop.exit()

        elif sRQName =="실시간미체결요청":
            rows = self.dynamicCall("GetRepeatCnt(QString,QString)",sTrCode,sRQName)

            for i in range(rows):
                code = self.dynamicCall("GetCommData(QString,QString,int,QString",sTrCode,sRQName,i,"종목코드")
                code_nm = self.dynamicCall("GetCommData(QString,QString,int,QString",sTrCode,sRQName,i,"종목명")
                order_no = self.dynamicCall("GetCommData(QString,QString,int,QString",sTrCode,sRQName,i,"주문번호")
                order_status = self.dynamicCall("GetCommData(QString,QString,int,QString",sTrCode,sRQName,i,"주문상태")
                code_quantity = self.dynamicCall("GetCommData(QString,QString,int,QString",sTrCode,sRQName,i,"주문수량")
                code_price = self.dynamicCall("GetCommData(QString,QString,int,QString",sTrCode,sRQName,i,"주문가격")
                code_gubun = self.dynamicCall("GetCommData(QString,QString,int,QString",sTrCode,sRQName,i,"주문구분")
                not_quantity = self.dynamicCall("GetCommData(QString,QString,int,QString",sTrCode,sRQName,i,"미체결수량")
                ok_quantity = self.dynamicCall("GetCommData(QString,QString,int,QString",sTrCode,sRQName,i,"체결량")

                code = code.strip()
                code_nm =  code_nm.strip()
                order_no = int(order_no.strip())
                order_status = order_status.strip()
                code_quantity = int(order_quantity.strip())
                code_price =  int(order_price.strip())
                code_gubun = order_gubun.strip().lstrip('+').lstrip('-')
                not_quantity = int(not_quantity.strip())
                ok_quantity = int(ok_quantity.strip())

                if order_no in self.not_account_stock_dict:
                    pass
                else:
                    self.not_account_stock_dict[order_no] = {}

                nasd = self.not_account_stock_dict[order_no]
                nasd.update({'종목코드':code})
                nasd.update({'종목명':code_nm})
                nasd.update({'주문번호':order_no})
                nasd.update({'주문상태':order_status})
                nasd.update({'주문수량':order_quantity})
                nasd.update({'주문가격':order_price})
                nasd.update({'주문구분':order_gubun})
                nasd.update({'미체결수량':not_quantity})
                nasd.update({'체결량':ok_quantity})

                print("미체결종목 : %s " %self.not_account_stock_dict[order_no])

            self.detail_account_info_event_loop.exit()

        elif sRQName == '주식일봉차트조회':
            print('일봉데이터 요청')
            code = self.dynamicCall("GetCommData(QString,QString,int,QString",sTrCode,sRQName,0, "종목코드")
            code = code.strip()
            print("%s 일봉데이터 요청" % code)

            rows = self.dynamicCall("GetRepeatCnt(QString,QString)",sTrCode,sRQName)
            print(rows)

            if sPrevNext == "2":
                self.day_kiwoom_db(code=code, sPrevNext=sPrevNext)
            else:
                self.calculator_event_loop.exit()


    def stop_screen_cancel(self,sScrNo=None):
        self.dynamicCall("DisconnectRealData(QString)", sScrNo)


    def get_code_list_by_matket(self, market_code):
        # 종목코드들 반환
        code_list = self.dynamicCall("GetCodeListByMarket(QString)",market_code)
        code_list = code_list.split(";")[:-1]
        return code_list


    def calculator_fnc(self):
        '''
        종목분석 실행용 함수
        :return:
        '''
        code_list = self.get_code_list_by_matket("10")
        print("코스닥 갯수 %s" %len(code_list))

        for idx, code in enumerate(code_list):
            self.dynamicCall("DisconnectRealData(QString)", self.screen_calculation_stock)
            print("%s/%s : KOSDAQ Strock Code : %s is updating... " % (idx+1,len(code_list),code))
            self.day_kiwoom_db(code=code)


    def day_kiwoom_db(self,code=None, date=None, sPrevNext="0"):

        QTest.qWait(3600) # 3.6초마다 딜레이

        self.dynamicCall("SetInputValue(QString,QString)","종목코드",code)
        self.dynamicCall("SetInputValue(QString,QString)","수정주가구분",1)

        if date != None:
            self.dynamicCall("SetInputValue(QString,QString)","기준일자",date)
        self.dynamicCall("CommRqData(QString,QString,int,QString","주식일봉차트조회","opt10081",sPrevNext,self.screen_calculation_stock)
        self.calculator_event_loop.exec_()





