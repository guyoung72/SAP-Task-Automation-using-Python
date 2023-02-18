from tkinter import *
from tkinter import scrolledtext
import webbrowser
import pyautogui
import subprocess
import time
import pygetwindow as gw
import os
from datetime import datetime
import pywinauto

def openErrorScreen1():
    ErrorScreen1=Toplevel(root)
    ErrorScreen1.title("90AUTO 에러")
    ErrorScreen1.iconbitmap('icon/error.ico')  # 아이콘 지정
    ErrorScreen1.resizable(False, False)
    Label(ErrorScreen1, text="\nSAP SV2창을 먼저 여세요.\n").pack()
    Button(ErrorScreen1, text ='      확인      ', command= ErrorScreen1.destroy).pack(pady=10)
    ErrorScreen1.lift
    ws = ErrorScreen1.winfo_screenwidth()  # width of the screen
    hs = ErrorScreen1.winfo_screenheight()  # height of the screen
    w = 300
    h = 100
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)
    ErrorScreen1.geometry('%dx%d+%d+%d' % (w, h, x, y))

def function_check_error():
    try:
        sv2 = gw.getWindowsWithTitle(')/500')[0]
        def show_sv2():
            if sv2.isActive == False:
                pywinauto.application.Application().connect(handle=sv2._hWnd).top_window().set_focus()
    except:
        openErrorScreen1()
    show_sv2()
    start = datetime.today().replace(day=1).strftime("%Y.%m.%d")
    end = datetime.today().strftime("%Y.%m.%d")
    os.system("vbs\\Jenny_vfx3_CheckingError.vbs {} {}".format(start, end))
    show_sv2()
    openCompleteScreen('vfx에러 확인')

def function_for_trade_error():
    try:
        sv2 = gw.getWindowsWithTitle(')/500')[0]
        def show_sv2():
            if sv2.isActive == False:
                pywinauto.application.Application().connect(handle=sv2._hWnd).top_window().set_focus()
    except:
        openErrorScreen1()
    show_sv2()
    print('For Tradedata Error를 해결할 빌링 넘버를 입력하세요 (최대 5개까지 붙여넣기 가능합니다.)')
    arg0 = input('빌링 넘버를 붙여넣으세요 1:')
    arg1 = input('빌링 넘버를 붙여넣으세요 2:')
    arg2 = input('빌링 넘버를 붙여넣으세요 3:')
    arg3 = input('빌링 넘버를 붙여넣으세요 4:')
    arg4 = input('빌링 넘버를 붙여넣으세요 5:')

    os.system("vbs\\Jenny_vf02_ForTradedataError.vbs {} {} {} {} {}".format(arg0, arg1, arg2, arg3, arg4))
    show_sv2()
    openCompleteScreen('For.Tradedata Error 해결')

def open_trade_error_entry_box():
    try:
        sv2 = gw.getWindowsWithTitle(')/500')[0]
        def show_sv2():
            if sv2.isActive == False:
                pywinauto.application.Application().connect(handle=sv2._hWnd).top_window().set_focus()
        trade_error_entry_box = Toplevel(root)
        trade_error_entry_box.title("ForTrade Error 해결사")
        trade_error_entry_box.iconbitmap('icon/title.ico')  # 아이콘 지정
        Label(trade_error_entry_box, text="\nBilling number들의 ForTrade/Customs \n상태를 L로 변경하여 에러를 해결합니다.\n",
              font='Helvetica 10 bold').pack()
        Label(trade_error_entry_box,
              text="For.Tradedata Error가 발생한 Billing number들을\n 복사 후 아래 빈 칸에 붙여넣은 뒤 '실행'을 누르세요.\n").pack()
        ws = trade_error_entry_box.winfo_screenwidth()  # width of the screen
        hs = trade_error_entry_box.winfo_screenheight()  # height of the screen
        w = 310
        h = 410
        x = (ws / 2) - (w / 2)
        y = (hs / 2) - (h / 2)
        trade_error_entry_box.geometry('%dx%d+%d+%d' % (w, h, x, y))
        trade_error_entry_box.lift
        # 여기에 multi-line input 입력
        text_area = scrolledtext.ScrolledText(trade_error_entry_box, wrap=WORD, width=20, height=7,
                                              font=("Calibri", 12))
        text_area.pack()
        text_area.focus()
        Button(trade_error_entry_box, text='    ▶ 실행     ', command=trade_error_entry_box.destroy).pack(pady=14)
        Button(trade_error_entry_box, text='      닫기      ', command=trade_error_entry_box.destroy).pack(pady=5)
    except:
        openErrorScreen1()

def openCompleteScreen(whatisdone):
    CompleteScreen = Toplevel(root)
    CompleteScreen.title(whatisdone+" 완료")
    CompleteScreen.iconbitmap('icon/complete.ico')  # 아이콘 지정
    CompleteScreen.resizable(False, False)
    CompleteScreen.lift
    Label(CompleteScreen, text="\n"+whatisdone+"을 완료하였습니다.").pack()
    Button(CompleteScreen, text='      확인      ', command=CompleteScreen.destroy).pack(pady=10)
    ws = CompleteScreen.winfo_screenwidth()  # width of the screen
    hs = CompleteScreen.winfo_screenheight()  # height of the screen
    w = 300
    h = 100
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)
    CompleteScreen.geometry('%dx%d+%d+%d' % (w, h, x, y))

root=Tk() #'객체(root)'를 생성한다
root.title("Billing Error 해결 Bot [vfx3]") #창 제목
Label(root, text='Billing Error 해결 Bot',font='Helvetica 16 bold',bg='light blue').pack() #제목
Label(root, text='\n설명:',font='Helvetica 10 bold').pack()
Label(root, text='▶Billing Error 통합 해결 Bot입니다.\n▶이 Bot이 현재 해결할 수 있는 에러는:\n1)For Tradedata Error  2)Pricing Error(메일전송용 표 만들기만)  3)Acct Determin. Error').pack()
Label(root, text='\n실행 방법:',font='Helvetica 10 bold').pack()
Label(root, text='1. SV2에 먼저 로그인해주세요.\n2. 아래 중 필요한 작업 버튼을 눌러주세요. 이후에 나오는 팝업창에 OK를 누르세요.\n').pack()


checkbtn = Button(root, text = '  vfx3 에러 확인  ', font='Helvetica 12 bold', command=function_check_error)
checkbtn.pack(pady = 5)

startbtn = Button(root, text = 'For.Tradedata Error 해결', font='Helvetica 12 bold', command=open_trade_error_entry_box)
startbtn.pack(pady = 5)

def openCredits():
    Credits=Toplevel(root)
    Credits.lift()
    Credits.title("90AUTO 정보")
    Credits.iconbitmap('icon/info.ico')  # 아이콘 지정
    Credits.resizable(False,False)
    Label(Credits, text='\nCreated by:', font='Helvetica 10 bold').pack()
    Label(Credits,text="김구영 인턴").pack()
    link1 = Label(Credits, text="linkedin.com/in/guyoung-kim/", fg="blue", cursor="hand2")
    link1.pack()
    def callback(url):
        webbrowser.open(url, new=0, autoraise=True)
    link1.bind("<Button-1>", lambda e: callback("https://www.linkedin.com/in/guyoung-kim/"))
    Label(Credits, text='\nMethods Used:', font='Helvetica 10 bold').pack()
    Label(Credits, text="Python(mainly pyautogui) & SAP Script Recording\n").pack()
    Label(Credits, text='\nSpecial Thanks to:', font='Helvetica 10 bold').pack()
    Label(Credits, text="홍두표 사원님\n").pack()
    Button(Credits, text = '      닫기      ', command= Credits.destroy).pack(pady=10)
    ws = Credits.winfo_screenwidth()  # width of the screen
    hs = Credits.winfo_screenheight()  # height of the screen
    w = 300
    h = 300
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)
    Credits.geometry('%dx%d+%d+%d' % (w,h,x,y))

creditsbtn = Button(root, text = '      정보      ', command= openCredits)
creditsbtn.pack(pady = 15)

quitbtn = Button(root, text = '      종료      ', command= root.quit)
quitbtn.pack(pady = 2)

root.iconbitmap('icon/title.ico') #아이콘 지정
root.geometry("500x400")
root.eval('tk::PlaceWindow . center')
root.lift

root.resizable(False,False) #창크기 조절 못해요

root.mainloop() #핵심. 빙빙돌려주기




