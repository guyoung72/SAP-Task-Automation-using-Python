from tkinter import * # 먼저 임포트해준다
import pyautogui
import subprocess
# import win32gui
import webbrowser
import time
import pygetwindow as gw
import pyperclip
import os
from datetime import datetime
from pywinauto import Application

def openErrorScreen2():
    ErrorScreen2=Toplevel(root)
    ErrorScreen2.title("90AUTO 에러")
    ErrorScreen2.iconbitmap('icon/error.ico')  # 아이콘 지정
    ErrorScreen2.resizable(False,False)
    Label(ErrorScreen2,text="\nSAP SV2창을 먼저 여세요.\n").pack()
    Button(ErrorScreen2, text = '      확인      ', command= ErrorScreen2.destroy).pack(pady=10)
    ws = ErrorScreen2.winfo_screenwidth()  # width of the screen
    hs = ErrorScreen2.winfo_screenheight()  # height of the screen
    w = 800
    h = 100
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)
    ErrorScreen2.geometry('%dx%d+%d+%d' % (w,h,x,y))

start = datetime.today().replace(day=1).strftime("%Y.%m.%d")
end = datetime.today().strftime("%Y.%m.%d")

def function1():
    try:
        sv2 = gw.getWindowsWithTitle(')/500')[0]
        def show_sv2():
            if sv2.isActive == False:
                pywinauto.application.Application().connect(handle=sv2._hWnd).top_window().set_focus()
    except:
        openErrorScreen2()
    os.system("vbs\\DeliveryAmount0.vbs {} {}".format(start, end))

    basis = gw.getWindowsWithTitle('Basis (1)')[0]
    basis.activate()

    xl = win32gui.FindWindow(None, "Basis (1)의 워크시트 - Excel")
    win32gui.MoveWindow(xl, 0, 0, 1000, 1000, True)

    basis.activate()
    basis.maximize()
    time.sleep(0.5)

    pyautogui.click(x=500, y=500)
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'A')
    time.sleep(1)
    pyautogui.press('alt')
    pyautogui.press('n')
    pyautogui.press('v')
    time.sleep(0.3)
    pyautogui.press('enter')
    time.sleep(4)

    pyautogui.click(x=1600, y=306)
    time.sleep(0.2)
    pyautogui.click(x=1600, y=306)

    pyperclip.copy("profit")
    pyautogui.hotkey("ctrl","v")
    time.sleep(0.5)
    pyautogui.moveTo(x=1600, y=342)
    pyautogui.dragTo(1600, 1000, 0.5, button='left')
    time.sleep(0.5)

    pyautogui.click(x=1600, y=306)
    pyautogui.hotkey('ctrl', 'A')

    pyperclip.copy("sold")
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.5)
    pyautogui.moveTo(x=1600, y=342)
    pyautogui.dragTo(1600, 1000, 0.5, button='left')
    time.sleep(0.5)

    pyautogui.click(x=1600, y=306)
    pyautogui.hotkey('ctrl', 'A')

    pyperclip.copy("material")
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.5)

    pyautogui.moveTo(x=1600, y=342)
    pyautogui.dragTo(1600, 1000, 0.5, button='left')
    time.sleep(0.5)

    pyautogui.moveTo(x=1635, y=385)
    pyautogui.dragTo(1600, 1015, 0.5, button='left')
    time.sleep(0.5)

    pyautogui.click(x=1600, y=306)
    pyautogui.hotkey('ctrl', 'A')
    pyperclip.copy("delivery")
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.5)

    pyautogui.moveTo(x=1600, y=342)
    pyautogui.dragTo(1600, 1030, 0.5, button='left')
    time.sleep(0.5)

    pyautogui.moveTo(x=1653, y=383)
    pyautogui.dragTo(1609, 800, 0.5, button='left')
    time.sleep(0.5)

    pyautogui.click(x=1600, y=306)
    pyautogui.hotkey('ctrl', 'A')
    pyperclip.copy("customer material")
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.5)

    pyautogui.moveTo(x=1600, y=342)
    pyautogui.dragTo(1600, 1030, 0.5, button='left')
    time.sleep(0.5)

    pyautogui.click(x=586, y=242)
    time.sleep(3)
    pyautogui.press('tab')
    pyautogui.write('"0"')
    pyautogui.press('enter')

    time.sleep(1)

    pyautogui.press('alt')
    pyautogui.press('j')
    pyautogui.press('y')
    pyautogui.press('p')
    pyautogui.press('t')

    def subtotals():
        pyautogui.hotkey('shift', 'f10')
        pyautogui.press('b')
        pyautogui.press('right')

    for _ in range(5):
        subtotals()

    pyautogui.hotkey('shift', 'f10')
    pyautogui.press('b')

    pyautogui.hotkey('ctrl', 'a')
    pyautogui.press('alt')
    pyautogui.press('h')
    pyautogui.press('f')
    pyautogui.press('k')
    time.sleep(0.1)
    pyautogui.press('alt')
    pyautogui.press('h')
    pyautogui.press('f')
    pyautogui.press('k')
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', ' ')
    pyautogui.press('alt')
    pyautogui.press('h')
    pyautogui.press('o')
    pyautogui.press('i')
    time.sleep(1)
    pyautogui.alert(text='피벗 테이블 완성', title='구영의 메시지', button='OK')


root=Tk() #'객체(root)'를 생성한다

root.title("Delivery Amount 0 [ZALL_SD_REPORT_004N]") #제목
Label(root, text='Delivery Amount=0 피벗으로 정리 Bot ',font='Helvetica 16 bold',bg='light blue').pack()
Label(root, text='\n설명:',font='Helvetica 10 bold').pack()
Label(root, text='▶사용자가 실행 버튼을 눌러주면, Bot이 알아서 SAP과 Excel을 조작하여\nDelivery Amount=0인 Billing number들만 피벗테이블로 정리하여 만들어 줍니다.\n\n▶완성된 피벗테이블은 각 Profit Center 담당자 분들께 보내는 메일에\n표로 첨부해 보내드리면 됩니다.').pack()
Label(root, text='\n실행 방법:',font='Helvetica 10 bold').pack()
Label(root, text='1. SV2에 먼저 로그인해주세요.\n\n2. 아래 실행 버튼을 눌러주세요.\n\n3. 피벗테이블이 완성될 때까지 기다리세요. (화장실 갔다오세요)').pack()

startbtn = Button(root, text = '    실행    ',font='Helvetica 12 bold', command=function1)
startbtn.pack(pady = 8)

def openCredits():
    Credits=Toplevel(root)
    Credits.title("90AUTO 정보")
    Credits.iconbitmap('icon/info.ico')  # 아이콘 지정
    Credits.resizable(False,False)
    Label(Credits, text='\nCreated by:', font='Helvetica 10 bold').pack()
    Label(Credits,text="인턴 김구영\n(2021.7~12)").pack()
    link1 = Label(Credits, text="linkedin.com/in/guyoung-kim/", fg="blue", cursor="hand2")
    link1.pack()
    def callback(url):
        webbrowser.open(url, new=0, autoraise=True)
    link1.bind("<Button-1>", lambda e: callback("https://www.linkedin.com/in/guyoung-kim/"))
    Label(Credits, text='\nMethods Used:', font='Helvetica 10 bold').pack()
    Label(Credits, text="Python(mainly pyautogui) & SAP Script Recording\n").pack()
    Label(Credits, text='\nSpecial Thanks to:', font='Helvetica 10 bold').pack()
    Label(Credits, text="HDP 사원님\n").pack()
    Button(Credits, text = '      닫기      ', command= Credits.destroy).pack(pady=10)
    ws = Credits.winfo_screenwidth()  # width of the screen
    hs = Credits.winfo_screenheight()  # height of the screen
    w = 300
    h = 300
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)
    Credits.geometry('%dx%d+%d+%d' % (w,h,x,y))

creditsbtn = Button(root, text = '      정보      ', command= openCredits)
creditsbtn.pack(pady = 8)

quitbtn = Button(root, text = '      종료      ', command= root.quit)
quitbtn.pack(pady = 8)

root.iconbitmap('icon/title.ico') #아이콘 지정
root.geometry("500x420") #루트창 크기
root.eval('tk::PlaceWindow . center')

root.resizable(False,False) #창크기 조절 못해요

root.mainloop() #핵심. 빙빙돌려주기
