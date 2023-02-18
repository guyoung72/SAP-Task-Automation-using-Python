from tkinter import *
import webbrowser
import os
import pyautogui
import pygetwindow as gw
import pywinauto

root=Tk() #'객체(root)'를 생성한다

root.title("이천(3180) 도번 sales view 생성하기 [mm01]") #제목
Label(root, text='[MM01] 이천 도번 sales view 생성 Bot ',font='Helvetica 16 bold',bg='light blue').pack()
Label(root, text='\n설명:',font='Helvetica 10 bold').pack()
Label(root, text='▶사용자가 복붙해놓은 도번의 Sales View를 생성해줍니다.').pack()
Label(root, text='\n실행 방법:',font='Helvetica 10 bold').pack()
Label(root, text='1. SV2에 먼저 로그인해주세요.\n2. 생성할 도번을 복사한 뒤, 다음 빈 칸에 붙여넣으세요:\n3.고객사 번호는 추후에 뜨는 창에 수동으로 입력해주셔야 합니다.\n').pack()

def callback(url):
   webbrowser.open(url, new=0, autoraise=True)

def openErrorScreen1():
    ErrorScreen1=Toplevel(root)
    ErrorScreen1.title("90AUTO 에러")
    ErrorScreen1.iconbitmap('icon/error.ico')  # 아이콘 지정
    ErrorScreen1.resizable(False, False)
    Label(ErrorScreen1, text="\nSAP SV2창을 먼저 여세요.\n").pack()
    Button(ErrorScreen1, text ='      확인      ', command= ErrorScreen1.destroy).pack(pady=10)
    ErrorScreen1.lift()
    ws = ErrorScreen1.winfo_screenwidth()  # width of the screen
    hs = ErrorScreen1.winfo_screenheight()  # height of the screen
    w = 300
    h = 100
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)
    ErrorScreen1.geometry('%dx%d+%d+%d' % (w, h, x, y))
    try:
        popup2 = gw.getWindowsWithTitle('90AUTO 에러')[0]
        popup2.activate()
    except:
        pass

def openCompleteScreen(whatisdone):
    CompleteScreen = Toplevel(root)
    CompleteScreen.title(whatisdone)
    CompleteScreen.iconbitmap('icon/complete.ico')  # 아이콘 지정
    CompleteScreen.resizable(False, False)
    Label(CompleteScreen, text="\n"+whatisdone+"을 완료하였습니다.").pack()
    Done=Button(CompleteScreen, text='      확인      ', command=CompleteScreen.destroy)
    Done.pack(pady=10)
    ws = CompleteScreen.winfo_screenwidth()  # width of the screen
    hs = CompleteScreen.winfo_screenheight()  # height of the screen
    w = 300
    h = 100
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)
    CompleteScreen.geometry('%dx%d+%d+%d' % (w, h, x, y))
    try:
        popup3 = gw.getWindowsWithTitle('도번 생성 완료')[0]
        popup3.activate()
    except:
        CompleteScreen.lift()
    def omgplzwork(dd):
        CompleteScreen.destroy()
    CompleteScreen.bind('<Return>', omgplzwork)

# 여기에 single-line input
searchbox=Entry(root, width=18)
searchbox.focus()

# 여기에 multi-line input
# text_area = scrolledtext.ScrolledText(root, wrap=WORD, width=20, height=10, font=("Calibri", 12))
# text_area.pack()
# text_area.focus()

def openCustomer():
    Customer = Toplevel(root)
    Customer.title("고객사 번호 입력")
    Customer.iconbitmap('icon/title.ico')  # 아이콘 지정
    Customer.resizable(False, False)
    Label(Customer, text="\n3자리수 고객사 번호를 입력해주세요:\n\n").pack()
    customerbox=Entry(Customer, width=8)
    def function2(company=None):
        try:
            os.system("vbs\\mm01_01.vbs {}".format(customerbox.get()))
        except:
            Customer.destroy()
            openErrorScreen1()
        Customer.destroy()
        blank = pyautogui.locateOnScreen('img\\SAP_Trans.Grp_blank.png')
        print(blank)
        if blank == None:
            os.system("vbs\\mm01_03.vbs")
        if not blank == None:
            os.system("vbs\\mm01_02.vbs")
        openCompleteScreen(searchbox.get() + ' 도번 생성 완료')
    customerbox.pack()
    customerbox.bind("<Return>", function2)
    Button(Customer, text='      입력      \n(or press enter)',font='Helvetica 9 bold', command=function2).pack(pady=10)
    Customer.lift()
    Customer.attributes('-topmost', True)
    Customer.after_idle(root.attributes, '-topmost', False)
    customerbox.focus()
    ws = Customer.winfo_screenwidth()  # width of the screen
    hs = Customer.winfo_screenheight()  # height of the screen
    w = 325
    h = 180
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)
    Customer.geometry('%dx%d+%d+%d' % (w, h, x, y))
    try:
        popup1=gw.getWindowsWithTitle('고객사 번호 입력')[0]
        popup1.activate()
    except:
        customerbox.focus()

#도번 1개를 입력하는 펑션
def function1(event=None):
    try:
        sv2 = gw.getWindowsWithTitle(')/500')[0]
        def show_sv2():
            if sv2.isActive == False:
                pywinauto.application.Application().connect(handle=sv2._hWnd).top_window().set_focus()
    except:
        openErrorScreen1()
    sv2.activate()
    os.system("vbs\\mm01.vbs {}".format(searchbox.get()))
    sv2.activate()  # 윈도우 활성화
    openCustomer()


searchbox.pack()

startbtn = Button(root, text = ' Sales view 생성 \n(or press enter)',font='Helvetica 10 bold', command=function1)
startbtn.pack(pady = 8)
searchbox.bind('<Return>',function1)


def openCredits():
    Credits=Toplevel(root)
    Credits.title("90AUTO 정보")
    Credits.iconbitmap('icon/info.ico') #아이콘 지정
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
    Label(Credits, text="HDP 사원님\n").pack()
    Button(Credits, text = '      닫기      ', command= Credits.destroy).pack(pady=10)
    ws = Credits.winfo_screenwidth()  # width of the screen
    hs = Credits.winfo_screenheight()  # height of the screen
    w = 300
    h = 300
    x = (ws / 2) - (w / 2)
    y = (hs / 2) - (h / 2)
    Credits.geometry('%dx%d+%d+%d' % (w,h,x,y))
    Credits.lift()

creditsbtn = Button(root, text = '      정보      ', command= openCredits)
creditsbtn.pack(pady = 5)

quitbtn = Button(root, text = '      종료      ', command= root.quit)
quitbtn.pack(pady = 7)

root.iconbitmap('icon/title.ico') #아이콘 지정
root.geometry("480x370")
root.eval('tk::PlaceWindow . center')
root.lift()
root.resizable(False,False) #창크기 조절 못해요



root.mainloop() #핵심. 빙빙돌려주기