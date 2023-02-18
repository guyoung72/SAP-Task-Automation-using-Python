import os
import time
import pyautogui
import pygetwindow as gw
import pywinauto

print("MM01 빠르게 도번 생성하기\n \n1. SV2창을 띄어 놓는다. \n2. 아래에 복사한 도번을 입력해주세요.\n")

sv2 = gw.getWindowsWithTitle('SV2(1)/500')[0]  # 윈도우 타이틀에 SV2 이 포함된 모든 윈도우 수집, 리스트로 리턴
logon = gw.getWindowsWithTitle('SAP Logon 760')[0]  # 윈도우 타이틀에 SAP Logon 이 포함된 모든 윈도우 수집, 리스트로 리턴

#sv2 창 보여주기
def show_sv2():
    if sv2.isActive == False:
        pywinauto.application.Application().connect(handle=sv2._hWnd).top_window().set_focus()


#도번 1개를 입력하는 펑션
def function1():
    material = input("도번을 붙여넣으세요 :")
    show_sv2()
    os.system("vbs\\mm01.vbs {}".format(material))

    #만약 로그인창이 띄어져버렸다면, sv2창으로 다시 가도록
    sv2.activate()  # 윈도우 활성화
    time.sleep(0.1) # 잠시 쉬고

    time.sleep(0.1) #잠시 쉬고
    customer = input("3자리 고객사 번호를 입력하세요 :")

    if sv2.isActive == False:
        pywinauto.application.Application().connect(handle=sv2._hWnd).top_window().set_focus()

    os.system("vbs\\mm01_01.vbs {}".format(customer))

    blank = pyautogui.locateOnScreen('img\\SAP_Trans.Grp_blank.png')
    print(blank)

    if blank == None:
        os.system("vbs\\mm01_03.vbs")

    if not blank == None:
        os.system("vbs\\mm01_02.vbs")

    pyautogui.alert(text=material + ' 도번 생성 완료.', title='구영의 메시지', button='OK')
    print(material+' 도번 생성 완료.')


function1()

for i in range(100):
    function1()