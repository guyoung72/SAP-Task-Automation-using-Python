# SAP Task Automation using Python
How to automate SAP tasks using its recording features and Python codes (inspired during my finance internship at Continental <img src="https://user-images.githubusercontent.com/79275984/218221656-6282a15e-7f25-46e7-9aa7-83e9cefc7b68.png" width="23">)<br><br>
Due to security reasons, this repo does not include any scripts (VBS files) that include the actual SAP execution codes that were recorded using SAP's recording feature.<br>

# ⬇️English Version Manual⬇️

2021.12.29 by Guyoung

## How it works<br>
1. Run exe file (or pyw) in the folder<br>
2. A window will pop up. Follow the instructions in the pop up window.<br>
3. Enter data as necessary. Tasks that are date-sensitive will automatically have today's date as the default value. For example, "export-then-create-pivot-table.py" will automatically create a pivot table that only contains this month's data (from first day of the month to today).
<br>
☞ py files The exe files work only with the vbs files<br>
☞ pyfiles folder is a folder containing code source that is intended to help others automate their work.<br><br>

<img width="693" alt="image" src="https://user-images.githubusercontent.com/79275984/222053369-6bb9c86a-6c62-4e3f-9e7c-d4e33561ffa6.png">
<br>
There will be VBS files in the "vbs" folder which includes pre-recorded SAP action scripts. (if you don't know how to record your own VBS file with SAP, google "how to record on SAP")

## Things I liked about automating SAP tasks
During my internship, I had tons of data entry, exporting, and reporting tasks that used Excel and SAP. After a few months of doing those, I recognized there were patterns. They were repetitive and I saw the potential to make a Python bot do the work. I thought it could be a fun little project. I essentially used the recording feature on the SAP to produce vbs files, and then created variables on Python so that those vbs codes could be re-run with different variables (such as dates, material numbers, and customer codes)<br><br>
One month into re-learning Python, I was able to automate a few tasks that interns do regularly. I shared the automation bot (the py files you see in the top of this repo) to my manager and other team members - my manager liked it. She told me to host a "learning session"(a thing in Continental where you teach others great skills and/or knowledge) where I get to teach the finance team how to automate repetitive tasks!<br><br>
It was a great experience, and although the codes being super-simple, this is when I really started to put more effort into learning Python!

## Wanna make your own automation? Try:
Download any Python IDE (ex: PyCharm, VS Code, Jupyter, etc.) and watch this playlist:<br> https://www.youtube.com/watch?v=7-kXlg59Fn0&list=PLMACxP0Btc7c9tEGvr5FKsJBwd3y1R6eI
<br>
<br>This series of videos taught me how to automate repetitive SAP data exporting/entry tasks.

# ⬇️Korean Version Manual⬇️

2021.12.29 by 구영

## 작동 방법<br>
1.	exe 파일 (혹은 pyw) 을 실행한다<br>
2.	창에 있는 안내 대로 한다<br>
<br>
☞'업무자동화'폴더 자체를 본인의 바탕화면에 두고 쓰셔도 작동합니다.<br>
☞ py files폴더가 없어도 exe파일들은 작동합니다<br>
☞ py files폴더는 다른 분들의 업무 자동화에 도움되라고 둔, 코드 원본이 담긴 폴더입니다<br><br>

<img width="693" alt="image" src="https://user-images.githubusercontent.com/79275984/219856126-c74f8517-62b6-4dec-8f5f-18e360914f29.png">
<br>
vbs폴더 안에는 이미 record된 SAP 동작들을 재현해 줄 비쥬얼 베이직 스크립트가 들어있습니다.

## 나도 자동화하기
1.	PyCharm Community Edition(무료공개판) 다운받은 뒤 실행합니다. PyCharm의 interpreter를 Anaconda나 Python 3.X버전으로 설정한다.
PyCharm Community Edition (무료공개판) 다운 후 설치할 때 꼭 community edition옵션을 택할 것(일반버전은 지불+관리자 권한 필요)
구글링해서 PyCharm Community Edition 설치하는 법 치면 엄청 많이 나올 것.
2.	[파이썬으로 SAP Script 실행하기 (날짜, 파일명 등 입력과 함께)]
https://kminito.tistory.com/10
위 링크의 글을 정독하고 따라하자. 끝.

위 링크에 나오는 아래 형태의 코드가 제일 핵심이다.
> os.system("vbs\\DeliveryAmount0.vbs {}".format(customer))


## 후기
SAP에 많이 존재하는 귀찮은 단순반복업무들의 패턴만 파악했다면, 위 링크 보고 1주 안에 짧지만 귀찮은 업무 하나 자동화 가능.<br>
학교에선 코딩 입문 수업 진짜 재미없었는데 이렇게 회사에 있는 업무 실제로 자동화 시키니깐 보람차고 재밌다.<br>
좀 늦게 알아버려서 아쉬울 따름.<br>
처음에 파이썬 설정하고 그럴 때 삽질 많이 할거임. 이미 할 줄 안다면 물론 삽질x<br>
물론 코드 짤 때는 매 순간 삽질할 수 있음.<br>
++추가: 인턴쉽 끝나고 나서 추후 들어온 인턴 동생한테 물어보니 UiPath로도 더 쉽게 자동화할 수 있다고 한다! 회사에서 해당 프로그램을 제공한다면 자동화하기 좋은 기회가 될 듯.
