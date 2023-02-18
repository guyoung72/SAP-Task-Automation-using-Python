# SAP-Task-Automation-using-Python
How to automate SAP tasks using its recording features and Python codes

업무자동화 exe 실행법 / 작동원리 / 코딩하는 법<br>

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
물론 코드 짤 때는 매 순간 삽질할 수 있음.
