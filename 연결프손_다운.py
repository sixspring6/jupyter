#!/usr/bin/env python
# coding: utf-8

# In[1]:


import win32com.client
import xlwings as xw
import pandas as pd
import win32api

from tkinter import *
from tkinter import simpledialog
from datetime import datetime, timedelta

# 데이터를 생성합니다.
data = {'시작일': [202301, 201001],
        '종료일': [202308, 202212],
        '구분': ['당기', '전기']}

# 데이터프레임을 생성합니다.
df = pd.DataFrame(data)

# Tkinter를 초기화합니다.
root = Tk()
root.withdraw()  # Tkinter 창 숨기기

# 사용자로부터 yyyy-mm 형식의 날짜를 메시지 박스로 입력 받습니다.
input_date = simpledialog.askstring("날짜 입력", "yyyy-mm 형식의 날짜를 입력하세요:")

# 입력 받은 날짜를 yyyy-mm 형식으로 파싱합니다.
user_date = datetime.strptime(input_date, '%Y-%m')

# '시작일' 열을 입력 받은 날짜의 연도의 1월로 변경합니다.
df.loc[0, '시작일'] = user_date.strftime('%Y') + '01'

# '종료일' 열의 0번 행을 입력 받은 날짜로 변경합니다.
df.loc[0, '종료일'] = user_date.strftime('%Y%m')

# 입력 받은 날짜의 전년도 12월을 계산합니다.
previous_year_date = user_date - timedelta(days=365)

# '종료일' 열의 1번 행을 전년도 12월로 변경합니다.
df.loc[1, '종료일'] = previous_year_date.strftime('%Y12')

# 수정된 데이터프레임을 출력합니다.
print(df)

# Tkinter 창을 닫습니다.
root.destroy()


app = xw.App(visible=False)
wb = xw.Book('C:/연결프로젝트손익명세/작업파일1.xlsx')
sht=wb.sheets['sheet1']
df = sht.range('a1').options(pd.DataFrame,expand='table',Header=1,index=False).value

손익명세 = df['종료일'].iloc[0]

# Excel 파일 열기


# 시트 선택
sheet = wb.sheets['현지화전기']

# 시트의 내용 삭제
sheet.clear_contents()

sheet = wb.sheets['원화전기']

# 시트의 내용 삭제
sheet.clear_contents()

# 시트 선택
sheet = wb.sheets['현지화당기']

# 시트의 내용 삭제
sheet.clear_contents()

sheet = wb.sheets['원화당기']

# 시트의 내용 삭제
sheet.clear_contents()

sheet = wb.sheets['손익명세']

# 시트의 내용 삭제
sheet.clear_contents()


#작업파일 비우고 다시 닫음
wb.close()
app.quit()

print("작업파일을 비웠습니다.")





import os

folder_path = 'c:/연결프로젝트손익명세'  # 파일이 있는 폴더 경로
search_string = '원화당기'  # 포함된 문자열

for filename in os.listdir(folder_path):
    if search_string in filename:
        file_path = os.path.join(folder_path, filename)
        os.remove(file_path)
        print(f"파일 {file_path}를 삭제했습니다.")




folder_path = 'c:/연결프로젝트손익명세'  # 파일이 있는 폴더 경로
search_string = '현지화당기'  # 포함된 문자열

for filename in os.listdir(folder_path):
    if search_string in filename:
        file_path = os.path.join(folder_path, filename)
        os.remove(file_path)
        print(f"파일 {file_path}를 삭제했습니다.")


        
folder_path = 'c:/연결프로젝트손익명세'  # 파일이 있는 폴더 경로
search_string = '원화전기'  # 포함된 문자열

for filename in os.listdir(folder_path):
    if search_string in filename:
        file_path = os.path.join(folder_path, filename)
        os.remove(file_path)
        print(f"파일 {file_path}를 삭제했습니다.")


folder_path = 'c:/연결프로젝트손익명세'  # 파일이 있는 폴더 경로
search_string = '현지화전기'  # 포함된 문자열

for filename in os.listdir(folder_path):
    if search_string in filename:
        file_path = os.path.join(folder_path, filename)
        os.remove(file_path)
        print(f"파일 {file_path}를 삭제했습니다.")


        
folder_path = 'c:/연결프로젝트손익명세'  # 파일이 있는 폴더 경로
search_string = '통합'  # 포함된 문자열

for filename in os.listdir(folder_path):
    if search_string in filename:
        file_path = os.path.join(folder_path, filename)
        os.remove(file_path)
        print(f"파일 {file_path}를 삭제했습니다.")
        
        
folder_path = 'c:/연결프로젝트손익명세'  # 파일이 있는 폴더 경로
search_string = '손익'  # 포함된 문자열

for filename in os.listdir(folder_path):
    if search_string in filename:
        file_path = os.path.join(folder_path, filename)
        os.remove(file_path)
        print(f"파일 {file_path}를 삭제했습니다.")

        

#파일 삭제 후 sap


SapGuiAuto = win32com.client.GetObject('SAPGUI')


# 현재 SAP 세션 가져오기
Application = SapGuiAuto.GetScriptingEngine

Connection = Application.Children(Application.connections.count - 1)


session = Connection.Children(0)

for index, row in df.iterrows():
    var1 = row['시작일']
    var2 = row['종료일']
    var3 = row['구분']



    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/usr/cntlIMAGE_CONTAINER/shellcont/shell/shellcont[0]/shell").selectedNode = "F00100"
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZHFIGLR3040"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 8
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "8"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
    session.findById("wnd[0]/usr/ctxtS_SPMON-LOW").text = var1
    session.findById("wnd[0]/usr/ctxtS_SPMON-HIGH").text = var2
    session.findById("wnd[0]/usr/radP_A").setFocus
    session.findById("wnd[0]").sendVKey (8)
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").pressToolbarContextButton ("&MB_EXPORT")
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectContextMenuItem ("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\연결프로젝트손익명세"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "원화"+var3+".xlsx"
    session.findById("wnd[1]").sendVKey (0)
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").setCurrentCell (12,"LTEXT")
    session.findById("wnd[0]").sendVKey (3)
    session.findById("wnd[0]").sendVKey (3)
    
    import time

    def wait():
        time.sleep(5)
        print("5초가 지났습니다!")
    wait()

    
#     import psutil
#     import os


#     for proc in psutil.process_iter():

#         try:
#             # 프로세스 이름을 가져옵니다.
#             process_name = proc.name()

#             # 만약 엑셀이 실행 중이면,
#             if 'EXCEL' in process_name.upper():
#                 # 프로세스를 종료합니다.
#                 os.kill(proc.pid, 9)
#                 print(f"{process_name} 프로세스를 종료했습니다.")

#         except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#             # 종료할 수 없는 프로세스는 건너뜁니다.
#             pass

    # 함수 호출
    

    session.findById("wnd[0]").maximize
    session.findById("wnd[0]/tbar[0]/okcd").text = "ZHFIGLR3040"
    session.findById("wnd[0]").sendVKey (0)
    session.findById("wnd[0]/tbar[1]/btn[17]").press()
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 8
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "8"
    session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
    session.findById("wnd[0]/usr/ctxtS_SPMON-LOW").text = var1
    session.findById("wnd[0]/usr/ctxtS_SPMON-HIGH").text = var2
    session.findById("wnd[0]/usr/radP_B").setFocus
    session.findById("wnd[0]/usr/radP_B").select()
    session.findById("wnd[0]").sendVKey (8)
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").pressToolbarContextButton ("&MB_EXPORT")
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectContextMenuItem ("&XXL")
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\연결프로젝트손익명세"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "현지화"+var3+".xlsx"
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 17
    session.findById("wnd[1]").sendVKey (0)
    session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").setCurrentCell (12,"LTEXT")
    session.findById("wnd[0]").sendVKey (3)
    session.findById("wnd[0]").sendVKey (3)
    
    wait()
    
#     for proc in psutil.process_iter():

#         try:
#             # 프로세스 이름을 가져옵니다.
#             process_name = proc.name()

#             # 만약 엑셀이 실행 중이면,
#             if 'EXCEL' in process_name.upper():
#                 # 프로세스를 종료합니다.
#                 os.kill(proc.pid, 9)
#                 print(f"{process_name} 프로세스를 종료했습니다.")

#         except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#             # 종료할 수 없는 프로세스는 건너뜁니다.
#             pass



#손익명세 다운

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "zufiumr0430"
session.findById("wnd[0]").sendVKey (0)
session.findById("wnd[0]/tbar[1]/btn[17]").press()
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 5
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "5"

session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
session.findById("wnd[0]/usr/ctxtPA_PERVV").text = 손익명세
session.findById("wnd[0]/usr/ctxtPA_PERVV").caretPosition = 6
session.findById("wnd[0]").sendVKey (8)
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").pressToolbarContextButton ("&MB_EXPORT")
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectContextMenuItem ("&XXL")
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\연결프로젝트손익명세"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "손익명세.xlsx"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 2
session.findById("wnd[1]/tbar[0]/btn[0]").press()

session.findById("wnd[0]").sendVKey (3)
session.findById("wnd[0]").sendVKey (3)


#별도통합손익명세 다운

# session.findById("wnd[0]").maximize
# session.findById("wnd[0]/tbar[0]/okcd").text = "ZUFIUMR0460"
# session.findById("wnd[0]").sendVKey (0)
# session.findById("wnd[0]/usr/ctxtPA_PERVV").text = 손익명세
# session.findById("wnd[0]/usr/ctxtPA_PERVV").caretPosition = 7
# session.findById("wnd[0]").sendVKey (8)
# session.findById("wnd[1]").sendVKey (0)
# session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").pressToolbarContextButton ("&MB_EXPORT")
# session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectContextMenuItem ("&XXL")
# session.findById("wnd[1]/tbar[0]/btn[0]").press()
# session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\연결프로젝트손익명세"
# session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "통합.XLSX"
# session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 2
# session.findById("wnd[1]").sendVKey (0)
# # session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").setCurrentCell 1,""
# session.findById("wnd[0]").sendVKey (3)
# session.findById("wnd[0]").sendVKey (3)


session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "ZUFIUMR0460"
session.findById("wnd[0]").sendVKey (0)
session.findById("wnd[0]").sendVKey (17)
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "0"
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").doubleClickCurrentCell()
session.findById("wnd[0]/usr/ctxtPA_PERVV").text = 손익명세
session.findById("wnd[0]").sendVKey (8)
session.findById("wnd[1]").sendVKey (0)
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").pressToolbarContextButton ("&MB_EXPORT")
session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").selectContextMenuItem ("&XXL")
session.findById("wnd[1]/tbar[0]/btn[0]").press()
session.findById("wnd[1]/usr/ctxtDY_PATH").text = "C:\\연결프로젝트손익명세"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "통합.XLSX"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 2
session.findById("wnd[1]").sendVKey (0)
# session.findById("wnd[0]/shellcont/shell/shellcont[1]/shell").setCurrentCell 1,""
session.findById("wnd[0]").sendVKey (3)
session.findById("wnd[0]").sendVKey (3)


    
    
import time

def wait():
    time.sleep(5)
    print("다운완료")
    
# import psutil
# import os


# for proc in psutil.process_iter():

#     try:
#         # 프로세스 이름을 가져옵니다.
#         process_name = proc.name()

#         # 만약 엑셀이 실행 중이면,
#         if 'EXCEL' in process_name.upper():
#             # 프로세스를 종료합니다.
#             os.kill(proc.pid, 9)
#             print(f"{process_name} 프로세스를 종료했습니다.")

#     except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
#         # 종료할 수 없는 프로세스는 건너뜁니다.
#         pass

# # 함수 호출
# wait()

