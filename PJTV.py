#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import pyodbc
import matplotlib.pyplot as plt


conn_str = (
    r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
    r'DBQ=d:\archive\손익\profit.accdb;'  # 실제 파일 경로로 변경
)
conn = pyodbc.connect(conn_str)
cursor = conn.cursor()


sql = "SELECT 기간, 손익센터명, 매출, 이익, 별도연결 FROM 손익2"  # 실제 테이블명으로 변경
df = pd.read_sql(sql, conn)
df


# In[2]:


df = df[df['손익센터명'] == 'HMC양재사옥']
df = df[df['별도연결'] == '연결']



# '대상월' 컬럼을 datetime 타입으로 변환 후, 연도와 월을 기준으로 정렬
df['기간'] = pd.to_datetime(df['기간'])
df.sort_values(by=['기간'], inplace=True)

# 연속된 달의 누계 값 차이를 계산하여 순수 월별 값 도출
df['순매출'] = df['매출'].diff().fillna(df['매출'])
df['순이익'] = df['이익'].diff().fillna(df['이익'])


# 첫 달의 경우 diff()로 인해 NaN이 나올 수 있으므로, 원래 값을 유지합니다 (fillna 사용)

# 각 연도별 첫 달의 순수 값 조정이 필요할 수 있습니다.
# 예를 들어, 각 연도의 시작에서는 diff()가 이전 연도의 마지막과 비교되므로,
# 첫 달의 순수 값이 잘못 계산될 수 있습니다. 이를 조정하기 위해:
for year in df['기간'].dt.year.unique():
    first_month_idx = df[df['기간'].dt.year == year].index.min()
    df.loc[first_month_idx, '순매출'] = df.loc[first_month_idx, '매출']
    df.loc[first_month_idx, '순이익'] = df.loc[first_month_idx, '이익']


# 순수 월별 값 시각화
plt.figure(figsize=(12, 8))
plt.plot(df['기간'], df['순매출'], label='순매출')
plt.plot(df['기간'], df['순이익'], label='순이익')

plt.legend()
plt.title('월별 순매출, 순이익')
plt.xlabel('대상월')
plt.ylabel('금액')
plt.xticks(rotation=45)
plt.show()


# In[3]:


df.to_excel('c:/x.xlsx')


# In[10]:


import pandas as pd
import matplotlib.pyplot as plt
import xlsxwriter

# DataFrame 'df_project_a'를 예로 들어 결과를 엑셀 파일에 저장
excel_file = '/project_a_analysis.xlsx'
sheet_name = 'Analysis'

# Excel writer 객체 생성 with xlsxwriter
writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')

# DataFrame을 엑셀 시트에 쓰기
df.to_excel(writer, sheet_name=sheet_name, index=False)

# xlsxwriter workbook과 worksheet 객체 가져오기
workbook = writer.book
worksheet = writer.sheets[sheet_name]

# 그래프 생성 및 이미지 파일로 저장
fig, ax = plt.subplots(figsize=(10, 6))
df_project_a.plot(x='기간', y=['순매출', '순이익'], kind='line', ax=ax)
plt.title('프로젝트 a의 월별 순매출, 순이익')
plt.xlabel('기간')
plt.ylabel('금액')
plt.legend()
plt.tight_layout()

# 그래프 이미지 파일로 저장
graph_image = '/mnt/data/project_a_graph.png'
plt.savefig(graph_image)

# 엑셀 파일에 그래프 이미지 삽입
worksheet.insert_image('H1', graph_image, {'x_scale': 0.5, 'y_scale': 0.5})

# Excel 파일 저장
writer.save()


# In[ ]:




