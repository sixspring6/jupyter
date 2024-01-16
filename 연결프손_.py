<<<<<<< HEAD



=======
>>>>>>> b6f02b0e521accf8e8639f7cfeeac1e9a28ff282
import win32com.client
import xlwings as xw
import pandas as pd
import numpy as np




data = {
    '주요pjt': ['AAID1902', 'TAAPL2101','TAOUS2201','TANUS2201'],
    '종속pjt': ['TAAID2101', 'AAPL2101','AOUS2201','ANUS2201'],
    '지분율': ['0.7', '0.55','1','1'],
    '환율': ['0.0868', '308.46','1344.8','1344.8'],
}

data_1 = {
    '주요pjt': ['AAID2102'],
    '종속pjt': ['TAAID2102'],
    '환율': ['0.0868'],}





consolidation = pd.DataFrame(data)

consolidation['지분율'] = consolidation['지분율'].astype(float)
consolidation['환율'] = consolidation['환율'].astype(float)
consolidation = consolidation.reset_index()


consolidation_1 = pd.DataFrame(data_1)


consolidation_1['환율'] = consolidation_1['환율'].astype(float)
consolidation_1 = consolidation_1.reset_index()






wb = xw.Book('C:/연결프로젝트손익명세/통합.xlsx')
sht=wb.sheets[0]
       #파일명어떻게 할지 
통합실행 = sht.range('a1:cg1000').options(pd.DataFrame,Header=1,index=False).value
print("파일 로딩을 시작합니다")
wb.close()


통합실행 = 통합실행.dropna(subset=['회사코드'])

통합실행 = 통합실행.rename(columns={'프로젝트코드':'손익센터'})

통합실행['예정원가율'] = 통합실행['예정원가율']/100

통합실행['예정원가율(KRW)'] = 통합실행['예정원가율(KRW)']/100



app = xw.App(visible=False)
wb = xw.Book('C:/연결프로젝트손익명세/원화당기.xlsx')
sht=wb.sheets[0]
       #파일명어떻게 할지 
원화당기 = sht.range('a1:af1000').options(pd.DataFrame,Header=1,index=False).value
print("원화당기 파일 로딩을 시작합니다")
wb.close()
app.quit()

app = xw.App(visible=False)
wb = xw.Book('C:/연결프로젝트손익명세/원화전기.xlsx')
sht=wb.sheets[0]
       #파일명어떻게 할지 
원화전기 = sht.range('a1:af1000').options(pd.DataFrame,Header=1,index=False).value
print("원화전기 파일 로딩을 시작합니다")
wb.close()
app.quit()

app = xw.App(visible=False)
wb = xw.Book('C:/연결프로젝트손익명세/현지화당기.xlsx')
sht=wb.sheets[0]
       #파일명어떻게 할지 
현지화당기 = sht.range('a1:af1000').options(pd.DataFrame,Header=1,index=False).value
print("현지화당기 파일 로딩을 시작합니다")
wb.close()
app.quit()

app = xw.App(visible=False)
wb = xw.Book('C:/연결프로젝트손익명세/현지화전기.xlsx')
sht=wb.sheets[0]
print("현지화전기 파일 로딩을 시작합니다")
현지화전기 = sht.range('a1:af1000').options(pd.DataFrame,Header=1,index=False).value
wb.close()
app.quit()


app = xw.App(visible=False)
wb = xw.Book('C:/연결프로젝트손익명세/손익명세.xlsx')
sht=wb.sheets[0]
손익명세 = sht.range('a1:bq400').options(pd.DataFrame,Header=1,index=False).value
wb.close()
app.quit()
print("손익명세파일 로딩을 완료했습니다")

app = xw.App(visible=False)
wb = xw.Book('C:/연결프로젝트손익명세/작업파일1.xlsx')

print("작업파일1을 불러들였습니다.")

sheet_name = 'new'   #new 시트 없애기

import pythoncom


try:
    sheet = wb.sheets[sheet_name]
    sheet.delete()
    print("new 시트를 삭제했습니다")
    
except (KeyError, pythoncom.com_error):
    pass








    
    



손익명세 = 손익명세.rename(columns={'프로젝트코드':'손익센터'})



손익명세.drop(['원가분석(KRW)','매출분석(KRW)','고객명','Project프로파일','상태','계약명','고객명','국내/해외','매출분석','원가분석','전표 번호','당기손익','주관','사업분야','통화'],axis=1,inplace=True)

손익명세['예정원가율'] = 손익명세['예정원가율']/100
손익명세['예정원가율(KRW)'] = 손익명세['예정원가율(KRW)']/100

손익명세['작업진행율'] = 손익명세['작업진행율']/100
손익명세['작업진행율(KRW)'] = 손익명세['작업진행율(KRW)']/100



원화당기 = 원화당기.drop('통화', axis=1)



원화당기['정리']= 원화당기['당기매출'].abs()+원화당기['당기원가'].abs()
원화당기 = 원화당기.drop(원화당기[원화당기['정리'] <= 0].index)

원화전기['정리']= 원화전기['당기매출'].abs()+원화전기['당기원가'].abs()
원화전기 = 원화전기.drop(원화전기[원화전기['정리'] <= 0].index)

현지화당기['정리']= 현지화당기['당기매출'].abs()+현지화당기['당기원가'].abs()
현지화당기 = 현지화당기.drop(현지화당기[현지화당기['정리'] <= 0].index)

현지화전기['정리']= 현지화전기['당기매출'].abs()+현지화전기['당기원가'].abs()
현지화전기 = 현지화전기.drop(현지화전기[현지화전기['정리'] <= 0].index)




원화당기.to_excel('c:/연결프로젝트손익명세/작성.xlsx',sheet_name='pot_table')

현지화당기.to_excel('c:/연결프로젝트손익명세/작성1.xlsx',sheet_name='pot_table')

현지화전기.to_excel('c:/연결프로젝트손익명세/작성3.xlsx',sheet_name='pot_table')

원화전기.to_excel('c:/연결프로젝트손익명세/작성2.xlsx',sheet_name='pot_table')


원화전기 =원화전기.rename(columns={'당기원가':'전기원가'})
원화전기 = 원화전기.rename(columns={'당기매출':'전기매출'})
원화전기 = 원화전기.rename(columns={'당기손익':'전기손익'})
원화전기 = 원화전기.rename(columns={'하자보수충당부채전입액':'전기 하자보수충당부채전입액'})
원화전기 = 원화전기.rename(columns={'손실충당금전입액':'전기 손실충당금전입액'})
원화전기 = 원화전기.rename(columns={'진행율대상원가':'전기 진행율대상원가'})

원화전기 = 원화전기[['손익센터','전기매출','전기원가','전기손익','전기 하자보수충당부채전입액','전기 손실충당금전입액',
             '전기 진행율대상원가']]

원화전기 = 원화전기.groupby(['손익센터'])[['전기매출','전기원가','전기손익','전기 하자보수충당부채전입액','전기 손실충당금전입액',
             '전기 진행율대상원가']].sum()

현지화전기 = 현지화전기.rename(columns={'당기원가':'전기원가_현지화'})
현지화전기 = 현지화전기.rename(columns={'당기매출':'전기매출_현지화'})
현지화전기 = 현지화전기.rename(columns={'당기손익':'전기손익_현지화'})
현지화전기 = 현지화전기.rename(columns={'하자보수충당부채전입액':'전기 하자보수충당부채전입액_현지화'})
현지화전기 = 현지화전기.rename(columns={'손실충당금전입액':'전기 손실충당금전입액_현지화'})
현지화전기 = 현지화전기.rename(columns={'진행율대상원가':'전기 진행율대상원가_현지화'})
현지화전기 = 현지화전기[['손익센터','통화','전기매출_현지화','전기원가_현지화','전기손익_현지화','전기 하자보수충당부채전입액_현지화'
               ,'전기 손실충당금전입액_현지화','전기 진행율대상원가_현지화']]

현지화전기 = 현지화전기.groupby(['손익센터','통화'])[['전기매출_현지화','전기원가_현지화','전기손익_현지화','전기 하자보수충당부채전입액_현지화','전기 손실충당금전입액_현지화',
             '전기 진행율대상원가_현지화']].sum()







현지화당기 = 현지화당기.rename(columns={'당기원가':'당기원가_현지화'})
현지화당기 = 현지화당기.rename(columns={'당기매출':'당기매출_현지화'})
현지화당기 = 현지화당기.rename(columns={'당기손익':'당기손익_현지화'})
현지화당기 = 현지화당기.rename(columns={'하자보수충당부채전입액':'당기 하자보수충당부채전입액_현지화'})
현지화당기 = 현지화당기.rename(columns={'손실충당금전입액':'당기 손실충당금전입액_현지화'})
현지화당기 = 현지화당기.rename(columns={'진행율대상원가':'당기 진행율대상원가_현지화'})
현지화당기 = 현지화당기[['손익센터','당기매출_현지화','당기원가_현지화','당기손익_현지화',
               '당기 하자보수충당부채전입액_현지화','당기 손실충당금전입액_현지화','당기 진행율대상원가_현지화']]

원화당기 = 원화당기.dropna(subset=['손익센터'])
현지화당기 = 현지화당기.dropna(subset=['손익센터'])

원화당기 = 원화당기.drop(원화당기[(원화당기['손익센터'].duplicated()) & (원화당기['CoCode'] == 'HD00')].index)




merge_outer1 = pd.merge(원화당기,원화전기, how='outer',on='손익센터') #1차 merge

현지화전기 = 현지화전기.reset_index()




merge_outer2 = pd.merge(merge_outer1,현지화전기, how='outer',on='손익센터')

merge_outer2.to_excel('c:/연결프로젝트손익명세/here.xlsx')

merge_outer3 = pd.merge(merge_outer2,현지화당기, how='outer',on='손익센터')

merge_outer3.to_excel('c:/연결프로젝트손익명세/1차.xlsx',sheet_name='pot_table')









현지화전기.to_excel('c:/연결프로젝트손익명세/wn1.xlsx')





merge = pd.merge(merge_outer3,손익명세, how='outer',on='손익센터')



merge = merge.dropna(subset=['CoCode'])
merge.to_excel('c:/연결프로젝트손익명세/1차.xlsx',sheet_name='pot_table')

merge['전기매출_현지화'] = merge['전기매출_현지화'].fillna(0)
merge['전기원가_현지화'] = merge['전기원가_현지화'].fillna(0)
merge['전기손익_현지화'] = merge['전기손익_현지화'].fillna(0)
merge['전기 진행율대상원가_현지화'] = merge['전기 진행율대상원가_현지화'].fillna(0)
merge['전기 하자보수충당부채전입액_현지화'] = merge['전기 하자보수충당부채전입액_현지화'].fillna(0)
merge['전기 손실충당금전입액_현지화'] = merge['전기 손실충당금전입액_현지화'].fillna(0)
merge['전기매출'] = merge['전기매출'].fillna(0)
merge['전기원가'] = merge['전기원가'].fillna(0)
merge['전기 진행율대상원가'] = merge['전기 진행율대상원가'].fillna(0)
merge['전기 하자보수충당부채전입액'] = merge['전기 하자보수충당부채전입액'].fillna(0)
merge['전기 손실충당금전입액'] = merge['전기 손실충당금전입액'].fillna(0)






현지화전기.to_excel('c:/연결프로젝트손익명세/wn2.xlsx')

merge["누계매출_현지화"]= np.nan

merge['누계매출_현지화']= merge['전기매출_현지화']+merge['당기매출_현지화']

merge["누계원가_현지화"]= np.nan

merge['누계원가_현지화']= merge['전기원가_현지화']+merge['당기원가_현지화']

merge["누계손익_현지화"]= np.nan

merge['누계손익_현지화']= merge['전기손익_현지화']+merge['당기손익_현지화']







merge["누계진행율대상원가_현지화"]= np.nan

merge['누계진행율대상원가_현지화']= merge['전기 진행율대상원가_현지화']+merge['당기 진행율대상원가_현지화']

merge["누계하자보수충당부채전입액_현지화"]= np.nan

merge['누계하자보수충당부채전입액_현지화']= merge['전기 하자보수충당부채전입액_현지화']+merge['당기 하자보수충당부채전입액_현지화']

merge["누계손실충당금전입액_현지화"]= np.nan

merge['누계손실충당금전입액_현지화']= merge['전기 손실충당금전입액_현지화']+merge['당기 손실충당금전입액_현지화']

merge["누계차이_현지화"]= np.nan

merge['누계차이_현지화']= merge['누계원가_현지화']-merge['누계진행율대상원가_현지화']



merge["누계매출"]= np.nan

merge['누계매출']= merge['전기매출']+merge['당기매출']


merge["누계원가"]= np.nan

merge['누계원가']= merge['당기원가']+merge['전기원가']

merge["누계손익"]= np.nan

merge['누계손익']= merge['누계매출']-merge['누계원가']






merge["누계진행율대상원가"]= np.nan

merge['누계진행율대상원가']= merge['전기 진행율대상원가']+merge['진행율대상원가']


merge["누계차이"]= np.nan

merge['누계차이']= merge['누계원가']-merge['누계진행율대상원가']


merge["누계하자보수충당부채전입액"]= np.nan

merge['누계하자보수충당부채전입액']= merge['하자보수충당부채전입액']+merge['전기 하자보수충당부채전입액']


merge["누계손실충당금전입액"]= np.nan

merge['누계손실충당금전입액']= merge['손실충당금전입액']+merge['전기 손실충당금전입액']

merge["당기검증"]= np.nan

merge['당기검증']= merge['당기매출'].abs()+merge['당기원가'].abs()

merge = merge.drop(merge[merge['당기검증'] <= 0].index)

merge['구분'] = "해당없음"

merge.loc[merge['신규'] == '예', '구분'] = '신규'
merge.loc[merge['이월'] == '예', '구분'] = '이월'
merge.loc[merge['정산'] == '예', '구분'] = '정산'




merge['하자율_현지화'] = np.where(merge['누계하자보수충당부채전입액_현지화'] != 0, 
                                  merge['누계하자보수충당부채전입액_현지화'] / merge['누계매출_현지화'], 0)

merge['계상미수(KRW)'] = merge['계상미수(KRW)'].fillna(0)
merge['계상선수(KRW)'] = merge['계상선수(KRW)'].fillna(0)




merge['결산계상검증'] = merge['계상미수(KRW)']+merge['계상선수(KRW)']-merge['결산계상 미수금']+merge['결산계상 선수금']





# 조건에 따라 열의 값을 채우기
merge.loc[(merge['주관'] == '플랜트 사업본부(화공)') & (merge['구분'] == '해당없음') & (merge['계약고'].isna()), '계약고'] = merge.loc[(merge['주관'] == '플랜트 사업본부(화공)') & (merge['구분'] == '해당없음') & (merge['계약고'].isna()), '누계매출_현지화']




#df.loc[(df['a'] == 'america') & (df['b'] == 'south') & (df['c'].isna()), 'c'] 
#= df.loc[(df['a'] == 'america') & (df['b'] == 'south') & (df['c'].isna()), 'd']



merge = merge.reindex(columns=['주관','지역','CoCode','지역2','손익센터','구분','손익센터명','사업분야','고객명','시작일','종료일','통화','환율',
                               '계약고','총공사예정비','예정원가율','작업진행율','전기매출_현지화','당기매출_현지화','누계기성청구합계',
                               '계상미수','계상선수',   '누계매출_현지화','계약잔고','당기원가_현지화',                           
                               '당기 하자보수충당부채전입액_현지화','당기 손실충당금전입액_현지화','전기원가_현지화','누계원가_현지화',
                                                            
                               '누계진행율대상원가_현지화','누계차이_현지화',
                               '누계하자보수충당부채전입액_현지화','누계손실충당금전입액_현지화','용지비_현지화','검증_현지화','당기손익_현지화','누계손익_현지화',
                               
                               
                               '계약고(KRW)','총공사예정비(KRW)','예정원가율(KRW)','작업진행율(KRW)','전기매출','당기매출',
                               '누계기성청구합계(KRW)','계상미수(KRW)',
                               '계상선수(KRW)','누계매출','계약잔고(KRW)','당기원가','하자보수충당부채전입액','손실충당금전입액','전기원가','누계원가',
                               '누계진행율대상원가','누계차이','누계하자보수충당부채전입액','누계손실충당금전입액','용지비','검증','당기손익'
                               ,'누계손익','하자율_현지화','당기검증','결산계상검증',
                               
                               
                               
                               
                               
                               
                               '진행율대상원가','하자보수충당부채전입액','손실충당금전입액','매출채권','매입채무',
                               '선수금','선급금','유보금(부채)','유보금(자산)','진행단계(입찰)명','진행단계(수행)명',
                               '진행단계(하자)명','전기손익','전기 하자보수충당부채전입액',
                               '전기 손실충당금전입액','전기손익_현지화',
                               '전기 하자보수충당부채전입액_현지화','전기 손실충당금전입액_현지화',
                               '회사코드','공사/용역','매출유형','EPC구분','신규',
                               '이월','정산','착공일','착공승인일',
                               '총공사예정비(전체)','총공사예정비 (전체) KRW',
                               '전기까지수입누계액(KRW)','실수입','실수입(KRW)',
                               '매출계(KRW)','당기말까지수입누계액(KRW)','전기까지원가(KRW)',
                               '전기까지원가','당기말까지원가누계',
                               '재료비','재료비(KRW)','노무비','노무비(KRW)','외주비','외주비(KRW)','경비','경비(KRW)',
                               '원가계(KRW)','당기말까지원가누계(KRW)','당기손익(KRW)','전기까지수입누계액','매출계','당기말까지수입누계액','원가계'])

#자료 합치기 별도 통합이므로, 단순 SUM이 가능하다.

merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '당기원가'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '당기원가'] + merge.loc[merge['손익센터'] == 'ANID2202', '당기원가'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '전기원가'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '전기원가'] + merge.loc[merge['손익센터'] == 'ANID2202', '전기원가'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계원가'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계원가'] + merge.loc[merge['손익센터'] == 'ANID2202', '누계원가'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계진행율대상원가'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계진행율대상원가'] + merge.loc[merge['손익센터'] == 'ANID2202', '누계진행율대상원가'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계손익'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계손익'] + merge.loc[merge['손익센터'] == 'ANID2202', '누계손익'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '당기매출'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '당기매출'] + merge.loc[merge['손익센터'] == 'ANID2202', '당기매출'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '전기매출'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '전기매출'] + merge.loc[merge['손익센터'] == 'ANID2202', '전기매출'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계기성청구합계(KRW)'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계기성청구합계(KRW)'] + merge.loc[merge['손익센터'] == 'ANID2202', '누계기성청구합계(KRW)'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계매출'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계매출'] + merge.loc[merge['손익센터'] == 'ANID2202', '누계매출'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '당기원가_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '당기원가_현지화'] + merge.loc[merge['손익센터'] == 'ANID2202', '당기원가_현지화'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '전기원가_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '전기원가_현지화'] + merge.loc[merge['손익센터'] == 'ANID2202', '전기원가_현지화'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계원가_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계원가_현지화'] + merge.loc[merge['손익센터'] == 'ANID2202', '누계원가_현지화'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계진행율대상원가_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계진행율대상원가_현지화'] + merge.loc[merge['손익센터'] == 'ANID2202', '누계진행율대상원가_현지화'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계손익_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계손익_현지화'] + merge.loc[merge['손익센터'] == 'ANID2202', '누계손익_현지화'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '당기매출_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '당기매출_현지화'] + merge.loc[merge['손익센터'] == 'ANID2202', '당기매출_현지화'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '전기매출_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '전기매출_현지화'] + merge.loc[merge['손익센터'] == 'ANID2202', '전기매출_현지화'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계기성청구합계'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계기성청구합계'] + merge.loc[merge['손익센터'] == 'ANID2202', '누계기성청구합계'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계매출_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '누계매출_현지화'] + merge.loc[merge['손익센터'] == 'ANID2202', '누계매출_현지화'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '당기손익'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '당기손익'] + merge.loc[merge['손익센터'] == 'ANID2202', '당기손익'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '당기손익_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '당기손익_현지화'] + merge.loc[merge['손익센터'] == 'ANID2202', '당기손익_현지화'].values

merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '당기원가'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '당기원가'] + merge.loc[merge['손익센터'] == 'ANID2203', '당기원가'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '전기원가'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '전기원가'] + merge.loc[merge['손익센터'] == 'ANID2203', '전기원가'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계원가'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계원가'] + merge.loc[merge['손익센터'] == 'ANID2203', '누계원가'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계진행율대상원가'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계진행율대상원가'] + merge.loc[merge['손익센터'] == 'ANID2203', '누계진행율대상원가'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계손익'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계손익'] + merge.loc[merge['손익센터'] == 'ANID2203', '누계손익'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '당기매출'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '당기매출'] + merge.loc[merge['손익센터'] == 'ANID2203', '당기매출'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '전기매출'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '전기매출'] + merge.loc[merge['손익센터'] == 'ANID2203', '전기매출'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계기성청구합계(KRW)'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계기성청구합계(KRW)'] + merge.loc[merge['손익센터'] == 'ANID2203', '누계기성청구합계(KRW)'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계매출'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계매출'] + merge.loc[merge['손익센터'] == 'ANID2203', '누계매출'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '당기원가_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '당기원가_현지화'] + merge.loc[merge['손익센터'] == 'ANID2203', '당기원가_현지화'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '전기원가_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '전기원가_현지화'] + merge.loc[merge['손익센터'] == 'ANID2203', '전기원가_현지화'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계원가_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계원가_현지화'] + merge.loc[merge['손익센터'] == 'ANID2203', '누계원가_현지화'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계진행율대상원가_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계진행율대상원가_현지화'] + merge.loc[merge['손익센터'] == 'ANID2203', '누계진행율대상원가_현지화'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계손익_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계손익_현지화'] + merge.loc[merge['손익센터'] == 'ANID2203', '누계손익_현지화'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '당기매출_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '당기매출_현지화'] + merge.loc[merge['손익센터'] == 'ANID2203', '당기매출_현지화'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '전기매출_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '전기매출_현지화'] + merge.loc[merge['손익센터'] == 'ANID2203', '전기매출_현지화'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계기성청구합계'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계기성청구합계'] + merge.loc[merge['손익센터'] == 'ANID2203', '누계기성청구합계'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계매출_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '누계매출_현지화'] + merge.loc[merge['손익센터'] == 'ANID2203', '누계매출_현지화'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '당기손익'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '당기손익'] + merge.loc[merge['손익센터'] == 'ANID2203', '당기손익'].values
merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '당기손익_현지화'] = merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2203']), '당기손익_현지화'] + merge.loc[merge['손익센터'] == 'ANID2203', '당기손익_현지화'].values

merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '당기원가'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '당기원가'] + merge.loc[merge['손익센터'] == 'AAID1903', '당기원가'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '전기원가'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '전기원가'] + merge.loc[merge['손익센터'] == 'AAID1903', '전기원가'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계원가'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계원가'] + merge.loc[merge['손익센터'] == 'AAID1903', '누계원가'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계진행율대상원가'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계진행율대상원가'] + merge.loc[merge['손익센터'] == 'AAID1903', '누계진행율대상원가'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계손익'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계손익'] + merge.loc[merge['손익센터'] == 'AAID1903', '누계손익'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '당기매출'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '당기매출'] + merge.loc[merge['손익센터'] == 'AAID1903', '당기매출'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '전기매출'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '전기매출'] + merge.loc[merge['손익센터'] == 'AAID1903', '전기매출'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계기성청구합계(KRW)'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계기성청구합계(KRW)'] + merge.loc[merge['손익센터'] == 'AAID1903', '누계기성청구합계(KRW)'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계매출'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계매출'] + merge.loc[merge['손익센터'] == 'AAID1903', '누계매출'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '당기원가_현지화'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '당기원가_현지화'] + merge.loc[merge['손익센터'] == 'AAID1903', '당기원가_현지화'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '전기원가_현지화'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '전기원가_현지화'] + merge.loc[merge['손익센터'] == 'AAID1903', '전기원가_현지화'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계원가_현지화'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계원가_현지화'] + merge.loc[merge['손익센터'] == 'AAID1903', '누계원가_현지화'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계진행율대상원가_현지화'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계진행율대상원가_현지화'] + merge.loc[merge['손익센터'] == 'AAID1903', '누계진행율대상원가_현지화'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계손익_현지화'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계손익_현지화'] + merge.loc[merge['손익센터'] == 'AAID1903', '누계손익_현지화'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '당기매출_현지화'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '당기매출_현지화'] + merge.loc[merge['손익센터'] == 'AAID1903', '당기매출_현지화'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '전기매출_현지화'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '전기매출_현지화'] + merge.loc[merge['손익센터'] == 'AAID1903', '전기매출_현지화'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계기성청구합계'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계기성청구합계'] + merge.loc[merge['손익센터'] == 'AAID1903', '누계기성청구합계'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계매출_현지화'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '누계매출_현지화'] + merge.loc[merge['손익센터'] == 'AAID1903', '누계매출_현지화'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '당기손익'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '당기손익'] + merge.loc[merge['손익센터'] == 'AAID1903', '당기손익'].values
merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '당기손익_현지화'] = merge.loc[merge['손익센터'].isin(['AAID1902', 'AAID1903']), '당기손익_현지화'] + merge.loc[merge['손익센터'] == 'AAID1903', '당기손익_현지화'].values


merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '당기원가'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '당기원가'] + merge.loc[merge['손익센터'] == 'CFKR1811', '당기원가'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '전기원가'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '전기원가'] + merge.loc[merge['손익센터'] == 'CFKR1811', '전기원가'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계원가'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계원가'] + merge.loc[merge['손익센터'] == 'CFKR1811', '누계원가'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계진행율대상원가'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계진행율대상원가'] + merge.loc[merge['손익센터'] == 'CFKR1811', '누계진행율대상원가'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계손익'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계손익'] + merge.loc[merge['손익센터'] == 'CFKR1811', '누계손익'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '당기매출'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '당기매출'] + merge.loc[merge['손익센터'] == 'CFKR1811', '당기매출'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '전기매출'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '전기매출'] + merge.loc[merge['손익센터'] == 'CFKR1811', '전기매출'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계기성청구합계(KRW)'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계기성청구합계(KRW)'] + merge.loc[merge['손익센터'] == 'CFKR1811', '누계기성청구합계(KRW)'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계매출'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계매출'] + merge.loc[merge['손익센터'] == 'CFKR1811', '누계매출'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '당기원가_현지화'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '당기원가_현지화'] + merge.loc[merge['손익센터'] == 'CFKR1811', '당기원가_현지화'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '전기원가_현지화'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '전기원가_현지화'] + merge.loc[merge['손익센터'] == 'CFKR1811', '전기원가_현지화'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계원가_현지화'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계원가_현지화'] + merge.loc[merge['손익센터'] == 'CFKR1811', '누계원가_현지화'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계진행율대상원가_현지화'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계진행율대상원가_현지화'] + merge.loc[merge['손익센터'] == 'CFKR1811', '누계진행율대상원가_현지화'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계손익_현지화'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계손익_현지화'] + merge.loc[merge['손익센터'] == 'CFKR1811', '누계손익_현지화'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '당기매출_현지화'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '당기매출_현지화'] + merge.loc[merge['손익센터'] == 'CFKR1811', '당기매출_현지화'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '전기매출_현지화'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '전기매출_현지화'] + merge.loc[merge['손익센터'] == 'CFKR1811', '전기매출_현지화'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계기성청구합계'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계기성청구합계'] + merge.loc[merge['손익센터'] == 'CFKR1811', '누계기성청구합계'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계매출_현지화'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '누계매출_현지화'] + merge.loc[merge['손익센터'] == 'CFKR1811', '누계매출_현지화'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '당기손익'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '당기손익'] + merge.loc[merge['손익센터'] == 'CFKR1811', '당기손익'].values
merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '당기손익_현지화'] = merge.loc[merge['손익센터'].isin(['CFKR1810', 'CFKR1811']), '당기손익_현지화'] + merge.loc[merge['손익센터'] == 'CFKR1811', '당기손익_현지화'].values


# merge.loc[merge['손익센터'].isin(['ANID2201', 'ANID2202']), '손익센터'] = 'ANID2202'

# df.loc[df['b'].isin(['y', 'z']), 'd'] = df.loc[df['b'].isin(['y', 'z']), 'd'] + df.loc[df['b'] == 'z', 'd'].values

# 'z'열을 없앱니다.
# df = df[df['b'] != 'z']

# 인덱스를 리셋합니다.
# df.reset_index(drop=True, inplace=True)

# 결과 출력
# print(df)


# 통합프손에서 계약고, 총공사 예정비, 예정원가율, 계상선수, 계상미수 가져오기 (별도 통합)



df1 = merge[merge['손익센터']=='AAID1902'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='AAID1902'].index
merge.loc[df1,'계약고'] = 통합실행.loc[통합실행1,'계약고'].values

df1 = merge[merge['손익센터']=='AAID1902'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='AAID1902'].index
merge.loc[df1,'계약고(KRW)'] = 통합실행.loc[통합실행1,'계약고(KRW)'].values

df1 = merge[merge['손익센터']=='AAID1902'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='AAID1902'].index
merge.loc[df1,'총공사예정비'] = 통합실행.loc[통합실행1,'총공사예정비'].values

df1 = merge[merge['손익센터']=='AAID1902'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='AAID1902'].index
merge.loc[df1,'총공사예정비(KRW)'] = 통합실행.loc[통합실행1,'총공사예정비(KRW)'].values

df1 = merge[merge['손익센터']=='AAID1902'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='AAID1902'].index
merge.loc[df1,'예정원가율'] = 통합실행.loc[통합실행1,'예정원가율'].values


df1 = merge[merge['손익센터']=='AAID1902'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='AAID1902'].index
merge.loc[df1,'예정원가율(KRW)'] = 통합실행.loc[통합실행1,'예정원가율(KRW)'].values

df1 = merge[merge['손익센터']=='AAID1902'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='AAID1902'].index
merge.loc[df1,'계상미수'] = 통합실행.loc[통합실행1,'계상미수'].values

df1 = merge[merge['손익센터']=='AAID1902'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='AAID1902'].index
merge.loc[df1,'계상선수'] = 통합실행.loc[통합실행1,'계상선수'].values

df1 = merge[merge['손익센터']=='AAID1902'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='AAID1902'].index
merge.loc[df1,'계상미수(KRW)'] = 통합실행.loc[통합실행1,'계상미수(KRW)'].values

df1 = merge[merge['손익센터']=='AAID1902'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='AAID1902'].index
merge.loc[df1,'계상선수(KRW)'] = 통합실행.loc[통합실행1,'계상선수(KRW)'].values

merge.loc[merge['손익센터'] == 'AAID1902', '누계기성청구합계'] = merge.loc[merge['손익센터'] == 'AAID1902', '누계매출_현지화'] + merge.loc[merge['손익센터'] == 'AAID1902', '계상선수'] - merge.loc[merge['손익센터'] == 'AAID1902', '계상미수']

merge.loc[merge['손익센터'] == 'AAID1902', '누계기성청구합계(KRW)'] = merge.loc[merge['손익센터'] == 'AAID1902', '누계매출'] + merge.loc[merge['손익센터'] == 'AAID1902', '계상선수(KRW)'] - merge.loc[merge['손익센터'] == 'AAID1902', '계상미수(KRW)']


df1 = merge[merge['손익센터']=='CFKR1810'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='CFKR1810'].index
merge.loc[df1,'계약고'] = 통합실행.loc[통합실행1,'계약고'].values

df1 = merge[merge['손익센터']=='CFKR1810'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='CFKR1810'].index
merge.loc[df1,'계약고(KRW)'] = 통합실행.loc[통합실행1,'계약고(KRW)'].values

df1 = merge[merge['손익센터']=='CFKR1810'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='CFKR1810'].index
merge.loc[df1,'총공사예정비'] = 통합실행.loc[통합실행1,'총공사예정비'].values

df1 = merge[merge['손익센터']=='CFKR1810'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='CFKR1810'].index
merge.loc[df1,'총공사예정비(KRW)'] = 통합실행.loc[통합실행1,'총공사예정비(KRW)'].values

df1 = merge[merge['손익센터']=='CFKR1810'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='CFKR1810'].index
merge.loc[df1,'예정원가율'] = 통합실행.loc[통합실행1,'예정원가율'].values


df1 = merge[merge['손익센터']=='CFKR1810'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='CFKR1810'].index
merge.loc[df1,'예정원가율(KRW)'] = 통합실행.loc[통합실행1,'예정원가율(KRW)'].values

df1 = merge[merge['손익센터']=='CFKR1810'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='CFKR1810'].index
merge.loc[df1,'계상미수'] = 통합실행.loc[통합실행1,'계상미수'].values

df1 = merge[merge['손익센터']=='CFKR1810'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='CFKR1810'].index
merge.loc[df1,'계상선수'] = 통합실행.loc[통합실행1,'계상선수'].values

df1 = merge[merge['손익센터']=='CFKR1810'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='CFKR1810'].index
merge.loc[df1,'계상미수(KRW)'] = 통합실행.loc[통합실행1,'계상미수(KRW)'].values

df1 = merge[merge['손익센터']=='CFKR1810'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='CFKR1810'].index
merge.loc[df1,'계상선수(KRW)'] = 통합실행.loc[통합실행1,'계상선수(KRW)'].values

merge.loc[merge['손익센터'] == 'CFKR1810', '누계기성청구합계'] = merge.loc[merge['손익센터'] == 'CFKR1810', '누계매출_현지화'] + merge.loc[merge['손익센터'] == 'CFKR1810', '계상선수'] - merge.loc[merge['손익센터'] == 'CFKR1810', '계상미수']

merge.loc[merge['손익센터'] == 'CFKR1810', '누계기성청구합계(KRW)'] = merge.loc[merge['손익센터'] == 'CFKR1810', '누계매출'] + merge.loc[merge['손익센터'] == 'CFKR1810', '계상선수(KRW)'] - merge.loc[merge['손익센터'] == 'CFKR1810', '계상미수(KRW)']

df1 = merge[merge['손익센터']=='ANID2202'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='ANID2202'].index
merge.loc[df1,'계약고'] = 통합실행.loc[통합실행1,'계약고'].values

df1 = merge[merge['손익센터']=='ANID2202'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='ANID2202'].index
merge.loc[df1,'계약고(KRW)'] = 통합실행.loc[통합실행1,'계약고(KRW)'].values

df1 = merge[merge['손익센터']=='ANID2202'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='ANID2202'].index
merge.loc[df1,'총공사예정비'] = 통합실행.loc[통합실행1,'총공사예정비'].values

df1 = merge[merge['손익센터']=='ANID2202'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='ANID2202'].index
merge.loc[df1,'총공사예정비(KRW)'] = 통합실행.loc[통합실행1,'총공사예정비(KRW)'].values

df1 = merge[merge['손익센터']=='ANID2202'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='ANID2202'].index
merge.loc[df1,'예정원가율'] = 통합실행.loc[통합실행1,'예정원가율'].values


df1 = merge[merge['손익센터']=='ANID2202'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='ANID2202'].index
merge.loc[df1,'예정원가율(KRW)'] = 통합실행.loc[통합실행1,'예정원가율(KRW)'].values

df1 = merge[merge['손익센터']=='ANID2202'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='ANID2202'].index
merge.loc[df1,'계상미수'] = 통합실행.loc[통합실행1,'계상미수'].values

df1 = merge[merge['손익센터']=='ANID2202'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='ANID2202'].index
merge.loc[df1,'계상선수'] = 통합실행.loc[통합실행1,'계상선수'].values

df1 = merge[merge['손익센터']=='ANID2202'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='ANID2202'].index
merge.loc[df1,'계상미수(KRW)'] = 통합실행.loc[통합실행1,'계상미수(KRW)'].values

df1 = merge[merge['손익센터']=='ANID2202'].index
통합실행1 = 통합실행[통합실행['손익센터'] =='ANID2202'].index
merge.loc[df1,'계상선수(KRW)'] = 통합실행.loc[통합실행1,'계상선수(KRW)'].values

merge.loc[merge['손익센터'] == 'ANID2202', '누계기성청구합계'] = merge.loc[merge['손익센터'] == 'ANID2202', '누계매출_현지화'] + merge.loc[merge['손익센터'] == 'ANID2202', '계상선수'] - merge.loc[merge['손익센터'] == 'ANID2202', '계상미수']

merge.loc[merge['손익센터'] == 'ANID2202', '누계기성청구합계(KRW)'] = merge.loc[merge['손익센터'] == 'ANID2202', '누계매출'] + merge.loc[merge['손익센터'] == 'ANID2202', '계상선수(KRW)'] - merge.loc[merge['손익센터'] == 'ANID2202', '계상미수(KRW)']

#여기까지 별도통실

# df1 = df.reset_index(drop=True)
merge.to_excel('c:/연결프로젝트손익명세/별도통실.xlsx',sheet_name='pot_table')





#아래부터 연결통합

for index, row in consolidation.iterrows():
    
    var1 = row['주요pjt']
    var2 = row['종속pjt']
    var3 = row['지분율']
    var4 = row['환율']



    share = var3
    minor_share = 1-share
    minor_share = round(minor_share,2)
    print(share)
    print(minor_share)

    merge.loc[merge['손익센터'] == var1, '계약고'] += merge.loc[merge['손익센터'] == var2, '계약고'].values[0] - round( merge.loc[merge['손익센터'] == var2, '계약고'].values[0]*share ,2)                                    #계약고는 종속 프로젝트의 비지배지분만큼만 더해진다, 지배지분은 흡수
    merge.loc[merge['손익센터'] == var1, '계약고(KRW)'] += merge.loc[merge['손익센터'] == var2, '계약고(KRW)'].values[0] -  merge.loc[merge['손익센터'] == var2, '계약고(KRW)'].values[0]*share                                     #계약고는 종속 프로젝트의 비지배지분만큼만 더해진다, 지배지분은 흡수

    
    merge.loc[merge['손익센터'] == var1, '총공사예정비'] += merge.loc[merge['손익센터'] == var2, '총공사예정비'].values[0] -  merge.loc[merge['손익센터'] == var2, '계약고'].values[0]*share                                          #총공사예정비를 더하고 종속프로젝트의 계약고 지분율만큼 흡수
    merge.loc[merge['손익센터'] == var1, '총공사예정비(KRW)'] += merge.loc[merge['손익센터'] == var2, '총공사예정비(KRW)'].values[0] -  merge.loc[merge['손익센터'] == var2, '계약고(KRW)'].values[0]*share                                          #총공사예정비를 더하고 종속프로젝트의 계약고 지분율만큼 흡수

    
    merge.loc[merge['손익센터'] == var1, '예정원가율'] = merge.loc[merge['손익센터'] == var1, '총공사예정비'].values[0] /merge.loc[merge['손익센터'] == var1, '계약고'].values[0]    #예정원가율은 통합 계약고와 예정원가의 비율
    

    merge.loc[merge['손익센터'] == var1, '누계기성청구합계'] += merge.loc[merge['손익센터'] == var2, '누계기성청구합계'].values[0] -  merge.loc[merge['손익센터'] == var2, '누계기성청구합계'].values[0]*share #기성청구누계는 종속 프로젝트의 기성청구 비지배지분만큼 더해진다
    merge.loc[merge['손익센터'] == var1, '누계기성청구합계(KRW)'] += merge.loc[merge['손익센터'] == var2, '누계기성청구합계(KRW)'].values[0] -  merge.loc[merge['손익센터'] == var2, '누계기성청구합계(KRW)'].values[0]*share #기성청구누계는 종속 프로젝트의 기성청구 비지배지분만큼 더해진다
    
    
    
    
    merge.loc[merge['손익센터'] == var1, '누계진행율대상원가_현지화'] += merge.loc[merge['손익센터'] == var2, '누계진행율대상원가_현지화'].values[0] - merge.loc[merge['손익센터'] == var2, '누계기성청구합계'].values[0]*share
    merge.loc[merge['손익센터'] == var1, '누계진행율대상원가'] += merge.loc[merge['손익센터'] == var2, '누계진행율대상원가'].values[0] - merge.loc[merge['손익센터'] == var2, '누계기성청구합계(KRW)'].values[0]*share
    
    merge.loc[merge['손익센터'] == var1, '누계매출_현지화'] = (merge.loc[merge['손익센터'] == var1, '누계진행율대상원가_현지화'] / merge.loc[merge['손익센터'] == var1, '예정원가율']).round(0)
    
    merge.loc[merge['손익센터'] == var1, '계상미수'] = 0
    merge.loc[merge['손익센터'] == var1, '계상선수'] = 0
    merge.loc[merge['손익센터'] == var1, '계상미수(KRW)'] = 0
    merge.loc[merge['손익센터'] == var1, '계상선수(KRW)'] = 0
    merge.loc[merge['손익센터'] == var1, '계상미수'] = 0
    merge.loc[merge['손익센터'] == var1, '계상미수(KRW)'] = 0
    
    계상매출var1 = merge.loc[merge['손익센터'] == var1, '누계매출_현지화'].values[0] - merge.loc[merge['손익센터'] == var1, '누계기성청구합계'].values[0]
    
    print(계상매출var1)
    
    
    if 계상매출var1 >= 0    :
        merge.loc[merge['손익센터'] == var1, '계상미수'] = 계상매출var1
        merge.loc[merge['손익센터'] == var1, '계상미수(KRW)'] = int(계상매출var1*var4)
    else:
        merge.loc[merge['손익센터'] == var1, '계상선수'] = 계상매출var1*-1
        merge.loc[merge['손익센터'] == var1, '계상선수(KRW)'] = int(계상매출var1*var4*-1)
        
    merge.loc[merge['손익센터'] == var1, '누계매출'] = merge.loc[merge['손익센터'] == var1, '누계기성청구합계(KRW)'].values[0] +  int(계상매출var1*var4)  
    
    
    
    
    
    merge.loc[merge['손익센터'] == var1, '누계하자보수충당부채전입액_현지화'] += merge.loc[merge['손익센터'] == var1, '누계하자보수충당부채전입액_현지화'].values[0]
    merge.loc[merge['손익센터'] == var1, '누계하자보수충당부채전입액'] += merge.loc[merge['손익센터'] == var1, '누계하자보수충당부채전입액'].values[0]
    
    
    merge.loc[merge['손익센터'] == var1, '작업진행율'] = round(merge.loc[merge['손익센터'] == var1, '누계매출_현지화'].values[0]) / merge.loc[merge['손익센터'] == var1, '계약고'].values[0]
    
    
    
    
    
    merge.loc[merge['손익센터'] == var1, '누계손실충당금전입액_현지화'] = 0 #공손충 리셋
    merge.loc[merge['손익센터'] == var1, '누계손실충당금전입액'] = 0 #공손충 리셋

    if merge.loc[merge['손익센터'] == var1, '예정원가율'].values[0] >= 1    :
        mask = merge['손익센터'] == var1
        merge.loc[mask, '누계손실충당금전입액_현지화'] = merge.loc[mask, '계약고'].values[0] * (1 - merge.loc[mask, '작업진행율'].values[0]) * (merge.loc[mask, '예정원가율'].values[0] - 1 + 0.0004)


    merge.loc[merge['손익센터'] == var1, '누계원가_현지화'] =merge.loc[merge['손익센터'] == var1, '누계진행율대상원가_현지화']        +     merge.loc[merge['손익센터'] == var1, '누계하자보수충당부채전입액_현지화'] + merge.loc[merge['손익센터'] == var1, '누계손실충당금전입액_현지화']
    merge.loc[merge['손익센터'] == var1, '누계원가'] =merge.loc[merge['손익센터'] == var1, '누계진행율대상원가']        +     merge.loc[merge['손익센터'] == var1, '누계하자보수충당부채전입액']

#아래부터 기존 통실

for index, row in consolidation_1.iterrows():
    
    var1 = row['주요pjt']
    var2 = row['종속pjt']
    var4 = row['환율']




    merge.loc[merge['손익센터'] == var1, '계약고'] += merge.loc[merge['손익센터'] == var2, '계약고'].values[0] 
    merge.loc[merge['손익센터'] == var1, '계약고(KRW)'] += merge.loc[merge['손익센터'] == var2, '계약고(KRW)'].values[0]

    
    merge.loc[merge['손익센터'] == var1, '총공사예정비'] += merge.loc[merge['손익센터'] == var2, '총공사예정비'].values[0]                                      #총공사예정비를 더하고 종속프로젝트의 계약고 지분율만큼 흡수
    merge.loc[merge['손익센터'] == var1, '총공사예정비(KRW)'] += merge.loc[merge['손익센터'] == var2, '총공사예정비(KRW)'].values[0]                       #총공사예정비를 더하고 종속프로젝트의 계약고 지분율만큼 흡수

    
    merge.loc[merge['손익센터'] == var1, '예정원가율'] = merge.loc[merge['손익센터'] == var1, '총공사예정비'].values[0] /merge.loc[merge['손익센터'] == var1, '계약고'].values[0]    #예정원가율은 통합 계약고와 예정원가의 비율
    

    merge.loc[merge['손익센터'] == var1, '누계기성청구합계'] += merge.loc[merge['손익센터'] == var2, '누계기성청구합계'].values[0] #기성청구누계는 종속 프로젝트의 기성청구 비지배지분만큼 더해진다
    merge.loc[merge['손익센터'] == var1, '누계기성청구합계(KRW)'] += merge.loc[merge['손익센터'] == var2, '누계기성청구합계(KRW)'].values[0] #기성청구누계는 종속 프로젝트의 기성청구 비지배지분만큼 더해진다
    
    
    
    
    merge.loc[merge['손익센터'] == var1, '누계진행율대상원가_현지화'] += merge.loc[merge['손익센터'] == var2, '누계진행율대상원가_현지화'].values[0]
    merge.loc[merge['손익센터'] == var1, '누계진행율대상원가'] += merge.loc[merge['손익센터'] == var2, '누계진행율대상원가'].values[0] 
    
    merge.loc[merge['손익센터'] == var1, '누계매출_현지화'] = (merge.loc[merge['손익센터'] == var1, '누계진행율대상원가_현지화'] / merge.loc[merge['손익센터'] == var1, '예정원가율']).round(2)
    
    merge.loc[merge['손익센터'] == var1, '계상미수'] = 0
    merge.loc[merge['손익센터'] == var1, '계상선수'] = 0
    merge.loc[merge['손익센터'] == var1, '계상미수(KRW)'] = 0
    merge.loc[merge['손익센터'] == var1, '계상선수(KRW)'] = 0


    
    계상매출var1 = merge.loc[merge['손익센터'] == var1, '누계매출_현지화'].values[0] - merge.loc[merge['손익센터'] == var1, '누계기성청구합계'].values[0]
    
    print(계상매출var1)
    
    
    if 계상매출var1 >= 0    :
        merge.loc[merge['손익센터'] == var1, '계상미수'] = 계상매출var1
        merge.loc[merge['손익센터'] == var1, '계상미수(KRW)'] = int(계상매출var1*var4)
    else:
        merge.loc[merge['손익센터'] == var1, '계상선수'] = 계상매출var1*-1
        merge.loc[merge['손익센터'] == var1, '계상선수(KRW)'] = int(계상매출var1*var4*-1)
        
    merge.loc[merge['손익센터'] == var1, '누계매출'] = merge.loc[merge['손익센터'] == var1, '누계기성청구합계(KRW)'].values[0] +  int(계상매출var1*var4)  
    
    
    
    
    
    merge.loc[merge['손익센터'] == var1, '누계하자보수충당부채전입액_현지화'] += merge.loc[merge['손익센터'] == var1, '누계하자보수충당부채전입액_현지화'].values[0]
    merge.loc[merge['손익센터'] == var1, '누계하자보수충당부채전입액'] += merge.loc[merge['손익센터'] == var1, '누계하자보수충당부채전입액'].values[0]
    
    
    merge.loc[merge['손익센터'] == var1, '작업진행율'] = round(merge.loc[merge['손익센터'] == var1, '누계매출_현지화'].values[0]) / merge.loc[merge['손익센터'] == var1, '계약고'].values[0]
    
    
    
    
    
    merge.loc[merge['손익센터'] == var1, '누계손실충당금전입액_현지화'] = 0 #공손충 리셋
    merge.loc[merge['손익센터'] == var1, '누계손실충당금전입액'] = 0 #공손충 리셋

    if merge.loc[merge['손익센터'] == var1, '예정원가율'].values[0] >= 1    :
        mask = merge['손익센터'] == var1
        merge.loc[mask, '누계손실충당금전입액_현지화'] = merge.loc[mask, '계약고'].values[0] * (1 - merge.loc[mask, '작업진행율'].values[0]) * (merge.loc[mask, '예정원가율'].values[0] - 1 + 0.0004)
        merge.loc[mask, '누계손실충당금전입액'] = merge.loc[mask, '누계손실충당금전입액_현지화'].values[0] * var4
    
    merge.loc[merge['손익센터'] == var1, '누계원가_현지화'] =merge.loc[merge['손익센터'] == var1, '누계진행율대상원가_현지화']        +     merge.loc[merge['손익센터'] == var1, '누계하자보수충당부채전입액_현지화'] + merge.loc[merge['손익센터'] == var1, '누계손실충당금전입액_현지화']
    merge.loc[merge['손익센터'] == var1, '누계원가'] =merge.loc[merge['손익센터'] == var1, '누계진행율대상원가']        +     merge.loc[merge['손익센터'] == var1, '누계하자보수충당부채전입액']

merge.to_excel('c:/연결프로젝트손익명세/전기통합전.xlsx')

merge['전기매출_현지화'] = 0
merge['전기원가_현지화'] = 0
merge['전기매출'] = 0
merge['전기원가'] = 0


#검색자리

wb = xw.Book('C:/연결프로젝트손익명세/전기.xlsx')
sht=wb.sheets[0]
previous = sht.range('a1').options(pd.DataFrame,expand='table',Header=1).value
print("전기 파일 로딩을 시작합니다")
wb.close()
prefix = 'previous_'
previous.columns = [prefix + str(col) for col in previous.columns]
previous

merge = merge.merge(previous, left_on='손익센터',right_on='previous_손익센터',  how='left')


#매출세팅
merge['전기매출_현지화'] = merge['previous_전기매출_현지화']
merge['당기매출_현지화'] = merge['누계매출_현지화'] - merge['전기매출_현지화']

merge['전기매출'] = merge['previous_전기매출']
merge['당기매출'] = merge['누계매출'] - merge['전기매출']








merge.to_excel('c:/연결검증.xlsx')















































#sht = wb.sheets['원화당기']
#sht.range('a1').options(index=False,Head=True).value = 원화당기

#sht = wb.sheets['원화전기']
#sht.range('a1').options(index=False,Head=True).value = 원화전기

#sht = wb.sheets['현지화전기']
#sht.range('a1').options(index=False,Head=True).value = 현지화전기

#sht = wb.sheets['현지화당기']
#sht.range('a1').options(index=False,Head=True).value = 현지화당기

#sht = wb.sheets['손익명세']
#sht.range('a1').options(index=False,Head=True).value = 손익명세
#print("데이터 처리를 완료했습니다")







#wb.sheets.add('new')






#분석치
print("분석을 시작합니다")


# sorted_df = merge.sort_values('당기매출', ascending=False)

# top_5_projects_sales = sorted_df.head(5)[['손익센터명','당기매출']]

# #sorted_df = merge.sort_values('당기매출', ascending=True)
# sorted_df = merge[merge['당기매출'] != 0].sort_values('당기매출', ascending=True)

# lowest_5_projects_sales = sorted_df.head(5)[['손익센터명','당기매출']]


# sorted_df = merge.sort_values('당기손익', ascending=False)

# top_5_projects_profit = sorted_df.head(5)[['손익센터명','당기손익']]

# sorted_df = merge.sort_values('당기손익', ascending=True)

# lowest_5_projects_profit = sorted_df.head(5)[['손익센터명','당기손익']]

# merge['원가율'] = np.where((merge['당기매출'] == 0) | (merge['당기원가'] == 0), 0, (merge['당기원가'] / merge['당기매출']) * 100)
# #merge['원가율'] = np.where((merge['원가율'] != '연산불가') & (merge['원가율'] != 0), merge['원가율'].round(2), '연산불가')

# merge['원가율'] = np.where(merge['원가율'] != 0, np.round(merge['원가율'] / 100, 4), merge['원가율'])


# merge['원가율_현지화'] = np.where((merge['당기매출_현지화'] == 0) | (merge['당기원가_현지화'] == 0), 0, (merge['당기원가_현지화'] / 
#                                                                                      merge['당기매출_현지화']) * 100)
# #merge['원가율'] = np.where((merge['원가율'] != '연산불가') & (merge['원가율'] != 0), merge['원가율'].round(2), '연산불가')

# merge['원가율_현지화'] = np.where(merge['원가율_현지화'] != 0, np.round(merge['원가율_현지화'] / 100, 4), merge['원가율_현지화'])


# novalue = merge[merge['원가율']==0][['손익센터명','당기매출','당기원가','당기손익']]


# sorted_df = merge.sort_values('원가율', ascending=False)

# top_5_projects_costrate = sorted_df.head(20)[['손익센터명','원가율','당기손익']]


# sorted_df = merge.sort_values('원가율', ascending=True)

# lowest_5_projects_costrate = sorted_df.head(20)[['손익센터명','원가율','당기손익']]



# merge['재계산매출']=merge['진행율대상원가'] / merge['원가율_현지화']

# merge['환율효과'] =merge['당기매출'] - merge['재계산매출'] 




# # sht = wb.sheets['new']
# # sht.range('a1').options(index=False,Head=True).value = merge

# sht = wb.sheets['분석']

# sht.range('b2').value = "매출 상위 5개 프로젝트는 아래와 같습니다."
# sht.range('b3').value = "프로젝트"
# sht.range('c3').value = "매출"
# sht.range('b4').options(index=False, header=False).value = top_5_projects_sales



# sht.range('b10').value = "매출 하위 5개 프로젝트는 아래와 같습니다."
# sht.range('b12').options(index=False, header=False).value = lowest_5_projects_sales
# sht.range('b11').value = "프로젝트"
# sht.range('c11').value = "매출"

# sht.range('b18').value = "손익 상위 5개 프로젝트는 아래와 같습니다."
# sht.range('b20').options(index=False, header=False).value = top_5_projects_profit
# sht.range('b19').value = "프로젝트"
# sht.range('c19').value = "손익"

# sht.range('b26').value = "손익 하위 5개 프로젝트는 아래와 같습니다."
# sht.range('b28').options(index=False, header=False).value = lowest_5_projects_profit
# sht.range('b27').value = "프로젝트"
# sht.range('c27').value = "손익"

# sht.range('b34').value = "원가율 상위 10개 프로젝트는 아래와 같습니다."
# sht.range('b36').options(index=False, header=False).value = top_5_projects_costrate
# sht.range('b35').value = "프로젝트"
# sht.range('c35').value = "원가율"
# sht.range('d35').value = "손익"

# sht.range('b57').value = "원가율 하위 10개 프로젝트는 아래와 같습니다."
# sht.range('b59').options(index=False, header=False).value = lowest_5_projects_costrate
# sht.range('b58').value = "프로젝트"
# sht.range('c58').value = "원가율"
# sht.range('d58').value = "손익"

# sht.range('b82').options(index=False, header=False).value = novalue
# sht.range('b80').value = "매출이나 원가만 발생한 프로젝트는 아래와 같습니다"
# sht.range('b81').value = "프로젝트"
# sht.range('c81').value = "당기매출"
# sht.range('d81').value = "당기원가"
# sht.range('e81').value = "당기손익"






print("분석을 완료했습니다")


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




merge.to_excel('c:/연결프로젝트손익명세/검토2.xlsx',sheet_name='pot_table')
      

#아래 일시 취소


# sht = wb.sheets['new']
# sht.range('a1').options(index=False,Head=True).value = merge

# # 기존의 칼럼이 있는 데이터프레임 생성


# # 칼럼 제거하여 값만 남은 데이터프레임 생성 df_values_only = df_with_columns.drop(columns=df_with_columns.columns)

# #df_values_only = pd.DataFrame(df_with_columns.values)
# merge1 = pd.DataFrame(merge.values)



# merge1 = merge1.values

# sht = wb.sheets['Final']



# range_to_clear = 'A8:DS1000'
# sht.range(range_to_clear).clear_contents()



# sht.range('a8').options(index=False,Head=False).value = merge1

# #merge.to_excel('c:/pyt.xlsx',sheet_name='pot_table')



# wb.save()
# wb.close()

# print("파일저장이 완료되었습니다")
