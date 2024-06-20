import streamlit as st
from jinja2 import Template
from trino.dbapi import connect
from trino.auth import BasicAuthentication
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.font_manager as fm
import seaborn as sns
from sklearn.ensemble import IsolationForest
import plotly.express as px
import plotly.graph_objects as go


# 페이지 설정
st.set_page_config(layout="wide")

# 한글 폰트 설정
font_path = 'C:/Users/hec/AppData/Local/Microsoft/Windows/Fonts/NanumBarunGothic.ttf'  # 사용할 폰트 경로
font_name = fm.FontProperties(fname=font_path).get_name()
plt.rc('font', family=font_name)

# Streamlit 앱 설정
st.title('Financial Analysis - Accounting Team')

# 기본 메뉴 추가
menu_options = ['기성미수금', '미청구', '선급금', '매입채무', '선수금', '초과청구']

# 선택한 페이지 상태 초기화
if 'page' not in st.session_state:
    st.session_state.page = None

# 사이드바에 버튼 추가
if st.sidebar.button('Account Statement'):
    st.session_state.page = 'account statement_menu'
if st.sidebar.button('Monitoring'):
    st.session_state.page = 'monitoring'
if st.sidebar.button('AI 분석'):
    st.session_state.page = 'ai_analysis'

# Trino에 연결 설정
connection = connect(
    host='trino.hec.co.kr', port=443, http_scheme='https',
    user='0900051', auth=BasicAuthentication('0900051', 'Hec!@345'),
    catalog='dw', schema='fie')
cursor = connection.cursor()

# 선택한 페이지에 따라 다른 내용을 표시
if st.session_state.page is None:
    st.write("왼쪽 화면에서 메뉴를 선택하세요.")

elif st.session_state.page == 'account statement_menu':
    st.subheader("Account Statement Menu")
    
    if st.button('기성미수금'):
        st.session_state.page = 'account statement'
        st.session_state.selected_option = '기성미수금'
    if st.button('미청구'):
        st.session_state.page = 'account statement'
        st.session_state.selected_option = '미청구'
    if st.button('선급금'):
        st.session_state.page = 'account statement'
        st.session_state.selected_option = '선급금'
    if st.button('매입채무'):
        st.session_state.page = 'account statement'
        st.session_state.selected_option = '매입채무'
    if st.button('선수금'):
        st.session_state.page = 'account statement'
        st.session_state.selected_option = '선수금'
    if st.button('초과청구'):
        st.session_state.page = 'account statement'
        st.session_state.selected_option = '초과청구'

elif st.session_state.page == 'account statement':
    selected_option = st.session_state.selected_option
    
    st.subheader(f"Account Statement - {selected_option}")

    # 공통 사용자 입력 받기
    end_date_prior = st.date_input("Enter the prior period end date", value=pd.to_datetime('2023-12-31'))
    end_date_current = st.date_input("Enter the current period end date", value=pd.to_datetime('2024-05-31'))

    # 날짜를 문자열로 변환
    end_date_prior_str = end_date_prior.strftime('%Y-%m-%d')
    end_date_current_str = end_date_current.strftime('%Y-%m-%d')


    cursor.execute("""
        SELECT DISTINCT COALESCE(p.sector, '비어있음') AS sector
        FROM s_accounting_ledger l
        LEFT JOIN s_accounting_project p ON l.profit_center = p.profit_center
    """)
    sectors = cursor.fetchall()
    sector_options = [sector[0] for sector in sectors]

    # 세션 상태에 selected_sectors가 없다면 모든 섹터를 선택된 상태로 초기화
    if 'selected_sectors' not in st.session_state:
        st.session_state.selected_sectors = sector_options

    # 사용자로부터 sector 선택받기
    selected_sectors = []
    for sector in sector_options:
        sector_str = str(sector)  # 문자열로 변환
        if st.checkbox(sector_str, key=sector_str, value=sector_str in st.session_state.selected_sectors):
            selected_sectors.append(sector_str)

    st.session_state.selected_sectors = selected_sectors

    # 선택된 sector를 쿼리 조건에 추가
    if selected_sectors:
        sector_filter = "AND COALESCE(p.sector, '비어있음') IN ({})".format(", ".join(["'{}'".format(sector) for sector in selected_sectors]))
    else:
        sector_filter = ""


    group_by_fields = "l.profit_center,l.co_code, l.acc_name, l.profit_center_name, p.sector, l.client_vendor_name" 
    select_fields = "p.sector,l.co_code, l.profit_center, l.acc_name,l.client_vendor_name, l.profit_center_name"




    exclude_acc_name = st.checkbox('계정과목 / 거래처 분류 제외')

    if exclude_acc_name:
        group_by_fields = "l.profit_center,l.co_code, l.profit_center_name, p.sector "
        select_fields = "p.sector,l.co_code, l.profit_center, l.profit_center_name"













    # 각 항목에 대한 쿼리 정의
    queries = {
        '기성미수금': f"""
            SELECT
              {select_fields},
              SUM(CASE 
                  WHEN l.posting_date <=  DATE '{end_date_prior_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                          ELSE l.gc
                      END 
                  ELSE 0 
              END) AS prior_period,
              SUM(CASE 
                  WHEN l.posting_date <=  DATE '{end_date_current_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                          ELSE l.gc
                      END 
                  ELSE 0 
              END) AS current_period,
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_current_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                          ELSE l.gc
                      END 
                  ELSE 0 
              END) - 
              SUM(CASE 
                  WHEN l.posting_date <=  DATE '{end_date_prior_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                          ELSE l.gc
                      END 
                  ELSE 0 
              END) AS difference
            FROM
              s_accounting_ledger l
            LEFT JOIN
              s_accounting_project p
            ON
              l.profit_center = p.profit_center
            WHERE
              CAST(l.acc_num AS VARCHAR) BETWEEN '11530110' AND '11530199'
              {sector_filter}
            GROUP BY
              {group_by_fields}
        """,
        '미청구': f"""
            SELECT
              {select_fields},
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_prior_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) AS prior_period,
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_current_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) AS current_period,
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_current_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) - 
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_prior_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) AS difference
            FROM
              s_accounting_ledger l
            LEFT JOIN
              s_accounting_project p
            ON
              l.profit_center = p.profit_center
            WHERE
              l.acc_num BETWEEN 11610110 AND 11610130
              {sector_filter}
            GROUP BY
              {group_by_fields}
        """,
        '선급금': f"""
            SELECT
              {select_fields},
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_prior_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) AS prior_period,
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_current_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) AS current_period,
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_current_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) - 
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_prior_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) AS difference
            FROM
              s_accounting_ledger l
            LEFT JOIN
              s_accounting_project p
            ON
              l.profit_center = p.profit_center
            WHERE
              l.acc_num BETWEEN 12010110 AND 12010199
              {sector_filter}
            GROUP BY
              {group_by_fields}
        """,
        '매입채무': f"""
            SELECT
              {select_fields},
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_prior_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) AS prior_period,
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_current_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) AS current_period,
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_current_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) - 
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_prior_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) AS difference
            FROM
              s_accounting_ledger l
            LEFT JOIN
              s_accounting_project p
            ON
              l.profit_center = p.profit_center
            WHERE
              l.acc_num BETWEEN 21030110 AND 21030199
              {sector_filter}
            GROUP BY
              {group_by_fields}
        """,
        '선수금': f"""
            SELECT
              {select_fields},
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_prior_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) AS prior_period,
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_current_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) AS current_period,
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_current_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) - 
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_prior_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) AS difference
            FROM
              s_accounting_ledger l
            LEFT JOIN
              s_accounting_project p
            ON
              l.profit_center = p.profit_center
            WHERE
              l.acc_num BETWEEN 21110110 AND 21110199
              {sector_filter}
            GROUP BY
              {group_by_fields}
        """,
        '초과청구': f"""
            SELECT
              {select_fields},
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_prior_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) AS prior_period,
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_current_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) AS current_period,
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_current_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) - 
              SUM(CASE 
                  WHEN l.posting_date <= DATE '{end_date_prior_str}' THEN 
                      CASE 
                          WHEN CAST(l.acc_num AS VARCHAR) LIKE '2%' OR CAST(l.acc_num AS VARCHAR) LIKE '3%' THEN -l.gc
                      ELSE l.gc 
                      END 
                  ELSE 0 
              END) AS difference
            FROM
              s_accounting_ledger l
            LEFT JOIN
              s_accounting_project p
            ON
              l.profit_center = p.profit_center
            WHERE
              l.acc_num BETWEEN 21210110 AND 21210120
              {sector_filter}
            GROUP BY
              {group_by_fields}
        """
    }

    # 선택된 항목에 따라 쿼리 실행
    ledger_query = queries[selected_option]

    # Pandas Styler 설정 변경
    pd.set_option("styler.render.max_elements", 305940)

    # 숫자 형식 지정 함수
    def format_numbers(val):
        color = 'red' if val < 0 else 'black'
        return f'color: {color};'






    # 쿼리 실행 및 결과 가져오기
     # 쿼리 실행 및 결과 가져오기
# 쿼리 실행 및 결과 가져오기
# 쿼리 실행 및 결과 가져오기
# 쿼리 실행 및 결과 가져오기
    # 쿼리 실행 및 결과 가져오기
    if st.button('Run Query'):
        cursor.execute(ledger_query)
        rows = cursor.fetchall()
        columns = [desc[0] for desc in cursor.description]

        # 결과를 데이터프레임으로 변환
        df = pd.DataFrame(rows, columns=columns)

        # 숫자 형식으로 변환할 열을 선택
        numeric_columns = ['prior_period', 'current_period', 'difference']
        df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric)

        # 모든 숫자 열이 0인 행을 제거
        df = df[~(df[numeric_columns] == 0).all(axis=1)]

        # current_period 기준으로 내림차순 정렬
        df = df.sort_values(by='current_period', ascending=False)

        # 조건부 서식 추가 (가로 막대 스타일)
        df_style = df.style.bar(subset=['current_period'], color='#d65f5f')\
                        .format("{:,.0f}", subset=numeric_columns)




        # 합계 계산
        summary = df[numeric_columns].sum().rename('Total')
        df_summary = pd.DataFrame(summary).transpose()
        

        # 1차 테이블 df_style
        st.dataframe(df_style, height=800, width=2000)  # 기본 크기를 더 넓힘

        # 2차 테이블 df_summary
        st.write(df_summary.style.format("{:,.0f}"))

        # 3차용 작업

        df = df.drop(columns=['sector', 'client_vendor_name', 'acc_name'], errors='ignore')
        df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric)
        df = df.groupby(['profit_center','profit_center_name'])[numeric_columns].sum().reset_index()
        df = df.sort_values(by='current_period', ascending=False)

        df['profit_center_full'] = df['profit_center_name'] + "-" + df['profit_center']


 
        # 전체 profit_center의 값을 가로 막대 그래프로 표시 df 
        fig = go.Figure()

        # 현재 기간과 이전 기간 값을 다른 색상으로 표시
        fig.add_trace(go.Bar(
            y=df['profit_center_full'],
            x=df['current_period'],
            name='당기',
            orientation='h',
            marker=dict(color='skyblue')))

        fig.add_trace(go.Bar(
            y=df['profit_center_full'],
            x=df['prior_period'],
            name='전기',
            orientation='h',
            marker=dict(color='lightcoral')     ))

        fig.update_layout(
            barmode='group',
            bargap=0.1,  # 막대 사이의 간격을 줄여 막대를 두껍게 만듦
            bargroupgap=0.3,  # 그룹 사이의 간격을 줄여 막대를 두껍게 만듦
            height=600 + len(df) * 20,  # 전체 그래프의 높이를 조정
            title='프로젝트별 금액(당기금액 순서)',
            xaxis_title='Amount',
            yaxis_title='Profit Center',
            yaxis=dict(autorange="reversed"),
            xaxis=dict(tickformat=',d')
        )

        st.plotly_chart(fig)

        # 엑셀 파일로 저장 버튼
        if st.button('Save to Excel'):
            with pd.ExcelWriter('c:/xx.xlsx') as writer:
                df.to_excel(writer, sheet_name='Data', index=False)
                df_summary.to_excel(writer, sheet_name='Summary', index=False)
            st.success('Data saved to c:/xx.xlsx')




elif st.session_state.page == 'monitoring':
    st.subheader("계정 모니터링 화면")

    # posting_date가 있는 모든 월을 가져오는 쿼리 실행
    cursor.execute("""
        SELECT DISTINCT DATE_FORMAT(posting_date, '%Y-%m') AS year_month
        FROM s_accounting_ledger
        ORDER BY year_month
    """)
    months = cursor.fetchall()
    month_options = [month[0] for month in months]

    # NoneType을 제외하고 날짜를 내림차순으로 정렬
    month_options_filtered = [month for month in month_options if month is not None]
    month_options_sorted = sorted(month_options_filtered, reverse=True)
    # 기본값을 가장 최근 월로 설정


    selected_month = st.selectbox("월 선택", month_options_sorted, index=0)

    if st.button('조회'):
        query = f"""
            WITH prior_data AS (
                SELECT
                    l.acc_name,
                    l.profit_center
                FROM
                    s_accounting_ledger l
                WHERE
                    DATE_FORMAT(l.posting_date, '%Y-%m') < '{selected_month}'
                GROUP BY
                    l.acc_name, l.profit_center
            ),
            current_data AS (
                SELECT
                    l.acc_name,
                    l.profit_center,
                    MIN(l.posting_date) as first_occurrence_date,
                    l.gc as amount,
                    l.description
                FROM
                    s_accounting_ledger l
                WHERE
                    DATE_FORMAT(l.posting_date, '%Y-%m') = '{selected_month}'
                GROUP BY
                    l.acc_name, l.profit_center, l.gc, l.description
            )
            SELECT
                c.acc_name,
                c.profit_center,
                c.first_occurrence_date,
                c.amount,
                c.description
            FROM
                current_data c
            LEFT JOIN
                prior_data p
            ON
                c.acc_name = p.acc_name
                AND c.profit_center = p.profit_center
            WHERE
                p.acc_name IS NULL
                AND p.profit_center IS NULL
        """

        cursor.execute(query)
        rows = cursor.fetchall()
        columns = [desc[0] for desc in cursor.description]

        # 결과를 데이터프레임으로 변환
        df = pd.DataFrame(rows, columns=columns)

        # 문구 추가
        st.markdown(f"### {selected_month} 최초 발생 프로젝트/계정")

        # 데이터프레임 출력
        st.dataframe(df, height=800, width=2000)

        # 추가 테이블 생성
        query_additional = f"""
            SELECT
                l.profit_center,
                l.profit_center_name,
                l.acc_name,
                l.description,
                SUM(CASE WHEN DATE_FORMAT(l.posting_date, '%Y-%m') = '{selected_month}' THEN l.gc ELSE 0 END) AS current_month,
                SUM(CASE WHEN DATE_FORMAT(l.posting_date, '%Y-%m') = DATE_FORMAT(DATE_TRUNC('month', DATE '{selected_month}-01') - INTERVAL '1' month, '%Y-%m') THEN l.gc ELSE 0 END) AS prior_month
            FROM
                s_accounting_ledger l
            WHERE
                l.description LIKE '%채권%' AND l.description LIKE '%채무%' AND l.description LIKE '%상계%'
            GROUP BY
                l.profit_center,
                l.profit_center_name,
                l.acc_name,
                l.description
        """

        cursor.execute(query_additional)
        rows_additional = cursor.fetchall()
        columns_additional = [desc[0] for desc in cursor.description]

        # 결과를 데이터프레임으로 변환
        df_additional = pd.DataFrame(rows_additional, columns=columns_additional)

        # 차이 계산
        df_additional['difference'] = df_additional['current_month'] - df_additional['prior_month']

        # prior_month와 current_month 값이 모두 0인 행 제거
        df_additional = df_additional[(df_additional['prior_month'] != 0) | (df_additional['current_month'] != 0)]

        # 데이터프레임 출력
        st.markdown(f"### {selected_month} 채권/채무/상계 프로젝트/계정")
        st.dataframe(df_additional, height=800, width=2000)

        # 엑셀 파일로 저장 버튼
        if st.button('Save Additional Data to Excel'):
            with pd.ExcelWriter('c:/monitoring_additional.xlsx') as writer:
                df_additional.to_excel(writer, sheet_name='Data', index=False)
            st.success('Data saved to c:/monitoring_additional.xlsx')
        # profit_center와 acc_name 기준으로 gc를 부분합
        query_sum_gc = f"""
            SELECT
                l.profit_center,
                l.profit_center_name,
                l.acc_name,
                SUM(l.gc) AS total_gc
            FROM
                s_accounting_ledger l
            WHERE
                l.posting_date BETWEEN DATE '{start_date_str}' AND DATE '{end_date_str}'
            GROUP BY
                l.profit_center,
                l.profit_center_name,
                l.acc_name
        """

        # 대상 기간의 월별 변동을 분석
        cursor.execute(query_sum_gc)
        result = cursor.fetchall()
        df = pd.DataFrame(result, columns=['profit_center', 'profit_center_name', 'acc_name', 'total_gc'])
        df['year_month'] = pd.to_datetime(df['posting_date']).dt.to_period('M')
        df_grouped = df.groupby(['profit_center', 'profit_center_name', 'acc_name', 'year_month']).sum().reset_index()

        # 이상치 그래프화
        fig = px.line(df_grouped, x='year_month', y='total_gc', color='profit_center', line_group='acc_name', title='이상치 분석')
        st.plotly_chart(fig)
elif st.session_state.page == 'ai_analysis':
    st.subheader("AI를 이용한 이상치 분석")

    # 사용자로부터 분석할 기간 입력받기
    start_date = st.date_input("Start date", value=pd.to_datetime('2022-01-01'))
    end_date = st.date_input("End date", value=pd.to_datetime('2023-12-31'))

    start_date_str = start_date.strftime('%Y-%m-%d')
    end_date_str = end_date.strftime('%Y-%m-%d')

    # 이상치 탐지 쿼리 실행
    query_outliers = f"""
        SELECT
            l.profit_center,
            l.profit_center_name,
            l.acc_name,
            l.gc,
            l.posting_date
        FROM
            s_accounting_ledger l
        WHERE
            l.posting_date BETWEEN DATE '{start_date_str}' AND DATE '{end_date_str}'
    """

    if st.button('분석 실행'):
        cursor.execute(query_outliers)
        rows_outliers = cursor.fetchall()
        columns_outliers = [desc[0] for desc in cursor.description]

        # 결과를 데이터프레임으로 변환
        df_outliers = pd.DataFrame(rows_outliers, columns=columns_outliers)

        # Isolation Forest를 사용하여 이상치 탐지
        isolation_forest = IsolationForest(contamination=0.05)
        df_outliers['gc'] = df_outliers['gc'].astype(float)
        isolation_forest.fit(df_outliers[['gc']])
        df_outliers['anomaly'] = isolation_forest.predict(df_outliers[['gc']])
        df_anomalies = df_outliers[df_outliers['anomaly'] == -1]

        # 이상치 데이터프레임 출력
        st.markdown("### 이상치 데이터")
        st.dataframe(df_anomalies, height=800, width=2000)

        # 시각화
        fig, ax = plt.subplots()
        sns.scatterplot(data=df_outliers, x='posting_date', y='gc', hue='anomaly', palette={1: 'blue', -1: 'red'}, ax=ax)
        ax.set_title('이상치 분석 결과')
        st.pyplot(fig)

        # 분석 문구 추가
        st.markdown(f"### 분석 결과")
        st.markdown(f"선택된 기간 동안 총 {len(df_anomalies)} 개의 이상치가 발견되었습니다.")

        # 엑셀 파일로 저장 버튼
        if st.button('Save Anomalies to Excel'):
            with pd.ExcelWriter('c:/anomalies.xlsx') as writer:
                df_anomalies.to_excel(writer, sheet_name='Anomalies', index=False)
            st.success('Data saved to c:/anomalies.xlsx')