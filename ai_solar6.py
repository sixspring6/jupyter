import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime
from trino.dbapi import connect
from trino.auth import BasicAuthentication
from dateutil.relativedelta import relativedelta
import os
import ssl
import certifi
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.ssl_ import create_urllib3_context
import tensorflow as tf
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import LSTM, Dense
from sklearn.preprocessing import MinMaxScaler

# SSL 설정
pem_path = r"C:\HMGSecureROOTCA.pem"
os.environ['CURL_CA_BUNDLE'] = ''
os.environ['SSL_CERT_FILE'] = ''
os.environ['REQUESTS_CA_BUNDLE'] = ''
ssl._create_default_https_context = ssl._create_unverified_context

class CustomHttpAdapter(HTTPAdapter):
    def __init__(self, ssl_context=None, **kwargs):
        self.ssl_context = ssl_context
        super().__init__(**kwargs)

    def init_poolmanager(self, *args, **kwargs):
        context = create_urllib3_context(cert_reqs=ssl.CERT_REQUIRED)
        try:
            if os.path.exists(pem_path):
                context.load_verify_locations(cafile=pem_path)
            else:
                print(f"Warning: PEM file not found at {pem_path}. Using default certificates.")
                context.load_verify_locations(cafile=certifi.where())
        except Exception as e:
            print(f"Error loading PEM file: {e}")
            print("Falling back to default certificates.")
            context.load_verify_locations(cafile=certifi.where())
        
        kwargs['ssl_context'] = self.ssl_context or context
        return super().init_poolmanager(*args, **kwargs)

# requests 세션 설정
session = requests.Session()
adapter = CustomHttpAdapter()
session.mount('https://', adapter)

st.set_page_config(layout="wide")

# 전역 변수 설정
TIME_STEPS = 12

def get_connection(host, port, user, password, catalog, schema):
    return connect(
        host=host,
        port=port,
        http_scheme='https',
        auth=BasicAuthentication(user, password),
        catalog=catalog,
        schema=schema,
    )

@st.cache_data(ttl=3600)
def get_data(year, month):
    try:
        st.write(f"데이터베이스 연결 시도 중... (조회 년도: {year}, 조회 월: {month})")
        conn = get_connection(
            host='data-query.hec.co.kr',
            port=443,
            user='0900051',
            password='roseofpA1!!!!!',
            catalog='dw',
            schema='fie'
        )
        st.write("데이터베이스 연결 성공")

        cursor = conn.cursor()
        
        query = f"""
        WITH monthly_data AS (
            SELECT 
                acc_name,
                acc_num,
                DATE_TRUNC('month', posting_date) AS month,
                SUM(gc) AS monthly_gc
            FROM s_accounting_ledger
            WHERE DATE_TRUNC('month', posting_date) <= DATE '{year}-{month:02d}-01'
              AND (CAST(acc_num AS VARCHAR) LIKE '1%' OR CAST(acc_num AS VARCHAR) LIKE '2%' OR CAST(acc_num AS VARCHAR) LIKE '3%')
            GROUP BY acc_name, acc_num, DATE_TRUNC('month', posting_date)
        ),
        cumulative_data AS (
            SELECT 
                acc_name,
                acc_num,
                month,
                SUM(monthly_gc) OVER (
                    PARTITION BY acc_num
                    ORDER BY month
                ) AS cumulative_gc
            FROM monthly_data
        )
        SELECT 
            acc_name,
            acc_num,
            CAST(DATE_FORMAT(month, '%Y-%m') AS VARCHAR) AS month,
            cumulative_gc
        FROM cumulative_data
        ORDER BY acc_num, month
        """
        
        cursor.execute(query)
        
        rows = cursor.fetchall()
        columns = [desc[0] for desc in cursor.description]
        df = pd.DataFrame(rows, columns=columns)
        
        cursor.close()
        conn.close()
        
        if not df.empty:
            df['account_type'] = df['acc_num'].astype(str).str[0].map({'1': '자산', '2': '부채', '3': '자본'})
            df['cumulative_gc'] = df['cumulative_gc'] * df['account_type'].map({'자산': 1, '부채': -1, '자본': -1})
            df['month'] = pd.to_datetime(df['month'])
        
        return df
    except Exception as e:
        st.error(f"데이터 로딩 중 오류 발생: {str(e)}")
        return pd.DataFrame()

def analyze_time_series(df, end_date):
    analysis_results = {}
    
    for acc_num in df['acc_num'].unique():
        acc_data = df[df['acc_num'] == acc_num]
        acc_name = acc_data['acc_name'].iloc[0]
        
        series = acc_data.set_index('month')['cumulative_gc']
        series = series[series.index <= end_date]
        
        if len(series) >= 2:
            start_date = series.index.min()
            last_value = series.iloc[-1]
            first_value = series.iloc[0]
            total_change = last_value - first_value
            percent_change = ((last_value - first_value) / abs(first_value)) * 100 if first_value != 0 else float('inf')
            
            is_anomaly = abs(percent_change) > 50
            
            trend = 'Increasing' if total_change > 0 else 'Decreasing' if total_change < 0 else 'Stable'
            
            analysis_results[acc_num] = {
                'acc_name': acc_name,
                'start_date': start_date,
                'end_date': end_date,
                'first_value': first_value,
                'last_value': last_value,
                'total_change': total_change,
                'percent_change': percent_change,
                'trend': trend,
                'is_anomaly': is_anomaly
            }
        else:
            analysis_results[acc_num] = {
                'acc_name': acc_name,
                'start_date': series.index.min() if not series.empty else None,
                'end_date': series.index.max() if not series.empty else None,
                'first_value': series.iloc[0] if not series.empty else None,
                'last_value': series.iloc[-1] if not series.empty else None,
                'total_change': None,
                'percent_change': None,
                'trend': 'Insufficient data',
                'is_anomaly': None
            }
    
    return analysis_results

def plot_time_series(df, acc_num, end_date):
    acc_data = df[df['acc_num'] == acc_num]
    series = acc_data.set_index('month')['cumulative_gc']
    series = series[series.index <= end_date]
    acc_name = acc_data['acc_name'].iloc[0]
    
    fig = go.Figure()
    
    fig.add_trace(go.Scatter(x=series.index, y=series,
                             mode='lines+markers', name='잔액', line=dict(color='blue')))

    fig.update_layout(title=f'{acc_num} - {acc_name} 월별 잔액 추이', 
                      xaxis_title='월', 
                      yaxis_title='잔액',
                      legend=dict(yanchor="top", y=0.99, xanchor="left", x=0.01))
    return fig

@st.cache_data(ttl=3600)
def get_liquidity_data(year, month):
    try:
        conn = get_connection(
            host='data-query.hec.co.kr',
            port=443,
            user='0900051',
            password='roseofpA1!!!!!',
            catalog='dw',
            schema='fie'
        )

        cursor = conn.cursor()
        
        query = f"""
        SELECT 
            sl.acc_num,
            sl.acc_name,
            sl.profit_center,
            sl.profit_center_name,
            ap.sector,
            SUM(sl.gc) AS total_gc
        FROM s_accounting_ledger sl
        LEFT JOIN s_accounting_project ap ON sl.profit_center = ap.profit_center
        WHERE DATE_TRUNC('month', sl.posting_date) <= DATE '{year}-{month:02d}-01'
          AND ((CAST(sl.acc_num AS BIGINT) >= 11010110 AND CAST(sl.acc_num AS BIGINT) <= 11110190)
           OR (CAST(sl.acc_num AS BIGINT) >= 17000000 AND CAST(sl.acc_num AS BIGINT) <= 17999999))
        GROUP BY sl.acc_num, sl.acc_name, sl.profit_center, sl.profit_center_name, ap.sector
        ORDER BY sl.acc_num, sl.profit_center
        """
        
        cursor.execute(query)
        
        rows = cursor.fetchall()
        
        columns = [desc[0] for desc in cursor.description]
        df = pd.DataFrame(rows, columns=columns)
        
        cursor.close()
        conn.close()
        
        return df
    except Exception as e:
        st.error(f"데이터 로딩 중 오류 발생: {str(e)}")
        return pd.DataFrame()

@st.cache_data(ttl=3600)
def get_monthly_liquidity_data(year, month, num_months=12):
    try:
        conn = get_connection(
            host='data-query.hec.co.kr',
            port=443,
            user='0900051',
            password='roseofpA1!!!!!',
            catalog='dw',
            schema='fie'
        )

        cursor = conn.cursor()
        
        end_date = datetime(year, month, 1)
        start_date = end_date - relativedelta(months=num_months-1)
        
        query = f"""
        WITH date_series AS (
            SELECT DATE_TRUNC('month', d) AS month
            FROM UNNEST(SEQUENCE(DATE '{start_date.strftime('%Y-%m-%d')}', DATE '{end_date.strftime('%Y-%m-%d')}', INTERVAL '1' MONTH)) AS t(d)
        )
        SELECT 
            ds.month,
            COALESCE(SUM(sl.gc), 0) AS total_gc
        FROM date_series ds
        LEFT JOIN s_accounting_ledger sl
            ON DATE_TRUNC('month', sl.posting_date) <= ds.month
            AND ((CAST(sl.acc_num AS BIGINT) >= 11010110 AND CAST(sl.acc_num AS BIGINT) <= 11110190)
                OR (CAST(sl.acc_num AS BIGINT) >= 17000000 AND CAST(sl.acc_num AS BIGINT) <= 17999999))
        GROUP BY ds.month
        ORDER BY ds.month DESC
        """
        
        cursor.execute(query)
        
        rows = cursor.fetchall()
        
        columns = [desc[0] for desc in cursor.description]
        df = pd.DataFrame(rows, columns=columns)
        
        cursor.close()
        conn.close()
        
        df['month'] = pd.to_datetime(df['month'])
        
        return df
    except Exception as e:
        st.error(f"월별 데이터 로딩 중 오류 발생: {str(e)}")
        return pd.DataFrame()

@st.cache_resource
def create_lstm_model(input_shape):
    model = Sequential([
        LSTM(50, activation='relu', input_shape=input_shape, return_sequences=True),
        LSTM(50, activation='relu'),
        Dense(1)
    ])
    model.compile(optimizer='adam', loss='mse')
    return model

def prepare_data_for_lstm(data, time_steps):
    X, y = [], []
    for i in range(len(data) - time_steps):
        X.append(data[i:(i + time_steps), 0])
        y.append(data[i + time_steps, 0])
    return np.array(X), np.array(y)

def predict_future(model, last_sequence, num_future_steps, time_steps):
    future_predictions = []
    current_sequence = last_sequence.copy()
    
    for _ in range(num_future_steps):
        prediction = model.predict(current_sequence.reshape(1, time_steps, 1))
        future_predictions.append(prediction[0, 0])
        current_sequence = np.roll(current_sequence, -1)
        current_sequence[-1] = prediction
    
    return future_predictions

def predictive_analysis(df, acc_num, year, month, future_months=6):
    acc_data = df[df['acc_num'] == acc_num].sort_values('month')
    time_series = acc_data['cumulative_gc'].values.reshape(-1, 1)
    
    scaler = MinMaxScaler()
    scaled_data = scaler.fit_transform(time_series)
    
    X, y = prepare_data_for_lstm(scaled_data, TIME_STEPS)
    X = X.reshape((X.shape[0], X.shape[1], 1))
    
    model = create_lstm_model((TIME_STEPS, 1))
    model.fit(X, y, epochs=100, batch_size=32, verbose=0)
    
    last_sequence = scaled_data[-TIME_STEPS:]
    future_scaled = predict_future(model, last_sequence, future_months, TIME_STEPS)
    future_predictions = scaler.inverse_transform(np.array(future_scaled).reshape(-1, 1))
    
    last_date = acc_data['month'].max()
    future_dates = pd.date_range(start=last_date, periods=future_months+1, freq='M')[1:]
    
    return pd.DataFrame({'month': future_dates, 'predicted_gc': future_predictions.flatten()})

def main():
    st.title('Accounting AI Monitoring System')

    st.markdown("""
        <style>
        div.stButton > button:first-child {
            background-color: #0099ff;
            color: white;
        }
        div.stButton > button:hover {
            background-color: #00ff00;
            color: white;
        }
        </style>""", unsafe_allow_html=True)

    if 'menu' not in st.session_state:
        st.session_state.menu = "계정추세 분석"

    st.sidebar.title("메뉴")
    col1, col2, col3 = st.sidebar.columns(3)
    
    if col1.button("계정추세 분석", key="btn_account"):
        st.session_state.menu = "계정추세 분석"
    
    if col2.button("유동성 분석", key="btn_liquidity"):
        st.session_state.menu = "유동성 분석"
    
    if col3.button("예측 분석", key="btn_prediction"):
        st.session_state.menu = "예측 분석"

    st.sidebar.markdown(f"**현재 메뉴: {st.session_state.menu}**")

    current_year = datetime.now().year
    year = st.selectbox("조회 년도", range(2010, current_year + 1), index=current_year - 2010, key='year')
    month = st.selectbox("조회 월", range(1, 13), index=datetime.now().month - 1, key='month')

    if st.session_state.menu == "계정추세 분석":
        if st.button("계정추세 분석 실행"):
            try:
                df = get_data(year, month)
                if df.empty:
                    st.warning("선택된 기간에 데이터가 없습니다.")
                    return

                end_date = pd.to_datetime(f"{year}-{month:02d}-01")
                analysis_results = analyze_time_series(df, end_date)

                st.subheader(f"{year}년 {month}월까지의 계정별 잔액 추세 분석")
                result_df = pd.DataFrame(analysis_results).T
                result_df = result_df.rename(columns={
                    'acc_name': '계정명',
                    'start_date': '시작일',
                    'end_date': '종료일',
                    'first_value': '시작 잔액',
                    'last_value': '최종 잔액',
                    'total_change': '총 변동액',
                    'percent_change': '변동률 (%)',
                    'trend': '추세',
                    'is_anomaly': '이상치 감지'
                })
                
                st.dataframe(result_df)

                anomalies = result_df[result_df['이상치 감지'] == True]
                if not anomalies.empty:
                    st.subheader("이상치가 감지된 계정 분석")
                    for acc_num in anomalies.index:
                        st.write(f"### {acc_num} - {anomalies.loc[acc_num, '계정명']}")
                        fig = plot_time_series(df, acc_num, end_date)
                        st.plotly_chart(fig, use_container_width=True)

                        st.write(f"시작 잔액: {anomalies.loc[acc_num, '시작 잔액']:,.2f}")
                        st.write(f"최종 잔액: {anomalies.loc[acc_num, '최종 잔액']:,.2f}")
                        st.write(f"총 변동액: {anomalies.loc[acc_num, '총 변동액']:,.2f}")
                        st.write(f"변동률: {anomalies.loc[acc_num, '변동률 (%)']:,.2f}%")
                        st.write(f"추세: {anomalies.loc[acc_num, '추세']}")

                        st.warning("이 계정은 변동률이 50% 이상으로, 이상치로 감지되었습니다. 자세한 검토가 필요할 수 있습니다.")
                else:
                    st.info("이상치가 감지된 계정이 없습니다.")

            except Exception as e:
                st.error(f"계정추세 분석 중 오류가 발생했습니다: {str(e)}")
                st.exception(e)



    elif st.session_state.menu == "유동성 분석":
        if st.button("유동성 분석 실행"):
            try:
                df = get_liquidity_data(year, month)
                monthly_df = get_monthly_liquidity_data(year, month)

                if df.empty or monthly_df.empty:
                    st.warning("선택된 기간에 데이터가 없습니다.")
                    return

                st.subheader(f"{year}년 {month}월까지의 유동성 분석")

                # Profit Center별 분석
                df['profit_center'] = df['profit_center'].fillna('비어있음')
                df['profit_center_name'] = df['profit_center_name'].fillna('비어있음')
                df['sector'] = df['sector'].fillna('비어있음')

                summary_df = df.groupby(['profit_center', 'profit_center_name'])['total_gc'].sum().reset_index()
                summary_df = summary_df.sort_values('total_gc', ascending=False)

                total_sum = summary_df['total_gc'].sum()

                st.write("Profit Center별 유동성 현황")
                st.dataframe(summary_df[['profit_center_name', 'total_gc']])
                st.write(f"**총계: {total_sum:,.2f}**")

                top_20 = summary_df.nlargest(20, 'total_gc')
                fig = go.Figure(data=[go.Bar(
                    x=top_20['profit_center_name'],
                    y=top_20['total_gc'],
                    hovertemplate='Profit Center: %{x}<br>금액: %{y:,.0f}'
                )])
                fig.update_layout(
                    title='Profit Center별 유동성 (상위 20개)',
                    xaxis_title='Profit Center Name',
                    yaxis_title='금액',
                    hoverlabel=dict(bgcolor="white", font_size=12)
                )
                st.plotly_chart(fig, use_container_width=True)

                # Sector별 분석
                sector_summary = df.groupby('sector')['total_gc'].sum().reset_index()
                sector_summary = sector_summary.sort_values('total_gc', ascending=False)

                st.write("Sector별 유동성 현황")
                st.dataframe(sector_summary)

                fig = go.Figure(data=[go.Bar(
                    x=sector_summary['sector'],
                    y=sector_summary['total_gc'],
                    hovertemplate='Sector: %{x}<br>금액: %{y:,.0f}'
                )])
                fig.update_layout(
                    title='Sector별 유동성',
                    xaxis_title='Sector',
                    yaxis_title='금액',
                    hoverlabel=dict(bgcolor="white", font_size=12)
                )
                st.plotly_chart(fig, use_container_width=True)

                # 월별 누적 금액 분석
                st.subheader("월별 누적 금액 분석")
                st.dataframe(monthly_df)

                fig = go.Figure(data=[go.Bar(
                    x=monthly_df['month'],
                    y=monthly_df['total_gc'],
                    hovertemplate='%{x}<br>누적 금액: %{y:,.0f}'
                )])
                fig.update_layout(
                    title='월별 누적 유동성',
                    xaxis_title='월',
                    yaxis_title='누적 금액',
                    hoverlabel=dict(bgcolor="white", font_size=12)
                )
                st.plotly_chart(fig, use_container_width=True)

                st.subheader("상세 데이터")
                st.dataframe(df)

            except Exception as e:
                st.error(f"유동성 분석 중 오류 발생했습니다: {str(e)}")
                st.exception(e)


    elif st.session_state.menu == "예측 분석":
        try:
            df = get_data(year, month)
            if df.empty:
                st.warning("선택된 기간에 데이터가 없습니다.")
                return
            
            if 'selected_acc_num' not in st.session_state:
                st.session_state.selected_acc_num = df['acc_num'].unique()[0]
            
            acc_num = st.selectbox("분석할 계정 선택", df['acc_num'].unique(), key='acc_num', index=list(df['acc_num'].unique()).index(st.session_state.selected_acc_num))
            st.session_state.selected_acc_num = acc_num

            future_months = st.slider("예측 개월 수", min_value=1, max_value=12, value=6, key='future_months')
            
            if st.button("예측 분석 실행"):
                predictions = predictive_analysis(df, acc_num, year, month, future_months)
                
                st.subheader(f"{acc_num} 계정의 향후 {future_months}개월 예측")
                st.write(predictions)
                
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=df[df['acc_num'] == acc_num]['month'], y=df[df['acc_num'] == acc_num]['cumulative_gc'],
                                         mode='lines+markers', name='실제 데이터'))
                fig.add_trace(go.Scatter(x=predictions['month'], y=predictions['predicted_gc'],
                                         mode='lines+markers', name='예측 데이터', line=dict(dash='dash')))
                fig.update_layout(title=f'{acc_num} 계정의 실제 데이터와 예측', xaxis_title='월', yaxis_title='금액')
                st.plotly_chart(fig, use_container_width=True)
                
        except Exception as e:
            st.error(f"예측 분석 중 오류가 발생했습니다: {str(e)}")
            st.exception(e)

    st.sidebar.info("이 대시보드는 회사의 자산/부채/자본 계정의 월별 누계 잔액을 분석합니다.")

if __name__ == "__main__":
    main()