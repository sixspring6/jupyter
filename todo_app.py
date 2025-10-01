#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
로컬에서 구동 가능한 투두 리스트 앱
Simple Todo List App that runs locally using Streamlit
"""

import streamlit as st
import json
import os
from datetime import datetime

# 파일 경로 설정
TODO_FILE = "todos.json"

# 투두 리스트 로드
def load_todos():
    """Load todos from JSON file"""
    if os.path.exists(TODO_FILE):
        try:
            with open(TODO_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return []
    return []

# 투두 리스트 저장
def save_todos(todos):
    """Save todos to JSON file"""
    with open(TODO_FILE, 'w', encoding='utf-8') as f:
        json.dump(todos, f, ensure_ascii=False, indent=2)

# 앱 설정
st.set_page_config(
    page_title="투두 리스트",
    page_icon="✅",
    layout="centered"
)

# 제목
st.title("✅ 투두 리스트")
st.markdown("---")

# 세션 상태 초기화
if 'todos' not in st.session_state:
    st.session_state.todos = load_todos()

# 새로운 할 일 추가
st.subheader("새로운 할 일 추가")
col1, col2 = st.columns([4, 1])

with col1:
    new_todo = st.text_input("할 일을 입력하세요", key="new_todo_input", label_visibility="collapsed")

with col2:
    if st.button("추가", use_container_width=True):
        if new_todo:
            todo_item = {
                "id": len(st.session_state.todos) + 1,
                "task": new_todo,
                "completed": False,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            st.session_state.todos.append(todo_item)
            save_todos(st.session_state.todos)
            st.rerun()

st.markdown("---")

# 필터 옵션
st.subheader("할 일 목록")
filter_option = st.radio(
    "필터",
    ["전체", "진행 중", "완료됨"],
    horizontal=True,
    label_visibility="collapsed"
)

# 투두 필터링
if filter_option == "진행 중":
    filtered_todos = [todo for todo in st.session_state.todos if not todo['completed']]
elif filter_option == "완료됨":
    filtered_todos = [todo for todo in st.session_state.todos if todo['completed']]
else:
    filtered_todos = st.session_state.todos

# 통계
total_count = len(st.session_state.todos)
completed_count = len([todo for todo in st.session_state.todos if todo['completed']])
active_count = total_count - completed_count

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("전체", total_count)
with col2:
    st.metric("진행 중", active_count)
with col3:
    st.metric("완료", completed_count)

st.markdown("---")

# 할 일 목록 표시
if not filtered_todos:
    st.info("할 일이 없습니다.")
else:
    for todo in filtered_todos:
        col1, col2, col3 = st.columns([0.5, 4, 1])
        
        with col1:
            # 체크박스로 완료 상태 토글
            completed = st.checkbox(
                "완료",
                value=todo['completed'],
                key=f"check_{todo['id']}",
                label_visibility="collapsed"
            )
            if completed != todo['completed']:
                # 상태 변경
                for t in st.session_state.todos:
                    if t['id'] == todo['id']:
                        t['completed'] = completed
                        break
                save_todos(st.session_state.todos)
                st.rerun()
        
        with col2:
            # 할 일 내용 표시
            if todo['completed']:
                st.markdown(f"~~{todo['task']}~~")
                st.caption(f"생성: {todo['created_at']}")
            else:
                st.markdown(f"{todo['task']}")
                st.caption(f"생성: {todo['created_at']}")
        
        with col3:
            # 삭제 버튼
            if st.button("🗑️", key=f"delete_{todo['id']}", use_container_width=True):
                st.session_state.todos = [t for t in st.session_state.todos if t['id'] != todo['id']]
                save_todos(st.session_state.todos)
                st.rerun()
        
        st.markdown("---")

# 하단 메뉴
st.markdown("###")
col1, col2 = st.columns(2)

with col1:
    if st.button("모두 완료로 표시", use_container_width=True):
        for todo in st.session_state.todos:
            todo['completed'] = True
        save_todos(st.session_state.todos)
        st.rerun()

with col2:
    if st.button("완료된 항목 삭제", use_container_width=True):
        st.session_state.todos = [todo for todo in st.session_state.todos if not todo['completed']]
        save_todos(st.session_state.todos)
        st.rerun()

# 앱 실행 안내
st.markdown("---")
st.caption("💡 이 앱을 실행하려면: `streamlit run todo_app.py`")
