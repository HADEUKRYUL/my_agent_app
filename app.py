import streamlit as st
from openai import OpenAI
import PyPDF2
import pandas as pd
from pptx import Presentation
import base64
import json
import os

# 1. 대화 기록 파일 경로 설정
HISTORY_FILE = "chat_history.json"

def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_history(messages):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(messages, f, ensure_ascii=False, indent=4)

# 설정 및 초기화
st.set_page_config(page_title="주인님의 전용 비서", page_icon="🤖", layout="wide")
st.title("🤖 나만의 특급 비서 에이전트")

try:
    my_key = st.secrets["OPENAI_API_KEY"]
except:
    my_key = None

if not my_key:
    st.error("🚨 API 키를 설정해 주십시오.")
    st.stop()

client = OpenAI(api_key=my_key)

# 대화 로드
if "messages" not in st.session_state:
    st.session_state.messages = load_history()

# 사이드바 및 파일 처리 (기존 기능 유지)
with st.sidebar:
    st.header("📁 데이터 분석실")
    if st.button("🗑️ 대화 기록 초기화"):
        st.session_state.messages = []
        save_history([])
        st.rerun()

    uploaded_file = st.file_uploader("파일/사진 업로드", type=["pdf", "csv", "txt", "pptx", "png", "jpg", "jpeg"])
    
    file_content = None
    image_data = None

    if uploaded_file:
        try:
            if uploaded_file.type.startswith("image"):
                image_data = base64.b64encode(uploaded_file.read()).decode('utf-8')
                st.image(uploaded_file, caption="분석 대기 중")
            else:
                if uploaded_file.name.endswith('.pdf'):
                    file_content = "\n".join([p.extract_text() for p in PyPDF2.PdfReader(uploaded_file).pages])
                elif uploaded_file.name.endswith('.csv'):
                    file_content = pd.read_csv(uploaded_file).to_string()
                elif uploaded_file.name.endswith('.pptx'):
                    prs = Presentation(uploaded_file)
                    file_content = "\n".join([shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")])
                else:
                    file_content = uploaded_file.read().decode("utf-8")
                st.success("✅ 파일 학습 완료!")
        except Exception as e:
            st.error(f"⚠️ 읽기 실패: {e}")

    st.markdown("---")
    st.link_button("🎨 Canva 바로가기", "https://www.canva.com/")

# 화면 출력
for msg in st.session_state.messages:
    if msg["role"] != "system":
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

# 입력 처리
if prompt := st.chat_input("질문을 입력하십시오..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    messages = [{"role": "system", "content": "과거 대화와 제공된 문서를 모두 기억하고 분석하는 비서입니다."}] + st.session_state.messages
    
    if image_data:
        messages[-1] = {"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_data}"}}]}
    elif file_content:
        messages[-1] = {"role": "user", "content": f"{prompt}\n\n[참고 내용]:\n{file_content[:5000]}"}

    with st.chat_message("assistant"):
        response = client.chat.completions.create(model="gpt-4o", messages=messages)
        reply = response.choices[0].message.content
        st.markdown(reply)
        st.session_state.messages.append({"role": "assistant", "content": reply})
        
        # [핵심] 대화 내용 파일로 저장
        save_history(st.session_state.messages)
