import streamlit as st
from openai import OpenAI
import PyPDF2
import pandas as pd
from pptx import Presentation
import base64

st.set_page_config(page_title="주인님의 전용 비서", page_icon="🤖", layout="wide")
st.title("🤖 나만의 특급 비서 에이전트")

# API 키 설정
try:
    my_key = st.secrets["OPENAI_API_KEY"]
except:
    my_key = None

if not my_key:
    st.error("🚨 API 키를 설정해 주십시오.")
    st.stop()

client = OpenAI(api_key=my_key)

# [핵심] 대화와 문서 내용을 모두 기억하는 영구 저장소
if "memory" not in st.session_state:
    st.session_state.memory = []

# 사이드바 (파일 업로드)
with st.sidebar:
    st.header("📁 데이터 분석실")
    uploaded_file = st.file_uploader("파일을 올려주십시오.", type=["pdf", "csv", "txt", "pptx", "png", "jpg", "jpeg"])
    
    if uploaded_file:
        file_text = ""
        # 파일 유형별 텍스트 추출 및 메모리 저장
        if uploaded_file.type.startswith("image"):
            image_data = base64.b64encode(uploaded_file.read()).decode('utf-8')
            st.session_state.last_image = image_data
            st.success(f"✅ [{uploaded_file.name}] 사진을 기억했습니다.")
        else:
            try:
                if uploaded_file.name.endswith('.pdf'):
                    reader = PyPDF2.PdfReader(uploaded_file)
                    file_text = "\n".join([p.extract_text() for p in reader.pages])
                elif uploaded_file.name.endswith('.csv'):
                    file_text = pd.read_csv(uploaded_file).to_string()
                elif uploaded_file.name.endswith('.pptx'):
                    prs = Presentation(uploaded_file)
                    file_text = "\n".join([shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")])
                else:
                    file_text = uploaded_file.read().decode("utf-8")
                
                # [기억] 분석한 문서 내용을 메모리에 영구 저장
                st.session_state.memory.append({"role": "system", "content": f"이전에 학습한 문서 내용: {file_text[:5000]}"})
                st.success(f"✅ [{uploaded_file.name}] 내용을 완전히 학습했습니다.")
            except Exception as e:
                st.error(f"⚠️ 읽기 실패: {e}")

# 대화 기록 표시
for msg in st.session_state.memory:
    if msg["role"] != "system":
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

# 채팅창 구현
if prompt := st.chat_input("질문을 입력하십시오..."):
    st.session_state.memory.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # 대화 맥락 구성 (메모리에 있는 모든 이전 기록 참조)
    messages = [{"role": "system", "content": "당신은 주인님의 충직한 비서입니다. 과거 대화와 제공된 문서를 모두 기억하고 분석하세요."}] + st.session_state.memory

    with st.chat_message("assistant"):
        response = client.chat.completions.create(model="gpt-4o", messages=messages)
        reply = response.choices[0].message.content
        st.markdown(reply)
        st.session_state.memory.append({"role": "assistant", "content": reply})
