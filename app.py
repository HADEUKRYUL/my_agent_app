import streamlit as st
from openai import OpenAI
import PyPDF2
import pandas as pd
from pptx import Presentation
import base64

# 설정 최적화
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

# 상태 저장
if "messages" not in st.session_state:
    st.session_state.messages = []

# 사이드바 (파일 업로드)
with st.sidebar:
    st.header("📁 데이터 분석실")
    uploaded_file = st.file_uploader("파일을 올려주십시오.", type=["pdf", "csv", "txt", "pptx", "png", "jpg", "jpeg"])
    
    file_content = None
    image_data = None

    if uploaded_file:
        # 이미지 파일 처리
        if uploaded_file.type.startswith("image"):
            image_data = base64.b64encode(uploaded_file.read()).decode('utf-8')
            st.image(uploaded_file, caption="분석 대기 중인 사진")
        # 문서 파일 처리
        else:
            try:
                if uploaded_file.name.endswith('.pdf'):
                    reader = PyPDF2.PdfReader(uploaded_file)
                    file_content = "\n".join([p.extract_text() for p in reader.pages])
                elif uploaded_file.name.endswith('.csv'):
                    file_content = pd.read_csv(uploaded_file).to_string()
                elif uploaded_file.name.endswith('.pptx'):
                    prs = Presentation(uploaded_file)
                    file_content = "\n".join([shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")])
                else:
                    file_content = uploaded_file.read().decode("utf-8")
                st.success("✅ 파일 내용을 성공적으로 추출했습니다!")
            except Exception as e:
                st.error(f"⚠️ 읽기 실패: {e}")

# 채팅창 구현
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

if prompt := st.chat_input("질문을 입력하십시오..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # 메시지 구성
    messages = [{"role": "system", "content": "당신은 주인님의 유능한 비서입니다. 제공된 문서나 이미지 데이터를 정밀하게 분석하여 답변하세요."}]
    for m in st.session_state.messages:
        messages.append(m)

    # 이미지나 문서가 있다면 추가
    if image_data:
        messages[-1] = {"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{image_data}"}}]}
    elif file_content:
        messages[-1] = {"role": "user", "content": f"{prompt}\n\n[참고 문서 내용]:\n{file_content[:15000]}"}

    with st.chat_message("assistant"):
        response = client.chat.completions.create(model="gpt-4o", messages=messages)
        reply = response.choices[0].message.content
        st.markdown(reply)
        st.session_state.messages.append({"role": "assistant", "content": reply})
