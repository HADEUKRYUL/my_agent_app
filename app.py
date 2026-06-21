import streamlit as st
from openai import OpenAI
import PyPDF2
import pandas as pd
from pptx import Presentation  # 파워포인트 해독 도구 로드

# 스마트폰 및 웹 화면 최적화 설정
st.set_page_config(page_title="주인님의 전용 비서", page_icon="🤖", layout="centered")
st.title("🤖 나만의 특급 비서 에이전트")
st.caption("파워포인트(PPTX) 분석 기능이 추가되었습니다. 문서나 명령을 내려주십시오.")

# 비밀 금고에서 열쇠 가져오기
try:
    my_key = st.secrets["OPENAI_API_KEY"]
except:
    my_key = None

if not my_key:
    st.error("🚨 비밀 금고(Secrets) 연결을 확인해 주세요.")
else:
    client = OpenAI(api_key=my_key)

    # 대화 기록 저장 공간 초기화
    if "messages" not in st.session_state:
        st.session_state.messages = []

    # 📁 왼쪽 사이드바 파일 업로드 기능 (pptx 확장자 추가)
    with st.sidebar:
        st.header("📁 파일 업로드 창")
        uploaded_file = st.file_uploader("분석할 파일(PDF, CSV, TXT, PPTX)을 올려주십시오.", type=["pdf", "csv", "txt", "pptx"])
        file_text = ""

        if uploaded_file is not None:
            try:
                # PDF 파일 읽기
                if uploaded_file.name.endswith('.pdf'):
                    pdf_reader = PyPDF2.PdfReader(uploaded_file)
                    for page in pdf_reader.pages:
                        file_text += page.extract_text() + "\n"
                # CSV 파일 읽기
                elif uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file)
                    file_text = df.to_string()
                # TXT 파일 읽기
                elif uploaded_file.name.endswith('.txt'):
                    file_text = uploaded_file.getvalue().decode("utf-8")
                # ◀ 파워포인트(PPTX) 파일 읽기 로직 추가
                elif uploaded_file.name.endswith('.pptx'):
                    prs = Presentation(uploaded_file)
                    for slide in prs.slides:
                        for shape in slide.shapes:
                            if hasattr(shape, "text"):
                                file_text += shape.text + "\n"
                
                st.success(f"✅ [{uploaded_file.name}] 파일 읽기 완료, 주인님!")
            except Exception as e:
                st.error(f"⚠️ 파일 읽기 에러: {e}")

    # 화면에 이전 대화 내용 표시
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # 주인님의 명령 입력창
    if prompt := st.chat_input("무엇이든 말씀하십시오, 주인님..."):
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        # AI 비서의 답변 처리
        with st.chat_message("assistant"):
            try:
                with st.spinner("보좌할 내용을 분석 중입니다..."):
                    system_instruction = (
                        "당신은 주인님의 유능하고 충직한 개인 비서 에이전트입니다. "
                        "항상 정중하게 '주인님'이라 부르며, 전문적인 기술 지식과 논리적인 사고를 바탕으로 답변하세요."
                    )
                    
                    # 업로드된 파일의 텍스트를 AI 두뇌에 주입
                    if file_text:
                        system_instruction += f"\n\n[주인님이 전달한 참고 문서 내용]\n{file_text[:10000]}"
                    
                    response = client.chat.completions.create(
                        model="gpt-4o",  
                        messages=[{"role": "system", "content": system_instruction}] + st.session_state.messages,
                        temperature=0.7
                    )
                    reply = response.choices[0].message.content
                    st.markdown(reply)
                    st.session_state.messages.append({"role": "assistant", "content": reply})
            except Exception as e:
                st.error(f"⚠️ OpenAI 통신 에러가 발생했습니다: {e}")
