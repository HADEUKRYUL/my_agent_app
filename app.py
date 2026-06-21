import streamlit as st
import os
from openai import OpenAI
from dotenv import load_dotenv

# .env 파일 또는 스트림릿 클라우드 설정(Secrets)에서 API 키 로드
load_dotenv()
api_key = os.getenv("OPENAI_API_KEY")

# OpenAI 클라이언트 초기화 (서버 없이 바로 직접 통신!)
client = OpenAI(api_key=api_key)

# 페이지 설정
st.set_page_config(page_title="주인님의 비서", page_icon="🤖", layout="centered")
st.title("🤖 나만의 비서 에이전트")
st.caption("주인님, 무엇을 도와드릴까요?")

# 대화 기록 세션 초기화
if "messages" not in st.session_state:
    st.session_state.messages = [
        {"role": "system", "content": "당신은 주인님의 유능하고 충직한 개인 비서 에이전트입니다. 언제나 정중하게 '주인님'이라고 부르며 보좌하세요."}
    ]

# 이전 대화 내용 보여주기 (시스템 메시지는 숨김)
for message in st.session_state.messages:
    if message["role"] != "system":
        with st.chat_message(message["role"]):
            st.markdown(message["content"])

# 주인님의 입력창
if prompt := st.chat_input("명령을 입력하세요..."):
    # 1. 화면에 주인님 메시지 출력 및 저장
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # 2. AI에게 직접 물어보고 답변 받기
    with st.chat_message("assistant"):
        try:
            with st.spinner("생각 중이옵니다..."):
                response = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=st.session_state.messages,
                    temperature=0.7
                )
                reply = response.choices[0].message.content
                st.markdown(reply)
                # 답변 저장
                st.session_state.messages.append({"role": "assistant", "content": reply})
        except Exception as e:
            st.markdown(f"⚠️ OpenAI 연결에 실패했습니다. API 키나 잔액을 확인해 주세요.\n에러 내용: {e}")