import streamlit as st
from openai import OpenAI

# 스마트폰 및 웹 화면 최적화 설정
st.set_page_config(page_title="주인님의 전용 비서", page_icon="🤖", layout="centered")
st.title("🤖 나만의 특급 비서 에이전트")
st.caption("최신 두뇌로 업그레이드 완료되었습니다. 주인님, 명령을 내려주십시오.")

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

    # 화면에 이전 대화 내용 표시
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # 주인님의 명령 입력창
    if prompt := st.chat_input("무엇이든 말씀하십시오, 주인님..."):
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        # AI 비서의 업그레이드 답변 처리
        with st.chat_message("assistant"):
            try:
                with st.spinner("보좌할 내용을 분석 중입니다..."):
                    # [★업그레이드 핵심 지침] 비서의 성격과 역할을 더 똑똑하게 규정합니다.
                    system_instruction = (
                        "당신은 주인님의 유능하고 충직한 개인 비서 에이전트입니다. "
                        "언제나 정중하고 공손하게 '주인님'이라고 부르며 보좌하세요. "
                        "비즈니스 이메일 작성, 일정 관리, 전문적인 기술 조언, 언어 학습 등 "
                        "모든 분야에서 최고 수준의 전문가로서 깊이 있고 정확한 답변을 제공해야 합니다."
                    )
                    
                    response = client.chat.completions.create(
                        model="gpt-4o",  # ◀ 기존 mini 모델에서 가장 똑똑한 최신 고성능 'gpt-4o'로 업그레이드!
                        messages=[{"role": "system", "content": system_instruction}] + st.session_state.messages,
                        temperature=0.7
                    )
                    reply = response.choices[0].message.content
                    st.markdown(reply)
                    st.session_state.messages.append({"role": "assistant", "content": reply})
            except Exception as e:
                st.error(f"⚠️ OpenAI 통신 에러가 발생했습니다: {e}")
