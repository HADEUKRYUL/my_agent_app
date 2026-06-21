import streamlit as st
from openai import OpenAI

st.set_page_config(page_title="주인님의 비서", page_icon="🤖")
st.title("🤖 나만의 비서 에이전트")

# 비밀 금고에서 열쇠를 강제로 찾아오는 마법의 코드
try:
    my_key = st.secrets["OPENAI_API_KEY"]
except:
    my_key = None

if not my_key:
    st.error("🚨 비밀 금고(Secrets)가 비어있거나 이름이 틀렸습니다! 우측 하단 Manage app -> Settings -> Secrets를 다시 확인해 주세요.")
else:
    st.success("✅ 비밀 금고 연결 성공! 비서가 깨어났습니다.")
    client = OpenAI(api_key=my_key)

    if "messages" not in st.session_state:
        st.session_state.messages = []

    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if prompt := st.chat_input("명령을 내려주십시오, 주인님."):
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        with st.chat_message("assistant"):
            try:
                response = client.chat.completions.create(
                    model="gpt-4o-mini",
                    messages=[{"role": "system", "content": "당신은 유능한 비서입니다. 항상 정중하게 대답하세요."}] + st.session_state.messages
                )
                reply = response.choices[0].message.content
                st.markdown(reply)
                st.session_state.messages.append({"role": "assistant", "content": reply})
            except Exception as e:
                st.error(f"⚠️ OpenAI 통신 에러: {e}")
