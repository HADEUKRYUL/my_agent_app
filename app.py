import streamlit as st
from openai import OpenAI
import PyPDF2
import pandas as pd
from pptx import Presentation
import base64
import json
import os
import datetime
from apscheduler.schedulers.background import BackgroundScheduler

# 💾 1. 영구 기억 저장소 설정
HISTORY_FILE = "chat_history.json"

def load_history():
    if os.path.exists(HISTORY_FILE):
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_history(messages):
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(messages, f, ensure_ascii=False, indent=4)

# ⏰ 2. 일정 관리 알람 (스케줄러) 활성화
if "scheduler" not in st.session_state:
    st.session_state.scheduler = BackgroundScheduler()
    st.session_state.scheduler.start()

# ⚙️ 3. 페이지 기본 설정
st.set_page_config(page_title="주인님의 전용 비서", page_icon="🤖", layout="wide")
st.title("🤖 나만의 특급 비서 에이전트")
st.caption("제조 현장의 데이터 분석부터 개인 일정 관리까지, 완벽하게 보좌하겠습니다.")

# 🔑 4. API 키 확인
try:
    my_key = st.secrets["OPENAI_API_KEY"]
except:
    my_key = None

if not my_key:
    st.error("🚨 우측 하단 Manage app -> Settings -> Secrets에서 API 키를 설정해 주십시오.")
    st.stop()

client = OpenAI(api_key=my_key)

# 🧠 5. 대화 기록 로드 (창을 닫아도 기억 유지)
if "messages" not in st.session_state:
    st.session_state.messages = load_history()

# 📁 6. 왼쪽 서랍(사이드바) 기능 (파일/사진 업로드, 캔바)
with st.sidebar:
    st.header("📁 데이터 분석 & 편의 기능")
    
    if st.button("🗑️ 대화 기록 완전히 지우기"):
        st.session_state.messages = []
        save_history([])
        st.rerun()

    uploaded_file = st.file_uploader("파일/사진 업로드 (PDF, CSV, TXT, PPTX, PNG, JPG)", type=["pdf", "csv", "txt", "pptx", "png", "jpg", "jpeg"])
    
    file_content = None
    image_data = None
    mime_type = None

    if uploaded_file:
        try:
            # 사진(Vision) 파일 처리
            if uploaded_file.type.startswith("image"):
                image_data = base64.b64encode(uploaded_file.read()).decode('utf-8')
                mime_type = uploaded_file.type
                st.image(uploaded_file, caption="분석 대기 중인 사진")
                st.success("✅ 사진을 기억했습니다! 무엇을 분석할까요?")
            # 일반 문서 파일 처리
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
                st.success("✅ 문서 학습 완료! 질문해 주십시오.")
        except Exception as e:
            st.error(f"⚠️ 파일 읽기 실패: {e}")

    st.markdown("---")
    st.markdown("🎨 **디자인 스튜디오**")
    st.link_button("Canva (캔바) 바로가기", "https://www.canva.com/")

# 💬 7. 화면에 과거 대화 기록 출력
for msg in st.session_state.messages:
    if msg["role"] != "system":
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

# ⌨️ 8. 채팅 입력 및 명령 처리
if prompt := st.chat_input("명령을 내려주십시오. (그림 요청: '/그리기 내용', 일정 요청: '일정/알람' 포함)"):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    # ▶ 기능 A: 그림 그리기 (DALL-E 3)
    if prompt.startswith("/그리기"):
        with st.chat_message("assistant"):
            with st.spinner("🎨 그림을 생성하고 있습니다..."):
                try:
                    draw_prompt = prompt.replace("/그리기", "").strip()
                    response = client.images.generate(
                        model="dall-e-3", prompt=draw_prompt, size="1024x1024", quality="standard", n=1
                    )
                    image_url = response.data[0].url
                    st.image(image_url, caption=f"완성된 그림: {draw_prompt}")
                    reply = "요청하신 그림을 완성했습니다! 꾹 눌러 저장하신 뒤 좌측 캔바(Canva)를 통해 활용해 보십시오."
                    st.markdown(reply)
                    st.session_state.messages.append({"role": "assistant", "content": reply})
                    save_history(st.session_state.messages)
                except Exception as e:
                    st.error(f"⚠️ 그림 생성 에러: {e}")
    
    # ▶ 기능 B: 일반 대화, 사진/문서 분석, 일정 알람 처리
    else:
        # 일정 관련 키워드 감지 시 시스템 메시지 팝업
        if "일정" in prompt or "알람" in prompt:
            st.toast("⏰ 일정 관리 모드가 감지되었습니다. 기억해 두겠습니다!", icon="🔔")
        
        with st.chat_message("assistant"):
            with st.spinner("분석 중입니다..."):
                system_instruction = "당신은 주인님의 충직한 비서입니다. 제공된 문서, 사진, 그리고 과거 대화를 모두 기억하여 맥락에 맞게 완벽하게 보좌하세요."
                messages_for_api = [{"role": "system", "content": system_instruction}] + st.session_state.messages[:-1]
                
                # 파일이나 사진이 업로드된 경우, AI 두뇌에 데이터 첨부
                if image_data:
                    messages_for_api.append({
                        "role": "user", 
                        "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{image_data}"}}]
                    })
                elif file_content:
                    messages_for_api.append({
                        "role": "user", 
                        "content": f"{prompt}\n\n[참고 문서 내용]:\n{file_content[:5000]}"
                    })
                else:
                    messages_for_api.append({"role": "user", "content": prompt})

                try:
                    response = client.chat.completions.create(model="gpt-4o", messages=messages_for_api, temperature=0.7)
                    reply = response.choices[0].message.content
                    st.markdown(reply)
                    
                    # 대화 저장소에 기록
                    st.session_state.messages.append({"role": "assistant", "content": reply})
                    save_history(st.session_state.messages)
                except Exception as e:
                    st.error(f"⚠️ 통신 에러가 발생했습니다: {e}")
