import streamlit as st
from openai import OpenAI
import PyPDF2
import pandas as pd
from pptx import Presentation
import base64

# 스마트폰 및 웹 화면 최적화 설정
st.set_page_config(page_title="주인님의 전용 비서", page_icon="🤖", layout="centered")
st.title("🤖 나만의 특급 비서 에이전트")
st.caption("사진 인식(Vision)과 그림 그리기 기능이 탑재되었습니다.")

# 비밀 금고에서 열쇠 가져오기
try:
    my_key = st.secrets["OPENAI_API_KEY"]
except:
    my_key = None

if not my_key:
    st.error("🚨 비밀 금고(Secrets) 연결을 확인해 주세요.")
else:
    client = OpenAI(api_key=my_key)

    if "messages" not in st.session_state:
        st.session_state.messages = []

    # 📁 왼쪽 사이드바: 파일/사진 업로드 및 캔바 바로가기
    with st.sidebar:
        st.header("📁 첨부 및 편의 기능")
        uploaded_file = st.file_uploader("파일이나 캡처 사진을 올려주십시오.", type=["pdf", "csv", "txt", "pptx", "png", "jpg", "jpeg"])
        
        st.markdown("---")
        st.markdown("🎨 **디자인 스튜디오**")
        st.link_button("Canva (캔바) 바로가기", "https://www.canva.com/")
        st.caption("비서가 만든 기획안이나 그림을 캔바에서 멋지게 꾸며보세요!")

        file_text = ""
        image_base64 = None
        mime_type = None

        if uploaded_file is not None:
            # 1. 사진/캡처 파일(이미지) 처리 로직
            if uploaded_file.name.lower().endswith(('.png', '.jpg', '.jpeg')):
                st.image(uploaded_file, caption="업로드된 사진", use_column_width=True)
                image_base64 = base64.b64encode(uploaded_file.getvalue()).decode('utf-8')
                mime_type = "image/png" if uploaded_file.name.lower().endswith('.png') else "image/jpeg"
                st.success("✅ 사진 인식 준비 완료! 무엇을 분석할까요?")
            # 2. 일반 문서 파일 처리 로직
            else:
                try:
                    if uploaded_file.name.endswith('.pdf'):
                        pdf_reader = PyPDF2.PdfReader(uploaded_file)
                        for page in pdf_reader.pages:
                            file_text += page.extract_text() + "\n"
                    elif uploaded_file.name.endswith('.csv'):
                        df = pd.read_csv(uploaded_file)
                        file_text = df.to_string()
                    elif uploaded_file.name.endswith('.txt'):
                        file_text = uploaded_file.getvalue().decode("utf-8")
                    elif uploaded_file.name.endswith('.pptx'):
                        prs = Presentation(uploaded_file)
                        for slide in prs.slides:
                            for shape in slide.shapes:
                                if hasattr(shape, "text"):
                                    file_text += shape.text + "\n"
                    st.success(f"✅ 문서 해독 완료!")
                except Exception as e:
                    st.error(f"⚠️ 파일 읽기 에러: {e}")

    # 화면에 이전 대화 내용 표시
    for msg in st.session_state.messages:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    # 주인님의 명령 입력창
    if prompt := st.chat_input("명령을 내려주십시오, 주인님... (그림을 원하시면 '/그리기 내용' 입력)"):
        # ▶ 기능 A: 그림 그리기 명령을 내렸을 때
        if prompt.startswith("/그리기"):
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)
            
            with st.chat_message("assistant"):
                with st.spinner("🎨 주인님의 명령에 따라 그림을 생성하고 있습니다..."):
                    try:
                        draw_prompt = prompt.replace("/그리기", "").strip()
                        response = client.images.generate(
                            model="dall-e-3",
                            prompt=draw_prompt,
                            size="1024x1024",
                            quality="standard",
                            n=1,
                        )
                        image_url = response.data[0].url
                        st.image(image_url, caption=f"완성된 그림: {draw_prompt}")
                        reply = f"요청하신 '{draw_prompt}' 이미지를 완성했습니다! 이미지를 꾹 눌러 저장하신 뒤 좌측 캔바(Canva) 버튼을 눌러 디자인에 활용해 보십시오."
                        st.markdown(reply)
                        st.session_state.messages.append({"role": "assistant", "content": reply})
                    except Exception as e:
                        st.error(f"⚠️ 그림 그리기 에러: {e}")
                        
        # ▶ 기능 B: 일반 대화 및 사진 분석 명령을 내렸을 때
        else:
            st.session_state.messages.append({"role": "user", "content": prompt})
            with st.chat_message("user"):
                st.markdown(prompt)

            with st.chat_message("assistant"):
                with st.spinner("분석 중입니다..."):
                    try:
                        system_instruction = "당신은 주인님의 유능하고 충직한 개인 비서 에이전트입니다. 항상 정중하게 답변하세요."
                        if file_text:
                            system_instruction += f"\n\n[참고 문서]\n{file_text[:5000]}"
                        
                        api_messages = [{"role": "system", "content": system_instruction}]
                        
                        # 이전 대화 내용 불러오기
                        for msg in st.session_state.messages[:-1]:
                            api_messages.append(msg)
                            
                        # 업로드된 사진이 있다면 비서의 눈(Vision)으로 전송
                        if image_base64:
                            api_messages.append({
                                "role": "user",
                                "content": [
                                    {"type": "text", "text": prompt},
                                    {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{image_base64}"}}
                                ]
                            })
                        else:
                            api_messages.append({"role": "user", "content": prompt})

                        response = client.chat.completions.create(
                            model="gpt-4o",  
                            messages=api_messages,
                            temperature=0.7
                        )
                        reply = response.choices[0].message.content
                        st.markdown(reply)
                        st.session_state.messages.append({"role": "assistant", "content": reply})
                    except Exception as e:
                        st.error(f"⚠️ 에러 발생: {e}")
