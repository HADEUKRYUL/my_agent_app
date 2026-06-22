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

# 💾 1. 영구 기억 저장소 설정 (대화 및 일정)
HISTORY_FILE = "chat_history.json"
SCHEDULE_FILE = "schedule_history.json"

def load_json(file_path):
    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_json(file_path, data):
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

# ⏰ 2. 일정 관리 알람 (스케줄러) 활성화
if "scheduler" not in st.session_state:
    st.session_state.scheduler = BackgroundScheduler()
    st.session_state.scheduler.start()

# ⚙️ 3. 페이지 기본 설정
st.set_page_config(page_title="주인님의 전용 비서", page_icon="🤖", layout="wide")
st.title("🤖 나만의 특급 비서 에이전트")
st.caption("대화, 일정 관리, 그리고 실적·품질 지표(가동률/불량PPM) 자동 연산 대시보드로 보좌합니다.")

# 🔑 4. API 키 설정
try:
    my_key = st.secrets["OPENAI_API_KEY"]
except:
    my_key = None

if not my_key:
    st.error("🚨 설정(Secrets)에서 API 키를 입력해 주십시오.")
    st.stop()

client = OpenAI(api_key=my_key)

# 🧠 5. 데이터 로드 (창을 닫아도 기억 유지)
if "messages" not in st.session_state:
    st.session_state.messages = load_json(HISTORY_FILE)
if "schedules" not in st.session_state:
    st.session_state.schedules = load_json(SCHEDULE_FILE)

# 📁 6. 왼쪽 서랍(사이드바) - 파일 업로드 및 편의 기능
with st.sidebar:
    st.header("📁 데이터 및 실적 업로드")
    
    if st.button("🗑️ 대화 기록 초기화"):
        st.session_state.messages = []
        save_json(HISTORY_FILE, [])
        st.rerun()

    uploaded_file = st.file_uploader("문서, 사진, 또는 생산실적(CSV/Excel) 업로드", type=["pdf", "csv", "xlsx", "txt", "pptx", "png", "jpg", "jpeg"])
    
    file_content = None
    image_data = None
    mime_type = None
    df_dashboard = None 

    if uploaded_file:
        try:
            # 사진(Vision) 파일 처리
            if uploaded_file.type.startswith("image"):
                image_data = base64.b64encode(uploaded_file.read()).decode('utf-8')
                mime_type = uploaded_file.type
                st.image(uploaded_file, caption="분석 대기 중인 사진")
                st.success("✅ 사진 인식 완료!")
            
            # 생산 실적 데이터 처리
            elif uploaded_file.name.endswith('.csv') or uploaded_file.name.endswith('.xlsx'):
                if uploaded_file.name.endswith('.csv'):
                    df_dashboard = pd.read_csv(uploaded_file)
                else:
                    df_dashboard = pd.read_excel(uploaded_file)
                file_content = df_dashboard.to_string()
                st.success("✅ 생산 실적 데이터 요약 성공! [📊 생산 실적 대시보드] 탭을 확인하십시오.")
            
            # 일반 문서 처리
            elif uploaded_file.name.endswith('.pdf'):
                file_content = "\n".join([p.extract_text() for p in PyPDF2.PdfReader(uploaded_file).pages])
                st.success("✅ PDF 학습 완료!")
            elif uploaded_file.name.endswith('.pptx'):
                prs = Presentation(uploaded_file)
                file_content = "\n".join([shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")])
                st.success("✅ PPTX 학습 완료!")
            else:
                file_content = uploaded_file.read().decode("utf-8")
                st.success("✅ 문서 학습 완료!")
        except Exception as e:
            st.error(f"⚠️ 파일 읽기 실패: {e}")

    st.markdown("---")
    st.link_button("🎨 Canva (캔바) 바로가기", "https://www.canva.com/")

# 📱 7. 화면을 3개의 탭(Tab)으로 분리
tab1, tab2, tab3 = st.tabs(["💬 비서와의 대화", "📅 일정 관리표", "📊 생산 실적 대시보드"])

# ==========================================
# 탭 1: 비서와의 대화
# ==========================================
with tab1:
    for msg in st.session_state.messages:
        if msg["role"] != "system":
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

    if prompt := st.chat_input("명령을 내려주십시오. (그림 요청: '/그리기 내용')"):
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        # 그림 그리기
        if prompt.startswith("/그리기"):
            with st.chat_message("assistant"):
                with st.spinner("🎨 그림을 생성하고 있습니다..."):
                    try:
                        draw_prompt = prompt.replace("/그리기", "").strip()
                        response = client.images.generate(model="dall-e-3", prompt=draw_prompt, size="1024x1024", quality="standard", n=1)
                        st.image(response.data[0].url, caption=f"완성된 그림: {draw_prompt}")
                        reply = "요청하신 그림을 완성했습니다!"
                        st.markdown(reply)
                        st.session_state.messages.append({"role": "assistant", "content": reply})
                        save_json(HISTORY_FILE, st.session_state.messages)
                    except Exception as e:
                        st.error(f"⚠️ 그림 생성 에러: {e}")
        # 일반 대화 및 문서 분석
        else:
            with st.chat_message("assistant"):
                with st.spinner("분석 중입니다..."):
                    system_instruction = "당신은 주인님의 충직한 개인 비서입니다. 과거 대화와 업로드된 데이터를 바탕으로 완벽히 보좌하세요."
                    messages_for_api = [{"role": "system", "content": system_instruction}] + st.session_state.messages[:-1]
                    
                    if image_data:
                        messages_for_api.append({"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{image_data}"}}]})
                    elif file_content:
                        messages_for_api.append({"role": "user", "content": f"{prompt}\n\n[업로드된 데이터/문서]:\n{file_content[:5000]}"})
                    else:
                        messages_for_api.append({"role": "user", "content": prompt})

                    try:
                        response = client.chat.completions.create(model="gpt-4o", messages=messages_for_api, temperature=0.7)
                        reply = response.choices[0].message.content
                        st.markdown(reply)
                        st.session_state.messages.append({"role": "assistant", "content": reply})
                        save_json(HISTORY_FILE, st.session_state.messages)
                    except Exception as e:
                        st.error(f"⚠️ 통신 에러: {e}")

# ==========================================
# 탭 2: 일정 관리표
# ==========================================
with tab2:
    st.subheader("📅 주인님의 일정표")
    with st.expander("➕ 새 일정 등록하기", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            new_date = st.date_input("날짜 선택")
            new_time = st.time_input("시간 선택")
        with col2:
            new_task = st.text_input("일정 내용")
            if st.button("일정 저장"):
                if new_task:
                    datetime_str = f"{new_date} {new_time.strftime('%H:%M')}"
                    st.session_state.schedules.append({"Task": new_task, "DateTime": datetime_str})
                    save_json(SCHEDULE_FILE, st.session_state.schedules)
                    st.success(f"✅ '{new_task}' 일정이 등록되었습니다!")
                    st.rerun()
                else:
                    st.warning("일정 내용을 입력해 주십시오.")

    if st.session_state.schedules:
        df_schedule = pd.DataFrame(st.session_state.schedules)
        df_schedule = df_schedule.sort_values(by="DateTime").reset_index(drop=True)
        st.dataframe(df_schedule, use_container_width=True)
        if st.button("🗑️ 완료된 지난 일정 모두 지우기"):
            st.session_state.schedules = []
            save_json(SCHEDULE_FILE, [])
            st.rerun()
    else:
        st.info("현재 등록된 일정이 없습니다.")

# ==========================================
# 탭 3: 생산 실적 대시보드 (수량 및 가동률/PPM 연산 강화)
# ==========================================
with tab3:
    st.subheader("📊 라인별 실적 정밀 대시보드")
    
    if df_dashboard is not None:
        # 1. 스마트 열 매핑 엔진 (주인님의 요구사항 반영)
        line_col, target_col, passed_col, defect_col = None, None, None, None
        for col in df_dashboard.columns:
            col_str = str(col).lower().strip()
            if any(k in col_str for k in ['라인', 'line', '구분', '공정', '라인명']):
                line_col = col
            elif any(k in col_str for k in ['목표', 'target', '목표수량', '목표량']):
                target_col = col
            elif any(k in col_str for k in ['합격', 'passed', '합격수량', '양품', '합격량', '정상수량']):
                passed_col = col
            elif any(k in col_str for k in ['불량', 'defect', '불량수량', '불량량']):
                defect_col = col

        # 2. 필수 열 매핑 성공 시 계산 수행
        if line_col and target_col and passed_col and defect_col:
            st.success("🎯 엑셀 항목 매핑 완료 (라인명, 목표수량, 합격수량, 불량수량 자동 인지)")
            
            # 동일 라인별 수량 데이터 선합산
            df_grouped = df_dashboard.groupby(line_col).agg({
                target_col: 'sum',
                passed_col: 'sum',
                defect_col: 'sum'
            }).reset_index()
            
            # 주인님께서 지정하신 수식 적용 연산 (분모 0 예방 처리 포함)
            # 가동률 = 목표수량 / 합격수량
            df_grouped['가동률'] = df_grouped.apply(
                lambda row: row[target_col] / row[passed_col] if row[passed_col] > 0 else 0, axis=1
            )
            # 불량PPM = (불량수량 / 합격수량) * 1000000
            df_grouped['불량PPM'] = df_grouped.apply(
                lambda row: (row[defect_col] / row[passed_col]) * 1000000 if row[passed_col] > 0 else 0, axis=1
            )
            
            # 한글 칼럼 명칭 통일 및 정렬
            df_display = df_grouped.rename(columns={
                line_col: '라인명',
                target_col: '목표수량',
                passed_col: '합격수량',
                defect_col: '불량수량'
            })
            
            # 📑 요약 레포트 출력
            st.markdown("### 📑 라인별 실적 분석 보고서")
            st.dataframe(df_display.style.format({
                '목표수량': '{:,.0f}',
                '합격수량': '{:,.0f}',
                '불량수량': '{:,.0f}',
                '가동률': '{:.1%}',
                '불량PPM': '{:,.0f} PPM'
            }), use_container_width=True)
            
            # 📈 대형 시각화 그래프 배치
            st.markdown("### 📈 항목별 지표 트렌드")
            g_col1, g_col2, g_col3 = st.columns(3)
            
            with g_col1:
                st.markdown("**📦 수량 종합 비교 (목표 vs 합격 vs 불량)**")
                st.bar_chart(df_display.set_index('라인명')[['목표수량', '합격수량', '불량수량']])
            
            with g_col2:
                st.markdown("**⚡ 라인별 가동률 (목표 / 합격)**")
                st.line_chart(df_display.set_index('라인명')[['가동률']])
                
            with g_col3:
                st.markdown("**🚨 라인별 불량 지수 (PPM)**")
                st.bar_chart(df_display.set_index('라인명')[['불량PPM']])
                
            with st.expander("📂 업로드 원본 데이터 보기"):
                st.dataframe(df_dashboard, use_container_width=True)
        else:
            st.warning("⚠️ 엑셀 파일 내에서 필요한 항목(라인명, 목표수량, 합격수량, 불량수량)의 열 이름을 완벽히 찾지 못했습니다.")
            with st.expander("💡 팁: 엑셀 파일 상단의 열 제목을 다음과 유사하게 맞춰주시면 자동 계산됩니다."):
                st.markdown("- **라인명:** `라인명`, `라인` 또는 `Line` \n- **목표수량:** `목표수량`, `목표` 또는 `Target` \n- **합격수량:** `합격수량`, `합격` 또는 `Passed` \n- **불량수량:** `불량수량`, `불량` 또는 `Defect`")
            st.dataframe(df_dashboard, use_container_width=True)
    else:
        st.info("👈 왼쪽 파일 업로드 창에 목표 및 합격/불량 수량이 기록된 실적 엑셀 파일을 업로드해 주십시오.")
