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
st.caption("대화, 일정 관리, 그리고 동일 라인별 수량·가동률 정밀 분석 대시보드로 보좌합니다.")

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
            
            # 생산 실적 데이터 처리 (대시보드 확장)
            elif uploaded_file.name.endswith('.csv') or uploaded_file.name.endswith('.xlsx'):
                if uploaded_file.name.endswith('.csv'):
                    df_dashboard = pd.read_csv(uploaded_file)
                else:
                    df_dashboard = pd.read_excel(uploaded_file)
                file_content = df_dashboard.to_string()
                st.success("✅ 생산 실적 데이터 분석 성공! 상단의 [📊 생산 실적 대시보드] 탭을 확인하십시오.")
            
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
# 탭 3: 생산 실적 대시보드 (수량 및 가동률 정밀 고도화)
# ==========================================
with tab3:
    st.subheader("📊 라인별 실적 정밀 대시보드")
    
    if df_dashboard is not None:
        # 1. 엑셀 파일 내 주요 키워드 자동 매핑 엔진
        line_col, qty_col, util_col = None, None, None
        for col in df_dashboard.columns:
            col_str = str(col).lower().strip()
            if any(k in col_str for k in ['라인', 'line', '구분', '공정']):
                line_col = col
            elif any(k in col_str for k in ['생산수량', '생산량', '수량', 'quantity', 'qty', '실적']):
                qty_col = col
            elif any(k in col_str for k in ['가동률', '가동율', 'utilization', 'rate', '가동', '가동비율']):
                util_col = col

        # 2. 매핑 성공 여부에 따른 스마트 대시보드 렌더링
        if line_col and (qty_col or util_col):
            st.success(f"🎯 동일 라인 식별 완료 (기준 열: {line_col})")
            
            # 동일 라인별 그룹화 연산 (수량은 '합계', 가동률은 '평균')
            agg_rules = {}
            if qty_col: agg_rules[qty_col] = 'sum'
            if util_col: agg_rules[util_col] = 'mean'
            
            df_grouped = df_dashboard.groupby(line_col).agg(agg_rules).reset_index()
            
            # 📊 스탯 요약 표 출력 (한눈에 보기)
            st.markdown("### 📑 동일 라인별 핵심 지표 요약표")
            # 가동률 데이터 형태(0~1 소수점 형태인지, 0~100 퍼센트 형태인지 자동 감지하여 포맷팅)
            display_format = {}
            if qty_col: display_format[qty_col] = "{:,.0f}"
            if util_col:
                is_percentage = df_grouped[util_col].max() > 1.0
                display_format[util_col] = "{:.1f}%" if is_percentage else "{:.1%}"
                
            st.dataframe(df_grouped.style.format(display_format), use_container_width=True)
            
            # 📉 시각화 영역 분할 배치 (수량과 가동률 스케일 분리)
            st.markdown("### 📈 라인별 추이 분석 그래프")
            chart_col1, chart_col2 = st.columns(2)
            
            with chart_col1:
                if qty_col:
                    st.markdown(f"**📦 라인별 총 생산수량 합계**")
                    st.bar_chart(df_grouped.set_index(line_col)[[qty_col]])
            
            with chart_col2:
                if util_col:
                    st.markdown(f"**⚡ 라인별 평균 가동률 추이**")
                    st.line_chart(df_grouped.set_index(line_col)[[util_col]])
                    
            with st.expander("📂 업로드 원본 데이터 전체 보기"):
                st.dataframe(df_dashboard, use_container_width=True)
        else:
            # 매핑 실패 시 범용 폴백 모드
            st.warning("⚠️ 엑셀 파일 내에서 '라인', '생산수량', '가동률'을 뜻하는 명확한 열 이름을 찾지 못했습니다. 일반 수치 자동 분석으로 전환합니다.")
            with st.expander("💡 대시보드 자동 연동을 위한 권장 엑셀 열 이름 팁"):
                st.markdown("- **라인 열 이름:** `라인` 또는 `Line` \n- **생산수량 열 이름:** `생산수량` 또는 `Quantity` \n- **가동률 열 이름:** `가동률` 또는 `Utilization`")
            
            numeric_cols = df_dashboard.select_dtypes(include=['float64', 'int64']).columns
            if len(numeric_cols) > 0:
                st.bar_chart(df_dashboard[numeric_cols])
            st.dataframe(df_dashboard, use_container_width=True)
    else:
        st.info("👈 왼쪽 업로드 창에 오늘자 생산 실적 파일(Excel / CSV)을 넣어주시면 동일 라인별 가동률 정밀 분석 대시보드가 활성화됩니다.")
