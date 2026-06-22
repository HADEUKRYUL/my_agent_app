import streamlit as st
from openai import OpenAI
import PyPDF2
import pandas as pd
from pptx import Presentation
import base64
import json
import os
import datetime
import math
from apscheduler.schedulers.background import BackgroundScheduler

# 💾 1. 영구 기억 저장소
HISTORY_FILE = "chat_history.json"
SCHEDULE_FILE = "schedule_history.json"
PRODUCTION_FILE = "production_history.json"

def load_json(file_path):
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            return []
    return []

def save_json(file_path, data):
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

# 에러 방지용 안전 숫자 변환 함수
def safe_float(val):
    try:
        if pd.isna(val):
            return 0.0
        val_str = str(val).replace(',', '').strip()
        if not val_str or val_str.lower() in ['nan', 'nat', 'null']:
            return 0.0
        f = float(val_str)
        return 0.0 if math.isnan(f) else f
    except:
        return 0.0

# ⏰ 2. 일정 관리 알람
if "scheduler" not in st.session_state:
    st.session_state.scheduler = BackgroundScheduler()
    st.session_state.scheduler.start()

# ⚙️ 3. 페이지 기본 설정
st.set_page_config(page_title="주인님의 전용 비서", page_icon="🤖", layout="wide")
st.title("🤖 나만의 특급 비서 에이전트")
st.caption("대화, 일정, 누적 실적 및 위험 작업장(가동율 80%↓, 500 PPM↑) 자동 감지 시스템")

# 🔑 4. API 키 설정
try:
    my_key = st.secrets["OPENAI_API_KEY"]
except:
    my_key = None

if not my_key:
    st.error("🚨 설정(Secrets)에서 API 키를 입력해 주십시오.")
    st.stop()

client = OpenAI(api_key=my_key)

# 🧠 5. 데이터 동기화
if "messages" not in st.session_state:
    st.session_state.messages = load_json(HISTORY_FILE)
if "schedules" not in st.session_state:
    st.session_state.schedules = load_json(SCHEDULE_FILE)
if "production_history" not in st.session_state:
    st.session_state.production_history = load_json(PRODUCTION_FILE)

# 📁 6. 왼쪽 사이드바 (실적 업로드)
with st.sidebar:
    st.header("📁 데이터 및 실적 업로드")
    
    if st.button("🗑️ 대화 기록 초기화"):
        st.session_state.messages = []
        save_json(HISTORY_FILE, [])
        st.rerun()

    uploaded_file = st.file_uploader("문서, 사진, 생산실적(Excel/CSV) 업로드", type=["pdf", "csv", "xlsx", "txt", "pptx", "png", "jpg", "jpeg"])
    
    file_content = None
    image_data = None
    mime_type = None

    if uploaded_file:
        try:
            if uploaded_file.type.startswith("image"):
                image_data = base64.b64encode(uploaded_file.read()).decode('utf-8')
                mime_type = uploaded_file.type
                st.image(uploaded_file, caption="분석 대기 사진")
                st.success("✅ 사진 인식 완료!")
            
            elif uploaded_file.name.endswith('.csv') or uploaded_file.name.endswith('.xlsx'):
                df_upload = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
                file_content = df_upload.to_string()
                
                # [핵심] 작업장명 우선 매핑 엔진으로 변경
                workplace_col, target_col, passed_col, defect_col = None, None, None, None
                for col in df_upload.columns:
                    col_str = str(col).lower().strip()
                    if any(k in col_str for k in ['작업장', '작업장명', 'workplace', '라인', 'line', '구분', '공정']):
                        if workplace_col is None or '작업장' in col_str: # 작업장을 최우선으로 잡음
                            workplace_col = col
                    elif any(k in col_str for k in ['목표', 'target', '목표수량', '목표량']):
                        target_col = col
                    elif any(k in col_str for k in ['양품', 'passed', '합격', '양품수량', '합격수량']):
                        passed_col = col
                    elif any(k in col_str for k in ['불량', 'defect', '불량수량', '불량량']):
                        defect_col = col

                if workplace_col and target_col and passed_col and defect_col:
                    today_str = datetime.date.today().strftime("%Y-%m-%d")
                    
                    for _, row in df_upload.iterrows():
                        wp_name = str(row[workplace_col]).strip()
                        if wp_name.lower() in ['nan', 'nat', 'null', ''] or pd.isna(row[workplace_col]):
                            continue
                            
                        # 작업장(Workplace) 단위로 데이터 저장
                        st.session_state.production_history.append({
                            "Date": today_str,
                            "Workplace": wp_name,
                            "Target": safe_float(row[target_col]),
                            "Passed": safe_float(row[passed_col]),
                            "Defect": safe_float(row[defect_col])
                        })
                    save_json(PRODUCTION_FILE, st.session_state.production_history)
                    st.success("✅ 오늘자 생산 실적이 [일자별 누적 데이터베이스]에 기록되었습니다!")
                else:
                    st.warning("⚠️ 엑셀 열 제목(작업장명, 목표수량, 양품수량, 불량수량)을 확인해 주십시오.")
            
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

# 📱 7. 상단 탭 구성
tab1, tab2, tab3 = st.tabs(["💬 비서와의 대화", "📅 일정 관리표", "📊 실적 현황"])

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

        if prompt.startswith("/그리기"):
            with st.chat_message("assistant"):
                with st.spinner("🎨 그림 생성 중..."):
                    try:
                        draw_prompt = prompt.replace("/그리기", "").strip()
                        response = client.images.generate(model="dall-e-3", prompt=draw_prompt, size="1024x1024")
                        st.image(response.data[0].url, caption=f"완성본: {draw_prompt}")
                        reply = "그림을 완성했습니다!"
                        st.markdown(reply)
                        st.session_state.messages.append({"role": "assistant", "content": reply})
                        save_json(HISTORY_FILE, st.session_state.messages)
                    except Exception as e:
                        st.error(f"⚠️ 그림 에러: {e}")
        else:
            with st.chat_message("assistant"):
                with st.spinner("분석 중..."):
                    system_instruction = "당신은 주인님의 충직한 개인 비서입니다. 베트남 현장의 정보와 업로드 데이터를 바탕으로 완벽히 보좌하세요."
                    messages_for_api = [{"role": "system", "content": system_instruction}] + st.session_state.messages[:-1]
                    
                    if image_data:
                        messages_for_api.append({"role": "user", "content": [{"type": "text", "text": prompt}, {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{image_data}"}}]})
                    elif file_content:
                        messages_for_api.append({"role": "user", "content": f"{prompt}\n\n[업로드 데이터]:\n{file_content[:5000]}"})
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
                    st.success(f"✅ 일정이 등록되었습니다!")
                    st.rerun()

    if st.session_state.schedules:
        df_schedule = pd.DataFrame(st.session_state.schedules)
        st.dataframe(df_schedule.sort_values(by="DateTime").reset_index(drop=True), use_container_width=True)

# ==========================================
# 탭 3: 실적 현황 (작업장명 기준 집계 및 필터링)
# ==========================================
with tab3:
    st.subheader("📊 실적 현황 (누적 및 위험 감지 모니터링)")
    
    if st.session_state.production_history:
        df_hist = pd.DataFrame(st.session_state.production_history)
        
        # [호환성 방어] 과거 데이터에 'Line'만 있고 'Workplace'가 없을 경우를 대비한 자동 매핑
        if 'Workplace' not in df_hist.columns and 'Line' in df_hist.columns:
            df_hist['Workplace'] = df_hist['Line']
            
        for col in ['Target', 'Passed', 'Defect']:
            if col not in df_hist.columns:
                df_hist[col] = 0.0
                
        df_hist['Date_dt'] = pd.to_datetime(df_hist['Date'])
        df_hist['Month'] = df_hist['Date_dt'].dt.strftime('%Y-%m')
        
        current_month = datetime.date.today().strftime('%Y-%m')
        st.markdown(f"### 📅 당월 ({current_month}) 작업장별 실적 종합")
        
        df_current = df_hist[df_hist['Month'] == current_month]
        
        if not df_current.empty:
            # 1. 작업장(Workplace) 기준으로 수량 합산 
            df_grouped = df_current.groupby('Workplace').agg({
                'Target': 'sum',
                'Passed': 'sum',
                'Defect': 'sum'
            }).reset_index()
            
            # 2. 가동율 공식 적용 (양품수량 / 목표수량)
            df_grouped['가동율'] = df_grouped.apply(lambda r: r['Passed'] / r['Target'] if r['Target'] > 0 else 0.0, axis=1)
            
            # 3. PPM 연산 (불량수량 / 양품수량 * 1,000,000)
            df_grouped['PPM'] = df_grouped.apply(lambda r: (r['Defect'] / r['Passed']) * 1000000 if r['Passed'] > 0 else 0.0, axis=1)
            
            # 4. 화면 표시용 칼럼명 변경
            df_display = df_grouped.rename(columns={
                'Workplace': '작업장명',
                'Target': '목표수량',
                'Passed': '양품수량',
                'Defect': '불량수량'
            })
            
            st.dataframe(df_display.style.format({
                '목표수량': '{:,.0f}',
                '양품수량': '{:,.0f}',
                '불량수량': '{:,.0f}',
                '가동율': '{:.1%}',
                'PPM': '{:,.0f} PPM'
            }), use_container_width=True)
            
            # 5. 필터링 차트 (가동율 80% 이하, PPM 500 이상 작업장만 추출)
            st.markdown("### 🚨 핵심 지표 모니터링 (위험 작업장 색출)")
            g_col1, g_col2 = st.columns(2)
            
            with g_col1:
                st.markdown("**📉 가동율 80% 이하 작업장**")
                df_under_80 = df_display[df_display['가동율'] <= 0.8]
                if not df_under_80.empty:
                    st.bar_chart(df_under_80.set_index('작업장명')[['가동율']])
                else:
                    st.success("전체 작업장 가동율 양호 (80% 초과)")
                    
            with g_col2:
                st.markdown("**🔥 PPM** (500 PPM 이상 불량 작업장)")
                df_over_500 = df_display[df_display['PPM'] >= 500]
                if not df_over_500.empty:
                    st.bar_chart(df_over_500.set_index('작업장명')[['PPM']])
                else:
                    st.success("전체 작업장 불량 상태 양호 (500 PPM 미만)")
        else:
            st.info("이번 달 실적 데이터가 아직 업로드되지 않았습니다.")
            
        # 6. 매일 누적되는 일자별 장기 트렌드 지표
        st.markdown("---")
        st.markdown("### 📈 일자별 누적 생산 지표 트렌드")
        df_daily = df_hist.groupby(df_hist['Date_dt'].dt.strftime('%Y-%m-%d')).agg({
            'Target': 'sum', 
            'Passed': 'sum'
        }).reset_index()
        
        df_daily = df_daily.sort_values('Date_dt')
        df_daily['누적 목표수량'] = df_daily['Target'].cumsum()
        df_daily['누적 양품수량'] = df_daily['Passed'].cumsum()
        
        st.line_chart(df_daily.set_index('Date_dt')[['누적 목표수량', '누적 양품수량']])
        
    else:
        st.info("👈 왼쪽 파일 업로드 창에 실적 데이터(Excel/CSV)를 투입해 주십시오.")
