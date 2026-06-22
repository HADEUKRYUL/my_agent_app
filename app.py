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

# 에러 방지용 숫자 변환 함수
def safe_float(val):
    try:
        return float(str(val).replace(',', '').strip())
    except:
        return 0.0

# ⏰ 2. 일정 관리 알람
if "scheduler" not in st.session_state:
    st.session_state.scheduler = BackgroundScheduler()
    st.session_state.scheduler.start()

# ⚙️ 3. 페이지 기본 설정
st.set_page_config(page_title="주인님의 전용 비서", page_icon="🤖", layout="wide")
st.title("🤖 나만의 특급 비서 에이전트")
st.caption("대화, 일정 관리, 누적 실적 및 위험 라인(가동율 80%↓, 500 PPM↑) 자동 감지 시스템")

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
            # 사진(Vision) 처리
            if uploaded_file.type.startswith("image"):
                image_data = base64.b64encode(uploaded_file.read()).decode('utf-8')
                mime_type = uploaded_file.type
                st.image(uploaded_file, caption="분석 대기 사진")
                st.success("✅ 사진 인식 완료!")
            
            # 생산 실적 파일 분석 및 '일자별' 누적 저장
            elif uploaded_file.name.endswith('.csv') or uploaded_file.name.endswith('.xlsx'):
                df_upload = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
                file_content = df_upload.to_string()
                
                line_col, target_col, passed_col, defect_col = None, None, None, None
                for col in df_upload.columns:
                    col_str = str(col).lower().strip()
                    if any(k in col_str for k in ['라인', 'line', '구분', '공정', '라인명']):
                        line_col = col
                    elif any(k in col_str for k in ['목표', 'target', '목표수량', '목표량']):
                        target_col = col
                    elif any(k in col_str for k in ['양품', 'passed', '합격', '양품수량', '합격수량']):
                        passed_col = col
                    elif any(k in col_str for k in ['불량', 'defect', '불량수량', '불량량']):
                        defect_col = col

                if line_col and target_col and passed_col and defect_col:
                    today_str = datetime.date.today().strftime("%Y-%m-%d")
                    
                    # 충돌 방지: JSON에는 핵심 수량 3가지만 깨끗하게 저장합니다.
                    for _, row in df_upload.iterrows():
                        line_name = str(row[line_col]).strip()
                        if line_name == 'nan' or not line_name:
                            continue
                            
                        st.session_state.production_history.append({
                            "Date": today_str,
                            "Line": line_name,
                            "Target": safe_float(row[target_col]),
                            "Passed": safe_float(row[passed_col]),
                            "Defect": safe_float(row[defect_col])
                        })
                    save_json(PRODUCTION_FILE, st.session_state.production_history)
                    st.success("✅ 오늘자 생산 실적이 [일자별 누적 데이터베이스]에 기록되었습니다!")
                else:
                    st.warning("⚠️ 엑셀 열 제목(라인명, 목표수량, 양품수량, 불량수량)을 확인해 주십시오.")
            
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
                    system_instruction = "당신은 주인님의 충직한 개인 비서입니다. 과거 대화와 업로드 데이터를 바탕으로 완벽히 보좌하세요."
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
# 탭 3: 실적 현황 (결측치 완벽 차단 및 평균 연산 적용)
# ==========================================
with tab3:
    st.subheader("📊 실적 현황 (누적 및 위험 모니터링)")
    
    if st.session_state.production_history:
        df_hist = pd.DataFrame(st.session_state.production_history)
        
        # 만약 과거 데이터에 누락된 항목이 있다면 0으로 안전하게 채움
        for col in ['Target', 'Passed', 'Defect']:
            if col not in df_hist.columns:
                df_hist[col] = 0.0
                
        df_hist['Date_dt'] = pd.to_datetime(df_hist['Date'])
        df_hist['Month'] = df_hist['Date_dt'].dt.strftime('%Y-%m')
        
        # [핵심 보완] 로우(Row) 단위로 가동율, 이용율을 미리 계산 (나중에 평균내기 위함)
        df_hist['OpRate'] = df_hist.apply(lambda r: r['Target'] / r['Passed'] if r['Passed'] > 0 else 0, axis=1)
        df_hist['UtilRate'] = df_hist.apply(lambda r: r['Passed'] / r['Target'] if r['Target'] > 0 else 0, axis=1)
        
        current_month = datetime.date.today().strftime('%Y-%m')
        st.markdown(f"### 📅 당월 ({current_month}) Line* 명 실적 종합")
        
        df_current = df_hist[df_hist['Month'] == current_month]
        
        if not df_current.empty:
            # 1. 수량은 합산(sum), 비율(가동율/이용율)은 평균(mean)으로 그룹화!
            df_grouped = df_current.groupby('Line').agg({
                'Target': 'sum',
                'Passed': 'sum',
                'Defect': 'sum',
                'OpRate': 'mean',    # 평균 연산
                'UtilRate': 'mean'   # 평균 연산
            }).reset_index()
            
            # 2. PPM 연산 (불량수량 / 양품수량 * 1,000,000)
            df_grouped['PPM'] = df_grouped.apply(lambda r: (r['Defect'] / r['Passed']) * 1000000 if r['Passed'] > 0 else 0, axis=1)
            
            # 3. 명칭 적용
            df_display = df_grouped.rename(columns={
                'Line': 'Line* 명',
                'Target': '목표수량',
                'Passed': '양품수량',
                'Defect': '불량수량',
                'OpRate': '가동율',
                'UtilRate': '이용율'
            })
            
            st.dataframe(df_display.style.format({
                '목표수량': '{:,.0f}',
                '양품수량': '{:,.0f}',
                '불량수량': '{:,.0f}',
                '가동율': '{:.1%}',
                '이용율': '{:.1%}',
                'PPM': '{:,.0f} PPM'
            }), use_container_width=True)
            
            # 4. 필터링 차트 (가동율 80% 이하, PPM 500 이상만 추출)
            st.markdown("### 🚨 라인별 핵심 지표 모니터링")
            g_col1, g_col2 = st.columns(2)
            
            with g_col1:
                st.markdown("**📉 가동율 80% 이하 라인**")
                df_under_80 = df_display[df_display['가동율'] <= 0.8]
                if not df_under_80.empty:
                    st.bar_chart(df_under_80.set_index('Line* 명')[['가동율']])
                else:
                    st.success("전체 라인 가동율 양호 (80% 초과)")
                    
            with g_col2:
                st.markdown("**🔥 PPM (500 이상 라인명 표출)**")
                df_over_500 = df_display[df_display['PPM'] >= 500]
                if not df_over_500.empty:
                    st.bar_chart(df_over_500.set_index('Line* 명')[['PPM']])
                else:
                    st.success("전체 라인 불량 상태 양호 (500 PPM 미만)")
        else:
            st.info("이번 달 실적 데이터가 아직 업로드되지 않았습니다.")
            
        # 5. 매일 누적되는 일자별 장기 트렌드 지표
        st.markdown("---")
        st.markdown("### 📈 일자별 누적 생산 지표 트렌드")
        df_daily = df_hist.groupby(df_hist['Date_dt'].dt.strftime('%Y-%m-%d')).agg({
            'Target': 'sum', 
            'Passed': 'sum'
        }).reset_index()
        
        # 일자별 누적(Cumulative Sum) 산출
        df_daily = df_daily.sort_values('Date_dt')
        df_daily['누적 목표수량'] = df_daily['Target'].cumsum()
        df_daily['누적 양품수량'] = df_daily['Passed'].cumsum()
        
        st.line_chart(df_daily.set_index('Date_dt')[['누적 목표수량', '누적 양품수량']])
        
    else:
        st.info("👈 왼쪽 파일 업로드 창에 오늘자 실적 데이터(Excel/CSV)를 투입해 주십시오.")
