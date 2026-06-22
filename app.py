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

# 💾 1. 영구 기억 저장소 파일 경로 설정
HISTORY_FILE = "chat_history.json"
SCHEDULE_FILE = "schedule_history.json"
PRODUCTION_FILE = "production_history.json"  # 매일 업로드되는 실적 데이터 누적 저장소

def load_json(file_path):
    if os.path.exists(file_path):
        with open(file_path, "r", encoding="utf-8") as f:
            return json.load(f)
    return []

def save_json(file_path, data):
    with open(file_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)

# ⏰ 2. 일정 관리 알람 스케줄러 활성화
if "scheduler" not in st.session_state:
    st.session_state.scheduler = BackgroundScheduler()
    st.session_state.scheduler.start()

# ⚙️ 3. 페이지 레이아웃 및 환경 설정
st.set_page_config(page_title="주인님의 전용 비서", page_icon="🤖", layout="wide")
st.title("🤖 나만의 특급 비서 에이전트")
st.caption("대화, 일정 관리, 그리고 누적 실적 및 품질 PPM 자동 모니터링 시스템")

# 🔑 4. API 키 유효성 확인
try:
    my_key = st.secrets["OPENAI_API_KEY"]
except:
    my_key = None

if not my_key:
    st.error("🚨 설정(Secrets)에서 API 키를 입력해 주십시오.")
    st.stop()

client = OpenAI(api_key=my_key)

# 🧠 5. 가상 컴퓨터 메모리에 기존 영구 데이터 동기화
if "messages" not in st.session_state:
    st.session_state.messages = load_json(HISTORY_FILE)
if "schedules" not in st.session_state:
    st.session_state.schedules = load_json(SCHEDULE_FILE)
if "production_history" not in st.session_state:
    st.session_state.production_history = load_json(PRODUCTION_FILE)

# 📁 6. 왼쪽 사이드바 (실적 및 데이터 업로드 창)
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

    if uploaded_file:
        try:
            # 사진(Vision) 파일 처리
            if uploaded_file.type.startswith("image"):
                image_data = base64.b64encode(uploaded_file.read()).decode('utf-8')
                mime_type = uploaded_file.type
                st.image(uploaded_file, caption="분석 대기 중인 사진")
                st.success("✅ 사진 인식 완료!")
            
            # 생산 실적 파일 분석 및 누적 기억 로직
            elif uploaded_file.name.endswith('.csv') or uploaded_file.name.endswith('.xlsx'):
                df_upload = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
                file_content = df_upload.to_string()
                
                # 유연한 열 이름 매핑 엔진 (Line* 명 와일드카드 처리)
                line_col, target_col, passed_col, defect_col = None, None, None, None
                for col in df_upload.columns:
                    col_str = str(col).lower().strip()
                    if any(k in col_str for k in ['라인', 'line', '구분', '공정', '라인명']):
                        line_col = col
                    elif any(k in col_str for k in ['목표', 'target', '목표수량', '목표량']):
                        target_col = col
                    elif any(k in col_str for k in ['합격', 'passed', '합격수량', '양품', '양품수량', '합격량']):
                        passed_col = col
                    elif any(k in col_str for k in ['불량', 'defect', '불량수량', '불량량']):
                        defect_col = col

                if line_col and target_col and passed_col and defect_col:
                    today_str = datetime.date.today().strftime("%Y-%m-%d")
                    
                    # 업로드된 데이터를 표준 규격으로 변환하여 영구 데이터베이스에 누적
                    for _, row in df_upload.iterrows():
                        st.session_state.production_history.append({
                            "Date": today_str,
                            "Line": str(row[line_col]),
                            "Target": float(row[target_col]),
                            "Passed": float(row[passed_col]),
                            "Defect": float(row[defect_col])
                        })
                    save_json(PRODUCTION_FILE, st.session_state.production_history)
                    st.success("✅ 오늘자 생산 실적이 누적 기억 저장소에 안전하게 기록되었습니다!")
                else:
                    st.warning("⚠️ 엑셀 파일 내 주요 열 제목(라인명, 목표수량, 양품수량, 불량수량)을 확인해 주십시오.")
            
            # 일반 문서 처리
            elif uploaded_file.name.endswith('.pdf'):
                file_content = "\n".join([p.extract_text() for p in PyPDF2.PdfReader(uploaded_file).pages])
                st.success("✅ PDF 문서 학습 완료!")
            elif uploaded_file.name.endswith('.pptx'):
                prs = Presentation(uploaded_file)
                file_content = "\n".join([shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")])
                st.success("✅ PPTX 문서 학습 완료!")
            else:
                file_content = uploaded_file.read().decode("utf-8")
                st.success("✅ 텍스트 파일 학습 완료!")
        except Exception as e:
            st.error(f"⚠️ 파일 읽기 실패: {e}")

    st.markdown("---")
    st.link_button("🎨 Canva (캔바) 바로가기", "https://www.canva.com/")

# 📱 7. 상단 탭 구성 (대시보드 이름 변경 반영)
tab1, tab2, tab3 = st.tabs(["💬 비서와의 대화", "📅 일정 관리표", "📊 실적 현황"])

# ==========================================
# 탭 1: 비서와의 대화 (기존 기능 완전 유지)
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
                        reply = "그림을 완성했습니다! 이미지 보관 후 캔바를 통해 편집해 보십시오."
                        st.markdown(reply)
                        st.session_state.messages.append({"role": "assistant", "content": reply})
                        save_json(HISTORY_FILE, st.session_state.messages)
                    except Exception as e:
                        st.error(f"⚠️ 그림 에러: {e}")
        else:
            with st.chat_message("assistant"):
                with st.spinner("보좌관 분석 중..."):
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
# 탭 2: 일정 관리표 (기존 기능 완전 유지)
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
                    datetime_str = f"{new_date} {new_time.strftime('%H('%M')')}"
                    st.session_state.schedules.append({"Task": new_task, "DateTime": datetime_str})
                    save_json(SCHEDULE_FILE, st.session_state.schedules)
                    st.success(f"✅ '{new_task}' 일정 등록 완료")
                    st.rerun()

    if st.session_state.schedules:
        df_schedule = pd.DataFrame(st.session_state.schedules)
        st.dataframe(df_schedule.sort_values(by="DateTime").reset_index(drop=True), use_container_width=True)

# ==========================================
# 탭 3: 실적 현황 (★요청사항 전면 반영 및 고도화)
# ==========================================
with tab3:
    st.subheader("📊 월별 누적 실적 현황 관리 체계")
    
    if st.session_state.production_history:
        df_hist = pd.DataFrame(st.session_state.production_history)
        df_hist['Date'] = pd.to_datetime(df_hist['Date'])
        df_hist['Month'] = df_hist['Date'].dt.strftime('%Y-%m')
        
        # 1. 자동 한 달 단위 분기 기능 ("한달이 지나면 다시 시작하면서")
        current_month = datetime.date.today().strftime('%Y-%m')
        
        st.markdown(f"### 📅 당월 ({current_month}) 동일 라인별 실적 합계")
        df_current = df_hist[df_hist['Month'] == current_month]
        
        if not df_current.empty:
            # 동일 라인명으로 그룹화하여 합계 계산
            df_grouped = df_current.groupby('Line').agg({
                'Target': 'sum',
                'Passed': 'sum',
                'Defect': 'sum'
            }).reset_index()
            
            # 지정 수식 정밀 연산 (가동율, 이용율, PPM)
            df_grouped['가동율'] = df_grouped.apply(lambda r: r['Target'] / r['Passed'] if r['Passed'] > 0 else 0, axis=1)
            df_grouped['이용율'] = df_grouped.apply(lambda r: r['Passed'] / r['Target'] if r['Target'] > 0 else 0, axis=1)
            df_grouped['PPM'] = df_grouped.apply(lambda r: (r['Defect'] / r['Passed']) * 1000000 if r['Passed'] > 0 else 0, axis=1)
            
            # 주인님께서 지정하신 명칭으로 정렬 및 변경
            df_display = df_grouped.rename(columns={
                'Line': 'Line* 명',
                'Target': '목표수량',
                'Passed': '양품수량',
                'Defect': '불량수량'
            })
            
            # 요약 지표 테이블 시각화
            st.dataframe(df_display.style.format({
                '목표수량': '{:,.0f}',
                '양품수량': '{:,.0f}',
                '불량수량': '{:,.0f}',
                '가동율': '{:.1%}',
                '이용율': '{:.1%}',
                'PPM': '{:,.0f} PPM'
            }), use_container_width=True)
            
            # 그래프 모니터링 기능
            st.markdown("### 📈 라인별 핵심 지표 모니터링")
            g_col1, g_col2 = st.columns(2)
            with g_col1:
                st.markdown("**📦 생산 수량 균형 (목표 vs 양품)**")
                st.bar_chart(df_display.set_index('Line* 명')[['목표수량', '양품수량']])
            with g_col2:
                st.markdown("**🚨 불량 상태 트렌드 (PPM)**")
                st.bar_chart(df_display.set_index('Line* 명')[['PPM']])
        else:
            st.info("이번 달에 누적된 실적 데이터가 없습니다. 새로운 생산 실적 파일을 업로드해 주십시오.")
            
        # 2. 장기 월별 모니터링 뷰 (과거 월 데이터 추이 추적)
        st.markdown("---")
        st.markdown("### 📉 월별 장기 누적 모니터링 추이")
        df_monthly_trend = df_hist.groupby('Month').agg({'Target': 'sum', 'Passed': 'sum', 'Defect': 'sum'}).reset_index()
        st.line_chart(df_monthly_trend.set_index('Month')[['Target', 'Passed']])
    else:
        st.info("👈 왼쪽 파일 업로드 창에 실적 데이터(Excel/CSV)를 투입하시면 영구 실적 관리 시스템이 즉시 가동됩니다.")
