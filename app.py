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

# 💾 1. 영구 기억 저장소 파일 경로 및 데이터 구조 정의
HISTORY_FILE = "chat_history.json"
SCHEDULE_FILE = "schedule_history.json"
PRODUCTION_FILE = "production_history.json"
MANUAL_KNOWLEDGE_FILE = "manual_knowledge.json"
MANUAL_CHAT_FILE = "manual_chat_history.json"

def load_json(file_path, default_type=dict):
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if default_type == dict and isinstance(data, list):
                    # 과거 리스트 형태의 데이터 구조 마이그레이션 방어 코드
                    return {"current_id": "default", "sessions": [{"id": "default", "title": "기본 대화방", "messages": data}]}
                return data
        except:
            return default_type()
    return default_type() if default_type == dict else []

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

# ⏰ 2. 일정 관리 알람 스케줄러 활성화
if "scheduler" not in st.session_state:
    st.session_state.scheduler = BackgroundScheduler()
    st.session_state.scheduler.start()

# ⚙️ 3. 페이지 기본 인프라 설정
st.set_page_config(page_title="주인님의 전용 비서", page_icon="🤖", layout="wide")
st.title("🤖 나만의 특급 비서 에이전트")
st.caption("멀티 채팅 세션, 실적 현황, 다국어 업무 매뉴얼 챗봇 및 코어워크-Canva 리포트 솔루션 통합 플랫폼")

# 🔑 4. API 키 검증 및 세팅
try:
    my_key = st.secrets["OPENAI_API_KEY"]
except:
    my_key = None

if not my_key:
    st.error("🚨 설정(Secrets)에서 OPENAI_API_KEY를 입력해 주십시오.")
    st.stop()

client = OpenAI(api_key=my_key)

# 🧠 5. 마스터 데이터 세션 동기화 (영구 백업 구조)
if "chats" not in st.session_state:
    st.session_state.chats = load_json(HISTORY_FILE, default_type=dict)
    if not st.session_state.chats or "sessions" not in st.session_state.chats:
        st.session_state.chats = {"current_id": "default", "sessions": [{"id": "default", "title": "기본 대화방", "messages": []}]}

if "schedules" not in st.session_state:
    st.session_state.schedules = load_json(SCHEDULE_FILE, default_type=list)
if "production_history" not in st.session_state:
    st.session_state.production_history = load_json(PRODUCTION_FILE, default_type=list)
if "manual_knowledge" not in st.session_state:
    st.session_state.manual_knowledge = load_json(MANUAL_KNOWLEDGE_FILE, default_type=list)
if "manual_chat" not in st.session_state:
    st.session_state.manual_chat = load_json(MANUAL_CHAT_FILE, default_type=list)

# 📁 6. 왼쪽 사이드바 구성 (멀티 채팅방 생성, 검색 및 데이터 업로드)
with st.sidebar:
    st.header("💬 대화 세션 관리")
    
    # ➕ 새 채팅방 생성 기능
    if st.button("➕ 새 대화 시작하기", use_container_width=True):
        new_id = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
        st.session_state.chats["sessions"].append({
            "id": new_id,
            "title": f"새 대화 {datetime.datetime.now().strftime('%m-%d %H:%M')}",
            "messages": []
        })
        st.session_state.chats["current_id"] = new_id
        save_json(HISTORY_FILE, st.session_state.chats)
        st.rerun()

    # 🔍 채팅 키워드 검색 엔진
    search_query = st.text_input("🔍 기존 대화 검색 (내용/제목)", "").strip()
    
    valid_sessions = []
    for sess in st.session_state.chats["sessions"]:
        if not search_query:
            valid_sessions.append(sess)
        else:
            title_match = search_query.lower() in sess["title"].lower()
            msg_match = any(search_query.lower() in str(m.get("content", "")).lower() for m in sess.get("messages", []))
            if title_match or msg_match:
                valid_sessions.append(sess)

    if valid_sessions:
        session_dict = {s["id"]: s["title"] for s in valid_sessions}
        curr_id = st.session_state.chats["current_id"]
        if curr_id not in session_dict:
            curr_id = valid_sessions[0]["id"]
            
        selected_session_id = st.selectbox(
            "이동할 대화방 선택",
            list(session_dict.keys()),
            format_func=lambda x: session_dict[x],
            index=list(session_dict.keys()).index(curr_id)
        )
        st.session_state.chats["current_id"] = selected_session_id
    else:
        st.caption("❌ 검색 조건에 부합하는 대화방이 없습니다.")

    st.markdown("---")
    st.header("📁 데이터 및 실적 업로드")
    
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
                
                workplace_col, target_col, passed_col, defect_col = None, None, None, None
                for col in df_upload.columns:
                    col_str = str(col).lower().strip()
                    if any(k in col_str for k in ['작업장', '작업장명', 'workplace', '라인', 'line', '구분', '공정']):
                        if workplace_col is None or '작업장' in col_str:
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
                        st.session_state.production_history.append({
                            "Date": today_str,
                            "Workplace": wp_name,
                            "Target": safe_float(row[target_col]),
                            "Passed": safe_float(row[passed_col]),
                            "Defect": safe_float(row[defect_col])
                        })
                    save_json(PRODUCTION_FILE, st.session_state.production_history)
                    st.success("✅ 생산 실적이 일자별 누적 DB에 안전하게 기록되었습니다!")
                else:
                    st.warning("⚠️ 열 제목(작업장명, 목표수량, 양품수량, 불량수량)을 확인하십시오.")
            
            elif uploaded_file.name.endswith('.pdf'):
                file_content = "\n".join([p.extract_text() for p in PyPDF2.PdfReader(uploaded_file).pages])
                st.success("✅ PDF 학습 성공!")
            elif uploaded_file.name.endswith('.pptx'):
                prs = Presentation(uploaded_file)
                file_content = "\n".join([shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")])
                st.success("✅ PPTX 학습 성공!")
            else:
                file_content = uploaded_file.read().decode("utf-8")
                st.success("✅ 문서 학습 성공!")
        except Exception as e:
            st.error(f"⚠️ 파일 처리 실패: {e}")

    st.markdown("---")
    st.link_button("🎨 Canva 스튜디오 직행", "https://www.canva.com/", use_container_width=True)

# 📱 7. 기능 분할 탭 마스터 구조 구현
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "💬 비서와의 대화", "📅 일정 관리표", "📊 실적 현황", "📚 업무 매뉴얼 챗봇", "💼 코어워크 & Canva"
])

current_session = next(s for s in st.session_state.chats["sessions"] if s["id"] == st.session_state.chats["current_id"])

# ==========================================
# 탭 1: 비서와의 대화 (구분 기록 및 멀티 세션 완벽 적용)
# ==========================================
with tab1:
    st.subheader(f"💬 현재 대화방: {current_session['title']}")
    
    with st.expander("📝 대화방 이름 변경"):
        new_title = st.text_input("새 대화방 이름 입력", current_session["title"])
        if st.button("이름 저장"):
            current_session["title"] = new_title
            save_json(HISTORY_FILE, st.session_state.chats)
            st.rerun()

    for msg in current_session["messages"]:
        if msg["role"] != "system":
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

    if prompt := st.chat_input("명령을 하명하십시오... (그림 요청: '/그리기 내용')", key="main_chat_input"):
        current_session["messages"].append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)

        if prompt.startswith("/그리기"):
            with st.chat_message("assistant"):
                with st.spinner("🎨 초고화질 일러스트 생성 중..."):
                    try:
                        draw_prompt = prompt.replace("/그리기", "").strip()
                        response = client.images.generate(model="dall-e-3", prompt=draw_prompt, size="1024x1024")
                        st.image(response.data[0].url, caption=f"완성작: {draw_prompt}")
                        reply = f"요청하신 '{draw_prompt}' 일러스트 작성을 마쳤습니다! 저장 후 Canva 탭에서 보고서에 활용해 보십시오."
                        st.markdown(reply)
                        current_session["messages"].append({"role": "assistant", "content": reply})
                        save_json(HISTORY_FILE, st.session_state.chats)
                    except Exception as e:
                        st.error(f"⚠️ DALL-E 3 에러: {e}")
        else:
            with st.chat_message("assistant"):
                with st.spinner("맥락 파악 및 심층 분석 중..."):
                    system_instruction = "당신은 주인님의 충직한 개인 비서입니다. 과거 대화 맥락과 업로드 데이터를 바탕으로 유능하게 보좌하세요."
                    messages_for_api = [{"role": "system", "content": system_instruction}] + current_session["messages"][:-1]
                    
                    if image_data:
                        messages_for_api.append({"role": "user", "content": [
                            {"type": "text", "text": prompt},
                            {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{image_data}"}}
                        ]})
                    elif file_content:
                        messages_for_api.append({"role": "user", "content": f"{prompt}\n\n[참고 데이터]:\n{file_content[:5000]}"})
                    else:
                        messages_for_api.append({"role": "user", "content": prompt})

                    try:
                        response = client.chat.completions.create(model="gpt-4o", messages=messages_for_api, temperature=0.7)
                        reply = response.choices[0].message.content
                        st.markdown(reply)
                        current_session["messages"].append({"role": "assistant", "content": reply})
                        save_json(HISTORY_FILE, st.session_state.chats)
                    except Exception as e:
                        st.error(f"⚠️ OpenAI API 통신 장애: {e}")

# ==========================================
# 탭 2: 일정 관리표
# ==========================================
with tab2:
    st.subheader("📅 주인님의 스케줄 관리 기구")
    with st.expander("➕ 새 일정 스케줄 등록", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            new_date = st.date_input("날짜 선택")
            new_time = st.time_input("시간 선택")
        with col2:
            new_task = st.text_input("스케줄 내용 명시")
            if st.button("스케줄 영구 저장"):
                if new_task:
                    datetime_str = f"{new_date} {new_time.strftime('%H:%M')}"
                    st.session_state.schedules.append({"Task": new_task, "DateTime": datetime_str})
                    save_json(SCHEDULE_FILE, st.session_state.schedules)
                    st.success("✅ 새로운 일정이 캘린더 데이터베이스에 기록되었습니다.")
                    st.rerun()

    if st.session_state.schedules:
        df_schedule = pd.DataFrame(st.session_state.schedules)
        st.dataframe(df_schedule.sort_values(by="DateTime").reset_index(drop=True), use_container_width=True)

# ==========================================
# 탭 3: 실적 현황
# ==========================================
with tab3:
    st.subheader("📊 실적 현황 (작업장 단위 고도화)")
    if st.session_state.production_history:
        df_hist = pd.DataFrame(st.session_state.production_history)
        if 'Workplace' not in df_hist.columns and 'Line' in df_hist.columns:
            df_hist['Workplace'] = df_hist['Line']
            
        df_hist['Date_dt'] = pd.to_datetime(df_hist['Date'])
        df_hist['Month'] = df_hist['Date_dt'].dt.strftime('%Y-%m')
        
        current_month = datetime.date.today().strftime('%Y-%m')
        st.markdown(f"### 📅 당월 ({current_month}) 작업장별 종합 정산")
        
        df_current = df_hist[df_hist['Month'] == current_month]
        if not df_current.empty:
            df_grouped = df_current.groupby('Workplace').agg({'Target': 'sum', 'Passed': 'sum', 'Defect': 'sum'}).reset_index()
            df_grouped['가동율'] = df_grouped.apply(lambda r: r['Passed'] / r['Target'] if r['Target'] > 0 else 0.0, axis=1)
            df_grouped['PPM'] = df_grouped.apply(lambda r: (r['Defect'] / r['Passed']) * 1000000 if r['Passed'] > 0 else 0.0, axis=1)
            
            df_display = df_grouped.rename(columns={'Workplace': '작업장명', 'Target': '목표수량', 'Passed': '양품수량', 'Defect': '불량수량'})
            st.dataframe(df_display.style.format({'목표수량': '{:,.0f}', '양품수량': '{:,.0f}', '불량수량': '{:,.0f}', '가동율': '{:.1%}', 'PPM': '{:,.0f} PPM'}), use_container_width=True)
            
            st.markdown("### 🚨 실시간 핵심 지표 위험 필터 모니터링")
            g_col1, g_col2 = st.columns(2)
            with g_col1:
                st.markdown("**📉 가동율 위험군 (80% 이하 작업장)**")
                df_under_80 = df_display[df_display['가동율'] <= 0.8]
                if not df_under_80.empty:
                    st.bar_chart(df_under_80.set_index('작업장명')[['가동율']])
                else:
                    st.success("🎉 모든 작업장 가동율 80% 초과 안정권")
            with g_col2:
                st.markdown("**🔥 PPM 불량 경보군 (500 PPM 이상 작업장)**")
                df_over_500 = df_display[df_display['PPM'] >= 500]
                if not df_over_500.empty:
                    st.bar_chart(df_over_500.set_index('작업장명')[['PPM']])
                else:
                    st.success("🎉 모든 작업장 불량 상태 500 PPM 미만 클린존")
        else:
            st.info("당월 누적 실적이 없습니다. 왼쪽 창에 실적 엑셀을 투입해 주십시오.")
            
        st.markdown("---")
        st.markdown("### 📈 일자별 생산 지표 누적 추이 그래프")
        df_daily = df_hist.groupby(df_hist['Date_dt'].dt.strftime('%Y-%m-%d')).agg({'Target': 'sum', 'Passed': 'sum'}).reset_index().sort_values('Date_dt')
        df_daily['누적 목표수량'] = df_daily['Target'].cumsum()
        df_daily['누적 양품수량'] = df_daily['Passed'].cumsum()
        st.line_chart(df_daily.set_index('Date_dt')[['누적 목표수량', '누적 양품수량']])
    else:
        st.info("업로드된 실적 데이터가 존재하지 않습니다.")

# ==========================================
# 탭 4: 📚 업무 매뉴얼 챗봇 (★베트남어 전문 번역 인격 탑재★)
# ==========================================
with tab4:
    st.subheader("📚 현장 업무 매뉴얼 전용 챗봇 (베트남어 자동 번역 지원)")
    with st.expander("🛠️ 매뉴얼 영구 학습 세션 구축"):
        manual_file = st.file_uploader("학습용 문서 주입 (PDF, TXT, PPTX)", type=["pdf", "txt", "pptx"], key="manual_tab_uploader")
        if st.button("🧠 비서 인공지능 뇌 훈련 실행"):
            if manual_file:
                with st.spinner("지식 각인 프로세스 가동 중..."):
                    try:
                        ext_text = ""
                        if manual_file.name.endswith('.pdf'):
                            ext_text = "\n".join([p.extract_text() for p in PyPDF2.PdfReader(manual_file).pages])
                        elif manual_file.name.endswith('.pptx'):
                            prs = Presentation(manual_file)
                            ext_text = "\n".join([shape.text for slide in prs.slides for shape in slide.shapes if hasattr(shape, "text")])
                        else:
                            ext_text = manual_file.read().decode("utf-8")
                        
                        st.session_state.manual_knowledge.append({"title": manual_file.name, "content": ext_text[:20000]})
                        save_json(MANUAL_KNOWLEDGE_FILE, st.session_state.manual_knowledge)
                        st.success(f"✅ '{manual_file.name}' 매뉴얼 지식화 완수!")
                        st.rerun()
                    except Exception as e:
                        st.error(f"⚠️ 지식 추출 에러: {e}")

    if st.session_state.manual_knowledge:
        st.info(f"📚 마스터한 규정: {', '.join([d['title'] for d in st.session_state.manual_knowledge])}")
    
    for msg in st.session_state.manual_chat:
        if msg["role"] != "system":
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

    if m_prompt := st.chat_input("매뉴얼 관련 질문 또는 '이 내용을 베트남어로 번역해줘' 입력...", key="manual_chat_input"):
        st.session_state.manual_chat.append({"role": "user", "content": m_prompt})
        with st.chat_message("user"):
            st.markdown(m_prompt)

        with st.chat_message("assistant"):
            with st.spinner("매뉴얼 데이터베이스 스캔 및 번역 중..."):
                all_context = "\n\n--- 참조 매뉴얼 원본 데이터 ---\n"
                for doc in st.session_state.manual_knowledge:
                    all_context += f"\n[문서: {doc['title']}]\n{doc['content']}\n"
                
                # 💡 핵심 수정: 매뉴얼 전문가이자 '베트남어 번역가' 인격을 강하게 부여
                m_instruction = (
                    "당신은 제조 현장의 완벽한 업무 매뉴얼 전문가이자 전문 베트남어 번역가입니다. "
                    "기본적인 질문에 대한 답변은 반드시 주어진 참조 매뉴얼 내용에만 근거하여 팩트 기반으로 작성하세요. "
                    "단, 사용자가 특정 내용이나 매뉴얼 규정을 '베트남어로 번역'해달라고 요청하는 경우, "
                    "매뉴얼의 내용을 벗어난 추측이라 판단하지 말고, 즉시 현지 베트남 작업자가 완벽히 이해할 수 있는 "
                    "가장 자연스럽고 정확한 베트남어로 번역해서 출력해 주십시오."
                ) + all_context
                
                api_m = [{"role": "system", "content": m_instruction}] + st.session_state.manual_chat[:-1] + [{"role": "user", "content": m_prompt}]
                
                try:
                    res = client.chat.completions.create(model="gpt-4o", messages=api_m, temperature=0.3)
                    reply = res.choices[0].message.content
                    st.markdown(reply)
                    st.session_state.manual_chat.append({"role": "assistant", "content": reply})
                    save_json(MANUAL_CHAT_FILE, st.session_state.manual_chat)
                except Exception as e:
                    st.error(f"⚠️ 매뉴얼 엔진 에러: {e}")

# ==========================================
# 탭 5: 💼 코어워크 & Canva 보고서
# ==========================================
with tab5:
    st.subheader("💼 생산 엔지니어링 코어워크 및 Canva 자동화")
    st.caption("제조 현장의 핵심 핵심 업무(Core Work)를 조율하고, Canva에 복사하여 즉시 제출할 수 있는 최고급 보고서 초안을 설계합니다.")
    
    core_menu = st.radio("실행할 코어워크 메뉴를 선택하십시오.", [
        "🎨 Canva 연동형 고품질 보고서 초안 생성기",
        "📋 핵심 기계 장치(열처리/컴프레서) 안전 점검",
        "✍️ 일일 근무 교대(Shift) 인계 보고서 작성"
    ], horizontal=True)
    
    st.markdown("---")
    
    # 1) Canva 활용 보고서 자동 생성 메뉴
    if core_menu == "🎨 Canva 연동형 고품질 보고서 초안 생성기":
        st.markdown("### 🎨 Canva 파워포인트/보고서 템플릿 맞춤형 텍스트 설계 가이드")
        st.info("원하시는 보고서 주제나 대시보드 데이터를 기반으로 Canva 레이아웃 박스에 최적화된 블록형 초안을 빌드합니다.")
        
        report_topic = st.text_input("보고서 주제 입력 (예: 당월 가동율 및 PPM 종합 분석 보고서)", "생산성 및 품질 지표 종합 분석 보고서")
        report_data_source = st.checkbox("현재 당월 누적 실적 데이터를 보고서에 자동 반영", value=True)
        
        if st.button("🚀 Canva 맞춤형 고급 초안 생성"):
            with st.spinner("Canva 프리젠테이션 레이아웃에 맞춤형 구조화 중..."):
                data_summary_str = ""
                if report_data_source and st.session_state.production_history:
                    df_h = pd.DataFrame(st.session_state.production_history)
                    data_summary_str = "\n[현재 시스템 누적 원본 데이터 요약]\n" + df_h.tail(10).to_string()
                
                canva_prompt = f"Canva 슬라이드나 보고서 문서 서식에 바로 카피앤페이스트하여 붙여넣을 수 있도록 슬라이드별/구역별 타이틀, 핵심 키워드 요약, 레이아웃 추천안을 한국어로 아름답게 짜줘. 주제: {report_topic} {data_summary_str}"
                
                try:
                    res = client.chat.completions.create(
                        model="gpt-4o",
                        messages=[
                            {"role": "system", "content": "당신은 대기업 제조공장의 최고 전문 기획관이자 Canva 디자인 레이아웃 컨설턴트입니다. 사용자가 Canva 박스에 바로 붙여넣을 수 있게 구획화된 구조적 텍스트를 제공하세요."},
                            {"role": "user", "content": canva_prompt}
                        ],
                        temperature=0.7
                    )
                    report_result = res.choices[0].message.content
                    st.success("✨ Canva 보고서 초안 설계 완료! 아래 텍스트 블록을 그대로 Canva 템플릿 텍스트 박스에 활용하십시오.")
                    st.text_area("📋 Canva 복사용 텍스트 공간", value=report_result, height=450)
                except Exception as e:
                    st.error(f"⚠️ 보고서 빌더 통신 에러: {e}")

    # 2) 핵심 기계 장치 안전 점검 코어워크 메뉴
    elif core_menu == "📋 핵심 기계 장치(열처리/컴프레서) 안전 점검":
        st.markdown("### 📋 현장 핵심 설비 체크리스트 운영")
        st.checkbox("1. 열처리 공정로 내부 압력 및 히터 인입 전류 측정 (정상)")
        st.checkbox("2. 메인 컴프레서 오일 레벨 및 토출 온도 점검 (정상)")
        st.checkbox("3. 자동화 가공 라인 기어 로봇 암 오정렬(Defect) 시각 감지 테스트")
        if st.button("점검 완료 로그 기록 및 비서에게 보고"):
            st.toast("✅ 설비 체크리스트가 마스터 로그에 정상 인계되었습니다.")

    # 3) 근무 교대 인계 보고서 코어워크 메뉴
    elif core_menu == "✍️ 일일 근무 교대(Shift) 인계 보고서 작성":
        st.markdown("### ✍️ Shift Handovers 공정 인계서 자동 포맷터")
        current_shift = st.selectbox("현재 근무조 선택", ["주간조 (Day Shift)", "야간조 (Night Shift)"])
        issues = st.text_area("인계 및 특이 전달사항 명시 (예: A작업장 설비 씰 교체로 가동율 일시 하락)", "특이사항 없음. 정상 생산 진행 완료.")
        
        if st.button("인계 양식 자동 빌딩"):
            formatted_handover = f"【근무 인계 보고서】\n- 일시: {datetime.date.today().strftime('%Y-%m-%d')}\n- 담당 근무조: {current_shift}\n- 공정 핵심 이슈: {issues}\n- 품질 가동 지표: 상단 대시보드 연동 완료."
            st.code(formatted_handover, language="text")
            st.success("위 포맷을 복사하여 내부 메신저나 보고 채널에 업로드하십시오.")
