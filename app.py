import streamlit as st
import pandas as pd

# 모바일 화면에 맞춘 기본 설정
st.set_page_config(page_title="2026 여름신앙학교 스케줄", page_icon="📅", layout="centered")
st.title("📅 여름신앙학교 스케줄 확인")

# 인쇄용 CSS 숨김 처리 및 A4 규격 세팅
st.markdown("""
<style>
@media print {
    /* 인쇄 시 웹사이트 기본 메뉴, 버튼, 헤더 등 불필요한 요소 숨김 */
    header, footer, .stSelectbox, .stAlert, [data-testid="stSidebar"] { 
        display: none !important; 
    }
    /* A4 용지에 맞춘 표 디자인 강제 적용 */
    table { 
        width: 100% !important; 
        border-collapse: collapse !important; 
        font-family: 'Malgun Gothic', sans-serif !important; 
        font-size: 11pt !important; 
    }
    th, td { 
        border: 1px solid #000 !important; 
        padding: 6px !important; 
        text-align: center !important; 
    }
    th { 
        background-color: #e6e6e6 !important; 
        -webkit-print-color-adjust: exact !important; 
        print-color-adjust: exact !important; 
    }
    /* 페이지 넘어갈 때 행이 반으로 잘리지 않도록 보호 */
    tr { page-break-inside: avoid !important; }
    h3 { page-break-before: always !important; }
    h3:first-of-type { page-break-before: auto !important; }
}
</style>
""", unsafe_allow_html=True)

# 엑셀 파일 로드
FILE_PATH = '2026 여름신앙학교 데일리 스케줄(최종).xlsx'

@st.cache_data
def load_data():
    return pd.read_excel(FILE_PATH, sheet_name='타임테이블')

try:
    df_raw = load_data()
    
    # Header 위치 탐색 ('TIME' 행 찾아내기)
    header_row_idx = 2
    for idx, row in df_raw.iterrows():
        if str(row.iloc[0]).strip() == 'TIME':
            header_row_idx = idx
            break
            
    # 인물 목록 추출
    header_row = df_raw.iloc[header_row_idx]
    people = [str(val).strip() for val in header_row.iloc[1:].dropna().values if str(val).strip()]
    
    # 보기 모드 리스트
    view_options = ["🌟 전체 스케줄 보기", "🕒 특정 시간대 스케줄 보기", "🖨️ 인쇄용 스케줄 (A4 최적화)"] + people
    selected_option = st.selectbox("👀 무엇을 확인하시겠어요?", view_options)
    
    # ---------------------------------------------------------
    # 1. 인쇄용 스케줄 (HTML 표 직접 렌더링)
    # ---------------------------------------------------------
    if selected_option == "🖨️ 인쇄용 스케줄 (A4 최적화)":
        st.subheader("🖨️ 인쇄용 화면 세팅")
        st.info("출력할 대상을 선택한 뒤, 인터넷 브라우저의 메뉴에서 **[인쇄]** 또는 **[PDF로 저장]**을 누르세요. 불필요한 메뉴가 사라지고 표만 깔끔하게 출력됩니다.")
        
        print_target = st.selectbox("출력할 대상:", ["전체 스케줄"] + people)
        
        all_data = []
        current_day = "DAY1"
        
        for idx in range(header_row_idx + 1, len(df_raw)):
            row = df_raw.iloc[idx]
            time_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            if "DAY" in time_val.upper():
                current_day = time_val
                continue
            if time_val.lower() in ["nan", "none"]: time_val = ""
            
            row_dict = {"DAY": current_day, "시간": time_val}
            has_task = False
            
            for person in people:
                p_idx = list(header_row).index(person)
                task_val = str(row.iloc[p_idx]).strip() if pd.notna(row.iloc[p_idx]) else ""
                if task_val.lower() in ["nan", "none"]: task_val = ""
                row_dict[person] = task_val
                if task_val: has_task = True
                
            if time_val or has_task:
                all_data.append(row_dict)
                
        df_all = pd.DataFrame(all_data)
        
        # 선택한 대상에 맞춰 데이터 정제
        if print_target != "전체 스케줄":
            df_print = df_all[["DAY", "시간", print_target]]
            df_print = df_print[df_print[print_target] != ""] # 공란 제거
        else:
            df_print = df_all
            
        days = df_print["DAY"].unique()
        for day in days:
            st.markdown(f"### {day} - {print_target}")
            day_df = df_print[df_print["DAY"] == day].drop(columns=["DAY"])
            # 인쇄에 가장 적합한 순수 HTML 테이블로 변환하여 출력
            html_table = day_df.to_html(index=False, escape=False)
            st.markdown(html_table, unsafe_allow_html=True)
            
    # ---------------------------------------------------------
    # 2. 전체 스케줄 보기
    # ---------------------------------------------------------
    elif selected_option == "🌟 전체 스케줄 보기":
        st.subheader("전체 시간대별 일정")
        st.caption("👈 모바일에서는 표를 좌우로 스크롤하여 모든 인물을 확인할 수 있습니다.")
        
        all_data = []
        current_day = "DAY1"
        
        for idx in range(header_row_idx + 1, len(df_raw)):
            row = df_raw.iloc[idx]
            time_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            if "DAY" in time_val.upper():
                current_day = time_val
                continue
            if time_val.lower() in ["nan", "none"]: time_val = ""
            
            row_dict = {"DAY": current_day, "시간": time_val}
            has_task = False
            for person in people:
                p_idx = list(header_row).index(person)
                task_val = str(row.iloc[p_idx]).strip() if pd.notna(row.iloc[p_idx]) else ""
                if task_val.lower() in ["nan", "none"]: task_val = ""
                row_dict[person] = task_val
                if task_val: has_task = True
                
            if time_val or has_task:
                all_data.append(row_dict)
                
        df_all = pd.DataFrame(all_data)
        days = df_all["DAY"].unique()
        if len(days) > 0:
            tabs = st.tabs(list(days))
            for i, day in enumerate(days):
                with tabs[i]:
                    day_df = df_all[df_all["DAY"] == day].drop(columns=["DAY"])
                    st.dataframe(day_df, hide_index=True, use_container_width=False)
                    
    # ---------------------------------------------------------
    # 3. 특정 시간대 스케줄 보기
    # ---------------------------------------------------------
    elif selected_option == "🕒 특정 시간대 스케줄 보기":
        st.subheader("🕒 시간대별 전체 업무 확인")
        time_data = {}
        current_day = "DAY1"
        
        for idx in range(header_row_idx + 1, len(df_raw)):
            row = df_raw.iloc[idx]
            time_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            if "DAY" in time_val.upper():
                current_day = time_val
                continue
            if time_val.lower() in ["nan", "none", ""]: 
                continue
                
            tasks = {}
            has_task = False
            for person in people:
                p_idx = list(header_row).index(person)
                task_val = str(row.iloc[p_idx]).strip() if pd.notna(row.iloc[p_idx]) else ""
                if task_val.lower() in ["nan", "none"]: task_val = ""
                tasks[person] = task_val
                if task_val: has_task = True
                
            if has_task:
                if current_day not in time_data:
                    time_data[current_day] = []
                time_data[current_day].append({"time": time_val, "tasks": tasks})
                
        days = list(time_data.keys())
        if days:
            selected_day = st.selectbox("📅 날짜(DAY)를 선택하세요:", days)
            day_times = [item["time"] for item in time_data[selected_day]]
            selected_time = st.selectbox("⏰ 시간을 선택하세요:", day_times)
            
            st.divider()
            st.markdown(f"### 📍 {selected_day} | {selected_time}")
            for item in time_data[selected_day]:
                if item["time"] == selected_time:
                    for person, task in item["tasks"].items():
                        if task:
                            st.write(f"- **{person}**: {task}")
                        else:
                            st.write(f"- **{person}**: (공란/휴식)")
                    break

    # ---------------------------------------------------------
    # 4. 개인별 스케줄 보기
    # ---------------------------------------------------------
    else:
        selected_person = selected_option
        person_col_idx = list(header_row).index(selected_person)
        schedule_data = []
        current_day = "DAY1"
        
        for idx in range(header_row_idx + 1, len(df_raw)):
            row = df_raw.iloc[idx]
            time_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            if "DAY" in time_val.upper():
                current_day = time_val
                continue
            task_val = str(row.iloc[person_col_idx]).strip() if pd.notna(row.iloc[person_col_idx]) else ""
            if task_val.lower() in ["nan", "none"]: task_val = ""
            if time_val.lower() in ["nan", "none"]: time_val = ""
            
            if time_val or task_val:
                schedule_data.append({"DAY": current_day, "시간": time_val, "담당 업무": task_val})
                
        df_schedule = pd.DataFrame(schedule_data)
        st.subheader(f"✅ {selected_person} 선생님 일정")
        days = df_schedule["DAY"].unique()
        if len(days) > 0:
            tabs = st.tabs(list(days))
            for i, day in enumerate(days):
                with tabs[i]:
                    day_df = df_schedule[df_schedule["DAY"] == day]
                    st.dataframe(day_df[["시간", "담당 업무"]], hide_index=True, use_container_width=True)

except Exception as e:
    st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
