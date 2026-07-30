import streamlit as st
import pandas as pd

# 모바일 화면에 맞춘 기본 설정
st.set_page_config(page_title="2026 여름신앙학교 스케줄", page_icon="📅", layout="centered")
st.title("📅 여름신앙학교 종합 안내")

FILE_PATH = '2026 여름신앙학교 데일리 스케줄(최종)_3.xlsx'

# 엑셀 데이터 로드 함수 (실시간 반영을 위해 캐시 미사용)
def load_schedule_data():
    df = pd.read_excel(FILE_PATH, sheet_name='타임테이블')
    return df

def load_contacts_data():
    teacher_df = None
    student_df = None
    
    try:
        teacher_df = pd.read_excel(FILE_PATH, sheet_name='교사 비상연락망')
    except Exception:
        pass
        
    try:
        student_df = pd.read_excel(FILE_PATH, sheet_name='학생 비상연락망')
        # 학년별 -> 이름별 오름차순 정렬
        if student_df is not None and not student_df.empty:
            sort_cols = []
            if '학년' in student_df.columns:
                sort_cols.append('학년')
            if '이름' in student_df.columns:
                sort_cols.append('이름')
            if sort_cols:
                student_df = student_df.sort_values(by=sort_cols, ascending=True).reset_index(drop=True)
    except Exception:
        pass
        
    return teacher_df, student_df

try:
    df_raw = load_schedule_data()
    
    # Header 위치 탐색 ('TIME' 행 찾아내기)
    header_row_idx = 2
    for idx, row in df_raw.iterrows():
        if str(row.iloc[0]).strip() == 'TIME':
            header_row_idx = idx
            break
            
    # 인물 목록 추출
    header_row = df_raw.iloc[header_row_idx]
    people = [str(val).strip() for val in header_row.iloc[1:].dropna().values if str(val).strip()]
    
    # 세로 및 가로 병합 셀(NaN) 자동 채우기
    df_body = df_raw.iloc[header_row_idx + 1:].copy()
    df_body.iloc[:, 0] = df_body.iloc[:, 0].ffill() 
    df_body.iloc[:, 1:] = df_body.iloc[:, 1:].ffill(axis=1)
    
    # ---------------------------------------------------------
    # 🌟 최상단 메인 탭 분리 (스케줄 vs 비상연락망)
    # ---------------------------------------------------------
    main_tab1, main_tab2 = st.tabs(["📅 신앙학교 스케줄", "📞 비상연락망"])
    
    # =========================================================
    # TAB 1: 신앙학교 스케줄
    # =========================================================
    with main_tab1:
        st.subheader("🗓️ 일정표 확인")
        
        # 스케줄 메뉴 전용 드롭다운
        schedule_options = ["🌟 전체 스케줄 보기", "🕒 특정 시간대 스케줄 보기", "🖨️ 인쇄용 스케줄 (A4 최적화)"] + people
        selected_schedule = st.selectbox("👀 확인할 스케줄 항목을 선택하세요:", schedule_options, key="schedule_select")
        
        # 1-1. 전체 스케줄 보기
        if selected_schedule == "🌟 전체 스케줄 보기":
            st.caption("👈 모바일에서는 표를 좌우로 스크롤하여 모든 인물을 확인할 수 있습니다.")
            all_data = []
            current_day = "DAY1"
            
            for idx in range(len(df_body)):
                row = df_body.iloc[idx]
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
                day_tabs = st.tabs(list(days))
                for i, day in enumerate(days):
                    with day_tabs[i]:
                        day_df = df_all[df_all["DAY"] == day].drop(columns=["DAY"])
                        st.dataframe(day_df, hide_index=True, use_container_width=False)
                        
        # 1-2. 특정 시간대 스케줄 보기
        elif selected_schedule == "🕒 특정 시간대 스케줄 보기":
            time_data = {}
            current_day = "DAY1"
            
            for idx in range(len(df_body)):
                row = df_body.iloc[idx]
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
                selected_day = st.selectbox("📅 날짜(DAY)를 선택하세요:", days, key="time_day_select")
                day_times = [item["time"] for item in time_data[selected_day]]
                selected_time = st.selectbox("⏰ 시간을 선택하세요:", day_times, key="time_hour_select")
                
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

        # 1-3. 인쇄용 스케줄 (다운로드 방식)
        elif selected_schedule == "🖨️ 인쇄용 스케줄 (A4 최적화)":
            st.success("웹사이트 화면 제약 없이 여러 장을 완벽하게 인쇄하려면 아래의 **[다운로드]** 버튼을 눌러 파일을 받아주세요!")
            print_target = st.selectbox("출력할 대상:", ["전체 스케줄"] + people, key="print_target_select")
            
            all_data = []
            current_day = "DAY1"
            
            for idx in range(len(df_body)):
                row = df_body.iloc[idx]
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
            
            if print_target != "전체 스케줄":
                df_print = df_all[["DAY", "시간", print_target]]
                df_print = df_print[df_print[print_target] != ""] 
            else:
                df_print = df_all
                
            days = df_print["DAY"].unique()
            
            html_content = """
            <!DOCTYPE html>
            <html>
            <head>
            <meta charset="utf-8">
            <title>스케줄 인쇄</title>
            <style>
                body { font-family: 'Malgun Gothic', sans-serif; padding: 20px; }
                table { width: 100%; border-collapse: collapse; margin-bottom: 40px; font-size: 11pt; }
                th, td { border: 1px solid #000; padding: 10px; text-align: center; }
                th { background-color: #e6e6e6; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                tr { page-break-inside: avoid; }
                thead { display: table-header-group; }
                h2 { page-break-before: always; margin-bottom: 10px; }
                h2:first-of-type { page-break-before: auto; }
                @page { size: auto; margin: 15mm; }
            </style>
            </head>
            <body>
            """
            
            for day in days:
                day_df = df_print[df_print["DAY"] == day].drop(columns=["DAY"])
                html_table = day_df.to_html(index=False, escape=False)
                html_content += f"<h2>{day} - {print_target}</h2>\n"
                html_content += html_table + "\n"
                
            html_content += "</body></html>"
            
            st.download_button(
                label="📥 완벽 인쇄용 파일 다운로드 (클릭)",
                data=html_content,
                file_name=f"여름신앙학교_스케줄_{print_target}.html",
                mime="text/html"
            )

        # 1-4. 개인별 스케줄 보기
        else:
            selected_person = selected_schedule
            person_col_idx = list(header_row).index(selected_person)
            schedule_data = []
            current_day = "DAY1"
            
            for idx in range(len(df_body)):
                row = df_body.iloc[idx]
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
                day_tabs = st.tabs(list(days))
                for i, day in enumerate(days):
                    with day_tabs[i]:
                        day_df = df_schedule[df_schedule["DAY"] == day]
                        st.dataframe(day_df[["시간", "담당 업무"]], hide_index=True, use_container_width=True)

    # =========================================================
    # TAB 2: 비상연락망
    # =========================================================
    with main_tab2:
        st.subheader("📞 비상연락망 안내")
        teacher_df, student_df = load_contacts_data()
        
        contact_tab1, contact_tab2 = st.tabs(["👨‍🏫 교사 비상연락망", "🧑‍🎓 학생 비상연락망"])
        
        with contact_tab1:
            if teacher_df is not None and not teacher_df.empty:
                st.dataframe(teacher_df, hide_index=True, use_container_width=True)
            else:
                st.info("교사 비상연락망 데이터가 없거나 시트명을 확인해 주세요.")
                
        with contact_tab2:
            if student_df is not None and not student_df.empty:
                st.caption("📌 학년별 ➔ 이름별 오름차순으로 자동 정렬되었습니다.")
                search_query = st.text_input("🔍 학생 이름 또는 조 검색:", "", key="student_search")
                filtered_df = student_df.copy()
                if search_query:
                    mask = filtered_df.astype(str).apply(lambda row: row.str.contains(search_query, case=False).any(), axis=1)
                    filtered_df = filtered_df[mask]
                    
                st.caption(f"총 {len(filtered_df)}건의 데이터가 조회되었습니다.")
                st.dataframe(filtered_df, hide_index=True, use_container_width=True)
            else:
                st.info("학생 비상연락망 데이터가 없거나 시트명을 확인해 주세요.")

except Exception as e:
    st.error(f"오류가 발생했습니다: {e}")
