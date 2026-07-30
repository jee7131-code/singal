import streamlit as st
import pandas as pd
import openpyxl

st.set_page_config(page_title="2026 여름신앙학교 스케줄", page_icon="📅", layout="centered")
st.title("📅 여름신앙학교 종합 안내")

FILE_PATH = '2026 여름신앙학교 데일리 스케줄(최종).xlsx'

def load_schedule_data():
    df = pd.read_excel(FILE_PATH, sheet_name='타임테이블')
    return df

def load_excel_unmerged(file_path, sheet_name):
    wb = openpyxl.load_workbook(file_path, data_only=True)
    if sheet_name not in wb.sheetnames:
        return None
    ws = wb[sheet_name]
    
    merged_info = []
    for rng in list(ws.merged_cells.ranges):
        top_left_value = ws.cell(row=rng.min_row, column=rng.min_col).value
        merged_info.append((rng, top_left_value))
        
    for rng, _ in merged_info:
        ws.unmerge_cells(str(rng))
        
    for rng, val in merged_info:
        for r in range(rng.min_row, rng.max_row + 1):
            for c in range(rng.min_col, rng.max_col + 1):
                ws.cell(row=r, column=c).value = val
                
    data = list(ws.values)
    if not data:
        return None
    return pd.DataFrame(data)

def load_contacts_data():
    teacher_df = None
    student_df = None
    
    try:
        teacher_df = pd.read_excel(FILE_PATH, sheet_name='교사 비상연락망')
        if teacher_df is not None and not teacher_df.empty:
            teacher_df.columns = teacher_df.columns.str.strip()
    except Exception:
        pass
        
    try:
        student_df = pd.read_excel(FILE_PATH, sheet_name='학생 비상연락망')
        if student_df is not None and not student_df.empty:
            student_df.columns = student_df.columns.str.strip()
            sort_cols = []
            if '학년' in student_df.columns: sort_cols.append('학년')
            if '이름' in student_df.columns: sort_cols.append('이름')
            if sort_cols:
                student_df = student_df.sort_values(by=sort_cols, ascending=True).reset_index(drop=True)
    except Exception:
        pass
        
    return teacher_df, student_df

try:
    df_raw = load_excel_unmerged(FILE_PATH, '타임테이블')
    
    if df_raw is None:
        st.error("엑셀 파일에서 '타임테이블' 시트를 찾을 수 없습니다.")
    else:
        header_row_idx = 2
        for idx, row in df_raw.iterrows():
            if str(row.iloc[0]).strip() == 'TIME':
                header_row_idx = idx
                break
                
        header_row = df_raw.iloc[header_row_idx]
        people = [str(val).strip() for val in header_row.iloc[1:].dropna().values if str(val).strip()]
        df_body = df_raw.iloc[header_row_idx + 1:].copy()
        
        main_tab1, main_tab2 = st.tabs(["📅 신앙학교 스케줄", "📞 비상연락망"])
        
        # =========================================================
        # TAB 1: 신앙학교 스케줄
        # =========================================================
        with main_tab1:
            st.subheader("🗓️ 일정표 확인")
            schedule_options = ["🌟 전체 스케줄 보기", "🕒 특정 시간대 스케줄 보기", "🖨️ 인쇄용 스케줄 (A4 최적화)"] + people
            selected_schedule = st.selectbox("👀 확인할 스케줄 항목을 선택하세요:", schedule_options, key="schedule_select")
            
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
        # TAB 2: 비상연락망 (실제 번호 직접 노출 적용)
        # =========================================================
        with main_tab2:
            st.subheader("📞 비상연락망 안내")
            teacher_df, student_df = load_contacts_data()
            
            contact_tab1, contact_tab2 = st.tabs(["👨‍🏫 교사 연락망", "🧑‍🎓 학생 연락망"])
            
            with contact_tab1:
                if teacher_df is not None and not teacher_df.empty:
                    for idx, row in teacher_df.iterrows():
                        role = str(row.get('직함', '') if pd.notna(row.get('직함')) else '')
                        name = str(row.get('이름', '') if pd.notna(row.get('이름')) else '')
                        b_name = str(row.get('세례명', '') if pd.notna(row.get('세례명')) else '')
                        phone = str(row.get('전화번호', '') if pd.notna(row.get('전화번호')) else '')
                        
                        with st.container(border=True):
                            c1, c2 = st.columns([1.5, 1.5])
                            with c1:
                                st.markdown(f"**{role}** | **{name}** ({b_name})")
                            with c2:
                                if phone and phone.lower() != 'nan':
                                    clean_p = phone.replace('-', '').strip()
                                    # 🛠️ '전화 걸기' 문구 대신 실제 전화번호(phone)를 노출
                                    st.markdown(f"📞 [{phone}](tel:{clean_p})")
                                else:
                                    st.caption("번호 없음")
                else:
                    st.info("교사 비상연락망 데이터가 없습니다.")
                    
            with contact_tab2:
                if student_df is not None and not student_df.empty:
                    search_q = st.text_input("🔍 학생 이름 또는 세례명 검색:", "", key="student_search_card")
                    
                    filtered_s = student_df.copy()
                    if search_q:
                        mask = filtered_s.astype(str).apply(lambda row: row.str.contains(search_q, case=False).any(), axis=1)
                        filtered_s = filtered_s[mask]
                    
                    view_mode = st.radio("📱 보기 모드:", ["📱 카드형 (모바일 추천)", "📊 표 전체보기"], horizontal=True)
                    
                    if view_mode == "📊 표 전체보기":
                        st.dataframe(filtered_s, hide_index=True, use_container_width=True)
                    else:
                        st.caption(f"총 {len(filtered_s)}명 | 학년별 접이식 메뉴를 눌러보세요.")
                        
                        if '학년' in filtered_s.columns:
                            grades = filtered_s['학년'].unique()
                            for grade in grades:
                                grade_data = filtered_s[filtered_s['학년'] == grade]
                                with st.expander(f"🎓 **{grade}** ({len(grade_data)}명)", expanded=True):
                                    for idx, s_row in grade_data.iterrows():
                                        s_name = str(s_row.get('이름', ''))
                                        s_bname = str(s_row.get('세례명', ''))
                                        s_gender = str(s_row.get('성별', ''))
                                        s_phone = str(s_row.get('학생 연락처', ''))
                                        p_phone = str(s_row.get('학부모 연락처', ''))
                                        
                                        with st.container(border=True):
                                            st.markdown(f"👤 **{s_name}** ({s_bname}) · {s_gender}")
                                            
                                            col_s, col_p = st.columns(2)
                                            with col_s:
                                                if s_phone and s_phone.lower() != 'nan':
                                                    clean_s = s_phone.replace('-', '').strip()
                                                    st.markdown(f"📱 **학생:** [{s_phone}](tel:{clean_s})")
                                                else:
                                                    st.caption("📱 학생: 번호 없음")
                                                    
                                            with col_p:
                                                if p_phone and p_phone.lower() != 'nan':
                                                    clean_p = p_phone.replace('-', '').strip()
                                                    st.markdown(f"👨‍👩‍👦 **부모님:** [{p_phone}](tel:{clean_p})")
                                                else:
                                                    st.caption("👨‍👩‍👦 부모님: 번호 없음")
                        else:
                            for idx, s_row in filtered_s.iterrows():
                                s_name = str(s_row.get('이름', ''))
                                s_bname = str(s_row.get('세례명', ''))
                                s_phone = str(s_row.get('학생 연락처', ''))
                                p_phone = str(s_row.get('학부모 연락처', ''))
                                
                                with st.container(border=True):
                                    st.markdown(f"👤 **{s_name}** ({s_bname})")
                                    col_s, col_p = st.columns(2)
                                    with col_s:
                                        if s_phone and s_phone.lower() != 'nan':
                                            st.markdown(f"📱 [{s_phone}](tel:{s_phone.replace('-', '').strip()})")
                                    with col_p:
                                        if p_phone and p_phone.lower() != 'nan':
                                            st.markdown(f"👨‍👩‍👦 [{p_phone}](tel:{p_phone.replace('-', '').strip()})")
                else:
                    st.info("학생 비상연락망 데이터가 없습니다.")

except Exception as e:
    st.error(f"오류가 발생했습니다: {e}")
