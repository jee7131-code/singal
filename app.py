import streamlit as st
import pandas as pd

# 모바일 화면에 맞춘 기본 설정
st.set_page_config(page_title="2026 여름신앙학교 스케줄", page_icon="📅", layout="centered")
st.title("📅 여름신앙학교 스케줄 확인")

# 엑셀 파일 로드 (캐싱을 통해 속도 향상)
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
            
    # 인물 목록 추출 (B열부터 끝까지)
    header_row = df_raw.iloc[header_row_idx]
    people = [str(val).strip() for val in header_row.iloc[1:].dropna().values if str(val).strip()]
    
    st.info("💡 선생님 개인 일정 또는 전체 스케줄을 선택하여 확인할 수 있습니다.")
    
    # 보기 모드 리스트에 '전체 스케줄 보기' 추가
    view_options = ["🌟 전체 스케줄 보기"] + people
    
    # 모바일 터치용 선택 드롭다운
    selected_option = st.selectbox("👀 무엇을 확인하시겠어요?", view_options)
    
    if selected_option == "🌟 전체 스케줄 보기":
        st.subheader("전체 시간대별 일정")
        st.caption("👈 모바일에서는 표를 좌우로 스크롤하여 모든 인물을 확인할 수 있습니다.")
        
        all_data = []
        current_day = "DAY1"
        
        # 전체 데이터 파싱
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
                
            # 시간이나 업무 중 하나라도 있는 행만 추가
            if time_val or has_task:
                all_data.append(row_dict)
                
        df_all = pd.DataFrame(all_data)
        
        days = df_all["DAY"].unique()
        if len(days) > 0:
            tabs = st.tabs(list(days))
            for i, day in enumerate(days):
                with tabs[i]:
                    day_df = df_all[df_all["DAY"] == day].drop(columns=["DAY"])
                    # use_container_width=False로 두어 열이 많을 때 좌우 스크롤이 자연스럽게 되도록 함
                    st.dataframe(day_df, hide_index=True, use_container_width=False)
                    
    else:
        # 기존 인물별 보기 로직
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
                schedule_data.append({
                    "DAY": current_day,
                    "시간": time_val,
                    "담당 업무": task_val
                })
                
        df_schedule = pd.DataFrame(schedule_data)
        st.subheader(f"✅ {selected_person} 선생님 일정")
        
        days = df_schedule["DAY"].unique()
        if len(days) > 0:
            tabs = st.tabs(list(days))
            for i, day in enumerate(days):
                with tabs[i]:
                    day_df = df_schedule[df_schedule["DAY"] == day]
                    st.dataframe(
                        day_df[["시간", "담당 업무"]],
                        hide_index=True,
                        use_container_width=True
                    )

except Exception as e:
    st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
