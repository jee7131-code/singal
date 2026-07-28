import streamlit as st
import pandas as pd

# 모바일 화면에 맞춘 기본 설정
st.set_page_config(page_title="2026 여름신앙학교 스케줄", page_icon="📅", layout="centered")
st.title("📅 인물별 스케줄 확인")

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
    
    st.info("💡 카카오톡 등에서 접속하여 본인의 이름만 선택하면 전체 일정이 정리됩니다. (조회 전용)")
    
    # 모바일 터치용 인물 선택 드롭다운
    selected_person = st.selectbox("👤 선생님(역할)을 선택하세요:", people)
    person_col_idx = list(header_row).index(selected_person)
    
    # 선택된 인물의 스케줄 정리
    schedule_data = []
    current_day = "DAY1"
    
    for idx in range(header_row_idx + 1, len(df_raw)):
        row = df_raw.iloc[idx]
        time_val = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        
        # 날짜(DAY) 구분선 확인
        if "DAY" in time_val.upper():
            current_day = time_val
            continue
            
        # 해당 인물의 업무 확인
        task_val = str(row.iloc[person_col_idx]).strip() if pd.notna(row.iloc[person_col_idx]) else ""
        
        if task_val.lower() in ["nan", "none"]: task_val = ""
        if time_val.lower() in ["nan", "none"]: time_val = ""
        
        # 시간이나 업무 내용이 하나라도 있으면 표에 추가
        if time_val or task_val:
            schedule_data.append({
                "DAY": current_day,
                "시간": time_val,
                "담당 업무": task_val
            })
            
    df_schedule = pd.DataFrame(schedule_data)
    
    st.subheader(f"✅ {selected_person} 선생님 일정")
    
    # DAY별 탭(Tab) 생성 및 데이터 출력
    days = df_schedule["DAY"].unique()
    if len(days) > 0:
        tabs = st.tabs(list(days))
        
        for i, day in enumerate(days):
            with tabs[i]:
                # 해당 DAY의 데이터만 필터링
                day_df = df_schedule[df_schedule["DAY"] == day]
                
                # 모바일에서 깔끔하게 보이도록 표 출력
                st.dataframe(
                    day_df[["시간", "담당 업무"]],
                    hide_index=True,
                    use_container_width=True
                )

except Exception as e:
    st.error(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")