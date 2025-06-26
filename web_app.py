# Streamlit 웹 애플리케이션: 연구과제 참여이력 통합
# 기존 project_participation_excel.py의 함수들을 재사용

import streamlit as st
import pandas as pd
import re
import os
import datetime
import tempfile
import io
from pathlib import Path

# 페이지 설정
st.set_page_config(
    page_title="연구과제 참여이력 통합",
    page_icon="📊",
    layout="wide"
)

# 기존 함수들 import (같은 폴더의 project_participation_excel.py에서)
from project_participation_excel import parse_txt_file, save_merged_to_excel

def main():
    st.title("📊 연구과제 참여이력 통합 시스템")
    st.markdown("---")
    
    # 사이드바 - 사용법 안내
    with st.sidebar:
        st.header("📖 사용법")
        st.markdown("""
        1. **파일 업로드**: 여러 txt 파일을 선택하세요
        2. **데이터 확인**: 탭에서 결과를 미리보기하세요
        3. **엑셀 다운로드**: 버튼을 클릭해 엑셀 파일을 받으세요
        
        **지원 형식**: 연구과제 참여확인서 txt 파일
        """)
        
        st.header("⚠️ 주의사항")
        st.markdown("""
        - txt 파일은 UTF-8 또는 CP949 인코딩이어야 합니다
        - 파일 형식이 다르면 오류가 발생할 수 있습니다
        - 다운로드 후 파일은 자동으로 삭제됩니다
        """)
    
    # 메인 영역
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.header("📁 파일 업로드")
        uploaded_files = st.file_uploader(
            "연구과제 참여확인서 txt 파일들을 선택하세요",
            type=['txt'],
            accept_multiple_files=True,
            help="여러 파일을 동시에 선택할 수 있습니다"
        )
    
    with col2:
        st.header("📈 처리 현황")
        if uploaded_files:
            st.success(f"📄 {len(uploaded_files)}개 파일 업로드 완료")
        else:
            st.info("📄 파일을 업로드해주세요")
    
    # 파일 처리 및 결과 표시
    if uploaded_files:
        try:
            with st.spinner("파일을 처리하고 있습니다..."):
                # 임시 디렉토리에 파일 저장
                temp_dir = tempfile.mkdtemp()
                temp_files = []
                
                for uploaded_file in uploaded_files:
                    temp_path = os.path.join(temp_dir, uploaded_file.name)
                    with open(temp_path, 'wb') as f:
                        f.write(uploaded_file.getbuffer())
                    temp_files.append(temp_path)
                
                # 기존 함수로 데이터 처리
                all_project_info = []
                all_researcher_info = []
                
                for temp_file in temp_files:
                    try:
                        project_info_list, researcher_info_list = parse_txt_file(temp_file)
                        all_project_info.extend(project_info_list)
                        all_researcher_info.extend(researcher_info_list)
                    except Exception as e:
                        st.error(f"파일 '{os.path.basename(temp_file)}' 처리 중 오류: {str(e)}")
                
                # 임시 파일들 삭제
                for temp_file in temp_files:
                    try:
                        os.remove(temp_file)
                    except:
                        pass
                try:
                    os.rmdir(temp_dir)
                except:
                    pass
                
                if all_project_info or all_researcher_info:
                    # 결과 표시
                    st.success("✅ 파일 처리 완료!")
                    
                    # 탭으로 결과 표시
                    tab1, tab2, tab3, tab4 = st.tabs(["📋 과제정보", "👥 연구원정보", "🔄 연구원정보_기간통합", "💾 엑셀 다운로드"])
                    
                    with tab1:
                        st.header("과제정보")
                        if all_project_info:
                            project_columns = ['과제번호', '연구기간', '과 제 명', '연구책임자', '지원기관', '지원사업', '소속연구소', '관리부서', '협약연구비', '공동연구원수', '연구보조원수']
                            project_df = pd.DataFrame(all_project_info).drop_duplicates().reindex(columns=project_columns)
                            st.dataframe(project_df, use_container_width=True)
                            st.info(f"총 {len(project_df)}개의 과제 정보가 있습니다.")
                        else:
                            st.warning("과제 정보가 없습니다.")
                    
                    with tab2:
                        st.header("연구원정보")
                        if all_researcher_info:
                            researcher_columns = ['과제번호', '과 제 명', '성명', '주민번호', '연구원구분', '과정구분', '소속', '참여기간']
                            researcher_df = pd.DataFrame(all_researcher_info).reindex(columns=researcher_columns)
                            # 지원기관 정보 추가
                            project_df_for_merge = pd.DataFrame(all_project_info)[['과제번호', '지원기관']].drop_duplicates()
                            researcher_df = pd.merge(researcher_df, project_df_for_merge, on='과제번호', how='left')
                            cols = list(researcher_df.columns)
                            cols.insert(2, cols.pop(cols.index('지원기관')))
                            researcher_df = researcher_df[cols]
                            st.dataframe(researcher_df, use_container_width=True)
                            st.info(f"총 {len(researcher_df)}개의 연구원 참여 기록이 있습니다.")
                        else:
                            st.warning("연구원 정보가 없습니다.")
                    
                    with tab3:
                        st.header("연구원정보_기간통합")
                        if all_researcher_info:
                            researcher_columns = ['과제번호', '과 제 명', '성명', '주민번호', '연구원구분', '과정구분', '소속', '참여기간']
                            researcher_df = pd.DataFrame(all_researcher_info).reindex(columns=researcher_columns)
                            # 지원기관 정보 추가
                            project_df_for_merge = pd.DataFrame(all_project_info)[['과제번호', '지원기관']].drop_duplicates()
                            researcher_df = pd.merge(researcher_df, project_df_for_merge, on='과제번호', how='left')
                            cols = list(researcher_df.columns)
                            cols.insert(2, cols.pop(cols.index('지원기관')))
                            researcher_df = researcher_df[cols]
                            # 기간 병합 함수
                            def parse_period(period):
                                start, end = period.split('~')
                                return start.strip(), end.strip()
                            def merge_periods(periods):
                                periods = sorted([parse_period(p) for p in periods], key=lambda x: x[0])
                                merged = []
                                for s, e in periods:
                                    if not merged:
                                        merged.append([s, e])
                                    else:
                                        last_s, last_e = merged[-1]
                                        last_end = datetime.datetime.strptime(last_e, "%Y-%m-%d")
                                        this_start = datetime.datetime.strptime(s, "%Y-%m-%d")
                                        if (this_start - last_end).days == 1:
                                            merged[-1][1] = e
                                        else:
                                            merged.append([s, e])
                                return ['{} ~ {}'.format(s, e) for s, e in merged]
                            # 연구원별로 기간 병합
                            researcher_names = researcher_df['성명'].dropna().unique()
                            all_merged_data = []
                            for name in researcher_names:
                                r_df = researcher_df[researcher_df['성명'] == name].copy()
                                group_cols = [col for col in r_df.columns if col != '참여기간']
                                merged_rows = []
                                for _, group in r_df.groupby(group_cols, dropna=False):
                                    periods = group['참여기간'].tolist()
                                    for merged_period in merge_periods(periods):
                                        row = group.iloc[0].to_dict()
                                        row['참여기간'] = merged_period
                                        merged_rows.append(row)
                                all_merged_data.extend(merged_rows)
                            if all_merged_data:
                                merged_df = pd.DataFrame(all_merged_data).reindex(columns=researcher_df.columns)
                                st.dataframe(merged_df, use_container_width=True)
                                st.info(f"총 {len(merged_df)}개의 기간통합된 연구원 참여 기록이 있습니다.")
                            else:
                                st.warning("기간통합할 연구원 정보가 없습니다.")
                        else:
                            st.warning("연구원 정보가 없습니다.")
                    
                    with tab4:
                        st.header("엑셀 파일 다운로드")
                        st.markdown("처리된 데이터를 엑셀 파일로 다운로드할 수 있습니다.")
                        # 엑셀 파일 생성
                        output_buffer = io.BytesIO()
                        with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                            # 과제정보 시트
                            if all_project_info:
                                project_columns = ['과제번호', '연구기간', '과 제 명', '연구책임자', '지원기관', '지원사업', '소속연구소', '관리부서', '협약연구비', '공동연구원수', '연구보조원수']
                                project_df = pd.DataFrame(all_project_info).drop_duplicates().reindex(columns=project_columns)
                                project_df.to_excel(writer, sheet_name='과제정보', index=False)
                            # 연구원별 시트 생성
                            if all_researcher_info:
                                researcher_columns = ['과제번호', '과 제 명', '성명', '주민번호', '연구원구분', '과정구분', '소속', '참여기간']
                                researcher_df = pd.DataFrame(all_researcher_info).reindex(columns=researcher_columns)
                                # 지원기관 정보 추가
                                project_df_for_merge = pd.DataFrame(all_project_info)[['과제번호', '지원기관']].drop_duplicates()
                                researcher_df = pd.merge(researcher_df, project_df_for_merge, on='과제번호', how='left')
                                cols = list(researcher_df.columns)
                                cols.insert(2, cols.pop(cols.index('지원기관')))
                                researcher_df = researcher_df[cols]
                                # 연구원별로 시트 분리
                                researcher_names = researcher_df['성명'].dropna().unique()
                                for name in researcher_names:
                                    r_df = researcher_df[researcher_df['성명'] == name].copy()
                                    r_df.to_excel(writer, sheet_name=name, index=False)
                                    # 기간 병합 시트
                                    def parse_period(period):
                                        start, end = period.split('~')
                                        return start.strip(), end.strip()
                                    def merge_periods(periods):
                                        periods = sorted([parse_period(p) for p in periods], key=lambda x: x[0])
                                        merged = []
                                        for s, e in periods:
                                            if not merged:
                                                merged.append([s, e])
                                            else:
                                                last_s, last_e = merged[-1]
                                                last_end = datetime.datetime.strptime(last_e, "%Y-%m-%d")
                                                this_start = datetime.datetime.strptime(s, "%Y-%m-%d")
                                                if (this_start - last_end).days == 1:
                                                    merged[-1][1] = e
                                                else:
                                                    merged.append([s, e])
                                        return ['{} ~ {}'.format(s, e) for s, e in merged]
                                    group_cols = [col for col in r_df.columns if col != '참여기간']
                                    merged_rows = []
                                    for _, group in r_df.groupby(group_cols, dropna=False):
                                        periods = group['참여기간'].tolist()
                                        for merged_period in merge_periods(periods):
                                            row = group.iloc[0].to_dict()
                                            row['참여기간'] = merged_period
                                            merged_rows.append(row)
                                    merged_researcher_df = pd.DataFrame(merged_rows).reindex(columns=r_df.columns)
                                    merged_sheet_name = f"{name}_기간통합"
                                    merged_researcher_df.to_excel(writer, sheet_name=merged_sheet_name, index=False)
                        
                        output_buffer.seek(0)
                        
                        # 다운로드 버튼
                        now_str = datetime.datetime.now().strftime('%Y%m%d_%H%M')
                        filename = f'연구과제_참여이력_통합_{now_str}.xlsx'
                        
                        st.download_button(
                            label="📥 엑셀 파일 다운로드",
                            data=output_buffer.getvalue(),
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            help="클릭하면 엑셀 파일이 다운로드됩니다"
                        )
                        
                        st.success("✅ 엑셀 파일이 준비되었습니다. 위 버튼을 클릭하여 다운로드하세요.")
                
                else:
                    st.error("❌ 처리할 수 있는 데이터가 없습니다. 파일 형식을 확인해주세요.")
                    
        except Exception as e:
            st.error(f"❌ 처리 중 오류가 발생했습니다: {str(e)}")
            st.info("파일 형식이나 인코딩을 확인해주세요.")

if __name__ == "__main__":
    main() 