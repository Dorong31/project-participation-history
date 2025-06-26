#ver.1_250625_연구과제 참여확인서에서 과제별 참여 이력 추출
#입력자료: 연구포털에서 연구과제 일괄출력(조회) 후 .txt 파일로 저장
#출력자료: 연구과제 참여이력 통합.xlsx

import pandas as pd
import re
import os
import datetime

# 텍스트 파일 하나에서 과제정보/연구원정보 추출
# 반환: (과제정보 리스트, 연구원정보 리스트)
def parse_txt_file(txt_path):
    with open(txt_path, encoding='cp949') as f:
        content = f.read()
    # 1. 과제별로 분리
    projects = [p for p in content.split('연구과제 참여확인서') if p.strip()]
    project_info_list = []
    researcher_info_list = []
    for project in projects:
        # 2. 과제정보, 연구원정보 블록 추출
        info_block = re.search(r'■ 과제정보(.*?)■ 연구원정보', project, re.DOTALL)
        researcher_block = re.search(r'■ 연구원정보(.*?)(?:-- 이하 여백 --|202\d{1,4}년|\Z)', project, re.DOTALL)
        if not info_block or not researcher_block:
            continue
        info_lines = info_block.group(1).strip().split('\n')
        researcher_lines = researcher_block.group(1).strip().split('\n')
        # 3. 과제정보 추출
        info_dict = {}
        for line in info_lines:
            for key in ['과제번호', '연구기간', '과 제 명', '연구책임자', '지원기관', '지원사업', '소속연구소', '관리부서', '협약연구비', '공동연구원수', '연구보조원수']:
                if key in line:
                    parts = line.split('\t')
                    for i, part in enumerate(parts):
                        if part.strip() == key and i+1 < len(parts):
                            info_dict[key] = parts[i+1].strip()
        project_number = info_dict.get('과제번호', '')
        project_name = info_dict.get('과 제 명', '')
        # 과제정보 리스트에 추가
        if project_number:
            project_info_list.append(info_dict)
        # 4. 연구원정보 추출
        if len(researcher_lines) < 3:
            continue
        name_line = researcher_lines[0]
        name = re.search(r'성명:\s*([^\t]+)', name_line)
        jumin = re.search(r'주민번호:\s*([^\t]+)', name_line)
        name = name.group(1).strip() if name else ''
        jumin = jumin.group(1).strip() if jumin else ''
        for line in researcher_lines[2:]:
            if not line.strip():
                continue
            parts = line.strip().split('\t')
            if len(parts) < 4:
                continue
            researcher_dict = {
                '과제번호': project_number,
                '과 제 명': project_name,
                '성명': name,
                '주민번호': jumin,
                '연구원구분': parts[0].strip(),
                '과정구분': parts[1].strip(),
                '소속': parts[2].strip(),
                '참여기간': parts[3].strip()
            }
            researcher_info_list.append(researcher_dict)
    return project_info_list, researcher_info_list

# 여러 txt 파일을 통합 처리하여 엑셀로 저장
# - 과제정보는 중복 없이 하나의 시트
# - 연구원정보/기간병합은 연구원별로 각각 시트 생성
def save_merged_to_excel(all_project_info, all_researcher_info, output_path):
    # 1. 과제정보 통합 및 중복 제거
    project_columns = ['과제번호', '연구기간', '과 제 명', '연구책임자', '지원기관', '지원사업', '소속연구소', '관리부서', '협약연구비', '공동연구원수', '연구보조원수']
    project_df = pd.DataFrame(all_project_info).drop_duplicates().reindex(columns=project_columns)
    # 2. 연구원정보 통합
    researcher_columns = ['과제번호', '과 제 명', '성명', '주민번호', '연구원구분', '과정구분', '소속', '참여기간']
    researcher_df = pd.DataFrame(all_researcher_info).reindex(columns=researcher_columns)
    # 2-1. 과제번호 기준으로 지원기관 정보 매핑
    project_df_for_merge = project_df[['과제번호', '지원기관']].drop_duplicates()
    researcher_df = pd.merge(researcher_df, project_df_for_merge, on='과제번호', how='left')
    # 지원기관을 세 번째 열로 이동
    cols = list(researcher_df.columns)
    cols.insert(2, cols.pop(cols.index('지원기관')))
    researcher_df = researcher_df[cols]
    # 3. 연구원별로 시트 분리
    researcher_names = researcher_df['성명'].dropna().unique()
    # 4. 참여기간 병합 함수
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
    # 5. 엑셀로 저장
    with pd.ExcelWriter(output_path) as writer:
        # 과제정보 시트
        project_df.to_excel(writer, sheet_name='과제정보', index=False)
        # 연구원별 시트 생성
        for name in researcher_names:
            r_df = researcher_df[researcher_df['성명'] == name].copy()
            r_df.to_excel(writer, sheet_name=name, index=False)
            # 기간 병합
            group_cols = [col for col in r_df.columns if col != '참여기간']
            merged_rows = []
            for _, group in r_df.groupby(group_cols, dropna=False):
                periods = group['참여기간'].tolist()
                for merged_period in merge_periods(periods):
                    row = group.iloc[0].to_dict()
                    row['참여기간'] = merged_period
                    merged_rows.append(row)
            merged_researcher_df = pd.DataFrame(merged_rows).reindex(columns=r_df.columns)
            merged_sheet_name = f"{name}(통합)"
            merged_researcher_df.to_excel(writer, sheet_name=merged_sheet_name, index=False)

# 현재 폴더 내 모든 txt 파일을 통합 처리
# - 모든 txt 파일에서 과제정보/연구원정보를 추출해 통합
# - 엑셀로 저장
def process_all_txt_files_and_merge():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    txt_files = [f for f in os.listdir(current_dir) if f.lower().endswith('.txt')]
    all_project_info = []
    all_researcher_info = []
    for txt_file in txt_files:
        txt_path = os.path.join(current_dir, txt_file)
        project_info_list, researcher_info_list = parse_txt_file(txt_path)
        all_project_info.extend(project_info_list)
        all_researcher_info.extend(researcher_info_list)
    # 결과 엑셀 파일명에 실행 시각 추가
    now_str = datetime.datetime.now().strftime('%Y%m%d_%H%M')
    output_path = os.path.join(current_dir, f'연구과제_참여이력_통합_{now_str}.xlsx')
    save_merged_to_excel(all_project_info, all_researcher_info, output_path)
    print(f"모든 txt 파일을 통합하여 {os.path.basename(output_path)}로 저장 완료.")

if __name__ == "__main__":
    process_all_txt_files_and_merge()