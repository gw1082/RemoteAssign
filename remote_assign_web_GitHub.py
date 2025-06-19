import streamlit as st
import pandas as pd
import numpy as np
import random
import math
from io import BytesIO

st.set_page_config(page_title="재택근무 자동배정", layout="wide")
st.title("재택근무 자동배정기 (웹버전)")

st.write("엑셀 파일을 업로드하면, 부서별 32~40% 미만 재택율과 직원별 3~4일 재택만 나오도록 자동 배정합니다.")

uploaded_file = st.file_uploader("엑셀 파일 업로드 (.xlsx)", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0)
    base_cols = ['부서', '직원', '문정', '잠실']
    days1 = ['월', '화', '수', '목', '금']
    days2 = ['Next_월', 'Next_화', 'Next_수', 'Next_목', 'Next_금']
    all_days = days1 + days2

    # Next 요일 열이 없으면 추가
    for day in days2:
        if day not in df.columns:
            df[day] = ""

    result_rows = []
    warnings = []

    munjung_limit = 14
    jamsil_limit = 30
    munjung_counts = {day: 0 for day in all_days}
    jamsil_counts = {day: 0 for day in all_days}

    def assign_office_days(row, all_days, munjung_limit, jamsil_limit, munjung_counts, jamsil_counts):
        try:
            office_munjung = int(row['문정'])
            office_jamsil = int(row['잠실'])
        except:
            office_munjung = 0
            office_jamsil = 0

        work_days = [day for day in all_days if row[day] != '재택']
        total_work = len(work_days)
        if total_work == 0:
            return row

        if office_munjung + office_jamsil == 0:
            munjung_days = total_work
            jamsil_days = 0
        else:
            munjung_days = int(total_work * office_munjung / (office_munjung + office_jamsil))
            jamsil_days = total_work - munjung_days

        day_types = ['문정'] * munjung_days + ['잠실'] * jamsil_days
        while len(day_types) < total_work:
            day_types.append(random.choice(['문정', '잠실']))
        random.shuffle(day_types)

        for i, day in enumerate(work_days):
            if day_types[i] == '문정':
                if munjung_counts[day] < munjung_limit:
                    row[day] = '문정'
                    munjung_counts[day] += 1
                elif jamsil_counts[day] < jamsil_limit:
                    row[day] = '잠실'
                    jamsil_counts[day] += 1
                else:
                    row[day] = ''
            elif day_types[i] == '잠실':
                if jamsil_counts[day] < jamsil_limit:
                    row[day] = '잠실'
                    jamsil_counts[day] += 1
                elif munjung_counts[day] < munjung_limit:
                    row[day] = '문정'
                    munjung_counts[day] += 1
                else:
                    row[day] = ''
        return row

    # 부서별로 재택 및 출근지 배정
    for dept, group in df.groupby('부서'):
        group = group.reset_index(drop=True)
        total_people = len(group)
        total_days = len(all_days)

        min_remote = math.ceil(total_people * 0.32)
        max_remote = math.floor(total_people * 0.4) - 1
        if max_remote < min_remote:
            max_remote = min_remote
        possible_nums = []
        for n in range(min_remote, total_people+1):
            percent = n / total_people * 100
            if 32 <= percent < 40:
                possible_nums.append(n)
        if not possible_nums:
            best_n = min(range(1, total_people+1), key=lambda n: abs((n/total_people*100)-36))
            warnings.append(f"⚠️ 부서 [{dept}]는 인원수가 {total_people}명이라 32~40%에 맞는 정수 인원 배정 불가. {best_n}명({round(best_n/total_people*100,1)}%)로 배정.")
            remote_per_day = best_n
        else:
            remote_per_day = possible_nums[0]

        total_remote = remote_per_day * total_days
        base_remote = total_remote // total_people
        extra_remote = total_remote % total_people

        # 3~4일만 나오도록 분배
        if base_remote < 3:
            base_remote = 3
            extra_remote = 0
        elif base_remote > 4:
            base_remote = 4
            extra_remote = 0
        elif base_remote == 3 and extra_remote > 0:
            pass
        elif base_remote == 4 and extra_remote > 0:
            base_remote = 3
            extra_remote = total_people - extra_remote

        indices = list(range(total_people))
        random.shuffle(indices)
        remote_goal = [base_remote] * total_people
        for idx in indices[:extra_remote]:
            remote_goal[idx] += 1

        person_left = remote_goal.copy()
        day_left = [remote_per_day for _ in range(total_days)]
        assign = np.zeros((total_people, total_days), dtype=int)

        for d in range(total_days):
            candidates = [i for i in range(total_people) if person_left[i] > 0]
            if len(candidates) < remote_per_day:
                extra_needed = remote_per_day - len(candidates)
                extra_candidates = [i for i in range(total_people) if person_left[i] == 0]
                candidates += random.sample(extra_candidates, extra_needed)
            random.shuffle(candidates)
            candidates = sorted(candidates, key=lambda x: person_left[x], reverse=True)
            selected = candidates[:remote_per_day]
            for idx in selected:
                assign[idx, d] = 1
                if person_left[idx] > 0:
                    person_left[idx] -= 1
                day_left[d] -= 1

        for i in range(total_people):
            for j, day in enumerate(all_days):
                if assign[i, j] == 1:
                    group.at[i, day] = '재택'
                else:
                    group.at[i, day] = ''

        # 출근지 배정
        for idx, row in group.iterrows():
            group.loc[idx] = assign_office_days(row, all_days, munjung_limit, jamsil_limit, munjung_counts, jamsil_counts)

        result_rows.append(group)

    result_df = pd.concat(result_rows, ignore_index=True)

    # 직원별 근무일수 열 추가
    munjung_col = []
    jamsil_col = []
    jaetaek_col = []
    total_col = []
    for idx, row in result_df.iterrows():
        munjung = sum(row[day] == '문정' for day in all_days)
        jamsil = sum(row[day] == '잠실' for day in all_days)
        jaetaek = sum(row[day] == '재택' for day in all_days)
        total = munjung + jamsil + jaetaek
        munjung_col.append(munjung)
        jamsil_col.append(jamsil)
        jaetaek_col.append(jaetaek)
        total_col.append(total)
    result_df['문정근무일'] = munjung_col
    result_df['잠실근무일'] = jamsil_col
    result_df['재택일'] = jaetaek_col
    result_df['총합'] = total_col

    # 하단 요약행 9개 추가
    munjung_counts_summary = {}
    jamsil_counts_summary = {}
    remote_counts_summary = {}
    total_counts_summary = {}
    for day in all_days:
        munjung_counts_summary[day] = result_df[day].tolist().count('문정')
        jamsil_counts_summary[day] = result_df[day].tolist().count('잠실')
        remote_counts_summary[day] = result_df[day].tolist().count('재택')
        total_counts_summary[day] = munjung_counts_summary[day] + jamsil_counts_summary[day] + remote_counts_summary[day]

    def make_row(label, values_dict):
        row = {col: '' for col in result_df.columns}
        row['부서'] = ''
        row['직원'] = label
        for day in all_days:
            row[day] = values_dict.get(day, '')
        return row

    munjung_row = make_row('문정근무일', munjung_counts_summary)
    jamsil_row = make_row('잠실근무일', jamsil_counts_summary)
    remote_row = make_row('재택근무일', remote_counts_summary)
    total_row = make_row('총합', total_counts_summary)
    empty_row = {col: '' for col in result_df.columns}
    empty_row['부서'] = ''
    empty_row['직원'] = ''

    def make_rate_row(label, source_dict):
        row = {col: '' for col in result_df.columns}
        row['부서'] = ''
        row['직원'] = label
        for day in all_days:
            total = total_counts_summary[day]
            val = source_dict[day]
            row[day] = f"{round(val/total*100,1)}%" if total else '0%'
        return row

    munjung_rate_row = make_rate_row('문정비율', munjung_counts_summary)
    jamsil_rate_row = make_rate_row('잠실비율', jamsil_counts_summary)
    remote_rate_row = make_rate_row('재택비율', remote_counts_summary)
    rate_sum_row = make_rate_row('총비율', {day: munjung_counts_summary[day]+jamsil_counts_summary[day]+remote_counts_summary[day] for day in all_days})

    extra_rows_df = pd.DataFrame([
        munjung_row,
        jamsil_row,
        remote_row,
        total_row,
        empty_row,
        munjung_rate_row,
        jamsil_rate_row,
        remote_rate_row,
        rate_sum_row
    ])
    result_df = pd.concat([result_df, extra_rows_df], ignore_index=True)

    # 열 순서 재정렬
    ordered_cols = (
        base_cols
        + days1 + days2
        + ['문정근무일', '잠실근무일', '재택일', '총합']
    )
    rest_cols = [c for c in result_df.columns if c not in ordered_cols]
    final_cols = ordered_cols + rest_cols
    result_df = result_df[final_cols]

    st.success("자동 배정 결과 미리보기")
    st.dataframe(result_df, use_container_width=True)

    # 다운로드 버튼
    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='자동배정결과')
        return output.getvalue()

    st.download_button(
        label="엑셀로 다운로드",
        data=to_excel(result_df),
        file_name="재택근무_자동배정결과.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if warnings:
        st.warning('\n'.join(warnings))

else:
    st.info("엑셀 파일(.xlsx)을 업로드해 주세요. 샘플 양식은 관리자에게 문의하세요.")
