# -*- coding: utf-8 -*-
"""
세그먼트별 감정 궤적 분석
- 상담유형별, 연령대별, 제품군별 감정 패턴
- 종료 감정 긍정/부정 그룹 비교
- 결과 엑셀 저장
"""
import sys, io, warnings, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
warnings.filterwarnings('ignore')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
import numpy as np
import json
from collections import Counter

# 데이터 로드
EXPORTS = 'outputs/exports'
df_calls = pd.read_csv(f'{EXPORTS}/전체_음성분석_데이터_LLM_콜단위요약.csv', encoding='utf-8-sig')
df_turns = pd.read_csv(f'{EXPORTS}/전체_음성분석_데이터_LLM_발화단위상세.csv', encoding='utf-8-sig')

# 원본 데이터에서 상담유형 가져오기
from data_loader import load_call_data
df_raw = load_call_data()

# df_calls에 없는 컬럼만 raw에서 가져오기
df_raw['CNID'] = df_raw['CNID'].astype(str).str.strip()
df_calls['CNID'] = df_calls['CNID'].astype(str).str.strip()

# 상담유형만 raw에서 추가 (연령대/성별/제품대는 df_calls에 이미 있음)
meta = df_raw[['CNID','상담번호: 상담유형(중)','컨설턴트 만족도']].copy()
meta.columns = ['CNID','상담유형','컨설턴트만족도']
df_merged = df_calls.merge(meta, on='CNID', how='left')

cust = df_turns[df_turns['화자']=='고객'].copy()
cust['CNID'] = cust['CNID'].astype(str).str.strip()
cust_merged = cust.merge(df_merged[['CNID','상담유형','연령대','성별','제품대']], on='CNID', how='left')

print(f'merge 완료: {len(df_merged)}건')

sep = '='*60

# ══ 1. 상담유형별 단계 감성 패턴 ══════════════════════════════════
print(f'{sep}\n1. 상담유형별 단계 감성 패턴\n{sep}')

top_types = df_merged['상담유형'].value_counts().head(5).index
for stype in top_types:
    cnids = df_merged[df_merged['상담유형']==stype]['CNID']
    sub = cust_merged[cust_merged['CNID'].isin(cnids)]
    if len(sub) < 20:
        continue
    print(f'\n  [{stype}] (N={len(cnids)})')
    for stage in ['초기','탐색','해결시도','결과제시','종료']:
        vals = sub[sub['단계']==stage]['융합Valence'].dropna()
        if len(vals) > 0:
            print(f'    {stage}: {vals.mean():+.3f} (n={len(vals)})')

# ══ 2. 연령대별 단계 감성 패턴 ══════════════════════════════════
print(f'\n{sep}\n2. 연령대별 단계 감성 패턴\n{sep}')

for age in ['20~39세','40~49세','50~64세','65~74세','75세 이상']:
    cnids = df_merged[df_merged['연령대']==age]['CNID']
    sub = cust_merged[cust_merged['CNID'].isin(cnids)]
    if len(sub) < 20:
        continue
    print(f'\n  [{age}] (N={len(cnids)})')
    for stage in ['초기','탐색','해결시도','결과제시','종료']:
        vals = sub[sub['단계']==stage]['융합Valence'].dropna()
        if len(vals) > 0:
            print(f'    {stage}: {vals.mean():+.3f} (n={len(vals)})')

# ══ 3. 제품군별 단계 감성 패턴 ══════════════════════════════════
print(f'\n{sep}\n3. 제품군별 단계 감성 패턴\n{sep}')

for prod in ['주방가전','생활가전','TV/AV','에어컨/에어케어']:
    cnids = df_merged[df_merged['제품대']==prod]['CNID']
    sub = cust_merged[cust_merged['CNID'].isin(cnids)]
    if len(sub) < 20:
        continue
    print(f'\n  [{prod}] (N={len(cnids)})')
    for stage in ['초기','탐색','해결시도','결과제시','종료']:
        vals = sub[sub['단계']==stage]['융합Valence'].dropna()
        if len(vals) > 0:
            print(f'    {stage}: {vals.mean():+.3f} (n={len(vals)})')

# ══ 4. 종료 감정 긍정 vs 부정 그룹 비교 ══════════════════════════
print(f'\n{sep}\n4. 종료 감정 긍정 vs 부정 그룹 비교\n{sep}')

# 종료 단계 감성 기준으로 분류
pos_cnids = df_merged[df_merged['종료_감성'] > 0.05]['CNID']
neg_cnids = df_merged[df_merged['종료_감성'] < -0.05]['CNID']
neu_cnids = df_merged[(df_merged['종료_감성'] >= -0.05) & (df_merged['종료_감성'] <= 0.05)]['CNID']

for label, cnids in [('긍정 종료', pos_cnids), ('부정 종료', neg_cnids), ('중립 종료', neu_cnids)]:
    sub = cust_merged[cust_merged['CNID'].isin(cnids)]
    nps = df_merged[df_merged['CNID'].isin(cnids)]['NPS'].mean()
    print(f'\n  [{label}] N={len(cnids)}, NPS={nps:.2f}')
    for stage in ['초기','탐색','해결시도','결과제시','종료']:
        vals = sub[sub['단계']==stage]['융합Valence'].dropna()
        if len(vals) > 0:
            print(f'    {stage}: {vals.mean():+.3f}')

    # 주요 트리거 (부정→긍정 or 긍정→부정 전환)
    transitions = 0
    for cnid in cnids:
        ct = sub[sub['CNID']==cnid].sort_values('시작(초)')
        fv = ct['융합Valence'].dropna().values
        for i in range(len(fv)-1):
            if abs(fv[i+1]-fv[i]) >= 0.15:
                transitions += 1
    if len(cnids) > 0:
        print(f'    전환 빈도: 콜당 {transitions/len(cnids):.1f}회')

# ══ 엑셀 저장 ════════════════════════════════════════════════════
output = os.path.join(os.path.dirname(os.path.abspath(__file__)), '세그먼트별_분석결과.xlsx')

with pd.ExcelWriter(output, engine='openpyxl') as writer:
    # 상담유형별
    rows = []
    for stype in top_types:
        cnids = df_merged[df_merged['상담유형']==stype]['CNID']
        sub = cust_merged[cust_merged['CNID'].isin(cnids)]
        row = {'상담유형': stype, 'N': len(cnids)}
        for stage in ['초기','탐색','해결시도','결과제시','종료']:
            vals = sub[sub['단계']==stage]['융합Valence'].dropna()
            row[stage] = round(vals.mean(), 3) if len(vals) > 0 else None
        rows.append(row)
    pd.DataFrame(rows).to_excel(writer, sheet_name='상담유형별', index=False)

    # 연령대별
    rows = []
    for age in ['20~39세','40~49세','50~64세','65~74세','75세 이상']:
        cnids = df_merged[df_merged['연령대']==age]['CNID']
        sub = cust_merged[cust_merged['CNID'].isin(cnids)]
        row = {'연령대': age, 'N': len(cnids)}
        for stage in ['초기','탐색','해결시도','결과제시','종료']:
            vals = sub[sub['단계']==stage]['융합Valence'].dropna()
            row[stage] = round(vals.mean(), 3) if len(vals) > 0 else None
        rows.append(row)
    pd.DataFrame(rows).to_excel(writer, sheet_name='연령대별', index=False)

    # 종료 그룹별
    rows = []
    for label, cnids in [('긍정', pos_cnids), ('부정', neg_cnids), ('중립', neu_cnids)]:
        sub = cust_merged[cust_merged['CNID'].isin(cnids)]
        nps = df_merged[df_merged['CNID'].isin(cnids)]['NPS'].mean()
        row = {'종료감정': label, 'N': len(cnids), 'NPS': round(nps, 2)}
        for stage in ['초기','탐색','해결시도','결과제시','종료']:
            vals = sub[sub['단계']==stage]['융합Valence'].dropna()
            row[stage] = round(vals.mean(), 3) if len(vals) > 0 else None
        rows.append(row)
    pd.DataFrame(rows).to_excel(writer, sheet_name='종료감정그룹별', index=False)

print(f'\n엑셀 저장: {output}')
