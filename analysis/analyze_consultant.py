"""
컨설턴트 만족도 vs 음성 감성 상관분석
- NPS와 비교
- 컨설턴트 만족도 그룹별 감성 패턴
"""
import sys, io, warnings, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
warnings.filterwarnings('ignore')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
import numpy as np
import json

EXPORTS = 'outputs/exports'
df_calls = pd.read_csv(f'{EXPORTS}/전체_음성분석_데이터_LLM_콜단위요약.csv', encoding='utf-8-sig')
df_turns = pd.read_csv(f'{EXPORTS}/전체_음성분석_데이터_LLM_발화단위상세.csv', encoding='utf-8-sig')

from data_loader import load_call_data
df_raw = load_call_data()
df_raw['CNID'] = df_raw['CNID'].astype(str).str.strip()
df_calls['CNID'] = df_calls['CNID'].astype(str).str.strip()

# 컨설턴트 만족도는 df_calls에 이미 있음, raw에서 원본 값 확인
meta = df_raw[['CNID','컨설턴트 만족도']].copy()
meta.columns = ['CNID','컨설턴트만족도_raw']
df_merged = df_calls.merge(meta, on='CNID', how='left')
# 컨설턴트만족도 컬럼명 통일
if '컨설턴트만족도' not in df_merged.columns and '컨설턴트만족도_raw' in df_merged.columns:
    df_merged['컨설턴트만족도'] = df_merged['컨설턴트만족도_raw']
elif '컨설턴트만족도' not in df_merged.columns:
    df_merged['컨설턴트만족도'] = df_merged.get('컨설턴트만족도_raw')

cust = df_turns[df_turns['화자']=='고객'].copy()
cust['CNID'] = cust['CNID'].astype(str)

sep = '='*60

# ══ 1. 컨설턴트 만족도 vs 감성 상관 ══════════════════════════════
print(f'{sep}\n1. 컨설턴트 만족도 vs 감성 상관\n{sep}')

valid = df_merged[['컨설턴트만족도','고객평균감성','NPS']].dropna()
corr_cons = valid['컨설턴트만족도'].corr(valid['고객평균감성'])
corr_nps = valid['NPS'].corr(valid['고객평균감성'])
corr_cons_nps = valid['컨설턴트만족도'].corr(valid['NPS'])

print(f'컨설턴트만족도 ↔ 감성: r = {corr_cons:.3f}')
print(f'NPS ↔ 감성:           r = {corr_nps:.3f}')
print(f'컨설턴트만족도 ↔ NPS:  r = {corr_cons_nps:.3f}')

# ══ 2. 컨설턴트 만족도 그룹별 감성 ═══════════════════════════════
print(f'\n{sep}\n2. 컨설턴트 만족도 그룹별\n{sep}')

for score in [1,2,3,4,5]:
    sub = df_merged[df_merged['컨설턴트만족도']==score]
    if len(sub) < 5:
        continue
    avg_v = sub['고객평균감성'].mean()
    avg_nps = sub['NPS'].mean()
    print(f'  만족도 {score}점: N={len(sub)}, 감성={avg_v:+.3f}, NPS={avg_nps:.2f}')

# ══ 3. NPS-컨설턴트 괴리 분석 ════════════════════════════════════
print(f'\n{sep}\n3. NPS-컨설턴트 괴리 분석\n{sep}')

# NPS 높은데 컨설턴트 낮은 케이스
high_nps_low_cons = df_merged[(df_merged['NPS']>=9) & (df_merged['컨설턴트만족도']<=3)]
low_nps_high_cons = df_merged[(df_merged['NPS']<=5) & (df_merged['컨설턴트만족도']>=4)]

print(f'NPS 높은데(9+) 컨설턴트 낮음(3-): {len(high_nps_low_cons)}건')
if len(high_nps_low_cons) > 0:
    print(f'  평균 감성: {high_nps_low_cons["고객평균감성"].mean():+.3f}')

print(f'NPS 낮은데(5-) 컨설턴트 높음(4+): {len(low_nps_high_cons)}건')
if len(low_nps_high_cons) > 0:
    print(f'  평균 감성: {low_nps_high_cons["고객평균감성"].mean():+.3f}')

# ══ 4. 단계별 감성 - 컨설턴트 만족도별 ═══════════════════════════
print(f'\n{sep}\n4. 단계별 감성 (컨설턴트 만족도별)\n{sep}')

for score_range, label in [((1,3),'불만족(1-3)'), ((4,4),'보통(4)'), ((5,5),'만족(5)')]:
    cnids = df_merged[(df_merged['컨설턴트만족도']>=score_range[0]) & (df_merged['컨설턴트만족도']<=score_range[1])]['CNID']
    sub = cust[cust['CNID'].isin(cnids)]
    if len(sub) < 20:
        continue
    print(f'\n  [{label}] N={len(cnids)}')
    for stage in ['초기','탐색','해결시도','결과제시','종료']:
        vals = sub[sub['단계']==stage]['융합Valence'].dropna()
        if len(vals) > 0:
            print(f'    {stage}: {vals.mean():+.3f}')

# 엑셀 저장
output = os.path.join(os.path.dirname(os.path.abspath(__file__)), '컨설턴트만족도_분석결과.xlsx')
results = {
    '상관분석': pd.DataFrame([
        {'지표쌍': '컨설턴트만족도 ↔ 감성', 'Pearson_r': round(corr_cons, 3), 'N': len(valid)},
        {'지표쌍': 'NPS ↔ 감성', 'Pearson_r': round(corr_nps, 3), 'N': len(valid)},
        {'지표쌍': '컨설턴트만족도 ↔ NPS', 'Pearson_r': round(corr_cons_nps, 3), 'N': len(valid)},
    ]),
}

with pd.ExcelWriter(output, engine='openpyxl') as writer:
    for sheet, data in results.items():
        data.to_excel(writer, sheet_name=sheet, index=False)

print(f'\n엑셀 저장: {output}')
