"""
짧은 발화("네", "아") 세부 감정 분석
- SER 감정 분포 (neu/hap/ang/sad)
- 전후 맥락별 "네"의 감정 차이
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

# 주의: SER 데이터는 재추출 후 사용 가능
# 현재는 기존 데이터에서 분석 가능한 항목만 수행
EXPORTS = 'outputs/exports'
df_turns = pd.read_csv(f'{EXPORTS}/전체_음성분석_데이터_LLM_발화단위상세.csv', encoding='utf-8-sig')

cust = df_turns[df_turns['화자']=='고객'].copy()
short = cust[cust['짧은발화']==True].copy()

sep = '='*60

# ══ 1. 짧은 발화 기본 통계 ════════════════════════════════════════
print(f'{sep}\n1. 짧은 발화 기본 통계\n{sep}')
print(f'전체 고객 발화: {len(cust)}')
print(f'짧은 발화: {len(short)} ({len(short)/len(cust):.1%})')

# 발화 내용별 분포
print(f'\n발화 내용 분포:')
top_utterances = short['발화내용'].value_counts().head(15)
for text, cnt in top_utterances.items():
    pct = cnt / len(short)
    print(f'  "{text}": {cnt}건 ({pct:.1%})')

# ══ 2. 짧은 발화별 감정 그룹 분포 ════════════════════════════════
print(f'\n{sep}\n2. 짧은 발화별 감정 그룹 분포\n{sep}')

for text in ['네', '네 네', '예', '네 네 네', '여보세요', '네 감사합니다']:
    sub = short[short['발화내용']==text]
    if len(sub) < 10:
        continue
    groups = sub['융합감정그룹'].value_counts()
    avg_v = sub['융합Valence'].mean()
    print(f'\n  "{text}" (N={len(sub)}, 평균V={avg_v:+.3f})')
    for g, c in groups.head(5).items():
        print(f'    {g}: {c}건 ({c/len(sub):.0%})')

# ══ 3. 단계별 "네"의 감정 차이 ═══════════════════════════════════
print(f'\n{sep}\n3. 단계별 "네"의 감정 차이\n{sep}')

ne_only = short[short['발화내용']=='네']
for stage in ['초기','탐색','해결시도','결과제시','종료']:
    sub = ne_only[ne_only['단계']==stage]
    if len(sub) < 5:
        continue
    avg_v = sub['융합Valence'].mean()
    groups = sub['융합감정그룹'].value_counts()
    top = groups.index[0] if len(groups) > 0 else '-'
    print(f'  {stage}: N={len(sub)}, V={avg_v:+.3f}, 최다감정={top}')

# ══ 4. 전후 맥락별 "네"의 감정 ═══════════════════════════════════
print(f'\n{sep}\n4. 전후 맥락별 "네"의 감정\n{sep}')

# "네" 직전 상담사 발화 유형별 감정 차이
ne_contexts = []
for _, row in ne_only.iterrows():
    cnid = row['CNID']
    turn_idx = row['turn_idx']
    # 직전 상담사 발화 찾기
    prev_agent = df_turns[(df_turns['CNID']==cnid) & (df_turns['turn_idx']<turn_idx) & (df_turns['화자']=='상담사')]
    if len(prev_agent) == 0:
        continue
    prev_text = str(prev_agent.iloc[-1]['발화내용'])

    # 맥락 분류
    if any(k in prev_text for k in ['죄송','불편','놀라']):
        context = '사과/공감 후'
    elif any(k in prev_text for k in ['확인','조회','알아']):
        context = '확인/조회 후'
    elif any(k in prev_text for k in ['기사','센터','방문','배송','수리']):
        context = '솔루션 안내 후'
    elif any(k in prev_text for k in ['비용','유상','금액']):
        context = '비용 안내 후'
    elif any(k in prev_text for k in ['감사','기다려']):
        context = '감사/인사 후'
    else:
        context = '기타 안내 후'

    ne_contexts.append({
        'context': context,
        'valence': row['융합Valence'],
        'group': row['융합감정그룹'],
    })

if ne_contexts:
    ctx_df = pd.DataFrame(ne_contexts)
    print(f'  "네" 맥락 분석 (N={len(ctx_df)})')
    for ctx in ['사과/공감 후','확인/조회 후','솔루션 안내 후','비용 안내 후','감사/인사 후','기타 안내 후']:
        sub = ctx_df[ctx_df['context']==ctx]
        if len(sub) < 5:
            continue
        avg_v = sub['valence'].mean()
        top_g = sub['group'].value_counts().index[0] if len(sub) > 0 else '-'
        print(f'    {ctx:>12}: N={len(sub):>4}, V={avg_v:+.3f}, 주요감정={top_g}')

# ══ 5. 음성 특징별 "네"의 감정 차이 ══════════════════════════════
print(f'\n{sep}\n5. 음성 특징별 "네"의 감정 차이\n{sep}')

# Energy 기준 상위/하위 비교
ne_with_audio = ne_only[ne_only['Energy_mean'].notna()].copy()
if len(ne_with_audio) > 20:
    median_energy = ne_with_audio['Energy_mean'].median()
    high_energy = ne_with_audio[ne_with_audio['Energy_mean'] > median_energy]
    low_energy = ne_with_audio[ne_with_audio['Energy_mean'] <= median_energy]

    print(f'  Energy 높은 "네" (밝은 톤): N={len(high_energy)}, V={high_energy["융합Valence"].mean():+.3f}')
    print(f'  Energy 낮은 "네" (어두운 톤): N={len(low_energy)}, V={low_energy["융합Valence"].mean():+.3f}')

    # F0 기준
    median_f0 = ne_with_audio['F0_mean'].median()
    high_f0 = ne_with_audio[ne_with_audio['F0_mean'] > median_f0]
    low_f0 = ne_with_audio[ne_with_audio['F0_mean'] <= median_f0]
    print(f'  F0 높은 "네" (높은 톤): N={len(high_f0)}, V={high_f0["융합Valence"].mean():+.3f}')
    print(f'  F0 낮은 "네" (낮은 톤): N={len(low_f0)}, V={low_f0["융합Valence"].mean():+.3f}')

# 엑셀 저장
output = os.path.join(os.path.dirname(os.path.abspath(__file__)), '짧은발화_분석결과.xlsx')
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    # 발화별 통계
    rows = []
    for text in top_utterances.index:
        sub = short[short['발화내용']==text]
        row = {'발화': text, 'N': len(sub), '평균V': round(sub['융합Valence'].mean(),3)}
        groups = sub['융합감정그룹'].value_counts()
        for g in ['감사/만족','안정/중립','불안/걱정','불만/짜증','혼란/당황']:
            row[g] = groups.get(g, 0)
        rows.append(row)
    pd.DataFrame(rows).to_excel(writer, sheet_name='발화별통계', index=False)

    # 맥락별 통계
    if ne_contexts:
        ctx_df.groupby('context').agg(
            N=('valence','count'),
            평균V=('valence','mean'),
        ).round(3).to_excel(writer, sheet_name='맥락별네감정')

print(f'\n엑셀 저장: {output}')
