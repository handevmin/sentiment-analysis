# -*- coding: utf-8 -*-
"""
상담사 응대 패턴과 고객 감정 변화 분석
- 부정 고객 발화 → 상담사 응대 → 다음 고객 발화의 3-turn 시퀀스 분석
- 응대 유형별, 응답 길이별, 단계별 감정 변화량 측정
- 부정 강도별 응대 효과 분석
"""
import sys, io, warnings, os
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
warnings.filterwarnings('ignore')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import pandas as pd
import numpy as np

EXPORTS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputs', 'exports')
df_turns = pd.read_csv(f'{EXPORTS}/전체_음성분석_데이터_LLM_발화단위상세.csv', encoding='utf-8-sig')
df_calls = pd.read_csv(f'{EXPORTS}/전체_음성분석_데이터_LLM_콜단위요약.csv', encoding='utf-8-sig')
df_calls['CNID'] = df_calls['CNID'].astype(str).str.strip()

# 키워드 사전
empathy_kw = ['죄송','불편','놀라','걱정','도와','이해','힘드','어려','맞으시']
resolution_kw = ['확인','조회','처리','접수','안내','해드','드리','조치']
schedule_kw = ['기사','센터','방문','배송','수리','내일','오전','오후']
cost_kw = ['비용','유상','금액','요금']
alt_kw = ['대신','다른','방법','대안','그러면']

sep = '='*60

# ══ 3-turn 시퀀스 구축 ════════════════════════════════════════════
print(f'{sep}\n3-turn 시퀀스 분석: 부정 고객 → 상담사 → 다음 고객\n{sep}')

sequences = []
for cnid in df_turns['CNID'].unique():
    sub = df_turns[df_turns['CNID']==cnid].sort_values('시작(초)').reset_index(drop=True)

    for i in range(len(sub)):
        if sub.iloc[i]['화자'] != '고객':
            continue
        v_before = sub.iloc[i]['융합Valence']
        if pd.isna(v_before) or v_before >= -0.05:
            continue

        next_agents = sub[(sub.index > i) & (sub['화자']=='상담사')]
        if len(next_agents) == 0:
            continue
        agent_row = next_agents.iloc[0]
        agent_text = str(agent_row['발화내용'])

        next_custs = sub[(sub.index > agent_row.name) & (sub['화자']=='고객')]
        if len(next_custs) == 0:
            continue
        v_after = next_custs.iloc[0]['융합Valence']
        if pd.isna(v_after):
            continue

        has_empathy = any(k in agent_text for k in empathy_kw)
        has_resolution = any(k in agent_text for k in resolution_kw)
        has_schedule = any(k in agent_text for k in schedule_kw)
        has_cost = any(k in agent_text for k in cost_kw)

        if has_empathy and has_resolution:
            resp_type = '공감+해결안내'
        elif has_empathy:
            resp_type = '공감만'
        elif has_schedule:
            resp_type = '일정/방문안내'
        elif has_resolution:
            resp_type = '해결안내만'
        elif has_cost:
            resp_type = '비용안내'
        elif len(agent_text) <= 10:
            resp_type = '짧은응답'
        else:
            resp_type = '일반안내'

        sequences.append({
            'cnid': cnid,
            'v_before': v_before,
            'v_after': v_after,
            'delta': v_after - v_before,
            'resp_type': resp_type,
            'agent_len': len(agent_text),
            'stage': sub.iloc[i]['단계'],
            'has_empathy': has_empathy,
            'has_resolution': has_resolution,
        })

sdf = pd.DataFrame(sequences)
print(f'시퀀스 총 {len(sdf)}건\n')

# ══ 1. 응대 유형별 감정 변화 ══════════════════════════════════════
print(f'{sep}\n1. 상담사 응대 유형별 고객 감정 변화\n{sep}')
print(f'{"응대 유형":>14} {"N":>5} {"평균변화":>8} {"개선율":>7} {"평균전":>7} {"평균후":>7}')
for rtype in ['공감+해결안내','공감만','해결안내만','일정/방문안내','비용안내','짧은응답','일반안내']:
    sub = sdf[sdf['resp_type']==rtype]
    if len(sub) < 10:
        continue
    print(f'{rtype:>14} {len(sub):>5} {sub["delta"].mean():>+8.3f} '
          f'{(sub["delta"]>0).mean():>6.0%} {sub["v_before"].mean():>+7.3f} {sub["v_after"].mean():>+7.3f}')

# ══ 2. 응답 길이별 효과 ══════════════════════════════════════════
print(f'\n{sep}\n2. 상담사 응답 길이별 감정 변화\n{sep}')
for label, lo, hi in [('짧은(<15자)',0,15), ('중간(15~40자)',15,40), ('긴(40~80자)',40,80), ('상세(80자+)',80,9999)]:
    sub = sdf[(sdf['agent_len'] >= lo) & (sdf['agent_len'] < hi)]
    if len(sub) < 10:
        continue
    print(f'  {label:>14}: N={len(sub):>4}, 변화={sub["delta"].mean():+.3f}, 개선율={(sub["delta"]>0).mean():.0%}')

# ══ 3. 부정 강도별 응대 효과 ═════════════════════════════════════
print(f'\n{sep}\n3. 부정 강도별 상담사 응대 효과\n{sep}')
for label, lo, hi in [('경미(-0.05~-0.15)',-0.15,-0.05), ('중간(-0.15~-0.3)',-0.30,-0.15),
                       ('심각(-0.3~-0.5)',-0.50,-0.30), ('극심(-0.5이하)',-99,-0.50)]:
    sub = sdf[(sdf['v_before'] >= lo) & (sdf['v_before'] < hi)]
    if len(sub) < 20:
        continue
    print(f'\n  [{label}] N={len(sub)}, 평균변화={sub["delta"].mean():+.3f}')
    for rtype in ['해결안내만','공감+해결안내','공감만','짧은응답']:
        rsub = sub[sub['resp_type']==rtype]
        if len(rsub) >= 5:
            print(f'    {rtype:>12}: N={len(rsub):>3}, 변화={rsub["delta"].mean():+.3f}, 개선율={(rsub["delta"]>0).mean():.0%}')

# ══ 4. 단계×응대유형 교차 효과 ═══════════════════════════════════
print(f'\n{sep}\n4. 단계×응대유형 교차 효과\n{sep}')
for stage in ['탐색','해결시도','결과제시']:
    sub = sdf[sdf['stage']==stage]
    if len(sub) < 30:
        continue
    print(f'\n  [{stage}]')
    empathy = sub[sub['has_empathy'] & ~sub['has_resolution']]
    resolution = sub[~sub['has_empathy'] & sub['has_resolution']]
    both = sub[sub['has_empathy'] & sub['has_resolution']]
    short = sub[~sub['has_empathy'] & ~sub['has_resolution'] & (sub['agent_len'] <= 10)]
    general = sub[~sub['has_empathy'] & ~sub['has_resolution'] & (sub['agent_len'] > 10)]

    for lbl, s in [('공감만',empathy),('해결만',resolution),('공감+해결',both),('짧은응답',short),('일반안내',general)]:
        if len(s) >= 10:
            print(f'    {lbl:>8}: N={len(s):>4}, 변화={s["delta"].mean():+.3f}, 개선율={(s["delta"]>0).mean():.0%}')

# ══ 5. 감정 회복률 (기존 분석 유지) ════════════════════════════════
print(f'\n{sep}\n5. 전체 감정 회복률\n{sep}')
cust = df_turns[df_turns['화자']=='고객'].copy()
recovery_count = 0
no_recovery_count = 0
for cnid in cust['CNID'].unique():
    fv = cust[cust['CNID']==cnid]['융합Valence'].dropna().values
    if len(fv) < 3:
        continue
    has_neg = any(v < -0.1 for v in fv)
    if not has_neg:
        continue
    neg_idx = next(i for i, v in enumerate(fv) if v < -0.1)
    recovered = any(v > 0.05 for v in fv[neg_idx+1:]) if neg_idx < len(fv)-1 else False
    if recovered:
        recovery_count += 1
    else:
        no_recovery_count += 1

total = recovery_count + no_recovery_count
print(f'부정 구간 있는 콜: {total}건')
print(f'  감정 회복: {recovery_count}건 ({recovery_count/total:.0%})')
print(f'  회복 안됨: {no_recovery_count}건 ({no_recovery_count/total:.0%})')

# ══ 6. 해결 불가 후 회복 속도 ════════════════════════════════════
print(f'\n{sep}\n6. 해결 불가 안내 후 회복 속도\n{sep}')
fail_kw = ['불가','안 되','어렵','없','못']
fast = slow = no_rec = 0
for cnid in cust['CNID'].unique():
    ct = df_turns[df_turns['CNID']==cnid].sort_values('시작(초)').reset_index(drop=True)
    fail_turns = ct[(ct['화자']=='상담사') & ct['발화내용'].astype(str).apply(lambda x: any(k in x for k in fail_kw))]
    if len(fail_turns) == 0:
        continue
    fail_idx = fail_turns.iloc[0].name
    after_c = ct[(ct.index > fail_idx) & (ct['화자']=='고객')]
    fv_a = after_c['융합Valence'].dropna().values
    if len(fv_a) == 0:
        continue
    rec_at = None
    for j, v in enumerate(fv_a):
        if v > 0.0:
            rec_at = j
            break
    if rec_at is not None:
        if rec_at <= 2: fast += 1
        else: slow += 1
    else:
        no_rec += 1

tf = fast + slow + no_rec
if tf > 0:
    print(f'해결 불가 안내 포함 콜: {tf}건')
    print(f'  빠른 회복 (2턴 이내): {fast}건 ({fast/tf:.0%})')
    print(f'  느린 회복 (3턴 이상): {slow}건 ({slow/tf:.0%})')
    print(f'  회복 안됨:           {no_rec}건 ({no_rec/tf:.0%})')

# ══ 엑셀 저장 ════════════════════════════════════════════════════
output = os.path.join(os.path.dirname(os.path.abspath(__file__)), '상담사응대패턴_분석결과.xlsx')
with pd.ExcelWriter(output, engine='openpyxl') as writer:
    # 응대 유형별
    rows = []
    for rtype in ['공감+해결안내','공감만','해결안내만','일정/방문안내','비용안내','짧은응답','일반안내']:
        sub = sdf[sdf['resp_type']==rtype]
        if len(sub) < 10:
            continue
        rows.append({
            '응대유형': rtype, 'N': len(sub),
            '평균변화': round(sub['delta'].mean(), 3),
            '개선율': f'{(sub["delta"]>0).mean():.0%}',
        })
    pd.DataFrame(rows).to_excel(writer, sheet_name='응대유형별', index=False)

    # 응답 길이별
    rows = []
    for label, lo, hi in [('짧은(<15자)',0,15),('중간(15~40자)',15,40),('긴(40~80자)',40,80),('상세(80자+)',80,9999)]:
        sub = sdf[(sdf['agent_len'] >= lo) & (sdf['agent_len'] < hi)]
        if len(sub) >= 10:
            rows.append({'길이': label, 'N': len(sub), '평균변화': round(sub['delta'].mean(),3), '개선율': f'{(sub["delta"]>0).mean():.0%}'})
    pd.DataFrame(rows).to_excel(writer, sheet_name='응답길이별', index=False)

print(f'\n엑셀 저장: {output}')
