"""LLM 교차검증 완료 데이터 종합 분석."""
import sys, io, warnings, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
warnings.filterwarnings('ignore')
import pandas as pd
import numpy as np

df_calls = pd.read_excel('전체_음성분석_데이터_LLM.xlsx', sheet_name='콜단위요약')
df_turns = pd.read_excel('전체_음성분석_데이터_LLM.xlsx', sheet_name='발화단위상세')
cust = df_turns[df_turns['화자']=='고객'].copy()
agent = df_turns[df_turns['화자']=='상담사'].copy()
sep = '='*60

print(sep); print('1. 기본 통계'); print(sep)
print(f'총 콜: {len(df_calls)}')
print(f'총 발화: {len(df_turns)} (고객 {len(cust)}, 상담사 {len(agent)})')
print(f'콜당 평균 발화: {len(df_turns)/len(df_calls):.1f}')
print(f'평균 통화시간: {df_calls["통화시간(초)"].mean():.0f}초 ({df_calls["통화시간(초)"].mean()/60:.1f}분)')
print(f'고객 평균 발화 길이: {cust["구간길이(초)"].mean():.1f}초')
print(f'상담사 평균 발화 길이: {agent["구간길이(초)"].mean():.1f}초')

print(f'\n{sep}'); print('2. 짧은 발화'); print(sep)
short = cust[cust['짧은발화']==True]
meaningful = cust[cust['짧은발화']!=True]
print(f'짧은 발화: {len(short)} ({len(short)/len(cust):.1%})')
print(f'의미있는: {len(meaningful)} ({len(meaningful)/len(cust):.1%})')

print(f'\n{sep}'); print('3. 감정 그룹 분포 (LLM 후)'); print(sep)
gd = cust['융합감정그룹'].value_counts()
for g, c in gd.items():
    print(f'  {g}: {c}건 ({c/len(cust):.1%})')

print(f'\n{sep}'); print('4. STT Only vs STT+Audio 변경'); print(sep)
both = cust[(cust['STT_Only그룹'].notna()) & (cust['융합감정그룹'].notna())]
changed = both[both['STT_Only그룹'] != both['융합감정그룹']]
print(f'비교 가능: {len(both)}, 변경: {len(changed)} ({len(changed)/len(both):.1%})')
cp = changed.groupby(['STT_Only그룹','융합감정그룹']).size().sort_values(ascending=False)
for (fr, to), cnt in cp.head(10).items():
    print(f'  {fr} -> {to}: {cnt}건')

print(f'\n{sep}'); print('5. 단계별 감성'); print(sep)
stage_data = {}
for stage in ['초기','탐색','해결시도','결과제시','종료']:
    st = cust[cust['단계']==stage]['융합Valence'].dropna()
    stage_data[stage] = {'mean': round(float(st.mean()),4), 'std': round(float(st.std()),4), 'n': int(len(st))}
    print(f'  {stage}: mean={st.mean():+.4f}, std={st.std():.4f}, n={len(st)}')

print(f'\n{sep}'); print('6. NPS vs 감성'); print(sep)
call_v = cust.groupby('CNID')['융합Valence'].mean().reset_index()
call_v.columns = ['CNID','avg_v']
mg = df_calls.merge(call_v, on='CNID', how='left')
valid = mg[['NPS','avg_v']].dropna()
corr = valid['NPS'].corr(valid['avg_v'])
print(f'Pearson r = {corr:.3f} (n={len(valid)})')
for lo,hi,lb in [(1,5,'비추천'),(6,8,'중립'),(9,10,'추천')]:
    s = valid[(valid['NPS']>=lo)&(valid['NPS']<=hi)]
    if len(s)>0: print(f'  {lb} (NPS {lo}-{hi}): valence={s["avg_v"].mean():+.4f}, n={len(s)}')

print(f'\n{sep}'); print('7. 성별'); print(sep)
for g in ['남성','여성']:
    cn = df_calls[df_calls['성별']==g]['CNID']
    v = cust[cust['CNID'].isin(cn)]['융합Valence'].dropna()
    n = df_calls[df_calls['성별']==g]['NPS'].mean()
    print(f'  {g}: NPS={n:.2f}, 감성={v.mean():+.4f}, n={len(cn)}')

print(f'\n{sep}'); print('8. 연령대'); print(sep)
for a in ['20~39세','40~49세','50~64세','65~74세','75세 이상']:
    cn = df_calls[df_calls['연령대']==a]['CNID']
    if len(cn)==0: continue
    v = cust[cust['CNID'].isin(cn)]['융합Valence'].dropna()
    n = df_calls[df_calls['연령대']==a]['NPS'].mean()
    print(f'  {a}: NPS={n:.2f}, 감성={v.mean():+.4f}, n={len(cn)}')

print(f'\n{sep}'); print('9. 제품군'); print(sep)
for p in df_calls['제품대'].dropna().value_counts().head(6).index:
    cn = df_calls[df_calls['제품대']==p]['CNID']
    v = cust[cust['CNID'].isin(cn)]['융합Valence'].dropna()
    n = df_calls[df_calls['제품대']==p]['NPS'].mean()
    print(f'  {p}: NPS={n:.2f}, 감성={v.mean():+.4f}, n={len(cn)}')

print(f'\n{sep}'); print('10. 감정 전환'); print(sep)
tr = 0; n2p = 0; p2n = 0
for cnid in cust['CNID'].unique():
    ct = cust[cust['CNID']==cnid].sort_values('시작(초)')
    fv = ct['융합Valence'].dropna().values
    for i in range(len(fv)-1):
        d = fv[i+1]-fv[i]
        if abs(d)>=0.15:
            tr += 1
            if fv[i]<-0.05 and fv[i+1]>0.05: n2p += 1
            elif fv[i]>0.05 and fv[i+1]<-0.05: p2n += 1
print(f'  총 전환: {tr}, 콜당 평균: {tr/len(df_calls):.1f}회')
print(f'  부정->긍정: {n2p} ({n2p/max(tr,1):.0%})')
print(f'  긍정->부정: {p2n} ({p2n/max(tr,1):.0%})')

print(f'\n{sep}'); print('11. 상담사 응대 패턴'); print(sep)
ek = ['죄송','불편','놀라','걱정','도와','감사','기다려']
pk = ['제가','확인해','조회해','안내','연락','처리','접수']
sk = ['센터','기사','교환','환불','수리','배송','설치']
e = agent['발화내용'].astype(str).apply(lambda x: any(k in x for k in ek)).sum()
p = agent['발화내용'].astype(str).apply(lambda x: any(k in x for k in pk)).sum()
s = agent['발화내용'].astype(str).apply(lambda x: any(k in x for k in sk)).sum()
print(f'  공감 멘트: {e}/{len(agent)} ({e/len(agent):.0%})')
print(f'  선제적 안내: {p}/{len(agent)} ({p/len(agent):.0%})')
print(f'  구체적 솔루션: {s}/{len(agent)} ({s/len(agent):.0%})')
print(f'\n  NPS별 상담사 공감률:')
for lo,hi,lb in [(1,5,'비추천'),(6,8,'중립'),(9,10,'추천')]:
    cn = df_calls[(df_calls['NPS']>=lo)&(df_calls['NPS']<=hi)]['CNID']
    sa = agent[agent['CNID'].isin(cn)]
    if len(sa)>0:
        er = sa['발화내용'].astype(str).apply(lambda x: any(k in x for k in ek)).mean()
        pr = sa['발화내용'].astype(str).apply(lambda x: any(k in x for k in pk)).mean()
        print(f'    {lb}: 공감={er:.0%}, 선제={pr:.0%}')

# JSON 저장
stats = {
    'n_calls': int(len(df_calls)), 'n_turns': int(len(df_turns)),
    'n_customer': int(len(cust)), 'short': int(len(short)),
    'short_pct': round(len(short)/len(cust),3),
    'changed': int(len(changed)), 'change_rate': round(len(changed)/max(len(both),1),3),
    'avg_duration': round(float(df_calls['통화시간(초)'].mean()),1),
    'nps_corr': round(float(corr),3),
    'groups': {k:int(v) for k,v in gd.items()},
    'stages': stage_data,
    'transitions': tr, 'n2p': n2p, 'p2n': p2n,
}
with open('outputs/data/full_stats_llm.json', 'w', encoding='utf-8') as f:
    json.dump(stats, f, ensure_ascii=False, indent=2)
print(f'\nJSON: outputs/data/full_stats_llm.json')
