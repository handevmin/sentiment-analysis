"""부정→긍정 / 긍정→부정 전환 트리거 분석 (686건 LLM 데이터)."""
import sys, io, warnings
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
warnings.filterwarnings('ignore')
import pandas as pd
import numpy as np
import json
from collections import Counter

df = pd.read_excel('전체_음성분석_데이터_LLM.xlsx', sheet_name='발화단위상세')

n2p_triggers = []
n2p_stages = []
n2p_patterns = []
p2n_triggers = []
p2n_stages = []
p2n_patterns = []

for cnid in df['CNID'].unique():
    ct = df[df['CNID']==cnid].sort_values('시작(초)').reset_index(drop=True)

    for i in range(1, len(ct)):
        curr = ct.iloc[i]
        if curr['화자'] != '고객' or pd.isna(curr['융합Valence']):
            continue

        prev_cust = None
        prev_agent = None
        for j in range(i-1, -1, -1):
            if ct.iloc[j]['화자'] == '고객' and prev_cust is None and pd.notna(ct.iloc[j]['융합Valence']):
                prev_cust = ct.iloc[j]
            if ct.iloc[j]['화자'] == '상담사' and prev_agent is None:
                prev_agent = ct.iloc[j]
            if prev_cust is not None and prev_agent is not None:
                break

        if prev_cust is None or prev_agent is None:
            continue

        prev_v = prev_cust['융합Valence']
        curr_v = curr['융합Valence']

        agent_text = str(prev_agent['발화내용'])

        def classify_n2p(txt):
            if any(k in txt for k in ['죄송','불편','놀라','걱정','도와']):
                return '공감/사과'
            elif any(k in txt for k in ['확인','조회','알아']):
                return '확인/조회'
            elif any(k in txt for k in ['기사','센터','방문','배송','설치','수리','교환']):
                return '구체적 솔루션'
            elif any(k in txt for k in ['안내','말씀','드리','설명']):
                return '정보 안내'
            elif any(k in txt for k in ['감사','기다려']):
                return '감사/인사'
            return '기타'

        def classify_p2n(txt):
            if any(k in txt for k in ['비용','유상','금액','요금','원']):
                return '비용 안내'
            elif any(k in txt for k in ['지연','소요','대기','시간','기다']):
                return '지연/대기'
            elif any(k in txt for k in ['불가','안 되','어렵','힘들','없']):
                return '불가/제한'
            elif any(k in txt for k in ['확인','조회','알아']):
                return '확인 요청'
            elif any(k in txt for k in ['혹시','양해','참고']):
                return '주의사항 안내'
            return '기타'

        # 부정→긍정
        if prev_v <= -0.1 and curr_v >= 0.05:
            n2p_triggers.append(agent_text)
            n2p_stages.append(curr['단계'])
            n2p_patterns.append(classify_n2p(agent_text))

        # 긍정→부정
        elif prev_v >= 0.05 and curr_v <= -0.1:
            p2n_triggers.append(agent_text)
            p2n_stages.append(curr['단계'])
            p2n_patterns.append(classify_p2n(agent_text))

sep = '='*60
print(sep)
print(f'부정 → 긍정 전환 분석 (N={len(n2p_triggers)}건)')
print(sep)

print('\n[전환 트리거 상담사 발화 패턴]')
for pat, cnt in Counter(n2p_patterns).most_common():
    print(f'  {pat}: {cnt}건 ({cnt/len(n2p_patterns):.0%})')

print('\n[전환 발생 단계]')
for stg, cnt in Counter(n2p_stages).most_common():
    print(f'  {stg}: {cnt}건 ({cnt/len(n2p_stages):.0%})')

print('\n[대표 상담사 발화 — 부정→긍정 직전]')
for i in range(min(10, len(n2p_triggers))):
    print(f'  [{n2p_patterns[i]}] {n2p_triggers[i][:70]}')

print(f'\n{sep}')
print(f'긍정 → 부정 전환 분석 (N={len(p2n_triggers)}건)')
print(sep)

print('\n[전환 트리거 상담사 발화 패턴]')
for pat, cnt in Counter(p2n_patterns).most_common():
    print(f'  {pat}: {cnt}건 ({cnt/len(p2n_patterns):.0%})')

print('\n[전환 발생 단계]')
for stg, cnt in Counter(p2n_stages).most_common():
    print(f'  {stg}: {cnt}건 ({cnt/len(p2n_stages):.0%})')

print('\n[대표 상담사 발화 — 긍정→부정 직전]')
for i in range(min(8, len(p2n_triggers))):
    print(f'  [{p2n_patterns[i]}] {p2n_triggers[i][:70]}')

# JSON 저장
result = {
    'n2p_total': len(n2p_triggers),
    'n2p_patterns': dict(Counter(n2p_patterns).most_common()),
    'n2p_stages': dict(Counter(n2p_stages).most_common()),
    'p2n_total': len(p2n_triggers),
    'p2n_patterns': dict(Counter(p2n_patterns).most_common()),
    'p2n_stages': dict(Counter(p2n_stages).most_common()),
}
with open('outputs/data/transition_analysis.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)
print(f'\nJSON 저장: outputs/data/transition_analysis.json')
