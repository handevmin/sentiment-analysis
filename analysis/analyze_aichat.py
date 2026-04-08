"""
AI Chat 데이터 종합 분석 + 음성 상담 비교
- 세션 단위 집계
- 음성 상담 데이터와 비교 분석
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
from datetime import datetime

# ── 데이터 로드 ──────────────────────────────────────────────────
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
df_raw = pd.read_csv(os.path.join(BASE, 'AI_CHAT_Data_Sheet0.csv'), encoding='utf-8-sig')

# 첫 행(컬럼 설명) 제거
df = df_raw[df_raw['SESSION_ID'] != '세션ID'].copy()
df['만족도점수'] = pd.to_numeric(df['만족도점수'], errors='coerce')

# ── 테스터 제외 ──────────────────────────────────────────────────
TESTERS = {
    '2fd5d28f-9102-4f29-b891-690fccca4f7a',
    'acf40b3d-8aeb-42fd-908f-e7a82d281a64',
    '587a3023-9849-4dc7-8e49-5868c2e61957',
    '72109591-26b5-45d8-9fdd-1d10e3e699c8',
    'c40c91f4-febe-4dd7-a055-1722d64f4993',
    '51700104-3903-480b-9a1e-98b4d956ff53',
    '009413ea-e3e8-4f00-b2f2-75e4799cb747',
    '71f651d4-4175-47bf-8557-a7b38e7a9ce6',
    '1d37b58a-a821-47af-990f-793ae51374dc',
    'fdba49d2-e6b3-4768-a112-ed7651eec7a6',
    '1f856d75-7645-4039-97b0-3dfe9a287770',
    'e7efaef1-c164-4a13-9250-083995197abf',
    'a2b2321d-987a-4792-93ce-cc4f25e26600',
    '207393d9-6260-4cc8-b9ed-ac3e9bde11ba',
    '831684b5-bc13-49cc-b4c2-5fb28caf5467',
    'ee76bdbb-4a48-4748-a53a-0326033fe431',
    'fe908d4a-09a2-4bda-b626-48d38d608636',
    '1721c54c-74a4-42c4-984b-7e4c0e0e3928',
    '18bb4745-b11b-4ef5-bcc4-be907552d247',
}

before = len(df)
df = df[~df['SESSION_ID'].isin(TESTERS)].reset_index(drop=True)
print(f'테스터 제거: {before} → {len(df)} ({before - len(df)}건 제거)')

# ── 시간 파싱 ────────────────────────────────────────────────────
df['발화일자'] = pd.to_datetime(df['발화일자'], errors='coerce')

sep = '=' * 60

# ══════════════════════════════════════════════════════════════════
# 1. 기본 통계
# ══════════════════════════════════════════════════════════════════
print(f'\n{sep}')
print('1. AI Chat 기본 통계')
print(sep)

n_sessions = df['SESSION_ID'].nunique()
n_turns = len(df)
print(f'세션 수: {n_sessions}')
print(f'총 발화(턴) 수: {n_turns}')
print(f'세션당 평균 턴: {n_turns / n_sessions:.1f}')

# 세션별 턴 수 분포
session_turns = df.groupby('SESSION_ID').size()
print(f'세션당 턴 수: min={session_turns.min()}, max={session_turns.max()}, median={session_turns.median():.0f}')

# 만족도
print(f'\n만족도 분포:')
sat = df.groupby('SESSION_ID')['만족도점수'].first().dropna()
print(f'  평균: {sat.mean():.2f}')
print(f'  분포:')
for score in sorted(sat.unique()):
    cnt = (sat == score).sum()
    print(f'    {score:.0f}점: {cnt}건 ({cnt/len(sat):.0%})')

# ══════════════════════════════════════════════════════════════════
# 2. 고객 질문 패턴 분석
# ══════════════════════════════════════════════════════════════════
print(f'\n{sep}')
print('2. 고객 질문 패턴 분석')
print(sep)

questions = df['QUESTION'].dropna().astype(str)
q_lengths = questions.str.len()
print(f'질문 평균 길이: {q_lengths.mean():.1f}자')
print(f'질문 중앙값 길이: {q_lengths.median():.0f}자')

# 키워드 vs 서사 분석
keyword_patterns = 0  # 모델명, 제품코드 포함
narrative_patterns = 0  # 서사적 표현
model_re_count = 0  # 모델명 직접 입력

for q in questions:
    q_lower = q.lower()
    # 모델명 패턴 (영문+숫자 조합, 6자 이상)
    import re
    if re.search(r'[A-Z]{1,5}[\-]?\d{3,}', q, re.IGNORECASE):
        model_re_count += 1
        keyword_patterns += 1
    elif any(kw in q for kw in ['어제','오늘','그래서','근데','했는데','있었','가지고']):
        narrative_patterns += 1
    elif len(q) < 15:
        keyword_patterns += 1

print(f'모델명 직접 입력: {model_re_count}건 ({model_re_count/len(questions):.1%})')
print(f'키워드형 질문: {keyword_patterns}건 ({keyword_patterns/len(questions):.1%})')
print(f'서사형 질문: {narrative_patterns}건 ({narrative_patterns/len(questions):.1%})')

# INPUT_TYPE 분석
print(f'\n입력 방식:')
for itype, cnt in df['INPUT_TYPE'].value_counts().items():
    print(f'  {itype}: {cnt}건 ({cnt/len(df):.0%})')

# ══════════════════════════════════════════════════════════════════
# 3. INTENT (상담 주제) 분석
# ══════════════════════════════════════════════════════════════════
print(f'\n{sep}')
print('3. 상담 주제 (INTENT) 분석')
print(sep)

intent_counts = df['INTENT_CD'].value_counts()
print('상위 15개:')
for intent, cnt in intent_counts.head(15).items():
    print(f'  {intent}: {cnt}건 ({cnt/len(df):.1%})')

# INTENT별 만족도
print(f'\nINTENT별 평균 만족도 (상위 10):')
intent_sat = df.groupby('INTENT_CD')['만족도점수'].mean().sort_values()
top_intents = df['INTENT_CD'].value_counts().head(10).index
for intent in top_intents:
    if intent in intent_sat.index:
        cnt = intent_counts[intent]
        avg = intent_sat[intent]
        print(f'  {intent}: {avg:.2f} (N={cnt})')

# ══════════════════════════════════════════════════════════════════
# 4. 세션 단위 집계
# ══════════════════════════════════════════════════════════════════
print(f'\n{sep}')
print('4. 세션 단위 집계')
print(sep)

session_df = df.groupby('SESSION_ID').agg(
    만족도=('만족도점수', 'first'),
    피드백=('만족도피드백', 'first'),
    턴수=('SESSION_ID', 'size'),
    시작시간=('발화일자', 'min'),
    종료시간=('발화일자', 'max'),
    직접입력비율=('INPUT_TYPE', lambda x: (x == 'CONVERSATION').mean()),
    주요인텐트=('INTENT_CD', lambda x: x.mode().iloc[0] if len(x.mode()) > 0 else ''),
    평균질문길이=('QUESTION', lambda x: x.astype(str).str.len().mean()),
    모델명포함=('QUESTION', lambda x: sum(1 for q in x.astype(str) if re.search(r'[A-Z]{1,5}[\-]?\d{3,}', q, re.IGNORECASE)) > 0),
).reset_index()

# 세션 소요 시간
session_df['소요시간(초)'] = (session_df['종료시간'] - session_df['시작시간']).dt.total_seconds()
session_df['소요시간(초)'] = session_df['소요시간(초)'].clip(lower=0)

print(f'세션 수: {len(session_df)}')
print(f'평균 턴 수: {session_df["턴수"].mean():.1f}')
print(f'평균 소요 시간: {session_df["소요시간(초)"].mean():.0f}초 ({session_df["소요시간(초)"].mean()/60:.1f}분)')
print(f'직접 입력 비율 평균: {session_df["직접입력비율"].mean():.0%}')
print(f'모델명 포함 세션: {session_df["모델명포함"].sum()} ({session_df["모델명포함"].mean():.0%})')

# ══════════════════════════════════════════════════════════════════
# 5. 만족도별 세그먼트 분석
# ══════════════════════════════════════════════════════════════════
print(f'\n{sep}')
print('5. 만족도별 세그먼트')
print(sep)

for lo, hi, label in [(1,5,'불만족(1-5)'),(6,8,'보통(6-8)'),(9,10,'만족(9-10)')]:
    sub = session_df[(session_df['만족도']>=lo)&(session_df['만족도']<=hi)]
    if len(sub) == 0:
        continue
    print(f'\n  [{label}] N={len(sub)}')
    print(f'    평균 턴: {sub["턴수"].mean():.1f}')
    print(f'    평균 질문 길이: {sub["평균질문길이"].mean():.1f}자')
    print(f'    직접 입력 비율: {sub["직접입력비율"].mean():.0%}')
    print(f'    모델명 포함: {sub["모델명포함"].mean():.0%}')
    print(f'    평균 소요 시간: {sub["소요시간(초)"].mean():.0f}초')

# ══════════════════════════════════════════════════════════════════
# 6. 질문-답변 간격 분석
# ══════════════════════════════════════════════════════════════════
print(f'\n{sep}')
print('6. 질문-답변 간격 분석')
print(sep)

# 세션 내 연속 턴 간 시간 간격
intervals = []
for sid in df['SESSION_ID'].unique():
    st = df[df['SESSION_ID']==sid].sort_values('발화일자')
    times = st['발화일자'].dropna()
    if len(times) < 2:
        continue
    diffs = times.diff().dt.total_seconds().dropna()
    intervals.extend(diffs.values)

intervals = [i for i in intervals if 0 < i < 3600]  # 1시간 이내만
if intervals:
    print(f'턴 간 평균 간격: {np.mean(intervals):.1f}초 ({np.mean(intervals)/60:.1f}분)')
    print(f'턴 간 중앙값 간격: {np.median(intervals):.1f}초')
    print(f'1분 이내: {sum(1 for i in intervals if i<=60)/len(intervals):.0%}')
    print(f'5분 이상: {sum(1 for i in intervals if i>=300)/len(intervals):.0%}')

# ══════════════════════════════════════════════════════════════════
# 7. 피드백 텍스트 분석
# ══════════════════════════════════════════════════════════════════
print(f'\n{sep}')
print('7. 만족도 피드백 분석')
print(sep)

feedback = session_df['피드백'].dropna()
feedback_counts = feedback.value_counts()
for fb, cnt in feedback_counts.head(10).items():
    print(f'  {fb}: {cnt}건 ({cnt/len(feedback):.0%})')

# ══════════════════════════════════════════════════════════════════
# 8. 음성 상담과 비교 데이터 준비
# ══════════════════════════════════════════════════════════════════
print(f'\n{sep}')
print('8. 음성 상담 vs AI Chat 비교')
print(sep)

# 음성 상담 데이터 로드
voice_calls = pd.read_excel(os.path.join(BASE, 'analysis', '전체_음성분석_데이터_LLM.xlsx'), sheet_name='콜단위요약')
voice_turns = pd.read_excel(os.path.join(BASE, 'analysis', '전체_음성분석_데이터_LLM.xlsx'), sheet_name='발화단위상세')

voice_cust = voice_turns[voice_turns['화자']=='고객']
voice_agent = voice_turns[voice_turns['화자']=='상담사']

# 비교 항목 산출
comparison = {
    '지표': [],
    '음성 상담': [],
    'AI Chat': [],
    '차이/비고': [],
}

def add_comp(metric, voice_val, chat_val, note=''):
    comparison['지표'].append(metric)
    comparison['음성 상담'].append(voice_val)
    comparison['AI Chat'].append(chat_val)
    comparison['차이/비고'].append(note)

# 기본 규모
add_comp('분석 건수', f'{len(voice_calls)}건', f'{len(session_df)}세션', '')
add_comp('총 발화(턴) 수', f'{len(voice_turns):,}', f'{len(df):,}', '')
add_comp('건당 평균 턴', f'{len(voice_turns)/len(voice_calls):.1f}', f'{session_df["턴수"].mean():.1f}', '')

# 만족도
voice_nps = voice_calls['NPS'].mean()
chat_sat = session_df['만족도'].mean()
add_comp('평균 만족도', f'{voice_nps:.2f}', f'{chat_sat:.2f}', '')

# 만족도 분포
voice_high = (voice_calls['NPS']>=9).mean()
chat_high = (session_df['만족도']>=9).mean()
add_comp('만족(9-10점) 비율', f'{voice_high:.0%}', f'{chat_high:.0%}', '')

voice_low = (voice_calls['NPS']<=5).mean()
chat_low = (session_df['만족도']<=5).mean()
add_comp('불만족(1-5점) 비율', f'{voice_low:.0%}', f'{chat_low:.0%}', '')

# 소요 시간
voice_dur = voice_calls['통화시간(초)'].mean()
chat_dur = session_df['소요시간(초)'].mean()
add_comp('평균 소요 시간', f'{voice_dur:.0f}초 ({voice_dur/60:.1f}분)', f'{chat_dur:.0f}초 ({chat_dur/60:.1f}분)', '')

# 고객 발화 길이
voice_cust_len = voice_cust['발화내용'].astype(str).str.len().mean()
chat_q_len = questions.str.len().mean()
add_comp('고객 발화 평균 길이', f'{voice_cust_len:.1f}자', f'{chat_q_len:.1f}자', '')

# 상담사/봇 응답 길이
voice_agent_len = voice_agent['발화내용'].astype(str).str.len().mean()
chat_a_len = df['ANSWER'].astype(str).str.len().mean()
add_comp('응답 평균 길이', f'{voice_agent_len:.1f}자', f'{chat_a_len:.1f}자', '')

# 짧은 응답 비율 (음성)
voice_short = (voice_cust['짧은발화']==True).mean()
add_comp('짧은 응답 비율 (고객)', f'{voice_short:.0%}', 'N/A (텍스트)', '음성에서만 해당')

# 모델명 포함
add_comp('모델명 직접 입력 비율', 'N/A (묘사 중심)', f'{session_df["모델명포함"].mean():.0%}', 'AI Chat에서 키워드 중심')

comp_df = pd.DataFrame(comparison)
print(comp_df.to_string(index=False))

# ══════════════════════════════════════════════════════════════════
# 9. 질적 차이 분석
# ══════════════════════════════════════════════════════════════════
print(f'\n{sep}')
print('9. 채널별 질적 차이 분석')
print(sep)

# 음성: 공감 멘트 비율
empathy_kw = ['죄송','불편','놀라','걱정','도와','감사','기다려']
voice_empathy = voice_agent['발화내용'].astype(str).apply(lambda x: any(k in x for k in empathy_kw)).mean()

# AI Chat: 공감 표현 비율
chat_empathy_kw = ['죄송','불편','도움','감사','양해','안타']
chat_empathy = df['ANSWER'].astype(str).apply(lambda x: any(k in x for k in chat_empathy_kw)).mean()

print(f'공감 표현 비율:')
print(f'  음성 상담사: {voice_empathy:.0%}')
print(f'  AI Chat: {chat_empathy:.0%}')

# 선제적 안내
voice_proactive = voice_agent['발화내용'].astype(str).apply(lambda x: any(k in x for k in ['제가','확인해','조회해','안내','처리'])).mean()
chat_proactive = df['ANSWER'].astype(str).apply(lambda x: any(k in x for k in ['추천','안내','확인','참고'])).mean()
print(f'\n선제적/정보제공 비율:')
print(f'  음성 상담사: {voice_proactive:.0%}')
print(f'  AI Chat: {chat_proactive:.0%}')

# 고객 질문 유형
print(f'\n고객 질문 유형:')
# 음성: 서사적
voice_narrative = voice_cust['발화내용'].astype(str).apply(lambda x: any(k in x for k in ['어제','오늘','그래서','근데','했는데','있었','가지고'])).mean()
chat_narrative = questions.apply(lambda x: any(k in x for k in ['어제','오늘','그래서','근데','했는데','있었','가지고'])).mean()
print(f'  서사적 표현: 음성 {voice_narrative:.0%} / AI Chat {chat_narrative:.0%}')

# 구체적 스펙 질문
voice_spec = voice_cust['발화내용'].astype(str).apply(lambda x: any(k in x for k in ['CPU','RAM','메모리','해상도','인치','용량'])).mean()
chat_spec = questions.apply(lambda x: any(k in x for k in ['CPU','RAM','메모리','해상도','인치','용량'])).mean()
print(f'  스펙/기술 질문: 음성 {voice_spec:.0%} / AI Chat {chat_spec:.0%}')

# 감정 표현
voice_emotion = voice_cust['발화내용'].astype(str).apply(lambda x: any(k in x for k in ['화','짜증','불편','답답','걱정','속상'])).mean()
chat_emotion = questions.apply(lambda x: any(k in x for k in ['화','짜증','불편','답답','걱정','속상'])).mean()
print(f'  감정 표현: 음성 {voice_emotion:.0%} / AI Chat {chat_emotion:.0%}')

# ══════════════════════════════════════════════════════════════════
# 10. 시간대별 분석
# ══════════════════════════════════════════════════════════════════
print(f'\n{sep}')
print('10. 시간대별 이용 패턴')
print(sep)

df['hour'] = df['발화일자'].dt.hour
hour_dist = df.groupby('hour').size()
peak_hour = hour_dist.idxmax()
print(f'피크 시간: {peak_hour}시 ({hour_dist[peak_hour]}건)')
print(f'시간대별:')
for h in range(9, 22):
    if h in hour_dist.index:
        bar = '█' * int(hour_dist[h] / hour_dist.max() * 20)
        print(f'  {h:2d}시: {hour_dist[h]:4d}건 {bar}')

# ══════════════════════════════════════════════════════════════════
# 엑셀 저장
# ══════════════════════════════════════════════════════════════════
output_path = os.path.join(BASE, 'analysis', 'AI_Chat_분석결과.xlsx')

with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    # 세션 단위 요약
    session_df.to_excel(writer, sheet_name='세션단위요약', index=False)

    # 비교 표
    comp_df.to_excel(writer, sheet_name='음성vsAIChat비교', index=False)

    # INTENT별 통계
    intent_stats = df.groupby('INTENT_CD').agg(
        건수=('INTENT_CD', 'size'),
        평균만족도=('만족도점수', 'mean'),
        직접입력비율=('INPUT_TYPE', lambda x: (x=='CONVERSATION').mean()),
        평균질문길이=('QUESTION', lambda x: x.astype(str).str.len().mean()),
    ).sort_values('건수', ascending=False).reset_index()
    intent_stats.to_excel(writer, sheet_name='INTENT별통계', index=False)

    # 피드백 분포
    fb_df = pd.DataFrame(feedback_counts).reset_index()
    fb_df.columns = ['피드백', '건수']
    fb_df.to_excel(writer, sheet_name='피드백분포', index=False)

    # 원본 데이터 (테스터 제거 후)
    df.to_excel(writer, sheet_name='원본데이터(테스터제거)', index=False)

print(f'\n엑셀 저장: {output_path}')

# JSON 저장
stats = {
    'n_sessions': int(n_sessions),
    'n_turns': int(n_turns),
    'avg_turns_per_session': round(n_turns/n_sessions, 1),
    'avg_satisfaction': round(float(chat_sat), 2),
    'avg_duration_sec': round(float(session_df['소요시간(초)'].mean()), 0),
    'model_name_ratio': round(float(session_df['모델명포함'].mean()), 3),
    'avg_q_length': round(float(chat_q_len), 1),
    'avg_a_length': round(float(chat_a_len), 1),
    'voice_empathy': round(float(voice_empathy), 3),
    'chat_empathy': round(float(chat_empathy), 3),
}
with open(os.path.join(BASE, 'analysis', 'outputs', 'data', 'aichat_stats.json'), 'w', encoding='utf-8') as f:
    json.dump(stats, f, ensure_ascii=False, indent=2)

print(f'JSON 저장: outputs/data/aichat_stats.json')
print('\n분석 완료!')
