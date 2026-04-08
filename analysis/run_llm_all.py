"""
전체 LLM 교차검증 실행
- extract_all.py로 추출된 엑셀 데이터를 읽어서
- 각 콜의 고객 발화에 대해 LLM 교차검증 수행
- 결과를 엑셀에 업데이트
"""
import os
import sys
import json
import warnings
import traceback
import pandas as pd
import numpy as np
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor, as_completed

warnings.filterwarnings('ignore')
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from llm_analyzer import analyze_with_llm, apply_llm_results

EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '전체_음성분석_데이터.xlsx')
OUTPUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), '전체_음성분석_데이터_LLM.xlsx')
PROGRESS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'outputs', 'data', 'llm_cache')


PARALLEL_WORKERS = 10


def _process_one_call(cnid, call_turns_df):
    """단일 콜 LLM 교차검증 (스레드에서 실행)."""
    cache_path = os.path.join(PROGRESS_DIR, f'{cnid}.json')

    # 캐시 확인
    if os.path.exists(cache_path):
        with open(cache_path, encoding='utf-8') as f:
            return cnid, 'cached', json.load(f)

    turns_list = _df_to_turns(call_turns_df)
    if not turns_list:
        return cnid, 'skip', None

    try:
        llm_results = analyze_with_llm(turns_list)
        if llm_results:
            apply_llm_results(turns_list, llm_results)

            # 캐시 저장
            cache_data = [
                {'turn_idx': r.get('turn_idx'), 'group': r.get('group'),
                 'valence': r.get('valence'), 'confidence': r.get('confidence'),
                 'reasoning': r.get('reasoning', '')}
                for r in llm_results
            ]
            with open(cache_path, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)

            # turns에서 업데이트할 데이터 추출
            updates = []
            for t in turns_list:
                if t.get('speaker') == '고객' and t.get('fusion_method') == 'LLM심판':
                    updates.append({
                        'turn_idx': t.get('turn_idx'),
                        'group': t.get('fusion_group', ''),
                        'valence': t.get('fusion_valence'),
                        'confidence': t.get('fusion_confidence'),
                    })
            return cnid, 'done', updates
        else:
            return cnid, 'error', None
    except Exception as e:
        return cnid, 'error', str(e)


def run():
    os.makedirs(PROGRESS_DIR, exist_ok=True)

    print('엑셀 로딩...')
    df_turns = pd.read_excel(EXCEL_PATH, sheet_name='발화단위상세')
    df_calls = pd.read_excel(EXCEL_PATH, sheet_name='콜단위요약')
    print(f'콜 {len(df_calls)}건, 발화 {len(df_turns)}건')
    print(f'병렬 처리: {PARALLEL_WORKERS} workers')

    cnids = df_calls['CNID'].unique()

    # 각 콜의 발화 DataFrame 미리 준비
    call_data = {}
    for cnid in cnids:
        call_data[cnid] = df_turns[df_turns['CNID'] == cnid].copy()

    done = 0
    errors = 0
    skipped = 0
    error_msgs = []

    with ThreadPoolExecutor(max_workers=PARALLEL_WORKERS) as executor:
        futures = {
            executor.submit(_process_one_call, cnid, call_data[cnid]): cnid
            for cnid in cnids
        }

        for future in tqdm(as_completed(futures), total=len(futures), desc=f'LLM 교차검증 (x{PARALLEL_WORKERS})'):
            cnid = futures[future]
            try:
                cnid_result, status, data = future.result()

                if status == 'cached':
                    skipped += 1
                    _apply_cached(df_turns, cnid_result, data)
                elif status == 'done':
                    done += 1
                    if data:
                        for u in data:
                            mask = (df_turns['CNID'] == cnid_result) & (df_turns['turn_idx'] == u['turn_idx'])
                            df_turns.loc[mask, '융합감정그룹'] = u['group']
                            df_turns.loc[mask, '융합Valence'] = u['valence']
                            df_turns.loc[mask, '융합신뢰도'] = u['confidence']
                elif status == 'skip':
                    skipped += 1
                else:
                    errors += 1
                    if isinstance(data, str) and len(error_msgs) < 5:
                        error_msgs.append(f'{cnid_result}: {data[:80]}')

            except Exception as e:
                errors += 1
                if len(error_msgs) < 5:
                    error_msgs.append(f'{cnid}: {str(e)[:80]}')

            # 100건마다 중간 저장
            processed = done + skipped + errors
            if processed % 100 == 0 and processed > 0:
                _save_excel(df_calls, df_turns)
                print(f'\n  중간 저장 ({processed}/{len(cnids)})')

    if error_msgs:
        print('\n오류 샘플:')
        for msg in error_msgs:
            print(f'  {msg}')

    # 최종 저장 — 콜단위 평균 감성 업데이트
    for _, row in df_calls.iterrows():
        cnid = row['CNID']
        ct = df_turns[(df_turns['CNID'] == cnid) & (df_turns['화자'] == '고객')]
        fv = ct['융합Valence'].dropna()
        if len(fv) > 0:
            df_calls.loc[df_calls['CNID'] == cnid, '고객평균감성'] = round(float(fv.mean()), 4)

    _save_excel(df_calls, df_turns)

    print(f'\n완료: {done}건 LLM 처리, {skipped}건 캐시/스킵, {errors}건 오류')
    print(f'저장: {OUTPUT_PATH}')


def _df_to_turns(call_df):
    """DataFrame → turns list (LLM 입력용)."""
    turns = []
    for _, row in call_df.iterrows():
        t = {
            'turn_idx': int(row.get('turn_idx', 0)),
            'speaker': row.get('화자', ''),
            'text': str(row.get('발화내용', '')),
            'start_sec': float(row.get('시작(초)', 0) or 0),
            'end_sec': float(row.get('종료(초)', 0) or 0),
            'is_short_utterance': bool(row.get('짧은발화', False)),
            'fusion_group': row.get('융합감정그룹', ''),
            'fusion_confidence': float(row.get('융합신뢰도', 0) or 0),
            'fusion_group_probs': {},
            'audio_valence': float(row.get('음성Valence(보정)', 0) or 0) if pd.notna(row.get('음성Valence(보정)')) else None,
            'audio_features': {
                'f0_mean': float(row.get('F0_mean', 0) or 0),
                'f0_slope': float(row.get('F0_slope', 0) or 0),
                'energy_mean': float(row.get('Energy_mean', 0) or 0),
                'zcr_mean': float(row.get('ZCR', 0) or 0),
                'voiced_ratio': float(row.get('VoicedRatio', 0) or 0),
                'jitter': float(row.get('Jitter(%)', 0) or 0),
                'shimmer': float(row.get('Shimmer(%)', 0) or 0),
                'hnr': float(row.get('HNR(dB)', 0) or 0),
                'f0_direction': float(row.get('F0_direction', 0) or 0),
                'energy_direction': float(row.get('Energy_direction', 0) or 0),
            } if row.get('화자') == '고객' and pd.notna(row.get('F0_mean')) else None,
        }
        turns.append(t)
    return turns


def _apply_cached(df_turns, cnid, cached_results):
    """캐시된 LLM 결과를 DataFrame에 적용."""
    for r in cached_results:
        idx = r.get('turn_idx')
        mask = (df_turns['CNID'] == cnid) & (df_turns['turn_idx'] == idx)
        if r.get('group'):
            df_turns.loc[mask, '융합감정그룹'] = r['group']
        if r.get('valence') is not None:
            df_turns.loc[mask, '융합Valence'] = r['valence']
        if r.get('confidence') is not None:
            df_turns.loc[mask, '융합신뢰도'] = r['confidence']


def _save_excel(df_calls, df_turns):
    """엑셀 저장."""
    with pd.ExcelWriter(OUTPUT_PATH, engine='openpyxl') as writer:
        df_calls.to_excel(writer, sheet_name='콜단위요약', index=False)
        df_turns.to_excel(writer, sheet_name='발화단위상세', index=False)


if __name__ == '__main__':
    run()
