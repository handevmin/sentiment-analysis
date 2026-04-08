"""
Batch Analysis Runner
- 배치 폴더별 분석 (1_100 등)
- 전체 콜 분석
- 콜별 JSON + 리포트 + 종합 리포트 생성
"""
import os
import sys
import json
import glob
import numpy as np
import pandas as pd
from tqdm import tqdm

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import SPEECH_BATCHES, DATA_DIR, OUTPUT_DIR, CALL_STAGES
from data_loader import load_call_data, build_wav_index, merge_wav_paths, get_gt_stage_valence
from emotion_pipeline import analyze_call
from report_generator import generate_report
from transition_detector import generate_one_line_insight
from score_calibrator import calibrate_scores


def _calibrate_scores_from_meta(meta_row: pd.Series, res: dict) -> dict:
    """기존 점수를 읽고 보정하여 반환."""
    score_keys = [
        '상담사_해결의지_100', '상담사_솔루션구체성_100', '상담사_설명명확성_100',
        '상담사_공감표현_100', '상담사_주도성_100', '상담사_다음단계명확성_100',
        '고객_문제구체성_100', '고객_문제객관성_100', '고객_감정강도_100', '고객_협조도_100',
        '상호작용_해결진척도_100', '상호작용_마찰도_100', '상호작용_감정회복력_100',
    ]
    raw_scores = {k.replace('_100', ''): meta_row.get(k) for k in score_keys}

    # STT 텍스트 추출 (턴에서 재조립)
    stt = meta_row.get('상담번호: (STT) 대화내역', '')
    if not stt and res.get('turns'):
        parts = []
        for t in res['turns']:
            speaker = '[상담사]' if t.get('speaker') == 'agent' else '[고객]'
            parts.append(f"{speaker} {t.get('text', '')}")
        stt = '\n'.join(parts)

    return calibrate_scores(raw_scores, stt)


def _serialize_result(res: dict, meta_row: pd.Series) -> dict:
    """analyze_call 결과 → JSON 직렬화 가능 dict."""
    return {
        'cnid': res['cnid'],
        'duration_sec': res['duration_sec'],
        'text_stage_valence': res['text_stage_valence'],
        'audio_stage_valence': res['audio_stage_valence'],
        'transitions': res.get('transitions', []),
        'transition_summary': res.get('transition_summary', {}),
        'turns': [
            {k: v for k, v in t.items() if k != 'audio_features'}
            for t in res['turns']
        ],
        'meta': {
            'nps': int(meta_row.get('NPS') or 0),
            'consultant_score': str(meta_row.get('컨설턴트 만족도', '')),
            'gender': meta_row.get('성별', ''),
            'age_group': meta_row.get('연령대', ''),
            'product_l1': meta_row.get('상담번호: 제품명(대)', ''),
            'product_l2': meta_row.get('상담번호: 제품명(중)', ''),
            'symptom': meta_row.get('상담번호: 접수증상(중)', ''),
            'call_type': meta_row.get('상담번호: 상담유형(중)', ''),
            'paid': meta_row.get('상담번호: 유/무상', ''),
            'summary': meta_row.get('상담번호: 상담요약', ''),
            'nps_reason': meta_row.get('NPS 긍정적 선택사유') or meta_row.get('NPS 부정적 선택사유', ''),
            'consultant_reason': meta_row.get('컨설턴트 긍정적 평가사유', ''),
            'recv_dt': str(meta_row.get('상담번호: (시스템)접수일시', '')),
            'talk_time': str(meta_row.get('통화초', '')),
        },
        'gt': get_gt_stage_valence(meta_row),
        'scores': _calibrate_scores_from_meta(meta_row, res)
    }


def _save_json(data: dict, path: str):
    """JSON 저장 (numpy 타입 처리)."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2,
                  default=lambda o: float(o) if isinstance(o, (np.floating, np.integer)) else str(o))


def run_batch(df: pd.DataFrame, batch_name: str,
              run_text: bool = True, run_audio: bool = True,
              gen_reports: bool = True) -> list[dict]:
    """
    DataFrame의 각 행 분석 → JSON 저장 → (옵션) 콜별 리포트 → 종합 리포트.

    Returns: serialized result dicts list
    """
    batch_dir = os.path.join(DATA_DIR, f"batch_{batch_name}")
    os.makedirs(batch_dir, exist_ok=True)

    all_serialized = []
    errors = 0

    print(f"\n{'='*60}")
    print(f"  배치 분석: {batch_name} | 대상: {len(df)} 콜")
    print(f"  텍스트: {'ON' if run_text else 'OFF'} | 음성: {'ON' if run_audio else 'OFF'}")
    print(f"{'='*60}\n")

    for idx, (_, row) in enumerate(tqdm(df.iterrows(), total=len(df), desc=f"[{batch_name}] 분석")):
        cnid = row['CNID']

        try:
            res = analyze_call(row, run_text=run_text, run_audio=run_audio)
            serialized = _serialize_result(res, row)

            # 한줄 인사이트 추가
            serialized['one_line_insight'] = generate_one_line_insight(
                res.get('transitions', []), res.get('turns', [])
            )

            # JSON 저장 (점진적)
            _save_json(serialized, os.path.join(batch_dir, f"call_{cnid}.json"))

            # 콜별 HTML 리포트
            if gen_reports:
                report_path = os.path.join(batch_dir, f"report_{cnid}.html")
                generate_report(serialized, report_path)

            all_serialized.append(serialized)

        except Exception as e:
            errors += 1
            print(f"\n  [오류] CNID {cnid}: {e}")
            continue

    print(f"\n  완료: {len(all_serialized)} 건 성공 / {errors} 건 오류")

    # 종합 리포트 생성
    if all_serialized:
        from batch_report_generator import generate_batch_report
        agg_path = os.path.join(OUTPUT_DIR, f"batch_report_{batch_name}.html")
        generate_batch_report(all_serialized, batch_name, agg_path)
        print(f"  종합 리포트: {agg_path}")

    # 종합 CSV 저장
    _save_batch_csv(all_serialized, os.path.join(batch_dir, f"summary_{batch_name}.csv"))

    return all_serialized


def _save_batch_csv(results: list[dict], path: str):
    """배치 분석 결과 → 요약 CSV."""
    rows = []
    for r in results:
        row = {
            'CNID': r['cnid'],
            'NPS': r['meta'].get('nps'),
            '통화시간': r['duration_sec'],
            '성별': r['meta'].get('gender'),
            '연령대': r['meta'].get('age_group'),
            '제품': r['meta'].get('product_l2'),
            '전환_총수': r.get('transition_summary', {}).get('total', 0),
            '긍정회복': r.get('transition_summary', {}).get('neg_to_pos_count', 0),
            '부정전환': r.get('transition_summary', {}).get('pos_to_neg_count', 0),
            '인사이트': r.get('one_line_insight', ''),
        }
        for stage in CALL_STAGES:
            row[f'텍스트_{stage}'] = r['text_stage_valence'].get(stage)
            row[f'음성_{stage}'] = r['audio_stage_valence'].get(stage)
        rows.append(row)

    pd.DataFrame(rows).to_csv(path, index=False, encoding='utf-8-sig')
    print(f"  요약 CSV: {path}")


def run_batch_by_folder(folder_key: str, df: pd.DataFrame, wav_index: dict,
                         run_text: bool = True, run_audio: bool = True):
    """특정 배치 폴더(예: '1_100')의 WAV 파일만 분석."""
    target_dir = None
    for batch_dir in SPEECH_BATCHES:
        if folder_key in batch_dir:
            target_dir = batch_dir
            break

    if not target_dir or not os.path.isdir(target_dir):
        print(f"[오류] 배치 폴더 '{folder_key}'를 찾을 수 없습니다.")
        return []

    # 해당 폴더의 CNID만 필터
    folder_cnids = set()
    for wav_path in glob.glob(os.path.join(target_dir, "*.wav")):
        folder_cnids.add(os.path.splitext(os.path.basename(wav_path))[0])

    df_batch = df[df['CNID'].isin(folder_cnids)].reset_index(drop=True)
    print(f"[배치 필터] {folder_key}: {len(folder_cnids)} WAV → {len(df_batch)} 건 매칭")

    return run_batch(df_batch, folder_key, run_text=run_text, run_audio=run_audio)


def run_all_calls(df: pd.DataFrame,
                   run_text: bool = True, run_audio: bool = True):
    """WAV 파일이 있는 전체 콜 분석."""
    df_with_wav = df[df['wav_path'].notna()].reset_index(drop=True)
    return run_batch(df_with_wav, "all", run_text=run_text, run_audio=run_audio)
