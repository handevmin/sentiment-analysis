"""
전체 WAV 파일 데이터 추출 → 엑셀
- Sheet1: 콜 단위 요약
- Sheet2: 발화 단위 상세 (모든 turn × 컬럼)
"""
import os
import sys
import warnings
import traceback
import numpy as np
import pandas as pd
from tqdm import tqdm

warnings.filterwarnings('ignore')
os.environ['TRANSFORMERS_VERBOSITY'] = 'error'

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config import CALL_STAGES, OUTPUT_DIR
from data_loader import load_call_data, build_wav_index, merge_wav_paths
from utils.audio_utils import load_wav, extract_segment_features, get_wav_duration
from utils.text_utils import parse_turns, assign_timestamps_whisper, assign_timestamps, predict_text_valence
from fusion import fuse_sentiment


def extract_all():
    df = load_call_data()
    wav_idx = build_wav_index()
    df = merge_wav_paths(df, wav_idx)
    df_matched = df[df['wav_path'].notna()].reset_index(drop=True)

    print(f"\n전체 추출 시작: {len(df_matched)}건")
    print("(Forced Alignment + 음성 특징 + BERT 감정)")

    call_rows = []
    turn_rows = []
    errors = 0

    for idx, (_, row) in enumerate(tqdm(df_matched.iterrows(), total=len(df_matched), desc="추출")):
        cnid = row['CNID']
        wav_path = row['wav_path']

        try:
            # ── 오디오 로드 ────────────────────────────────────────
            audio, sr = load_wav(wav_path)
            duration = len(audio) / sr

            # ── STT 파싱 ──────────────────────────────────────────
            stt_text = row.get('상담번호: (STT) 대화내역', '')
            turns = parse_turns(stt_text)
            if not turns:
                errors += 1
                continue

            # ── Forced Alignment ──────────────────────────────────
            total_dur = float(row.get('duration_sec', 0) or 0)
            if total_dur <= 0:
                total_dur = duration

            try:
                turns = assign_timestamps_whisper(turns, audio, sr, total_dur)
            except Exception:
                turns = assign_timestamps(turns, total_dur)

            # ── 음성 특징 추출 (고객만) ───────────────────────────
            cust_feats_list = []
            for turn in turns:
                if turn['speaker'] == '고객':
                    feats = extract_segment_features(
                        audio, sr,
                        turn.get('start_sec', 0),
                        turn.get('end_sec', 0)
                    )
                    turn['audio_features'] = feats
                    cust_feats_list.append(feats)
                else:
                    turn['audio_features'] = None

            # ── Baseline 계산 ─────────────────────────────────────
            baseline = {}
            if cust_feats_list:
                for key in ['f0_mean', 'f0_slope', 'energy_mean', 'energy_slope',
                            'zcr_mean', 'voiced_ratio', 'jitter', 'shimmer', 'hnr']:
                    vals = [f.get(key, 0) for f in cust_feats_list if f.get(key, 0) != 0]
                    baseline[key] = float(np.mean(vals)) if vals else 0
                    baseline[f'{key}_std'] = float(np.std(vals)) if len(vals) > 1 else 1.0

            # ── Baseline 보정 Valence ─────────────────────────────
            from emotion_pipeline import _baseline_corrected_valence
            for turn in turns:
                if turn['speaker'] == '고객' and turn.get('audio_features'):
                    turn['audio_valence'] = _baseline_corrected_valence(turn['audio_features'], baseline)
                else:
                    turn['audio_valence'] = None

            # ── BERT 텍스트 감정 (고객만) ─────────────────────────
            SHORT_FILLER_WORDS = {'네', '예', '응', '어', '아', '음'}
            for turn in turns:
                if turn['speaker'] == '고객':
                    text_clean = turn['text'].replace(' ', '')
                    words = set(turn['text'].split())
                    is_short = len(text_clean) <= 6 or words.issubset(SHORT_FILLER_WORDS)
                    turn['is_short_utterance'] = is_short

                    if is_short:
                        turn['text_valence'] = None
                        turn['text_emotion_label'] = '음성전용'
                        turn['text_emotion_probs'] = {}
                    else:
                        v, label, probs = predict_text_valence(turn['text'])
                        turn['text_valence'] = v
                        turn['text_emotion_label'] = label
                        turn['text_emotion_probs'] = probs
                else:
                    turn['text_valence'] = None
                    turn['text_emotion_label'] = ''
                    turn['text_emotion_probs'] = {}
                    turn['is_short_utterance'] = False

            # ── 융합 판정 ─────────────────────────────────────────
            for turn in turns:
                if turn['speaker'] == '고객':
                    is_short = turn.get('is_short_utterance', False)
                    fusion = fuse_sentiment(
                        turn.get('text_emotion_probs', {}),
                        turn.get('audio_features'),
                        is_short_utterance=is_short
                    )
                    turn['fusion_group'] = fusion['group']
                    turn['fusion_valence'] = fusion['valence']
                    turn['fusion_confidence'] = fusion['confidence']

                    # STT-only 비교용
                    if not is_short:
                        stt_only = fuse_sentiment(turn.get('text_emotion_probs', {}), None, False)
                        turn['stt_only_group'] = stt_only['group']
                        turn['stt_only_valence'] = stt_only['valence']
                    else:
                        turn['stt_only_group'] = None
                        turn['stt_only_valence'] = None

            # ── 콜 단위 행 ────────────────────────────────────────
            cust_turns = [t for t in turns if t['speaker'] == '고객']
            cust_fv = [t.get('fusion_valence', 0) for t in cust_turns if t.get('fusion_valence') is not None]

            call_row = {
                'CNID': cnid,
                'NPS': row.get('NPS'),
                '컨설턴트만족도': row.get('컨설턴트 만족도'),
                '성별': row.get('성별'),
                '연령대': row.get('연령대'),
                '제품대': row.get('상담번호: 제품명(대)'),
                '제품중': row.get('상담번호: 제품명(중)'),
                '통화시간(초)': total_dur,
                '오디오길이(초)': round(duration, 1),
                '총발화수': len(turns),
                '고객발화수': len(cust_turns),
                '타임스탬프방식': turns[0].get('timestamp_method', ''),
                '고객평균감성': round(float(np.mean(cust_fv)), 3) if cust_fv else None,
                # baseline 특징
                'baseline_F0': round(baseline.get('f0_mean', 0), 1),
                'baseline_Energy': round(baseline.get('energy_mean', 0), 4),
                'baseline_ZCR': round(baseline.get('zcr_mean', 0), 4),
                'baseline_Jitter': round(baseline.get('jitter', 0), 2),
                'baseline_Shimmer': round(baseline.get('shimmer', 0), 2),
                'baseline_HNR': round(baseline.get('hnr', 0), 1),
            }

            # 단계별 감성
            stage_vals = {s: [] for s in CALL_STAGES}
            for t in cust_turns:
                stage = t.get('stage', '')
                fv = t.get('fusion_valence')
                if stage in stage_vals and fv is not None:
                    stage_vals[stage].append(fv)
            for s in CALL_STAGES:
                call_row[f'{s}_감성'] = round(float(np.mean(stage_vals[s])), 3) if stage_vals[s] else None

            call_rows.append(call_row)

            # ── 발화 단위 행 ──────────────────────────────────────
            for turn in turns:
                af = turn.get('audio_features') or {}
                turn_row = {
                    'CNID': cnid,
                    'turn_idx': turn.get('turn_idx'),
                    '화자': turn.get('speaker'),
                    '시작(초)': turn.get('start_sec'),
                    '종료(초)': turn.get('end_sec'),
                    '구간길이(초)': round(turn.get('end_sec', 0) - turn.get('start_sec', 0), 2),
                    '단계': turn.get('stage'),
                    '발화내용': turn.get('text'),
                    '짧은발화': turn.get('is_short_utterance', False),
                    # 텍스트 감정
                    'BERT라벨': turn.get('text_emotion_label', ''),
                    '텍스트Valence': turn.get('text_valence'),
                    # 음성 특징 (원시)
                    'F0_mean': round(af.get('f0_mean', 0), 1) if af else None,
                    'F0_std': round(af.get('f0_std', 0), 1) if af else None,
                    'F0_slope': round(af.get('f0_slope', 0), 1) if af else None,
                    'F0_direction': round(af.get('f0_direction', 0), 1) if af else None,
                    'Energy_mean': round(af.get('energy_mean', 0), 5) if af else None,
                    'Energy_std': round(af.get('energy_std', 0), 5) if af else None,
                    'Energy_slope': round(af.get('energy_slope', 0), 5) if af else None,
                    'Energy_direction': round(af.get('energy_direction', 0), 5) if af else None,
                    'ZCR': round(af.get('zcr_mean', 0), 4) if af else None,
                    'VoicedRatio': round(af.get('voiced_ratio', 0), 3) if af else None,
                    'Jitter(%)': round(af.get('jitter', 0), 2) if af else None,
                    'Shimmer(%)': round(af.get('shimmer', 0), 2) if af else None,
                    'HNR(dB)': round(af.get('hnr', 0), 1) if af else None,
                    'SpectralCentroid': round(af.get('spectral_centroid_mean', 0), 1) if af else None,
                    'SpeechRate': round(af.get('speech_rate', 0), 2) if af else None,
                    # 음성 Valence (baseline 보정)
                    '음성Valence(보정)': round(turn.get('audio_valence', 0), 3) if turn.get('audio_valence') is not None else None,
                    # 융합 결과
                    '융합감정그룹': turn.get('fusion_group'),
                    '융합Valence': round(turn.get('fusion_valence', 0), 3) if turn.get('fusion_valence') is not None else None,
                    '융합신뢰도': round(turn.get('fusion_confidence', 0), 3) if turn.get('fusion_confidence') is not None else None,
                    # STT-only 비교
                    'STT_Only그룹': turn.get('stt_only_group'),
                    'STT_Only_Valence': round(turn.get('stt_only_valence', 0), 3) if turn.get('stt_only_valence') is not None else None,
                }

                # MFCC 1~13
                for i in range(1, 14):
                    turn_row[f'MFCC_{i}'] = round(af.get(f'mfcc_{i}', 0), 3) if af else None

                turn_rows.append(turn_row)

        except Exception as e:
            errors += 1
            if errors <= 5:
                print(f"\n  [오류] {cnid}: {e}")
            continue

    print(f"\n완료: {len(call_rows)}건 성공, {errors}건 오류")

    # ── 엑셀 저장 ─────────────────────────────────────────────────
    output_path = os.path.join(os.path.dirname(OUTPUT_DIR), "전체_음성분석_데이터.xlsx")

    df_calls = pd.DataFrame(call_rows)
    df_turns = pd.DataFrame(turn_rows)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df_calls.to_excel(writer, sheet_name='콜단위요약', index=False)
        df_turns.to_excel(writer, sheet_name='발화단위상세', index=False)

    print(f"\n저장: {output_path}")
    print(f"  콜단위요약: {len(df_calls)}행 × {len(df_calls.columns)}열")
    print(f"  발화단위상세: {len(df_turns)}행 × {len(df_turns.columns)}열")


if __name__ == "__main__":
    extract_all()
