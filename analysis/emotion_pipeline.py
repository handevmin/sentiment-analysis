"""
Emotion Analysis Pipeline
- 단일 콜: STT 파싱 → 타임스탬프 배분 → 텍스트 감성 → 음성 특징 → 단계별 집계
- 전체 배치: 병렬 처리
"""
import os
import sys
import traceback
import numpy as np
import pandas as pd
from tqdm import tqdm

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import CALL_STAGES, STAGE_RATIO
from utils.text_utils  import (parse_turns, assign_timestamps, assign_timestamps_whisper,
                                predict_text_valence, aggregate_stage_valence)
from utils.audio_utils import (load_wav, extract_segment_features,
                                extract_stage_features, audio_to_valence,
                                get_wav_duration)


def _baseline_corrected_valence(feats: dict, baseline: dict) -> float:
    """
    고객 통화 내 평균(baseline) 대비 상대적 변화로 Valence 산출.

    절대값 판정의 문제:
      - 원래 조용히 말하는 사람 → Energy 낮음 → "부정"으로 오판
      - 원래 피치가 높은 사람 → F0 높음 → 의미 없는 "긍정"

    baseline 보정:
      - 해당 통화 내 고객 전체 평균을 기준점(0)으로 설정
      - 평균보다 높으면 +, 낮으면 - (z-score 방식)
      - "이 사람 평소보다 에너지가 올랐다/내렸다"만 측정
    """
    v = 0.0

    # F0 기울기: baseline 대비 변화
    f0_slope = feats.get("f0_slope", 0)
    f0_base  = baseline.get("f0_slope", 0)
    f0_std   = baseline.get("f0_slope_std", 1) or 1
    f0_z = np.clip((f0_slope - f0_base) / max(f0_std, 5.0), -1, 1)
    v += 0.25 * f0_z

    # 에너지: baseline 대비 변화
    energy     = feats.get("energy_mean", 0)
    energy_base = baseline.get("energy_mean", 0)
    energy_std  = baseline.get("energy_mean_std", 1) or 1
    energy_z = np.clip((energy - energy_base) / max(energy_std, 0.005), -1, 1)
    # 에너지 변화 방향은 감정 "강도"이지 극성이 아님 → 절대값 약하게 반영
    v += 0.10 * energy_z

    # 에너지 기울기: baseline 대비
    e_slope     = feats.get("energy_slope", 0)
    e_slope_base = baseline.get("energy_slope", 0)
    e_slope_std  = baseline.get("energy_slope_std", 1) or 1
    e_slope_z = np.clip((e_slope - e_slope_base) / max(e_slope_std, 0.005), -1, 1)
    v += 0.10 * e_slope_z

    # ZCR: baseline 대비 변화 (높아지면 긴장)
    zcr      = feats.get("zcr_mean", 0)
    zcr_base = baseline.get("zcr_mean", 0)
    zcr_std  = baseline.get("zcr_mean_std", 1) or 1
    zcr_z = np.clip((zcr - zcr_base) / max(zcr_std, 0.02), -1, 1)
    v -= 0.10 * zcr_z  # ZCR 증가 = 긴장 → 부정

    # Voiced Ratio: baseline 대비
    vr      = feats.get("voiced_ratio", 0.5)
    vr_base = baseline.get("voiced_ratio", 0.5)
    vr_z    = (vr - vr_base)
    v += 0.05 * np.clip(vr_z * 5, -1, 1)

    # Jitter: baseline 대비 (증가 = 음성 떨림/동요 → 부정)
    jitter      = feats.get("jitter", 0)
    jitter_base = baseline.get("jitter", 0)
    jitter_std  = baseline.get("jitter_std", 1) or 1
    jitter_z = np.clip((jitter - jitter_base) / max(jitter_std, 0.5), -1, 1)
    v -= 0.08 * jitter_z

    # Shimmer: baseline 대비 (증가 = 에너지 떨림 → 부정)
    shimmer      = feats.get("shimmer", 0)
    shimmer_base = baseline.get("shimmer", 0)
    shimmer_std  = baseline.get("shimmer_std", 1) or 1
    shimmer_z = np.clip((shimmer - shimmer_base) / max(shimmer_std, 1.0), -1, 1)
    v -= 0.07 * shimmer_z

    # HNR: baseline 대비 (감소 = 잡음 증가/음성 품질 저하 → 부정)
    hnr      = feats.get("hnr", 0)
    hnr_base = baseline.get("hnr", 0)
    hnr_std  = baseline.get("hnr_std", 1) or 1
    hnr_z = np.clip((hnr - hnr_base) / max(hnr_std, 2.0), -1, 1)
    v += 0.05 * hnr_z  # HNR 증가 = 음성 맑음 → 긍정

    # 발화 내 방향 (전반→후반 변화)
    f0_dir = feats.get("f0_direction", 0)
    f0_dir_z = np.clip(f0_dir / 30.0, -1, 1)  # 30Hz 기준 정규화
    v += 0.05 * f0_dir_z  # 후반 피치 상승 = 긍정 마무리

    return float(np.clip(v, -1.0, 1.0))


# ── 단일 콜 분석 ──────────────────────────────────────────────────────────────

def analyze_call(row: pd.Series,
                 run_text:  bool = True,
                 run_audio: bool = True) -> dict:
    """
    하나의 상담 행(row)에 대해 전체 분석 수행.

    Returns:
        dict with keys:
            cnid, duration_sec
            turns: list[dict] (발화 단위, 텍스트/음성 감성 포함)
            text_stage_valence:  dict {단계: float}
            audio_stage_valence: dict {단계: float}
            audio_features:      list[dict] (단계별 음향 특징)
            error: str | None
    """
    result = {
        'cnid':                row['CNID'],
        'duration_sec':        float(row.get('duration_sec', 0) or 0),
        'turns':               [],
        'text_stage_valence':  {s: None for s in CALL_STAGES},
        'audio_stage_valence': {s: None for s in CALL_STAGES},
        'audio_features':      [],
        'error':               None,
    }

    # ── 1. STT 파싱 + 타임스탬프 배분 ────────────────────────────────
    stt_text = row.get('상담번호: (STT) 대화내역', '')
    turns    = parse_turns(stt_text)

    if not turns:
        result['error'] = "STT 텍스트 없음"
        return result

    # 실제 통화 시간 우선, 없으면 STT 문자 수 기반 추정
    total_dur = result['duration_sec']
    if total_dur <= 0:
        wav_path = row.get('wav_path', '')
        if wav_path and os.path.exists(wav_path):
            try:
                total_dur = get_wav_duration(wav_path)
            except Exception:
                total_dur = len(stt_text) * 0.05   # rough estimate

    # WAV가 있으면 Whisper 실제 타임스탬프, 없으면 비례 배분
    wav_path = row.get('wav_path', '')
    _cached_audio = None
    _cached_sr    = None
    if run_audio and wav_path and os.path.exists(wav_path):
        try:
            _cached_audio, _cached_sr = load_wav(wav_path)
            turns = assign_timestamps_whisper(turns, _cached_audio, _cached_sr, total_dur)
        except Exception:
            turns = assign_timestamps(turns, total_dur)
    else:
        turns = assign_timestamps(turns, total_dur)

    result['duration_sec'] = total_dur

    # ── 2. 텍스트 감성 분석 (고객 발화만) ────────────────────────────────
    # 짧은 응답은 텍스트 분석 스킵 → 4단계에서 음성만으로 판정
    SHORT_FILLER_WORDS = {'네', '예', '응', '어', '아', '음'}
    MIN_MEANINGFUL_LEN = 6

    if run_text:
        try:
            for turn in turns:
                if turn['speaker'] == '고객':
                    text_clean = turn['text'].replace(' ', '')
                    words = set(turn['text'].split())
                    is_short = (
                        len(text_clean) <= MIN_MEANINGFUL_LEN
                        or words.issubset(SHORT_FILLER_WORDS)
                    )

                    if is_short:
                        # 텍스트 판정 불가 → 음성 전용으로 표시
                        turn['text_valence']       = None
                        turn['text_emotion_label'] = '음성전용(짧은발화)'
                        turn['text_emotion_probs'] = {}
                        turn['is_short_utterance'] = True
                    else:
                        valence, label, probs = predict_text_valence(turn['text'])
                        turn['text_valence']       = valence
                        turn['text_emotion_label'] = label
                        turn['text_emotion_probs'] = probs
                        turn['is_short_utterance'] = False
                else:
                    turn['text_valence']       = None
                    turn['text_emotion_label'] = ''
                    turn['text_emotion_probs'] = {}
                    turn['is_short_utterance'] = False

            result['text_stage_valence'] = aggregate_stage_valence(turns)

            # ── [A] STT-only 결과 저장 (비교용) ──────────────────────
            from fusion import fuse_sentiment
            for turn in turns:
                if turn['speaker'] == '고객' and not turn.get('is_short_utterance'):
                    stt_only = fuse_sentiment(
                        turn.get('text_emotion_probs', {}),
                        audio_feats=None,
                        is_short_utterance=False
                    )
                    turn['stt_only_group']   = stt_only['group']
                    turn['stt_only_valence'] = stt_only['valence']
                else:
                    turn['stt_only_group']   = None
                    turn['stt_only_valence'] = None

        except Exception as e:
            result['error'] = f"텍스트 분석 오류: {e}"
            traceback.print_exc()

    # ── 3. 오디오 분석 (고객 발화만) ──────────────────────────────────
    if run_audio and wav_path and os.path.exists(wav_path):
        try:
            # Whisper 단계에서 이미 로드한 오디오 재사용
            if _cached_audio is not None:
                audio, sr = _cached_audio, _cached_sr
            else:
                audio, sr = load_wav(wav_path)

            # 단계별 음향 특징 추출
            stage_feats = extract_stage_features(audio, sr, total_dur, STAGE_RATIO)
            result['audio_features'] = stage_feats

            # 단계별 오디오 Valence 산출
            for i, stage in enumerate(CALL_STAGES):
                result['audio_stage_valence'][stage] = audio_to_valence(stage_feats[i])

            # 발화 단위 음향 특징 추출 (고객만)
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

            # 고객 baseline 계산 (해당 통화 내 평균)
            if cust_feats_list:
                import numpy as np
                baseline = {}
                for key in ['f0_mean', 'f0_slope', 'energy_mean', 'energy_slope',
                            'zcr_mean', 'voiced_ratio',
                            'jitter', 'shimmer', 'hnr']:
                    vals = [f.get(key, 0) for f in cust_feats_list if f.get(key, 0) != 0]
                    baseline[key] = float(np.mean(vals)) if vals else 0
                    baseline[f'{key}_std'] = float(np.std(vals)) if len(vals) > 1 else 1.0
            else:
                baseline = {}

            # baseline 보정된 Valence 산출
            for turn in turns:
                if turn['speaker'] == '고객' and turn.get('audio_features'):
                    feats = turn['audio_features']
                    turn['audio_valence'] = _baseline_corrected_valence(feats, baseline)
                    turn['audio_baseline'] = baseline  # 근거용 저장
                else:
                    turn['audio_valence'] = None

        except Exception as e:
            result['error'] = (result['error'] or '') + f" | 오디오 분석 오류: {e}"
            traceback.print_exc()

    # ── 4. STT+음성 융합 판정 (고객 발화만) ────────────────────────────
    try:
        from fusion import fuse_sentiment
        for turn in turns:
            if turn['speaker'] == '고객':
                is_short = turn.get('is_short_utterance', False)
                fusion = fuse_sentiment(
                    turn.get('text_emotion_probs', {}),
                    turn.get('audio_features'),
                    is_short_utterance=is_short
                )
                turn['fusion_group']      = fusion['group']
                turn['fusion_valence']    = fusion['valence']
                turn['fusion_confidence'] = fusion['confidence']
                turn['fusion_group_probs'] = fusion['group_probs']
                turn['fusion_reasoning']  = fusion['reasoning']
            else:
                turn['fusion_group']      = None
                turn['fusion_valence']    = None
                turn['fusion_confidence'] = None
                turn['fusion_group_probs'] = {}
                turn['fusion_reasoning']  = ''
    except Exception:
        pass

    # ── 4.5. LLM 기반 감정 판정 (Claude API) ────────────────────────
    try:
        from llm_analyzer import analyze_with_llm, apply_llm_results
        llm_results = analyze_with_llm(turns)
        if llm_results:
            apply_llm_results(turns, llm_results)
    except Exception:
        pass  # LLM 실패 시 기존 BERT+Audio fusion 유지

    # ── 5. 분석 근거 생성 ────────────────────────────────────────────
    try:
        from reasoning import generate_text_reasoning, generate_audio_reasoning
        for turn in turns:
            turn['text_reasoning']  = generate_text_reasoning(turn)
            turn['audio_reasoning'] = generate_audio_reasoning(turn)
    except Exception:
        pass

    # ── 6. 감정 전환 탐지 ─────────────────────────────────────────────
    try:
        from transition_detector import detect_transitions, summarize_transitions
        result['transitions']        = detect_transitions(turns)
        result['transition_summary'] = summarize_transitions(result['transitions'])
    except Exception:
        result['transitions']        = []
        result['transition_summary'] = {}

    result['turns'] = turns
    return result


# ── 전체 배치 분석 ────────────────────────────────────────────────────────────

def run_batch_analysis(df: pd.DataFrame,
                        run_text:  bool = True,
                        run_audio: bool = True) -> list[dict]:
    """
    DataFrame의 각 행에 대해 analyze_call 실행.
    """
    results = []
    for _, row in tqdm(df.iterrows(), total=len(df), desc="콜 분석"):
        res = analyze_call(row, run_text=run_text, run_audio=run_audio)
        results.append(res)
    return results


# ── 결과 → DataFrame 변환 ─────────────────────────────────────────────────────

def results_to_dataframe(results: list[dict],
                          df_meta: pd.DataFrame) -> pd.DataFrame:
    """
    analyze_call 결과 리스트 → 비교 분석용 DataFrame.
    df_meta: 원본 Call_Data DataFrame (NPS 등 포함)
    """
    rows = []
    for res in results:
        cnid = res['cnid']
        meta_row = df_meta[df_meta['CNID'] == cnid]

        row_dict = {'cnid': cnid}

        # 메타 컬럼 추가
        if not meta_row.empty:
            m = meta_row.iloc[0]
            row_dict['nps']              = m.get('NPS', np.nan)
            row_dict['consultant_score'] = m.get('컨설턴트 만족도', np.nan)
            row_dict['gender']           = m.get('성별', '')
            row_dict['age_group']        = m.get('연령대', '')
            row_dict['product']          = m.get('상담번호: 제품명(대)', '')
            row_dict['duration_sec']     = res['duration_sec']

        # 기존 GT Valence (텍스트 기반 어노테이션)
        for stage in CALL_STAGES:
            if not meta_row.empty:
                row_dict[f"gt_{stage}"] = float(meta_row.iloc[0].get(f"{stage}_Valence", 0) or 0)
            else:
                row_dict[f"gt_{stage}"] = np.nan

        # 텍스트 감성 단계별
        for stage in CALL_STAGES:
            row_dict[f"text_{stage}"] = res['text_stage_valence'].get(stage, np.nan)

        # 음성 감성 단계별
        for stage in CALL_STAGES:
            row_dict[f"audio_{stage}"] = res['audio_stage_valence'].get(stage, np.nan)

        # 음향 특징 단계별 (첫 단계만 예시)
        if res['audio_features']:
            first_feats = res['audio_features'][0]
            for k, v in first_feats.items():
                row_dict[f"feat_init_{k}"] = v

        rows.append(row_dict)

    return pd.DataFrame(rows)


def extract_audio_feature_matrix(results: list[dict]) -> tuple[np.ndarray, list[str]]:
    """
    모든 콜의 음향 특징 행렬 추출 (머신러닝 입력용).
    Returns: (X: ndarray [n_calls, n_features], feature_names: list)
    """
    all_rows = []
    feat_names = None

    for res in results:
        if not res['audio_features']:
            continue

        # 5단계 특징을 concat
        row_feats = []
        for i, stage in enumerate(CALL_STAGES):
            if i < len(res['audio_features']):
                feats = res['audio_features'][i]
            else:
                feats = {}
            for k, v in feats.items():
                if feat_names is None or f"{stage}_{k}" not in feat_names:
                    pass
                row_feats.append(float(v) if v is not None else 0.0)

        if feat_names is None and res['audio_features']:
            feat_names = []
            for stage in CALL_STAGES:
                for k in res['audio_features'][0].keys():
                    feat_names.append(f"{stage}_{k}")

        all_rows.append(row_feats)

    if not all_rows:
        return np.array([]), []

    return np.array(all_rows), feat_names or []
