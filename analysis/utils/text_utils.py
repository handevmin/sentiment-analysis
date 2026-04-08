"""
Text Utility Functions
- STT 텍스트 파싱 (발화 단위 분리)
- 타임스탬프 비례 배분
- Korean BERT 감성 분석
"""
import re
import os
import sys
import numpy as np
from typing import Optional

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import EMOTION_VALENCE_MAP, CALL_STAGES, STAGE_RATIO, TEXT_MODEL_ID

# ── STT 파싱 ──────────────────────────────────────────────────────────────────

SPEAKER_RE = re.compile(r'\[(상담사|고객)\]\s*')


def parse_turns(stt_text: str) -> list[dict]:
    """
    STT 텍스트 → 발화 단위 리스트 파싱.
    형식: "[상담사] 텍스트\n \n [고객] 텍스트\n ..."

    Returns:
        list of {
            'speaker': '상담사' | '고객',
            'text': str,
            'char_count': int,
            'turn_idx': int
        }
    """
    if not stt_text or not isinstance(stt_text, str):
        return []

    # \n \n 기준으로 분리 후 [상담사]/[고객] 태그 파싱
    segments = re.split(r'\n\s*\n', stt_text.strip())
    turns = []
    for i, seg in enumerate(segments):
        seg = seg.strip()
        if not seg:
            continue
        m = SPEAKER_RE.match(seg)
        if not m:
            continue
        speaker = m.group(1)
        text    = SPEAKER_RE.sub('', seg).strip()
        if not text:
            continue
        turns.append({
            'speaker':    speaker,
            'text':       text,
            'char_count': len(text),
            'turn_idx':   i,
        })
    return turns


def assign_timestamps(turns: list[dict], total_duration_sec: float) -> list[dict]:
    """
    발화 단위에 시작/종료 시각을 문자 수 비례로 배분 (fallback).

    Args:
        turns: parse_turns() 결과
        total_duration_sec: 전체 통화 시간(초)

    Returns:
        각 turn에 'start_sec', 'end_sec', 'mid_sec', 'stage' 추가
    """
    total_chars = sum(t['char_count'] for t in turns)
    if total_chars == 0:
        return turns

    cursor = 0.0
    stage_boundaries = _stage_boundaries(total_duration_sec)

    for turn in turns:
        ratio    = turn['char_count'] / total_chars
        duration = total_duration_sec * ratio
        turn['start_sec'] = cursor
        turn['end_sec']   = cursor + duration
        turn['mid_sec']   = cursor + duration / 2
        turn['stage']     = _get_stage(turn['mid_sec'], stage_boundaries)
        turn['timestamp_method'] = 'proportional'
        cursor += duration

    return turns


# ── Forced Alignment 기반 실제 타임스탬프 ─────────────────────────────────────

_whisper_model = None


def _load_stable_whisper(model_size: str = "base"):
    """stable-ts 모델 로드 (최초 1회)."""
    global _whisper_model
    if _whisper_model is None:
        import stable_whisper
        print(f"[Forced Alignment 모델 로딩] stable-ts {model_size}...")
        _whisper_model = stable_whisper.load_model(model_size)
        print("[Forced Alignment 모델 로딩 완료]")
    return _whisper_model


def assign_timestamps_whisper(turns: list[dict], audio: np.ndarray,
                               sr: int, total_duration_sec: float,
                               model_size: str = "base") -> list[dict]:
    """
    Forced Alignment으로 기존 STT 텍스트의 실제 발화 시각을 추출.

    stable-ts의 align()을 사용하여 기존 STT 텍스트를 오디오에
    단어 단위로 강제 정렬(±0.1초 정확도).

    Args:
        turns: parse_turns() 결과 (화자 태그 포함)
        audio: float32 PCM array (16kHz)
        sr: sample rate
        total_duration_sec: 전체 통화 시간
        model_size: Whisper 모델 크기 ("base" 권장)
    """
    import warnings
    warnings.filterwarnings('ignore')

    model = _load_stable_whisper(model_size)

    # 전체 STT 텍스트 합치기 (화자 태그 제거 상태)
    full_text = " ".join(t['text'] for t in turns)
    if not full_text.strip():
        return assign_timestamps(turns, total_duration_sec)

    try:
        result = model.align(audio, full_text, language='ko')
    except Exception:
        return assign_timestamps(turns, total_duration_sec)

    # 정렬된 모든 단어를 flat list로 추출
    all_words = []
    for seg in result.segments:
        for word in seg.words:
            if hasattr(word, 'start') and hasattr(word, 'end'):
                all_words.append({
                    'word':  word.word.strip(),
                    'start': word.start,
                    'end':   word.end,
                })

    if not all_words:
        return assign_timestamps(turns, total_duration_sec)

    # ── 각 turn의 단어를 정렬된 단어 리스트에서 순서대로 매칭 ─────────
    stage_boundaries = _stage_boundaries(total_duration_sec)
    word_cursor = 0
    n_total_words = len(all_words)

    for turn in turns:
        turn_words = turn['text'].split()
        n_turn_words = max(len(turn_words), 1)

        idx_start = min(word_cursor, n_total_words - 1)
        idx_end   = min(word_cursor + n_turn_words - 1, n_total_words - 1)

        turn_start = all_words[idx_start]['start']
        turn_end   = all_words[idx_end]['end']

        if turn_end <= turn_start:
            turn_end = turn_start + 0.3

        turn['start_sec'] = round(turn_start, 2)
        turn['end_sec']   = round(turn_end, 2)
        turn['mid_sec']   = round((turn_start + turn_end) / 2, 2)
        turn['stage']     = _get_stage(turn['mid_sec'], stage_boundaries)
        turn['timestamp_method'] = 'forced_alignment'

        word_cursor += n_turn_words

    return turns


def _stage_boundaries(total_dur: float) -> list[float]:
    """상담 단계 경계 시간(초) 리스트 반환."""
    boundaries = [0.0]
    acc = 0.0
    for r in STAGE_RATIO:
        acc += r * total_dur
        boundaries.append(acc)
    return boundaries


def _get_stage(t: float, boundaries: list[float]) -> str:
    """시각 t가 속하는 상담 단계 반환."""
    for i, (start, end) in enumerate(zip(boundaries[:-1], boundaries[1:])):
        if start <= t < end:
            return CALL_STAGES[i]
    return CALL_STAGES[-1]


def parse_duration_str(dur_str: str) -> float:
    """
    '03분 41초' 또는 '3분41초' 형식 → 초(float) 변환.
    """
    if not dur_str or not isinstance(dur_str, str):
        return 0.0
    m = re.search(r'(\d+)분\s*(\d+)초', dur_str)
    if m:
        return int(m.group(1)) * 60 + int(m.group(2))
    m2 = re.search(r'(\d+):(\d+)', dur_str)
    if m2:
        return int(m2.group(1)) * 60 + int(m2.group(2))
    return 0.0


# ── Korean BERT 감성 분석 ─────────────────────────────────────────────────────

_model  = None
_tokenizer = None


def _load_model():
    global _model, _tokenizer
    if _model is None:
        from transformers import pipeline
        print(f"[텍스트 모델 로딩] {TEXT_MODEL_ID}")
        _classifier = pipeline(
            "text-classification",
            model=TEXT_MODEL_ID,
            device=-1,           # CPU
            top_k=None,          # 전체 클래스 확률 반환
        )
        _model = _classifier
        print("[텍스트 모델 로딩 완료]")
    return _model


def predict_text_valence(text: str) -> tuple[float, str, dict]:
    """
    텍스트 → Valence 점수, 감정 레이블, 전체 확률 분포 반환.

    Returns:
        (valence: float, label: str, probs: dict)
    """
    if not text or len(text.strip()) < 2:
        return 0.0, "중립", {}

    clf = _load_model()

    # 모델 최대 토큰 제한 (512) 고려해 잘라내기
    text_clip = text[:500]
    results   = clf(text_clip)[0]

    # results: [{'label': 'xxx', 'score': 0.xx}, ...]
    probs = {r['label']: r['score'] for r in results}

    # 최고 확률 레이블
    top_label = max(probs, key=probs.get)

    # Valence: 가중 평균
    valence = 0.0
    for label, score in probs.items():
        v = EMOTION_VALENCE_MAP.get(label, 0.0)
        valence += v * score

    return float(np.clip(valence, -1.0, 1.0)), top_label, probs


def predict_batch_valence(texts: list[str],
                           batch_size: int = 32) -> list[tuple]:
    """배치 단위 텍스트 감성 분석."""
    clf = _load_model()
    results = []
    for i in range(0, len(texts), batch_size):
        batch = [t[:500] if t else " " for t in texts[i:i+batch_size]]
        try:
            batch_results = clf(batch)
        except Exception as e:
            batch_results = [[{"label": "중립", "score": 1.0}]] * len(batch)

        for res in batch_results:
            probs     = {r['label']: r['score'] for r in res}
            top_label = max(probs, key=probs.get)
            valence   = sum(EMOTION_VALENCE_MAP.get(l, 0.0) * s for l, s in probs.items())
            results.append((float(np.clip(valence, -1.0, 1.0)), top_label, probs))
    return results


# ── 발화 단계별 집계 ──────────────────────────────────────────────────────────

def aggregate_stage_valence(turns: list[dict]) -> dict:
    """
    타임스탬프가 배정된 발화 리스트에서 단계별 평균 Valence 집계.
    turns에 'text_valence', 'stage' 키가 있어야 함.
    """
    stage_vals = {s: [] for s in CALL_STAGES}
    for t in turns:
        stage = t.get('stage', '초기')
        val   = t.get('text_valence')
        if val is not None:
            stage_vals[stage].append(val)

    return {
        stage: float(np.mean(vals)) if vals else 0.0
        for stage, vals in stage_vals.items()
    }
