"""
STT + Audio Fusion Sentiment Analyzer
- 60개 BERT 라벨 → 5개 상담 감정 그룹 확률 합산
- 음성 특징 → 감정 강도/방향 보정
- 최종: 그룹명 + Valence + 신뢰도
"""
import os
import sys
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import COUNSEL_EMOTION_GROUPS, LABEL_TO_GROUP, EMOTION_VALENCE_MAP


def compute_group_probs(text_probs: dict) -> dict:
    """
    BERT 60개 라벨 확률 → 5개 상담 감정 그룹 확률 합산.

    Args:
        text_probs: {label: prob} from Korean BERT

    Returns:
        {group_name: summed_probability}
    """
    group_probs = {g: 0.0 for g in COUNSEL_EMOTION_GROUPS}

    for label, prob in text_probs.items():
        group = LABEL_TO_GROUP.get(label)
        if group:
            group_probs[group] += prob

    return group_probs


def fuse_sentiment(text_probs: dict, audio_feats: dict = None,
                    text_weight: float = 0.6,
                    audio_weight: float = 0.4,
                    is_short_utterance: bool = False) -> dict:
    """
    텍스트 감정 확률 + 음성 특징 → 융합 감정 판정.

    짧은 발화(is_short_utterance=True)의 경우 텍스트 분석이 없으므로
    음성 특징 100%로 감정 판정.

    Args:
        text_probs: BERT 60-class 확률 분포 (짧은 발화면 빈 dict)
        audio_feats: 음성 특징 dict (f0_mean, energy_mean, zcr_mean 등)
        text_weight: 텍스트 가중치 (기본 0.6)
        audio_weight: 음성 가중치 (기본 0.4)
        is_short_utterance: True이면 음성 100%로 판정

    Returns:
        {
            "group": "감사/만족" | "안정/중립" | "불안/걱정" | "불만/짜증" | "혼란/당황",
            "group_probs": {group: prob},
            "valence": float [-1, +1],
            "confidence": float [0, 1],
            "audio_adjustment": float,
            "reasoning": str
        }
    """
    # ── 짧은 발화: 음성 100%로 판정 ─────────────────────────────────
    if is_short_utterance or not text_probs:
        return _audio_only_sentiment(audio_feats)

    # ── 1. 텍스트: 5그룹 확률 합산 + 라벨 편향 보정 ─────────────────────
    raw_group_probs = compute_group_probs(text_probs)

    # 라벨 개수 편향 보정:
    # 모델의 60개 라벨 중 부정이 48개라서 확률이 부정에 쏠림
    # → 그룹별 라벨 수로 나누어 정규화 (라벨당 평균 확률로 변환)
    group_label_counts = {
        g: max(len(info["labels"]), 1)
        for g, info in COUNSEL_EMOTION_GROUPS.items()
    }
    # 합산 확률 → 라벨당 평균 확률
    group_probs = {
        g: raw_group_probs[g] / group_label_counts[g]
        for g in raw_group_probs
    }
    # 재정규화 (합=1)
    total = sum(group_probs.values())
    if total > 0:
        group_probs = {k: v / total for k, v in group_probs.items()}

    # 텍스트 기반 Valence (그룹 Valence × 그룹 확률의 가중합)
    text_valence = 0.0
    for group, prob in group_probs.items():
        text_valence += COUNSEL_EMOTION_GROUPS[group]["valence"] * prob

    # ── 2. 음성: 감정 강도/방향 보정 ──────────────────────────────────
    audio_adjustment = 0.0
    audio_reasoning_parts = []

    if audio_feats and isinstance(audio_feats, dict):
        from reasoning import audio_to_valence_decomposed
        audio_valence, breakdown = audio_to_valence_decomposed(audio_feats)

        # 음성이 긍정 방향이면 "감사/만족", "안정/중립" 그룹 확률 보강
        # 음성이 부정 방향이면 "불안/걱정", "불만/짜증" 그룹 확률 보강
        audio_adjustment = audio_valence

        # F0 기울기: 상승 = 긍정 경향
        f0_slope = audio_feats.get("f0_slope", 0)
        if f0_slope > 10:
            audio_reasoning_parts.append(f"피치 상승(+{f0_slope:.0f}Hz)→긍정 보강")
        elif f0_slope < -10:
            audio_reasoning_parts.append(f"피치 하강({f0_slope:.0f}Hz)→부정 보강")

        # 에너지: 높은 에너지 = 감정 강도 증가
        energy = audio_feats.get("energy_mean", 0)
        if energy > 0.04:
            audio_reasoning_parts.append(f"높은 음량({energy:.3f})→감정 강도 높음")

        # ZCR: 높으면 긴장
        zcr = audio_feats.get("zcr_mean", 0)
        if zcr > 0.12:
            audio_reasoning_parts.append(f"높은 ZCR({zcr:.3f})→긴장/불안 보강")

    # ── 3. 융합 Valence ───────────────────────────────────────────────
    if audio_feats:
        fused_valence = text_weight * text_valence + audio_weight * audio_adjustment
    else:
        fused_valence = text_valence

    fused_valence = float(np.clip(fused_valence, -1.0, 1.0))

    # ── 4. 최종 그룹 판정 ─────────────────────────────────────────────
    # 음성 보정 반영한 그룹 확률 조정
    adjusted_probs = dict(group_probs)
    if audio_feats and abs(audio_adjustment) > 0.05:
        if audio_adjustment > 0:
            adjusted_probs["감사/만족"] += audio_adjustment * 0.3
            adjusted_probs["안정/중립"] += audio_adjustment * 0.2
        else:
            adjusted_probs["불안/걱정"] += abs(audio_adjustment) * 0.2
            adjusted_probs["불만/짜증"] += abs(audio_adjustment) * 0.2

    # 정규화
    total = sum(adjusted_probs.values())
    if total > 0:
        adjusted_probs = {k: v / total for k, v in adjusted_probs.items()}

    top_group = max(adjusted_probs, key=adjusted_probs.get)
    confidence = adjusted_probs[top_group]

    # ── 5. 판정 근거 ─────────────────────────────────────────────────
    text_top3 = sorted(group_probs.items(), key=lambda x: x[1], reverse=True)[:3]
    text_reason = ", ".join(f"{g} {p:.0%}" for g, p in text_top3)

    reasoning = f"텍스트 그룹 확률: [{text_reason}]"
    if audio_reasoning_parts:
        reasoning += f" | 음성 보정: {', '.join(audio_reasoning_parts)}"
    reasoning += f" → 최종 '{top_group}' ({confidence:.0%})"

    return {
        "group":            top_group,
        "group_probs":      adjusted_probs,
        "valence":          fused_valence,
        "confidence":       round(confidence, 3),
        "audio_adjustment": round(audio_adjustment, 4),
        "reasoning":        reasoning,
    }


def _audio_only_sentiment(audio_feats: dict) -> dict:
    """
    짧은 발화: 음성 특징 100%로 감정 판정.
    F0(피치), Energy(음량), ZCR(긴장도)로 감정 그룹 추정.
    """
    if not audio_feats or not isinstance(audio_feats, dict):
        return _empty_result()

    from reasoning import audio_to_valence_decomposed
    valence, breakdown = audio_to_valence_decomposed(audio_feats)

    f0_mean  = audio_feats.get("f0_mean", 0)
    energy   = audio_feats.get("energy_mean", 0)
    zcr      = audio_feats.get("zcr_mean", 0)
    f0_slope = audio_feats.get("f0_slope", 0)

    # 음성 특징 → 감정 그룹 판정 규칙
    reasons = []

    if valence > 0.1:
        group = "감사/만족"
        if f0_slope > 5:
            reasons.append(f"피치 상승(+{f0_slope:.0f}Hz)=밝은 톤")
        if energy > 0.03:
            reasons.append(f"높은 음량({energy:.3f})=활기")
    elif valence > -0.05:
        group = "안정/중립"
        reasons.append("음성 특징 중립 범위")
    elif valence > -0.15:
        group = "불안/걱정"
        if zcr > 0.1:
            reasons.append(f"높은 ZCR({zcr:.3f})=긴장")
        if f0_slope < -5:
            reasons.append(f"피치 하강({f0_slope:.0f}Hz)=위축")
    else:
        group = "불만/짜증"
        if energy > 0.04:
            reasons.append(f"높은 음량({energy:.3f})+부정 프로소디=불만")
        elif energy < 0.01:
            reasons.append(f"낮은 음량({energy:.3f})=무기력/실망")

    if not reasons:
        reasons.append(f"음성 Valence={valence:+.3f}")

    # 그룹 확률: 판정 그룹에 높은 확률 배정
    group_probs = {g: 0.05 for g in COUNSEL_EMOTION_GROUPS}
    group_probs[group] = 0.80
    total = sum(group_probs.values())
    group_probs = {k: v / total for k, v in group_probs.items()}

    reasoning = f"[음성 전용 판정] {', '.join(reasons)} → '{group}'"

    return {
        "group":            group,
        "group_probs":      group_probs,
        "valence":          float(np.clip(valence, -1.0, 1.0)),
        "confidence":       0.6,   # 음성 단독이므로 신뢰도 제한
        "audio_adjustment": round(valence, 4),
        "reasoning":        reasoning,
    }


def _empty_result() -> dict:
    return {
        "group": "안정/중립",
        "group_probs": {g: 0.2 for g in COUNSEL_EMOTION_GROUPS},
        "valence": 0.0,
        "confidence": 0.0,
        "audio_adjustment": 0.0,
        "reasoning": "분석 데이터 없음",
    }
