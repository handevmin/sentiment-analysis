"""
Per-turn Analysis Reasoning Generator
- 텍스트 감성 값의 산출 근거 (감정 확률 분포 + 키워드)
- 음성 감성 값의 산출 근거 (4개 컴포넌트 기여도 분해)
"""
import os
import sys
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import EMOTION_VALENCE_MAP

# ── 감정 키워드 사전 ──────────────────────────────────────────────────────────
POSITIVE_KEYWORDS = [
    "감사", "고맙", "좋", "만족", "편리", "편하", "도움", "해결", "친절",
    "빠르", "정확", "완료", "성공", "기쁘", "다행", "안심", "안도",
]
NEGATIVE_KEYWORDS = [
    "불편", "불만", "짜증", "화", "걱정", "우려", "답답", "실망", "안되",
    "못", "지연", "오래", "어렵", "문제", "고장", "파손", "불량", "느리",
]
NEUTRAL_KEYWORDS = ["확인", "네", "예", "알겠", "여보세요"]


# ── 텍스트 감성 근거 ──────────────────────────────────────────────────────────

def generate_text_reasoning(turn: dict) -> str:
    """
    텍스트 감성 값의 산출 근거를 한국어 설명 문장으로 생성.

    Args:
        turn: dict with text_valence, text_emotion_label, text_emotion_probs, text

    Returns:
        한국어 설명 문자열
    """
    valence = turn.get("text_valence")
    label   = turn.get("text_emotion_label", "")
    probs   = turn.get("text_emotion_probs", {})
    text    = turn.get("text", "")

    if valence is None:
        return "텍스트 감성 분석 불가: 발화 텍스트가 없거나 모델 오류 발생."

    parts = []

    # 1) 최종 Valence 및 주 감정
    parts.append(f"텍스트 감성 {valence:+.3f}")
    if label:
        top_prob = probs.get(label, 0)
        parts.append(f"Korean BERT가 '{label}' 감정을 {top_prob:.0%} 확률로 감지.")

    # 2) 상위 3개 감정 확률 분포 + Valence 기여도
    if probs:
        sorted_probs = sorted(probs.items(), key=lambda x: x[1], reverse=True)[:3]
        contrib_parts = []
        for emo, prob in sorted_probs:
            v_map = EMOTION_VALENCE_MAP.get(emo, 0.0)
            contrib = v_map * prob
            contrib_parts.append(f"{emo} {prob:.0%}(기여 {contrib:+.3f})")
        parts.append("상위 감정 분포: " + ", ".join(contrib_parts) + ".")

    # 3) 키워드 분석
    found_pos = [kw for kw in POSITIVE_KEYWORDS if kw in text]
    found_neg = [kw for kw in NEGATIVE_KEYWORDS if kw in text]

    if found_pos or found_neg:
        kw_parts = []
        if found_pos:
            kw_parts.append(f"긍정 키워드 [{', '.join(found_pos[:3])}]")
        if found_neg:
            kw_parts.append(f"부정 키워드 [{', '.join(found_neg[:3])}]")
        parts.append("감지된 " + " / ".join(kw_parts) + ".")

    # 4) 짧은 발화 보정 안내
    if len(text) < 5:
        parts.append("(매우 짧은 발화로 감성 추정 신뢰도 낮음)")

    return " ".join(parts)


# ── 음성 감성 근거 ──────────────────────────────────────────────────────────

def audio_to_valence_decomposed(feats: dict) -> tuple[float, dict]:
    """
    audio_to_valence와 동일한 계산을 수행하되,
    각 컴포넌트별 기여도를 분해하여 함께 반환.

    Returns:
        (valence, {component_name: {"value": raw_value, "contribution": float, "description": str}})
    """
    breakdown = {}

    # F0 기울기
    f0_slope = feats.get("f0_slope", 0)
    f0_norm  = float(np.clip(f0_slope / 50.0, -1, 1))
    f0_contrib = 0.25 * f0_norm
    breakdown["F0 기울기"] = {
        "value": f0_slope,
        "contribution": f0_contrib,
        "description": f"F0 기울기 {f0_slope:+.1f}Hz → 정규화 {f0_norm:+.2f} × 가중치 0.25 = 기여 {f0_contrib:+.3f}"
                       + (" (피치 상승=긍정 추세)" if f0_slope > 5 else
                          " (피치 하강=부정 추세)" if f0_slope < -5 else
                          " (피치 안정)"),
    }

    # 에너지 기울기
    e_slope = feats.get("energy_slope", 0)
    e_norm  = float(np.clip(e_slope / 0.05, -1, 1))
    e_contrib = 0.15 * e_norm
    breakdown["에너지 기울기"] = {
        "value": e_slope,
        "contribution": e_contrib,
        "description": f"에너지 기울기 {e_slope:+.4f} → 기여 {e_contrib:+.3f}"
                       + (" (음량 증가=감정 고조)" if e_slope > 0.005 else
                          " (음량 감소=감정 진정)" if e_slope < -0.005 else
                          " (음량 안정)"),
    }

    # ZCR
    zcr = feats.get("zcr_mean", 0)
    zcr_norm = float(np.clip((zcr - 0.1) / 0.2, -1, 1))
    zcr_contrib = -0.10 * zcr_norm
    breakdown["ZCR"] = {
        "value": zcr,
        "contribution": zcr_contrib,
        "description": f"ZCR {zcr:.3f} → 기여 {zcr_contrib:+.3f}"
                       + (" (높은 ZCR=긴장/불안)" if zcr > 0.15 else " (안정적 음성)"),
    }

    # Voiced Ratio
    vr = feats.get("voiced_ratio", 0.5)
    vr_contrib = 0.10 * (vr - 0.5)
    breakdown["유성구간 비율"] = {
        "value": vr,
        "contribution": float(vr_contrib),
        "description": f"유성구간 {vr:.0%} → 기여 {vr_contrib:+.3f}"
                       + (" (활발한 발화)" if vr > 0.6 else " (침묵/무성구간 많음)" if vr < 0.3 else ""),
    }

    valence = float(np.clip(f0_contrib + e_contrib + zcr_contrib + vr_contrib, -1.0, 1.0))
    return valence, breakdown


def generate_audio_reasoning(turn: dict) -> str:
    """
    음성 감성 값의 산출 근거를 한국어 설명 문장으로 생성.

    Args:
        turn: dict with audio_valence, audio_features

    Returns:
        한국어 설명 문자열
    """
    av    = turn.get("audio_valence")
    feats = turn.get("audio_features")

    if av is None or feats is None:
        return "음성 분석 불가: WAV 파일 없음 또는 구간이 너무 짧음."

    _, breakdown = audio_to_valence_decomposed(feats)

    parts = [f"음성 감성 {av:+.3f}:"]

    # 기여도 절대값 내림차순
    sorted_items = sorted(breakdown.items(), key=lambda x: abs(x[1]["contribution"]), reverse=True)
    for name, info in sorted_items:
        parts.append(f"{name} {info['contribution']:+.3f}")

    detail = " | ".join(parts[1:])
    summary = parts[0] + " " + detail + "."

    # 종합 해석
    if av > 0.15:
        summary += " 음성 프로소디에서 긍정적 신호가 우세합니다."
    elif av < -0.15:
        summary += " 음성 프로소디에서 부정적 신호가 우세합니다."
    else:
        summary += " 음성 프로소디 중립 범위 내입니다."

    return summary


# ── 전환 근거 ─────────────────────────────────────────────────────────────────

def generate_transition_reasoning(prev_turn: dict, curr_turn: dict,
                                   direction: str) -> str:
    """
    감성 전환 지점의 원인 분석 문장 생성.

    Args:
        prev_turn, curr_turn: 전환 전후 발화 dict
        direction: "neg_to_pos" or "pos_to_neg"
    """
    parts = []

    # 화자 변화
    spk_prev = prev_turn.get("speaker", "")
    spk_curr = curr_turn.get("speaker", "")
    speaker_changed = spk_prev != spk_curr

    if direction == "neg_to_pos":
        parts.append("부정 → 긍정 전환 감지.")
        if speaker_changed and spk_curr == "상담사":
            parts.append("상담사의 응대가 고객 감정 개선에 기여한 것으로 판단됩니다.")
        elif speaker_changed and spk_curr == "고객":
            parts.append("고객이 이전 상담사 발화에 긍정적으로 반응하였습니다.")
    else:
        parts.append("긍정 → 부정 전환 감지.")
        if speaker_changed and spk_curr == "고객":
            parts.append("고객이 불만/우려를 표현한 것으로 보입니다.")
        elif speaker_changed and spk_curr == "상담사":
            parts.append("상담사 발화 후 감성이 하락하였습니다.")

    # 감정 레이블 변화
    emo_prev = prev_turn.get("text_emotion_label", "")
    emo_curr = curr_turn.get("text_emotion_label", "")
    if emo_prev and emo_curr and emo_prev != emo_curr:
        parts.append(f"감정 레이블이 '{emo_prev}' → '{emo_curr}'(으)로 변화하였습니다.")

    # 키워드 트리거
    text_curr = curr_turn.get("text", "")
    triggers = []
    if direction == "neg_to_pos":
        triggers = [kw for kw in POSITIVE_KEYWORDS if kw in text_curr]
        if triggers:
            parts.append(f"긍정 전환 트리거 키워드: [{', '.join(triggers[:3])}].")
    else:
        triggers = [kw for kw in NEGATIVE_KEYWORDS if kw in text_curr]
        if triggers:
            parts.append(f"부정 전환 트리거 키워드: [{', '.join(triggers[:3])}].")

    # 음성 변화
    av_prev = prev_turn.get("audio_valence")
    av_curr = curr_turn.get("audio_valence")
    if av_prev is not None and av_curr is not None:
        av_delta = av_curr - av_prev
        if abs(av_delta) > 0.1:
            agreement = "일치" if (av_delta > 0) == (direction == "neg_to_pos") else "불일치"
            parts.append(f"음성 감성도 {av_delta:+.3f} 변화 (텍스트와 {agreement}).")

    return " ".join(parts)
