"""
Sentiment Transition Detector
- 연속 발화 간 감성 전환 탐지
- 부정→긍정 / 긍정→부정 분류 및 원인 분석
- 전환 패턴 요약 통계
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import CALL_STAGES

# 기본 임계값
TRANSITION_THRESHOLD    = 0.15   # 최소 Valence delta
SIGN_FLIP_MARGIN        = 0.05   # 부호 전환 판정 margin


def detect_transitions(turns: list[dict],
                        threshold: float = TRANSITION_THRESHOLD,
                        margin: float = SIGN_FLIP_MARGIN) -> list[dict]:
    """
    연속 발화 사이 감성 전환 이벤트를 탐지.

    Returns:
        list of transition dicts:
        {
            turn_idx_from, turn_idx_to,
            direction: "neg_to_pos" | "pos_to_neg" | "intensification",
            text_delta, audio_delta, magnitude,
            modality_agreement: bool,
            stage, speaker_changed,
            prev_turn, curr_turn (참조용),
            trigger_analysis: str (reasoning.py에서 생성)
        }
    """
    if len(turns) < 2:
        return []

    transitions = []

    for i in range(len(turns) - 1):
        prev = turns[i]
        curr = turns[i + 1]

        tv_prev = prev.get("text_valence", 0) or 0
        tv_curr = curr.get("text_valence", 0) or 0
        text_delta = tv_curr - tv_prev

        av_prev = prev.get("audio_valence")
        av_curr = curr.get("audio_valence")
        audio_delta = (av_curr - av_prev) if (av_prev is not None and av_curr is not None) else None

        # 전환 조건 판정
        is_large_shift = abs(text_delta) >= threshold
        is_sign_flip = (tv_prev < -margin and tv_curr > margin) or \
                       (tv_prev > margin and tv_curr < -margin)

        if not (is_large_shift or is_sign_flip):
            continue

        # 방향 분류
        if tv_prev < -margin and tv_curr > margin:
            direction = "neg_to_pos"
        elif tv_prev > margin and tv_curr < -margin:
            direction = "pos_to_neg"
        elif text_delta > 0:
            direction = "neg_to_pos"
        else:
            direction = "pos_to_neg"

        # 모달리티 일치 여부
        modality_agreement = None
        if audio_delta is not None:
            modality_agreement = (audio_delta > 0) == (text_delta > 0)

        transition = {
            "turn_idx_from":      i,
            "turn_idx_to":        i + 1,
            "direction":          direction,
            "text_delta":         round(text_delta, 4),
            "audio_delta":        round(audio_delta, 4) if audio_delta is not None else None,
            "magnitude":          round(abs(text_delta), 4),
            "modality_agreement": modality_agreement,
            "stage":              curr.get("stage", ""),
            "speaker_changed":    prev.get("speaker") != curr.get("speaker"),
            "prev_speaker":       prev.get("speaker", ""),
            "curr_speaker":       curr.get("speaker", ""),
            "prev_text":          prev.get("text", "")[:60],
            "curr_text":          curr.get("text", "")[:60],
            "prev_emotion":       prev.get("text_emotion_label", ""),
            "curr_emotion":       curr.get("text_emotion_label", ""),
            "prev_valence":       round(tv_prev, 4),
            "curr_valence":       round(tv_curr, 4),
            "time_sec":           curr.get("start_sec", 0),
            "trigger_analysis":   "",  # reasoning.py에서 채움
        }
        transitions.append(transition)

    # reasoning 생성
    from reasoning import generate_transition_reasoning
    for t in transitions:
        prev_t = turns[t["turn_idx_from"]]
        curr_t = turns[t["turn_idx_to"]]
        t["trigger_analysis"] = generate_transition_reasoning(prev_t, curr_t, t["direction"])

    return transitions


def summarize_transitions(transitions: list[dict]) -> dict:
    """
    전환 이벤트 목록의 요약 통계.

    Returns:
        {
            total, neg_to_pos_count, pos_to_neg_count,
            recovery_rate, avg_magnitude,
            by_stage: {stage: count},
            top_triggers: list[str],
            multimodal_agreement_rate: float
        }
    """
    if not transitions:
        return {
            "total": 0,
            "neg_to_pos_count": 0,
            "pos_to_neg_count": 0,
            "recovery_rate": 0.0,
            "avg_magnitude": 0.0,
            "by_stage": {s: 0 for s in CALL_STAGES},
            "multimodal_agreement_rate": 0.0,
        }

    n2p = [t for t in transitions if t["direction"] == "neg_to_pos"]
    p2n = [t for t in transitions if t["direction"] == "pos_to_neg"]

    total = len(transitions)
    n2p_count = len(n2p)
    p2n_count = len(p2n)
    recovery_rate = n2p_count / total if total > 0 else 0.0

    avg_magnitude = sum(t["magnitude"] for t in transitions) / total

    by_stage = {s: 0 for s in CALL_STAGES}
    for t in transitions:
        stage = t.get("stage", "")
        if stage in by_stage:
            by_stage[stage] += 1

    # 모달리티 일치율
    agreed = [t for t in transitions if t.get("modality_agreement") is True]
    total_with_audio = [t for t in transitions if t.get("modality_agreement") is not None]
    agreement_rate = len(agreed) / len(total_with_audio) if total_with_audio else 0.0

    return {
        "total": total,
        "neg_to_pos_count": n2p_count,
        "pos_to_neg_count": p2n_count,
        "recovery_rate": round(recovery_rate, 3),
        "avg_magnitude": round(avg_magnitude, 4),
        "by_stage": by_stage,
        "multimodal_agreement_rate": round(agreement_rate, 3),
    }


def find_most_dramatic_transition(transitions: list[dict]) -> dict | None:
    """가장 큰 감성 변화가 발생한 전환 이벤트 반환."""
    if not transitions:
        return None
    return max(transitions, key=lambda t: t["magnitude"])


def generate_one_line_insight(transitions: list[dict], turns: list[dict]) -> str:
    """
    전환 분석 기반 한줄 인사이트 문장 생성.
    배치 리포트의 콜별 요약에 사용.
    """
    if not transitions:
        return "감성 전환 없음 — 안정적 상담 흐름."

    summary = summarize_transitions(transitions)
    n2p = summary["neg_to_pos_count"]
    p2n = summary["pos_to_neg_count"]

    # 가장 극적인 전환 위치
    dramatic = find_most_dramatic_transition(transitions)
    stage = dramatic["stage"] if dramatic else ""

    if n2p > 0 and p2n == 0:
        return f"긍정 회복 {n2p}회 — {stage} 단계에서 감정 개선 확인."
    elif p2n > 0 and n2p == 0:
        return f"부정 전환 {p2n}회 — {stage} 단계에서 감정 악화 주의."
    elif n2p > p2n:
        return f"전환 {summary['total']}회 (회복 {n2p} > 악화 {p2n}) — 전반적 개선 추세."
    elif p2n > n2p:
        return f"전환 {summary['total']}회 (악화 {p2n} > 회복 {n2p}) — 불만 관리 필요."
    else:
        return f"전환 {summary['total']}회 (회복·악화 동수) — {stage} 단계 주목."
