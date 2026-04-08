"""
LLM as Fusion Judge (Multi-model: Claude / GPT)
- LLM은 STT 원문을 직접 분석하지 않음
- BERT 텍스트 결과 + 음성 분석 결과가 이미 나온 상태에서
  둘이 충돌할 때 어떤 쪽을 신뢰할지 교차검증 역할만 수행
- Claude API 한도 초과 시 자동으로 GPT로 fallback
"""
import os
import sys
import json

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import COUNSEL_EMOTION_GROUPS

ANTHROPIC_API_KEY = ""
OPENAI_API_KEY = ""

# 사용할 모델 우선순위 (첫 번째 실패 시 다음으로)
MODEL_PRIORITY = [
    {"provider": "openai", "model": "gpt-5.4-mini", "key": OPENAI_API_KEY},
    {"provider": "anthropic", "model": "claude-sonnet-4-20250514", "key": ANTHROPIC_API_KEY},
]

SYSTEM_PROMPT = """당신은 콜센터 감성 분석 융합 심판관입니다.

두 가지 독립 분석 결과가 주어집니다:
1. BERT 텍스트 분석: 발화 텍스트에서 추출한 감정 그룹 확률
2. 음성 프로소디 분석: 목소리의 물리적 특징 (피치, 에너지, 긴장도)

당신의 역할은 이 두 분석을 비교·검증하여 최종 감정을 판정하는 것입니다.
STT 원문 텍스트를 직접 분석하지 마세요. 이미 BERT가 분석한 결과를 활용하세요.

## 판정 원칙
1. 텍스트와 음성이 일치하면 → 높은 신뢰도로 그대로 채택
2. 텍스트와 음성이 충돌하면 → 상황에 따라 판단:
   - 짧은 발화("네", "예" 등)는 음성을 우선 신뢰
   - 감정 표현이 명확한 발화("감사합니다", "화가 나요")는 텍스트를 우선 신뢰
   - 음성이 baseline 대비 긴장 증가(ZCR↑, 에너지 급변)를 보이면 부정 고려
   - 음성이 baseline 대비 안정적이면(Valence ≈ 0) 텍스트 분석 오류 가능성 고려
3. BERT가 확신이 낮은 경우(최대 그룹 확률 < 35%) → 음성 비중을 높임
4. 음성 Valence가 0에 가까우면(±0.05) 해당 발화는 고객 평소 톤과 유사 → 중립 신호
4. 대화 맥락에서 앞뒤 감정 흐름도 고려 (급격한 전환이 자연스러운지)

## 감정 그룹 (5개)
- 감사/만족 (Valence +0.3 ~ +1.0)
- 안정/중립 (Valence -0.1 ~ +0.2)
- 불안/걱정 (Valence -0.2 ~ -0.5)
- 불만/짜증 (Valence -0.5 ~ -1.0)
- 혼란/당황 (Valence -0.1 ~ -0.3)

## 출력 형식 (JSON 배열만, 다른 텍스트 없이)
[
  {
    "turn_idx": 1,
    "group": "안정/중립",
    "valence": 0.0,
    "confidence": 0.9,
    "reasoning": "BERT 불안/걱정 40%이나 음성이 차분(Energy 0.008, ZCR 낮음). 단순 확인 발화로 판단하여 음성 쪽 채택."
  }
]"""


def analyze_with_llm(turns: list[dict]) -> list[dict]:
    """
    Claude API를 융합 심판관으로 사용.
    BERT 결과 + 음성 결과를 전달하고 최종 판정을 받음.
    """
    import anthropic

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    # 각 고객 턴의 분석 결과 구성
    analysis_parts = []
    customer_indices = []

    for t in turns:
        speaker = t.get("speaker", "")
        idx = t.get("turn_idx", 0)

        if speaker != "고객":
            # 상담사 발화는 맥락용으로만 포함
            analysis_parts.append(
                f"Turn {idx} [상담사] (맥락): \"{t.get('text', '')[:60]}\""
            )
            continue

        customer_indices.append(idx)
        text = t.get("text", "")
        start = t.get("start_sec", 0)
        end = t.get("end_sec", 0)
        is_short = t.get("is_short_utterance", False)

        # BERT 분석 결과
        bert_group = t.get("fusion_group", "—")
        bert_confidence = t.get("fusion_confidence", 0)
        group_probs = t.get("fusion_group_probs", {})
        bert_top3 = sorted(group_probs.items(), key=lambda x: x[1], reverse=True)[:3]
        bert_str = ", ".join(f"{g}:{p:.0%}" for g, p in bert_top3) if bert_top3 else "N/A"

        # 음성 분석 결과
        af = t.get("audio_features")
        av = t.get("audio_valence")
        if af and isinstance(af, dict) and av is not None:
            audio_str = (f"F0={af.get('f0_mean',0):.0f}Hz(기울기{af.get('f0_slope',0):+.0f}), "
                        f"Energy={af.get('energy_mean',0):.3f}, "
                        f"ZCR={af.get('zcr_mean',0):.3f}, "
                        f"유성구간={af.get('voiced_ratio',0):.0%}, "
                        f"Jitter={af.get('jitter',0):.1f}%, "
                        f"Shimmer={af.get('shimmer',0):.1f}%, "
                        f"HNR={af.get('hnr',0):.1f}dB, "
                        f"F0방향={af.get('f0_direction',0):+.0f}Hz, "
                        f"에너지방향={af.get('energy_direction',0):+.4f}, "
                        f"AudioValence(baseline보정)={av:+.3f}")
        else:
            audio_str = "음성 데이터 없음"

        line = (f"Turn {idx} [{start:.1f}s~{end:.1f}s] {'(짧은발화)' if is_short else ''}\n"
                f"  발화: \"{text}\"\n"
                f"  BERT 판정: {bert_group} (신뢰 {bert_confidence:.0%}) [{bert_str}]\n"
                f"  음성 분석: {audio_str}")
        analysis_parts.append(line)

    if not customer_indices:
        return []

    analysis_text = "\n\n".join(analysis_parts)

    user_prompt = f"""아래 상담 콜의 각 고객 발화에 대해 BERT 텍스트 분석과 음성 프로소디 분석이 완료되었습니다.
두 분석 결과를 비교·검증하여 최종 감정을 판정해주세요.

음성 Valence는 해당 고객의 통화 내 평균 대비 상대적 변화입니다.
- AudioValence ≈ 0: 고객의 평소 톤과 유사 → 감정 변화 없음
- AudioValence > 0: 평소보다 밝아짐/활기
- AudioValence < 0: 평소보다 긴장/위축

특히 BERT와 음성이 충돌하는 경우, 어느 쪽을 신뢰할지 근거와 함께 판정하세요.

## 분석 결과
{analysis_text}

Turn {', '.join(str(i) for i in customer_indices)}에 대해 JSON 배열로 최종 판정을 출력하세요."""

    # 모델 우선순위대로 시도
    for model_config in MODEL_PRIORITY:
        try:
            response_text = _call_llm(model_config, SYSTEM_PROMPT, user_prompt)
            if not response_text:
                continue

            # JSON 파싱
            if response_text.startswith("["):
                return json.loads(response_text)
            start_pos = response_text.find("[")
            end_pos = response_text.rfind("]") + 1
            if start_pos >= 0 and end_pos > start_pos:
                return json.loads(response_text[start_pos:end_pos])

        except json.JSONDecodeError:
            continue
        except Exception as e:
            err_str = str(e)
            if "usage limits" in err_str or "rate" in err_str.lower() or "429" in err_str:
                continue  # 다음 모델로 fallback
            print(f"[AI 교차검증 오류] {model_config['provider']}/{model_config['model']}: {err_str[:80]}")
            continue

    return []


def _call_llm(config: dict, system_prompt: str, user_prompt: str) -> str:
    """provider별 API 호출."""
    provider = config["provider"]
    model = config["model"]
    key = config["key"]

    if provider == "anthropic":
        import anthropic
        client = anthropic.Anthropic(api_key=key)
        resp = client.messages.create(
            model=model, max_tokens=2000,
            system=system_prompt,
            messages=[{"role": "user", "content": user_prompt}]
        )
        return resp.content[0].text.strip()

    elif provider == "openai":
        from openai import OpenAI
        client = OpenAI(api_key=key)
        resp = client.chat.completions.create(
            model=model, max_completion_tokens=4096,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ]
        )
        return resp.choices[0].message.content.strip()

    return ""


def apply_llm_results(turns: list[dict], llm_results: list[dict]):
    """LLM 최종 판정을 turns에 적용."""
    result_map = {r["turn_idx"]: r for r in llm_results}

    for turn in turns:
        if turn.get("speaker") != "고객":
            continue

        idx = turn.get("turn_idx")
        llm = result_map.get(idx)

        if llm:
            turn["fusion_group"]      = llm.get("group", "안정/중립")
            turn["fusion_valence"]    = llm.get("valence", 0.0)
            turn["fusion_confidence"] = llm.get("confidence", 0.0)
            turn["fusion_reasoning"]  = llm.get("reasoning", "")
            turn["fusion_method"]     = "LLM심판"
