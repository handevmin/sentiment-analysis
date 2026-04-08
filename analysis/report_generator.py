"""
Call Sentiment Analysis Report Generator
- 단일 콜 분석 결과를 전문 HTML 리포트로 생성
- Plotly CDN 기반 인터랙티브 차트 임베딩
- 발화 테이블, 단계별 비교, 레이더 차트 포함
"""
import os
import sys
import json
import warnings
warnings.filterwarnings('ignore')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import CALL_STAGES, STAGE_RATIO, DATA_DIR, OUTPUT_DIR


# ── 색상 팔레트 ───────────────────────────────────────────────────────────────
C = {
    "blue":    "#1565C0",
    "blue_lt": "#42A5F5",
    "orange":  "#E65100",
    "green":   "#2E7D32",
    "green_lt":"#66BB6A",
    "purple":  "#6A1B9A",
    "purple_lt":"#AB47BC",
    "gray":    "#546E7A",
    "red":     "#C62828",
    "bg":      "#F8F9FA",
    "white":   "#FFFFFF",
    "border":  "#DEE2E6",
    "text":    "#212529",
    "muted":   "#6C757D",
}

STAGE_BG = ["#E3F2FD","#BBDEFB","#90CAF9","#64B5F6","#42A5F5"]


def valence_color(v: float) -> str:
    """Valence 값 → 배경색 (부정=빨강, 중립=흰색, 긍정=파랑)."""
    if v is None:
        return "#F5F5F5"
    v = float(v)
    if v >= 0.3:   return "#C8E6C9"
    if v >= 0.1:   return "#DCEDC8"
    if v >= -0.1:  return "#F5F5F5"
    if v >= -0.3:  return "#FFCCBC"
    return "#FFCDD2"

def valence_label(v: float) -> str:
    if v is None: return "N/A"
    v = float(v)
    if v >= 0.5:  return "매우 긍정"
    if v >= 0.2:  return "긍정"
    if v >= -0.2: return "중립"
    if v >= -0.5: return "부정"
    return "매우 부정"

def score_color(s: float) -> str:
    """0~100 점수 → 색상."""
    if s is None: return C["gray"]
    s = float(s)
    if s >= 70: return C["green"]
    if s >= 40: return C["blue"]
    return C["orange"]

def fmt_valence(v) -> str:
    if v is None: return "N/A"
    return f"{float(v):+.3f}"

def fmt_score(s) -> str:
    if s is None: return "—"
    return f"{float(s):.1f}"


# ── 차트 데이터 직렬화 ────────────────────────────────────────────────────────

def build_trajectory_chart_data(data: dict) -> str:
    """감정 궤적 Plotly JSON — 고객 발화의 융합 Valence 기준."""
    turns   = data["turns"]
    gt      = data["gt"]
    dur     = data["duration_sec"]

    boundaries = [0.0]
    acc = 0.0
    for r in STAGE_RATIO:
        acc += r * dur
        boundaries.append(acc)

    traces = []

    # ── 단계 배경 ─────────────────────────────────────────────────────
    shapes, annotations = [], []
    stage_alphas = ["rgba(227,242,253,0.5)","rgba(187,222,251,0.5)",
                    "rgba(144,202,249,0.5)","rgba(100,181,246,0.5)","rgba(66,165,245,0.4)"]
    for i, (stage, color) in enumerate(zip(CALL_STAGES, stage_alphas)):
        shapes.append({
            "type":"rect","xref":"x","yref":"paper",
            "x0":boundaries[i],"x1":boundaries[i+1],
            "y0":0,"y1":1,
            "fillcolor":color,"line":{"width":0},"layer":"below"
        })
        annotations.append({
            "x":(boundaries[i]+boundaries[i+1])/2,"y":1.02,
            "xref":"x","yref":"paper","text":f"<b>{stage}</b>",
            "showarrow":False,"font":{"size":11,"color":"#1565C0"},
            "align":"center"
        })

    # ── GT 수평선 ─────────────────────────────────────────────────────
    gt_x, gt_y = [], []
    for i, stage in enumerate(CALL_STAGES):
        gt_x += [boundaries[i], boundaries[i+1], None]
        gt_y += [gt.get(stage, 0), gt.get(stage, 0), None]
    traces.append({
        "type":"scatter","x":gt_x,"y":gt_y,
        "mode":"lines",
        "name":"기존 어노테이션 (GT)",
        "line":{"color":C["gray"],"width":2,"dash":"dot"},
        "hovertemplate":"GT: %{y:.3f}<extra></extra>"
    })

    # ── 고객 발화 융합 감성 궤적 ──────────────────────────────────────
    from config import COUNSEL_EMOTION_GROUPS
    cust_turns = [t for t in turns if t.get("speaker") == "고객"]
    cust_mid   = [t.get("mid_sec", 0) for t in cust_turns]
    cust_fv    = [t.get("fusion_valence", 0) or 0 for t in cust_turns]
    cust_group = [t.get("fusion_group", "") for t in cust_turns]
    cust_text  = [f"{t.get('fusion_group','')} | {t.get('text','')[:35]}" for t in cust_turns]
    cust_colors = [COUNSEL_EMOTION_GROUPS.get(g, {}).get("color", "#546E7A") for g in cust_group]

    traces.append({
        "type":"scatter","x":cust_mid,"y":cust_fv,
        "mode":"lines+markers","name":"고객 감성 (융합)",
        "line":{"color":"#1565C0","width":3},
        "marker":{"size":12,"color":cust_colors,
                  "line":{"color":"white","width":2}},
        "text":cust_text,
        "hovertemplate":"<b>%{x:.1f}초</b><br>Valence: %{y:.3f}<br>%{text}<extra></extra>"
    })

    # ── 중립선 ───────────────────────────────────────────────────────
    shapes.append({
        "type":"line","xref":"x","yref":"y",
        "x0":0,"x1":dur,"y0":0,"y1":0,
        "line":{"color":"rgba(0,0,0,0.2)","width":1,"dash":"solid"}
    })

    layout = {
        "height":400,"margin":{"t":40,"b":80,"l":60,"r":20},
        "template":"plotly_white",
        "legend":{"orientation":"h","y":-0.28,"x":0.5,"xanchor":"center","yanchor":"top"},
        "hovermode":"x unified",
        "xaxis":{"title":"시간 (초)","range":[0, dur]},
        "yaxis":{"title":"Valence","range":[-1.1,1.1],
                 "tickvals":[-1,-0.5,0,0.5,1],
                 "ticktext":["−1.0<br>(매우부정)","−0.5","0<br>(중립)","0.5","1.0<br>(매우긍정)"]},
        "shapes":shapes,"annotations":annotations,
        "font":{"family":"Malgun Gothic, Apple SD Gothic Neo, sans-serif","size":12}
    }

    return json.dumps({"data": traces, "layout": layout})


def build_stage_bar_chart_data(data: dict) -> str:
    """단계별 고객 감성 막대 차트 (융합 Valence 기준)."""
    # 단계별 융합 Valence 집계 (고객 발화만)
    from config import COUNSEL_EMOTION_GROUPS
    stage_vals = {s: [] for s in CALL_STAGES}
    stage_groups = {s: [] for s in CALL_STAGES}
    for t in data["turns"]:
        if t.get("speaker") != "고객":
            continue
        stage = t.get("stage", "")
        fv = t.get("fusion_valence")
        fg = t.get("fusion_group", "")
        if stage in stage_vals and fv is not None:
            stage_vals[stage].append(fv)
            stage_groups[stage].append(fg)

    import numpy as np
    fusion_v = [float(np.mean(stage_vals[s])) if stage_vals[s] else 0 for s in CALL_STAGES]
    # 각 단계의 지배적 감정 그룹
    from collections import Counter
    dominant = []
    for s in CALL_STAGES:
        if stage_groups[s]:
            dominant.append(Counter(stage_groups[s]).most_common(1)[0][0])
        else:
            dominant.append("—")
    bar_colors = [COUNSEL_EMOTION_GROUPS.get(d, {}).get("color", "#546E7A") for d in dominant]

    gt_v = [data["gt"].get(s, 0) or 0 for s in CALL_STAGES]

    traces = [
        {"type":"bar","name":"고객 감성 (융합)","x":CALL_STAGES,"y":fusion_v,
         "marker":{"color":bar_colors},
         "text":[f"{v:+.3f}<br>{d}" for v, d in zip(fusion_v, dominant)],
         "textposition":"outside"},
        {"type":"scatter","name":"GT 어노테이션","x":CALL_STAGES,"y":gt_v,
         "mode":"lines+markers","line":{"color":C["gray"],"dash":"dot","width":2},
         "marker":{"size":8,"color":C["gray"]}},
    ]
    layout = {
        "height":320,"margin":{"t":20,"b":60,"l":60,"r":20},
        "template":"plotly_white",
        "legend":{"orientation":"h","y":-0.28,"x":0.5,"xanchor":"center"},
        "yaxis":{"title":"Valence","range":[-1.1,1.1],
                 "zeroline":True,"zerolinecolor":"rgba(0,0,0,0.3)"},
        "font":{"family":"Malgun Gothic, Apple SD Gothic Neo, sans-serif","size":12}
    }
    return json.dumps({"data": traces, "layout": layout})


def build_radar_chart_data(data: dict) -> str:
    """상담사/고객/상호작용 지표 레이더 차트."""
    scores = data["scores"]
    counselor_keys = ["상담사_해결의지","상담사_솔루션구체성","상담사_설명명확성",
                      "상담사_공감표현","상담사_주도성","상담사_다음단계명확성"]
    customer_keys  = ["고객_문제구체성","고객_문제객관성","고객_감정강도","고객_협조도"]
    interact_keys  = ["상호작용_해결진척도","상호작용_마찰도","상호작용_감정회복력"]

    all_keys  = counselor_keys + customer_keys + interact_keys
    all_vals  = [scores.get(k, 0) or 0 for k in all_keys]
    all_labels = [k.split("_")[1] for k in all_keys]

    # 닫기 위해 첫 값 반복
    all_labels += [all_labels[0]]
    all_vals   += [all_vals[0]]

    traces = [{
        "type":"scatterpolar","r":all_vals,"theta":all_labels,
        "fill":"toself","name":"점수 (0~100)",
        "line":{"color":C["blue"]},"fillcolor":"rgba(21,101,192,0.15)",
        "marker":{"color":C["blue"]}
    }]
    layout = {
        "height":320,"margin":{"t":30,"b":30,"l":40,"r":40},
        "polar":{"radialaxis":{"visible":True,"range":[0,100],
                               "tickvals":[0,25,50,75,100]}},
        "template":"plotly_white","showlegend":False,
        "font":{"family":"Malgun Gothic, Apple SD Gothic Neo, sans-serif","size":11}
    }
    return json.dumps({"data": traces, "layout": layout})


def build_gantt_chart_data(data: dict) -> str:
    """화자별 발화 Gantt 차트."""
    turns = data["turns"]
    shapes, annotations = [], []

    for t in turns:
        color = C["blue"] if t["speaker"] == "상담사" else C["orange"]
        y = 1 if t["speaker"] == "상담사" else 0
        shapes.append({
            "type":"rect","xref":"x","yref":"y",
            "x0":t["start_sec"],"x1":t["end_sec"],
            "y0":y-0.4,"y1":y+0.4,
            "fillcolor":color,"opacity":0.7,
            "line":{"width":0}
        })

    traces = [
        {"type":"scatter","x":[None],"y":[None],"mode":"markers",
         "showlegend":False,"marker":{"color":C["blue"],"size":1}},
    ]
    layout = {
        "height":120,"margin":{"t":6,"b":36,"l":70,"r":20},
        "template":"plotly_white","shapes":shapes,
        "showlegend":False,
        "xaxis":{"title":"시간 (초)","range":[0, data["duration_sec"]],"tickfont":{"size":11}},
        "yaxis":{"tickvals":[0,1],"ticktext":["고객","상담사"],"range":[-0.6,1.6],
                 "tickfont":{"size":12}},
        "font":{"family":"Malgun Gothic, Apple SD Gothic Neo, sans-serif","size":12}
    }
    return json.dumps({"data": traces, "layout": layout})


# ── HTML 렌더링 ────────────────────────────────────────────────────────────────

def render_turn_table(turns: list) -> str:
    rows = ""
    for t in turns:
        spk   = t.get("speaker", "")
        text  = t.get("text", "")
        start = t.get("start_sec", 0)
        end   = t.get("end_sec", 0)
        stage = t.get("stage", "")

        # 상담사는 감정분석 없음
        if spk == "상담사":
            spk_badge = '<span class="badge-counselor">상담사</span>'
            rows += f"""
            <tr>
              <td class="center">{spk_badge}</td>
              <td class="time-cell">{start:.1f}s ~ {end:.1f}s</td>
              <td class="center"><span class="stage-pill">{stage}</span></td>
              <td class="text-cell">{text}</td>
              <td class="center" style="color:var(--gray-300)">—</td>
              <td class="center" style="color:var(--gray-300)">—</td>
              <td class="center" style="color:var(--gray-300)">—</td>
            </tr>"""
            continue

        # 고객: 융합 감정 표시
        spk_badge = '<span class="badge-customer">고객</span>'
        fusion_group = t.get("fusion_group") or "—"
        fv = t.get("fusion_valence")
        fc = t.get("fusion_confidence")
        is_short = t.get("is_short_utterance", False)

        # 감정 그룹 색상
        from config import COUNSEL_EMOTION_GROUPS
        group_color = COUNSEL_EMOTION_GROUPS.get(fusion_group, {}).get("color", "#546E7A")
        method_tag = "음성" if is_short else "융합"

        fv_display = fmt_valence(fv) if fv is not None else "—"
        fc_display = f"{fc:.0%}" if fc is not None else "—"
        fv_color = '#1B5E20' if (fv or 0) > 0.1 else '#B71C1C' if (fv or 0) < -0.1 else '#37474F'

        text_reason   = t.get("text_reasoning", "")
        audio_reason  = t.get("audio_reasoning", "")
        fusion_reason = t.get("fusion_reasoning", "")

        rows += f"""
        <tr>
          <td class="center">{spk_badge}</td>
          <td class="time-cell">{start:.1f}s ~ {end:.1f}s</td>
          <td class="center"><span class="stage-pill">{stage}</span></td>
          <td class="text-cell">{text}</td>
          <td class="center">
            <span style="display:inline-block;padding:2px 8px;border-radius:3px;font-size:11px;
                          font-weight:600;color:white;background:{group_color}">{fusion_group}</span>
            <br><small style="color:var(--gray-500)">{method_tag} · {fc_display}</small>
          </td>
          <td class="center" style="background:{valence_color(fv)}">
            <span class="valence-val" style="color:{fv_color}">{fv_display}</span>
          </td>
          <td class="reason-cell">
            <details><summary>근거 보기</summary>
              <div class="reason-detail">
                <div class="reason-item"><span class="reason-tag" style="background:#E8F5E9;color:#1B5E20">FUSION</span>{fusion_reason}</div>
                <div class="reason-item"><span class="reason-tag reason-tag-text">TEXT</span>{text_reason}</div>
                <div class="reason-item"><span class="reason-tag reason-tag-audio">AUDIO</span>{audio_reason}</div>
              </div>
            </details>
          </td>
        </tr>"""
    return rows


def render_stage_table(data: dict) -> str:
    """단계별 고객 감정 그룹 + 융합 Valence 테이블."""
    import numpy as np
    from collections import Counter
    from config import COUNSEL_EMOTION_GROUPS

    # 단계별 집계
    stage_vals = {s: [] for s in CALL_STAGES}
    stage_groups = {s: [] for s in CALL_STAGES}
    for t in data["turns"]:
        if t.get("speaker") != "고객":
            continue
        stage = t.get("stage", "")
        fv = t.get("fusion_valence")
        fg = t.get("fusion_group", "")
        if stage in stage_vals and fv is not None:
            stage_vals[stage].append(fv)
            stage_groups[stage].append(fg)

    rows = ""
    for i, stage in enumerate(CALL_STAGES):
        gt = data["gt"].get(stage, 0) or 0
        vals = stage_vals[stage]
        fv_mean = float(np.mean(vals)) if vals else 0
        n_turns = len(vals)

        # 지배적 감정 그룹
        if stage_groups[stage]:
            dominant = Counter(stage_groups[stage]).most_common(1)[0][0]
        else:
            dominant = "—"
        group_color = COUNSEL_EMOTION_GROUPS.get(dominant, {}).get("color", "#546E7A")

        rows += f"""
        <tr>
          <td style="background:{STAGE_BG[i]};font-weight:600;text-align:center">{stage}</td>
          <td style="text-align:center">
            <span style="display:inline-block;padding:2px 8px;border-radius:3px;font-size:11px;
                          font-weight:600;color:white;background:{group_color}">{dominant}</span></td>
          <td style="text-align:center;background:{valence_color(fv_mean)}">
            <b>{fmt_valence(fv_mean)}</b><br><small>{valence_label(fv_mean)}</small></td>
          <td style="text-align:center;background:{valence_color(gt)}">
            <b>{fmt_valence(gt)}</b></td>
          <td style="text-align:center;color:var(--gray-500)">{n_turns}</td>
        </tr>"""
    return rows


def render_score_bars(data: dict) -> str:
    scores = data["scores"]
    groups = [
        ("상담사 역량", ["상담사_해결의지","상담사_솔루션구체성","상담사_설명명확성",
                        "상담사_공감표현","상담사_주도성","상담사_다음단계명확성"], C["blue"]),
        ("고객 특성",  ["고객_문제구체성","고객_문제객관성","고객_감정강도","고객_협조도"], C["orange"]),
        ("상호작용",   ["상호작용_해결진척도","상호작용_마찰도","상호작용_감정회복력"], C["green"]),
    ]
    html = ""
    for group_name, keys, color in groups:
        html += f'<div class="score-group"><h4 style="color:{color};margin:0 0 8px 0">{group_name}</h4>'
        for key in keys:
            val = scores.get(key)
            display = float(val) if val is not None else 0
            label   = key.split("_",1)[1]
            bar_color = score_color(display)
            html += f"""
            <div class="score-row">
              <div class="score-label">{label}</div>
              <div class="score-bar-wrap">
                <div class="score-bar" style="width:{display}%;background:{bar_color}"></div>
              </div>
              <div class="score-num" style="color:{bar_color}">{display:.1f}</div>
            </div>"""
        html += "</div>"
    return html


def render_transition_section(data: dict) -> str:
    """감성 전환 분석 HTML 생성."""
    transitions = data.get("transitions", [])
    summary     = data.get("transition_summary", {})

    if not transitions:
        return """
        <div class="card">
          <div class="card-title">전환 감지 결과</div>
          <p style="color:var(--gray-500);font-size:13px;padding:8px 0">
            이 통화에서 유의미한 감성 전환이 감지되지 않았습니다. 전체적으로 안정적인 감성 흐름입니다.
          </p>
        </div>"""

    total  = summary.get("total", 0)
    n2p    = summary.get("neg_to_pos_count", 0)
    p2n    = summary.get("pos_to_neg_count", 0)
    rate   = summary.get("recovery_rate", 0)
    agreement = summary.get("multimodal_agreement_rate", 0)

    # 요약 카드
    html = f"""
    <div class="grid-3" style="margin-bottom:12px">
      <div class="card" style="text-align:center">
        <div class="card-title">총 전환 횟수</div>
        <div style="font-size:28px;font-weight:700;color:var(--navy-800)">{total}</div>
      </div>
      <div class="card" style="text-align:center">
        <div class="card-title">긍정 회복 / 부정 전환</div>
        <div style="font-size:22px;font-weight:700">
          <span style="color:var(--green)">{n2p}</span>
          <span style="color:var(--gray-300);margin:0 6px">/</span>
          <span style="color:var(--red)">{p2n}</span>
        </div>
        <div style="font-size:11px;color:var(--gray-500);margin-top:4px">회복률 {rate:.0%}</div>
      </div>
      <div class="card" style="text-align:center">
        <div class="card-title">텍스트-음성 방향 일치율</div>
        <div style="font-size:28px;font-weight:700;color:var(--navy-400)">{agreement:.0%}</div>
      </div>
    </div>"""

    # 개별 전환 타임라인
    html += '<div class="card"><div class="card-title">전환 이벤트 상세</div>'
    for i, t in enumerate(transitions):
        direction_label = "긍정 회복" if t["direction"] == "neg_to_pos" else "부정 전환"
        direction_color = "var(--green)" if t["direction"] == "neg_to_pos" else "var(--red)"
        arrow = "&#8593;" if t["direction"] == "neg_to_pos" else "&#8595;"

        html += f"""
        <div style="border-left:3px solid {direction_color};padding:10px 16px;margin-bottom:10px;
                     background:{'var(--green-bg)' if t['direction']=='neg_to_pos' else 'var(--red-bg)'};
                     border-radius:0 6px 6px 0">
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px">
            <span style="font-size:16px;color:{direction_color};font-weight:700">{arrow}</span>
            <span style="font-size:10px;font-weight:700;letter-spacing:0.5px;text-transform:uppercase;
                          color:{direction_color}">{direction_label}</span>
            <span style="font-size:11px;color:var(--gray-500)">{t.get('time_sec',0):.1f}초</span>
            <span class="stage-pill">{t.get('stage','')}</span>
            <span style="font-size:11px;color:var(--gray-500)">
              Valence {t.get('prev_valence',0):+.3f} → {t.get('curr_valence',0):+.3f}
              (delta {t.get('text_delta',0):+.3f})
            </span>
          </div>
          <div style="font-size:12px;color:var(--gray-700);line-height:1.7">
            <div style="margin-bottom:4px">
              <span class="badge-{'counselor' if t.get('prev_speaker')=='상담사' else 'customer'}">{t.get('prev_speaker','')}</span>
              &ldquo;{t.get('prev_text','')}&rdquo;
            </div>
            <div style="margin-bottom:6px">
              <span class="badge-{'counselor' if t.get('curr_speaker')=='상담사' else 'customer'}">{t.get('curr_speaker','')}</span>
              &ldquo;{t.get('curr_text','')}&rdquo;
            </div>
            <div style="font-size:11.5px;color:var(--navy-700);background:var(--navy-50);
                         padding:6px 10px;border-radius:4px">
              {t.get('trigger_analysis','')}
            </div>
          </div>
        </div>"""

    html += '</div>'
    return html


def render_ab_comparison(data: dict) -> str:
    """STT-only vs STT+Audio 비교 섹션 HTML."""
    turns = data.get("turns", [])
    cust_turns = [t for t in turns if t.get("speaker") == "고객"]

    # 비교 데이터 수집
    rows_html = ""
    changed = 0
    total   = 0
    for t in cust_turns:
        stt_g = t.get("stt_only_group")
        stt_v = t.get("stt_only_valence")
        fus_g = t.get("fusion_group")
        fus_v = t.get("fusion_valence")

        if stt_g is None and fus_g is None:
            continue

        total += 1
        stt_g_str = stt_g or "—"
        fus_g_str = fus_g or "—"
        stt_v_str = f"{stt_v:+.2f}" if stt_v is not None else "—"
        fus_v_str = f"{fus_v:+.2f}" if fus_v is not None else "—"

        is_changed = (stt_g != fus_g) and stt_g is not None and fus_g is not None
        if is_changed:
            changed += 1

        from config import COUNSEL_EMOTION_GROUPS
        stt_color = COUNSEL_EMOTION_GROUPS.get(stt_g_str, {}).get("color", "#546E7A")
        fus_color = COUNSEL_EMOTION_GROUPS.get(fus_g_str, {}).get("color", "#546E7A")

        highlight = "background:#FFF8E1;" if is_changed else ""
        arrow = "→" if is_changed else "="
        arrow_color = "var(--amber)" if is_changed else "var(--gray-300)"

        rows_html += f"""
        <tr style="{highlight}">
          <td class="time-cell">{t.get('start_sec',0):.1f}s</td>
          <td class="text-cell" style="max-width:200px">{t.get('text','')[:35]}</td>
          <td style="text-align:center">
            <span style="padding:2px 6px;border-radius:3px;font-size:10px;font-weight:600;
                          color:white;background:{stt_color}">{stt_g_str}</span>
            <br><small style="font-family:monospace">{stt_v_str}</small></td>
          <td style="text-align:center;font-size:16px;color:{arrow_color}">{arrow}</td>
          <td style="text-align:center">
            <span style="padding:2px 6px;border-radius:3px;font-size:10px;font-weight:600;
                          color:white;background:{fus_color}">{fus_g_str}</span>
            <br><small style="font-family:monospace">{fus_v_str}</small></td>
          <td style="font-size:11px;color:var(--gray-700);max-width:220px">{t.get('fusion_reasoning','')[:60]}</td>
        </tr>"""

    change_rate = f"{changed}/{total}" if total > 0 else "0/0"
    change_pct  = f"{changed/total:.0%}" if total > 0 else "0%"

    return f"""
    <div class="grid-2" style="margin-bottom:12px">
      <div class="card" style="text-align:center">
        <div class="card-title">음성 추가로 판정 변경된 발화</div>
        <div style="font-size:32px;font-weight:700;color:var(--navy-800)">{change_rate}</div>
        <div style="font-size:12px;color:var(--gray-500);margin-top:4px">변경률 {change_pct}</div>
      </div>
      <div class="card">
        <div class="card-title">해석</div>
        <div style="font-size:12.5px;color:var(--gray-700);line-height:1.8">
          변경률이 높을수록 음성 분석이 텍스트만으로는 포착하지 못하는 추가 감정 정보를 제공하고 있습니다.
          변경된 발화(노란 배경)에서 어떤 음성 근거로 판정이 바뀌었는지 확인하세요.
        </div>
      </div>
    </div>
    <div class="card" style="padding:0;overflow:hidden">
      <div class="table-scroll" style="max-height:350px">
        <table class="turn-table">
          <thead>
            <tr>
              <th style="width:60px">시간</th>
              <th>발화</th>
              <th style="width:100px">STT Only</th>
              <th style="width:30px"></th>
              <th style="width:100px">STT+Audio</th>
              <th>변경 근거</th>
            </tr>
          </thead>
          <tbody>{rows_html}</tbody>
        </table>
      </div>
    </div>"""


def generate_report(data: dict, output_path: str = None) -> str:
    """전체 HTML 리포트 생성."""
    meta = data["meta"]
    cnid = data["cnid"]

    traj_json  = build_trajectory_chart_data(data)
    bar_json   = build_stage_bar_chart_data(data)
    radar_json = build_radar_chart_data(data)
    gantt_json = build_gantt_chart_data(data)

    turn_rows  = render_turn_table(data["turns"])
    stage_rows = render_stage_table(data)
    score_bars = render_score_bars(data)

    n_turns = len(data["turns"])
    n_counselor = sum(1 for t in data["turns"] if t["speaker"] == "상담사")
    n_customer  = n_turns - n_counselor
    dur_min = int(data["duration_sec"] // 60)
    dur_sec = int(data["duration_sec"] % 60)

    # 고객 감정 요약 집계
    import numpy as np
    from collections import Counter
    cust_turns = [t for t in data["turns"] if t.get("speaker") == "고객"]
    cust_fv = [t.get("fusion_valence", 0) or 0 for t in cust_turns if t.get("fusion_valence") is not None]
    avg_valence = float(np.mean(cust_fv)) if cust_fv else 0
    cust_groups = [t.get("fusion_group", "") for t in cust_turns if t.get("fusion_group")]
    dominant_group = Counter(cust_groups).most_common(1)[0][0] if cust_groups else "—"
    n_transitions = data.get("transition_summary", {}).get("total", 0)

    transition_html      = render_transition_section(data)
    ab_comparison_html   = render_ab_comparison(data)

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>상담 감성 분석 리포트 — {cnid}</title>
<script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700&display=swap');

  :root {{
    --navy-900: #0A1628;
    --navy-800: #0F2044;
    --navy-700: #152D5F;
    --navy-600: #1B3A7A;
    --navy-500: #2350A0;
    --navy-400: #3A6BC4;
    --navy-300: #6892D5;
    --navy-100: #D6E4F7;
    --navy-50:  #EEF4FC;
    --accent:   #2563EB;
    --green:    #15803D;
    --green-bg: #F0FDF4;
    --amber:    #B45309;
    --amber-bg: #FFFBEB;
    --red:      #B91C1C;
    --red-bg:   #FEF2F2;
    --gray-700: #374151;
    --gray-500: #6B7280;
    --gray-300: #D1D5DB;
    --gray-100: #F3F4F6;
    --white:    #FFFFFF;
    --border:   #E5E7EB;
    --text:     #111827;
  }}

  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: 'Noto Sans KR', 'Malgun Gothic', sans-serif;
    background: #F1F5F9;
    color: var(--text);
    font-size: 13.5px;
    line-height: 1.6;
  }}

  /* ── 최상단 로고바 ───────────────────────── */
  .topbar {{
    background: var(--navy-900);
    padding: 0 48px;
    height: 48px;
    display: flex;
    align-items: center;
    justify-content: space-between;
  }}
  .topbar .brand {{
    font-size: 13px; font-weight: 600; letter-spacing: 1.5px;
    color: rgba(255,255,255,0.55); text-transform: uppercase;
  }}
  .topbar .report-id {{
    font-size: 12px; color: rgba(255,255,255,0.4);
    font-family: 'Consolas', monospace;
  }}

  /* ── 헤더 ────────────────────────────────── */
  .report-header {{
    background: var(--navy-800);
    color: white;
    padding: 28px 48px 0 48px;
    border-bottom: 3px solid var(--navy-500);
  }}
  .header-top {{
    display: flex;
    align-items: flex-start;
    justify-content: space-between;
    padding-bottom: 20px;
    border-bottom: 1px solid rgba(255,255,255,0.08);
  }}
  .header-top h1 {{
    font-size: 20px; font-weight: 700; letter-spacing: -0.3px;
    color: #FFFFFF; line-height: 1.3;
  }}
  .header-top .subtitle {{
    font-size: 12px; color: rgba(255,255,255,0.5);
    margin-top: 4px; letter-spacing: 0.2px;
  }}
  .header-badge {{
    display: flex; gap: 8px; flex-shrink: 0; margin-top: 2px;
  }}
  .hbadge {{
    padding: 4px 12px; border-radius: 4px; font-size: 11px;
    font-weight: 600; letter-spacing: 0.3px;
    border: 1px solid rgba(255,255,255,0.15);
    color: rgba(255,255,255,0.7);
    background: rgba(255,255,255,0.06);
  }}

  .meta-row {{
    display: flex; gap: 0; margin-top: 0;
  }}
  .meta-item {{
    padding: 14px 28px 14px 0;
    margin-right: 28px;
    border-right: 1px solid rgba(255,255,255,0.08);
  }}
  .meta-item:last-child {{ border-right: none; }}
  .meta-item .lbl {{
    font-size: 10px; color: rgba(255,255,255,0.4);
    text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 3px;
  }}
  .meta-item .val {{
    font-size: 13.5px; font-weight: 600; color: rgba(255,255,255,0.9);
  }}

  /* ── KPI 카드 ─────────────────────────────── */
  .kpi-row {{
    display: flex; gap: 12px;
    padding: 20px 48px; flex-wrap: wrap;
    background: #F1F5F9;
    border-bottom: 1px solid var(--border);
  }}
  .kpi-card {{
    background: var(--white);
    border: 1px solid var(--border);
    border-radius: 6px;
    flex: 1; min-width: 130px; text-align: left;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    overflow: hidden;
  }}
  .kpi-card .kpi-lbl {{
    font-size: 10px; color: var(--navy-400);
    text-transform: uppercase; letter-spacing: 0.8px;
    background: var(--navy-50);
    border-bottom: 1px solid var(--navy-100);
    padding: 7px 16px;
    font-weight: 600;
  }}
  .kpi-card .kpi-body {{
    padding: 14px 16px 14px 16px;
  }}
  .kpi-card .kpi-val {{
    font-size: 30px; font-weight: 700; line-height: 1;
    color: var(--navy-800);
  }}
  .kpi-card .kpi-val span {{ font-size: 14px; color: var(--gray-500); font-weight: 400; }}
  .kpi-card .kpi-sub {{
    font-size: 11.5px; margin-top: 6px;
    color: var(--gray-500);
  }}
  .kpi-card .kpi-tag {{
    display: inline-block; margin-top: 7px;
    padding: 3px 9px; border-radius: 3px; font-size: 10.5px; font-weight: 600;
  }}

  /* ── 콘텐츠 영역 ──────────────────────────── */
  .content-area {{ padding: 24px 48px; }}

  /* ── 섹션 ─────────────────────────────────── */
  .section {{ margin-bottom: 28px; }}
  .section-header {{
    display: flex; align-items: center; gap: 10px;
    margin-bottom: 14px;
  }}
  .section-num {{
    width: 24px; height: 24px; border-radius: 4px;
    background: var(--navy-700); color: white;
    font-size: 12px; font-weight: 700;
    display: flex; align-items: center; justify-content: center;
    flex-shrink: 0;
  }}
  .section-title {{
    font-size: 14px; font-weight: 700; color: var(--navy-700);
    letter-spacing: -0.1px;
  }}

  /* ── 카드 ─────────────────────────────────── */
  .card {{
    background: var(--white); border-radius: 8px;
    border: 1px solid var(--border); padding: 20px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    margin-bottom: 12px;
  }}
  .card-title {{
    font-size: 11px; font-weight: 600; color: var(--gray-500);
    text-transform: uppercase; letter-spacing: 0.6px;
    margin-bottom: 14px; padding-bottom: 10px;
    border-bottom: 1px solid var(--border);
  }}

  /* ── 차트 ─────────────────────────────────── */
  .chart-wrap {{ width: 100%; }}

  /* ── 발화 테이블 ──────────────────────────── */
  .turn-table {{ width: 100%; border-collapse: collapse; font-size: 12.5px; }}
  .turn-table th {{
    background: var(--navy-800); color: rgba(255,255,255,0.85);
    padding: 10px 12px; text-align: left;
    font-weight: 600; font-size: 11px; letter-spacing: 0.3px;
    position: sticky; top: 0;
  }}
  .turn-table td {{
    padding: 9px 12px; border-bottom: 1px solid var(--border);
    vertical-align: middle;
  }}
  .turn-table tr:nth-child(even) td {{ background: #FAFBFC; }}
  .turn-table tr:hover td {{ background: var(--navy-50) !important; }}
  .center {{ text-align: center !important; }}
  .reason-cell {{ font-size: 11px; max-width: 200px; }}
  .reason-cell details {{ cursor: pointer; }}
  .reason-cell summary {{
    font-size: 10.5px; color: var(--navy-400); font-weight: 600;
    padding: 2px 0; user-select: none;
  }}
  .reason-detail {{
    margin-top: 6px; font-size: 11px; line-height: 1.6;
    color: var(--gray-700);
  }}
  .reason-item {{
    padding: 4px 0; border-bottom: 1px solid var(--border);
  }}
  .reason-item:last-child {{ border-bottom: none; }}
  .reason-tag {{
    display: inline-block; padding: 1px 5px; border-radius: 2px;
    font-size: 9px; font-weight: 700; letter-spacing: 0.4px;
    margin-right: 4px; vertical-align: middle;
  }}
  .reason-tag-text {{ background: var(--navy-100); color: var(--navy-700); }}
  .reason-tag-audio {{ background: #F3E8FF; color: #6B21A8; }}
  .time-cell {{
    color: var(--gray-500); font-size: 11.5px;
    white-space: nowrap; font-family: 'Consolas', monospace;
  }}
  .text-cell {{ color: var(--text); max-width: 340px; }}
  .valence-val {{
    font-size: 13px; font-weight: 700;
    font-family: 'Consolas', monospace;
  }}
  .badge-counselor {{
    background: var(--navy-100); color: var(--navy-700);
    padding: 2px 9px; border-radius: 3px;
    font-size: 10.5px; font-weight: 600; white-space: nowrap;
  }}
  .badge-customer {{
    background: #FEF3C7; color: #92400E;
    padding: 2px 9px; border-radius: 3px;
    font-size: 10.5px; font-weight: 600; white-space: nowrap;
  }}
  .stage-pill {{
    background: #EDE9FE; color: #5B21B6;
    padding: 2px 7px; border-radius: 3px;
    font-size: 10.5px; white-space: nowrap;
  }}
  .table-scroll {{
    overflow-x: auto; max-height: 400px; overflow-y: auto;
    border-radius: 6px; border: 1px solid var(--border);
  }}

  /* ── 단계 비교 테이블 ────────────────────── */
  .stage-table {{ width: 100%; border-collapse: collapse; font-size: 12.5px; }}
  .stage-table th {{
    background: var(--navy-800); color: rgba(255,255,255,0.85);
    padding: 10px 14px; text-align: center;
    font-weight: 600; font-size: 11px; letter-spacing: 0.3px;
  }}
  .stage-table td {{ padding: 11px 14px; border-bottom: 1px solid var(--border); }}

  /* ── 그리드 ──────────────────────────────── */
  .grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }}
  .grid-3 {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px; }}
  @media (max-width: 900px) {{ .grid-2, .grid-3 {{ grid-template-columns: 1fr; }} }}

  /* ── 점수 바 ─────────────────────────────── */
  .score-group {{ margin-bottom: 18px; }}
  .score-group-title {{
    font-size: 10px; font-weight: 700; letter-spacing: 0.8px;
    text-transform: uppercase; margin-bottom: 10px;
    padding-bottom: 6px; border-bottom: 1px solid var(--border);
  }}
  .score-row {{ display: flex; align-items: center; margin-bottom: 7px; gap: 10px; }}
  .score-label {{
    width: 104px; font-size: 11.5px; color: var(--gray-700);
    flex-shrink: 0; text-align: right;
  }}
  .score-bar-wrap {{
    flex: 1; height: 7px; background: var(--gray-100);
    border-radius: 3px; overflow: hidden;
  }}
  .score-bar {{ height: 100%; border-radius: 3px; }}
  .score-num {{
    width: 38px; font-size: 11.5px; font-weight: 700;
    text-align: right; flex-shrink: 0;
  }}

  /* ── 인포 박스 ────────────────────────────── */
  .info-box {{
    background: var(--navy-50); border-left: 3px solid var(--navy-500);
    padding: 12px 16px; border-radius: 0 6px 6px 0;
    margin-bottom: 12px; font-size: 12.5px; line-height: 1.8;
    color: var(--gray-700);
  }}
  /* ── 요약 카드 ────────────────────────────── */
  .summary-card {{
    background: var(--white);
    border: 1px solid var(--border);
    border-radius: 6px;
    overflow: hidden;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
  }}
  .summary-card-header {{
    background: var(--navy-50);
    border-bottom: 1px solid var(--navy-100);
    padding: 8px 18px;
    font-size: 10px; font-weight: 600;
    color: var(--navy-400);
    text-transform: uppercase; letter-spacing: 0.8px;
  }}
  .summary-card-body {{
    padding: 0;
  }}
  .summary-point {{
    display: flex; align-items: flex-start; gap: 14px;
    padding: 13px 18px;
    border-bottom: 1px solid var(--border);
    font-size: 13px; line-height: 1.7; color: var(--gray-700);
  }}
  .summary-point:last-child {{ border-bottom: none; }}
  .summary-point-num {{
    flex-shrink: 0; width: 20px; height: 20px;
    border-radius: 50%;
    background: var(--navy-700); color: white;
    font-size: 10px; font-weight: 700;
    display: flex; align-items: center; justify-content: center;
    margin-top: 1px;
  }}
  .summary-card-footer {{
    display: flex; align-items: center; gap: 12px;
    padding: 11px 18px;
    background: var(--white);
    border-top: 1px solid var(--border);
  }}
  .summary-card-footer .footer-label {{
    flex-shrink: 0;
    background: var(--navy-800); color: white;
    font-size: 10px; font-weight: 700;
    letter-spacing: 0.5px; text-transform: uppercase;
    padding: 3px 10px; border-radius: 3px;
  }}
  .summary-card-footer .footer-text {{
    font-size: 12.5px; color: var(--gray-700);
  }}

  /* ── 방법론 항목 ──────────────────────────── */
  .method-label {{
    display: inline-block; width: 20px; height: 20px;
    background: var(--navy-700); color: white;
    border-radius: 3px; font-size: 11px; font-weight: 700;
    text-align: center; line-height: 20px; margin-right: 6px;
    flex-shrink: 0;
  }}

  /* ── 구분선 ──────────────────────────────── */
  .divider {{ height: 1px; background: var(--border); margin: 4px 0 16px 0; }}

  /* ── 푸터 ─────────────────────────────────── */
  .report-footer {{
    background: var(--navy-900);
    color: rgba(255,255,255,0.35);
    padding: 16px 48px; font-size: 11px;
    display: flex; justify-content: space-between; align-items: center;
    margin-top: 8px; border-top: 1px solid #1E2E4A;
  }}
</style>
</head>
<body>

<!-- ══ 로고바 ══════════════════════════════════════════════════════════════ -->
<div class="topbar">
  <span class="brand">Speech Sentiment Analysis</span>
  <span class="report-id">REPORT ID: {cnid} &nbsp;|&nbsp; {meta.get('recv_dt','')}</span>
</div>

<!-- ══ 헤더 ══════════════════════════════════════════════════════════════ -->
<div class="report-header">
  <div class="header-top">
    <div>
      <h1>상담 감성 분석 리포트</h1>
      <div class="subtitle">STT 텍스트 감성 분석 &nbsp;+&nbsp; 음성 프로소디 분석 비교 (STT vs STT+Audio)</div>
    </div>
    <div class="header-badge">
      <span class="hbadge">Korean BERT</span>
      <span class="hbadge">librosa</span>
      <span class="hbadge">Plotly</span>
    </div>
  </div>
  <div class="meta-row">
    <div class="meta-item">
      <div class="lbl">통화 ID</div>
      <div class="val">{cnid}</div>
    </div>
    <div class="meta-item">
      <div class="lbl">접수 일시</div>
      <div class="val">{meta.get('recv_dt','—')}</div>
    </div>
    <div class="meta-item">
      <div class="lbl">통화 시간</div>
      <div class="val">{meta.get('talk_time', f'{dur_min}분 {dur_sec}초')}</div>
    </div>
    <div class="meta-item">
      <div class="lbl">제품 분류</div>
      <div class="val">{meta.get('product_l1','—')} / {meta.get('product_l2','—')}</div>
    </div>
    <div class="meta-item">
      <div class="lbl">고객</div>
      <div class="val">{meta.get('gender','—')} · {meta.get('age_group','—')}</div>
    </div>
    <div class="meta-item">
      <div class="lbl">접수 증상</div>
      <div class="val">{meta.get('symptom','—')}</div>
    </div>
  </div>
</div>

<!-- ══ KPI 카드 ══════════════════════════════════════════════════════════ -->
<div class="kpi-row">
  <div class="kpi-card">
    <div class="kpi-lbl">NPS 점수</div>
    <div class="kpi-body">
      <div class="kpi-val" style="color:{'#15803D' if meta.get('nps',0)>=9 else '#B45309' if meta.get('nps',0)>=7 else '#B91C1C'}">{meta.get('nps','—')}<span>/10</span></div>
      <div class="kpi-tag" style="background:{'#F0FDF4' if meta.get('nps',0)>=9 else '#FFFBEB'};color:{'#15803D' if meta.get('nps',0)>=9 else '#B45309'}">
        {'추천 고객' if meta.get('nps',0)>=9 else '중립 고객' if meta.get('nps',0)>=7 else '비추천 고객'}
      </div>
    </div>
  </div>
  <div class="kpi-card">
    <div class="kpi-lbl">컨설턴트 만족도</div>
    <div class="kpi-body">
      <div class="kpi-val" style="color:{'#15803D' if float(meta.get('consultant_score') or 0)>=4 else '#B45309'}">{meta.get('consultant_score','—')}<span>/5</span></div>
      <div class="kpi-sub">{meta.get('consultant_reason','—')}</div>
    </div>
  </div>
  <div class="kpi-card">
    <div class="kpi-lbl">총 발화 수</div>
    <div class="kpi-body">
      <div class="kpi-val">{n_turns}</div>
      <div class="kpi-sub">상담사 {n_counselor}회 &nbsp;/&nbsp; 고객 {n_customer}회</div>
    </div>
  </div>
  <div class="kpi-card">
    <div class="kpi-lbl">고객 평균 감성</div>
    <div class="kpi-body">
      <div class="kpi-val" style="color:{'#15803D' if avg_valence>0.1 else '#B91C1C' if avg_valence<-0.1 else 'var(--navy-800)'}">{avg_valence:+.2f}</div>
      <div class="kpi-sub">{dominant_group} &nbsp;·&nbsp; 전환 {n_transitions}회</div>
    </div>
  </div>
  <div class="kpi-card">
    <div class="kpi-lbl">서비스 유형</div>
    <div class="kpi-body">
      <div class="kpi-val" style="font-size:20px;margin-top:4px">{meta.get('paid','—')}</div>
      <div class="kpi-sub">{meta.get('call_type','—')}</div>
    </div>
  </div>
</div>

<!-- ══ 콘텐츠 ═════════════════════════════════════════════════════════════ -->
<div class="content-area">

<!-- ── 1. 상담 요약 ────────────────────────────────────────────────────── -->
<div class="section">
  <div class="section-header">
    <div class="section-num">1</div>
    <div class="section-title">상담 내용 요약</div>
  </div>
  <div class="summary-card">
    <div class="summary-card-header">상담 내용 요약</div>
    <div class="summary-card-body">
      {"".join(
        f'<div class="summary-point"><div class="summary-point-num">{i+1}</div><div>{pt.strip()}</div></div>'
        for i, pt in enumerate(
          [p.strip() for p in (meta.get("summary") or "").split("\n") if p.strip()]
        )
      ) or '<div class="summary-point"><div>요약 정보 없음</div></div>'}
    </div>
    {"<div class='summary-card-footer'><span class='footer-label'>NPS 선택 사유</span><span class='footer-text'>" + str(meta.get('nps_reason','')) + "</span></div>" if meta.get('nps_reason') else ''}
  </div>
</div>

<!-- ── 2. 감정 변화 궤적 ────────────────────────────────────────────────── -->
<div class="section">
  <div class="section-header">
    <div class="section-num">2</div>
    <div class="section-title">고객 감정 변화 궤적</div>
  </div>
  <div class="card">
    <div class="card-title">고객 발화별 융합 감성 (STT + Audio + LLM) · 점 색상 = 감정 그룹 · 회색 점선 = GT 어노테이션</div>
    <div id="chart-traj" class="chart-wrap"></div>
  </div>
  <div class="card">
    <div class="card-title">화자별 발화 구간 (Gantt)</div>
    <div id="chart-gantt" class="chart-wrap"></div>
  </div>
  <div class="info-box" style="line-height:1.9">
    <b>분석가 해석 가이드</b><br>
    <b>차트의 점</b>: 고객 발화별 융합 감성(STT 텍스트 + 음성 프로소디 + LLM 맥락 분석)입니다.
    점 색상은 감정 그룹을 나타냅니다 — 녹색(감사/만족), 파랑(안정/중립), 주황(불안/걱정), 빨강(불만/짜증), 보라(혼란/당황).<br>
    <b>분석 방식</b>: 의미 있는 발화는 BERT 텍스트 60% + 음성 특징 40%를 융합하고, LLM(Claude)이 전체 대화 맥락을 보고
    최종 감정 그룹과 Valence를 판정합니다. 짧은 응답("네", "예")은 음성 톤만으로 판정합니다.<br>
    <b>감정 흐름 읽기</b>: 선이 0 아래로 내려가면 고객 불안/불만 구간이며,
    급격한 상승은 상담사의 해결 제시나 공감이 효과를 발휘한 시점입니다.
  </div>
</div>

<!-- ── 2.5. 감성 전환 분석 ───────────────────────────────────────────────── -->
<div class="section">
  <div class="section-header">
    <div class="section-num" style="background:var(--green)">T</div>
    <div class="section-title">감성 전환 분석 (부정 ↔ 긍정)</div>
  </div>
  {transition_html}
  <div class="info-box" style="line-height:1.9">
    <b>전환 분석 해석</b><br>
    <b>긍정 회복(녹색)</b>: 상담사의 공감 표현, 구체적 해결 방안 제시, 사과 후 발생합니다.
    회복이 많을수록 상담사의 감정 관리 역량이 높습니다.<br>
    <b>부정 전환(적색)</b>: 고객의 문제 재진술, 예상과 다른 답변, 대기 시간 불만 등에서 발생합니다.
    부정 전환이 해결시도·결과제시 단계에 집중되면 솔루션 적합성을 재검토해야 합니다.<br>
    <b>모달리티 일치율</b>: 텍스트와 음성 감성이 같은 방향으로 전환된 비율입니다.
    일치율이 높으면 전환이 실제 감정 변화를 반영하고, 낮으면 표면적 표현과 실제 감정 사이 괴리가 있습니다.
  </div>
</div>

<!-- ── 3. 단계별 감성 비교 ──────────────────────────────────────────────── -->
<div class="section">
  <div class="section-header">
    <div class="section-num">3</div>
    <div class="section-title">상담 단계별 고객 감정 분석</div>
  </div>
  <div class="grid-2">
    <div class="card">
      <div class="card-title">단계별 고객 감성 (융합) · 막대 색상 = 지배 감정 그룹</div>
      <div id="chart-bar" class="chart-wrap"></div>
    </div>
    <div class="card">
      <div class="card-title">단계별 수치 비교</div>
      <table class="stage-table">
        <thead>
          <tr>
            <th>단계</th>
            <th>감정 그룹</th>
            <th>융합 Valence</th>
            <th>GT</th>
            <th>발화 수</th>
          </tr>
        </thead>
        <tbody>{stage_rows}</tbody>
      </table>
    </div>
  </div>
</div>

<!-- ── 3.5. STT Only vs STT+Audio 비교 ─────────────────────────────────── -->
<div class="section">
  <div class="section-header">
    <div class="section-num" style="background:var(--amber)">AB</div>
    <div class="section-title">STT Only vs STT+Audio 비교 분석</div>
  </div>
  {ab_comparison_html}
  <div class="info-box" style="line-height:1.8">
    <b>연구 해석 가이드</b> — 이 섹션은 텍스트(STT)만으로 감정 분석한 결과[A]와
    음성 프로소디를 추가한 결과[B]를 직접 비교합니다.
    노란 배경으로 표시된 행은 음성 추가로 감정 판정이 변경된 발화이며,
    이 변경이 합리적인지 검토하는 것이 연구 핵심입니다.
  </div>
</div>

<!-- ── 4. 발화 단위 상세 ────────────────────────────────────────────────── -->
<div class="section">
  <div class="section-header">
    <div class="section-num">4</div>
    <div class="section-title">발화 단위 상세 분석</div>
  </div>
  <div class="card" style="padding:0;overflow:hidden">
    <div class="table-scroll">
      <table class="turn-table">
        <thead>
          <tr>
            <th style="width:72px">화자</th>
            <th style="width:115px">구간</th>
            <th style="width:64px">단계</th>
            <th>발화 내용</th>
            <th style="width:120px">감정 그룹</th>
            <th style="width:80px">Valence</th>
            <th style="width:180px">분석 근거</th>
          </tr>
        </thead>
        <tbody>{turn_rows}</tbody>
      </table>
    </div>
  </div>
</div>

<!-- ── 5. 상담 품질 지표 ────────────────────────────────────────────────── -->
<div class="section">
  <div class="section-header">
    <div class="section-num">5</div>
    <div class="section-title">상담 품질 지표 (0 ~ 100점)</div>
  </div>
  <div class="grid-2">
    <div class="card">
      <div class="card-title">레이더 차트 — 13개 지표 종합</div>
      <div id="chart-radar" class="chart-wrap"></div>
    </div>
    <div class="card">
      <div class="card-title">항목별 점수</div>
      {score_bars}
    </div>
  </div>
</div>

<!-- ── 6. 분석 방법론 ────────────────────────────────────────────────────── -->
<div class="section">
  <div class="section-header">
    <div class="section-num">6</div>
    <div class="section-title">분석 방법론</div>
  </div>
  <div class="grid-3">
    <div class="card">
      <div class="card-title">STT 텍스트 감성</div>
      <ul style="padding-left:16px;font-size:12.5px;line-height:2;color:var(--gray-700)">
        <li>모델: <code style="font-size:11px;background:#F3F4F6;padding:1px 5px;border-radius:3px">hun3359/klue-bert-base-sentiment</code></li>
        <li>출력: 7-class 감정 확률 분포</li>
        <li>Valence: 가중 평균 → [−1, +1]</li>
        <li>타임스탬프: 문자 수 비례 배분</li>
      </ul>
    </div>
    <div class="card">
      <div class="card-title">음성 프로소디 분석</div>
      <ul style="padding-left:16px;font-size:12.5px;line-height:2;color:var(--gray-700)">
        <li>WAV 포맷: GSM 8kHz → PCM 16kHz (ffmpeg)</li>
        <li>F0(피치): pyin 알고리즘</li>
        <li>에너지(RMS): 프레임별 평균·기울기</li>
        <li>MFCC 13차: 음색 특성 벡터</li>
        <li>ZCR, Speaking Rate, Voiced Ratio</li>
      </ul>
    </div>
    <div class="card">
      <div class="card-title">상담 단계 구분 기준</div>
      <ul style="padding-left:0;font-size:12.5px;line-height:1;color:var(--gray-700);list-style:none">
        <li style="padding:7px 0;border-bottom:1px solid var(--border);display:flex;justify-content:space-between"><span>초기</span><span style="color:var(--gray-500)">10% &nbsp;·&nbsp; 인사 및 고객 확인</span></li>
        <li style="padding:7px 0;border-bottom:1px solid var(--border);display:flex;justify-content:space-between"><span>탐색</span><span style="color:var(--gray-500)">25% &nbsp;·&nbsp; 문제 파악</span></li>
        <li style="padding:7px 0;border-bottom:1px solid var(--border);display:flex;justify-content:space-between"><span>해결시도</span><span style="color:var(--gray-500)">35% &nbsp;·&nbsp; 솔루션 제시</span></li>
        <li style="padding:7px 0;border-bottom:1px solid var(--border);display:flex;justify-content:space-between"><span>결과제시</span><span style="color:var(--gray-500)">20% &nbsp;·&nbsp; 처리 결과 전달</span></li>
        <li style="padding:7px 0;display:flex;justify-content:space-between"><span>종료</span><span style="color:var(--gray-500)">10% &nbsp;·&nbsp; 마무리 및 인사</span></li>
      </ul>
    </div>
  </div>
</div>

</div><!-- /content-area -->

<!-- ══ 푸터 ═══════════════════════════════════════════════════════════════ -->
<div class="report-footer">
  <div>Speech Sentiment Analysis Report &nbsp;·&nbsp; CNID: {cnid}</div>
  <div>Korean BERT Sentiment Model &nbsp;+&nbsp; Audio Prosody Analysis &nbsp;·&nbsp; librosa · transformers · plotly</div>
</div>

<script>
  var config = {{responsive: true, displayModeBar: false}};

  var trajData = {traj_json};
  Plotly.newPlot('chart-traj', trajData.data, trajData.layout, config);

  var ganttData = {gantt_json};
  Plotly.newPlot('chart-gantt', ganttData.data, ganttData.layout, config);

  var barData = {bar_json};
  Plotly.newPlot('chart-bar', barData.data, barData.layout, config);

  var radarData = {radar_json};
  Plotly.newPlot('chart-radar', radarData.data, radarData.layout, config);
</script>
</body>
</html>"""

    if output_path is None:
        output_path = os.path.join(OUTPUT_DIR, f"report_{cnid}.html")

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    return output_path


def generate_from_json(json_path: str, output_path: str = None) -> str:
    """JSON 파일에서 리포트 생성."""
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)
    return generate_report(data, output_path)


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--json",  type=str, help="분석 결과 JSON 경로")
    parser.add_argument("--call",  type=str, help="CNID (자동으로 JSON 검색)")
    parser.add_argument("--out",   type=str, default=None, help="출력 HTML 경로")
    args = parser.parse_args()

    if args.json:
        path = generate_from_json(args.json, args.out)
    elif args.call:
        json_path = os.path.join(DATA_DIR, f"call_{args.call}.json")
        if not os.path.exists(json_path):
            print(f"[오류] {json_path} 없음. 먼저 run.py --call {args.call} 실행하세요.")
            sys.exit(1)
        path = generate_from_json(json_path, args.out)
    else:
        # 기본값: 저장된 첫 번째 JSON
        jsons = [f for f in os.listdir(DATA_DIR) if f.startswith("call_") and f.endswith(".json")]
        if not jsons:
            print("[오류] 분석 결과 JSON 없음. run.py --call CNID 먼저 실행하세요.")
            sys.exit(1)
        path = generate_from_json(os.path.join(DATA_DIR, jsons[0]), args.out)

    print(f"[리포트 생성 완료] {path}")
    import webbrowser
    webbrowser.open(f"file:///{path.replace(os.sep, '/')}")
