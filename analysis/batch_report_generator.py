"""
Batch Aggregate Report Generator
- 배치 종합 HTML 리포트 (KPI, 분포, 히트맵, 전환 대시보드, 콜별 테이블)
"""
import os
import sys
import json
import numpy as np

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import CALL_STAGES, OUTPUT_DIR


def _valence_color(v):
    if v is None: return "#F5F5F5"
    v = float(v)
    if v >= 0.1:  return "#DCEDC8"
    if v >= -0.1: return "#F5F5F5"
    return "#FFCCBC"


def _fmt(v, d=3):
    if v is None: return "—"
    return f"{float(v):+.{d}f}"


def generate_batch_report(results: list[dict], batch_name: str,
                           output_path: str = None) -> str:
    """배치 종합 HTML 리포트 생성."""

    n = len(results)
    if n == 0:
        return ""

    # ── 집계 ──────────────────────────────────────────────────────────
    nps_vals = [r['meta'].get('nps', 0) for r in results if r['meta'].get('nps') is not None]
    nps_mean = np.mean(nps_vals) if nps_vals else 0
    nps_std  = np.std(nps_vals)  if nps_vals else 0

    # 단계별 평균
    text_means, audio_means = {}, {}
    for stage in CALL_STAGES:
        tv = [r['text_stage_valence'].get(stage, 0) or 0 for r in results]
        av = [r['audio_stage_valence'].get(stage, 0) or 0 for r in results]
        text_means[stage]  = float(np.mean(tv))
        audio_means[stage] = float(np.mean(av))

    # 전환 집계
    total_transitions = sum(r.get('transition_summary', {}).get('total', 0) for r in results)
    total_n2p = sum(r.get('transition_summary', {}).get('neg_to_pos_count', 0) for r in results)
    total_p2n = sum(r.get('transition_summary', {}).get('pos_to_neg_count', 0) for r in results)
    recovery_rate = total_n2p / total_transitions if total_transitions > 0 else 0

    # ── Plotly 차트 데이터 ────────────────────────────────────────────

    # 1) 단계별 평균 궤적
    traj_data = json.dumps({
        "data": [
            {"type": "scatter", "x": CALL_STAGES, "y": [text_means[s] for s in CALL_STAGES],
             "mode": "lines+markers", "name": "텍스트 감성 평균",
             "line": {"color": "#1565C0", "width": 3},
             "marker": {"size": 10}},
            {"type": "scatter", "x": CALL_STAGES, "y": [audio_means[s] for s in CALL_STAGES],
             "mode": "lines+markers", "name": "음성 감성 평균",
             "line": {"color": "#6A1B9A", "width": 3, "dash": "dash"},
             "marker": {"size": 10, "symbol": "diamond"}},
        ],
        "layout": {
            "height": 300, "margin": {"t": 20, "b": 60, "l": 60, "r": 20},
            "template": "plotly_white",
            "yaxis": {"title": "Valence (평균)", "range": [-0.5, 0.5],
                      "zeroline": True, "zerolinecolor": "rgba(0,0,0,0.2)"},
            "legend": {"orientation": "h", "y": -0.25, "x": 0.5, "xanchor": "center"},
            "font": {"family": "Malgun Gothic, sans-serif", "size": 12}
        }
    })

    # 2) NPS 분포
    nps_dist = json.dumps({
        "data": [{"type": "histogram", "x": nps_vals, "nbinsx": 11,
                  "marker": {"color": "#1565C0", "opacity": 0.8}}],
        "layout": {
            "height": 250, "margin": {"t": 10, "b": 40, "l": 50, "r": 20},
            "template": "plotly_white",
            "xaxis": {"title": "NPS 점수", "dtick": 1}, "yaxis": {"title": "건수"},
            "font": {"family": "Malgun Gothic, sans-serif", "size": 12}
        }
    })

    # 3) 히트맵 데이터
    hm_cnids = [r['cnid'] for r in results[:100]]  # 최대 100건
    hm_z = [[r['text_stage_valence'].get(s, 0) or 0 for s in CALL_STAGES]
            for r in results[:100]]
    heatmap_data = json.dumps({
        "data": [{"type": "heatmap", "z": hm_z, "x": CALL_STAGES, "y": hm_cnids,
                  "colorscale": [[0, "#C62828"], [0.5, "#F5F5F5"], [1, "#1565C0"]],
                  "zmid": 0, "zmin": -0.5, "zmax": 0.5,
                  "colorbar": {"title": "Valence"}}],
        "layout": {
            "height": max(300, len(hm_cnids) * 18),
            "margin": {"t": 10, "b": 50, "l": 100, "r": 20},
            "template": "plotly_white",
            "xaxis": {"title": "상담 단계"}, "yaxis": {"title": "통화 ID"},
            "font": {"family": "Malgun Gothic, sans-serif", "size": 11}
        }
    })

    # ── 콜별 요약 테이블 ──────────────────────────────────────────────
    call_rows = ""
    for r in sorted(results, key=lambda x: x['meta'].get('nps', 0)):
        cnid = r['cnid']
        nps  = r['meta'].get('nps', '—')
        dur  = f"{r['duration_sec']:.0f}s"
        ts   = r.get('transition_summary', {})
        n_trans = ts.get('total', 0)
        insight = r.get('one_line_insight', '')

        # 단계별 mini values
        stage_cells = ""
        for s in CALL_STAGES:
            tv = r['text_stage_valence'].get(s, 0) or 0
            stage_cells += f'<td style="text-align:center;background:{_valence_color(tv)};font-size:11px;font-family:monospace">{_fmt(tv)}</td>'

        call_rows += f"""
        <tr>
          <td style="font-family:monospace;font-size:11px">{cnid}</td>
          <td style="text-align:center;font-weight:700">{nps}</td>
          <td style="text-align:center;font-size:11px">{dur}</td>
          {stage_cells}
          <td style="text-align:center">{n_trans}</td>
          <td style="font-size:11px;max-width:250px">{insight}</td>
        </tr>"""

    # ── HTML ──────────────────────────────────────────────────────────
    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>배치 분석 종합 리포트 — {batch_name}</title>
<script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
<style>
  @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700&display=swap');
  :root {{
    --navy-900: #0A1628; --navy-800: #0F2044; --navy-700: #152D5F;
    --navy-500: #2350A0; --navy-400: #3A6BC4; --navy-100: #D6E4F7; --navy-50: #EEF4FC;
    --green: #15803D; --red: #B91C1C; --amber: #B45309;
    --gray-500: #6B7280; --gray-300: #D1D5DB; --gray-100: #F3F4F6;
    --white: #FFFFFF; --border: #E5E7EB; --text: #111827;
  }}
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Noto Sans KR', sans-serif; background: #F1F5F9; color: var(--text); font-size: 13px; line-height: 1.6; }}
  .topbar {{ background: var(--navy-900); padding: 0 48px; height: 44px; display: flex; align-items: center; justify-content: space-between; }}
  .topbar .brand {{ font-size: 12px; font-weight: 600; letter-spacing: 1.5px; color: rgba(255,255,255,0.55); text-transform: uppercase; }}
  .topbar .info {{ font-size: 11px; color: rgba(255,255,255,0.35); }}
  .header {{ background: var(--navy-800); color: white; padding: 24px 48px; border-bottom: 3px solid var(--navy-500); }}
  .header h1 {{ font-size: 20px; font-weight: 700; }}
  .header .sub {{ font-size: 12px; color: rgba(255,255,255,0.5); margin-top: 4px; }}
  .content {{ padding: 24px 48px; }}
  .kpi-row {{ display: flex; gap: 12px; margin-bottom: 20px; flex-wrap: wrap; }}
  .kpi {{ background: var(--white); border: 1px solid var(--border); border-radius: 6px; flex: 1; min-width: 120px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.06); }}
  .kpi-lbl {{ background: var(--navy-50); border-bottom: 1px solid var(--navy-100); padding: 6px 14px; font-size: 10px; font-weight: 600; color: var(--navy-400); text-transform: uppercase; letter-spacing: 0.6px; }}
  .kpi-body {{ padding: 12px 14px; }}
  .kpi-val {{ font-size: 28px; font-weight: 700; color: var(--navy-800); }}
  .kpi-val span {{ font-size: 13px; color: var(--gray-500); font-weight: 400; }}
  .kpi-sub {{ font-size: 11px; color: var(--gray-500); margin-top: 4px; }}
  .section {{ margin-bottom: 24px; }}
  .sec-header {{ display: flex; align-items: center; gap: 8px; margin-bottom: 12px; }}
  .sec-num {{ width: 22px; height: 22px; border-radius: 4px; background: var(--navy-700); color: white; font-size: 11px; font-weight: 700; display: flex; align-items: center; justify-content: center; }}
  .sec-title {{ font-size: 14px; font-weight: 700; color: var(--navy-700); }}
  .card {{ background: var(--white); border: 1px solid var(--border); border-radius: 6px; padding: 16px; box-shadow: 0 1px 3px rgba(0,0,0,0.06); margin-bottom: 12px; }}
  .card-title {{ font-size: 10px; font-weight: 600; color: var(--gray-500); text-transform: uppercase; letter-spacing: 0.6px; margin-bottom: 10px; padding-bottom: 8px; border-bottom: 1px solid var(--border); }}
  .grid-2 {{ display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }}
  .grid-3 {{ display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 12px; }}
  table.summary {{ width: 100%; border-collapse: collapse; font-size: 12px; }}
  table.summary th {{ background: var(--navy-800); color: rgba(255,255,255,0.85); padding: 8px 10px; text-align: left; font-size: 10px; letter-spacing: 0.3px; font-weight: 600; position: sticky; top: 0; }}
  table.summary td {{ padding: 7px 10px; border-bottom: 1px solid var(--border); vertical-align: middle; }}
  table.summary tr:nth-child(even) td {{ background: #FAFBFC; }}
  table.summary tr:hover td {{ background: var(--navy-50) !important; }}
  .scroll-wrap {{ overflow-x: auto; max-height: 500px; overflow-y: auto; border-radius: 6px; border: 1px solid var(--border); }}
  .info-box {{ background: var(--navy-50); border-left: 3px solid var(--navy-500); padding: 10px 14px; border-radius: 0 6px 6px 0; font-size: 12px; line-height: 1.8; color: #374151; margin-bottom: 12px; }}
  .footer {{ background: var(--navy-900); color: rgba(255,255,255,0.35); padding: 14px 48px; font-size: 11px; display: flex; justify-content: space-between; margin-top: 8px; }}
</style>
</head>
<body>

<div class="topbar">
  <span class="brand">Speech Sentiment Analysis</span>
  <span class="info">BATCH: {batch_name} | {n} calls analyzed</span>
</div>

<div class="header">
  <h1>배치 분석 종합 리포트</h1>
  <div class="sub">배치: {batch_name} | 분석 대상: {n}건 | STT 텍스트 + 음성 프로소디 비교 분석</div>
</div>

<div class="content">

<!-- KPI -->
<div class="kpi-row">
  <div class="kpi"><div class="kpi-lbl">분석 건수</div><div class="kpi-body"><div class="kpi-val">{n}</div></div></div>
  <div class="kpi"><div class="kpi-lbl">평균 NPS</div><div class="kpi-body"><div class="kpi-val">{nps_mean:.1f}<span>/10</span></div><div class="kpi-sub">std {nps_std:.1f}</div></div></div>
  <div class="kpi"><div class="kpi-lbl">총 감성 전환</div><div class="kpi-body"><div class="kpi-val">{total_transitions}</div><div class="kpi-sub" style="color:var(--green)">회복 {total_n2p} <span style="color:var(--gray-300)">|</span> <span style="color:var(--red)">악화 {total_p2n}</span></div></div></div>
  <div class="kpi"><div class="kpi-lbl">감성 회복률</div><div class="kpi-body"><div class="kpi-val">{recovery_rate:.0%}</div><div class="kpi-sub">긍정 전환 / 전체 전환</div></div></div>
</div>

<!-- 1. 단계별 평균 궤적 -->
<div class="section">
  <div class="sec-header"><div class="sec-num">1</div><div class="sec-title">단계별 평균 감성 궤적</div></div>
  <div class="grid-2">
    <div class="card">
      <div class="card-title">텍스트 vs 음성 평균 Valence</div>
      <div id="chart-traj"></div>
    </div>
    <div class="card">
      <div class="card-title">NPS 분포</div>
      <div id="chart-nps"></div>
    </div>
  </div>
  <div class="info-box">
    <b>해석</b> — 텍스트 감성은 상담 언어 특성상 0 근처의 작은 값이 정상입니다.
    음성 감성의 단계별 추이 변화가 더 크게 나타나며, 해결시도 이후 상승 추세가 높은 NPS와 상관됩니다.
  </div>
</div>

<!-- 2. 히트맵 -->
<div class="section">
  <div class="sec-header"><div class="sec-num">2</div><div class="sec-title">콜별 단계 감성 히트맵</div></div>
  <div class="card">
    <div class="card-title">텍스트 Valence 히트맵 (행=CNID, 열=단계, 색상=Valence)</div>
    <div id="chart-heatmap"></div>
  </div>
</div>

<!-- 3. 콜별 근거 테이블 -->
<div class="section">
  <div class="sec-header"><div class="sec-num">3</div><div class="sec-title">콜별 분석 근거 테이블</div></div>
  <div class="card" style="padding:0;overflow:hidden">
    <div class="scroll-wrap">
      <table class="summary">
        <thead>
          <tr>
            <th>CNID</th><th>NPS</th><th>시간</th>
            {"".join(f'<th>{s}</th>' for s in CALL_STAGES)}
            <th>전환</th><th>인사이트</th>
          </tr>
        </thead>
        <tbody>{call_rows}</tbody>
      </table>
    </div>
  </div>
  <div class="info-box">
    <b>분석 근거 읽는 법</b> — 각 행은 하나의 통화입니다.
    5단계 컬럼의 색상이 녹색일수록 긍정, 붉은색일수록 부정 감성입니다.
    인사이트 컬럼은 전환 분석 기반 자동 생성 요약이며, 개별 콜 리포트에서 상세 근거를 확인할 수 있습니다.
  </div>
</div>

</div><!-- /content -->

<div class="footer">
  <div>Speech Sentiment Analysis — Batch Report: {batch_name}</div>
  <div>Korean BERT + Audio Prosody | {n} calls</div>
</div>

<script>
  var cfg = {{responsive: true, displayModeBar: false}};
  var d1 = {traj_data}; Plotly.newPlot('chart-traj', d1.data, d1.layout, cfg);
  var d2 = {nps_dist}; Plotly.newPlot('chart-nps', d2.data, d2.layout, cfg);
  var d3 = {heatmap_data}; Plotly.newPlot('chart-heatmap', d3.data, d3.layout, cfg);
</script>
</body></html>"""

    if output_path is None:
        output_path = os.path.join(OUTPUT_DIR, f"batch_report_{batch_name}.html")

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

    return output_path
