"""
Speech Sentiment Analysis — Main Pipeline Runner

실행 방법:
    cd "c:/Users/User/Desktop/Speech Sentiment Analysis/analysis"
    python run.py [--sample N] [--no-audio] [--no-text] [--call CNID]

옵션:
    --sample N    : 분석할 콜 수 (기본: config.SAMPLE_N = 50)
    --no-audio    : 음성 분석 건너뜀 (빠른 텍스트 전용 실행)
    --no-text     : 텍스트 모델 건너뜀 (GT Valence만 사용)
    --call CNID   : 특정 콜 하나만 분석 + 궤적 시각화
"""
import os
import sys
import json
import argparse
import warnings
warnings.filterwarnings('ignore')

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import numpy as np
import pandas as pd

from config import OUTPUT_DIR, DATA_DIR, FIGURE_DIR, CALL_STAGES, SAMPLE_N
from data_loader import load_call_data, build_wav_index, merge_wav_paths, sample_calls, get_gt_stage_valence
from emotion_pipeline import analyze_call, run_batch_analysis, results_to_dataframe, extract_audio_feature_matrix
from modality_comparison import run_comparison, stage_correlation_analysis, trajectory_statistics, get_feature_importance
from utils.viz_utils import (plot_emotion_trajectory, plot_stage_distribution,
                              plot_modality_comparison, plot_feature_importance,
                              plot_valence_heatmap, plot_scatter_text_vs_audio)


def parse_args():
    parser = argparse.ArgumentParser(description="Speech Sentiment Analysis Pipeline")
    parser.add_argument('--sample', type=int, default=SAMPLE_N,
                        help=f'분석 콜 수 (기본: {SAMPLE_N})')
    parser.add_argument('--no-audio', action='store_true', help='음성 분석 스킵')
    parser.add_argument('--no-text',  action='store_true', help='텍스트 모델 스킵 (GT 사용)')
    parser.add_argument('--call',     type=str, default=None,
                        help='특정 CNID 하나만 분석')
    parser.add_argument('--batch',    type=str, default=None,
                        help='배치 폴더 분석 (예: 1_100, 101_200)')
    parser.add_argument('--all',      action='store_true',
                        help='전체 594개 콜 분석 + 종합 리포트')
    return parser.parse_args()


def run_single_call(cnid: str, df: pd.DataFrame,
                     run_text: bool, run_audio: bool):
    """단일 콜 상세 분석 + 궤적 시각화."""
    rows = df[df['CNID'] == cnid]
    if rows.empty:
        print(f"[오류] CNID {cnid} 를 데이터에서 찾을 수 없습니다.")
        return

    row = rows.iloc[0]
    print(f"\n{'='*60}")
    print(f"콜 분석: {cnid}")
    print(f"  통화 시간: {row.get('duration_sec', 0):.0f}초")
    print(f"  NPS: {row.get('NPS', 'N/A')}")
    print(f"  제품: {row.get('상담번호: 제품명(중)', 'N/A')}")
    print(f"  WAV: {row.get('wav_path', '없음')}")
    print(f"{'='*60}")

    res = analyze_call(row, run_text=run_text, run_audio=run_audio)

    if res['error']:
        print(f"[경고] {res['error']}")

    # 발화 단위 결과 출력
    print(f"\n발화 단위 분석 ({len(res['turns'])} 턴):")
    print(f"{'idx':>3} {'화자':^5} {'시작':>6} {'종료':>6} {'단계':^6} {'텍스트Valence':>12} {'음성Valence':>10}  텍스트(앞 40자)")
    print("-" * 100)
    for t in res['turns']:
        tv = f"{t.get('text_valence', 0):.3f}" if t.get('text_valence') is not None else '  N/A'
        av = f"{t.get('audio_valence', 0):.3f}" if t.get('audio_valence') is not None else '  N/A'
        print(f"{t['turn_idx']:>3} {t['speaker']:^5} "
              f"{t.get('start_sec',0):>6.1f} {t.get('end_sec',0):>6.1f} "
              f"{t.get('stage',''):^6} {tv:>12} {av:>10}  "
              f"{t['text'][:40]}")

    # 단계별 집계 출력
    print(f"\n단계별 감성 집계:")
    print(f"{'단계':^8} {'GT Valence':>10} {'텍스트':>10} {'음성':>10}")
    print("-" * 45)
    gt = get_gt_stage_valence(row)
    for stage in CALL_STAGES:
        g  = f"{gt.get(stage, 0):.3f}"
        tv = f"{res['text_stage_valence'].get(stage, 0) or 0:.3f}"
        av = f"{res['audio_stage_valence'].get(stage, 0) or 0:.3f}"
        print(f"{stage:^8} {g:>10} {tv:>10} {av:>10}")

    # 궤적 시각화
    os.makedirs(FIGURE_DIR, exist_ok=True)
    fig = plot_emotion_trajectory(
        res['turns'], cnid,
        total_dur=res['duration_sec'],
        gt_stage_valence=gt,
        save=True
    )
    out_path = os.path.join(FIGURE_DIR, f"trajectory_{cnid}.html")
    print(f"\n[시각화] {out_path}")
    fig.show()


def run_full_pipeline(df: pd.DataFrame, n_sample: int,
                       run_text: bool, run_audio: bool):
    """전체 배치 분석 파이프라인."""

    print(f"\n{'='*60}")
    print(f"전체 파이프라인 시작 | 대상: {n_sample} 콜")
    print(f"  텍스트 분석: {'ON' if run_text else 'OFF'}")
    print(f"  음성 분석:   {'ON' if run_audio else 'OFF'}")
    print(f"{'='*60}\n")

    df_sample = sample_calls(df, n=n_sample)
    print(f"[샘플] {len(df_sample)} 콜 선택됨 (WAV 보유)")

    # ── Step 1: 배치 분석 ──────────────────────────────────────────
    print("\n[Step 1] 발화 분석 실행 중...")
    results = run_batch_analysis(df_sample, run_text=run_text, run_audio=run_audio)

    errors = [r for r in results if r.get('error')]
    print(f"  완료: {len(results)} 건 / 오류: {len(errors)} 건")

    # ── Step 2: 결과 DataFrame ─────────────────────────────────────
    print("\n[Step 2] 결과 정리...")
    df_result = results_to_dataframe(results, df_sample)
    out_csv = os.path.join(DATA_DIR, "call_analysis_results.csv")
    df_result.to_csv(out_csv, index=False, encoding='utf-8-sig')
    print(f"  저장: {out_csv}")

    # ── Step 3: 음향 특징 행렬 ────────────────────────────────────
    audio_X, audio_feat_names = extract_audio_feature_matrix(results)
    print(f"  음향 특징 행렬: {audio_X.shape}")

    # ── Step 4: 모달리티 비교 ─────────────────────────────────────
    print("\n[Step 4] 모달리티 비교 (NPS 예측 기반)...")
    for target in ['nps', 'consultant_score']:
        if df_result[target].notna().sum() < 10:
            continue
        print(f"\n  타겟: {target}")
        comparison_results = run_comparison(df_result, audio_X, audio_feat_names,
                                             target=target, model_name='ridge')

        # 결과 저장
        out_json = os.path.join(DATA_DIR, f"modality_comparison_{target}.json")
        with open(out_json, 'w', encoding='utf-8') as f:
            json.dump(comparison_results, f, ensure_ascii=False, indent=2)

        # 시각화
        fig_comp = plot_modality_comparison(comparison_results)
        fig_comp.write_html(os.path.join(FIGURE_DIR, f"modality_comparison_{target}.html"))

    # ── Step 5: 단계별 상관 분석 ──────────────────────────────────
    print("\n[Step 5] 단계별 상관 분석...")
    df_corr = stage_correlation_analysis(df_result)
    print(df_corr.to_string(index=False))
    df_corr.to_csv(os.path.join(DATA_DIR, "stage_correlation.csv"),
                   index=False, encoding='utf-8-sig')

    # ── Step 6: 감정 변화 추이 통계 ──────────────────────────────
    print("\n[Step 6] 감정 변화 추이 통계...")
    df_traj = trajectory_statistics(df_result)
    print(df_traj.to_string(index=False))
    df_traj.to_csv(os.path.join(DATA_DIR, "trajectory_statistics.csv"),
                   index=False, encoding='utf-8-sig')

    # ── Step 7: 시각화 ───────────────────────────────────────────
    print("\n[Step 7] 시각화 생성...")

    # 단계별 감성 분포
    plot_stage_distribution(df_result)
    print(f"  저장: {FIGURE_DIR}/stage_distribution.html")

    # 감성 히트맵 (텍스트)
    text_cols  = {s: f"text_{s}"  for s in CALL_STAGES}
    audio_cols = {s: f"audio_{s}" for s in CALL_STAGES}
    df_hm_text  = df_result.set_index('cnid')[[f"text_{s}"  for s in CALL_STAGES]].rename(columns={f"text_{s}": s for s in CALL_STAGES})
    df_hm_audio = df_result.set_index('cnid')[[f"audio_{s}" for s in CALL_STAGES]].rename(columns={f"audio_{s}": s for s in CALL_STAGES})
    plot_valence_heatmap(df_hm_text,  modality="텍스트")
    plot_valence_heatmap(df_hm_audio, modality="음성")
    print(f"  저장: {FIGURE_DIR}/heatmap_텍스트.html")
    print(f"  저장: {FIGURE_DIR}/heatmap_음성.html")

    # 텍스트 vs 음성 산점도 (첫 단계)
    if 'text_초기' in df_result.columns and 'audio_초기' in df_result.columns:
        plot_scatter_text_vs_audio(df_result, 'text_초기', 'audio_초기', color_col='product')
        print(f"  저장: {FIGURE_DIR}/scatter_text_vs_audio.html")

    # 특징 중요도 (음향 특징)
    if len(audio_X) > 0:
        y_nps = df_result['nps'].values.astype(float)
        feat_names, importances = get_feature_importance(audio_X, y_nps, audio_feat_names)
        plot_feature_importance(feat_names, importances,
                                 title="음향 특징 중요도 (NPS 예측)",
                                 filename="feature_importance_nps.html")
        print(f"  저장: {FIGURE_DIR}/feature_importance_nps.html")

    # ── 샘플 3개 콜 개별 궤적 시각화 ─────────────────────────────
    print("\n[샘플 궤적] 개별 콜 감정 궤적 시각화 (3건)...")
    sample_results = [r for r in results if not r.get('error')][:3]
    for res in sample_results:
        cnid     = res['cnid']
        meta_row = df_sample[df_sample['CNID'] == cnid]
        gt       = get_gt_stage_valence(meta_row.iloc[0]) if not meta_row.empty else {}
        plot_emotion_trajectory(
            res['turns'], cnid,
            total_dur=res['duration_sec'],
            gt_stage_valence=gt,
            save=True
        )
        print(f"  저장: {FIGURE_DIR}/trajectory_{cnid}.html")

    print(f"\n{'='*60}")
    print("분석 완료!")
    print(f"  결과 디렉토리: {OUTPUT_DIR}")
    print(f"  데이터: {DATA_DIR}")
    print(f"  시각화: {FIGURE_DIR}")
    print(f"{'='*60}")


def main():
    os.makedirs(DATA_DIR,   exist_ok=True)
    os.makedirs(FIGURE_DIR, exist_ok=True)

    args = parse_args()

    # 데이터 로딩
    df_call  = load_call_data()
    wav_idx  = build_wav_index()
    df_call  = merge_wav_paths(df_call, wav_idx)

    run_text  = not args.no_text
    run_audio = not args.no_audio

    if args.call:
        # 단일 콜 분석 모드
        run_single_call(args.call, df_call, run_text, run_audio)
    elif args.batch:
        # 배치 폴더 분석 모드
        from batch_runner import run_batch_by_folder
        run_batch_by_folder(args.batch, df_call, wav_idx,
                             run_text=run_text, run_audio=run_audio)
    elif args.all:
        # 전체 콜 분석 모드
        from batch_runner import run_all_calls
        run_all_calls(df_call, run_text=run_text, run_audio=run_audio)
    else:
        # 기존 파이프라인
        run_full_pipeline(df_call, n_sample=args.sample,
                           run_text=run_text, run_audio=run_audio)


if __name__ == "__main__":
    main()
