"""
Modality Comparison Analysis
- STT-only vs STT+Audio vs Audio-only 예측 성능 비교
- NPS / 컨설턴트 만족도 예측 기반 평가
- Cross-validation + Feature Importance
"""
import os
import sys
import numpy as np
import pandas as pd
from scipy.stats import pearsonr, spearmanr
from sklearn.preprocessing import StandardScaler
from sklearn.linear_model import Ridge
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
from sklearn.model_selection import cross_val_score, KFold
from sklearn.metrics import mean_squared_error, r2_score
from sklearn.pipeline import Pipeline

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import CALL_STAGES, CV_FOLDS, RANDOM_SEED

TARGET_COLS = ['nps', 'consultant_score']


# ── 특징 행렬 구성 ────────────────────────────────────────────────────────────

def build_feature_sets(df: pd.DataFrame,
                        audio_feat_matrix: np.ndarray,
                        audio_feat_names:  list) -> dict:
    """
    세 가지 특징 세트 구성:
    - 'text':   텍스트 감성 단계별 5개 값
    - 'audio':  음향 특징 행렬
    - 'fusion': text + audio 결합
    """
    text_cols  = [f"text_{s}"  for s in CALL_STAGES]
    audio_cols = [f"audio_{s}" for s in CALL_STAGES]

    X_text     = df[text_cols].fillna(0).values
    X_audio_v  = df[audio_cols].fillna(0).values  # 음성 Valence (5단계)

    feature_sets = {
        'STT':    (X_text,    text_cols),
        'Audio':  (X_audio_v, audio_cols),
    }

    # Fusion: 텍스트 Valence + 음성 Valence + 음향 특징
    if audio_feat_matrix is not None and len(audio_feat_matrix) == len(df):
        X_fusion = np.hstack([X_text, X_audio_v, audio_feat_matrix])
        fusion_names = text_cols + audio_cols + audio_feat_names
        feature_sets['Fusion'] = (X_fusion, fusion_names)

    return feature_sets


# ── 교차검증 평가 ─────────────────────────────────────────────────────────────

def evaluate_modality(X: np.ndarray, y: np.ndarray,
                       model_name: str = 'ridge',
                       cv_folds: int = CV_FOLDS) -> dict:
    """
    단일 특징 세트로 타겟 예측 성능 교차 검증.

    Returns:
        {'r2': float, 'rmse': float, 'corr': float,
         'r2_std': float, 'rmse_std': float}
    """
    valid_mask = ~np.isnan(y)
    X_v = X[valid_mask]
    y_v = y[valid_mask]

    if len(y_v) < cv_folds * 2:
        return {'r2': np.nan, 'rmse': np.nan, 'corr': np.nan}

    if model_name == 'ridge':
        model = Pipeline([('scaler', StandardScaler()),
                          ('reg',    Ridge(alpha=1.0))])
    elif model_name == 'rf':
        model = Pipeline([('scaler', StandardScaler()),
                          ('reg',    RandomForestRegressor(
                              n_estimators=100, random_state=RANDOM_SEED, n_jobs=-1))])
    else:
        model = Pipeline([('scaler', StandardScaler()),
                          ('reg',    GradientBoostingRegressor(
                              n_estimators=100, random_state=RANDOM_SEED))])

    kf  = KFold(n_splits=cv_folds, shuffle=True, random_state=RANDOM_SEED)
    r2_scores, rmse_scores, corr_scores = [], [], []

    for train_idx, test_idx in kf.split(X_v):
        X_tr, X_te = X_v[train_idx], X_v[test_idx]
        y_tr, y_te = y_v[train_idx], y_v[test_idx]

        model.fit(X_tr, y_tr)
        y_pred = model.predict(X_te)

        r2_scores.append(r2_score(y_te, y_pred))
        rmse_scores.append(np.sqrt(mean_squared_error(y_te, y_pred)))
        if len(y_te) > 2:
            r, _ = pearsonr(y_te, y_pred)
            corr_scores.append(r)

    return {
        'r2':       float(np.mean(r2_scores)),
        'r2_std':   float(np.std(r2_scores)),
        'rmse':     float(np.mean(rmse_scores)),
        'rmse_std': float(np.std(rmse_scores)),
        'corr':     float(np.mean(corr_scores)) if corr_scores else np.nan,
    }


def run_comparison(df: pd.DataFrame,
                    audio_feat_matrix: np.ndarray,
                    audio_feat_names:  list,
                    target: str = 'nps',
                    model_name: str = 'ridge') -> dict:
    """
    전체 모달리티 비교 실행.

    Returns:
        {모달리티명: 성능지표 dict}
    """
    feature_sets = build_feature_sets(df, audio_feat_matrix, audio_feat_names)
    y = df[target].values.astype(float)

    results = {}
    for modality, (X, _) in feature_sets.items():
        print(f"  [{modality}] {X.shape[1]}개 특징 → 교차검증 중...")
        res = evaluate_modality(X, y, model_name=model_name)
        results[modality] = res
        print(f"    R²={res['r2']:.3f}±{res.get('r2_std',0):.3f}  "
              f"RMSE={res['rmse']:.3f}  r={res.get('corr',0):.3f}")
    return results


# ── 특징 중요도 분석 ──────────────────────────────────────────────────────────

def get_feature_importance(X: np.ndarray, y: np.ndarray,
                             feature_names: list) -> tuple[list, np.ndarray]:
    """
    RandomForest 특징 중요도 반환.
    """
    valid_mask = ~np.isnan(y)
    X_v, y_v   = X[valid_mask], y[valid_mask]

    scaler = StandardScaler()
    X_s    = scaler.fit_transform(X_v)

    rf = RandomForestRegressor(n_estimators=200, random_state=RANDOM_SEED, n_jobs=-1)
    rf.fit(X_s, y_v)

    return feature_names, rf.feature_importances_


# ── 단계별 Valence 상관 분석 ──────────────────────────────────────────────────

def stage_correlation_analysis(df: pd.DataFrame) -> pd.DataFrame:
    """
    각 단계에서 텍스트/음성 Valence의 상관관계 분석.
    Returns DataFrame with columns: stage, text_audio_corr, text_nps_corr, audio_nps_corr
    """
    rows = []
    for stage in CALL_STAGES:
        text_col  = f"text_{stage}"
        audio_col = f"audio_{stage}"
        gt_col    = f"gt_{stage}"

        sub = df[[text_col, audio_col, gt_col, 'nps']].dropna()
        if len(sub) < 5:
            continue

        row = {'stage': stage}

        # 텍스트 vs 음성 감성 상관
        r, p = pearsonr(sub[text_col], sub[audio_col])
        row['text_audio_r'] = round(r, 3)
        row['text_audio_p'] = round(p, 3)

        # 텍스트 감성 vs GT Valence
        r, p = pearsonr(sub[text_col], sub[gt_col])
        row['text_gt_r']  = round(r, 3)
        row['text_gt_p']  = round(p, 3)

        # 음성 감성 vs GT Valence
        r, p = pearsonr(sub[audio_col], sub[gt_col])
        row['audio_gt_r'] = round(r, 3)
        row['audio_gt_p'] = round(p, 3)

        # 텍스트/음성 감성 vs NPS
        r_t, _ = pearsonr(sub[text_col],  sub['nps'])
        r_a, _ = pearsonr(sub[audio_col], sub['nps'])
        row['text_nps_r']  = round(r_t, 3)
        row['audio_nps_r'] = round(r_a, 3)

        rows.append(row)

    return pd.DataFrame(rows)


# ── 감정 변화 추이 통계 ───────────────────────────────────────────────────────

def trajectory_statistics(df: pd.DataFrame) -> pd.DataFrame:
    """
    단계별 감성 평균/표준편차 집계 (텍스트/음성/GT 각각).
    """
    rows = []
    for stage in CALL_STAGES:
        for prefix in ['text', 'audio', 'gt']:
            col = f"{prefix}_{stage}"
            if col not in df.columns:
                continue
            vals = df[col].dropna()
            rows.append({
                'stage':    stage,
                'modality': prefix,
                'mean':     round(float(vals.mean()), 4),
                'std':      round(float(vals.std()),  4),
                'median':   round(float(vals.median()), 4),
                'n':        len(vals),
            })
    return pd.DataFrame(rows)
