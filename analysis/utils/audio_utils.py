"""
Audio Utility Functions
- GSM WAV → PCM 변환 (ffmpeg)
- librosa 기반 음향 특징 추출
- 발화 구간별 특징 집계
"""
import os
import io
import subprocess
import tempfile
import numpy as np
import librosa
import imageio_ffmpeg

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from config import SAMPLE_RATE, FRAME_MS, HOP_MS, N_MFCC, MIN_PITCH_HZ, MAX_PITCH_HZ

FFMPEG = imageio_ffmpeg.get_ffmpeg_exe()


# ── 오디오 로딩 ───────────────────────────────────────────────────────────────

def load_wav(wav_path: str, sr: int = SAMPLE_RATE) -> tuple[np.ndarray, int]:
    """
    GSM-MS WAV → float32 PCM (16kHz, mono) 로드.
    ffmpeg으로 디코딩 후 librosa 리샘플링.
    """
    # ffmpeg으로 raw PCM s16le 변환
    cmd = [
        FFMPEG, "-y", "-i", wav_path,
        "-ac", "1",            # mono
        "-ar", str(sr),        # 목표 샘플레이트
        "-f", "s16le",         # raw signed 16-bit little-endian
        "-loglevel", "error",
        "pipe:1"
    ]
    proc = subprocess.run(cmd, capture_output=True)
    if proc.returncode != 0:
        raise RuntimeError(f"ffmpeg 변환 실패: {proc.stderr.decode()}")

    audio = np.frombuffer(proc.stdout, dtype=np.int16).astype(np.float32) / 32768.0
    return audio, sr


def get_wav_duration(wav_path: str) -> float:
    """ffprobe 없이 ffmpeg으로 파일 길이(초) 추출."""
    cmd = [FFMPEG, "-i", wav_path, "-hide_banner"]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    # ffmpeg는 출력 파일 없으면 에러지만 stderr에 Duration을 출력함
    for line in (proc.stderr + proc.stdout).split("\n"):
        if "Duration:" in line:
            dur_str = line.strip().split("Duration:")[1].split(",")[0].strip()
            h, m, s = dur_str.split(":")
            return float(h) * 3600 + float(m) * 60 + float(s)
    return 0.0


# ── 음향 특징 추출 ─────────────────────────────────────────────────────────────

def extract_segment_features(audio: np.ndarray, sr: int,
                              start_sec: float, end_sec: float) -> dict:
    """
    주어진 구간 [start_sec, end_sec]의 음향 특징 추출.

    Returns dict with keys:
        f0_mean, f0_std, f0_slope          ← 기본 주파수 (피치)
        energy_mean, energy_std            ← 에너지 (음량)
        zcr_mean                           ← Zero Crossing Rate (목소리 거칠기)
        mfcc_1..N_MFCC                     ← MFCC 계수 (음색)
        spectral_centroid_mean             ← 스펙트럼 무게중심
        speech_rate                        ← 발화 속도 추정 (에너지 피크/초)
        duration                           ← 구간 길이(초)
        voiced_ratio                       ← 유성 구간 비율
    """
    start_samp = int(start_sec * sr)
    end_samp   = int(end_sec   * sr)
    segment    = audio[start_samp:end_samp]

    if len(segment) < sr * 0.1:   # 0.1초 미만은 무시
        return _empty_features()

    frame_len = int(FRAME_MS * sr / 1000)
    hop_len   = int(HOP_MS   * sr / 1000)

    feats = {}
    dur   = (end_sec - start_sec)
    feats["duration"] = dur

    # ── 1. 기본 주파수 (F0 / 피치) ─────────────────────────────────────────
    try:
        f0, voiced_flag, _ = librosa.pyin(
            segment, fmin=MIN_PITCH_HZ, fmax=MAX_PITCH_HZ,
            sr=sr, frame_length=frame_len, hop_length=hop_len
        )
        voiced_f0 = f0[voiced_flag]
        feats["f0_mean"]   = float(np.nanmean(voiced_f0)) if len(voiced_f0) > 0 else 0.0
        feats["f0_std"]    = float(np.nanstd(voiced_f0))  if len(voiced_f0) > 0 else 0.0
        # F0 기울기: 후반부 평균 - 전반부 평균 (상승=긍정 추세)
        if len(voiced_f0) >= 4:
            half = len(voiced_f0) // 2
            feats["f0_slope"] = float(np.nanmean(voiced_f0[half:]) - np.nanmean(voiced_f0[:half]))
        else:
            feats["f0_slope"] = 0.0
        feats["voiced_ratio"] = float(np.mean(voiced_flag))
    except Exception:
        feats["f0_mean"] = feats["f0_std"] = feats["f0_slope"] = 0.0
        feats["voiced_ratio"] = 0.0

    # ── 2. 에너지 (RMS) ────────────────────────────────────────────────────
    rms = librosa.feature.rms(y=segment, frame_length=frame_len, hop_length=hop_len)[0]
    feats["energy_mean"] = float(np.mean(rms))
    feats["energy_std"]  = float(np.std(rms))
    # 에너지 기울기 (증가=감정 고조)
    feats["energy_slope"] = float(np.mean(rms[len(rms)//2:]) - np.mean(rms[:len(rms)//2]))

    # ── 3. Zero Crossing Rate (목소리 거칠기/긴장도) ──────────────────────
    zcr = librosa.feature.zero_crossing_rate(segment, frame_length=frame_len, hop_length=hop_len)[0]
    feats["zcr_mean"] = float(np.mean(zcr))

    # ── 4. MFCC (음색/음질 - 감정의 주요 특징) ──────────────────────────
    mfcc = librosa.feature.mfcc(y=segment, sr=sr, n_mfcc=N_MFCC,
                                  n_fft=frame_len, hop_length=hop_len)
    for i in range(N_MFCC):
        feats[f"mfcc_{i+1}"] = float(np.mean(mfcc[i]))

    # ── 5. 스펙트럼 무게중심 (밝기/긴장 지표) ─────────────────────────────
    sc = librosa.feature.spectral_centroid(y=segment, sr=sr,
                                            n_fft=frame_len, hop_length=hop_len)[0]
    feats["spectral_centroid_mean"] = float(np.mean(sc))

    # ── 6. 발화 속도 추정 (에너지 피크 카운트 / 초) ───────────────────────
    from scipy.signal import find_peaks
    peaks, _ = find_peaks(rms, height=np.mean(rms), distance=int(0.15 * sr / hop_len))
    feats["speech_rate"] = len(peaks) / dur if dur > 0 else 0.0

    # ── 7. Jitter (피치 불안정성 — 감정 동요 지표) ─────────────────────
    try:
        voiced_f0_clean = voiced_f0[~np.isnan(voiced_f0)] if 'voiced_f0' in dir() else np.array([])
        if len(voiced_f0_clean) >= 3:
            periods = 1.0 / voiced_f0_clean
            jitter_abs = np.mean(np.abs(np.diff(periods)))
            jitter_rel = jitter_abs / np.mean(periods) * 100  # %
            feats["jitter"] = float(jitter_rel)
        else:
            feats["jitter"] = 0.0
    except Exception:
        feats["jitter"] = 0.0

    # ── 8. Shimmer (에너지 불안정성 — 음성 떨림 지표) ──────────────────
    try:
        if len(rms) >= 3:
            shimmer_abs = np.mean(np.abs(np.diff(rms)))
            shimmer_rel = shimmer_abs / np.mean(rms) * 100 if np.mean(rms) > 0 else 0
            feats["shimmer"] = float(shimmer_rel)
        else:
            feats["shimmer"] = 0.0
    except Exception:
        feats["shimmer"] = 0.0

    # ── 9. HNR (조화 대 잡음비 — 음성 품질/긴장도) ────────────────────
    try:
        autocorr = np.correlate(segment, segment, mode='full')
        autocorr = autocorr[len(autocorr)//2:]
        if len(autocorr) > sr // MIN_PITCH_HZ:
            peak_range = autocorr[sr // MAX_PITCH_HZ : sr // MIN_PITCH_HZ]
            if len(peak_range) > 0 and autocorr[0] > 0:
                r_max = np.max(peak_range)
                hnr = 10 * np.log10(r_max / (autocorr[0] - r_max + 1e-10)) if r_max < autocorr[0] else 0
                feats["hnr"] = float(np.clip(hnr, -10, 40))
            else:
                feats["hnr"] = 0.0
        else:
            feats["hnr"] = 0.0
    except Exception:
        feats["hnr"] = 0.0

    # ── 10. 발화 내 전반/후반 비교 (감정 방향 변화) ───────────────────
    half_samp = len(segment) // 2
    if half_samp > sr * 0.05:  # 최소 0.05초
        seg_1st = segment[:half_samp]
        seg_2nd = segment[half_samp:]

        rms_1st = float(np.sqrt(np.mean(seg_1st**2)))
        rms_2nd = float(np.sqrt(np.mean(seg_2nd**2)))
        feats["energy_direction"] = rms_2nd - rms_1st  # + = 후반 에너지 증가

        try:
            f0_1, vf_1, _ = librosa.pyin(seg_1st, fmin=MIN_PITCH_HZ, fmax=MAX_PITCH_HZ, sr=sr)
            f0_2, vf_2, _ = librosa.pyin(seg_2nd, fmin=MIN_PITCH_HZ, fmax=MAX_PITCH_HZ, sr=sr)
            f0_1_mean = float(np.nanmean(f0_1[vf_1])) if np.any(vf_1) else 0
            f0_2_mean = float(np.nanmean(f0_2[vf_2])) if np.any(vf_2) else 0
            feats["f0_direction"] = f0_2_mean - f0_1_mean  # + = 후반 피치 상승
        except Exception:
            feats["f0_direction"] = 0.0
    else:
        feats["energy_direction"] = 0.0
        feats["f0_direction"] = 0.0

    return feats


def _empty_features() -> dict:
    """구간이 너무 짧을 때 반환하는 기본값 dict."""
    feats = {"duration": 0.0, "f0_mean": 0.0, "f0_std": 0.0, "f0_slope": 0.0,
             "voiced_ratio": 0.0, "energy_mean": 0.0, "energy_std": 0.0,
             "energy_slope": 0.0, "zcr_mean": 0.0,
             "spectral_centroid_mean": 0.0, "speech_rate": 0.0,
             "jitter": 0.0, "shimmer": 0.0, "hnr": 0.0,
             "energy_direction": 0.0, "f0_direction": 0.0}
    for i in range(N_MFCC):
        feats[f"mfcc_{i+1}"] = 0.0
    return feats


def extract_stage_features(audio: np.ndarray, sr: int,
                            total_dur: float, stage_ratios: list) -> list[dict]:
    """
    전체 오디오를 상담 단계 비율로 분할해 각 단계의 음향 특징 추출.
    stage_ratios: [초기, 탐색, 해결시도, 결과제시, 종료] 합=1.0
    """
    stage_feats = []
    cursor = 0.0
    for ratio in stage_ratios:
        start = cursor
        end   = min(cursor + total_dur * ratio, total_dur)
        feats = extract_segment_features(audio, sr, start, end)
        stage_feats.append(feats)
        cursor = end
    return stage_feats


def audio_to_valence(feats: dict) -> float:
    """
    음향 특징 → 단순 휴리스틱 Valence 추정 [-1, +1]
    (모델 학습 전 baseline으로 사용)

    규칙:
    - 높은 에너지 + 높은 F0 = 감정 고조 (분노 or 기쁨)
    - F0 기울기 상승 = 긍정/안도 추세
    - ZCR 높음 = 긴장/불안 → Valence 하락
    - 에너지 기울기 = 감정 변화 방향
    """
    v = 0.0

    # F0 기울기 기여 (정규화: 200Hz 기준)
    f0_slope_norm = np.clip(feats.get("f0_slope", 0) / 50.0, -1, 1)
    v += 0.25 * f0_slope_norm

    # 에너지 기울기 기여
    energy_slope_norm = np.clip(feats.get("energy_slope", 0) / 0.05, -1, 1)
    v += 0.15 * energy_slope_norm

    # ZCR 높으면 불안/긴장 → 음수
    zcr = feats.get("zcr_mean", 0)
    zcr_norm = np.clip((zcr - 0.1) / 0.2, -1, 1)
    v -= 0.10 * zcr_norm

    # voiced_ratio 높을수록 발화 활발 → 약한 긍정 신호
    v += 0.10 * (feats.get("voiced_ratio", 0.5) - 0.5)

    return float(np.clip(v, -1.0, 1.0))
