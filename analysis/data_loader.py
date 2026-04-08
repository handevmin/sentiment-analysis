"""
Data Loader
- Call_Data.xlsx 로딩 및 WAV 파일 매핑
- CNID ↔ WAV 경로 인덱스 구축
"""
import os
import sys
import glob
import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from config import EXCEL_CALL, SPEECH_BATCHES, SAMPLE_N, RANDOM_SEED, CALL_STAGES


def load_call_data() -> pd.DataFrame:
    """
    Call_Data.xlsx 또는 Call_Data_시트1.csv 로딩.
    CNID를 str로 변환하고, 통화 시간을 초(float)로 파싱.
    """
    from config import RAW_DATA_DIR
    csv_fallback = os.path.join(RAW_DATA_DIR, "Call_Data_시트1.csv")
    
    if os.path.exists(EXCEL_CALL):
        print("[데이터 로딩] Call_Data.xlsx ...")
        df = pd.read_excel(EXCEL_CALL, sheet_name=0, engine='openpyxl')
    elif os.path.exists(csv_fallback):
        print("[데이터 로딩] Call_Data_시트1.csv (xlsx 없음, CSV fallback) ...")
        df = pd.read_csv(csv_fallback, encoding='utf-8-sig')
    else:
        raise FileNotFoundError(f"Call_Data.xlsx 또는 Call_Data_시트1.csv 를 찾을 수 없습니다.")

    # 컬럼 정리
    df.columns = df.columns.str.strip()

    # CNID str 변환 (WAV 파일명 매핑용)
    df['CNID'] = df['CNID'].astype(str).str.strip().str.split('.').str[0]

    # 통화 시간 초 변환
    from utils.text_utils import parse_duration_str
    df['duration_sec'] = df['통화초'].apply(parse_duration_str)

    # 기존 Valence 컬럼 정리
    for stage in CALL_STAGES:
        col = f"{stage}_Valence"
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

    print(f"  총 {len(df)} 건 로딩 완료")
    return df


def build_wav_index() -> dict:
    """
    CNID → WAV 파일 절대 경로 딕셔너리 구축.
    """
    index = {}
    for batch_dir in SPEECH_BATCHES:
        if not os.path.isdir(batch_dir):
            continue
        for wav_path in glob.glob(os.path.join(batch_dir, "*.wav")):
            cnid = os.path.splitext(os.path.basename(wav_path))[0]
            index[cnid] = wav_path
    print(f"[WAV 인덱스] {len(index)} 개 파일 인덱싱 완료")
    return index


def merge_wav_paths(df: pd.DataFrame, wav_index: dict) -> pd.DataFrame:
    """
    DataFrame에 'wav_path' 컬럼 추가.
    매핑되지 않는 CNID는 NaN.
    """
    df = df.copy()
    df['wav_path'] = df['CNID'].map(wav_index)
    matched = df['wav_path'].notna().sum()
    print(f"[WAV 매핑] {matched} / {len(df)} 건 매핑됨")
    return df


def sample_calls(df: pd.DataFrame, n: int = None,
                 require_wav: bool = True) -> pd.DataFrame:
    """
    분석용 샘플 추출.
    require_wav=True이면 WAV 파일이 있는 건만 포함.
    """
    if require_wav:
        df = df[df['wav_path'].notna()].reset_index(drop=True)

    if n is None or n >= len(df):
        return df

    return df.sample(n=n, random_state=RANDOM_SEED).reset_index(drop=True)


def get_gt_stage_valence(row: pd.Series) -> dict:
    """
    단일 행에서 기존 5단계 Valence 어노테이션 딕셔너리 반환.
    """
    return {
        stage: float(row.get(f"{stage}_Valence", 0.0) or 0.0)
        for stage in CALL_STAGES
    }
