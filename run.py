#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
학교급식 식중독 분석 원클릭 스크립트

사용 방법
---------
1. 이 파일 이름을 run.py로 둡니다.
2. 같은 폴더에 아래 두 파일을 둡니다.
   - 1.xlsx : 학교안전사고 데이터
   - 2.xlsx : 학교안전사고 보상 데이터
3. 터미널에서 실행합니다.

   python3 -m pip install pandas openpyxl matplotlib
   python3 run.py

결과 폴더
---------
food_poisoning_result/
├── 01_식중독_학생기록.xlsx
├── 02_추정_집단발생사건.xlsx
├── 03_분석요약.xlsx
└── charts/
    ├── 01_연도별_추이.png
    ├── 02_지역별_피해규모.png
    ├── 03_지역별_추정사건수.png
    ├── 04_월별_피해규모.png
    └── 05_학교급별_피해규모.png

주의
----
- 원본 데이터에 학교명이 없다면 같은 지역·학교급·날짜·시간·장소·활동을
  하나의 집단발생 사건으로 추정합니다.
- 한 행은 집단발생 사건 1건이 아니라 학생별 사고·청구 기록일 수 있습니다.
"""

import re
import sys
from pathlib import Path

try:
    import pandas as pd
    import matplotlib.pyplot as plt
except ImportError:
    print("필수 패키지가 없습니다.")
    print("아래 명령을 먼저 실행하세요:")
    print("python3 -m pip install pandas openpyxl matplotlib")
    raise

BASE = Path(__file__).resolve().parent
DATA_FILE = BASE / "1.xlsx"
COMP_FILE = BASE / "2.xlsx"
OUT_DIR = BASE / "food_poisoning_result"
CHART_DIR = OUT_DIR / "charts"

FONT_CANDIDATES = [
    "AppleGothic",
    "Malgun Gothic",
    "NanumGothic",
    "Noto Sans CJK KR",
    "DejaVu Sans",
]


def set_korean_font():
    try:
        from matplotlib import font_manager, rcParams
        available = {f.name for f in font_manager.fontManager.ttflist}
        for name in FONT_CANDIDATES:
            if name in available:
                rcParams["font.family"] = name
                break
        rcParams["axes.unicode_minus"] = False
    except Exception:
        pass


def text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def detect_header_row(path: Path, sheet_name):
    preview = pd.read_excel(path, sheet_name=sheet_name, header=None, nrows=20)
    for idx, row in preview.iterrows():
        joined = " ".join(text(v) for v in row.tolist())
        if "사고형태" in joined and ("학교급" in joined or "지역" in joined):
            return idx
    return 0


def load_workbook(path: Path):
    print(f"\n[파일 읽기] {path.name}")
    xls = pd.ExcelFile(path)
    frames = []

    for sheet in xls.sheet_names:
        try:
            header_row = detect_header_row(path, sheet)
            df = pd.read_excel(path, sheet_name=sheet, header=header_row)
            df = df.dropna(how="all")
            if df.empty:
                continue

            df.columns = [text(c) for c in df.columns]
            df["원본파일"] = path.name
            df["원본시트"] = str(sheet)
            frames.append(df)
            print(f"  - {sheet}: {len(df):,}행")
        except Exception as exc:
            print(f"  - {sheet}: 건너뜀 ({exc})")

    if not frames:
        raise RuntimeError(f"{path.name}에서 읽을 수 있는 데이터가 없습니다.")

    return pd.concat(frames, ignore_index=True, sort=False)


def choose_col(columns, candidates):
    columns = list(columns)

    for candidate in candidates:
        if candidate in columns:
            return candidate

    for candidate in candidates:
        for col in columns:
            if candidate in col:
                return col

    return None


def standardize(df):
    candidates = {
        "지역": ["지역", "시도", "지역명"],
        "학교급": ["학교급", "학교급별", "학교구분"],
        "사고자구분": ["사고자구분", "사고자"],
        "학년": ["학년"],
        "성별": ["성별"],
        "사고연월": ["사고연월", "사고발생일", "사고일자", "발생일자", "사고일"],
        "시각": ["시각", "사고시각", "발생시각"],
        "요일": ["요일"],
        "사고시간": ["사고시간", "사고시간대"],
        "사고장소": ["사고장소", "발생장소", "장소"],
        "사고부위": ["사고부위", "부위"],
        "사고형태": ["사고형태", "사고유형", "형태"],
        "사고당시활동": ["사고당시활동", "활동"],
    }

    out = pd.DataFrame(index=df.index)

    for new_col, names in candidates.items():
        old_col = choose_col(df.columns, names)
        out[new_col] = df[old_col] if old_col else ""

    out["원본파일"] = df.get("원본파일", "")
    out["원본시트"] = df.get("원본시트", "")

    for col in out.columns:
        out[col] = out[col].map(text)

    date_text = (
        out["사고연월"]
        .str.replace(".", "-", regex=False)
        .str.replace("/", "-", regex=False)
    )

    out["사고일자"] = pd.to_datetime(date_text, errors="coerce")
    out["연도"] = out["사고일자"].dt.year
    out["월"] = out["사고일자"].dt.month

    missing = out["사고일자"].isna()
    extracted = out.loc[missing, "사고연월"].str.extract(
        r"(?P<year>20\d{2})\D*(?P<month>\d{1,2})?"
    )

    out.loc[missing, "연도"] = pd.to_numeric(extracted["year"], errors="coerce")
    out.loc[missing, "월"] = pd.to_numeric(extracted["month"], errors="coerce")

    return out


def filter_food_poisoning(df):
    mask = pd.Series(False, index=df.index)

    for col in ["사고형태", "사고부위", "사고당시활동", "사고장소"]:
        mask |= df[col].str.contains("식중독", na=False)

    return df[mask].copy()


def reconstruct_events(food):
    work = food.copy()

    date_key = work["사고일자"].dt.strftime("%Y-%m-%d")
    fallback = (
        work["연도"].fillna("").astype(str).str.replace(".0", "", regex=False)
        + "-"
        + work["월"].fillna("").astype(str).str.replace(".0", "", regex=False)
    )
    work["날짜키"] = date_key.fillna(fallback)

    work["시간키"] = work["시각"]
    empty_time = work["시간키"].eq("")
    work.loc[empty_time, "시간키"] = work.loc[empty_time, "사고시간"]

    group_cols = [
        "지역",
        "학교급",
        "날짜키",
        "시간키",
        "사고장소",
        "사고당시활동",
    ]

    events = (
        work.groupby(group_cols, dropna=False)
        .agg(
            피해기록수=("사고형태", "size"),
            연도=("연도", "first"),
            월=("월", "first"),
            남학생수=("성별", lambda s: (s == "남").sum()),
            여학생수=("성별", lambda s: (s == "여").sum()),
            원본시트수=("원본시트", "nunique"),
        )
        .reset_index()
    )

    events.insert(
        0,
        "추정사건ID",
        [f"E{i:04d}" for i in range(1, len(events) + 1)],
    )

    events["대규모집단발생"] = events["피해기록수"] >= 10

    return events.sort_values(
        ["피해기록수", "날짜키"],
        ascending=[False, True],
    )


def save_excel(food, events):
    OUT_DIR.mkdir(exist_ok=True)
    CHART_DIR.mkdir(exist_ok=True)

    food.to_excel(
        OUT_DIR / "01_식중독_학생기록.xlsx",
        index=False,
    )

    events.to_excel(
        OUT_DIR / "02_추정_집단발생사건.xlsx",
        index=False,
    )

    year_summary = pd.concat(
        [
            food.groupby("연도").size().rename("학생별기록수"),
            events.groupby("연도").size().rename("추정사건수"),
            events.groupby("연도")["피해기록수"].sum().rename("피해기록합계"),
        ],
        axis=1,
    ).fillna(0).reset_index()

    school_summary = (
        events.groupby("학교급")
        .agg(
            추정사건수=("추정사건ID", "count"),
            피해기록합계=("피해기록수", "sum"),
            사건당평균피해=("피해기록수", "mean"),
            최대피해규모=("피해기록수", "max"),
        )
        .sort_values("피해기록합계", ascending=False)
        .reset_index()
    )

    month_summary = (
        events.groupby("월")
        .agg(
            추정사건수=("추정사건ID", "count"),
            피해기록합계=("피해기록수", "sum"),
            사건당평균피해=("피해기록수", "mean"),
        )
        .reindex(range(1, 13), fill_value=0)
        .reset_index()
    )

    region_summary = (
        events.groupby("지역")
        .agg(
            추정사건수=("추정사건ID", "count"),
            피해기록합계=("피해기록수", "sum"),
            사건당평균피해=("피해기록수", "mean"),
            최대피해규모=("피해기록수", "max"),
        )
        .sort_values("피해기록합계", ascending=False)
        .reset_index()
    )

    top_events = events.nlargest(30, "피해기록수").copy()

    with pd.ExcelWriter(
        OUT_DIR / "03_분석요약.xlsx",
        engine="openpyxl",
    ) as writer:
        year_summary.to_excel(writer, sheet_name="연도별", index=False)
        school_summary.to_excel(writer, sheet_name="학교급별", index=False)
        month_summary.to_excel(writer, sheet_name="월별", index=False)
        region_summary.to_excel(writer, sheet_name="지역별", index=False)
        top_events.to_excel(writer, sheet_name="대형사건_TOP30", index=False)

    return {
        "year": year_summary,
        "school": school_summary,
        "month": month_summary,
        "region": region_summary,
        "top": top_events,
    }


def make_charts(summary):
    set_korean_font()

    year = summary["year"]
    plt.figure(figsize=(8, 5))
    plt.plot(
        year["연도"],
        year["학생별기록수"],
        marker="o",
        label="학생별 기록 수",
    )
    plt.plot(
        year["연도"],
        year["추정사건수"],
        marker="o",
        label="추정 사건 수",
    )
    plt.title("연도별 식중독 기록과 추정 집단발생 사건")
    plt.xlabel("연도")
    plt.ylabel("건수")
    plt.legend()
    plt.grid(axis="y", alpha=0.25)
    plt.tight_layout()
    plt.savefig(CHART_DIR / "01_연도별_추이.png", dpi=180)
    plt.close()

    region = summary["region"].head(15).sort_values("피해기록합계")
    plt.figure(figsize=(9, 6))
    plt.barh(region["지역"], region["피해기록합계"])
    plt.title("지역별 식중독 피해 기록 합계")
    plt.xlabel("학생별 사고·청구 기록 수")
    plt.tight_layout()
    plt.savefig(CHART_DIR / "02_지역별_피해규모.png", dpi=180)
    plt.close()

    region = summary["region"].head(15).sort_values("추정사건수")
    plt.figure(figsize=(9, 6))
    plt.barh(region["지역"], region["추정사건수"])
    plt.title("지역별 추정 집단발생 사건 수")
    plt.xlabel("추정 사건 수")
    plt.tight_layout()
    plt.savefig(CHART_DIR / "03_지역별_추정사건수.png", dpi=180)
    plt.close()

    month = summary["month"]
    plt.figure(figsize=(8, 5))
    plt.bar(month["월"].astype(int), month["피해기록합계"])
    plt.title("월별 식중독 피해 기록")
    plt.xlabel("월")
    plt.ylabel("학생별 기록 수")
    plt.xticks(range(1, 13))
    plt.tight_layout()
    plt.savefig(CHART_DIR / "04_월별_피해규모.png", dpi=180)
    plt.close()

    school = summary["school"].sort_values("피해기록합계")
    plt.figure(figsize=(8, 5))
    plt.barh(school["학교급"], school["피해기록합계"])
    plt.title("학교급별 식중독 피해 기록")
    plt.xlabel("학생별 기록 수")
    plt.tight_layout()
    plt.savefig(CHART_DIR / "05_학교급별_피해규모.png", dpi=180)
    plt.close()


def main():
    if not DATA_FILE.exists():
        print("오류: run.py와 같은 폴더에서 1.xlsx를 찾지 못했습니다.")
        sys.exit(1)

    if not COMP_FILE.exists():
        print("경고: 2.xlsx를 찾지 못했습니다.")
        print("현재 분석은 1.xlsx만으로도 진행할 수 있습니다.")

    OUT_DIR.mkdir(exist_ok=True)
    CHART_DIR.mkdir(exist_ok=True)

    raw = load_workbook(DATA_FILE)
    standardized = standardize(raw)
    food = filter_food_poisoning(standardized)

    if food.empty:
        print("\n식중독 기록을 찾지 못했습니다.")
        print("원본 엑셀의 컬럼명이나 값 형식을 확인해야 합니다.")
        sys.exit(2)

    events = reconstruct_events(food)
    summary = save_excel(food, events)
    make_charts(summary)

    print("\n========================")
    print("분석 완료")
    print("========================")
    print(f"식중독 학생별 기록: {len(food):,}건")
    print(f"추정 집단발생 사건: {len(events):,}건")
    print(
        f"10명 이상 대규모 사건: "
        f"{(events['피해기록수'] >= 10).sum():,}건"
    )

    print("\n[연도별 요약]")
    print(summary["year"].to_string(index=False))

    print("\n[피해 규모 상위 추정 사건 10개]")
    cols = [
        "추정사건ID",
        "지역",
        "학교급",
        "날짜키",
        "시간키",
        "사고장소",
        "피해기록수",
    ]
    print(summary["top"][cols].head(10).to_string(index=False))

    print("\n결과 폴더:")
    print(OUT_DIR)


if __name__ == "__main__":
    main()
