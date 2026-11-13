# 상담원장 정리 스크립트
#
# 결정 기록: my-wiki/decisions/d-006-상담원장-정리-다섯-결정.md (d-006)
#   1. 처리상태 «종결»→«완료», «대기»→«보류» 로 합친다 (CS기획팀 답)
#   2. 처리시간 음수는 지우지 않는다. 값을 두고 맨 끝 _표시 칸에 이유를 적는다
#   3. 만족도 6~10점은 지우지 않고 _표시 칸에 이유를 적는다. 평균 만족도에서는 표시한 줄을 뺀다
#   4. 칸 이름은 원본 그대로 둔다
#   5. 고객번호 앞뒤 공백을 없앤다
#
# 그 밖의 처리는 원장 시트 «작성기준»과 대화에서 정한 것 그대로다.
#
# 원본은 읽기만 한다. 쓰는 파일은 정리본/ 아래 셋뿐이다:
#   정리본/상담원장_정리.csv · 정리본/지운줄.csv · 정리본/변경요약.md
#
# 돌리기:  python 정리/정리_상담원장.py

import sys
from pathlib import Path

import numpy as np
import pandas as pd

sys.stdout.reconfigure(encoding="utf-8")

ROOT = Path(__file__).resolve().parent.parent          # cleaner/
SRC = ROOT.parent / "messy" / "상담원장_2025.xlsx"
OUT_DIR = ROOT / "정리본"
OUT_CLEAN = OUT_DIR / "상담원장_정리.csv"
OUT_DROP = OUT_DIR / "지운줄.csv"
OUT_SUMMARY = OUT_DIR / "변경요약.md"

FLAG = "_표시"
FLAG_NEG = "처리시간 음수 — 틀린 값이지만 진짜 값을 모름, 값은 그대로 둠 (d-006 #2)"
FLAG_SAT = "만족도 6~10 — NPS를 잘못 넣은 것으로 보이나 확인 못함, 평균 만족도에서 뺌 (d-006 #3)"

CHANNEL = {"유선": "전화", "전화상담": "전화", "TEL": "전화",
           "chatbot": "챗봇", "Chatbot": "챗봇", "챗 봇": "챗봇",
           "app": "앱", "APP": "앱", "모바일앱": "앱",
           "e-mail": "이메일", "E-mail": "이메일", "메일": "이메일"}
REGION = {"경기도": "경기", "서울시": "서울", "서울특별시": "서울"}
STATUS = {"처리완료": "완료", "처리중": "진행중",
          "종결": "완료",          # d-006 #1
          "대기": "보류"}          # d-006 #1
REASK = {"y": "Y", "예": "Y", "TRUE": "Y", "n": "N", "아니오": "N", "FALSE": "N"}


def main():
    raw = pd.read_excel(SRC, sheet_name="원장", dtype=object)
    df = raw.copy()
    changes = []   # (설명, 바뀐 줄의 원본 index)
    dropped = []   # (원본 index, 지운 이유)

    def changed(label, mask):
        changes.append((label, df.index[mask]))

    # 테스트 접수 — 작성기준: AG999 는 시스템 점검용
    t = df["상담원ID"].eq("AG999")
    dropped.append((df.index[t], "테스트 접수(상담원ID AG999)"))
    df = df[~t].copy()

    # 글자 칸 앞뒤 공백 (고객번호 포함 — d-006 #5)
    for c in df.columns:
        m = df[c].map(lambda v: isinstance(v, str) and v != v.strip())
        if m.any():
            label = f"{c}: 앞뒤 공백 지움"
            if c == "고객번호":
                label += " (d-006 #5)"
            changed(label, m)
        df[c] = df[c].map(lambda v: v.strip() if isinstance(v, str) else v)
        df[c] = df[c].map(lambda v: np.nan if isinstance(v, str) and v == "" else v)

    # 접수일 형식
    m = df["접수일"].str.contains(r"[./]")
    changed("접수일: 2025/01/01 · 2025.01.01 → 2025-01-01", m)
    df["접수일"] = pd.to_datetime(df["접수일"].str.replace(r"[./]", "-", regex=True), format="%Y-%m-%d")

    # 표기 통일
    for col, mp, label in [("지역", REGION, "지역: 경기도→경기, 서울시·서울특별시→서울"),
                           ("접수채널", CHANNEL, "접수채널: 전화·챗봇·앱·이메일 넷으로"),
                           ("재문의", REASK, "재문의: y·예·TRUE→Y, n·아니오·FALSE→N")]:
        changed(label, df[col].isin(mp.keys()))
        df[col] = df[col].replace(mp)
    for k, v in STATUS.items():
        note = " (d-006 #1)" if k in ("종결", "대기") else ""
        changed(f"처리상태: {k}→{v}{note}", df["처리상태"].eq(k))
    df["처리상태"] = df["처리상태"].replace(STATUS)
    m = df["상담원팀"].str.contains(" ")
    changed("상담원팀: 사이 공백 지움 (요금 팀→요금팀)", m)
    df["상담원팀"] = df["상담원팀"].str.replace(" ", "", regex=False)
    m = df["월정액"].map(lambda v: isinstance(v, str))
    changed("월정액: 89,000원 → 89000", m)
    df["월정액"] = df["월정액"].map(lambda v: int(str(v).replace(",", "").replace("원", "")))

    # 상담시간 — 작성기준: 2025-07-01 전화 시스템 교체 뒤로 전화 채널은 초 단위
    st = df["상담시간"].astype(float)
    m = st == 9999
    changed("상담시간: 9999 → 빈칸", m)
    st[m] = np.nan
    sec = df["접수채널"].eq("전화") & (df["접수일"] >= "2025-07-01") & st.notna()
    changed("상담시간: 2025-07-01 이후 전화 — 초 → 분 (60으로 나누고 소수 한 자리)", sec)
    st[sec] = (st[sec] / 60).round(1)
    df["상담시간"] = st

    # 표시 칸 (d-006 #2 · #3)
    df[FLAG] = ""

    # 처리시간 — 음수는 값을 두고 표시 (d-006 #2)
    pt = df["처리시간"].astype(float)
    m = pt < 0
    df.loc[m, FLAG] = FLAG_NEG
    changes.append(("_표시: 처리시간 음수 (값은 그대로, d-006 #2)", df.index[m]))
    df["처리시간"] = pt

    # 만족도 — 작성기준: 옛 시스템 이관분 미응답이 0 · «-» · «미응답»
    raw_sat = df["만족도"]
    m0 = raw_sat.astype(str).isin(["0", "-", "미응답"])
    changed("만족도: 0 · «-» · «미응답» → 빈칸 (작성기준: 미응답)", m0)
    mj = raw_sat.map(lambda v: isinstance(v, str) and v.endswith("점"))
    changed("만족도: «5점» → 5", mj)

    def sat(v):
        if isinstance(v, str):
            v = v[:-1] if v.endswith("점") else v
            if not v.isdigit():
                return np.nan
        if pd.isna(v):
            return np.nan
        v = int(v)
        return np.nan if v == 0 else v          # 6~10 은 값을 둔다 (d-006 #3)
    df["만족도"] = raw_sat.map(sat).astype("Int64")
    m = df["만족도"].between(6, 10).fillna(False).astype(bool)
    df.loc[m, FLAG] = np.where(df.loc[m, FLAG] == "", FLAG_SAT, df.loc[m, FLAG] + "; " + FLAG_SAT)
    changes.append(("_표시: 만족도 6~10 (값은 그대로, d-006 #3)", df.index[m]))
    df["NPS"] = df["NPS"].astype("Int64")

    # 행 전체 중복 — 작성기준: 12월 초 접수분 일부가 두 번 추출됐을 수 있음
    d = df.duplicated()
    dropped.append((df.index[d], "행 전체 중복(날짜 형식·표기를 맞춘 뒤 앞 줄과 같음)"))
    df = df[~d]

    # 같은 접수번호 — 작성기준: 최종수정일이 늦은 줄이 최신
    df["최종수정일"] = pd.to_datetime(df["최종수정일"])
    df = df.sort_values(["접수번호", "최종수정일"], kind="stable")
    ties = int(df.duplicated(["접수번호", "최종수정일"], keep=False).sum())
    d = df.duplicated("접수번호", keep="last")
    dropped.append((df.index[d], "같은 접수번호의 옛 상태(최종수정일이 더 늦은 줄이 있음)"))
    df = df[~d]

    df["완료여부"] = np.where(df["처리상태"].eq("완료"), "Y", "N")
    df = df[[c for c in df.columns if c != FLAG] + [FLAG]]      # _표시 는 맨 끝
    df = df.sort_values(["접수일", "접수시각", "접수번호"], kind="stable")
    for c in ["접수일", "최종수정일"]:
        df[c] = df[c].dt.strftime("%Y-%m-%d")

    OUT_DIR.mkdir(exist_ok=True)
    df.to_csv(OUT_CLEAN, index=False, encoding="utf-8-sig", lineterminator="\n")
    parts = []
    for idx, why in dropped:
        x = raw.loc[idx].copy()
        x.insert(0, "원본행번호", idx + 2)
        x["지운이유"] = why
        parts.append(x)
    drop_df = pd.concat(parts).sort_values("원본행번호")
    drop_df.to_csv(OUT_DROP, index=False, encoding="utf-8-sig", lineterminator="\n")

    # 다섯 숫자와 고객 수
    kept = set(df.index)
    sat_ok = df.loc[~df[FLAG].str.contains(FLAG_SAT, regex=False), "만족도"].astype(float)   # 만족도 6~10 으로 표시한 줄은 평균에서 뺀다 (d-006 #3)
    r = df.loc[df["접수일"] >= "2025-03-01", "재문의"]
    metrics = [
        ("상담 건수", f"{len(df):,}건", "접수번호 하나당 한 건"),
        ("평균 상담시간", f"{df['상담시간'].mean():.2f}분", f"값 있는 {df['상담시간'].notna().sum():,}건"),
        ("재문의율", f"{(r == 'Y').mean() * 100:.2f}%", f"{(r == 'Y').sum():,} / {r.notna().sum():,}, 3월 이후 접수분"),
        ("평균 만족도", f"{sat_ok.mean():.2f}", f"1~5점 응답 {sat_ok.notna().sum():,}건, 만족도 6~10 표시 줄 제외"),
        ("완료율", f"{df['완료여부'].eq('Y').mean() * 100:.2f}%", f"{df['완료여부'].eq('Y').sum():,} / {len(df):,}"),
        ("고객 수", f"{df['고객번호'].nunique():,}명", "고객번호 고유 수"),
    ]

    L = ["# 상담원장 변경 요약", "",
         "결정 기록: d-006 (my-wiki/decisions/d-006-상담원장-정리-다섯-결정.md)", "",
         f"- 원본: `../messy/상담원장_2025.xlsx` «원장» 시트 {len(raw):,}줄 (머리글 제외)",
         f"- 정리본: `정리본/상담원장_정리.csv` {len(df):,}줄 · 칸 {len(df.columns)}개 (원본 {len(raw.columns)}칸 + 완료여부 + _표시)",
         f"- 지운 줄: `정리본/지운줄.csv` {len(drop_df):,}줄", "",
         "## 줄 수", "", "| 단계 | 줄 수 |", "|---|---:|", f"| 원본 | {len(raw):,} |"]
    for idx, why in dropped:
        L.append(f"| − {why} | −{len(idx):,} |")
    L += [f"| **정리본** | **{len(df):,}** |", "",
          f"확인: {len(raw):,} − {len(drop_df):,} = {len(raw) - len(drop_df):,} "
          f"→ 정리본 {len(df):,}줄과 {'같음' if len(raw) - len(drop_df) == len(df) else '다름'}", "",
          f"같은 접수번호·같은 최종수정일인 줄: {ties}", "",
          "## 바꾼 것", "",
          "«바꾼 줄»은 테스트 줄을 뺀 뒤, 중복을 지우기 전에 센 수다. «정리본에 남은 줄»은 그중 정리본에 남은 줄 수다.", "",
          "| 무엇을 | 바꾼 줄 | 정리본에 남은 줄 |", "|---|---:|---:|"]
    for label, idx in changes:
        L.append(f"| {label} | {len(idx):,} | {sum(i in kept for i in idx):,} |")
    L += ["", "## 다섯 숫자와 고객 수", "", "| 지표 | 값 | 기준 |", "|---|---:|---|"]
    L += [f"| {a} | {b} | {c} |" for a, b, c in metrics]
    OUT_SUMMARY.write_text("\n".join(L) + "\n", encoding="utf-8")

    print("\n".join(L))


if __name__ == "__main__":
    main()
