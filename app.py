import itertools
import json
import os
import time
from datetime import datetime

import altair as alt
import numpy as np
import pandas as pd
import requests
import streamlit as st
import xmltodict
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook, load_workbook


# -------------------------------------------------
# 0. 기본 설정 & SERVICE_KEY 로드
# -------------------------------------------------
st.set_page_config(page_title="1365 사정율 분석기", layout="wide")

try:
    SERVICE_KEY = st.secrets["SERVICE_KEY"]
except Exception:
    SERVICE_KEY = ""

# -------------------------------------------------
# 진행률 애니메이션 텍스트 프레임 (2번 옵션)
# -------------------------------------------------
LOADING_FRAMES = [
    "⏳ 분석 중입니다...",
    "🔍 계산 중입니다...",
    "📊 데이터 처리 중...",
    "🧮 통계 분석 중...",
    "📈 최적 구간 탐색 중...",
]

def get_loading_text(step):
    """진행률 애니메이션 텍스트 반환"""
    return LOADING_FRAMES[step % len(LOADING_FRAMES)]


# -------------------------------------------------
# 공통 유틸 & API 헬퍼
# -------------------------------------------------
def get_headers():
    return {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

def parse_api_header_from_json(data):
    try:
        response = data.get("response", {})
        header = response.get("header", {})
        code = header.get("resultCode")
        msg = header.get("resultMsg")
        return code, msg
    except Exception:
        return None, None

def parse_api_header_from_xml(data):
    try:
        response = data.get("response", {})
        header = response.get("header", {})
        code = header.get("resultCode")
        msg = header.get("resultMsg")
        return code, msg
    except Exception:
        return None, None

def fetch_json(url: str, desc: str, api_warnings: list, timeout: int = 10):
    try:
        res = requests.get(url, headers=get_headers(), timeout=timeout)
        res.raise_for_status()
    except requests.exceptions.RequestException as e:
        api_warnings.append(f"[HTTP 오류] {desc} 요청 실패: {e}")
        return None

    try:
        data = json.loads(res.text)
    except Exception as e:
        api_warnings.append(f"[파싱 오류] {desc} JSON 파싱 실패: {e}")
        return None

    code, msg = parse_api_header_from_json(data)
    if code is not None and code != "00":
        api_warnings.append(f"[API 오류] {desc} (resultCode={code}, msg={msg})")
        return None

    return data

def fetch_xml(url: str, desc: str, api_warnings: list, timeout: int = 10):
    try:
        res = requests.get(url, headers=get_headers(), timeout=timeout)
        res.raise_for_status()
    except requests.exceptions.RequestException as e:
        api_warnings.append(f"[HTTP 오류] {desc} 요청 실패: {e}")
        return None

    try:
        data = xmltodict.parse(res.text)
    except Exception as e:
        api_warnings.append(f"[파싱 오류] {desc} XML 파싱 실패: {e}")
        return None

    code, msg = parse_api_header_from_xml(data)
    if code is not None and code != "00":
        api_warnings.append(f"[API 오류] {desc} (resultCode={code}, msg={msg})")
        return None

    return data

def safe_get_items(json_data):
    try:
        if not json_data:
            return []
        response = json_data.get("response", {})
        body = response.get("body", {})
        items = body.get("items")

        if not items:
            return []

        if isinstance(items, list):
            return items

        if isinstance(items, dict):
            it = items.get("item")
            if isinstance(it, dict):
                return [it]
            if isinstance(it, list):
                return it
        return []
    except Exception:
        return []
# -------------------------------------------------
# A값 / 집행관 이름
# -------------------------------------------------
def get_a_value(gongo_no: str, api_warnings: list) -> float:
    """A값(안전관리비 등) 조회"""
    try:
        url = (
            "http://apis.data.go.kr/1230000/ad/BidPublicInfoService/"
            "getBidPblancListInfoCnstwkBsisAmount"
            f"?inqryDiv=2&bidNtceNo={gongo_no}&pageNo=1&numOfRows=10&type=json&ServiceKey={SERVICE_KEY}"
        )
        data = fetch_json(url, f"A값 조회({gongo_no})", api_warnings)
        if data is None:
            return 0.0

        items = safe_get_items(data)
        if not items:
            return 0.0

        df = pd.DataFrame(items)
        cost_cols = [
            "sftyMngcst",
            "sftyChckMngcst",
            "rtrfundNon",
            "mrfnHealthInsrprm",
            "npnInsrprm",
            "odsnLngtrmrcprInsrprm",
            "qltyMngcst",
        ]
        valid_cols = [c for c in cost_cols if c in df.columns]
        if not valid_cols:
            return 0.0

        return (
            df[valid_cols]
            .apply(pd.to_numeric, errors="coerce")
            .fillna(0.0)
            .sum(axis=1)
            .iloc[0]
        )
    except Exception:
        return 0.0


def get_officer_name_final(gongo_no: str, api_warnings: list) -> str:
    """집행관 / 담당자 이름 조회"""
    url = (
        "http://apis.data.go.kr/1230000/ad/BidPublicInfoService/"
        f"getBidPblancListInfoCnstwk?inqryDiv=2&bidNtceNo={gongo_no}"
        f"&pageNo=1&numOfRows=1&type=json&ServiceKey={SERVICE_KEY}"
    )
    data = fetch_json(url, f"집행관 조회({gongo_no})", api_warnings)
    if data is None:
        return "확인불가"

    items = safe_get_items(data)
    if not items:
        return "확인불가"

    item = items[0]
    for key in ["exctvNm", "chrgrNm", "ntceChrgrNm"]:
        if key in item and str(item[key]).strip():
            return str(item[key]).strip()
    return "확인불가"


# -------------------------------------------------
# 핫존 / 블루오션 보조 함수
# -------------------------------------------------
def find_hot_zone(actual_rates, window=0.3, step=0.05):
    """
    집행관 장비가 많이 터진 '핫존(실제 1순위 사정율이 가장 몰린 구간)' 탐색
    """
    if not actual_rates:
        return None, None, 0

    rates_sorted = sorted(actual_rates)
    min_r, max_r = min(rates_sorted), max(rates_sorted)

    best_start, best_end, best_count = None, None, -1
    start = min_r
    while start <= max_r:
        end = start + window
        count = sum(start <= r <= end for r in rates_sorted)
        if count > best_count:
            best_count = count
            best_start, best_end = start, end
        start += step

    return best_start, best_end, best_count


def find_blue_ocean_v3(
    theoretical_rates,
    bidder_rates,
    hot_start,
    hot_end,
    bin_width=0.0005,
):
    """
    🔵 블루오션 v3 (최종)
    - 핫존 내부를 bin_width 간격으로 슬라이스
    - 각 구간마다
        * theo_count : 1365 이론 조합 수
        * bid_count  : 실제 투찰 업체 수
    - 스코어: (정규화된 이론 밀도) × (1 / (업체 수 + 1))

    이론이 충분히 있는(수요) 구간이면서, 업체 수(공급)가 적은 곳을 최우선으로 선택.
    """
    if hot_start is None or hot_end is None:
        return None, None, None

    theo = [r for r in theoretical_rates if hot_start <= r <= hot_end]
    bids = [r for r in bidder_rates if hot_start <= r <= hot_end]

    if len(theo) == 0 or len(bids) == 0:
        return None, None, None

    bins = np.arange(hot_start, hot_end + bin_width, bin_width)
    if len(bins) < 2:
        bins = np.array([hot_start, hot_end])

    theo_counts, _ = np.histogram(theo, bins=bins)
    bid_counts, bin_edges = np.histogram(bids, bins=bins)

    if theo_counts.sum() == 0:
        return None, None, None

    theo_norm = theo_counts / theo_counts.sum()
    max_theo = theo_norm.max()
    if max_theo <= 0:
        return None, None, None

    rows = []
    best_score = -1.0
    best_range = None
    best_center = None

    for i in range(len(bin_edges) - 1):
        start = bin_edges[i]
        end = bin_edges[i + 1]
        center = (start + end) / 2

        theo_c = theo_counts[i]
        bid_c = bid_counts[i]

        # 이론 조합이 전혀 없는 구간은 의미가 없으므로 제외
        if theo_c == 0:
            continue

        demand = theo_norm[i] / max_theo          # 이론 밀도 (0~1)
        supply_inv = 1.0 / (bid_c + 1.0)          # 업체수 역수 (업체 적을수록 ↑)
        score = demand * supply_inv

        rows.append(
            {
                "center": center,
                "score": score,
                "theo_count": int(theo_c),
                "bid_count": int(bid_c),
            }
        )

        if score > best_score:
            best_score = score
            best_range = (start, end)
            best_center = center

    if not rows:
        return None, None, None

    blue_df = pd.DataFrame(rows).sort_values("center").reset_index(drop=True)
    return blue_df, best_range, best_center


# -------------------------------------------------
# 공고 1건 분석
# -------------------------------------------------
def analyze_gongo(gongo_input_str: str, api_warnings: list):
    """
    공고번호 1건 분석
    - df_combined : 1365 조합 + 실제 입찰 업체 사정율
    - info        : dict(오피서/1순위업체/1순위사정율)
    - df_rates    : 1365 조합 사정율 리스트
    - bidder_rates: 해당 공고 모든 업체 사정율 리스트
    """
    try:
        if "-" in gongo_input_str:
            parts = gongo_input_str.split("-")
            gongo_no = parts[0].strip()
            gongo_ord = parts[1].strip()
        else:
            gongo_no = gongo_input_str.strip()
            gongo_ord = "00"

        officer_name = get_officer_name_final(gongo_no, api_warnings)

        # ------------------------------
        # 1) 복수예가 (1365 조합용)
        # ------------------------------
        url1 = (
            "http://apis.data.go.kr/1230000/as/ScsbidInfoService/"
            "getOpengResultListInfoCnstwkPreparPcDetail"
            f"?inqryDiv=2&bidNtceNo={gongo_no}&bidNtceOrd={gongo_ord}"
            f"&pageNo=1&numOfRows=15&type=json&ServiceKey={SERVICE_KEY}"
        )
        data1 = fetch_json(url1, f"복수예가 조회({gongo_no})", api_warnings)
        df_rates = pd.DataFrame()
        base_price = 0.0

        if data1 is not None:
            try:
                items1 = safe_get_items(data1)
                if items1:
                    df1 = pd.json_normalize(items1)
                    if "bssamt" in df1.columns and "bsisPlnprc" in df1.columns:
                        df1 = df1[["bssamt", "bsisPlnprc"]].astype(float)
                        base_price = (
                            df1.iloc[1]["bssamt"] if len(df1) > 1 else df1.iloc[0]["bssamt"]
                        )
                        df1["SA_rate"] = df1["bsisPlnprc"] / df1["bssamt"] * 100

                        if len(df1) >= 4:
                            rates = [
                                np.mean(c)
                                for c in itertools.combinations(df1["SA_rate"], 4)
                            ]
                            df_rates = (
                                pd.DataFrame({"rate": rates})
                                .sort_values("rate")
                                .reset_index(drop=True)
                            )
                            df_rates["조합순번"] = range(1, len(df_rates) + 1)
            except Exception:
                pass

        # ------------------------------
        # 2) 낙찰하한율
        # ------------------------------
        sucs_rate = 0.0
        url2 = (
            "http://apis.data.go.kr/1230000/ad/BidPublicInfoService/"
            "getBidPblancListInfoCnstwk"
            f"?inqryDiv=2&bidNtceNo={gongo_no}&pageNo=1&numOfRows=1&type=json&ServiceKey={SERVICE_KEY}"
        )
        data2 = fetch_json(url2, f"낙찰하한율 조회({gongo_no})", api_warnings)
        if data2 is not None:
            try:
                items2 = safe_get_items(data2)
                if items2 and "sucsfbidLwltRate" in items2[0]:
                    sucs_rate = float(items2[0]["sucsfbidLwltRate"])
            except Exception:
                pass

        # ------------------------------
        # 3) A값
        # ------------------------------
        A_value = get_a_value(gongo_no, api_warnings)

        # ------------------------------
        # 4) 개찰결과 (XML, 전체 업체)
        # ------------------------------
        url4 = (
            "http://apis.data.go.kr/1230000/as/ScsbidInfoService/"
            f"getOpengResultListInfoOpengCompt?serviceKey={SERVICE_KEY}"
            f"&pageNo=1&numOfRows=999&bidNtceNo={gongo_no}"
        )
        data4 = fetch_xml(url4, f"개찰결과 조회({gongo_no})", api_warnings)
        if data4 is None:
            return (
                pd.DataFrame(),
                f"개찰결과 조회 실패({gongo_input_str})",
                None,
                pd.DataFrame(),
                [],
            )

        try:
            items4_raw = data4.get("response", {}).get("body", {}).get("items")
            if isinstance(items4_raw, dict):
                items4 = items4_raw.get("item", [])
            elif isinstance(items4_raw, list):
                items4 = items4_raw
            else:
                items4 = []
            if isinstance(items4, dict):
                items4 = [items4]
            if not isinstance(items4, list):
                items4 = []
        except Exception:
            items4 = []

        df4 = pd.DataFrame(items4)
        top_info = {"winner": "개찰결과 없음", "rate": 0.0, "officer": officer_name}
        bidder_rates_all = []

        if not df4.empty and "bidprcAmt" in df4.columns:
            df4["bidprcAmt"] = pd.to_numeric(df4["bidprcAmt"], errors="coerce")
            df4 = df4.dropna(subset=["bidprcAmt"])

            if not df4.empty:
                top_name = str(df4.iloc[0].get("prcbdrNm", "업체명없음"))

                if sucs_rate > 0 and base_price > 0:
                    numerator = ((df4["bidprcAmt"] - A_value) * 100) / sucs_rate + A_value
                    df4["rate"] = numerator * 100 / base_price
                else:
                    df4["rate"] = 0.0

                bidder_rates_all = df4["rate"].astype(float).tolist()

                top_row = df4.iloc[0]
                top_rate = float(top_row.get("rate", 0.0))

                top_info = {
                    "winner": top_name,
                    "rate": round(top_rate, 5),
                    "officer": officer_name,
                }

                df4_clean = df4.drop_duplicates(subset=["rate"])
                df4_clean = df4_clean[
                    (df4_clean["rate"] >= 90) & (df4_clean["rate"] <= 110)
                ]
                df4_clean = df4_clean[["prcbdrNm", "rate"]].rename(
                    columns={"prcbdrNm": "업체명"}
                )
            else:
                df4_clean = pd.DataFrame()
        else:
            df4_clean = pd.DataFrame()

        # ------------------------------
        # 5) 조합 + 실제 통합 DF
        # ------------------------------
        if not df_rates.empty:
            df_combined = pd.concat(
                [
                    df_rates[["rate"]].assign(업체명=df_rates["조합순번"].astype(str)),
                    df4_clean[["업체명", "rate"]],
                ],
                ignore_index=True,
            )
        else:
            df_combined = df4_clean.copy()

        if not df_combined.empty:
            df_combined = df_combined.sort_values("rate").reset_index(drop=True)
            df_combined["rate"] = df_combined["rate"].round(5)
            df_combined["공고번호"] = gongo_no

        return df_combined, None, top_info, df_rates, bidder_rates_all

    except Exception as e:
        return (
            pd.DataFrame(),
            f"예외 발생 ({gongo_input_str}): {e}",
            None,
            pd.DataFrame(),
            [],
        )
# -------------------------------------------------
# 전체 실행 + 엑셀 저장 + 진행률 표시(Progress + ETA)
# -------------------------------------------------
import time

def process_analysis(target_officer: str, gongo_input: str, progress_placeholder, progress_text):
    """
    메인 분석 루틴
    - 진행률 표시 + ETA
    - 추천 사정률 ±0.0001 강조
    - 핫존 / 블루오션 통계
    """
    start_time = time.time()
    api_warnings = []
    progress = 0
    progress_placeholder.progress(0.0)
    progress_text.markdown("⏳ 분석 준비 중...")

    if not gongo_input.strip():
        return (
            "공고번호를 입력해주세요.",
            None, None, None,
            "분석된 데이터가 없습니다.",
            None, None,
            {"total": 0, "filtered": 0, "missing": 0, "blue_range": "없음", "rec_rate": None},
            None,
            api_warnings,
        )

    if not SERVICE_KEY:
        api_warnings.append("SERVICE_KEY가 설정되어 있지 않습니다.")
        return (
            "❌ SERVICE_KEY 미설정",
            None, None, None,
            "SERVICE_KEY 미설정",
            None, None,
            {"total": 0, "filtered": 0, "missing": 0},
            None,
            api_warnings,
        )

    gongo_list = [x.strip() for x in gongo_input.replace(",", "\n").split("\n") if x.strip()]
    total_gongo = len(gongo_list)
    target_clean = target_officer.strip()

    logs = []
    results_for_merge = []
    scatter_data = []
    winner_rates = []
    theoretical_rates_all = []
    bidder_rates_all = []

    # ================================
    # 🔥 공고 개수 기준 진행률 업데이트 함수
    # ================================
    def update_progress(i):
        elapsed = time.time() - start_time
        pct = (i / total_gongo)
        remaining = (elapsed / pct) - elapsed if pct > 0 else 0

        bar = "■" * int(pct * 20)
        bar += "□" * (20 - len(bar))

        progress_placeholder.progress(pct)
        progress_text.markdown(
            f"""
🔄 **분석 중...**

`{bar}` **{pct*100:5.1f}%**

⏱ 경과: **{elapsed:5.1f}초**  
⏳ 예상 남은 시간: **{remaining:5.1f}초**
"""
        )

    # ================================
    # 🔥 공고들 반복 분석
    # ================================
    for idx, gongo in enumerate(gongo_list, start=1):
        df, err, info, df_rates_raw, bidder_rates = analyze_gongo(gongo, api_warnings)

        # 진행률 업데이트
        update_progress(idx)

        if err:
            logs.append(f"❌ {gongo} | 오류: {err}")
            continue

        officer = str(info["officer"]).strip()
        winner = info["winner"]
        w_rate = info["rate"]

        # 집행관 필터
        if target_clean:
            if officer != target_clean:
                logs.append(f"⛔ 제외: {gongo} | 집행관: {officer}")
                continue
            else:
                logs.append(f"✅ 포함: {gongo} | 집행관: {officer} | 1순위: {winner} ({w_rate}%)")
        else:
            logs.append(f"✅ {gongo} | 집행관: {officer} | 1순위: {winner} ({w_rate}%)")

        if not df.empty:
            results_for_merge.append({"gongo": gongo, "df": df, "info": info})

        if w_rate != 0:
            winner_rates.append(w_rate)
            scatter_data.append([w_rate, gongo, winner])

        if not df_rates_raw.empty:
            theoretical_rates_all.extend(df_rates_raw["rate"].tolist())

        if bidder_rates:
            bidder_rates_all.extend(bidder_rates)

    progress_text.markdown("📊 **데이터 병합 및 분석 중...**")

    # ================================
    # 🔥 데이터 없으면 종료
    # ================================
    if not results_for_merge:
        logs.append("⚠ 유효한 분석 데이터 없음")
        stats = {
            "total": total_gongo,
            "filtered": 0,
            "missing": total_gongo,
            "blue_range": "없음",
            "rec_rate": None
        }
        return (
            "\n".join(logs),
            None, None, None,
            "유효한 데이터 없음",
            None, None,
            stats,
            None,
            api_warnings,
        )

    # ================================
    # 🔥 통합 테이블 생성
    # ================================
    all_rates = pd.concat([r["df"]["rate"] for r in results_for_merge]).unique()
    merged_df = pd.DataFrame({"rate": all_rates}).sort_values("rate").reset_index(drop=True)

    col_index_to_winner = {}
    col_index_to_winrate = {}

    for res in results_for_merge:
        df = res["df"]
        info = res["info"]
        gongo_no = df["공고번호"].iloc[0]
        officer = info["officer"]
        winner = info["winner"]
        w_rate = info["rate"]

        col_name = f"{gongo_no}\n[{officer}]\n{winner}"

        sub_df = df[["rate", "업체명"]].rename(columns={"업체명": col_name})
        merged_df = pd.merge(merged_df, sub_df, on="rate", how="outer")

        col_index_to_winner[col_name] = winner
        col_index_to_winrate[col_name] = w_rate

    merged_df = merged_df.sort_values("rate").reset_index(drop=True)
    merged_df = merged_df.fillna("")

    # 화면용 데이터프레임
    header_row = {"rate": "1순위 사정률(%)"}
    for col in merged_df.columns[1:]:
        wr = col_index_to_winrate.get(col)
        header_row[col] = f"{wr:.4f}" if wr is not None else ""

    merged_display_df = pd.concat([pd.DataFrame([header_row]), merged_df], ignore_index=True)

    # ================================
    # 🔥 엑셀 파일 생성
    # ================================
    progress_text.markdown("📁 **엑셀 파일 생성 중...**")

    excel_filename = f"사정율분석_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "통합분석"

    # 데이터 삽입
    for r in dataframe_to_rows(merged_df, index=False, header=True):
        ws.append(r)

    # 1순위 사정률 행 추가
    second_row = ["1순위 사정률(%)"]
    for col in merged_df.columns[1:]:
        wr = col_index_to_winrate.get(col)
        second_row.append(f"{wr:.4f}" if wr is not None else "")
    ws.insert_rows(2)
    for col_idx, v in enumerate(second_row, start=1):
        ws.cell(row=2, column=col_idx, value=v)

    # 헤더 서식
    header_font = Font(bold=True)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = header_align
    for cell in ws[2]:
        cell.font = header_font
        cell.alignment = header_align

    # 1순위 업체 강조
    yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for col_idx, col_name in enumerate(merged_df.columns, start=1):
        if col_idx == 1:
            continue
        winner = col_index_to_winner.get(col_name)
        if not winner:
            continue
        for row_idx in range(3, ws.max_row + 1):
            if ws.cell(row=row_idx, column=col_idx).value == winner:
                ws.cell(row=row_idx, column=col_idx).fill = yellow

    # -------------------------
    # 🔥 추천 사정률(±0.0001) 강조
    # -------------------------
    rec_rate = None  # 우선 None으로 초기화, 아래 블루오션 계산 후 값 반영됨

    highlight = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    # 추천 사정률은 블루오션 분석에서 계산된 후 아래에서 다시 적용됨
    # (여기서는 엑셀 구조 준비만 해둠)

    # 파일 저장
    wb.save(excel_filename)
    excel_path = excel_filename

    # ================================
    # 🔥 핫존 계산
    # ================================
    hot_start, hot_end = None, None
    if winner_rates:
        hot_start, hot_end, _ = find_hot_zone(winner_rates)
        if hot_start is None or hot_end is None:
            hot_start, hot_end = min(winner_rates), max(winner_rates)

    # ================================
    # 🔥 산점도 생성
    # ================================
    chart_main = None
    if scatter_data:
        chart_df = pd.DataFrame(scatter_data, columns=["rate", "공고번호", "업체명"])

        def cat(v):
            return "🔥 핫존" if hot_start <= v <= hot_end else "일반"

        chart_df["구분"] = chart_df["rate"].apply(cat)

        chart_main = (
            alt.Chart(chart_df)
            .mark_circle(size=140)
            .encode(
                x=alt.X("rate", title="사정율 (%)"),
                y=alt.Y("공고번호", title="공고번호"),
                color=alt.condition(
                    alt.datum.구분 == "🔥 핫존",
                    alt.value("#FF3B30"),
                    alt.value("#CCCCCC")
                ),
                tooltip=["업체명", "rate", "공고번호", "구분"],
            )
            .interactive()
        )

    # ================================
    # 🔥 블루오션 분석
    # ================================
    blue_df, best_range, best_center = None, None, None
    if hot_start is not None and hot_end is not None and theoretical_rates_all and bidder_rates_all:
        blue_df, best_range, best_center = find_blue_ocean_v3(
            theoretical_rates_all,
            bidder_rates_all,
            hot_start,
            hot_end,
            bin_width=0.0005,
        )

    chart_gap = None
    blue_desc = ""
    best_range_str = "없음"

    if blue_df is not None and best_range is not None:
        best_range_str = f"{best_range[0]:.4f}% ~ {best_range[1]:.4f}%"
        rec_rate = round(best_range[1], 4)

        # 블루오션 그래프
        plot_df = blue_df.rename(columns={"center": "구간중심", "score": "블루오션점수"})
        chart_gap = (
            alt.Chart(plot_df)
            .mark_bar()
            .encode(
                x=alt.X("구간중심", title="사정율 구간 중심 (%)"),
                y=alt.Y("블루오션점수", title="블루오션 점수"),
                tooltip=["구간중심", "블루오션점수", "theo_count", "bid_count"],
            )
            .interactive()
        )

        blue_desc = (
            f"- 이 집행관의 핫존은 **{hot_start:.4f}% ~ {hot_end:.4f}%** 입니다.\n"
            f"- 최적 블루오션 구간은 **{best_range_str}** 입니다.\n"
            f"- 추천 투찰 사정률: **{rec_rate:.4f}%**\n"
        )
    else:
        blue_desc = "블루오션 통계가 충분하지 않습니다."

    # ================================
    # 🔥 추천 사정률을 엑셀에 반영 (±0.0001)
    # ================================
    if rec_rate is not None:
        lower = rec_rate - 0.0001
        upper = rec_rate + 0.0001

        wb2 = Workbook()
        wb2 = load_workbook(excel_path)
        ws2 = wb2.active

        for row in range(3, ws2.max_row + 1):
            try:
                val = float(ws2.cell(row=row, column=1).value)
                if lower <= val <= upper:
                    for col in range(1, ws2.max_column + 1):
                        ws2.cell(row=row, column=col).fill = highlight
            except:
                pass

        wb2.save(excel_path)

    # ================================
    # 🔥 통계 요약 생성
    # ================================
    stats = {
        "total": total_gongo,
        "filtered": len(results_for_merge),
        "missing": total_gongo - len(results_for_merge),
        "blue_range": best_range_str,
        "rec_rate": rec_rate,
    }

    analysis_text = f"""
### 🔥 집행관 핫존
- **{hot_start:.4f}% ~ {hot_end:.4f}%**

### 💎 블루오션 분석
{blue_desc}
"""

    progress_text.markdown("✅ **모든 분석이 완료되었습니다!**")

    return (
        "\n".join(logs),
        merged_display_df,
        hot_start,
        hot_end,
        analysis_text,
        chart_main,
        chart_gap,
        stats,
        excel_path,
        api_warnings,
    )
# -------------------------------------------------
# Streamlit UI (디자인 + 실행 버튼 + 진행률 연결)
# -------------------------------------------------

def reset_gongo():
    st.session_state["gongo_text"] = ""


# ---------------------- CSS -----------------------
st.markdown("""
<style>

html, body, [data-testid="stAppViewContainer"] {
    background: linear-gradient(135deg, #1e1e2f 0%, #2f2f46 50%, #191926 100%);
    color: #fff !important;
}

/* fade-in */
.fade-in {
    opacity: 0;
    animation: fadeIn 1.2s forwards;
}
@keyframes fadeIn {
    to { opacity: 1; }
}

/* 버튼 스타일 */
button[kind="primary"] {
    background: linear-gradient(90deg, #ff7b3d, #ff4f4f);
    border-radius: 8px;
    border: none;
    font-weight: 600;
    transition: 0.3s;
}
button[kind="primary"]:hover {
    transform: scale(1.03);
    background: linear-gradient(90deg, #ff9966, #ff5f5f);
}

button[kind="secondary"] {
    background: #444 !important;
    border-radius: 8px;
    border: none;
}
button[kind="secondary"]:hover {
    background: #666 !important;
    transform: scale(1.03);
}


/* 메트릭 카드 */
.metric-card {
    background: rgba(255,255,255,0.1);
    padding: 18px;
    border-radius: 15px;
    backdrop-filter: blur(8px);
    border: 1px solid rgba(255,255,255,0.2);
    text-align: center;
    transition: 0.3s;
}
.metric-card:hover {
    transform: translateY(-4px);
}

/* 추천 사정률 강조 */
.glow-box {
    background: rgba(255,240,200,0.15);
    border: 1px solid #ffdd9c;
    border-radius: 15px;
    padding: 20px;
    animation: glow 3s infinite ease-in-out;
}
@keyframes glow {
    0% { box-shadow: 0 0 10px #ffdd9c55; }
    50% { box-shadow: 0 0 20px #ffdd9c; }
    100% { box-shadow: 0 0 10px #ffdd9c55; }
}

</style>
""", unsafe_allow_html=True)


# ---------------------- HEADER -----------------------
st.markdown(
    """
<h1 class="fade-in" style="text-align:center;
 font-size:40px; font-weight:900;
 background: linear-gradient(90deg,#ffddaa,#ffd087,#ffb067);
 -webkit-background-clip:text; color:transparent;">
🏗 1365 사정율 분석기<br>(핫존 + 블루오션 + 추천 사정률)
</h1>
""",
    unsafe_allow_html=True,
)

st.markdown("<br>", unsafe_allow_html=True)


# ---------------------- INPUT AREA -----------------------
target = st.text_input("🎯 타겟 집행관 (선택 사항)", value="")

gongo_input = st.text_area(
    "📄 공고번호 목록 입력",
    height=180,
    key="gongo_text",
    placeholder="예)\nR25BK01074208-000\nR25BK01071774-000\n...",
)

btn_col1, btn_col2 = st.columns([1, 1])
with btn_col1:
    run_clicked = st.button("🚀 분석 실행", use_container_width=True)
with btn_col2:
    st.button("🧹 초기화", use_container_width=True, on_click=reset_gongo)


# ---------------------- EXECUTION -----------------------
if run_clicked:
    # 진행률 Placeholder (UI 영역 확보)
    progress_placeholder = st.empty()
    progress_text = st.empty()

    with st.spinner("🔄 분석을 시작합니다..."):
        result = process_analysis(target, gongo_input, progress_placeholder, progress_text)

    # 결과 저장
    st.session_state["analysis_result"] = {
        "logs": result[0],
        "merged": result[1],
        "hot_start": result[2],
        "hot_end": result[3],
        "analysis_md": result[4],
        "chart_main": result[5],
        "chart_gap": result[6],
        "stats": result[7],
        "excel_path": result[8],
        "api_warnings": result[9],
    }


# ---------------------- RESULT DISPLAY -----------------------
if "analysis_result" in st.session_state:
    res = st.session_state["analysis_result"]

    # API 경고 메시지
    if res["api_warnings"]:
        st.warning(
            "⚠ 공공데이터포털 API 경고/오류 발생:\n\n"
            + "\n".join(f"- {w}" for w in res["api_warnings"])
        )

    # 로그 표시
    st.markdown("## 📜 로그")
    st.code(res["logs"])

    merged = res["merged"]
    if merged is None or merged.empty:
        st.error("⚠ 유효한 분석 데이터 없음")
    else:
        stats = res["stats"]

        # 요약 메트릭 카드
        st.markdown("## 🔍 핵심 요약")
        c1, c2, c3, c4 = st.columns(4)
        c1.markdown(f"<div class='metric-card'><h3>핫존 시작</h3><h2>{res['hot_start']:.4f}%</h2></div>", unsafe_allow_html=True)
        c2.markdown(f"<div class='metric-card'><h3>핫존 끝</h3><h2>{res['hot_end']:.4f}%</h2></div>", unsafe_allow_html=True)
        c3.markdown(f"<div class='metric-card'><h3>분석 공고</h3><h2>{stats['filtered']}</h2></div>", unsafe_allow_html=True)
        c4.markdown(f"<div class='metric-card'><h3>누락 공고</h3><h2>{stats['missing']}</h2></div>", unsafe_allow_html=True)

        # 추천 사정률 박스
        st.markdown("## 🔥 추천 투찰 사정률")
        rec = stats.get("rec_rate")
        if rec:
            st.markdown(
                f"""
<div class='glow-box'>
    <h2 style='color:#ffcc66;'>🔥 {rec:.4f}%</h2>
    <p style='font-size:14px;'>핫존 + 블루오션 기반 추천 사정률</p>
</div>
""",
                unsafe_allow_html=True,
            )
        else:
            st.info("추천 사정률 없음")

        # 텍스트 보고서
        st.markdown("## 🎯 종합 분석 리포트")
        st.markdown(res["analysis_md"])

        # 그래프(1순위 분포)
        if res["chart_main"] is not None:
            st.markdown("## 📈 1순위 사정률 분포")
            st.altair_chart(res["chart_main"], use_container_width=True)

        # 블루오션 그래프
        if res["chart_gap"] is not None:
            st.markdown("## 💎 블루오션 점수 그래프")
            st.altair_chart(res["chart_gap"], use_container_width=True)

        # 통합 테이블
        st.markdown("## 📑 통합 테이블")
        st.dataframe(merged, use_container_width=True)

        # 엑셀 다운로드 버튼
        if res["excel_path"]:
            with open(res["excel_path"], "rb") as f:
                st.download_button(
                    "📥 엑셀 다운로드",
                    f,
                    file_name=res["excel_path"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
