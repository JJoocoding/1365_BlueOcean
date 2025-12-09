import itertools
import json
import os
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

# -------------------------------------------------
# 0. 기본 설정 & SERVICE_KEY 로드
# -------------------------------------------------
st.set_page_config(page_title="1365 사정율 분석기", layout="wide")

try:
    SERVICE_KEY = st.secrets["SERVICE_KEY"]
except Exception:
    SERVICE_KEY = ""


# -------------------------------------------------
# 공통 유틸
# -------------------------------------------------
def get_headers():
    return {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}


def safe_get_items(json_data):
    """response.body.items.item 에서 item 리스트만 안전하게 추출"""
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
            item_list = items.get("item")
            if not item_list:
                return []
            if isinstance(item_list, dict):
                return [item_list]
            if isinstance(item_list, list):
                return item_list

        return []
    except Exception:
        return []


# -------------------------------------------------
# A값 / 집행관 이름
# -------------------------------------------------
def get_a_value(gongo_no: str) -> float:
    """A값(안전관리비 등) 조회"""
    try:
        url = (
            "http://apis.data.go.kr/1230000/ad/BidPublicInfoService/"
            "getBidPblancListInfoCnstwkBsisAmount"
            f"?inqryDiv=2&bidNtceNo={gongo_no}&pageNo=1&numOfRows=10&type=json&ServiceKey={SERVICE_KEY}"
        )
        res = requests.get(url, headers=get_headers(), timeout=7)
        data = json.loads(res.text)
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


def get_officer_name_final(gongo_no: str) -> str:
    """집행관 / 담당자 이름 조회"""
    url = (
        "http://apis.data.go.kr/1230000/ad/BidPublicInfoService/"
        f"getBidPblancListInfoCnstwk?inqryDiv=2&bidNtceNo={gongo_no}"
        f"&pageNo=1&numOfRows=1&type=json&ServiceKey={SERVICE_KEY}"
    )
    try:
        res = requests.get(url, headers=get_headers(), timeout=7)
        data = json.loads(res.text)
        items = safe_get_items(data)
        if not items:
            return "확인불가"
        item = items[0]
        for key in ["exctvNm", "chrgrNm", "ntceChrgrNm"]:
            if key in item and str(item[key]).strip():
                return str(item[key]).strip()
        return "확인불가"
    except Exception:
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


def find_blue_ocean_v3(theoretical_rates, bidder_rates, hot_start, hot_end, bin_width=0.0005):
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
def analyze_gongo(gongo_input_str: str):
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

        headers = get_headers()
        officer_name = get_officer_name_final(gongo_no)

        # ------------------------------
        # 1) 복수예가 (1365 조합용)
        # ------------------------------
        url1 = (
            "http://apis.data.go.kr/1230000/as/ScsbidInfoService/"
            "getOpengResultListInfoCnstwkPreparPcDetail"
            f"?inqryDiv=2&bidNtceNo={gongo_no}&bidNtceOrd={gongo_ord}"
            f"&pageNo=1&numOfRows=15&type=json&ServiceKey={SERVICE_KEY}"
        )
        res1 = requests.get(url1, headers=headers, timeout=10)

        df_rates = pd.DataFrame()
        base_price = 0.0

        try:
            data1 = json.loads(res1.text)
            items1 = safe_get_items(data1)
            if items1:
                df1 = pd.json_normalize(items1)
                if "bssamt" in df1.columns and "bsisPlnprc" in df1.columns:
                    df1 = df1[["bssamt", "bsisPlnprc"]].astype(float)
                    base_price = df1.iloc[1]["bssamt"] if len(df1) > 1 else df1.iloc[0]["bssamt"]
                    df1["SA_rate"] = df1["bsisPlnprc"] / df1["bssamt"] * 100

                    if len(df1) >= 4:
                        rates = [
                            np.mean(c) for c in itertools.combinations(df1["SA_rate"], 4)
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
        try:
            url2 = (
                "http://apis.data.go.kr/1230000/ad/BidPublicInfoService/"
                "getBidPblancListInfoCnstwk"
                f"?inqryDiv=2&bidNtceNo={gongo_no}&pageNo=1&numOfRows=1&type=json&ServiceKey={SERVICE_KEY}"
            )
            res2 = requests.get(url2, headers=headers, timeout=10)
            data2 = json.loads(res2.text)
            items2 = safe_get_items(data2)
            if items2 and "sucsfbidLwltRate" in items2[0]:
                sucs_rate = float(items2[0]["sucsfbidLwltRate"])
        except Exception:
            pass

        # ------------------------------
        # 3) A값
        # ------------------------------
        A_value = get_a_value(gongo_no)

        # ------------------------------
        # 4) 개찰결과 (XML, 전체 업체)
        # ------------------------------
        url4 = (
            "http://apis.data.go.kr/1230000/as/ScsbidInfoService/"
            f"getOpengResultListInfoOpengCompt?serviceKey={SERVICE_KEY}"
            f"&pageNo=1&numOfRows=999&bidNtceNo={gongo_no}"
        )
        try:
            res4 = requests.get(url4, headers=headers, timeout=10)
        except Exception as e:
            return (
                pd.DataFrame(),
                f"HTTP 오류 ({gongo_input_str}): {e}",
                None,
                pd.DataFrame(),
                [],
            )

        try:
            data4 = xmltodict.parse(res4.text)
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

                # 모든 업체 사정률 (블루오션용)
                bidder_rates_all = df4["rate"].astype(float).tolist()

                top_row = df4.iloc[0]
                top_rate = float(top_row.get("rate", 0.0))

                top_info = {
                    "winner": top_name,
                    "rate": round(top_rate, 5),
                    "officer": officer_name,
                }

                # 통합테이블용
                df4_clean = df4.drop_duplicates(subset=["rate"])
                df4_clean = df4_clean[(df4_clean["rate"] >= 90) & (df4_clean["rate"] <= 110)]
                df4_clean = df4_clean[["prcbdrNm", "rate"]].rename(columns={"prcbdrNm": "업체명"})
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
        return pd.DataFrame(), f"예외 발생 ({gongo_input_str}): {e}", None, pd.DataFrame(), []


# -------------------------------------------------
# 전체 실행 + 엑셀 저장
# -------------------------------------------------
def process_analysis(target_officer: str, gongo_input: str):
    if not gongo_input.strip():
        return (
            "공고번호를 입력해주세요.",
            None,
            None,
            None,
            "분석된 데이터가 없습니다.",
            None,
            None,
            {"total": 0, "filtered": 0, "missing": 0, "blue_range": "없음", "rec_rate": None},
            None,
        )

    if not SERVICE_KEY:
        return (
            "❌ SERVICE_KEY 미설정 (secrets.toml 확인)",
            None,
            None,
            None,
            "SERVICE_KEY 미설정으로 분석 중단",
            None,
            None,
            {"total": 0, "filtered": 0, "missing": 0, "blue_range": "없음", "rec_rate": None},
            None,
        )

    gongo_list = [x.strip() for x in gongo_input.replace(",", "\n").split("\n") if x.strip()]
    target_clean = target_officer.strip()

    logs = []
    results_for_merge = []
    scatter_data = []   # 1순위 산점도
    winner_rates = []   # 핫존용
    theoretical_rates_all = []
    bidder_rates_all = []

    for gongo in gongo_list:
        df, err, info, df_rates_raw, bidder_rates = analyze_gongo(gongo)

        if err:
            logs.append(f"❌ {gongo} | 오류: {err}")
            continue

        officer = str(info["officer"]).strip()
        winner = info["winner"]
        w_rate = info["rate"]

        # 집행관 필터
        if target_clean:
            if officer != target_clean:
                logs.append(f"⛔ [제외] {gongo} | 집행관: {officer}")
                continue
            else:
                logs.append(
                    f"✅ [포함] {gongo} | 집행관: {officer} | 1순위: {winner} ({w_rate}%)"
                )
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

    if not results_for_merge:
        logs.append("⚠ 유효한 분석 데이터가 없습니다.")
        return (
            "\n".join(logs),
            None,
            None,
            None,
            "분석된 데이터가 없습니다.",
            None,
            None,
            {
                "total": len(gongo_list),
                "filtered": 0,
                "missing": len(gongo_list),
                "blue_range": "없음",
                "rec_rate": None,
            },
            None,
        )

    # ---------------------------
    # 통합 테이블 생성
    # ---------------------------
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

    # 화면용: 1행에 1순위 사정률을 한 번 더 보여주는 행 추가
    header_row = {"rate": "1순위 사정률(%)"}
    for col in merged_df.columns[1:]:
        wr = col_index_to_winrate.get(col)
        header_row[col] = f"{wr:.4f}" if wr is not None else ""
    merged_display_df = pd.concat(
        [pd.DataFrame([header_row]), merged_df], ignore_index=True
    )

    # ---------------------------
    # 엑셀 파일 생성
    # ---------------------------
    excel_filename = f"사정율분석_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

    wb = Workbook()
    ws = wb.active
    ws.title = "통합분석"

    # DataFrame → Worksheet
    for r in dataframe_to_rows(merged_df, index=False, header=True):
        ws.append(r)

    # 두 번째 행에 1순위 사정률 추가
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

    # 1순위 업체 하이라이트
    fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for col_idx, col_name in enumerate(merged_df.columns, start=1):
        if col_idx == 1:
            continue
        winner = col_index_to_winner.get(col_name)
        if not winner:
            continue
        for row_idx in range(3, ws.max_row + 1):
            if ws.cell(row=row_idx, column=col_idx).value == winner:
                ws.cell(row=row_idx, column=col_idx).fill = fill

    wb.save(excel_filename)
    excel_path = excel_filename

    # ---------------------------
    # 그래프 + 블루오션 분석
    # ---------------------------
    hot_start, hot_end, _ = find_hot_zone(winner_rates)
    if hot_start is None or hot_end is None:
        hot_start, hot_end = min(winner_rates), max(winner_rates)

    # 메인 산점도 (1순위 분포)
    chart_main = None
    if scatter_data:
        chart_df = pd.DataFrame(scatter_data, columns=["rate", "공고번호", "업체명"])
        min_rate = chart_df["rate"].min()
        max_rate = chart_df["rate"].max()

        def cat(v):
            return "🔥 핫존" if hot_start <= v <= hot_end else "일반"

        chart_df["구분"] = chart_df["rate"].apply(cat)

        base_chart = alt.Chart(chart_df).encode(
            x=alt.X(
                "rate",
                title="사정율 (%)",
                scale=alt.Scale(domain=[min(min_rate, 98) - 0.2, max(max_rate, 102) + 0.2]),
            ),
            y=alt.Y("공고번호", sort=None, title="공고번호"),
            tooltip=["업체명", "rate", "공고번호", "구분"],
        )

        chart_main = (
            base_chart.mark_circle(size=120)
            .encode(
                color=alt.Color(
                    "구분",
                    scale=alt.Scale(domain=["🔥 핫존", "일반"], range=["red", "lightgray"]),
                    legend=alt.Legend(title="구분"),
                )
            )
            .interactive()
        )

    # 블루오션 v3 (이론 밀도 우선 + 업체수 보정)
    blue_df, best_range, best_center = find_blue_ocean_v3(
        theoretical_rates_all, bidder_rates_all, hot_start, hot_end, bin_width=0.0005
    )

    chart_gap = None
    blue_desc = ""
    best_range_str = "없음"
    rec_rate = None

    if blue_df is not None and best_range is not None:
        best_range_str = f"{best_range[0]:.3f}% ~ {best_range[1]:.3f}%"
        # best_range = (start, end)
        rec_rate = round(best_range[1], 4) if best_range is not None else None  # 🔥 최댓값 사용!


        # 블루오션 점수 막대 그래프
        blue_plot_df = blue_df.rename(columns={"center": "구간중심", "score": "블루오션점수"})
        chart_gap = (
            alt.Chart(blue_plot_df)
            .mark_bar()
            .encode(
                x=alt.X(
                    "구간중심",
                    title="사정율 구간 중심 (%)",
                    scale=alt.Scale(domain=[hot_start, hot_end]),
                ),
                y=alt.Y("블루오션점수", title="블루오션 점수"),
                tooltip=[
                    "구간중심",
                    "블루오션점수",
                    "theo_count",
                    "bid_count",
                ],
            )
            .properties(title="💎 블루오션 탐지 (핫존 내부)")
            .interactive()
        )

        blue_desc = (
            f"- 이 집행관의 핫존(**{hot_start:.3f}% ~ {hot_end:.3f}%**) 안에서\n"
            f"  1365 이론 조합 밀도는 높지만 실제 투찰 업체 수는 상대적으로 적은\n"
            f"  **최상위 블루오션 구간**은 👉 **{best_range_str}** 입니다.\n"
        )
        if rec_rate is not None:
            blue_desc += (
                f"- 이 구간의 중심값을 기준으로 **추천 투찰 사정율**은 "
                f"👉 **{rec_rate:.4f}%** 입니다.\n"
            )
    else:
        blue_desc = (
            "- 현재 데이터로는 뚜렷한 블루오션 구간이 통계적으로 드러나지 않았습니다. "
            "공고 수를 더 늘려 보시는 것도 좋습니다.\n"
        )

    total_input = len(gongo_list)
    filtered = len(results_for_merge)
    missing = total_input - filtered

    stats = {
        "total": total_input,
        "filtered": filtered,
        "missing": missing,
        "blue_range": best_range_str,
        "rec_rate": rec_rate,
    }

    analysis_text = f"""
- 입력 공고 수: **{total_input}건**
- 집행관 필터 통과 공고 수: **{filtered}건**
- 분석에 사용된 1순위 사정율 개수: **{len(winner_rates)}개**

### 🔥 집행관 핫존
- 실제 1순위 사정율이 가장 많이 몰린 구간(핫존)은  
  👉 **{hot_start:.3f}% ~ {hot_end:.3f}%** 입니다.

### 💎 블루오션 해석
{blue_desc}
"""

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
    )


# -------------------------------------------------
# Streamlit UI
# -------------------------------------------------
def reset_gongo():
    st.session_state["gongo_text"] = ""


st.markdown(
    "<h1 style='font-size:32px;'>🏗 1365 사정율 분석기 (핫존 + 블루오션 + 추천 사정률)</h1>",
    unsafe_allow_html=True,
)

target = st.text_input("🎯 타겟 집행관 (선택 사항, 비우면 전체)", value="")

gongo_input = st.text_area(
    "📄 공고번호 목록 (줄바꿈/콤마 구분)",
    height=200,
    key="gongo_text",
    placeholder="예)\nR25BK01074208-000\nR25BK01071774-000\n...",
)

btn_col1, btn_col2 = st.columns([1, 1])
with btn_col1:
    run_clicked = st.button("🚀 분석 실행", use_container_width=True)
with btn_col2:
    st.button("🧹 초기화", use_container_width=True, on_click=reset_gongo)

# ----- 분석 실행 버튼을 누른 경우에만 API 호출 & 결과 저장 -----
if run_clicked:
    with st.spinner("분석 중입니다... 잠시만 기다려 주세요."):
        result = process_analysis(target, gongo_input)

    # 결과를 세션에 저장해서, 엑셀 다운로드 등으로 rerun 되어도 유지
    (
        logs,
        merged,
        hot_start,
        hot_end,
        analysis_md,
        chart_main,
        chart_gap,
        stats,
        excel_path,
    ) = result

    st.session_state["analysis_result"] = {
        "logs": logs,
        "merged": merged,
        "hot_start": hot_start,
        "hot_end": hot_end,
        "analysis_md": analysis_md,
        "chart_main": chart_main,
        "chart_gap": chart_gap,
        "stats": stats,
        "excel_path": excel_path,
    }

# ----- 세션에 저장된 결과가 있다면 항상 화면에 표시 -----
if "analysis_result" in st.session_state:
    res = st.session_state["analysis_result"]

    logs = res["logs"]
    merged = res["merged"]
    hot_start = res["hot_start"]
    hot_end = res["hot_end"]
    analysis_md = res["analysis_md"]
    chart_main = res["chart_main"]
    chart_gap = res["chart_gap"]
    stats = res["stats"]
    excel_path = res["excel_path"]

    # 로그
    st.markdown("### 📜 로그")
    st.code(logs or "로그 없음", language="text")

    if merged is None or merged.empty:
        st.warning("⚠ 유효한 분석 데이터가 없습니다.")
    else:
        # 요약 카드
        st.markdown("### 📊 요약 카드")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            if hot_start is not None and hot_end is not None:
                st.metric("핫존 시작", f"{hot_start:.4f}%")
        with c2:
            if hot_start is not None and hot_end is not None:
                st.metric("핫존 끝", f"{hot_end:.4f}%")
        with c3:
            st.metric("분석 공고 수", stats.get("filtered", 0))
        with c4:
            st.metric("누락 공고 수", stats.get("missing", 0))

        # 추천 투찰 사정률 카드
        rec_rate = stats.get("rec_rate")
        st.markdown("### 🔥 추천 투찰 사정률")
        if rec_rate is not None:
            st.markdown(
                f"""
<div style="background-color:#FFEFB5; padding:18px; border-radius:12px;
     border:1px solid #E0C772;">
  <div style="font-size:22px; font-weight:700; color:#333;">
    🔥 추천 투찰 사정율 : <span style="color:#C0392B;">{rec_rate:.4f}%</span>
  </div>
  <div style="font-size:14px; margin-top:6px; color:#444;">
    (핫존 + 블루오션 통계 기반 자동 추천 값)
  </div>
</div>
""",
                unsafe_allow_html=True,
            )
        else:
            st.info("블루오션 통계가 부족하여 추천 사정률을 계산할 수 없습니다.")

        # 종합 분석 리포트
        st.markdown("### 🎯 종합 분석 리포트")
        st.markdown(analysis_md)

        # 그래프
        if chart_main is not None:
            st.markdown("### 📈 1순위 사정율 분포 (줌/이동 가능)")
            st.altair_chart(chart_main, use_container_width=True)

        if chart_gap is not None:
            st.markdown("### 💎 블루오션 점수 분포 (핫존 기준)")
            st.altair_chart(chart_gap, use_container_width=True)

        # 통합 테이블
        st.markdown("### 📑 통합 사정율 비교 테이블")
        st.dataframe(merged, use_container_width=True)

        # 엑셀 다운로드 (여기에서 클릭해도 세션에 결과가 남아 있어서 초기화 안 됨)
        if excel_path and os.path.exists(excel_path):
            with open(excel_path, "rb") as f:
                st.download_button(
                    label="📥 엑셀 다운로드",
                    data=f,
                    file_name=os.path.basename(excel_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
