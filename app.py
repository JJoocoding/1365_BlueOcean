import itertools
import json
import os
from io import BytesIO
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


def find_blue_ocean(theoretical_rates, actual_rates, hot_start, hot_end, bin_width=0.1):
    """
    🔵 블루오션 정의 (옵션 1 / 점수 방식 A)
    1) 핫존 안에서
    2) 이론상 1365 조합이 많이 몰린 구간
    3) 그 구간에 실제 1순위는 상대적으로 적은 구간
    """
    if hot_start is None or hot_end is None:
        return [], None, None

    theo = [r for r in theoretical_rates if hot_start <= r <= hot_end]
    act = [r for r in actual_rates if hot_start <= r <= hot_end]

    if len(theo) == 0:
        return [], None, None

    bins = np.arange(hot_start, hot_end + bin_width, bin_width)
    if len(bins) < 2:
        bins = np.array([hot_start, hot_end])

    theo_counts, _ = np.histogram(theo, bins=bins)
    act_counts, bin_edges = np.histogram(act, bins=bins)

    theo_norm = theo_counts / theo_counts.sum()
    act_norm = act_counts / act_counts.sum() if act_counts.sum() > 0 else np.zeros_like(act_counts)

    results = []
    best_range = None
    best_center = None
    best_score = -1.0

    for i in range(len(bin_edges) - 1):
        start = bin_edges[i]
        end = bin_edges[i + 1]
        center = (start + end) / 2

        p_theo = theo_norm[i]
        p_act = act_norm[i]

        if p_theo < 1e-6:
            continue

        # 옵션 A 점수: 이론이 많이 몰릴수록, 실제는 적을수록 점수↑
        score = (p_theo ** 2) * (1 - p_act)

        results.append(
            {
                "start": start,
                "end": end,
                "center": center,
                "p_theo": p_theo,
                "p_act": p_act,
                "score": score,
            }
        )

        if score > best_score:
            best_score = score
            best_range = (start, end)
            best_center = center

    results.sort(key=lambda x: x["score"], reverse=True)
    return results, best_range, best_center


# -------------------------------------------------
# 공고 1건 분석
# -------------------------------------------------
def analyze_gongo(gongo_input_str: str):
    """
    공고 1건 분석
    - df_combined : 1365 조합 + 실제 입찰 업체 사정율
    - top_info    : 1순위 업체 / 사정율 / 집행관
    - df_rates    : 1365 조합 사정율 리스트
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

        # -----------------------
        # 1) 복수예가 (1365 조합용)
        # -----------------------
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
                        combs = itertools.combinations(df1["SA_rate"], 4)
                        rates = [np.mean(c) for c in combs]
                        df_rates = (
                            pd.DataFrame(rates, columns=["rate"])
                            .sort_values("rate")
                            .reset_index(drop=True)
                        )
                        df_rates["조합순번"] = range(1, len(df_rates) + 1)
        except Exception:
            pass

        # -----------------------
        # 2) 낙찰하한율
        # -----------------------
        sucsfbidLwltRate = 0.0
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
                sucsfbidLwltRate = float(items2[0]["sucsfbidLwltRate"])
        except Exception:
            pass

        # -----------------------
        # 3) A값
        # -----------------------
        A_value = get_a_value(gongo_no)

        # -----------------------
        # 4) 개찰결과 (XML)
        # -----------------------
        url4 = (
            "http://apis.data.go.kr/1230000/as/ScsbidInfoService/"
            f"getOpengResultListInfoOpengCompt?serviceKey={SERVICE_KEY}&pageNo=1&numOfRows=999&bidNtceNo={gongo_no}"
        )
        try:
            res4 = requests.get(url4, headers=headers, timeout=10)
        except Exception as e:
            return pd.DataFrame(), f"HTTP 오류 ({gongo_input_str}): {e}", None, pd.DataFrame()

        items4 = []
        try:
            data4 = xmltodict.parse(res4.text)
            items4_raw = data4.get("response", {}).get("body", {}).get("items")
            if isinstance(items4_raw, dict):
                items4 = items4_raw.get("item", [])
            elif isinstance(items4_raw, list):
                items4 = items4_raw
            if isinstance(items4, dict):
                items4 = [items4]
            if not isinstance(items4, list):
                items4 = []
        except Exception:
            items4 = []

        df4 = pd.DataFrame(items4)
        top_info = {"name": "개찰결과 없음", "rate": 0.0, "officer": officer_name}

        if not df4.empty and "bidprcAmt" in df4.columns:
            df4["bidprcAmt"] = pd.to_numeric(df4["bidprcAmt"], errors="coerce")
            df4 = df4.dropna(subset=["bidprcAmt"])

            if not df4.empty:
                top_name = str(df4.iloc[0].get("prcbdrNm", "업체명없음"))

                if sucsfbidLwltRate > 0 and base_price > 0:
                    numerator = ((df4["bidprcAmt"] - A_value) * 100) / sucsfbidLwltRate + A_value
                    df4["rate"] = numerator * 100 / base_price
                else:
                    df4["rate"] = 0.0

                top_row = df4.iloc[0]
                top_rate = float(top_row.get("rate", 0.0))

                top_info = {
                    "name": top_name,
                    "rate": round(top_rate, 5),
                    "officer": officer_name,
                }

                df4 = df4.drop_duplicates(subset=["rate"])
                df4 = df4[(df4["rate"] >= 90) & (df4["rate"] <= 110)]
                df4 = df4[["prcbdrNm", "rate"]].rename(columns={"prcbdrNm": "업체명"})

        # -----------------------
        # 5) 조합 + 실제 통합
        # -----------------------
        if not df_rates.empty:
            df_combined = pd.concat(
                [
                    # 1365 조합은 '조합번호만' 표시
                    df_rates[["rate"]].assign(업체명=df_rates["조합순번"].astype(str)),
                    df4[["업체명", "rate"]],
                ],
                ignore_index=True,
            ).sort_values("rate").reset_index(drop=True)
        else:
            if not df4.empty and "rate" in df4.columns:
                df_combined = df4.sort_values("rate").reset_index(drop=True)
            else:
                df_combined = pd.DataFrame()

        if not df_combined.empty:
            df_combined["rate"] = df_combined["rate"].round(5)
            df_combined["공고번호"] = gongo_no

        return df_combined, None, top_info, df_rates

    except Exception as e:
        return pd.DataFrame(), f"❌ 예외 ({gongo_input_str}): {e}", None, pd.DataFrame()


# -------------------------------------------------
# 전체 실행 + 엑셀 + 그래프 + 추천사정율
# -------------------------------------------------
def process_analysis(target_officer: str, gongo_input: str):
    if not gongo_input.strip():
        return "공고번호를 입력해주세요.", None, None, None, None, None, None, None, None

    if not SERVICE_KEY:
        return (
            "❌ SERVICE_KEY 미설정 (secrets.toml 확인)",
            None,
            None,
            None,
            None,
            None,
            None,
            None,
            None,
        )

    gongo_list = [x.strip() for x in gongo_input.replace(",", "\n").split("\n") if x.strip()]
    target_clean = target_officer.strip()

    logs = []
    results_for_merge = []
    scatter_data = []
    total_actual_rates = []
    total_theoretical_rates = []

    for gongo in gongo_list:
        df, err, top, df_rates_raw = analyze_gongo(gongo)

        if err:
            logs.append(f"❌ {gongo} {err}")
            continue

        if not top:
            logs.append(f"⚠ {gongo}: 1순위 정보 없음")
            continue

        officer = str(top["officer"]).strip()

        # 집행관 필터
        if target_clean:
            if officer != target_clean:
                logs.append(f"⛔ [제외] {gongo} | 집행관: {officer}")
                continue
            else:
                logs.append(
                    f"✅ [포함] {gongo} | 집행관: {officer} | 1순위: {top['name']} ({top['rate']}%)"
                )
        else:
            logs.append(
                f"✅ {gongo} | 집행관: {officer} | 1순위: {top['name']} ({top['rate']}%)"
            )

        if not df.empty:
            results_for_merge.append({"gongo": gongo, "df": df, "top": top})

        if top["rate"] != 0:
            scatter_data.append([top["rate"], gongo, top["name"], officer])
            total_actual_rates.append(top["rate"])

        if not df_rates_raw.empty:
            total_theoretical_rates.extend(df_rates_raw["rate"].tolist())

    if not results_for_merge:
        logs.append("⚠ 집행관 필터 및 데이터 조건을 만족하는 공고가 없습니다.")
        return "\n".join(logs), None, None, None, None, None, None, None, None

    # -----------------------
    # 통합 테이블(가로비교용)
    # -----------------------
    all_rates = pd.concat([r["df"]["rate"] for r in results_for_merge]).unique()
    merged_df = pd.DataFrame({"rate": all_rates}).sort_values("rate").reset_index(drop=True)

    col_index_to_winner = {}
    winner_rate_map = {}

    for res in results_for_merge:
        gn = res["df"]["공고번호"].iloc[0] if "공고번호" in res["df"].columns else res["gongo"]
        winner_name = res["top"]["name"]
        winner_rate = res["top"]["rate"]
        officer_nm = res["top"]["officer"]

        # 🔹 엑셀/화면 헤더에 1순위 업체명 + 사정율 같이 표시
        col_name = f"{gn}\n[{officer_nm}]\n{winner_name}\n({winner_rate:.4f}%)"

        sub_df = res["df"][["rate", "업체명"]].rename(columns={"업체명": col_name})
        merged_df = pd.merge(merged_df, sub_df, on="rate", how="outer")
        col_index_to_winner[col_name] = winner_name
        winner_rate_map[col_name] = winner_rate

    merged_df = merged_df.fillna("")

    # -----------------------
    # 그래프 및 블루오션 / 추천 사정율
    # -----------------------
    chart_main = None
    chart_gap = None
    hot_start = None
    hot_end = None
    best_range = None
    recommended_rate = None

    if scatter_data:
        chart_df = pd.DataFrame(scatter_data, columns=["rate", "공고번호", "업체명", "집행관"])
        min_rate = chart_df["rate"].min()
        max_rate = chart_df["rate"].max()

        hot_start, hot_end, _ = find_hot_zone(total_actual_rates)
        if hot_start is None or hot_end is None:
            hot_start, hot_end = min_rate, max_rate

        def cat(r):
            return "🔥 집중구간" if hot_start <= r <= hot_end else "일반"

        chart_df["구분"] = chart_df["rate"].apply(cat)

        base_chart = alt.Chart(chart_df).encode(
            x=alt.X(
                "rate",
                title="사정율 (%)",
                scale=alt.Scale(domain=[min(min_rate, 98) - 0.2, max(max_rate, 102) + 0.2]),
            ),
            y=alt.Y("공고번호", sort=None, title="공고번호"),
            tooltip=["업체명", "rate", "공고번호", "집행관", "구분"],
        )

        chart_main = (
            base_chart.mark_circle(size=120)
            .encode(
                color=alt.Color(
                    "구분",
                    scale=alt.Scale(domain=["🔥 집중구간", "일반"], range=["red", "lightgray"]),
                    legend=alt.Legend(title="구분"),
                )
            )
            .interactive()
        )

        # 블루오션 + 추천 사정율
        if total_theoretical_rates and total_actual_rates:
            blue_results, best_range, best_center = find_blue_ocean(
                total_theoretical_rates,
                total_actual_rates,
                hot_start,
                hot_end,
                bin_width=0.1,
            )

            if best_center is not None:
                recommended_rate = round(best_center, 4)

            if blue_results:
                gap_df = pd.DataFrame(
                    [
                        {"구간중심": r["center"], "블루오션점수": r["score"]}
                        for r in blue_results
                    ]
                )
                chart_gap = (
                    alt.Chart(gap_df)
                    .mark_bar()
                    .encode(
                        x=alt.X(
                            "구간중심",
                            title="사정율 구간 중심 (%)",
                            scale=alt.Scale(domain=[hot_start, hot_end]),
                        ),
                        y=alt.Y("블루오션점수", title="이론 대비 실제 부족 정도"),
                        tooltip=["구간중심", "블루오션점수"],
                    )
                    .properties(title="💎 블루오션 탐지 (핫존 내)")
                    .interactive()
                )

    # -----------------------
    # 엑셀 파일 생성
    # -----------------------
    excel_buffer = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "통합분석"

    # DF → Worksheet
    for r in dataframe_to_rows(merged_df, index=False, header=True):
        ws.append(r)

    # 헤더 서식
    header_font = Font(bold=True)
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for cell in ws[1]:
        cell.font = header_font
        cell.alignment = header_align

    # 1순위 하이라이트
    highlight_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    for col_idx, col_name in enumerate(merged_df.columns, start=1):
        if col_idx == 1:
            continue
        winner = col_index_to_winner.get(col_name)
        if not winner:
            continue
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if str(cell.value).strip() == winner:
                cell.fill = highlight_fill
                cell.font = Font(bold=True)

    # 공고별 시트(선택사항) - 조합/업체 세부 확인용
    for res in results_for_merge:
        sheet_name = res["gongo"].split("-")[0][:31]
        ws_sub = wb.create_sheet(title=sheet_name)
        for r in dataframe_to_rows(res["df"], index=False, header=True):
            ws_sub.append(r)

    wb.save(excel_buffer)
    excel_buffer.seek(0)

    # -----------------------
    # 분석 리포트 텍스트
    # -----------------------
    total_input = len(gongo_list)
    filtered_count = len(results_for_merge)

    if recommended_rate is not None and best_range is not None:
        blue_text = (
            f"- 이 집행관의 핫존은 **{hot_start:.3f}% ~ {hot_end:.3f}%** 입니다.\n"
            f"- 그 안에서 **이론(1365) 조합은 많이 몰렸지만 실제 1순위는 상대적으로 적은 최상위 블루오션 구간**은\n"
            f"  👉 **{best_range[0]:.3f}% ~ {best_range[1]:.3f}%** 입니다.\n"
            f"- 이 데이터를 기반으로 추천하는 **투찰 사정율**은\n"
            f"  👉 **{recommended_rate:.4f}%** 입니다."
        )
    else:
        blue_text = "- 블루오션을 도출하기에 통계 데이터가 다소 부족합니다. 공고를 더 많이 넣어 보세요."

    analysis_text = f"""
### 🎯 종합 분석 리포트

- 입력 공고 수: **{total_input}건**
- 집행관 필터 통과 공고 수: **{filtered_count}건** (집행관: `{target_clean or "전체"}`)

#### 1. 🔥 집행관 장비 핫존
- 가장 많이 몰린 실제 1순위 사정율 구간: **{hot_start:.3f}% ~ {hot_end:.3f}%**

#### 2. 💎 블루오션 & 추천 투찰 사정율
{blue_text}
"""

    return "\n".join(logs), merged_df, analysis_text, chart_main, chart_gap, hot_start, hot_end, recommended_rate, excel_buffer


# -------------------------------------------------
# Streamlit UI
# -------------------------------------------------
st.markdown("## 🏗 1365 사정율 분석기 (핫존 + 블루오션 + 추천 투찰사정율)")

target = st.text_input("🎯 타겟 집행관 (비우면 전체)", value="")
gongo_input = st.text_area("📄 공고번호 목록 (줄바꿈/콤마 구분)", height=200)

if st.button("🚀 분석 실행"):
    with st.spinner("🔍 분석을 실행하고 있습니다. 잠시만 기다려주세요..."):
        logs, merged, analysis_md, chart_main, chart_gap, hot_start, hot_end, rec_rate, excel_buf = process_analysis(
            target, gongo_input
        )

    # 로그
    st.subheader("📋 로그")
    st.code(logs or "로그 없음", language="text")

    if merged is None or merged.empty:
        st.warning("⚠ 유효한 분석 데이터가 없습니다.")
        st.stop()

    # 🔹 상단 요약 카드
    st.subheader("📊 요약 카드")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("핫존 시작", f"{hot_start:.4f}%" if hot_start else "-")
    with c2:
        st.metric("핫존 끝", f"{hot_end:.4f}%" if hot_end else "-")
    with c3:
        st.metric("추천 투찰사정율", f"{rec_rate:.4f}%" if rec_rate else "-")

    # 🎯 추천 투찰사정율 강조 박스
    if rec_rate is not None:
        st.markdown(
            f"""
        <div style="
            padding:18px;
            background-color:#FFF3CD;
            border-left:6px solid #FFB800;
            border-radius:6px;
            font-size:20px;
	    color:#333333;
            Line-height:1.6;
        ">
            🔥 <strong>추천 투찰 사정율 :</strong> 
            <span style="color:#C0392B; font-size:26px; font-weight:700;">{rec_rate:.4f}%</span>
            <br>
            (핫존 + 블루오션 통계 기반 자동 추천 값)
        </div>
        """,
            unsafe_allow_html=True,
        )

    # 📊 분석 리포트
    st.markdown(analysis_md)

    # 그래프
    if chart_main is not None:
        st.subheader("📈 사정율 분포 (1순위 기준, 줌/이동 가능)")
        st.altair_chart(chart_main, use_container_width=True)

    if chart_gap is not None:
        st.subheader("💎 블루오션 점수 차트 (핫존 내)")
        st.altair_chart(chart_gap, use_container_width=True)

    # 테이블
    st.subheader("📑 통합 사정율 비교 테이블")
    st.dataframe(merged, use_container_width=True)

    # 📥 엑셀 다운로드
    if excel_buf is not None:
        st.download_button(
            label="📥 엑셀 다운로드",
            data=excel_buf,
            file_name=f"사정율분석_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
