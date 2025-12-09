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
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

# -------------------------------------------------
# Page Setup
# -------------------------------------------------
st.set_page_config(page_title="1365 사정율 분석기", layout="wide")

# -------------------------------------------------
# Secrets Load
# -------------------------------------------------
try:
    SERVICE_KEY = st.secrets["SERVICE_KEY"]
except Exception:
    SERVICE_KEY = ""


# -------------------------------------------------
# Utility Functions
# -------------------------------------------------
def get_headers():
    return {"User-Agent": "Mozilla/5.0"}

def safe_get_items(json_data):
    try:
        body = json_data.get("response", {}).get("body", {})
        items = body.get("items")
        if isinstance(items, list):
            return items
        if isinstance(items, dict):
            item = items.get("item")
            if isinstance(item, dict):
                return [item]
            return item or []
        return []
    except:
        return []


# -------------------------------------------------
# A값 / 집행관
# -------------------------------------------------
def get_a_value(gongo_no: str) -> float:
    try:
        url = (
            "http://apis.data.go.kr/1230000/ad/BidPublicInfoService/"
            "getBidPblancListInfoCnstwkBsisAmount"
            f"?inqryDiv=2&bidNtceNo={gongo_no}&pageNo=1&numOfRows=10&type=json&ServiceKey={SERVICE_KEY}"
        )
        res = requests.get(url, headers=get_headers(), timeout=7)
        items = safe_get_items(res.json())
        if not items:
            return 0.0
        df = pd.DataFrame(items)
        cols = [
            "sftyMngcst","sftyChckMngcst","rtrfundNon",
            "mrfnHealthInsrprm","npnInsrprm","odsnLngtrmrcprInsrprm","qltyMngcst"
        ]
        valid = [c for c in cols if c in df.columns]
        return (
            df[valid]
            .apply(pd.to_numeric, errors="coerce")
            .fillna(0)
            .sum(axis=1)
            .iloc[0]
        )
    except:
        return 0.0


def get_officer_name_final(gongo_no: str) -> str:
    try:
        url = (
            "http://apis.data.go.kr/1230000/ad/BidPublicInfoService/"
            f"getBidPblancListInfoCnstwk?inqryDiv=2&bidNtceNo={gongo_no}"
            "&pageNo=1&numOfRows=1&type=json&ServiceKey={SERVICE_KEY}"
        )
        res = requests.get(url, headers=get_headers(), timeout=7)
        items = safe_get_items(res.json())
        if not items:
            return "확인불가"
        item = items[0]
        for key in ["exctvNm", "chrgrNm", "ntceChrgrNm"]:
            if key in item and str(item[key]).strip():
                return str(item[key]).strip()
        return "확인불가"
    except:
        return "확인불가"


# -------------------------------------------------
# 핫존 / 블루오션
# -------------------------------------------------
def find_hot_zone(actual_rates, window=0.3, step=0.05):
    """ 1순위 사정율이 가장 몰린 구간 탐색 """
    if not actual_rates:
        return None, None, 0
    rates_sorted = sorted(actual_rates)
    min_r, max_r = min(rates_sorted), max(rates_sorted)

    best_s, best_e, best_count = None, None, -1

    cur = min_r
    while cur <= max_r:
        end = cur + window
        count = sum(cur <= r <= end for r in rates_sorted)
        if count > best_count:
            best_s, best_e, best_count = cur, end, count
        cur += step

    return best_s, best_e, best_count


def find_blue_ocean(theoretical, actual, hot_s, hot_e, bw=0.1):
    if hot_s is None or hot_e is None:
        return [], None, None

    theo = [r for r in theoretical if hot_s <= r <= hot_e]
    act = [r for r in actual if hot_s <= r <= hot_e]

    if len(theo) == 0:
        return [], None, None

    bins = np.arange(hot_s, hot_e + bw, bw)
    theo_counts, _ = np.histogram(theo, bins=bins)
    act_counts, edges = np.histogram(act, bins=bins)

    theo_norm = theo_counts / theo_counts.sum()
    act_norm = act_counts / max(act_counts.sum(), 1)

    results = []
    best_score = -1
    best_range = None

    for i in range(len(edges) - 1):
        s, e = edges[i], edges[i + 1]
        center = (s + e) / 2

        p_theo = theo_norm[i]
        p_act = act_norm[i]

        if p_theo < 1e-5:
            continue

        # 옵션 1 + Score-A
        score = p_theo * (p_theo - p_act)
        if score <= 0:
            continue

        results.append({
            "start": s, "end": e, "center": center,
            "p_theo": p_theo, "p_act": p_act, "score": score
        })

        if score > best_score:
            best_score = score
            best_range = (s, e)

    results.sort(key=lambda x: x["score"], reverse=True)

    # 추천 투찰 사정률 (중심값)
    recommended = round(best_range[0] + (best_range[1] - best_range[0]) / 2, 4) if best_range else None

    return results, best_range, recommended


# -------------------------------------------------
# 공고 1건 분석
# -------------------------------------------------
def analyze_gongo(gongo_no_full: str):
    try:
        if "-" in gongo_no_full:
            gongo_no, gongo_ord = gongo_no_full.split("-")
        else:
            gongo_no, gongo_ord = gongo_no_full, "00"

        officer = get_officer_name_final(gongo_no)

        # 1) 복수예가 → 1365
        url1 = (
            "http://apis.data.go.kr/1230000/as/ScsbidInfoService/"
            "getOpengResultListInfoCnstwkPreparPcDetail"
            f"?inqryDiv=2&bidNtceNo={gongo_no}&bidNtceOrd={gongo_ord}"
            f"&pageNo=1&numOfRows=30&type=json&ServiceKey={SERVICE_KEY}"
        )
        items1 = safe_get_items(requests.get(url1, headers=get_headers()).json())

        df_rates = pd.DataFrame()
        base_price = 0

        if items1:
            df1 = pd.json_normalize(items1)
            df1 = df1.astype(float)
            base_price = df1["bssamt"].iloc[0]
            df1["SA_rate"] = df1["bsisPlnprc"] / df1["bssamt"] * 100

            if len(df1) >= 4:
                rates = [np.mean(c) for c in itertools.combinations(df1["SA_rate"], 4)]
                df_rates = pd.DataFrame({"rate": rates}).sort_values("rate")
                df_rates["idx"] = range(1, len(df_rates) + 1)

        # 2) A값
        A_value = get_a_value(gongo_no)

        # 3) 개찰결과
        url4 = (
            "http://apis.data.go.kr/1230000/as/ScsbidInfoService/"
            f"getOpengResultListInfoOpengCompt?serviceKey={SERVICE_KEY}"
            f"&pageNo=1&numOfRows=999&bidNtceNo={gongo_no}"
        )
        xml_data = xmltodict.parse(requests.get(url4, headers=get_headers()).text)
        items4 = xml_data.get("response", {}).get("body", {}).get("items", {})
        items4 = items4.get("item", []) if isinstance(items4, dict) else items4
        if isinstance(items4, dict): items4 = [items4]

        df4 = pd.DataFrame(items4)
        df4["bidprcAmt"] = pd.to_numeric(df4["bidprcAmt"], errors="coerce")
        df4 = df4.dropna(subset=["bidprcAmt"])

        top_row = df4.iloc[0]
        sucsfbid = float(top_row.get("sucsfbidLwltRate", 0)) or 0

        df4["rate"] = ((df4["bidprcAmt"] - A_value) * 100 / sucsfbid + A_value) * 100 / base_price

        top_name = top_row["prcbdrNm"]
        top_rate = float(df4["rate"].iloc[0])

        df4 = df4[["prcbdrNm","rate"]].rename(columns={"prcbdrNm":"업체명"})

        # 결합
        if not df_rates.empty:
            combined = pd.concat(
                [df_rates[["rate"]].assign(업체명=df_rates["idx"].astype(str)), df4],
                ignore_index=True
            ).sort_values("rate")
        else:
            combined = df4.copy()

        combined["공고"] = gongo_no

        return combined, officer, top_name, top_rate, df_rates

    except Exception as e:
        return pd.DataFrame(), None, None, 0, pd.DataFrame()


# -------------------------------------------------
# 전체 프로세스 (session_state 사용!)
# -------------------------------------------------
def run_analysis(target, gongo_text):

    # 입력 공고번호 정리
    gongo_list = [g.strip() for g in gongo_text.replace(",", "\n").split("\n") if g.strip()]

    logs = []
    merged_list = []
    actual_rates = []
    theoretical_rates = []

    for g in gongo_list:
        df, officer, top_name, top_rate, df_rates = analyze_gongo(g)

        if officer is None:
            logs.append(f"❌ {g}: 분석 실패")
            continue

        logs.append(f"📌 {g} | 집행관={officer} | 1순위={top_name}({top_rate:.4f})")

        if target and officer != target:
            logs.append(f"➡ 제외: 집행관 불일치")
            continue

        if not df.empty:
            merged_list.append({"gongo": g, "df": df, "top": top_name, "rate": top_rate})

        actual_rates.append(top_rate)
        if not df_rates.empty:
            theoretical_rates.extend(df_rates["rate"].tolist())

    if not merged_list:
        return logs, None, None, None, None, None, None, None, None

    # 통합 DF
    all_rates = sorted({r for m in merged_list for r in m["df"]["rate"].tolist()})
    merged_df = pd.DataFrame({"rate": all_rates})

    name_map = {}
    for m in merged_list:
        g = m["gongo"]
        col = f"{g}\n{m['top']}\n{m['rate']:.4f}"
        sub = m["df"][["rate","업체명"]].rename(columns={"업체명": col})
        merged_df = merged_df.merge(sub, on="rate", how="left")
        name_map[col] = m["top"]

    # 핫존
    hot_s, hot_e, _ = find_hot_zone(actual_rates)

    # 블루오션
    blue_results, blue_range, recommended = find_blue_ocean(
        theoretical_rates, actual_rates, hot_s, hot_e
    )

    # 리포트
    if blue_range:
        report = (
            f"### 🔍 블루오션 분석 결과\n"
            f"• 핫존: **{hot_s:.3f} ~ {hot_e:.3f}%**\n"
            f"• 블루오션 구간: **{blue_range[0]:.3f} ~ {blue_range[1]:.3f}%**\n"
            f"• ⭐ 추천 투찰 사정률: **{recommended:.4f}%**"
        )
    else:
        report = "블루오션 구간 부족"

    # 그래프 생성
    chart_df = pd.DataFrame({
        "rate":[m["rate"] for m in merged_list],
        "공고":[m["gongo"] for m in merged_list]
    })
    chart = alt.Chart(chart_df).mark_circle(size=120).encode(
        x="rate",
        y="공고",
        tooltip=["rate","공고"]
    ).interactive()

    # Gap 차트
    if blue_results:
        gap_df = pd.DataFrame(blue_results)
        gap_chart = alt.Chart(gap_df).mark_bar().encode(
            x="center",
            y="score",
            tooltip=["start","end","score"]
        ).interactive()
    else:
        gap_chart = None

    # 엑셀 생성
    excel_name = f"사정율분석_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    wb = Workbook()
    ws = wb.active; ws.title = "통합"

    for r in dataframe_to_rows(merged_df, index=False, header=True):
        ws.append(r)

    # 헤더 Bold
    for c in ws[1]:
        c.font = Font(bold=True)
        c.alignment = Alignment(horizontal="center")

    # highlight
    fill = PatternFill(start_color="FFFF00", fill_type="solid")
    for col_idx, col_name in enumerate(merged_df.columns, start=1):
        if col_idx == 1: continue
        winner = name_map.get(col_name)
        for row_idx in range(2, ws.max_row+1):
            if ws.cell(row=row_idx, column=col_idx).value == winner:
                ws.cell(row=row_idx, column=col_idx).fill = fill

    wb.save(excel_name)

    return (
        logs, merged_df, hot_s, hot_e, report,
        chart, gap_chart, recommended, excel_name
    )


# -------------------------------------------------
# Streamlit UI
# -------------------------------------------------
st.title("🏗 1365 사정율 분석기 (핫존 + 블루오션 + 추천 사정률)")

# -------- 입력 UI --------
c1, c2 = st.columns([3,1])
with c1:
    target = st.text_input("🎯 타겟 집행관 (비우면 전체)")
with c2:
    if st.button("🧹 초기화"):
        st.session_state.clear()
        st.experimental_rerun()

gongo_input = st.text_area("📄 공고번호 목록 (줄바꿈/콤마)", height=180)

run_btn = st.button("🚀 분석 실행")

# -------- 실행 --------
if run_btn:
    with st.spinner("분석 중입니다..."):
        logs, merged_df, hot_s, hot_e, report, chart, gap_chart, recommended, excel_name = run_analysis(
            target, gongo_input
        )

    st.session_state["logs"] = logs
    st.session_state["merged_df"] = merged_df
    st.session_state["hot_s"] = hot_s
    st.session_state["hot_e"] = hot_e
    st.session_state["report"] = report
    st.session_state["chart"] = chart
    st.session_state["gap_chart"] = gap_chart
    st.session_state["recommended"] = recommended
    st.session_state["excel_name"] = excel_name

# -------- 출력 영역 --------
if "merged_df" in st.session_state and st.session_state["merged_df"] is not None:

    st.subheader("📋 로그")
    st.code("\n".join(st.session_state["logs"]))

    st.markdown(st.session_state["report"])

    # 추천 사정률 박스
    st.success(f"✨ **추천 투찰 사정률: {st.session_state['recommended']:.4f}%**")

    st.subheader("📊 통합 사정율 비교표")
    st.dataframe(st.session_state["merged_df"], use_container_width=True)

    # 그래프
    if st.session_state["chart"] is not None:
        st.subheader("📈 사정율 분포도")
        st.altair_chart(st.session_state["chart"], use_container_width=True)

    if st.session_state["gap_chart"] is not None:
        st.subheader("💎 블루오션 Gap 차트")
        st.altair_chart(st.session_state["gap_chart"], use_container_width=True)

    # 다운로드 버튼
    if "excel_name" in st.session_state:
        with open(st.session_state["excel_name"], "rb") as f:
            st.download_button(
                label="📥 엑셀 다운로드",
                data=f,
                file_name=st.session_state["excel_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_excel"
            )
