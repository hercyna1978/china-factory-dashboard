import os
import pandas as pd
import plotly.express as px
import streamlit as st

# =========================================================
# 설정
# =========================================================
st.set_page_config(
    page_title="중국공장 물류 운영 대시보드",
    page_icon="📊",
    layout="wide",
)

FILE_NAME = "data(2).xlsx"
SHEETS = {
    "master": "01_상품Master",
    "snapshot": "02_재고스냅샷",
    "inbound": "03_입고",
    "outbound": "04_출고",
    "defect": "05_불량관리",
    "current": "08_현재고",
}
FACTORIES = ["C2", "C5", "C2-S", "미상"]
DEFECT_TYPES = ["테불량", "렌즈불량", "전체불량", "분류어려움"]

# 공장 / 불량유형 / 페이지별 컬러 팔레트 (라이트·다크 모드 모두에서 잘 보이도록
# 채도가 있는 accent 컬러만 고정하고, 배경/글자색은 Streamlit 기본값을 그대로 사용)
FACTORY_COLORS = {
    "C2": "#3B82F6",
    "C5": "#14B8A6",
    "C2-S": "#F59E0B",
    "미상": "#94A3B8",
}
DEFECT_COLORS = {
    "테불량": "#F43F5E",
    "렌즈불량": "#8B5CF6",
    "전체불량": "#F59E0B",
    "분류어려움": "#94A3B8",
}
PAGE_ACCENT = {
    "경영진 요약": "#6366F1",
    "공장별 운영": "#3B82F6",
    "재고 현황": "#14B8A6",
    "입출고 추이": "#22C55E",
    "품질 현황": "#F43F5E",
    "데이터 품질": "#64748B",
}
METRIC_COLORS = {
    "SKU": "#3B82F6",
    "현재고": "#14B8A6",
    "누적 입고": "#22C55E",
    "입고": "#22C55E",
    "누적 출고": "#6366F1",
    "출고": "#6366F1",
    "불량 수량": "#F43F5E",
    "불량": "#F43F5E",
    "불량률": "#F59E0B",
    "재고 보유 수량": "#14B8A6",
    "재고 0 SKU": "#94A3B8",
    "마이너스 재고 SKU": "#F43F5E",
    "출고 수량": "#6366F1",
}

# 라이트/다크 모드에서 모두 자연스럽게 보이도록 배경은 반투명(rgba)만 사용하고
# 글자색은 지정하지 않아 Streamlit 테마의 글자색을 그대로 상속받는다.
st.markdown("""
<style>
.block-container {padding-top: 1.4rem; padding-bottom: 2rem;}
[data-testid="stMetric"] {
    border: 1px solid rgba(128,128,128,.22);
    border-radius: 12px;
    padding: 12px 16px;
    background: rgba(128,128,128,.04);
}
.metric-card {
    border: 1px solid rgba(128,128,128,.22);
    border-left: 4px solid var(--accent, #6366F1);
    border-radius: 12px;
    padding: 14px 16px;
    background: rgba(128,128,128,.04);
    height: 100%;
}
.metric-card .m-label {
    font-size: 13px;
    opacity: .72;
    margin-bottom: 6px;
    font-weight: 600;
}
.metric-card .m-value {
    font-size: 26px;
    font-weight: 700;
    line-height: 1.1;
}
.section-title {
    display: flex;
    align-items: center;
    gap: 9px;
    margin: 22px 0 10px 0;
}
.section-title .bar {
    width: 5px;
    height: 20px;
    border-radius: 3px;
}
.section-title .txt {
    font-size: 17px;
    font-weight: 700;
}
.page-badge {
    display: inline-block;
    padding: 3px 12px;
    border-radius: 999px;
    font-size: 12.5px;
    font-weight: 700;
    background: rgba(128,128,128,.12);
    border: 1px solid rgba(128,128,128,.25);
    margin-bottom: 6px;
}
</style>
""", unsafe_allow_html=True)


def metric_card(col, label, value, color=None):
    """색상 강조 막대가 있는 커스텀 지표 카드 (라이트/다크 모드 자동 대응)."""
    accent = color or METRIC_COLORS.get(label, "#6366F1")
    col.markdown(
        f"""
        <div class="metric-card" style="--accent:{accent};">
            <div class="m-label">{label}</div>
            <div class="m-value">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def section_title(text, color="#6366F1"):
    st.markdown(
        f"""
        <div class="section-title">
            <div class="bar" style="background:{color};"></div>
            <div class="txt">{text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def page_badge(text, color):
    st.markdown(
        f'<span class="page-badge" style="color:{color};border-color:{color}55;">{text}</span>',
        unsafe_allow_html=True,
    )


# =========================================================
# 기본 함수
# =========================================================
def code(x):
    if pd.isna(x):
        return ""
    s = str(x).strip().upper()
    return s[:-2] if s.endswith(".0") else s


def factory(x):
    """
    Master의 공급처상품명으로 공장 분류.
    C2-S를 먼저 검사하는 이유: C2-S 안에 C2가 포함되어 있기 때문.
    """
    if pd.isna(x):
        return "미상"
    s = str(x).strip().upper().replace(" ", "")
    if not s or s in {"NAN", "NONE", "0"}:
        return "미상"
    if "C2-S" in s:
        return "C2-S"
    if "C2" in s:
        return "C2"
    if "C5" in s:
        return "C5"
    return "미상"


def defect_type_group(x):
    """
    불량유형 텍스트를 4가지로 단순 분류.
    '테' -> 테불량, '렌즈' -> 렌즈불량, '전체' -> 전체불량, 그 외 -> 분류어려움.
    (텍스트에 여러 키워드가 섞여 있지 않아 우선순위 이슈는 없음)
    """
    if pd.isna(x):
        return "분류어려움"
    s = str(x)
    if "테" in s:
        return "테불량"
    if "렌즈" in s:
        return "렌즈불량"
    if "전체" in s:
        return "전체불량"
    return "분류어려움"


def num(x):
    return f"{x:,.0f}"


def pct(x):
    return "-" if pd.isna(x) else f"{x:.1f}%"


def filter_df(df, fac, cat):
    out = df
    if fac != "전체":
        out = out[out["공장"] == fac]
    if cat != "전체":
        out = out[out["카테고리"] == cat]
    return out


def chart_bar(df, x, y, title, horizontal=False):
    """공장/불량유형처럼 정해진 컬러 팔레트가 있는 컬럼은 자동으로 색을 입혀준다."""
    color_map = None
    if x == "공장":
        color_map = FACTORY_COLORS
    elif x == "불량유형_그룹":
        color_map = DEFECT_COLORS

    common_kwargs = dict(title=title, text=y)
    if color_map is not None:
        common_kwargs.update(color=x, color_discrete_map=color_map)

    if horizontal:
        fig = px.bar(df.sort_values(y), x=y, y=x, orientation="h", **common_kwargs)
    else:
        fig = px.bar(df, x=x, y=y, **common_kwargs)

    fig.update_traces(texttemplate="%{text:,.0f}", textposition="outside")
    fig.update_layout(
        height=390,
        margin=dict(l=20, r=20, t=60, b=20),
        xaxis_title=None,
        yaxis_title="수량" if not horizontal else None,
        showlegend=False,
    )
    st.plotly_chart(fig, use_container_width=True, theme="streamlit")


def add_master_info(df, maps):
    out = df.copy()
    out["_상품코드"] = out["상품코드"].map(code)
    out["공장"] = out["_상품코드"].map(maps["factory"]).fillna("미상")
    out["상품명_Master"] = out["_상품코드"].map(maps["name"]).fillna(out["상품명"])
    out["카테고리"] = out["_상품코드"].map(maps["category"]).fillna("미상")
    out["판매상태"] = out["_상품코드"].map(maps["status"]).fillna("미상")
    return out


# =========================================================
# 데이터 읽기
# =========================================================
@st.cache_data(show_spinner=False)
def read_excel(source_path):
    excel = pd.ExcelFile(source_path)
    missing = [v for v in SHEETS.values() if v not in excel.sheet_names]
    if missing:
        raise ValueError("필수 시트가 없습니다: " + ", ".join(missing))

    master = pd.read_excel(excel, SHEETS["master"])
    snapshot = pd.read_excel(excel, SHEETS["snapshot"])
    inbound = pd.read_excel(excel, SHEETS["inbound"])
    outbound = pd.read_excel(excel, SHEETS["outbound"])
    defect = pd.read_excel(excel, SHEETS["defect"])
    current = pd.read_excel(excel, SHEETS["current"])

    required = {
        "상품Master": ["상품코드", "상품명", "카테고리", "출시일", "판매상태", "공급처상품명"],
        "재고스냅샷": ["기준일", "상품코드", "상품명", "현재고수량"],
        "입고": ["입고일", "상품코드", "상품명", "입고수량", "중국공장"],
        "출고": ["출고일", "상품코드", "상품명", "출고수량", "판매채널"],
        "불량관리": ["발생일", "상품코드", "상품명", "불량수량", "불량유형", "중국공장"],
        "현재고": ["기준일", "상품코드", "상품명", "현재고수량"],
    }
    loaded = {
        "상품Master": master, "재고스냅샷": snapshot, "입고": inbound,
        "출고": outbound, "불량관리": defect, "현재고": current,
    }
    for name, cols in required.items():
        missing_cols = [c for c in cols if c not in loaded[name].columns]
        if missing_cols:
            raise ValueError(f"{name} 시트에 열이 없습니다: {', '.join(missing_cols)}")

    # 날짜/숫자
    for df, col in [
        (master, "출시일"), (snapshot, "기준일"), (inbound, "입고일"),
        (outbound, "출고일"), (defect, "발생일"), (current, "기준일")
    ]:
        df[col] = pd.to_datetime(df[col], errors="coerce")

    for df, col in [
        (snapshot, "현재고수량"), (inbound, "입고수량"),
        (outbound, "출고수량"), (defect, "불량수량"), (current, "현재고수량")
    ]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # 상품Master 기준 공장
    master["_상품코드"] = master["상품코드"].map(code)
    master["공장"] = master["공급처상품명"].apply(factory)

    maps = {
        "factory": master.drop_duplicates("_상품코드").set_index("_상품코드")["공장"].to_dict(),
        "name": master.drop_duplicates("_상품코드").set_index("_상품코드")["상품명"].to_dict(),
        "category": master.drop_duplicates("_상품코드").set_index("_상품코드")["카테고리"].to_dict(),
        "status": master.drop_duplicates("_상품코드").set_index("_상품코드")["판매상태"].to_dict(),
    }

    snapshot = add_master_info(snapshot, maps)
    inbound = add_master_info(inbound, maps)
    outbound = add_master_info(outbound, maps)
    defect = add_master_info(defect, maps)
    current = add_master_info(current, maps)

    # 불량유형 4종 재분류 (테불량 / 렌즈불량 / 전체불량 / 분류어려움)
    defect["불량유형_그룹"] = defect["불량유형"].apply(defect_type_group)

    # 현재고는 최신 기준일만 사용
    if current["기준일"].notna().any():
        latest_current = current["기준일"].max()
        current = current[current["기준일"] == latest_current].copy()
    else:
        latest_current = pd.NaT

    return {
        "master": master, "snapshot": snapshot, "inbound": inbound,
        "outbound": outbound, "defect": defect, "current": current,
        "latest_current": latest_current,
    }


def monthly(df, date_col, qty_col, label):
    x = df.dropna(subset=[date_col]).copy()
    if x.empty:
        return pd.DataFrame(columns=["월", label])
    x["월"] = x[date_col].dt.to_period("M").astype(str)
    return x.groupby("월", as_index=False)[qty_col].sum().rename(columns={qty_col: label})


def factory_kpi(data):
    rows = []
    for f in FACTORIES:
        m = data["master"][data["master"]["공장"] == f]
        c = data["current"][data["current"]["공장"] == f]
        i = data["inbound"][data["inbound"]["공장"] == f]
        o = data["outbound"][data["outbound"]["공장"] == f]
        d = data["defect"][data["defect"]["공장"] == f]

        out_qty = o["출고수량"].sum()
        defect_qty = d["불량수량"].sum()

        rows.append({
            "공장": f,
            "SKU": m["_상품코드"].nunique(),
            "현재고": c["현재고수량"].sum(),
            "입고": i["입고수량"].sum(),
            "출고": out_qty,
            "불량": defect_qty,
            "불량률": defect_qty / out_qty * 100 if out_qty else 0,
        })
    return pd.DataFrame(rows)


def style_negative(df, col):
    """마이너스 값을 은은한 붉은색 배경으로 강조 (라이트/다크 모드 모두 대응하도록 반투명 사용)."""
    def _hl(v):
        try:
            return "background-color: rgba(244,63,94,0.18);" if v < 0 else ""
        except TypeError:
            return ""
    return df.style.map(_hl, subset=[col])


# =========================================================
# 파일 선택
# =========================================================
with st.sidebar:
    st.markdown("### 데이터")
    uploaded = st.file_uploader("Excel 파일 선택", type=["xlsx"])

if uploaded:
    try:
        # 업로드 파일은 현재 세션에서만 사용
        data = read_excel(uploaded)
        source_name = uploaded.name
    except Exception as e:
        st.error(str(e))
        st.stop()
else:
    if not os.path.exists(FILE_NAME):
        st.error(f"'{FILE_NAME}' 파일이 없습니다. app.py와 같은 폴더에 넣어주세요.")
        st.stop()
    try:
        # 파일 수정시간이 바뀌면 cache도 새로 계산
        data = read_excel(FILE_NAME)
        source_name = FILE_NAME
    except Exception as e:
        st.error(str(e))
        st.stop()


master = data["master"]
current = data["current"]
inbound = data["inbound"]
outbound = data["outbound"]
defect = data["defect"]
snapshot = data["snapshot"]


# =========================================================
# 사이드바
# =========================================================
with st.sidebar:
    st.markdown("---")
    st.markdown("### 조회")

    page = st.radio(
        "화면",
        ["경영진 요약", "공장별 운영", "재고 현황", "입출고 추이", "품질 현황", "데이터 품질"],
        label_visibility="collapsed",
    )

    selected_factory = st.selectbox("공장", ["전체"] + FACTORIES)

    categories = sorted(master["카테고리"].dropna().astype(str).unique())
    selected_category = st.selectbox("카테고리", ["전체"] + categories)

    st.markdown("---")
    st.caption(f"데이터: {source_name}")
    if pd.notna(data["latest_current"]):
        st.caption(f"현재고 기준일: {data['latest_current'].strftime('%Y-%m-%d')}")

    dates = pd.concat([inbound["입고일"], outbound["출고일"], defect["발생일"]]).dropna()
    if not dates.empty:
        st.caption(f"거래기간: {dates.min():%Y-%m-%d} ~ {dates.max():%Y-%m-%d}")


# =========================================================
# 제목
# =========================================================
accent = PAGE_ACCENT.get(page, "#6366F1")
page_badge(page, accent)
st.title("중국공장 물류 운영 대시보드")
st.caption("공장 구분은 상품Master의 '공급처상품명'을 기준으로 자동 분류합니다.")


# =========================================================
# 1. 경영진 요약
# =========================================================
if page == "경영진 요약":
    m = filter_df(master, selected_factory, selected_category)
    c = filter_df(current, selected_factory, selected_category)
    i = filter_df(inbound, selected_factory, selected_category)
    o = filter_df(outbound, selected_factory, selected_category)
    d = filter_df(defect, selected_factory, selected_category)

    sku = m["_상품코드"].nunique()
    stock = c["현재고수량"].sum()
    in_qty = i["입고수량"].sum()
    out_qty = o["출고수량"].sum()
    defect_qty = d["불량수량"].sum()
    defect_rate = defect_qty / out_qty * 100 if out_qty else 0

    cols = st.columns(6)
    for col, title, value in [
        (cols[0], "SKU", num(sku)),
        (cols[1], "현재고", num(stock)),
        (cols[2], "누적 입고", num(in_qty)),
        (cols[3], "누적 출고", num(out_qty)),
        (cols[4], "불량 수량", num(defect_qty)),
        (cols[5], "불량률", pct(defect_rate)),
    ]:
        metric_card(col, title, value)

    section_title("공장별 핵심 현황", accent)
    kpi = factory_kpi(data)
    if selected_factory != "전체":
        kpi = kpi[kpi["공장"] == selected_factory]

    a, b = st.columns(2)
    with a:
        chart_bar(kpi, "공장", "현재고", "공장별 현재고")
    with b:
        chart_bar(kpi, "공장", "출고", "공장별 출고")

    table = kpi.copy()
    for col in ["SKU", "현재고", "입고", "출고", "불량"]:
        table[col] = table[col].map(num)
    table["불량률"] = table["불량률"].map(pct)

    section_title("공장별 KPI", accent)
    st.dataframe(table, use_container_width=True, hide_index=True)


# =========================================================
# 2. 공장별 운영
# =========================================================
elif page == "공장별 운영":
    section_title("공장별 운영 실적", accent)
    kpi = factory_kpi(data)
    if selected_factory != "전체":
        kpi = kpi[kpi["공장"] == selected_factory]

    f = filter_df(inbound, selected_factory, selected_category)
    g = filter_df(outbound, selected_factory, selected_category)
    h = filter_df(current, selected_factory, selected_category)
    q = filter_df(defect, selected_factory, selected_category)

    cols = st.columns(4)
    metric_card(cols[0], "입고", num(f["입고수량"].sum()))
    metric_card(cols[1], "출고", num(g["출고수량"].sum()))
    metric_card(cols[2], "현재고", num(h["현재고수량"].sum()))
    metric_card(cols[3], "불량", num(q["불량수량"].sum()))

    for y, title in [("입고", "공장별 입고"), ("출고", "공장별 출고"), ("불량", "공장별 불량")]:
        chart_bar(kpi, "공장", y, title)

    sku = (
        master.groupby("공장")["_상품코드"].nunique()
        .reindex(FACTORIES, fill_value=0)
        .reset_index(name="SKU")
    )
    if selected_factory != "전체":
        sku = sku[sku["공장"] == selected_factory]
    chart_bar(sku, "공장", "SKU", "공장별 SKU 수")


# =========================================================
# 3. 재고 현황
# =========================================================
elif page == "재고 현황":
    section_title("현재 재고 현황", accent)
    df = filter_df(current, selected_factory, selected_category)

    cols = st.columns(4)
    metric_card(cols[0], "현재고", num(df["현재고수량"].sum()))
    metric_card(cols[1], "재고 보유 수량", num(df.loc[df["현재고수량"] > 0, "현재고수량"].sum()))
    metric_card(cols[2], "재고 0 SKU", num((df["현재고수량"] == 0).sum()))
    metric_card(cols[3], "마이너스 재고 SKU", num((df["현재고수량"] < 0).sum()))

    stock = df.groupby("공장", as_index=False)["현재고수량"].sum().rename(columns={"현재고수량": "현재고"})
    chart_bar(stock, "공장", "현재고", "공장별 현재고")

    top = (
        df.groupby(["_상품코드", "상품명_Master", "공장", "카테고리"], as_index=False)["현재고수량"]
        .sum()
        .sort_values("현재고수량", ascending=False)
        .head(20)
        .rename(columns={"_상품코드": "상품코드", "상품명_Master": "상품명", "현재고수량": "현재고"})
    )

    section_title("현재고 상위 20 SKU", accent)
    st.dataframe(
        top.style.format({"현재고": num}),
        use_container_width=True,
        hide_index=True,
    )

    # NOTE: 원본 코드는 `.sort_values(...)[colA, colB, ...]` 형태로 되어 있어
    # (튜플로 컬럼을 선택 -> KeyError) 이 페이지에서 오류가 발생했습니다.
    # 아래처럼 이중 대괄호 `[[...]]`로 열을 선택해야 정상 동작합니다.
    negative = (
        df[df["현재고수량"] < 0]
        .sort_values("현재고수량")[
            ["_상품코드", "상품명_Master", "공장", "카테고리", "판매상태", "현재고수량"]
        ]
        .rename(columns={"_상품코드": "상품코드", "상품명_Master": "상품명", "현재고수량": "현재고"})
    )

    section_title("마이너스 재고 SKU", accent)
    if negative.empty:
        st.success("마이너스 재고 SKU가 없습니다.")
    else:
        st.dataframe(
            style_negative(negative, "현재고").format({"현재고": num}),
            use_container_width=True,
            hide_index=True,
        )


# =========================================================
# 4. 입출고 추이
# =========================================================
elif page == "입출고 추이":
    section_title("입출고 추이", accent)
    i = filter_df(inbound, selected_factory, selected_category)
    o = filter_df(outbound, selected_factory, selected_category)

    mi = monthly(i, "입고일", "입고수량", "입고")
    mo = monthly(o, "출고일", "출고수량", "출고")
    trend = mi.merge(mo, on="월", how="outer").fillna(0).sort_values("월")

    if trend.empty:
        st.info("조회할 데이터가 없습니다.")
    else:
        fig = px.line(
            trend, x="월", y=["입고", "출고"], markers=True, title="월별 입고 / 출고",
            color_discrete_map={"입고": "#22C55E", "출고": "#6366F1"},
        )
        fig.update_layout(
            height=450, margin=dict(l=20, r=20, t=60, b=20),
            xaxis_title="월", yaxis_title="수량", legend_title_text="",
        )
        st.plotly_chart(fig, use_container_width=True, theme="streamlit")

        a, b = st.columns(2)
        with a:
            chart_bar(trend, "월", "입고", "월별 입고")
        with b:
            chart_bar(trend, "월", "출고", "월별 출고")

        show = trend.copy()
        st.dataframe(
            show.style.format({"입고": num, "출고": num}),
            use_container_width=True,
            hide_index=True,
        )

    channel = (
        o.groupby("판매채널", as_index=False)["출고수량"]
        .sum().sort_values("출고수량", ascending=False)
    )
    chart_bar(channel, "판매채널", "출고수량", "판매채널별 출고")


# =========================================================
# 5. 품질 현황
# =========================================================
elif page == "품질 현황":
    section_title("품질 현황", accent)
    d = filter_df(defect, selected_factory, selected_category)
    o = filter_df(outbound, selected_factory, selected_category)

    defect_qty = d["불량수량"].sum()
    out_qty = o["출고수량"].sum()
    rate = defect_qty / out_qty * 100 if out_qty else 0

    cols = st.columns(3)
    metric_card(cols[0], "불량 수량", num(defect_qty))
    metric_card(cols[1], "출고 수량", num(out_qty))
    metric_card(cols[2], "불량률", pct(rate))

    md = monthly(d, "발생일", "불량수량", "불량")
    if not md.empty:
        fig = px.line(
            md, x="월", y="불량", markers=True, title="월별 불량 추이",
            color_discrete_sequence=["#F43F5E"],
        )
        fig.update_layout(height=430, margin=dict(l=20, r=20, t=60, b=20),
                          xaxis_title="월", yaxis_title="불량 수량")
        st.plotly_chart(fig, use_container_width=True, theme="streamlit")

    fd = d.groupby("공장", as_index=False)["불량수량"].sum().rename(columns={"불량수량": "불량"})
    chart_bar(fd, "공장", "불량", "공장별 불량")

    # 불량유형: '테' -> 테불량 / '렌즈' -> 렌즈불량 / '전체' -> 전체불량 / 그외 -> 분류어려움
    dt = (
        d.groupby("불량유형_그룹", as_index=False)["불량수량"]
        .sum()
        .rename(columns={"불량수량": "불량"})
    )
    dt["불량유형_그룹"] = pd.Categorical(dt["불량유형_그룹"], categories=DEFECT_TYPES, ordered=True)
    dt = dt.sort_values("불량유형_그룹").reset_index(drop=True)
    dt["불량유형_그룹"] = dt["불량유형_그룹"].astype(str)

    chart_bar(dt, "불량유형_그룹", "불량", "불량유형별 불량 현황")

    detail = dt.rename(columns={"불량유형_그룹": "불량유형"})
    detail["비중"] = (detail["불량"] / detail["불량"].sum() * 100) if detail["불량"].sum() else 0

    section_title("불량유형별 상세", accent)
    st.dataframe(
        detail.style.format({"불량": num, "비중": pct}),
        use_container_width=True,
        hide_index=True,
    )


# =========================================================
# 6. 데이터 품질
# =========================================================
elif page == "데이터 품질":
    section_title("데이터 품질 점검", accent)

    master_codes = set(master["_상품코드"])
    rows = []
    for name, df in [
        ("현재고", current), ("입고", inbound), ("출고", outbound),
        ("불량관리", defect), ("재고스냅샷", snapshot)
    ]:
        rows.append({
            "데이터": name,
            "행 수": len(df),
            "상품코드 수": df["_상품코드"].nunique(),
            "Master 미매칭 행": int((~df["_상품코드"].isin(master_codes)).sum()),
        })

    section_title("Master 상품코드 매칭", accent)
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    fc = (
        master.groupby("공장")["_상품코드"].nunique()
        .reindex(FACTORIES, fill_value=0)
        .reset_index(name="SKU")
    )
    chart_bar(fc, "공장", "SKU", "공급처상품명 기준 공장 분류")

    unknown = master[master["공장"] == "미상"][
        ["_상품코드", "상품명", "공급처상품명", "카테고리", "판매상태"]
    ].rename(columns={"_상품코드": "상품코드"})
    # 공급처상품명 열에 숫자(0)와 문자열이 섞여 있어 표시 시 타입 오류가 날 수 있어 문자열로 통일
    unknown["공급처상품명"] = unknown["공급처상품명"].fillna("").astype(str)

    section_title("미상 공장 SKU", accent)
    if unknown.empty:
        st.success("미상으로 분류된 SKU가 없습니다.")
    else:
        st.dataframe(unknown, use_container_width=True, hide_index=True)

    date_rows = [
        ["현재고", data["latest_current"]],
        ["입고", inbound["입고일"].max()],
        ["출고", outbound["출고일"].max()],
        ["불량", defect["발생일"].max()],
    ]
    dates = pd.DataFrame(date_rows, columns=["데이터", "최신일"])
    dates["최신일"] = dates["최신일"].apply(
        lambda x: x.strftime("%Y-%m-%d") if pd.notna(x) else "-"
    )

    section_title("데이터 기준일", accent)
    st.dataframe(dates, use_container_width=True, hide_index=True)
