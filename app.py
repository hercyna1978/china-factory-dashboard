import os
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st


# =========================================================
# 페이지 설정
# =========================================================
st.set_page_config(
    page_title="중국공장 물류 운영 대시보드",
    layout="wide",
    initial_sidebar_state="expanded",
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

CHART_COLORS = [
    "#2563EB",
    "#06B6D4",
    "#10B981",
    "#F59E0B",
    "#8B5CF6",
    "#EF4444",
    "#EC4899",
    "#64748B",
]

try:
    THEME_TYPE = st.context.theme.type
except Exception:
    THEME_TYPE = "light"

PLOTLY_TEMPLATE = "plotly_dark" if THEME_TYPE == "dark" else "plotly_white"


# =========================================================
# 화면 디자인
# =========================================================
st.markdown(
    """
<style>
.block-container {
    padding-top: 1.1rem;
    padding-bottom: 2rem;
    max-width: 1600px;
}

[data-testid="stSidebar"] {
    border-right: 1px solid var(--secondary-background-color);
}

[data-testid="stSidebar"] > div:first-child {
    padding-top: 1rem;
}

.dashboard-header {
    padding: 22px 26px;
    margin-bottom: 18px;
    border-radius: 18px;
    border: 1px solid var(--secondary-background-color);
    background:
        linear-gradient(
            135deg,
            rgba(37, 99, 235, 0.16),
            rgba(6, 182, 212, 0.08)
        ),
        var(--background-color);
    box-shadow: 0 6px 24px rgba(0, 0, 0, 0.08);
}

.dashboard-title {
    font-size: 2rem;
    line-height: 1.15;
    font-weight: 850;
    letter-spacing: -0.04em;
    margin: 0;
}

.dashboard-subtitle {
    margin-top: 7px;
    font-size: 0.93rem;
    opacity: 0.68;
}

.section-title {
    display: flex;
    align-items: center;
    gap: 9px;
    margin: 22px 0 10px 0;
    font-size: 1.12rem;
    font-weight: 800;
    letter-spacing: -0.02em;
}

.section-title::before {
    content: "";
    display: inline-block;
    width: 5px;
    height: 21px;
    border-radius: 5px;
    background: linear-gradient(180deg, #2563EB, #06B6D4);
}

.kpi-grid {
    display: grid;
    grid-template-columns: repeat(6, minmax(0, 1fr));
    gap: 12px;
    margin: 10px 0 16px 0;
}

.kpi-card {
    position: relative;
    overflow: hidden;
    min-height: 108px;
    padding: 16px 17px;
    border-radius: 16px;
    border: 1px solid var(--secondary-background-color);
    background: var(--secondary-background-color);
    box-shadow: 0 4px 16px rgba(0, 0, 0, 0.06);
}

.kpi-card::after {
    content: "";
    position: absolute;
    left: 0;
    right: 0;
    bottom: 0;
    height: 4px;
    background: var(--kpi-color);
}

.kpi-label {
    font-size: 0.82rem;
    font-weight: 700;
    opacity: 0.68;
}

.kpi-value {
    margin-top: 9px;
    font-size: 1.65rem;
    line-height: 1.1;
    font-weight: 850;
    letter-spacing: -0.035em;
}

.info-card {
    border: 1px solid var(--secondary-background-color);
    border-radius: 14px;
    padding: 14px 16px;
    background: var(--secondary-background-color);
    margin-bottom: 14px;
}

.danger-card {
    border: 1px solid rgba(239, 68, 68, 0.35);
    border-radius: 14px;
    padding: 14px 16px;
    background: rgba(239, 68, 68, 0.08);
    margin-bottom: 14px;
}

div[data-testid="stDataFrame"] {
    border-radius: 12px;
    overflow: hidden;
}

[data-testid="stMetric"] {
    border-radius: 14px;
    border: 1px solid var(--secondary-background-color);
    background: var(--secondary-background-color);
    padding: 13px 15px;
}

@media (max-width: 1100px) {
    .kpi-grid {
        grid-template-columns: repeat(3, minmax(0, 1fr));
    }
}

@media (max-width: 700px) {
    .kpi-grid {
        grid-template-columns: repeat(2, minmax(0, 1fr));
    }
}
</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# 기본 함수
# =========================================================
def code(value):
    if pd.isna(value):
        return ""
    value = str(value).strip().upper()
    if value.endswith(".0"):
        value = value[:-2]
    return value


def classify_factory(value):
    """
    상품Master의 공급처상품명을 기준으로 공장 분류.
    C2-S가 C2보다 먼저 검사되어야 한다.
    """
    if pd.isna(value):
        return "미상"

    value = str(value).strip().upper().replace(" ", "")

    if not value or value in {"NAN", "NONE", "0"}:
        return "미상"

    if "C2-S" in value:
        return "C2-S"
    if "C2" in value:
        return "C2"
    if "C5" in value:
        return "C5"

    return "미상"


def classify_defect(value):
    """
    불량유형 통합 분류 기준
    테라 포함 -> 테불량
    렌즈 포함 -> 렌즈불량
    전체 포함 -> 전체 불량
    그 외 -> 분류어려움
    """
    if pd.isna(value):
        return "분류어려움"

    value = str(value).strip()

    if "테라" in value:
        return "테불량"
    if "렌즈" in value:
        return "렌즈불량"
    if "전체" in value:
        return "전체 불량"

    return "분류어려움"


def num(value):
    return f"{float(value):,.0f}"


def pct(value):
    return "-" if pd.isna(value) else f"{float(value):.1f}%"


def filter_df(df, factory, category):
    result = df

    if factory != "전체":
        result = result[result["공장"] == factory]

    if category != "전체":
        result = result[result["카테고리"] == category]

    return result


def section(title):
    st.markdown(
        f'<div class="section-title">{title}</div>',
        unsafe_allow_html=True,
    )


def kpi_cards(items):
    cards = []

    for label, value, color in items:
        cards.append(
            f"""
            <div class="kpi-card" style="--kpi-color:{color};">
                <div class="kpi-label">{label}</div>
                <div class="kpi-value">{value}</div>
            </div>
            """
        )

    st.markdown(
        '<div class="kpi-grid">' + "".join(cards) + "</div>",
        unsafe_allow_html=True,
    )


def chart_bar(df, x, y, title, horizontal=False):
    if df.empty:
        st.info("조회할 데이터가 없습니다.")
        return

    work = df.copy()

    if horizontal:
        work = work.sort_values(y)

        fig = px.bar(
            work,
            x=y,
            y=x,
            orientation="h",
            title=title,
            text=y,
            color_discrete_sequence=CHART_COLORS,
        )
    else:
        fig = px.bar(
            work,
            x=x,
            y=y,
            title=title,
            text=y,
            color_discrete_sequence=CHART_COLORS,
        )

    fig.update_traces(
        texttemplate="%{text:,.0f}",
        textposition="outside",
        cliponaxis=False,
    )

    fig.update_layout(
        template=PLOTLY_TEMPLATE,
        height=390,
        margin=dict(l=25, r=25, t=60, b=30),
        xaxis_title=None,
        yaxis_title=None,
        showlegend=False,
        hovermode="x unified",
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"displayModeBar": False},
    )


def add_master_info(df, maps):
    result = df.copy()

    result["_상품코드"] = result["상품코드"].map(code)
    result["공장"] = (
        result["_상품코드"].map(maps["factory"]).fillna("미상")
    )
    result["상품명_Master"] = (
        result["_상품코드"].map(maps["name"]).fillna(result["상품명"])
    )
    result["카테고리"] = (
        result["_상품코드"].map(maps["category"]).fillna("미상")
    )
    result["판매상태"] = (
        result["_상품코드"].map(maps["status"]).fillna("미상")
    )

    return result


def monthly(df, date_col, qty_col, result_name):
    result = df.dropna(subset=[date_col]).copy()

    if result.empty:
        return pd.DataFrame(columns=["월", result_name])

    result["월"] = result[date_col].dt.to_period("M").astype(str)

    return (
        result.groupby("월", as_index=False)[qty_col]
        .sum()
        .rename(columns={qty_col: result_name})
        .sort_values("월")
    )


def factory_kpi(data):
    rows = []

    for factory_name in FACTORIES:
        master_df = data["master"][
            data["master"]["공장"] == factory_name
        ]
        current_df = data["current"][
            data["current"]["공장"] == factory_name
        ]
        inbound_df = data["inbound"][
            data["inbound"]["공장"] == factory_name
        ]
        outbound_df = data["outbound"][
            data["outbound"]["공장"] == factory_name
        ]
        defect_df = data["defect"][
            data["defect"]["공장"] == factory_name
        ]

        outbound_qty = outbound_df["출고수량"].sum()
        defect_qty = defect_df["불량수량"].sum()

        rows.append(
            {
                "공장": factory_name,
                "SKU": master_df["_상품코드"].nunique(),
                "현재고": current_df["현재고수량"].sum(),
                "입고": inbound_df["입고수량"].sum(),
                "출고": outbound_qty,
                "불량": defect_qty,
                "불량률": (
                    defect_qty / outbound_qty * 100
                    if outbound_qty
                    else 0
                ),
            }
        )

    return pd.DataFrame(rows)


def styled_numbers(df, integer_columns=None, percent_columns=None):
    result = df.copy()

    integer_columns = integer_columns or []
    percent_columns = percent_columns or []

    for column in integer_columns:
        if column in result.columns:
            result[column] = pd.to_numeric(
                result[column], errors="coerce"
            ).map(
                lambda x: f"{x:,.0f}" if pd.notna(x) else "-"
            )

    for column in percent_columns:
        if column in result.columns:
            result[column] = pd.to_numeric(
                result[column], errors="coerce"
            ).map(
                lambda x: f"{x:.1f}%" if pd.notna(x) else "-"
            )

    return result


# =========================================================
# 엑셀 읽기
# =========================================================
def read_excel(source):
    excel = pd.ExcelFile(source)

    missing_sheets = [
        sheet for sheet in SHEETS.values()
        if sheet not in excel.sheet_names
    ]

    if missing_sheets:
        raise ValueError(
            "필수 시트가 없습니다: " + ", ".join(missing_sheets)
        )

    master = pd.read_excel(excel, SHEETS["master"])
    snapshot = pd.read_excel(excel, SHEETS["snapshot"])
    inbound = pd.read_excel(excel, SHEETS["inbound"])
    outbound = pd.read_excel(excel, SHEETS["outbound"])
    defect = pd.read_excel(excel, SHEETS["defect"])
    current = pd.read_excel(excel, SHEETS["current"])

    required_columns = {
        "상품Master": [
            "상품코드", "상품명", "카테고리",
            "출시일", "판매상태", "공급처상품명"
        ],
        "재고스냅샷": [
            "기준일", "상품코드", "상품명", "현재고수량"
        ],
        "입고": [
            "입고일", "상품코드", "상품명",
            "입고수량", "중국공장"
        ],
        "출고": [
            "출고일", "상품코드", "상품명",
            "출고수량", "판매채널"
        ],
        "불량관리": [
            "발생일", "상품코드", "상품명",
            "불량수량", "불량유형", "중국공장"
        ],
        "현재고": [
            "기준일", "상품코드", "상품명", "현재고수량"
        ],
    }

    loaded = {
        "상품Master": master,
        "재고스냅샷": snapshot,
        "입고": inbound,
        "출고": outbound,
        "불량관리": defect,
        "현재고": current,
    }

    for sheet_name, columns in required_columns.items():
        missing_columns = [
            column
            for column in columns
            if column not in loaded[sheet_name].columns
        ]

        if missing_columns:
            raise ValueError(
                f"{sheet_name} 시트에 열이 없습니다: "
                + ", ".join(missing_columns)
            )

    date_columns = [
        (master, "출시일"),
        (snapshot, "기준일"),
        (inbound, "입고일"),
        (outbound, "출고일"),
        (defect, "발생일"),
        (current, "기준일"),
    ]

    for df, column in date_columns:
        df[column] = pd.to_datetime(
            df[column],
            errors="coerce",
        )

    numeric_columns = [
        (snapshot, "현재고수량"),
        (inbound, "입고수량"),
        (outbound, "출고수량"),
        (defect, "불량수량"),
        (current, "현재고수량"),
    ]

    for df, column in numeric_columns:
        df[column] = pd.to_numeric(
            df[column],
            errors="coerce",
        ).fillna(0)

    # -----------------------------------------------------
    # 상품Master 기준 공장 분류
    # -----------------------------------------------------
    master["_상품코드"] = master["상품코드"].map(code)
    master["공장"] = master["공급처상품명"].apply(
        classify_factory
    )

    master_unique = master.drop_duplicates(
        "_상품코드",
        keep="first",
    )

    maps = {
        "factory": master_unique.set_index(
            "_상품코드"
        )["공장"].to_dict(),

        "name": master_unique.set_index(
            "_상품코드"
        )["상품명"].to_dict(),

        "category": master_unique.set_index(
            "_상품코드"
        )["카테고리"].to_dict(),

        "status": master_unique.set_index(
            "_상품코드"
        )["판매상태"].to_dict(),
    }

    snapshot = add_master_info(snapshot, maps)
    inbound = add_master_info(inbound, maps)
    outbound = add_master_info(outbound, maps)
    defect = add_master_info(defect, maps)
    current = add_master_info(current, maps)

    # -----------------------------------------------------
    # 불량유형 통합 분류
    # -----------------------------------------------------
    defect["불량유형_통합"] = defect["불량유형"].apply(
        classify_defect
    )

    # -----------------------------------------------------
    # 현재고는 최신 기준일 데이터만 사용
    # -----------------------------------------------------
    if current["기준일"].notna().any():
        latest_current = current["기준일"].max()
        current = current[
            current["기준일"] == latest_current
        ].copy()
    else:
        latest_current = pd.NaT

    return {
        "master": master,
        "snapshot": snapshot,
        "inbound": inbound,
        "outbound": outbound,
        "defect": defect,
        "current": current,
        "latest_current": latest_current,
    }


# =========================================================
# 데이터 준비
# =========================================================
with st.sidebar:
    st.markdown("## 데이터")
    uploaded_file = st.file_uploader(
        "Excel 파일 선택",
        type=["xlsx"],
    )

if uploaded_file is not None:
    source = uploaded_file
    source_name = uploaded_file.name
else:
    if not os.path.exists(FILE_NAME):
        st.error(
            f"{FILE_NAME} 파일이 app.py와 같은 폴더에 없습니다."
        )
        st.stop()

    source = FILE_NAME
    source_name = FILE_NAME

try:
    data = read_excel(source)
except Exception as error:
    st.error("데이터를 읽는 중 오류가 발생했습니다.")
    st.code(str(error))
    st.stop()

master = data["master"]
snapshot = data["snapshot"]
inbound = data["inbound"]
outbound = data["outbound"]
defect = data["defect"]
current = data["current"]


# =========================================================
# 사이드바 필터
# =========================================================
with st.sidebar:
    st.markdown("---")
    st.markdown("## 조회")

    page = st.radio(
        "화면",
        [
            "경영진 요약",
            "공장별 운영",
            "재고 현황",
            "입출고 추이",
            "품질 현황",
            "데이터 품질",
        ],
        label_visibility="collapsed",
    )

    selected_factory = st.selectbox(
        "공장",
        ["전체"] + FACTORIES,
    )

    category_values = sorted(
        master["카테고리"]
        .dropna()
        .astype(str)
        .unique()
        .tolist()
    )

    selected_category = st.selectbox(
        "카테고리",
        ["전체"] + category_values,
    )

    st.markdown("---")
    st.markdown("## 데이터 기준")

    if pd.notna(data["latest_current"]):
        st.caption(
            "현재고 기준일: "
            + data["latest_current"].strftime("%Y-%m-%d")
        )

    transaction_dates = pd.concat(
        [
            inbound["입고일"],
            outbound["출고일"],
            defect["발생일"],
        ],
        ignore_index=True,
    ).dropna()

    if not transaction_dates.empty:
        st.caption(
            "거래기간: "
            + transaction_dates.min().strftime("%Y-%m-%d")
            + " ~ "
            + transaction_dates.max().strftime("%Y-%m-%d")
        )

    st.caption("파일: " + source_name)


# =========================================================
# 공통 헤더
# =========================================================
st.markdown(
    """
    <div class="dashboard-header">
        <div class="dashboard-title">중국공장 물류 운영 대시보드</div>
        <div class="dashboard-subtitle">
            상품Master 기준 공장 분류 · 재고 · 입출고 · 품질 통합 현황
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)


# =========================================================
# 1. 경영진 요약
# =========================================================
if page == "경영진 요약":
    master_f = filter_df(
        master,
        selected_factory,
        selected_category,
    )
    current_f = filter_df(
        current,
        selected_factory,
        selected_category,
    )
    inbound_f = filter_df(
        inbound,
        selected_factory,
        selected_category,
    )
    outbound_f = filter_df(
        outbound,
        selected_factory,
        selected_category,
    )
    defect_f = filter_df(
        defect,
        selected_factory,
        selected_category,
    )

    sku = master_f["_상품코드"].nunique()
    stock = current_f["현재고수량"].sum()
    inbound_qty = inbound_f["입고수량"].sum()
    outbound_qty = outbound_f["출고수량"].sum()
    defect_qty = defect_f["불량수량"].sum()

    defect_rate = (
        defect_qty / outbound_qty * 100
        if outbound_qty
        else 0
    )

    kpi_cards(
        [
            ("SKU", num(sku), "#2563EB"),
            ("현재고", num(stock), "#06B6D4"),
            ("누적 입고", num(inbound_qty), "#10B981"),
            ("누적 출고", num(outbound_qty), "#F59E0B"),
            ("불량 수량", num(defect_qty), "#EF4444"),
            ("불량률", pct(defect_rate), "#8B5CF6"),
        ]
    )

    section("공장별 핵심 현황")

    kpi = factory_kpi(data)

    if selected_factory != "전체":
        kpi = kpi[
            kpi["공장"] == selected_factory
        ]

    left, right = st.columns(2)

    with left:
        chart_bar(
            kpi,
            "공장",
            "현재고",
            "공장별 현재고",
        )

    with right:
        chart_bar(
            kpi,
            "공장",
            "출고",
            "공장별 출고",
        )

    section("공장별 KPI")

    display = styled_numbers(
        kpi,
        integer_columns=[
            "SKU",
            "현재고",
            "입고",
            "출고",
            "불량",
        ],
        percent_columns=["불량률"],
    )

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
    )


# =========================================================
# 2. 공장별 운영
# =========================================================
elif page == "공장별 운영":
    section("공장별 운영 실적")

    kpi = factory_kpi(data)

    if selected_factory != "전체":
        kpi = kpi[
            kpi["공장"] == selected_factory
        ]

    inbound_f = filter_df(
        inbound,
        selected_factory,
        selected_category,
    )
    outbound_f = filter_df(
        outbound,
        selected_factory,
        selected_category,
    )
    current_f = filter_df(
        current,
        selected_factory,
        selected_category,
    )
    defect_f = filter_df(
        defect,
        selected_factory,
        selected_category,
    )

    kpi_cards(
        [
            ("입고", num(inbound_f["입고수량"].sum()), "#10B981"),
            ("출고", num(outbound_f["출고수량"].sum()), "#F59E0B"),
            ("현재고", num(current_f["현재고수량"].sum()), "#06B6D4"),
            ("불량", num(defect_f["불량수량"].sum()), "#EF4444"),
        ]
    )

    left, right = st.columns(2)

    with left:
        chart_bar(
            kpi,
            "공장",
            "입고",
            "공장별 입고",
        )

    with right:
        chart_bar(
            kpi,
            "공장",
            "출고",
            "공장별 출고",
        )

    left, right = st.columns(2)

    with left:
        chart_bar(
            kpi,
            "공장",
            "불량",
            "공장별 불량",
        )

    with right:
        sku_by_factory = (
            master.groupby("공장")["_상품코드"]
            .nunique()
            .reindex(FACTORIES, fill_value=0)
            .reset_index(name="SKU")
        )

        if selected_factory != "전체":
            sku_by_factory = sku_by_factory[
                sku_by_factory["공장"] == selected_factory
            ]

        chart_bar(
            sku_by_factory,
            "공장",
            "SKU",
            "공장별 SKU",
        )

    section("공장별 KPI 상세")

    display = styled_numbers(
        kpi,
        integer_columns=[
            "SKU",
            "현재고",
            "입고",
            "출고",
            "불량",
        ],
        percent_columns=["불량률"],
    )

    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
    )


# =========================================================
# 3. 재고 현황
# =========================================================
elif page == "재고 현황":
    section("현재 재고 현황")

    current_f = filter_df(
        current,
        selected_factory,
        selected_category,
    )

    total_stock = current_f["현재고수량"].sum()
    positive_stock = current_f.loc[
        current_f["현재고수량"] > 0,
        "현재고수량",
    ].sum()
    zero_sku = (
        current_f["현재고수량"] == 0
    ).sum()
    negative_sku = (
        current_f["현재고수량"] < 0
    ).sum()

    kpi_cards(
        [
            ("현재고", num(total_stock), "#2563EB"),
            ("재고 보유 수량", num(positive_stock), "#10B981"),
            ("재고 0 SKU", num(zero_sku), "#F59E0B"),
            ("마이너스 재고 SKU", num(negative_sku), "#EF4444"),
        ]
    )

    stock_by_factory = (
        current_f.groupby(
            "공장",
            as_index=False,
        )["현재고수량"]
        .sum()
        .rename(
            columns={"현재고수량": "현재고"}
        )
    )

    chart_bar(
        stock_by_factory,
        "공장",
        "현재고",
        "공장별 현재고",
    )

    section("현재고 상위 20 SKU")

    top = (
        current_f.groupby(
            [
                "_상품코드",
                "상품명_Master",
                "공장",
                "카테고리",
            ],
            as_index=False,
        )["현재고수량"]
        .sum()
        .sort_values(
            "현재고수량",
            ascending=False,
        )
        .head(20)
        .rename(
            columns={
                "_상품코드": "상품코드",
                "상품명_Master": "상품명",
                "현재고수량": "현재고",
            }
        )
    )

    top["현재고"] = top["현재고"].map(num)

    st.dataframe(
        top,
        use_container_width=True,
        hide_index=True,
    )

    section("마이너스 재고 SKU")

    # 중요: 여러 컬럼 선택은 반드시 이중 대괄호를 사용한다.
    # 기존 오류의 원인이었던 df["a", "b"] 형태를 제거했다.
    negative = (
        current_f[
            current_f["현재고수량"] < 0
        ]
        .sort_values("현재고수량")
        [
            [
                "_상품코드",
                "상품명_Master",
                "공장",
                "카테고리",
                "판매상태",
                "현재고수량",
            ]
        ]
        .rename(
            columns={
                "_상품코드": "상품코드",
                "상품명_Master": "상품명",
                "현재고수량": "현재고",
            }
        )
    )

    if negative.empty:
        st.markdown(
            '<div class="info-card">마이너스 재고 SKU가 없습니다.</div>',
            unsafe_allow_html=True,
        )
    else:
        negative["현재고"] = negative["현재고"].map(num)

        st.markdown(
            f"""
            <div class="danger-card">
                마이너스 재고 SKU <strong>{len(negative):,}</strong>건
            </div>
            """,
            unsafe_allow_html=True,
        )

        st.dataframe(
            negative,
            use_container_width=True,
            hide_index=True,
        )


# =========================================================
# 4. 입출고 추이
# =========================================================
elif page == "입출고 추이":
    section("입출고 추이")

    inbound_f = filter_df(
        inbound,
        selected_factory,
        selected_category,
    )
    outbound_f = filter_df(
        outbound,
        selected_factory,
        selected_category,
    )

    inbound_qty = inbound_f["입고수량"].sum()
    outbound_qty = outbound_f["출고수량"].sum()

    kpi_cards(
        [
            ("누적 입고", num(inbound_qty), "#10B981"),
            ("누적 출고", num(outbound_qty), "#F59E0B"),
            (
                "입출고 차이",
                num(inbound_qty - outbound_qty),
                "#2563EB",
            ),
        ]
    )

    monthly_inbound = monthly(
        inbound_f,
        "입고일",
        "입고수량",
        "입고",
    )

    monthly_outbound = monthly(
        outbound_f,
        "출고일",
        "출고수량",
        "출고",
    )

    trend = (
        monthly_inbound
        .merge(
            monthly_outbound,
            on="월",
            how="outer",
        )
        .fillna(0)
        .sort_values("월")
    )

    if trend.empty:
        st.info("조회할 데이터가 없습니다.")
    else:
        fig = px.line(
            trend,
            x="월",
            y=["입고", "출고"],
            markers=True,
            title="월별 입고 / 출고",
            color_discrete_sequence=CHART_COLORS[:2],
        )

        fig.update_layout(
            template=PLOTLY_TEMPLATE,
            height=450,
            margin=dict(l=25, r=25, t=60, b=30),
            xaxis_title="월",
            yaxis_title="수량",
            hovermode="x unified",
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
            config={"displayModeBar": False},
        )

        left, right = st.columns(2)

        with left:
            chart_bar(
                trend,
                "월",
                "입고",
                "월별 입고",
            )

        with right:
            chart_bar(
                trend,
                "월",
                "출고",
                "월별 출고",
            )

        display = trend.copy()
        display["입고"] = display["입고"].map(num)
        display["출고"] = display["출고"].map(num)

        st.dataframe(
            display,
            use_container_width=True,
            hide_index=True,
        )

    section("판매채널별 출고")

    channel = (
        outbound_f.groupby(
            "판매채널",
            as_index=False,
        )["출고수량"]
        .sum()
        .sort_values(
            "출고수량",
            ascending=False,
        )
    )

    chart_bar(
        channel,
        "판매채널",
        "출고수량",
        "판매채널별 출고",
    )


# =========================================================
# 5. 품질 현황
# =========================================================
elif page == "품질 현황":
    section("품질 현황")

    defect_f = filter_df(
        defect,
        selected_factory,
        selected_category,
    )
    outbound_f = filter_df(
        outbound,
        selected_factory,
        selected_category,
    )

    defect_qty = defect_f["불량수량"].sum()
    outbound_qty = outbound_f["출고수량"].sum()

    defect_rate = (
        defect_qty / outbound_qty * 100
        if outbound_qty
        else 0
    )

    kpi_cards(
        [
            ("불량 수량", num(defect_qty), "#EF4444"),
            ("출고 수량", num(outbound_qty), "#F59E0B"),
            ("불량률", pct(defect_rate), "#8B5CF6"),
        ]
    )

    monthly_defect = monthly(
        defect_f,
        "발생일",
        "불량수량",
        "불량",
    )

    if not monthly_defect.empty:
        fig = px.line(
            monthly_defect,
            x="월",
            y="불량",
            markers=True,
            title="월별 불량 추이",
            color_discrete_sequence=["#EF4444"],
        )

        fig.update_layout(
            template=PLOTLY_TEMPLATE,
            height=430,
            margin=dict(l=25, r=25, t=60, b=30),
            xaxis_title="월",
            yaxis_title="불량 수량",
            hovermode="x unified",
        )

        st.plotly_chart(
            fig,
            use_container_width=True,
            config={"displayModeBar": False},
        )

    left, right = st.columns(2)

    with left:
        factory_defect = (
            defect_f.groupby(
                "공장",
                as_index=False,
            )["불량수량"]
            .sum()
            .rename(
                columns={"불량수량": "불량"}
            )
        )

        chart_bar(
            factory_defect,
            "공장",
            "불량",
            "공장별 불량",
        )

    with right:
        defect_type = (
            defect_f.groupby(
                "불량유형_통합",
                as_index=False,
            )["불량수량"]
            .sum()
            .rename(
                columns={
                    "불량유형_통합": "불량유형",
                    "불량수량": "불량",
                }
            )
            .sort_values(
                "불량",
                ascending=True,
            )
        )

        chart_bar(
            defect_type,
            "불량유형",
            "불량",
            "불량유형별 불량",
            horizontal=True,
        )

    section("불량유형별 상세")

    detail = (
        defect_f.groupby(
            "불량유형_통합",
            as_index=False,
        )["불량수량"]
        .sum()
        .rename(
            columns={
                "불량유형_통합": "불량유형",
                "불량수량": "불량수량",
            }
        )
        .sort_values(
            "불량수량",
            ascending=False,
        )
    )

    detail["불량수량"] = detail[
        "불량수량"
    ].map(num)

    st.dataframe(
        detail,
        use_container_width=True,
        hide_index=True,
    )


# =========================================================
# 6. 데이터 품질
# =========================================================
elif page == "데이터 품질":
    section("데이터 품질 점검")

    master_codes = set(
        master["_상품코드"]
    )

    rows = []

    for data_name, df in [
        ("현재고", current),
        ("입고", inbound),
        ("출고", outbound),
        ("불량관리", defect),
        ("재고스냅샷", snapshot),
    ]:
        rows.append(
            {
                "데이터": data_name,
                "행 수": len(df),
                "상품코드 수": df[
                    "_상품코드"
                ].nunique(),
                "Master 미매칭 행": int(
                    (
                        ~df["_상품코드"].isin(
                            master_codes
                        )
                    ).sum()
                ),
            }
        )

    st.dataframe(
        pd.DataFrame(rows),
        use_container_width=True,
        hide_index=True,
    )

    section("공장별 SKU 분류")

    factory_count = (
        master.groupby("공장")["_상품코드"]
        .nunique()
        .reindex(
            FACTORIES,
            fill_value=0,
        )
        .reset_index(name="SKU")
    )

    chart_bar(
        factory_count,
        "공장",
        "SKU",
        "공급처상품명 기준 공장 분류",
    )

    section("미상 공장 SKU")

    unknown = master[
        master["공장"] == "미상"
    ][
        [
            "_상품코드",
            "상품명",
            "공급처상품명",
            "카테고리",
            "판매상태",
        ]
    ].rename(
        columns={
            "_상품코드": "상품코드"
        }
    )

    if unknown.empty:
        st.markdown(
            '<div class="info-card">미상으로 분류된 SKU가 없습니다.</div>',
            unsafe_allow_html=True,
        )
    else:
        st.dataframe(
            unknown,
            use_container_width=True,
            hide_index=True,
        )

    section("데이터 최신일")

    latest_dates = pd.DataFrame(
        [
            [
                "현재고",
                data["latest_current"],
            ],
            [
                "입고",
                inbound["입고일"].max(),
            ],
            [
                "출고",
                outbound["출고일"].max(),
            ],
            [
                "불량",
                defect["발생일"].max(),
            ],
        ],
        columns=["데이터", "최신일"],
    )

    latest_dates["최신일"] = latest_dates[
        "최신일"
    ].apply(
        lambda value:
            value.strftime("%Y-%m-%d")
            if pd.notna(value)
            else "-"
    )

    st.dataframe(
        latest_dates,
        use_container_width=True,
        hide_index=True,
    )
