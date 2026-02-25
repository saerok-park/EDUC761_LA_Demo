import os
import glob
import pandas as pd
import streamlit as st

DATA_DIR = "DATA"
REQUIRED_COLS = [
    "TimeStamp", "Turn", "Speaker", "Sentence",
    "Teacher_Tag", "Student_Tag", "DialogAct"
]


# -----------------------
# Utility
# -----------------------

def get_role(speaker: str) -> str:
    s = (speaker or "").strip().lower()
    if s in {"t", "teacher", "instructor"}:
        return "teacher"
    if "teacher" in s:
        return "teacher"
    return "student"


@st.cache_data(show_spinner=False)
def load_and_merge_xlsx(data_dir: str) -> pd.DataFrame:
    xlsx_files = [
        f for f in glob.glob(os.path.join(data_dir, "*.xlsx"))
        if not os.path.basename(f).startswith("~$")
    ]

    dfs = []
    for path in xlsx_files:
        df = pd.read_excel(path, engine="openpyxl")
        lesson_id = os.path.splitext(os.path.basename(path))[0]
        df["lesson_id"] = lesson_id
        dfs.append(df[REQUIRED_COLS + ["lesson_id"]])

    data = pd.concat(dfs, ignore_index=True)

    for c in ["Speaker", "Sentence", "Teacher_Tag", "Student_Tag", "DialogAct"]:
        data[c] = data[c].astype("string").str.strip()

    data["Turn_num"] = pd.to_numeric(data["Turn"], errors="coerce")
    data = data.sort_values(by=["lesson_id", "Turn_num"]).reset_index(drop=True)
    data["role"] = data["Speaker"].apply(get_role)

    return data


# -----------------------
# Colorful Bar Chart
# -----------------------

def colorful_bar_chart(df, category_col, value_col, title):

    spec = {
        "title": title,
        "data": {"values": df.to_dict(orient="records")},
        "mark": {
            "type": "bar",
            "cornerRadiusTopLeft": 6,
            "cornerRadiusTopRight": 6
        },
        "encoding": {
            "x": {
                "field": category_col,
                "type": "nominal",
                "axis": {"labelAngle": -45}
            },
            "y": {"field": value_col, "type": "quantitative"},
            "color": {
                "field": category_col,
                "type": "nominal",
                "scale": {
                    "range": [
                        "#4E79A7", "#F28E2B", "#E15759", "#76B7B2",
                        "#59A14F", "#EDC948", "#B07AA1", "#FF9DA7",
                        "#9C755F", "#BAB0AC"
                    ]
                }
            },
            "tooltip": [
                {"field": category_col},
                {"field": value_col}
            ],
        },
        "config": {"view": {"stroke": "transparent"}}
    }

    st.vega_lite_chart(spec, use_container_width=True)


# -----------------------
# Heatmap
# -----------------------

def vega_heatmap(df_props: pd.DataFrame, title: str):

    long = (
        df_props.reset_index()
        .melt(id_vars="Teacher_Tag",
              var_name="DialogAct",
              value_name="Proportion")
        .dropna()
    )

    spec = {
        "title": title,
        "data": {"values": long.to_dict(orient="records")},
        "mark": "rect",
        "encoding": {
            "x": {
                "field": "DialogAct",
                "type": "nominal",
                "axis": {"labelAngle": -45}
            },
            "y": {
                "field": "Teacher_Tag",
                "type": "nominal"
            },
            "color": {
                "field": "Proportion",
                "type": "quantitative",
                "scale": {
                    "range": ["#f7fbff", "#08306b"]
                }
            },
            "tooltip": [
                {"field": "Teacher_Tag"},
                {"field": "DialogAct"},
                {"field": "Proportion", "format": ".0%"}
            ],
        },
        "config": {"view": {"stroke": "transparent"}}
    }

    st.vega_lite_chart(spec, use_container_width=True)


# =========================
# UI
# =========================

st.set_page_config(
    page_title="Classroom Discourse Dashboard",
    layout="wide"
)

st.title("📊 Classroom Discourse Dashboard")

data = load_and_merge_xlsx(DATA_DIR)

# Sidebar
st.sidebar.header("Filters")

lessons = sorted(data["lesson_id"].unique())
selected_lessons = st.sidebar.multiselect(
    "Select lessons",
    lessons,
    default=lessons
)

filtered = data[data["lesson_id"].isin(selected_lessons)].copy()


# -----------------------
# Data Preview
# -----------------------

with st.expander("📄 View Raw Data (First 50 Rows)"):
    st.dataframe(
        filtered[
            [
                "lesson_id", "TimeStamp", "Turn",
                "Speaker", "role",
                "Sentence", "Teacher_Tag",
                "Student_Tag", "DialogAct"
            ]
        ].head(50),
        use_container_width=True
    )


# =========================
# RQ1
# =========================

st.header("RQ1. Classroom discourse distribution & dialog acts")

col1, col2 = st.columns(2)

with col1:
    rq1_turns = (
        filtered.groupby("role")
        .size()
        .reset_index(name="turn_count")
    )

    colorful_bar_chart(
        rq1_turns,
        "role",
        "turn_count",
        "Teacher vs Student Turn Frequency"
    )

with col2:
    da_counts = (
        filtered["DialogAct"]
        .dropna()
        .value_counts()
        .head(10)
        .reset_index()
    )

    da_counts.columns = ["DialogAct", "count"]

    colorful_bar_chart(
        da_counts,
        "DialogAct",
        "count",
        "DialogAct Distribution"
    )


# =========================
# RQ2
# =========================

st.header("RQ2. Patterns in students’ discourse contributions")

students = filtered[filtered["role"] == "student"]

st_counts = (
    students["Student_Tag"]
    .dropna()
    .value_counts()
    .head(10)
    .reset_index()
)

st_counts.columns = ["Student_Tag", "count"]

colorful_bar_chart(
    st_counts,
    "Student_Tag",
    "count",
    "Student Tag Distribution"
)


# =========================
# RQ3
# =========================

st.header("RQ3. Teacher instructional intentions × DialogAct")

teachers = filtered[filtered["role"] == "teacher"]

ct_counts = pd.crosstab(
    teachers["Teacher_Tag"],
    teachers["DialogAct"]
)

ct_props = ct_counts.div(ct_counts.sum(axis=1), axis=0).fillna(0)
ct_props.index.name = "Teacher_Tag"

vega_heatmap(
    ct_props,
    "Teacher_Tag × DialogAct (Row Proportions)"
)