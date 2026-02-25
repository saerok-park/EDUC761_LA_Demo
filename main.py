import os
import glob
import pandas as pd
import streamlit as st

DATA_DIR = "DATA"
REQUIRED_COLS = ["TimeStamp", "Turn", "Speaker", "Sentence", "Teacher_Tag", "Student_Tag", "DialogAct"]


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

    if not xlsx_files:
        raise FileNotFoundError(f"No .xlsx files found in: {data_dir}")

    dfs = []
    for path in xlsx_files:
        df = pd.read_excel(path, engine="openpyxl")

        missing = [c for c in REQUIRED_COLS if c not in df.columns]
        if missing:
            raise ValueError(
                f"File '{os.path.basename(path)}' is missing columns: {missing}"
            )

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


def vega_heatmap(df_props: pd.DataFrame, title: str):
    long = (
        df_props.reset_index()
        .melt(id_vars="Teacher_Tag", var_name="DialogAct", value_name="Proportion")
        .dropna()
    )

    spec = {
        "title": title,
        "data": {"values": long.to_dict(orient="records")},
        "mark": "rect",
        "encoding": {
            "x": {"field": "DialogAct", "type": "nominal", "axis": {"labelAngle": -45}},
            "y": {"field": "Teacher_Tag", "type": "nominal"},
            "color": {"field": "Proportion", "type": "quantitative"},
            "tooltip": [
                {"field": "Teacher_Tag", "type": "nominal"},
                {"field": "DialogAct", "type": "nominal"},
                {"field": "Proportion", "type": "quantitative", "format": ".0%"},
            ],
        },
        "config": {"view": {"stroke": "transparent"}},
    }

    st.vega_lite_chart(spec, use_container_width=True)


# =========================
# UI
# =========================

st.set_page_config(page_title="Classroom Discourse Dashboard", layout="wide")
st.title("📊 Classroom Discourse Dashboard")

data = load_and_merge_xlsx(DATA_DIR)

# Sidebar
st.sidebar.header("Filters")

lessons = sorted(data["lesson_id"].unique())
selected_lessons = st.sidebar.multiselect(
    "Select lessons", lessons, default=lessons
)

role_filter = st.sidebar.selectbox(
    "Role filter",
    ["all", "teacher", "student"]
)

# Apply filters
filtered = data[data["lesson_id"].isin(selected_lessons)].copy()

if role_filter != "all":
    filtered = filtered[filtered["role"] == role_filter]


# =========================
# RQ1
# =========================

st.header("RQ1. How is classroom discourse distributed and dialog act types?")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Turn frequency")

    if role_filter == "all":
        rq1_turns = (
            filtered.groupby("role")
            .size()
            .rename("turn_count")
            .reset_index()
        )
        st.bar_chart(rq1_turns.set_index("role")["turn_count"])
    else:
        st.metric(
            label=f"Total turns ({role_filter})",
            value=len(filtered)
        )

with col2:
    st.subheader("DialogAct distribution (Top 10)")

    da_counts = (
        filtered["DialogAct"]
        .dropna()
        .value_counts()
        .head(10)
    )

    st.bar_chart(da_counts)


# =========================
# RQ2
# =========================

st.header("RQ2. What patterns characterize students’ discourse contributions?")

if role_filter == "teacher":
    st.info("Switch Role filter to 'student' or 'all' to view student analysis.")
else:
    students = filtered[filtered["role"] == "student"]

    st.subheader("Student_Tag distribution (Top 10)")

    st_counts = (
        students["Student_Tag"]
        .dropna()
        .value_counts()
        .head(10)
    )

    st.bar_chart(st_counts)


# =========================
# RQ3
# =========================

st.header("RQ3. How are teachers’ instructional intentions enacted in their discourse?")

if role_filter == "student":
    st.info("Switch Role filter to 'teacher' or 'all' to view teacher analysis.")
else:
    teachers = filtered[filtered["role"] == "teacher"]

    if len(teachers) == 0:
        st.warning("No teacher data available for selected filter.")
    else:
        st.subheader("Teacher_Tag × DialogAct (Proportion Heatmap)")

        ct_counts = pd.crosstab(
            teachers["Teacher_Tag"],
            teachers["DialogAct"]
        )

        ct_props = ct_counts.div(ct_counts.sum(axis=1), axis=0).fillna(0)
        ct_props.index.name = "Teacher_Tag"

        vega_heatmap(ct_props, "Teacher_Tag × DialogAct (Row Proportions)")