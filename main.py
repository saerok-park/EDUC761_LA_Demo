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
    xlsx_files = [f for f in glob.glob(os.path.join(data_dir, "*.xlsx"))
                  if not os.path.basename(f).startswith("~$")]
    if not xlsx_files:
        raise FileNotFoundError(f"No .xlsx files found in: {data_dir}")

    dfs = []
    for path in xlsx_files:
        df = pd.read_excel(path, engine="openpyxl")

        # 컬럼 체크
        missing = [c for c in REQUIRED_COLS if c not in df.columns]
        if missing:
            raise ValueError(
                f"File '{os.path.basename(path)}' is missing columns: {missing}\n"
                f"Found columns: {list(df.columns)}"
            )

        lesson_id = os.path.splitext(os.path.basename(path))[0]
        df["lesson_id"] = lesson_id
        dfs.append(df[REQUIRED_COLS + ["lesson_id"]])

    data = pd.concat(dfs, ignore_index=True)

    # 기본 정리
    for c in ["Speaker", "Sentence", "Teacher_Tag", "Student_Tag", "DialogAct"]:
        data[c] = data[c].astype("string").str.strip()

    data["Turn_num"] = pd.to_numeric(data["Turn"], errors="coerce")
    data["_row"] = range(len(data))
    data = data.sort_values(by=["lesson_id", "Turn_num", "_row"]).reset_index(drop=True)

    data["role"] = data["Speaker"].apply(get_role)

    # 빈 문자열 -> NA
    for c in ["Teacher_Tag", "Student_Tag", "DialogAct"]:
        data[c] = data[c].replace({"": pd.NA, "nan": pd.NA, "NaN": pd.NA})

    return data


def vega_heatmap(df_props: pd.DataFrame, title: str):
    """
    df_props: index=Teacher_Tag, columns=DialogAct, values=proportions (0~1)
    Streamlit 내장 Vega-Lite로 heatmap
    """
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
        "config": {"view": {"stroke": "transparent"}}
    }
    st.vega_lite_chart(spec, use_container_width=True)


# =========================
# UI
# =========================
st.set_page_config(page_title="Classroom Discourse Dashboard", layout="wide")
st.title("Classroom Discourse Dashboard")

data = load_and_merge_xlsx(DATA_DIR)

# Sidebar filters
st.sidebar.header("Filters")
lessons = sorted(data["lesson_id"].unique().tolist())
selected_lessons = st.sidebar.multiselect("Select lessons", lessons, default=lessons)

role_filter = st.sidebar.selectbox("Role filter (for preview only)", ["all", "teacher", "student"])
show_preview = st.sidebar.checkbox("Show data preview", value=False)

filtered = data[data["lesson_id"].isin(selected_lessons)].copy()

if show_preview:
    st.subheader("Preview (first 50 rows)")
    preview = filtered.copy()
    if role_filter != "all":
        preview = preview[preview["role"] == role_filter]
    st.dataframe(
        preview[["lesson_id", "TimeStamp", "Turn", "Speaker", "role", "Sentence", "Teacher_Tag", "Student_Tag", "DialogAct"]].head(50),
        use_container_width=True
    )

# =========================
# RQ1
# =========================
st.header("RQ1. How is classroom discourse distributed and dialog act types?")

c1, c2 = st.columns(2)

with c1:
    st.subheader("Teacher vs. student turn frequency")
    rq1_turns = (
        filtered.groupby("role")
        .size()
        .rename("turn_count")
        .reset_index()
        .sort_values("turn_count", ascending=False)
    )
    st.dataframe(rq1_turns, use_container_width=True)
    st.bar_chart(rq1_turns.set_index("role")["turn_count"])

with c2:
    st.subheader("Overall DialogAct distribution (Top 10)")
    da_counts = (
        filtered["DialogAct"]
        .dropna()
        .value_counts()
        .rename("count")
        .reset_index()
        .rename(columns={"index": "DialogAct"})
    )
    top_da = da_counts.head(10)
    st.dataframe(top_da, use_container_width=True)
    st.bar_chart(top_da.set_index("DialogAct")["count"])

# =========================
# RQ2
# =========================
st.header("RQ2. What patterns characterize students’ discourse contributions?")

students = filtered[filtered["role"] == "student"].copy()
st.subheader("Student_Tag distribution (Top 10)")
st_counts = (
    students["Student_Tag"]
    .dropna()
    .value_counts()
    .rename("count")
    .reset_index()
    .rename(columns={"index": "Student_Tag"})
)
top_st = st_counts.head(10)
st.dataframe(top_st, use_container_width=True)
st.bar_chart(top_st.set_index("Student_Tag")["count"])

# =========================
# RQ3
# =========================
st.header("RQ3. How are teachers’ instructional intentions enacted in their discourse?")

teachers = filtered[filtered["role"] == "teacher"].copy()

st.subheader("Teacher_Tag × DialogAct (Row Proportion Heatmap)")
ct_counts = pd.crosstab(teachers["Teacher_Tag"], teachers["DialogAct"], dropna=False)
ct_props = ct_counts.div(ct_counts.sum(axis=1), axis=0).fillna(0)
ct_props.index.name = "Teacher_Tag"

vega_heatmap(ct_props, "Teacher_Tag × DialogAct (Proportions)")

with st.expander("See tables (counts & proportions)"):
    st.write("Counts")
    st.dataframe(ct_counts.fillna(0).astype(int), use_container_width=True)
    st.write("Row proportions")
    st.dataframe(ct_props.round(3), use_container_width=True)