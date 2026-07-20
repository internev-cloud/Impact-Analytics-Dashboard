"""
topic_dashboard.py
===================
Topic & SubTopic Analytics — eVidyaloka
Standalone home-page module (previously duplicated as ops_dashboard.py's
5th tab). Data source: a single workbook sitting alongside this script —
see TOPIC_FILE below.

Sheets used
───────────
  Topic-SubTopic     — primary session-level log (1 row per session)
  Cancelled Sessions — sessions that were cancelled (loaded, not yet charted)
  Offline Sessions   — sessions that went offline (loaded, not yet charted)
"""

import streamlit as st
import pandas as pd
import os

DATA_DIR = os.path.dirname(os.path.abspath(__file__))
TOPIC_FILE = os.path.join(DATA_DIR, "Topic_SubTopic_Cancelled_Offline_May_2026.xlsx")


@st.cache_data(show_spinner=False)
def load_topic_data(path: str):
    """Returns (topic_df, cancelled_df, offline_df), or (None, None, None) if missing."""
    if not os.path.exists(path):
        return None, None, None
    topic_df     = pd.read_excel(path, sheet_name="Topic-SubTopic")
    cancelled_df = pd.read_excel(path, sheet_name="Cancelled Sessions")
    offline_df   = pd.read_excel(path, sheet_name="Offline Sessions")
    return topic_df, cancelled_df, offline_df


def render_topic_dashboard():
    st.markdown("""
    <style>
    .metric-card{
        background:#ffffff; padding:18px; border-radius:12px;
        box-shadow:0 2px 10px rgba(0,0,0,.08);
    }
    [data-testid="stMetricValue"]{
        font-size:34px; font-weight:700; color:#0f766e;
    }
    </style>
    """, unsafe_allow_html=True)

    st.title("📚 Topic & SubTopic Analytics Dashboard")
    st.markdown(
        "<p style='color:gray;font-size:1.1em;margin-top:-10px;'>"
        "Session-level Topic & Sub-topic Coverage Analysis</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    topic_df, cancelled_df, offline_df = load_topic_data(TOPIC_FILE)

    if topic_df is None:
        st.error(
            f"⚠️ Topic/SubTopic data file not found. "
            f"Place `{os.path.basename(TOPIC_FILE)}` in `{DATA_DIR}`."
        )
        return

    base_df = topic_df.copy()

    # ── Filters ────────────────────────────────────────────────────────────
    with st.sidebar:
        st.markdown("---")
        st.header("📊 Topic Filters")

        state_filter = st.multiselect(
            "State", sorted(base_df["state"].dropna().unique()), key="topic_state"
        )
        state_df = base_df[base_df["state"].isin(state_filter)] if state_filter else base_df.copy()

        subject_filter = st.multiselect(
            "Subject", sorted(state_df["subject"].dropna().unique()), key="topic_subject"
        )
        subject_df = state_df[state_df["subject"].isin(subject_filter)] if subject_filter else state_df.copy()

        grade_filter = st.multiselect(
            "Grade", sorted(subject_df["grade"].dropna().unique()), key="topic_grade"
        )
        grade_df = subject_df[subject_df["grade"].isin(grade_filter)] if grade_filter else subject_df.copy()

        status_filter = st.multiselect(
            "Session Status", sorted(grade_df["session_status"].dropna().unique()), key="topic_status"
        )
        status_df = grade_df[grade_df["session_status"].isin(status_filter)] if status_filter else grade_df.copy()

        centre_filter = st.multiselect(
            "Centre", sorted(status_df["center_name"].dropna().unique()), key="topic_centre"
        )
        filtered_df = (
            status_df[status_df["center_name"].isin(centre_filter)]
            if centre_filter else status_df.copy()
        )

    if filtered_df.empty:
        st.warning("⚠️ No data for the selected filters. Please adjust your criteria.")
        return

    # ── KPI cards ──────────────────────────────────────────────────────────
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Sessions",    f"{len(filtered_df):,}")
    col2.metric("Unique Topics",     f"{filtered_df['topic_name'].nunique():,}")
    col3.metric("Unique Subtopics",  f"{filtered_df['sub_topic_name'].nunique():,}")
    col4.metric("Active Centers",    f"{filtered_df['center_id'].nunique():,}")

    st.markdown("---")

    # ── Topic / sub-topic summary table ───────────────────────────────────
    st.subheader("📋 Topic / Sub-topic Session Summary")

    pivot_table = (
        filtered_df
        .groupby(["topic_name", "sub_topic_name"], dropna=False)
        .agg(**{
            "#Sessions": ("topic_name", "count"),
            "#Schools":  ("center_name", "nunique"),
        })
        .reset_index()
        .sort_values("#Sessions", ascending=False)
        .rename(columns={"topic_name": "Topic", "sub_topic_name": "Sub-topic"})
    )

    st.dataframe(pivot_table, use_container_width=True, hide_index=True, height=700)

    st.download_button(
        "📥 Download Topic Data",
        filtered_df.to_csv(index=False),
        file_name="topic_data.csv",
        key="topic_download",
    )
