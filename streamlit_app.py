"""Streamlit web app for WorkoutParser."""

import io

import altair as alt
import pandas as pd
import streamlit as st

from workout_parser.parser import parse_files, validate_file

st.title("Workout Parser")
st.write("Upload one or more workout spreadsheet files (.xlsx or .numbers) to parse and consolidate them.")

with st.expander("Expected file format"):
    st.markdown(
        """
**Filename** must follow the pattern: `Month-Year-...` (e.g., `December-2024-4-Day-Full-Gym-Routine.xlsx` or `.numbers`)

**Workbook structure:**
- Sheets named "Week 1", "Week 2", etc.
- Exercise data is read from the **Client Notes** section
- Sets/reps/weight in the format `{sets}x{reps}x{weight}` (e.g., `3x10x135`)
"""
    )

uploaded_files = st.file_uploader(
    "Choose .xlsx or .numbers files", type=["xlsx", "numbers"], accept_multiple_files=True
)

if uploaded_files and st.button("Parse"):
    errors = []
    for f in uploaded_files:
        error = validate_file(f.name)
        if error:
            errors.append(error)

    if errors:
        for error in errors:
            st.error(error)
    else:
        with st.spinner("Parsing..."):
            df = parse_files(uploaded_files)
        st.session_state["parsed_df"] = df

if "parsed_df" in st.session_state:
    df = st.session_state["parsed_df"]

    if df.empty:
        st.warning("No exercises found in the uploaded files.")
    else:
        st.success(f"Parsed {len(df)} exercises.")
        st.dataframe(df)

        buf = io.BytesIO()
        df.to_excel(buf, index=False, sheet_name="Monthly Exercises")
        st.download_button(
            "Download Excel",
            data=buf.getvalue(),
            file_name="parsed_workouts.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # --- Trend Chart ---
        st.subheader("Exercise Trends")

        exercises = sorted(df["Name"].unique())
        selected = st.multiselect("Select exercises", exercises, default=exercises[:1])

        if selected:
            metric = st.radio("Y-axis", ["Weight", "Reps"], horizontal=True)
            filtered = df[df["Name"].isin(selected)].copy()
            filtered["Date"] = pd.to_datetime(filtered["Date"])

            chart = (
                alt.Chart(filtered)
                .mark_line(point=True)
                .encode(
                    x=alt.X("Date:T", title="Date"),
                    y=alt.Y(f"{metric}:Q", title=metric),
                    color=alt.Color("Name:N", title="Exercise"),
                    tooltip=["Date:T", "Name:N", "Sets:Q", "Reps:Q", "Weight:Q"],
                )
                .interactive()
            )

            st.altair_chart(chart, use_container_width=True)
