"""Streamlit web app for WorkoutParser."""

import io

import streamlit as st

from workout_parser.parser import parse_files, validate_file

st.title("Workout Parser")
st.write("Upload one or more workout Excel files to parse and consolidate them.")

with st.expander("Expected file format"):
    st.markdown(
        """
**Filename** must follow the pattern: `Month-Year-...` (e.g., `December-2024-4-Day-Full-Gym-Routine.xlsx`)

**Workbook structure:**
- Sheets named "Week 1", "Week 2", etc.
- Exercise data is read from the **Client Notes** section
- Sets/reps/weight in the format `{sets}x{reps}x{weight}` (e.g., `3x10x135`)
"""
    )

uploaded_files = st.file_uploader(
    "Choose .xlsx files", type="xlsx", accept_multiple_files=True
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
