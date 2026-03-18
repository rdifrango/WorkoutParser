"""Streamlit web app for WorkoutParser."""

import io

import streamlit as st

from workout_parser.parser import parse_files, validate_file

st.title("Workout Parser")
st.write("Upload one or more workout Excel files to parse and consolidate them.")

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
