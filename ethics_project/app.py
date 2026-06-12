
from __future__ import annotations

from pathlib import Path
import csv
import re
import pandas as pd
import streamlit as st

from comparison_tables import (
    DEFAULT_MAX_ROWS,
    DEFAULT_MODELS,
    build_pairs_from_directory,
    create_side_by_side_workbook,
    flatten_average_columns,
    make_average_scores_dataframe,
    make_incident_average_scores_dataframe,
    make_side_by_side_dataframe,
    read_comparison_file,
)

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
RATINGS_PATH = DATA_DIR / "ratings.csv"
SECRET_PATH = DATA_DIR / "secret_table.csv"

APP_COLUMNS = ["News Title", "My Ground Truth", "ChatGPT", "Claude", "Gemini"]
RATING_COLUMNS = APP_COLUMNS[1:]


def clean_cell(value: object) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip().replace("\ufeff", "")
    if len(text) >= 2 and text[0] == '"' and text[-1] == '"':
        text = text[1:-1]
    elif text.startswith('"'):
        text = text[1:]
    elif text.endswith('"'):
        text = text[:-1]
    return text.replace('""', '"').strip()


def clean_rating(value: object):
    text = clean_cell(value)
    if text == "" or text.lower() in {"none", "nan", "null", "<na>"}:
        return pd.NA
    match = re.search(r"-?\d+", text)
    if not match:
        return pd.NA
    number = int(match.group())
    if 0 <= number <= 10:
        return number
    return pd.NA


def parse_rating_csv_text(text: str) -> pd.DataFrame:
    """Parse a CSV even if titles contain commas or a row was wrongly quoted.

    It always treats the last 4 comma-separated fields as the rating columns.
    Everything before those final 4 fields is the title.
    """
    rows = []
    lines = [line.rstrip("\r\n") for line in text.splitlines() if line.strip()]
    for line in lines[1:]:
        parts = line.rsplit(",", 4)
        if len(parts) != 5:
            # Fallback for a correctly quoted CSV row.
            try:
                fields = next(csv.reader([line]))
            except Exception:
                continue
            if len(fields) != 5:
                continue
            parts = fields
        row = [clean_cell(parts[0])] + [clean_rating(x) for x in parts[1:]]
        if row[0]:
            rows.append(row)
    df = pd.DataFrame(rows, columns=APP_COLUMNS)
    return normalise_rating_table(df)


def read_rating_file(path_or_buffer) -> pd.DataFrame:
    name = getattr(path_or_buffer, "name", str(path_or_buffer)).lower()
    if name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(path_or_buffer)
        return normalise_rating_table(df)

    if hasattr(path_or_buffer, "seek"):
        path_or_buffer.seek(0)
        raw = path_or_buffer.read()
        if isinstance(raw, bytes):
            text = raw.decode("utf-8-sig", errors="replace")
        else:
            text = raw
    else:
        text = Path(path_or_buffer).read_text(encoding="utf-8-sig", errors="replace")
    return parse_rating_csv_text(text)


def normalise_rating_table(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [clean_cell(c) for c in df.columns]
    if "News Title" not in df.columns:
        candidates = [c for c in df.columns if "title" in c.lower() or "headline" in c.lower()]
        if not candidates:
            raise ValueError("Missing News Title column.")
        df = df.rename(columns={candidates[0]: "News Title"})

    df["News Title"] = df["News Title"].map(clean_cell)
    df = df[df["News Title"].ne("")].drop_duplicates(subset=["News Title"], keep="first")

    for col in RATING_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA
        df[col] = df[col].map(clean_rating).astype("Int64")

    return df[APP_COLUMNS].head(100).reset_index(drop=True)


def load_rating_table(force: bool = False) -> pd.DataFrame:
    DATA_DIR.mkdir(exist_ok=True)
    if not RATINGS_PATH.exists():
        empty = pd.DataFrame({"News Title": []})
        for col in RATING_COLUMNS:
            empty[col] = pd.Series(dtype="Int64")
        return normalise_rating_table(empty)

    mtime = RATINGS_PATH.stat().st_mtime
    if force or "ratings_df" not in st.session_state or st.session_state.get("ratings_mtime") != mtime:
        st.session_state.ratings_df = read_rating_file(RATINGS_PATH)
        st.session_state.ratings_mtime = mtime
    return st.session_state.ratings_df


def save_rating_table(df: pd.DataFrame) -> None:
    DATA_DIR.mkdir(exist_ok=True)
    clean = normalise_rating_table(df)
    clean.to_csv(RATINGS_PATH, index=False, encoding="utf-8-sig")
    st.session_state.ratings_df = clean
    st.session_state.ratings_mtime = RATINGS_PATH.stat().st_mtime


def render_incident_page() -> None:
    st.title("AI Ethics Incident Rating Table")
    st.caption("Scale: 0 = extremely unethical, 10 = highly ethical. Your ground truth is the benchmark column.")

    df = load_rating_table()
    counts = df[RATING_COLUMNS].notna().sum()
    st.info("Loaded markings — " + " | ".join(f"{c}: {int(counts[c])}/{len(df)}" for c in RATING_COLUMNS))

    with st.expander("Data file status", expanded=False):
        st.write(f"Using file: `{RATINGS_PATH}`")
        st.dataframe(df.head(5), use_container_width=True, hide_index=True)

    uploaded_file = st.file_uploader("Optional: import ratings CSV/XLSX", type=["csv", "txt", "xlsx", "xls"])
    if uploaded_file is not None:
        if st.button("Import uploaded file"):
            imported = read_rating_file(uploaded_file)
            save_rating_table(imported)
            st.success("Imported and saved. The table below now uses the imported markings.")
            st.rerun()

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("Reload ratings from disk"):
            load_rating_table(force=True)
            st.success("Reloaded ratings.csv from disk.")
            st.rerun()
    with c2:
        if st.button("Save ratings", type="primary"):
            save_rating_table(st.session_state.ratings_df)
            st.success("Saved ratings.csv.")
    with c3:
        csv_bytes = df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        st.download_button("Download current table", data=csv_bytes, file_name="ai_ethics_ratings.csv", mime="text/csv")

    edited = st.data_editor(
        st.session_state.ratings_df,
        use_container_width=True,
        hide_index=True,
        disabled=["News Title"],
        num_rows="fixed",
        column_config={
            "News Title": st.column_config.TextColumn("News Title", width="large"),
            "My Ground Truth": st.column_config.NumberColumn("My Ground Truth", min_value=0, max_value=10, step=1, format="%d"),
            "ChatGPT": st.column_config.NumberColumn("ChatGPT", min_value=0, max_value=10, step=1, format="%d"),
            "Claude": st.column_config.NumberColumn("Claude", min_value=0, max_value=10, step=1, format="%d"),
            "Gemini": st.column_config.NumberColumn("Gemini", min_value=0, max_value=10, step=1, format="%d"),
        },
    )
    st.session_state.ratings_df = normalise_rating_table(edited)



def render_comparison_page() -> None:
    st.title("Israeli vs Russian Crimes")
    st.caption(
        "Upload six files: Israel and Russia files for ChatGPT, Claude, and Gemini. "
        "The app creates three sheets with one Israel title column and paired Israel/Russia columns for Ethicality, Visibility, and Accountability."
    )

    uploaded: dict[tuple[str, int], object] = {}
    for model in DEFAULT_MODELS:
        st.subheader(model)
        c1, c2 = st.columns(2)
        with c1:
            uploaded[(model, 1)] = st.file_uploader(
                f"{model} Israel file",
                type=["csv", "txt", "xlsx", "xls"],
                key=f"{model.lower()}_news1_upload",
            )
        with c2:
            uploaded[(model, 2)] = st.file_uploader(
                f"{model} Russia file",
                type=["csv", "txt", "xlsx", "xls"],
                key=f"{model.lower()}_news2_upload",
            )

    all_files_uploaded = all(file is not None for file in uploaded.values())
    if not all_files_uploaded:
        st.info("Upload all six files to generate the comparison workbook.")
        return

    try:
        pairs = {}
        source_info = {}
        for model in DEFAULT_MODELS:
            news1_file = uploaded[(model, 1)]
            news2_file = uploaded[(model, 2)]
            news1_df = read_comparison_file(news1_file, model_hint=model)
            news2_df = read_comparison_file(news2_file, model_hint=model)
            pairs[model] = (news1_df, news2_df)
            source_info[model] = {
                "news1_file": getattr(news1_file, "name", "Israel"),
                "news2_file": getattr(news2_file, "name", "Russia"),
            }

        st.session_state.comparison_pairs = pairs
        workbook_bytes = create_side_by_side_workbook(
            pairs,
            source_info=source_info,
            max_rows=DEFAULT_MAX_ROWS,
            incident_ratings=load_rating_table(),
        )
        st.success("Comparison workbook generated.")
        st.download_button(
            "Download Israeli vs Russian Crimes workbook",
            data=workbook_bytes,
            file_name="israeli_vs_russian_crimes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        with st.expander("All comparison rows", expanded=True):
            for model, (news1_df, news2_df) in pairs.items():
                st.markdown(f"**{model} Comparison**")
                full_table = make_side_by_side_dataframe(news1_df, news2_df, max_rows=DEFAULT_MAX_ROWS)
                st.dataframe(full_table, use_container_width=True, hide_index=True, height=720)
                st.caption(f"Rows shown: {DEFAULT_MAX_ROWS} | Rows loaded — Israel: {len(news1_df)} | Russia: {len(news2_df)}")
    except Exception as exc:
        st.error(f"Could not generate comparison workbook: {exc}")


def load_secret_average_pairs() -> dict[str, tuple[pd.DataFrame, pd.DataFrame]]:
    """Use uploaded comparison files when present; otherwise use bundled inputs."""
    if "comparison_pairs" in st.session_state:
        return st.session_state.comparison_pairs
    pairs, _source_info = build_pairs_from_directory(DATA_DIR / "comparison_inputs")
    return pairs


def render_secret_page() -> None:
    st.title("Final Comparison")
    st.caption(
        "Average scores in two compact summary tables: Israeli vs Russian Crimes by AI model, "
        "then incident-rating averages for your ground truth and each AI."
    )

    try:
        pairs = load_secret_average_pairs()
        average_table = make_average_scores_dataframe(pairs, max_rows=DEFAULT_MAX_ROWS)
    except Exception as exc:
        st.error(f"Could not build the final comparison table: {exc}")
        st.info("Upload the six comparison files on the Israeli vs Russian Crimes page, or keep them in data/comparison_inputs.")
        return

    st.subheader("Israeli vs Russian Crimes Average Scores")
    st.dataframe(average_table, use_container_width=True, hide_index=True)

    incident_table = make_incident_average_scores_dataframe(load_rating_table())
    st.subheader("Incident Ratings Average Scores")
    st.caption("Average score for your ground truth and each AI model across the incident ratings table. Blank cells are ignored.")
    st.dataframe(incident_table, use_container_width=True, hide_index=True)

    csv_table = pd.concat(
        {
            "Israeli vs Russian Crimes": flatten_average_columns(average_table),
            "Incident Ratings": flatten_average_columns(incident_table),
        },
        names=["Table", "Row"],
    ).reset_index(level="Table")
    csv_bytes = csv_table.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
    st.download_button(
        "Download Final Comparison",
        data=csv_bytes,
        file_name="final_comparison_average_scores.csv",
        mime="text/csv",
    )

def main() -> None:
    st.set_page_config(page_title="AI Ethics Rating App", layout="wide")
    page = st.sidebar.radio("Page", ["Incident Ratings", "Israeli vs Russian Crimes", "Final Comparison"])
    if page == "Incident Ratings":
        render_incident_page()
    elif page == "Israeli vs Russian Crimes":
        render_comparison_page()
    else:
        render_secret_page()


if __name__ == "__main__":
    main()
