import streamlit as st
import pandas as pd
from io import BytesIO, StringIO


# ---------------------------------------------------------
# Page configuration
# ---------------------------------------------------------

st.set_page_config(
    page_title="CSV Consolidator",
    page_icon="📈",
    layout="wide"
)


# ---------------------------------------------------------
# Custom styles
# ---------------------------------------------------------

st.markdown(
    """
    <style>
        .main {
            background-color: #f7f9fc;
        }

        .block-container {
            padding-top: 2rem;
        }

        .title {
            font-size: 2.2rem;
            font-weight: 700;
            color: #1f3c88;
        }

        .subtitle {
            color: #667085;
            font-size: 1.05rem;
            margin-bottom: 1.5rem;
        }

        div[data-testid="stMetric"] {
            background-color: white;
            padding: 15px;
            border-radius: 10px;
            box-shadow: 0 2px 8px rgba(0, 0, 0, 0.06);
        }
    </style>
    """,
    unsafe_allow_html=True
)


# ---------------------------------------------------------
# Header
# ---------------------------------------------------------

st.markdown(
    '<div class="title">📈 CSV File Consolidator</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="subtitle">'
    'Upload filtered CSV files and generate a single consolidated file.'
    '</div>',
    unsafe_allow_html=True
)


# ---------------------------------------------------------
# CSV reading functions
# ---------------------------------------------------------

def detect_encoding(file_bytes):
    """
    Identifies the file encoding.
    """

    for encoding in ["utf-8-sig", "utf-8", "latin-1"]:
        try:
            text = file_bytes.decode(encoding)
            return text, encoding

        except UnicodeDecodeError:
            continue

    raise ValueError(
        "The file encoding could not be identified."
    )


def find_header_line(text):
    """
    Finds the actual header line.

    The function ignores empty lines or metadata before
    the actual column names.

    Examples of expected column names:
    Product[MSPN Nbr]
    Product[CAI Nbr]
    Customer[BT Name]
    """

    lines = text.splitlines()

    # First, look for the expected header pattern
    for index, line in enumerate(lines):
        clean_line = line.strip()

        if not clean_line:
            continue

        contains_column_name = (
            "Product[" in clean_line
            or "Customer[" in clean_line
        )

        contains_separators = (
            clean_line.count(",") >= 2
            or clean_line.count(";") >= 2
            or clean_line.count("\t") >= 2
        )

        if contains_column_name and contains_separators:
            return index

    # Fallback: find the first non-empty line
    # containing multiple separators
    for index, line in enumerate(lines):
        clean_line = line.strip()

        if not clean_line:
            continue

        if clean_line.count(",") >= 2:
            return index

        if clean_line.count(";") >= 2:
            return index

        if clean_line.count("\t") >= 2:
            return index

    # If no header is found, use the first line
    return 0


def detect_separator(header_line):
    """
    Detects the CSV separator based on the header line.
    """

    comma_count = header_line.count(",")
    semicolon_count = header_line.count(";")
    tab_count = header_line.count("\t")

    separators = {
        ",": comma_count,
        ";": semicolon_count,
        "\t": tab_count
    }

    separator = max(
        separators,
        key=separators.get
    )

    if separators[separator] == 0:
        raise ValueError(
            "The CSV separator could not be identified."
        )

    return separator


def read_csv_automatically(uploaded_file):
    """
    Reads the CSV file and automatically identifies:

    - File encoding
    - Separator
    - Actual header line
    """

    file_bytes = uploaded_file.getvalue()

    text, encoding = detect_encoding(file_bytes)

    lines = text.splitlines()

    if not lines:
        raise ValueError("The file is empty.")

    header_index = find_header_line(text)

    header_line = lines[header_index]

    separator = detect_separator(header_line)

    # Remove all lines before the actual header
    content_from_header = "\n".join(
        lines[header_index:]
    )

    dataframe = pd.read_csv(
        StringIO(content_from_header),
        sep=separator,
        header=0,
        skip_blank_lines=True,
        quotechar='"',
        dtype=str,
        engine="python"
    )

    # Remove extra spaces from column names
    dataframe.columns = (
        dataframe.columns
        .astype(str)
        .str.strip()
    )

    # Remove columns generated as "Unnamed"
    dataframe = dataframe.loc[
        :,
        ~dataframe.columns.str.match(r"^Unnamed")
    ]

    # Remove completely empty rows
    dataframe = (
        dataframe
        .dropna(how="all")
        .reset_index(drop=True)
    )

    return (
        dataframe,
        separator,
        encoding,
        header_index
    )


# ---------------------------------------------------------
# Cached file generation
# ---------------------------------------------------------

@st.cache_data
def generate_csv_file(dataframe):
    """
    Generates the CSV file only once per unique dataframe.
    Streamlit caches the result based on the dataframe content,
    avoiding regeneration on every widget interaction.
    """

    return dataframe.to_csv(
        index=False,
        sep=",",
        encoding="utf-8-sig"
    )


@st.cache_data
def generate_excel_file(dataframe):
    """
    Generates the Excel file only once per unique dataframe.
    Streamlit caches the result based on the dataframe content,
    avoiding regeneration on every widget interaction.

    Uses the xlsxwriter engine, which is generally faster than
    openpyxl for writing large volumes of rows.
    """

    excel_buffer = BytesIO()

    with pd.ExcelWriter(
        excel_buffer,
        engine="xlsxwriter"
    ) as writer:

        dataframe.to_excel(
            writer,
            index=False,
            sheet_name="Consolidated"
        )

    return excel_buffer.getvalue()


# ---------------------------------------------------------
# Uploader reset control
# ---------------------------------------------------------

# This counter is used to change the file_uploader's key.
# Changing the key forces Streamlit to recreate the widget
# from scratch, which effectively clears any uploaded files.
if "uploader_key" not in st.session_state:
    st.session_state.uploader_key = 0


def clear_uploaded_files():
    """
    Clears the uploaded files by resetting the file_uploader key.
    """

    st.session_state.uploader_key += 1


# ---------------------------------------------------------
# File upload
# ---------------------------------------------------------

st.subheader("1. Upload CSV files")

upload_column, clear_button_column = st.columns([5, 1])

with upload_column:
    uploaded_files = st.file_uploader(
        "Select one or more CSV files",
        type=["csv"],
        accept_multiple_files=True,
        help="The files should have compatible columns.",
        key=f"file_uploader_{st.session_state.uploader_key}"
    )

with clear_button_column:
    st.write("")
    st.write("")

    st.button(
        "🗑️ Clear files",
        on_click=clear_uploaded_files,
        use_container_width=True
    )


if uploaded_files:

    dataframes = []
    errors = []

    with st.spinner("Reading files..."):

        for uploaded_file in uploaded_files:

            try:
                (
                    dataframe,
                    separator,
                    encoding,
                    header_index
                ) = read_csv_automatically(uploaded_file)

                if dataframe.empty:
                    errors.append(
                        f"{uploaded_file.name}: the file is empty."
                    )
                    continue

                if len(dataframe.columns) <= 1:
                    errors.append(
                        f"{uploaded_file.name}: only one column was "
                        "identified. Check the file separator."
                    )
                    continue

                # Add the source file name
                dataframe["source_file"] = uploaded_file.name

                dataframes.append(dataframe)

                st.success(
                    f"**{uploaded_file.name}** — "
                    f"header found on line {header_index + 1} — "
                    f"separator `{separator}` — "
                    f"encoding `{encoding}` — "
                    f"{len(dataframe.columns)} columns — "
                    f"{len(dataframe)} rows"
                )

            except Exception as error:
                errors.append(
                    f"{uploaded_file.name}: {error}"
                )

    # Display errors
    if errors:
        st.warning("Some files could not be processed:")

        for error in errors:
            st.write(f"- {error}")

    # Continue only if at least one file was successfully read
    if dataframes:

        st.success(
            f"{len(dataframes)} file(s) successfully loaded."
        )

        # Consolidate all dataframes
        consolidated_dataframe = pd.concat(
            dataframes,
            ignore_index=True,
            sort=False
        )

        # ---------------------------------------------------------
        # Summary
        # ---------------------------------------------------------

        st.subheader("2. Consolidation summary")

        metric_1, metric_2, metric_3, metric_4 = st.columns(4)

        with metric_1:
            st.metric(
                "Files loaded",
                len(dataframes)
            )

        with metric_2:
            st.metric(
                "Total rows",
                len(consolidated_dataframe)
            )

        with metric_3:
            st.metric(
                "Columns",
                len(consolidated_dataframe.columns)
            )

        with metric_4:
            st.metric(
                "Duplicate rows",
                int(
                    consolidated_dataframe
                    .duplicated()
                    .sum()
                )
            )

        # ---------------------------------------------------------
        # Processing options
        # ---------------------------------------------------------

        st.subheader("3. Processing options")

        option_1, option_2 = st.columns(2)

        with option_1:
            remove_duplicates = st.checkbox(
                "Remove duplicate rows",
                value=False
            )

        with option_2:
            keep_source_file = st.checkbox(
                "Keep source file column",
                value=True
            )

        final_dataframe = consolidated_dataframe.copy()

        if remove_duplicates:
            final_dataframe = (
                final_dataframe
                .drop_duplicates()
                .reset_index(drop=True)
            )

        if (
            not keep_source_file
            and "source_file" in final_dataframe.columns
        ):
            final_dataframe = final_dataframe.drop(
                columns=["source_file"]
            )

        # ---------------------------------------------------------
        # Column selection
        # ---------------------------------------------------------

        st.subheader("4. Column selection")

        selected_columns = st.multiselect(
            "Choose the columns you want to keep",
            options=list(final_dataframe.columns),
            default=list(final_dataframe.columns)
        )

        if selected_columns:
            final_dataframe = final_dataframe[
                selected_columns
            ]

        # ---------------------------------------------------------
        # Data preview
        # ---------------------------------------------------------

        st.subheader("5. Data preview")

        st.dataframe(
            final_dataframe.head(100),
            use_container_width=True,
            hide_index=True
        )

        st.caption(
            f"Displaying up to 100 rows out of "
            f"{len(final_dataframe)} total rows."
        )

        # ---------------------------------------------------------
        # Download files
        # ---------------------------------------------------------

        st.subheader("6. Download")

        download_csv_column, download_excel_column = st.columns(2)

        # Generate consolidated CSV (cached)
        final_csv = generate_csv_file(final_dataframe)

        with download_csv_column:
            st.download_button(
                label="⬇️ Download consolidated CSV",
                data=final_csv,
                file_name="consolidated_file.csv",
                mime="text/csv",
                use_container_width=True
            )

        # Generate consolidated Excel file (cached)
        final_excel = generate_excel_file(final_dataframe)

        with download_excel_column:
            st.download_button(
                label="⬇️ Download consolidated Excel",
                data=final_excel,
                file_name="consolidated_file.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument"
                    ".spreadsheetml.sheet"
                ),
                use_container_width=True
            )

else:
    st.info(
        "Upload your CSV files to start the consolidation."
    )