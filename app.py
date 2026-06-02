import streamlit as st
import pandas as pd
import pdfplumber
import pytesseract
pytesseract.pytesseract.tesseract_cmd = "tesseract"
import re
import tempfile

from pdf2image import convert_from_path
from io import BytesIO

# CONFIG

st.set_page_config(
    page_title="PDF to Excel Extractor",
    layout="wide"
)

st.title("PDF to Excel Extractor")

st.write("""
Upload scanned or normal PDFs and extract structured Excel data.
""")

# PDF TEXT EXTRACTION

def extract_text_from_pdf(pdf_path):

    full_text = ""

    try:

        with pdfplumber.open(pdf_path) as pdf:

            for page in pdf.pages:

                text = page.extract_text(
                    x_tolerance=2,
                    y_tolerance=2
                )

                if text:
                    full_text += "\n" + text

    except Exception as e:

        st.error(f"Text Extraction Error: {e}")

    # OCR FALLBACK

    if len(full_text.strip()) < 50:

        st.info("Scanned PDF detected. Running OCR...")

        try:

            images = convert_from_path(pdf_path)

            ocr_text = ""

            for img in images:

                text = pytesseract.image_to_string(img)

                ocr_text += "\n" + text

            full_text += ocr_text

        except Exception as e:

            st.error(f"OCR Error: {e}")

    return full_text

# CLEAN TEXT

def clean_text(text):

    # Normalize line breaks
    text = text.replace("\r", "\n")

    # OCR cleanup
    text = text.replace("|", " ")

    text = text.replace(";", ":")

    # Remove tabs ONLY
    text = re.sub(r"\t+", " ", text)

    # IMPORTANT:
    # DO NOT collapse all spaces globally

    # Remove excessive blank lines only
    text = re.sub(r"\n{2,}", "\n", text)

    return text.strip()

# TABLE EXTRACTION

def extract_structured_tables(pdf_path, fields):

    all_rows = []

    structured_table_found = False

    previous_headers = None

    try:

        with pdfplumber.open(pdf_path) as pdf:

            for page in pdf.pages:

                tables = page.extract_tables()

                for table in tables:

                    if not table or len(table) < 1:
                        continue

                    first_row = table[0]

                    # CLEAN HEADERS
                    current_headers = []

                    for h in first_row:

                        if h:
                            current_headers.append(
                                str(h).strip().lower()
                            )
                        else:
                            current_headers.append("")

                    # CHECK IF THIS PAGE HAS REAL HEADERS
                    matched_fields = 0

                    for field in fields:

                        field_lower = field.lower()

                        for header in current_headers:

                            if field_lower in header:

                                matched_fields += 1
                                break

                    # CASE 1:
                    # PAGE HAS ACTUAL HEADERS

                    if matched_fields >= 2:

                        headers = current_headers

                        previous_headers = headers

                        data_rows = table[1:]

                        structured_table_found = True

                    # CASE 2:
                    # CONTINUATION PAGE WITHOUT HEADERS

                    elif previous_headers:

                        headers = previous_headers

                        data_rows = table

                    else:
                        continue

                    # PROCESS ROWS

                    for row_data in data_rows:

                        row = {}

                        for field in fields:

                            field_lower = field.lower()

                            for i, header in enumerate(headers):

                                if field_lower in header:

                                    if i < len(row_data):

                                        value = row_data[i]

                                        row[field] = (
                                            str(value).strip()
                                            if value
                                            else ""
                                        )

                        # Infer Type
                        if "Type" in fields:

                            full_row = " ".join(
                                [
                                    str(x)
                                    for x in row_data
                                    if x
                                ]
                            )

                            if "Type" not in row:
                                row["Type"] = infer_type(
                                    full_row
                                )

                        # Remove blank rows
                        non_empty = any(
                            str(v).strip()
                            for v in row.values()
                        )

                        if non_empty:
                            all_rows.append(row)

    except Exception as e:

        st.error(f"Table Extraction Error: {e}")

    if not structured_table_found:
        return pd.DataFrame()

    return pd.DataFrame(all_rows)

# SPLIT LABEL-STYLE RECORDS

def split_records(text):

    lines = text.split("\n")

    records = []

    current_record = []

    current_sr_no = ""

    for i, line in enumerate(lines):

        line_clean = line.strip()

        # TRUE RECORD START
        # Line contains ONLY serial number
        # Example:
        # 1
        # 2

        if re.match(r"^\d+$", line_clean):

            # Save previous record
            if current_record:

                block = "\n".join(current_record)

                records.append({
                    "Sr No": current_sr_no,
                    "Raw_Text": block
                })

            current_sr_no = line_clean

            current_record = []

            continue

        # ADD CONTENT

        if current_sr_no:

            current_record.append(line_clean)

    # LAST RECORD

    if current_record:

        block = "\n".join(current_record)

        records.append({
            "Sr No": current_sr_no,
            "Raw_Text": block
        })

    return records

# MULTI-LINE FIELD EXTRACTION

def extract_field(block, field):

    lines = block.split("\n")

    capturing = False

    value_lines = []

    for i, line in enumerate(lines):

        line_clean = line.strip()

        # FIELD START

        pattern = rf"^{re.escape(field)}\s*:\s*(.*)"

        match = re.search(
            pattern,
            line_clean,
            re.IGNORECASE
        )

        if match:

            capturing = True

            first_value = match.group(1).strip()

            # Add same-line value if exists
            if first_value:
                value_lines.append(first_value)

            continue

        # CONTINUE CAPTURE

        if capturing:

            # Stop ONLY if another proper field begins
            if ":" in line_clean:

                left_side = line_clean.split(":")[0].strip()

                # Likely another field
                if len(left_side.split()) <= 5:

                    break

            # Stop only if true next record begins
            if re.match(
                r"^\d+\s+[A-Za-z]",
                line_clean
            ):
                break

            # Append continuation lines
            if line_clean:

                value_lines.append(line_clean)

    # FINAL CLEANING

    final_value = " ".join(value_lines)

    final_value = re.sub(
        r"\s+",
        " ",
        final_value
    ).strip()

    return final_value

# TYPE INFERENCE - Can/ Should expand in future if required

def infer_type(block):

    block_lower = block.lower()

    if "veterinary" in block_lower:
        return "Veterinary"

    if "export" in block_lower:
        return "Export"

    if "consumption" in block_lower:
        return "Consumption"

    return ""

# PROCESS LABEL-STYLE RECORDS

def process_records(records, fields):

    final_rows = []

    for record in records:

        block = record["Raw_Text"]

        row = {
            "Sr No": record["Sr No"]
        }

        # EXTRACT FIELDS

        for field in fields:

            field = field.strip()

            if field.lower() == "type":

                row[field] = infer_type(block)

            else:

                row[field] = extract_field(
                    block,
                    field
                )

        # IMPORTANT:
        # REMOVE EMPTY/GARBAGE ROWS

        has_actual_data = False

        for field in fields:

            value = str(
                row.get(field, "")
            ).strip()

            if value:
                has_actual_data = True
                break

        # Keep ONLY meaningful rows
        if has_actual_data:

            final_rows.append(row)

    return pd.DataFrame(final_rows)

# STREAMLIT UI

uploaded_file = st.file_uploader(
    "Upload PDF",
    type=["pdf"]
)

fields_input = st.text_input(
    "Enter fields to extract",
    placeholder="Name, Brand, Type"
)

# MAIN PROCESS

if uploaded_file and fields_input:

    if st.button("Extract Data"):

        with st.spinner("Processing PDF..."):

            # SAVE TEMP PDF

            with tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".pdf"
            ) as tmp:

                tmp.write(uploaded_file.read())

                pdf_path = tmp.name

            # EXTRACT TEXT

            text = extract_text_from_pdf(pdf_path)

            cleaned_text = clean_text(text)

            # USER FIELDS

            fields = [
                f.strip()
                for f in fields_input.split(",")
            ]

            # TRY TABLE EXTRACTION FIRST

            table_df = extract_structured_tables(
                pdf_path,
                fields
            )

            # TABLE PDF

            if not table_df.empty:

                st.success(
                    "Structured table detected"
                )

                extracted_df = table_df

            # LABEL-STYLE PDF

            else:

                st.success(
                    "Label-style records detected"
                )

                records = split_records(
                    cleaned_text
                )

                extracted_df = process_records(
                    records,
                    fields
                )

            # KEEP ONLY REQUIRED COLUMNS

            final_columns = []

            if "Sr No" in extracted_df.columns:
                final_columns.append("Sr No")

            for field in fields:

                if field in extracted_df.columns:
                    final_columns.append(field)

            extracted_df = extracted_df[
                final_columns
            ]

            # DISPLAY OUTPUT

            st.subheader("Extracted Data")

            st.dataframe(
                extracted_df,
                width='stretch'
            )

            # EXCEL DOWNLOAD

            output = BytesIO()

            with pd.ExcelWriter(
                output,
                engine="openpyxl"
            ) as writer:

                extracted_df.to_excel(
                    writer,
                    index=False
                )

            st.download_button(
                label="Download Excel",
                data=output.getvalue(),
                file_name="Extracted_data.xlsx",
                mime=(
                    "application/vnd.openxmlformats-"
                    "officedocument.spreadsheetml.sheet"
                )
            )