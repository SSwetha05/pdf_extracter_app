import streamlit as st
import pandas as pd
import pdfplumber
import pytesseract
pytesseract.pytesseract.tesseract_cmd = "tesseract"
import re
import tempfile

from pdf2image import convert_from_path
from io import BytesIO


st.set_page_config(
    page_title="PDF to Excel Extractor",
    layout="wide"
)

st.title("PDF to Excel Extractor")

st.write("""
Upload scanned or normal PDFs and extract structured Excel data.
""")


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


def clean_text(text):

    text = text.replace("\r", "\n")

    text = text.replace("|", " ")

    text = text.replace(";", ":")

    text = re.sub(r"\t+", " ", text)

    text = re.sub(r"\n{2,}", "\n", text)

    return text.strip()


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

                    current_headers = []

                    for h in first_row:

                        if h:
                            current_headers.append(
                                str(h).strip().lower()
                            )
                        else:
                            current_headers.append("")

                    matched_fields = 0

                    for field in fields:

                        field_lower = field.lower()

                        for header in current_headers:

                            if field_lower in header:

                                matched_fields += 1
                                break


                    if matched_fields >= 2:

                        headers = current_headers

                        previous_headers = headers

                        data_rows = table[1:]

                        structured_table_found = True


                    elif previous_headers:

                        headers = previous_headers

                        data_rows = table

                    else:
                        continue


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


def split_records(text):

    lines = text.split("\n")

    records = []

    current_record = []

    current_sr_no = ""

    for i, line in enumerate(lines):

        line_clean = line.strip()

        if re.match(r"^\d+$", line_clean):

            if current_record:

                block = "\n".join(current_record)

                records.append({
                    "Sr No": current_sr_no,
                    "Raw_Text": block
                })

            current_sr_no = line_clean

            current_record = []

            continue

        if current_sr_no:

            current_record.append(line_clean)


    if current_record:

        block = "\n".join(current_record)

        records.append({
            "Sr No": current_sr_no,
            "Raw_Text": block
        })

    return records


def extract_field(block, field):

    lines = block.split("\n")

    capturing = False

    value_lines = []

    for i, line in enumerate(lines):

        line_clean = line.strip()

        pattern = rf"^{re.escape(field)}\s*:\s*(.*)"

        match = re.search(
            pattern,
            line_clean,
            re.IGNORECASE
        )

        if match:

            capturing = True

            first_value = match.group(1).strip()

            if first_value:
                value_lines.append(first_value)

            continue

        if capturing:

            if ":" in line_clean:

                left_side = line_clean.split(":")[0].strip()

                if len(left_side.split()) <= 5:

                    break

            if re.match(
                r"^\d+\s+[A-Za-z]",
                line_clean
            ):
                break

            if line_clean:

                value_lines.append(line_clean)


    final_value = " ".join(value_lines)

    final_value = re.sub(
        r"\s+",
        " ",
        final_value
    ).strip()

    return final_value


def infer_type(block):

    block_lower = block.lower()

    if "veterinary" in block_lower:
        return "Veterinary"

    if "export" in block_lower:
        return "Export"

    if "consumption" in block_lower:
        return "Consumption"

    return ""


def process_records(records, fields):

    final_rows = []

    for record in records:

        block = record["Raw_Text"]

        row = {
            "Sr No": record["Sr No"]
        }


        for field in fields:

            field = field.strip()

            if field.lower() == "type":

                row[field] = infer_type(block)

            else:

                row[field] = extract_field(
                    block,
                    field
                )


        has_actual_data = False

        for field in fields:

            value = str(
                row.get(field, "")
            ).strip()

            if value:
                has_actual_data = True
                break

        if has_actual_data:

            final_rows.append(row)

    return pd.DataFrame(final_rows)

# Streamlit UI

uploaded_file = st.file_uploader(
    "Upload PDF",
    type=["pdf"]
)

fields_input = st.text_input(
    "Enter fields to extract",
    placeholder="Name, Brand, Type"
)

# Main Process

if uploaded_file and fields_input:

    if st.button("Extract Data"):

        with st.spinner("Processing PDF..."):


            with tempfile.NamedTemporaryFile(
                delete=False,
                suffix=".pdf"
            ) as tmp:

                tmp.write(uploaded_file.read())

                pdf_path = tmp.name


            text = extract_text_from_pdf(pdf_path)

            cleaned_text = clean_text(text)


            fields = [
                f.strip()
                for f in fields_input.split(",")
            ]


            table_df = extract_structured_tables(
                pdf_path,
                fields
            )


            if not table_df.empty:

                st.success(
                    "Structured table detected"
                )

                extracted_df = table_df


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


            final_columns = []

            if "Sr No" in extracted_df.columns:
                final_columns.append("Sr No")

            for field in fields:

                if field in extracted_df.columns:
                    final_columns.append(field)

            extracted_df = extracted_df[
                final_columns
            ]


            st.subheader("Extracted Data")

            st.dataframe(
                extracted_df,
                width='stretch'
            )


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