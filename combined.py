import streamlit as st
import pandas as pd
import re
from docx import Document
from io import BytesIO
from copy import deepcopy
from docx.enum.text import WD_COLOR_INDEX

# === Helper Functions ===

def classify(percentile):
    try:
        if isinstance(percentile, str) and ">" in percentile:
            percentile = float(percentile.replace(">", ""))
        elif percentile == "-" or pd.isna(percentile):
            return "-"
        else:
            percentile = float(percentile)
    except:
        return "-"

    if percentile <= 1:
        return "Extremely Low"
    elif 2 <= percentile <= 8:
        return "Unusually Low"
    elif 9 <= percentile <= 24:
        return "Low Average"
    elif 25 <= percentile <= 74:
        return "Average"
    elif 75 <= percentile <= 90:
        return "High Average"
    elif 91 <= percentile <= 97:
        return "Unusually High"
    elif percentile >= 98:
        return "Extremely High"
    else:
        return "-"

def format_percentile_with_suffix(percentile):
    try:
        if isinstance(percentile, str) and ">" in percentile:
            percentile = float(percentile.replace(">", ""))
        elif percentile == "-" or pd.isna(percentile):
            return "-"
        else:
            percentile = float(percentile)
    except:
        return "-"

    if percentile.is_integer():
        integer_part = int(percentile)
    else:
        decimal_first_digit = int(str(percentile).split(".")[1][0])
        integer_part = decimal_first_digit

    if 10 <= integer_part % 100 <= 20:
        suffix = 'th'
    else:
        last_digit = integer_part % 10
        if last_digit == 1:
            suffix = 'st'
        elif last_digit == 2:
            suffix = 'nd'
        elif last_digit == 3:
            suffix = 'rd'
        else:
            suffix = 'th'

    if percentile.is_integer():
        return f"{int(percentile)}{suffix}"
    else:
        return f"{percentile}{suffix}"

def replace_placeholders(doc, lookup):
    def replace_in_runs(runs, lookup):
        n = len(runs)
        i = 0
        while i < n:
            found = False
            combined_text = ''
            for j in range(i, n):
                combined_text += runs[j].text
                for key, value in lookup.items():
                    placeholder = f"{{{{{key}}}}}"
                    if placeholder in combined_text:
                        new_text = combined_text.replace(placeholder, value)
                        idx = 0
                        for k in range(i, j + 1):
                            run_length = len(runs[k].text)
                            runs[k].text = new_text[idx:idx+run_length]
                            idx += run_length
                        found = True
                        break
                if found:
                    break
            i = j + 1 if found else i + 1

    for para in doc.paragraphs:
        replace_in_runs(para.runs, lookup)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_in_runs(para.runs, lookup)

def superscript_suffixes(doc):
    pattern = re.compile(r'(\d+(?:\.\d+)?)(st|nd|rd|th)')

    def copy_font_settings(source_run, target_run):
        target_run.font.bold = source_run.font.bold
        target_run.font.italic = source_run.font.italic
        target_run.font.underline = source_run.font.underline
        target_run.font.size = source_run.font.size
        target_run.font.name = source_run.font.name
        if source_run.font.color.rgb:
            target_run.font.color.rgb = source_run.font.color.rgb

    def process_runs(paragraph):
        new_runs = []
        for run in paragraph.runs:
            text = run.text
            last_end = 0
            matches = list(pattern.finditer(text))

            if not matches:
                new_runs.append((text, False, run))
            else:
                for match in matches:
                    start, end = match.span()
                    if start > last_end:
                        new_runs.append((text[last_end:start], False, run))
                    new_runs.append((match.group(1), False, run))
                    new_runs.append((match.group(2), True, run))
                    last_end = end
                if last_end < len(text):
                    new_runs.append((text[last_end:], False, run))

        for run in paragraph.runs:
            run.text = ''

        for text, is_super, original_run in new_runs:
            if text == '':
                continue
            new_run = paragraph.add_run(text)
            copy_font_settings(original_run, new_run)
            if is_super:
                new_run.font.superscript = True

    for para in doc.paragraphs:
        process_runs(para)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    process_runs(para)

def delete_rows_with_dash(doc):
    for table in doc.tables:
        rows_to_delete = []
        for row_idx, row in enumerate(table.rows):
            for cell in row.cells:
                if cell.text.strip() == "#":
                    rows_to_delete.append(row_idx)
                    break
        for row_idx in sorted(rows_to_delete, reverse=True):
            tbl = table._tbl
            tr = table.rows[row_idx]._tr
            tbl.remove(tr)

def delete_rows_with_unfilled_placeholders(doc):
    pattern = re.compile(r"\{\{.*?\}\}")
    for table in doc.tables:
        rows_to_delete = []
        for row_idx, row in enumerate(table.rows):
            for cell in row.cells:
                if pattern.search(cell.text):
                    rows_to_delete.append(row_idx)
                    break
        for row_idx in sorted(rows_to_delete, reverse=True):
            tbl = table._tbl
            tr = table.rows[row_idx]._tr
            tbl.remove(tr)

def highlight_unfilled_placeholders(doc):
    placeholder_pattern = re.compile(r"\{\{.*?\}\}")
    missing_symbol_pattern = re.compile(r"#")

    def highlight_placeholder_in_runs(runs):
        combined_text = ''
        run_indices = []

        for idx, run in enumerate(runs):
            combined_text += run.text
            run_indices.append(idx)

        matches = list(placeholder_pattern.finditer(combined_text)) + list(missing_symbol_pattern.finditer(combined_text))
        if not matches:
            return

        current_pos = 0
        for idx in run_indices:
            run = runs[idx]
            text_len = len(run.text)
            run_end_pos = current_pos + text_len

            for match in matches:
                match_start, match_end = match.span()
                if match_start < run_end_pos and match_end > current_pos:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW

            current_pos = run_end_pos

    for para in doc.paragraphs:
        highlight_placeholder_in_runs(para.runs)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    highlight_placeholder_in_runs(para.runs)

# === Streamlit App ===

st.title("📄 WIAT-4 and WISC Report Combiner")

uploaded_doc = st.file_uploader("Upload WIAT-4 Report (.docx)", type="docx")
uploaded_wisc = st.file_uploader("Upload WISC Report (.docx)", type="docx")

if uploaded_doc and uploaded_wisc:
    gender_selection = st.radio(
        "Select WIAT Report Gender Template:",
        ("Male", "Female")
    )

    if st.button("Generate Combined Report"):
        # === WIAT Processing ===
        input_doc = Document(uploaded_doc)
        template_path = "template_male.docx" if gender_selection == "Male" else "template_female.docx"
        template_doc = Document(template_path)

        target_table_indices = [2, 3, 4, 5, 6, 7, 8, 9, 10]
        ae_combined = pd.DataFrame()

        for i in target_table_indices:
            if i < len(input_doc.tables):
                table = input_doc.tables[i]
                data = []
                for row in table.rows:
                    data.append([cell.text.strip() for cell in row.cells])
                df = pd.DataFrame(data)

                if df.shape[0] > 1:
                    df.columns = df.iloc[0]
                    df = df.drop(index=0).reset_index(drop=True)

                if df.shape[1] >= 5:
                    ae_df = df.iloc[:, [0, 4]].copy()
                    ae_df.columns = ['Name', 'Percentile']
                    ae_df['Name'] = ae_df['Name'].str.replace(r'[^A-Za-z\s]', '', regex=True).str.strip()
                    ae_combined = pd.concat([ae_combined, ae_df], ignore_index=True)

        if not ae_combined.empty:
            ae_combined.drop_duplicates(subset='Name', inplace=True)
            ae_combined["Classification"] = ae_combined["Percentile"].apply(classify)
            ae_combined["Percentile*"] = ae_combined["Percentile"].apply(format_percentile_with_suffix)
            ae_combined = ae_combined.replace("-", "#")

        lookup = {}
        for idx, row in ae_combined.iterrows():
            name = row['Name'].strip()
            lookup[f"{name} Classification"] = row['Classification']
            lookup[f"{name} Percentile"] = str(row['Percentile']).strip()
            lookup[f"{name} Percentile*"] = str(row['Percentile*']).strip()

        replace_placeholders(template_doc, lookup)
        superscript_suffixes(template_doc)
        delete_rows_with_dash(template_doc)
        delete_rows_with_unfilled_placeholders(template_doc)
        highlight_unfilled_placeholders(template_doc)

        # === WISC Processing ===
        input_wisc_doc = Document(uploaded_wisc)
        wisc_template = Document("wisc_template.docx")

        wisc_combined = pd.DataFrame()

        for i, table in enumerate(input_wisc_doc.tables):
            data = []
            for row in table.rows:
                data.append([cell.text.strip() for cell in row.cells])
            df = pd.DataFrame(data)

       
            
            if df.shape[0] > 1:
                df.columns = df.iloc[0]
                df = df.drop(index=0).reset_index(drop=True)
                st.write(f"📊 Raw WISC Table {i+1}")
                st.dataframe(df)

                if df.shape[1] >= 5:
                    if i == 5:  # Table 6 (0-indexed)
                        ae_df = df.iloc[:, [1, 4]].copy()
                    elif df.shape[1] >= 6:
                        ae_df = df.iloc[:, [1, 5]].copy()
                    else:
                        continue  # Not enough columns
                
                    ae_df.columns = ['Name', 'Percentile']
                    ae_df['Name'] = ae_df['Name'].str.replace(r'[^A-Za-z\s]', '', regex=True).str.strip()
                    st.write(f"🧠 A/E Data from WISC Table {i+1}")
                    st.dataframe(ae_df)
                
                    wisc_combined = pd.concat([wisc_combined, ae_df], ignore_index=True)

        if not wisc_combined.empty:
            wisc_combined.drop_duplicates(subset='Name', inplace=True)
            wisc_combined["Classification"] = wisc_combined["Percentile"].apply(classify)
            wisc_combined["Percentile*"] = wisc_combined["Percentile"].apply(format_percentile_with_suffix)
            wisc_combined = wisc_combined.replace("-", "#")

        wisc_lookup = {}
        for idx, row in wisc_combined.iterrows():
            name = row['Name'].strip()
            wisc_lookup[f"{name} Classification"] = row['Classification']
            wisc_lookup[f"{name} Percentile"] = str(row['Percentile']).strip()
            wisc_lookup[f"{name} Percentile*"] = str(row['Percentile*']).strip()

        replace_placeholders(wisc_template, wisc_lookup)
        superscript_suffixes(wisc_template)
        delete_rows_with_dash(wisc_template)
        delete_rows_with_unfilled_placeholders(wisc_template)
        highlight_unfilled_placeholders(wisc_template)

        # === Append WISC to WIAT ===
        for element in wisc_template.element.body:
            template_doc.element.body.append(element)

        # === Output ===
        output = BytesIO()
        template_doc.save(output)
        output.seek(0)

        st.success("✅ Combined document generated successfully!")

        st.download_button(
            label="📥 Download Combined Report",
            data=output,
            file_name="combined_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
