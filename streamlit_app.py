import os
import pandas as pd
from datetime import datetime
import streamlit as st

# Ensure required directories exist
os.makedirs('./uploads', exist_ok=True)

def extract_keywords(files, keywords):
    results = []
    total_files = len(files)
    files_processed = 0
    sheets_processed = 0
    rows_processed = 0
    matches_found = 0

    progress_bar = st.progress(0)
    progress_text = st.empty()

    for i, file_name in enumerate(files, start=1):
        file_path = os.path.join('./uploads', file_name.name)
        with open(file_path, "wb") as f:
            f.write(file_name.getbuffer())

        files_processed += 1
        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                sheets_processed += 1
                df = xls.parse(sheet_name).fillna("")  # Fill empty cells
                if df.shape[1] > 1:  # Ensure at least two columns exist
                    for index, value in df.iloc[:, 1].items():  # Second column (B)
                        rows_processed += 1
                        if isinstance(value, str) and any(keyword in value.lower() for keyword in keywords):
                            matched_keyword = next(kw for kw in keywords if kw in value.lower())
                            right_cells = df.iloc[index, 2:5].fillna("").tolist()  # Columns C, D, E
                            results.append({
                                "File Name": file_name.name,
                                "Sheet Name": sheet_name,
                                "Row Number": index + 1,
                                "Origin Keyword": matched_keyword,
                                "B Column Content": value,
                                "C Column Content": right_cells[0] if len(right_cells) > 0 else "",
                                "D Column Content": right_cells[1] if len(right_cells) > 1 else "",
                                "E Column Content": right_cells[2] if len(right_cells) > 2 else "",
                            })
                            matches_found += 1

        except Exception as e:
            st.error(f"Error processing {file_name.name}: {e}")
            continue

        # Update progress
        progress_text.text(f"Processing file {i}/{total_files}: {file_name.name}")
        progress_bar.progress(i / total_files)

    # Save results to a temporary file
    if results:
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        temp_file = f"extracted_{timestamp}.xlsx"
        pd.DataFrame(results).to_excel(temp_file, index=False, engine="openpyxl")
        return temp_file, files_processed, sheets_processed, rows_processed, matches_found
    else:
        return None, files_processed, sheets_processed, rows_processed, matches_found


# Streamlit App Interface
st.title("📊 Excel Keyword Extractor")
st.markdown("**Upload Excel files, specify keywords, and extract matching rows.**")

# Input Section
keywords = st.text_area(
    "Enter keywords (comma-separated):",
    help="Enter keywords to search for in the Excel files. Separate multiple keywords with commas."
)

keyword_file = st.file_uploader(
    "Or upload a .txt file with keywords:",
    type=["txt"],
    help="Upload a text file containing keywords (one keyword per line)."
)

uploaded_files = st.file_uploader(
    "Drag and drop Excel files here, or click to upload.",
    type=["xls", "xlsx"],
    accept_multiple_files=True,
    help="You can upload multiple Excel files."
)

if keyword_file:
    keywords = keyword_file.read().decode("utf-8").strip().replace("\n", ", ")

keywords_list = [kw.lower().strip() for kw in keywords.split(",") if kw.strip()]

if st.button("Start Extraction"):
    if not uploaded_files or not keywords_list:
        st.warning("Please provide all inputs!")
    else:
        st.info("Extraction started... Please wait.")
        result_file, files_processed, sheets_processed, rows_processed, matches_found = extract_keywords(
            uploaded_files, keywords_list
        )

        # Show results
        st.success("Extraction completed! 🎉")
        st.write(f"**Files Processed**: {files_processed}")
        st.write(f"**Sheets Processed**: {sheets_processed}")
        st.write(f"**Rows Processed**: {rows_processed}")
        st.write(f"**Matches Found**: {matches_found}")

        if result_file:
            st.write("**Download Results:**")
            with open(result_file, "rb") as f:
                st.download_button(
                    label="Download Extracted File",
                    data=f,
                    file_name=os.path.basename(result_file),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            os.remove(result_file)  # Cleanup the temporary file after download
        else:
            st.warning("No matches found.")
