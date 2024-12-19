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
                                "Fayl Adı": file_name.name,
                                "Vərəq Adı": sheet_name,
                                "Sətir Nömrəsi": index + 1,
                                "Açar Söz": matched_keyword,
                                "B Sütun Məzmunu": value,
                                "C Sütun Məzmunu": right_cells[0] if len(right_cells) > 0 else "",
                                "D Sütun Məzmunu": right_cells[1] if len(right_cells) > 1 else "",
                                "E Sütun Məzmunu": right_cells[2] if len(right_cells) > 2 else "",
                            })
                            matches_found += 1

        except Exception as e:
            st.error(f"Xəta {file_name.name} faylında: {e}")
            continue

        # Update progress
        progress_text.text(f"{i}/{total_files} fayl işlənir: {file_name.name}")
        progress_bar.progress(i / total_files)

    # Save results to a temporary file
    if results:
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        temp_file = f"çıxarış_{timestamp}.xlsx"
        pd.DataFrame(results).to_excel(temp_file, index=False, engine="openpyxl")
        return temp_file, files_processed, sheets_processed, rows_processed, matches_found
    else:
        return None, files_processed, sheets_processed, rows_processed, matches_found


# Streamlit App Interface
# Add custom CSS for design
st.markdown(
    """
    <style>
    /* Background color */
    .main {
        background-color: #f0f4f5;
    }

    /* Style for headers */
    h1 {
        color: #4CAF50;
        font-size: 2.5em;
        text-align: center;
    }

    h3 {
        color: #333;
        text-align: left;
        margin-top: 20px;
    }

    /* Customize file uploader */
    .stFileUploader {
        border: 2px dashed #4CAF50 !important;
        border-radius: 10px;
        padding: 10px;
    }

    /* Center alignment */
    .stButton button {
        background-color: #4CAF50 !important;
        color: white !important;
        border-radius: 10px !important;
        width: 100% !important;
    }

    /* Progress bar customization */
    .stProgress {
        height: 25px;
    }

    /* Footer alignment */
    footer {
        visibility: hidden;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📊 Excel Açar Söz Ekstraktoru")
st.markdown("**Excel fayllarınızı yükləyin, açar sözləri daxil edin və uyğun olan sətirləri çıxarın.**")

# Layout for instructions
st.markdown("""
    ### İstifadə üçün:
    1. **Açar sözləri daxil edin** və ya .txt faylını yükləyin. 🟢
    2. **Excel fayllarınızı yükləyin** və ya sürüşdürərək buraya atın. 🟡
    3. `Ekstraksiya Başlasın` düyməsini sıxın və nəticələri yükləyin. 🔵
""")

# Create two columns for inputs
col1, col2 = st.columns(2)

with col1:
    keywords = st.text_area(
        "Açar sözləri daxil edin (vergüllə ayrılmış):",
        help="Açar sözləri axtarmaq üçün yazın, vergüllə ayırın."
    )

with col2:
    keyword_file = st.file_uploader(
        "Və ya açar sözləri olan .txt faylı yükləyin:",
        type=["txt"],
        help="Bir sətirdə bir açar söz olan .txt faylı yükləyin."
    )

uploaded_files = st.file_uploader(
    "Excel fayllarınızı yükləyin və ya sürüşdürərək buraya atın:",
    type=["xls", "xlsx"],
    accept_multiple_files=True,
    help="Bir neçə Excel faylı yükləyə bilərsiniz."
)

if keyword_file:
    keywords = keyword_file.read().decode("utf-8").strip().replace("\n", ", ")

keywords_list = [kw.lower().strip() for kw in keywords.split(",") if kw.strip()]

# Start extraction button
start_button = st.button("Ekstraksiya Başlasın")
if start_button:
    if not uploaded_files or not keywords_list:
        st.warning("Bütün məlumatları daxil edin!")
    else:
        st.info("Ekstraksiya başladı... Gözləyin.")
        result_file, files_processed, sheets_processed, rows_processed, matches_found = extract_keywords(
            uploaded_files, keywords_list
        )

        # Show results
        st.success("Ekstraksiya tamamlandı! 🎉")
        st.write(f"**İşlənən Fayllar:** {files_processed}")
        st.write(f"**İşlənən Vərəqlər:** {sheets_processed}")
        st.write(f"**İşlənən Sətirlər:** {rows_processed}")
        st.write(f"**Uyğunluqlar Tapıldı:** {matches_found}")

        if result_file:
            st.write("**Nəticələri Yükləyin:**")
            with open(result_file, "rb") as f:
                st.download_button(
                    label="Çıxarış Faylını Yükləyin",
                    data=f,
                    file_name=os.path.basename(result_file),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            os.remove(result_file)  # Cleanup the temporary file after download
        else:
            st.warning("Heç bir uyğunluq tapılmadı.")
