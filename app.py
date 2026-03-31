import streamlit as st
import pdfplumber
import pandas as pd
import io
import random

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title="Internal Marks Processing System",
    layout="wide",
    page_icon="📊"
)

# ---------------- CUSTOM CSS ----------------
st.markdown("""
<style>
body { background-color: #f5f7fa; }
h1 { text-align: center; color: #2c3e50; }
.stButton>button {
    background-color: #4CAF50;
    color: white;
    border-radius: 8px;
    height: 45px;
}
.stDownloadButton>button {
    background-color: #2196F3;
    color: white;
    border-radius: 8px;
    height: 45px;
}
</style>
""", unsafe_allow_html=True)

# ---------------- HEADER ----------------
st.markdown("""
<h1>📊 Internal Marks Processing System</h1>
<p style='text-align:center; font-size:18px; color:gray;'>
Upload PDF → Preview → Download Excel
</p>
<hr>
""", unsafe_allow_html=True)

# ---------------- USER INPUT ----------------
st.markdown("### 🏫 Institution Details")
colA, colB = st.columns(2)

with colA:
    college_name = st.text_input("College Name", "Your College Name")

with colB:
    department_name = st.text_input("Department Name", "Department of Computer Science")

# ---------------- SIDEBAR ----------------
st.sidebar.header("⚙️ Settings")
random_range = st.sidebar.slider("Random Variation (±)", 0, 10, 5)

# ---------------- FILE UPLOAD ----------------
st.subheader("📁 Upload File")
uploaded_file = st.file_uploader("Choose Internal Marks PDF", type=["pdf"])

# ---------------- VALID PREFIX ----------------
valid_regd_prefixes = [
    "25BSCS", "25BSAI", "25BSMT",
    "24BSCS", "24BSAI", "24BSMT",
    "23BSCS", "23BSAI", "23BSMT"
]

# ---------------- MAIN LOGIC ----------------
if uploaded_file:

    st.info("⏳ Processing PDF... Please wait")

    all_student_data = []
    subject_name = "Internal Assessment"

    try:
        with pdfplumber.open(uploaded_file) as pdf:

            first_page_text = pdf.pages[0].extract_text()

            for line in first_page_text.split('\n'):
                if "Subject : " in line:
                    subject_name = line.replace("Subject : ", "").strip()

            for page in pdf.pages:
                table = page.extract_table()

                if table:
                    for row in table[1:]:

                        clean_row = [
                            str(item).replace('\n', ' ').strip() if item else ""
                            for item in row[:4]
                        ]

                        if len(clean_row) > 1 and any(
                            clean_row[1].startswith(prefix)
                            for prefix in valid_regd_prefixes
                        ):

                            s_no = clean_row[0]
                            regd_no = clean_row[1]
                            student_name = clean_row[2]
                            total_marks_str = clean_row[3]

                            try:
                                total_marks = int(float(total_marks_str))

                                total_1_plus_2 = int(round(total_marks / 0.40))
                                scaled_to_40 = int(round(total_1_plus_2 * 0.40))

                                half_total = total_1_plus_2 // 2
                                random_offset = random.randint(-random_range, random_range)

                                total_1 = int(min(half_total + random_offset, 50))
                                total_2 = int(total_1_plus_2 - total_1)

                                # Total 1
                                ncc1 = 10
                                a1 = int(round(total_1 * 0.20))
                                m1 = int(round(total_1 * 0.40))
                                sq1 = int(total_1 - a1 - m1 - ncc1)
                                if sq1 < 0:
                                    m1 += sq1
                                    sq1 = 0

                                # Total 2
                                ncc2 = 10
                                a2 = int(round(total_2 * 0.20))
                                m2 = int(round(total_2 * 0.40))
                                sq2 = int(total_2 - a2 - m2 - ncc2)
                                if sq2 < 0:
                                    m2 += sq2
                                    sq2 = 0

                            except:
                                continue

                            all_student_data.append([
                                s_no, regd_no, student_name,
                                a1, sq1, ncc1, m1, total_1,
                                a2, sq2, ncc2, m2, total_2,
                                total_1_plus_2, scaled_to_40
                            ])

        # ---------------- DATAFRAME ----------------
        columns = [
            "S.No.", "Regd.No.", "Student Name",
            "Assignment 1", "Seminar/Quiz 1", "NCC/NSS 1", "Mid I", "Total 1",
            "Assignment 2", "Seminar/Quiz 2", "NCC/NSS 2", "Mid II", "Total 2",
            "Total 1 + Total 2", "Scaled to 40"
        ]

        df = pd.DataFrame(all_student_data, columns=columns)

        if df.empty:
            st.error("❌ No valid student data found in PDF")
            st.stop()

        st.success("✅ Processing Completed Successfully!")

        # ---------------- METRICS ----------------
        col1, col2, col3 = st.columns(3)
        col1.metric("👨‍🎓 Students", len(df))
        col2.metric("📘 Subject", subject_name)
        col3.metric("📊 Max Marks", df["Scaled to 40"].max())

        st.markdown("---")

        # ---------------- TABLE ----------------
        st.subheader("📋 Student Marks Preview")
        st.dataframe(df, use_container_width=True, height=450)

        # ---------------- EXCEL OUTPUT ----------------
        output = io.BytesIO()

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, startrow=3)

            workbook  = writer.book
            worksheet = writer.sheets['Sheet1']

            header_format = workbook.add_format({
                'bold': True,
                'align': 'center',
                'font_size': 14
            })

            # Merge and write headers
            worksheet.merge_range('A1:O1', college_name, header_format)
            worksheet.merge_range('A2:O2', department_name, header_format)
            worksheet.merge_range('A3:O3', f"Subject: {subject_name}", header_format)

        safe_name = subject_name.replace(" ", "_")

        st.markdown("### 📥 Download Result")

        st.download_button(
            label="Download Excel File",
            data=output.getvalue(),
            file_name=f"{safe_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"⚠️ Error: {str(e)}")

else:
    st.warning("📌 Please upload a PDF file to begin")
