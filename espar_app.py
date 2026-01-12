import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io
import re

# Page Configuration
st.set_page_config(page_title="ESPAR Report Generator", layout="wide")

# --- CSS for Printing ---
st.markdown("""
<style>
    @media print {
        .stButton {display: none;}
        .css-15zrgzn {display: none;}
        .stSidebar {display: none;}
    }
</style>
""", unsafe_allow_html=True)

# --- Helper Functions ---

def extract_course_code(filename):
    """Attempts to extract a pattern like DMIM1033 from filename."""
    match = re.search(r'([A-Z]{4}\d{4})', filename)
    if match:
        return match.group(1)
    return "Unknown_Course"

def load_file(uploaded_file):
    """Reads CSV or Excel file into a dataframe or returns raw content for parsing."""
    fname = uploaded_file.name
    if fname.endswith('.xlsx'):
        try:
            return pd.read_excel(uploaded_file, header=None) # Read raw first
        except Exception as e:
            st.error(f"Error reading Excel {fname}: {e}")
            return None
    else:
        # For CSV, we might need raw text parsing or pandas reading
        return uploaded_file

def parse_dashboard(file, is_excel=False, excel_df=None):
    """Extracts Pass Rate and Student Count from Dashboard."""
    try:
        total_students = 0
        pass_rate = 0
        
        if is_excel and excel_df is not None:
            # Iterate through rows to find "Total Students"
            df = excel_df
            header_row = -1
            # Search for header in first few columns/rows
            for r_idx, row in df.iterrows():
                row_str = row.astype(str).str.cat(sep=' ')
                if "Total Students" in row_str and "Pass Rate" in row_str:
                    header_row = r_idx
                    break
            
            if header_row != -1:
                # Reload with correct header
                df.columns = df.iloc[header_row]
                df = df.iloc[header_row+1:].reset_index(drop=True)
                df = df.dropna(subset=['Total Students'])
                
                if not df.empty:
                    total_students = pd.to_numeric(df['Total Students'].iloc[0], errors='coerce')
                    # Find Pass Rate col
                    pass_col = [c for c in df.columns if "Pass Rate" in str(c)]
                    if pass_col:
                        raw_rate = df[pass_col[0]].iloc[0]
                        pass_rate = pd.to_numeric(raw_rate, errors='coerce')
                        if pd.notna(pass_rate):
                            if pass_rate <= 1.0: pass_rate *= 100
                            return total_students, pass_rate

        else: # CSV Processing
            content = file.getvalue().decode('utf-8', errors='ignore')
            lines = content.split('\n')
            
            header_row = -1
            for i, line in enumerate(lines):
                if "Total Students" in line and "Pass Rate" in line:
                    header_row = i
                    break
            
            if header_row != -1:
                file.seek(0)
                df = pd.read_csv(file, skiprows=header_row)
                df = df.dropna(subset=['Total Students']).reset_index(drop=True)
                if not df.empty:
                    total_students = pd.to_numeric(df['Total Students'].iloc[0], errors='coerce')
                    pass_col = [c for c in df.columns if "Pass Rate" in c]
                    if pass_col:
                        raw_rate = df[pass_col[0]].iloc[0]
                        pass_rate = pd.to_numeric(raw_rate, errors='coerce')
                        if pd.notna(pass_rate):
                            if pass_rate <= 1.0: pass_rate *= 100
                            return total_students, pass_rate

    except Exception as e:
        print(f"Error parsing dashboard: {e}")
    return 0, 0

def parse_plo(file, is_excel=False, excel_df=None):
    """Extracts PLO scores from Table 3."""
    plo_scores = {}
    try:
        df = None
        if is_excel and excel_df is not None:
             # Find header row
            raw_df = excel_df
            header_row = -1
            for r_idx, row in raw_df.iterrows():
                row_str = row.astype(str).str.cat(sep=' ')
                if "PLO 1" in row_str or "PLO1" in row_str:
                    header_row = r_idx
                    break
            
            if header_row != -1:
                raw_df.columns = raw_df.iloc[header_row]
                df = raw_df.iloc[header_row+1:].reset_index(drop=True)
                
        else: # CSV
            content = file.getvalue().decode('utf-8', errors='ignore')
            lines = content.split('\n')
            header_row = -1
            for i, line in enumerate(lines):
                if "PLO 1" in line or "PLO1" in line:
                    header_row = i
                    break
            if header_row != -1:
                file.seek(0)
                df = pd.read_csv(file, skiprows=header_row)

        if df is not None:
            # Find Achievement Row
            target_row = None
            
            # Check first column for label
            first_col = df.columns[0]
            for idx, val in df[first_col].items():
                if isinstance(val, str) and ("Achievement" in val or "Average" in val):
                    target_row = df.iloc[idx]
                    break
            
            if target_row is not None:
                for col in df.columns:
                    if "PLO" in str(col):
                        clean_plo = str(col).strip()
                        try:
                            val = pd.to_numeric(target_row[col], errors='coerce')
                            if pd.notna(val) and val > 0:
                                if val <= 1.0: val *= 100
                                plo_scores[clean_plo] = val
                        except:
                            pass

    except Exception as e:
        print(f"Error parsing PLO: {e}")
    return plo_scores

# --- Main App Interface ---

st.title("📊 ESPAR Report Generator")
st.markdown("""
**Instructions:**
1. Drag and drop all your semester files (CSV or Excel) here.
2. The app will group them by course code and generate the report text.
""")

# UPDATED: Accept 'xlsx' now
uploaded_files = st.file_uploader("Upload Course Files", accept_multiple_files=True, type=['csv', 'xlsx'])

if uploaded_files:
    # Data Containers
    course_data = {} 
    
    # 1. Process Files
    for uploaded_file in uploaded_files:
        code = extract_course_code(uploaded_file.name)
        if code not in course_data:
            course_data[code] = {'students': 0, 'pass_rate': 0, 'plo': {}, 'has_dashboard': False}
            
        fname = uploaded_file.name
        is_excel = fname.endswith('.xlsx')
        
        # Helper: Read excel once if needed
        excel_df = None
        if is_excel:
            excel_df = load_file(uploaded_file)

        if "Dashboard" in fname or "CRR" in fname:
            students, rate = parse_dashboard(uploaded_file, is_excel, excel_df)
            if students > 0 or rate > 0:
                course_data[code]['students'] = students
                course_data[code]['pass_rate'] = rate
                course_data[code]['has_dashboard'] = True
                
        elif "Table 3" in fname or "PLO" in fname:
            if not is_excel: uploaded_file.seek(0)
            plos = parse_plo(uploaded_file, is_excel, excel_df)
            course_data[code]['plo'] = plos

    # 2. Aggregation Logic
    df_courses = []
    all_plo_scores = {} 
    
    total_students_cohort = 0
    passed_students_cohort = 0
    
    for code, data in course_data.items():
        # Calculate stats
        studs = pd.to_numeric(data.get('students', 0), errors='coerce')
        if pd.isna(studs): studs = 0
        
        rate = pd.to_numeric(data.get('pass_rate', 0), errors='coerce')
        if pd.isna(rate): rate = 0

        passed = round(studs * (rate / 100))
        total_students_cohort += studs
        passed_students_cohort += passed
        
        # PLO aggregation
        for plo_name, score in data['plo'].items():
            if plo_name not in all_plo_scores:
                all_plo_scores[plo_name] = []
            all_plo_scores[plo_name].append(score)
            
        df_courses.append({
            'Course Code': code,
            'Pass Rate': rate,
            'Fail Rate': 100 - rate if rate <= 100 else 0, 
            'Students': studs
        })

    # Create DataFrames
    results_df = pd.DataFrame(df_courses)
    
    # Avoid div by zero
    if total_students_cohort > 0:
        overall_pass_rate = (passed_students_cohort / total_students_cohort) * 100
    elif not results_df.empty:
        overall_pass_rate = results_df['Pass Rate'].mean()
    else:
        overall_pass_rate = 0
        
    # --- VISUALIZATION SECTION ---
    
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.subheader("Cohort Performance")
        st.metric("Overall Pass Rate", f"{overall_pass_rate:.1f}%")
        st.metric("Total Students", f"{int(total_students_cohort)}")
        
        # Pie Chart
        fig, ax = plt.subplots()
        labels = ['Pass', 'Fail']
        sizes = [overall_pass_rate, 100 - overall_pass_rate]
        sizes = [max(0, s) for s in sizes]
        
        colors = ['#66b3ff', '#ff9999']
        explode = (0.1, 0)
        
        if sum(sizes) > 0:
            ax.pie(sizes, explode=explode, labels=labels, colors=colors, autopct='%1.1f%%',
                   shadow=True, startangle=90)
            ax.axis('equal') 
            st.pyplot(fig)
        else:
            st.warning("No data for chart")

    with col2:
        st.subheader("Course Breakdown")
        st.dataframe(results_df.style.format({"Pass Rate": "{:.1f}%", "Fail Rate": "{:.1f}%"}), use_container_width=True)
        
        high_fail_df = results_df[results_df['Fail Rate'] > 15]
        
    # --- REPORT GENERATION SECTION ---
    
    st.markdown("---")
    st.header("📝 Generated Report Text")
    st.info("Copy and paste these sections directly into your ESPAR Word document.")

    # Calculate PLO Averages
    plo_averages = {k: sum(v)/len(v) for k, v in all_plo_scores.items()}
    sorted_plos = sorted(plo_averages.items(), key=lambda x: x[0])
    
    # Strength/Weakness Logic
    strength_text = "Students showed consistent performance across core modules."
    weakness_text = "No critical failure rates observed."
    
    if plo_averages:
        best_plo = max(plo_averages, key=plo_averages.get)
        strength_text = f"Students performed best in **{best_plo}** (Average: {plo_averages[best_plo]:.1f}%), indicating strong achievement in this domain."
        
    full_pass_courses = results_df[results_df['Pass Rate'] == 100]['Course Code'].tolist()
    if full_pass_courses:
        strength_text += f" Additionally, a **100% pass rate** was recorded in subjects such as {', '.join(full_pass_courses[:3])}."

    if not high_fail_df.empty:
        failed_list = [f"{row['Course Code']} ({row['Fail Rate']:.1f}% Fail)" for _, row in high_fail_df.iterrows()]
        weakness_text = f"High failure rates were observed in **{', '.join(failed_list)}**, suggesting students struggled with the specific requirements of these courses."
    elif plo_averages:
        worst_plo = min(plo_averages, key=plo_averages.get)
        if plo_averages[worst_plo] < 50:
            weakness_text = f"Students struggled in **{worst_plo}** (Average: {plo_averages[worst_plo]:.1f}%), falling below the target KPI."

    # 1.0 Executive Summary
    st.subheader("1.0 EXECUTIVE SUMMARY")
    exec_summary = f"""
* **Overall Status:** Satisfactory. The cohort achieved an **Overall Pass Rate of {overall_pass_rate:.1f}%**.
* **Key Strength:** {strength_text}
* **Key Weakness:** {weakness_text}
"""
    st.text_area("1.0 Executive Summary", value=exec_summary, height=150)

    # 3.1 CLO Analysis
    st.subheader("3.1 CLO Analysis (Course Level)")
    
    fail_list_formatted = ""
    if not high_fail_df.empty:
        for _, row in high_fail_df.iterrows():
            fail_list_formatted += f"* **{row['Course Code']}** (Failure Rate: **{row['Fail Rate']:.1f}%**)\n"
    else:
        fail_list_formatted = "* No courses exceeded the 15% failure threshold."

    clo_analysis = f"""
* **Overall Pass Rate:** **{overall_pass_rate:.1f}%**
* **High Failure Rate Courses (>15% Failure):**
{fail_list_formatted}

> *[Insert the Pie Chart generated on the left here]*
"""
    st.text_area("3.1 CLO Analysis", value=clo_analysis, height=200)

    # 3.2 PLO Analysis
    st.subheader("3.2 PLO Analysis (Programme Level)")
    
    plo_rows = ""
    for plo, score in sorted_plos:
        status = "ACHIEVED" if score >= 50 else "ATTENTION REQUIRED"
        plo_rows += f"| **{plo}** | **{score:.1f}** | 50 | {status} |\n"
        
    plo_table = f"""
| PLO Domain | Achievement (%) | Target (%) | Status |
| :--- | :--- | :--- | :--- |
{plo_rows}
"""
    st.text_area("3.2 PLO Analysis Table", value=plo_table, height=300)

    # 5.0 Conclusion
    st.subheader("5.0 CONCLUSION")
    conclusion = f"""
The academic session concluded with an **Overall Pass Rate of {overall_pass_rate:.1f}%**. 
{strength_text}
However, {weakness_text.lower()} 
These weaknesses are being addressed through the Strategic CQI Action Plan. 
The curriculum remains relevant and compliant with MQA standards.
"""
    st.text_area("5.0 Conclusion", value=conclusion, height=150)

else:
    st.info("Waiting for files... Please upload CSVs or Excel files.")
