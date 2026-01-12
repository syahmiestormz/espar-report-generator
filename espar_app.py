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

def parse_dashboard(file):
    """Extracts Pass Rate and Student Count from Dashboard CSV."""
    try:
        # Read the file. Dashboard usually has metadata at top.
        # We search for the header row containing 'Total Students'
        content = file.getvalue().decode('utf-8', errors='ignore')
        lines = content.split('\n')
        
        header_row = -1
        for i, line in enumerate(lines):
            if "Total Students" in line and "Pass Rate" in line:
                header_row = i
                break
        
        if header_row != -1:
            file.seek(0) # Reset file pointer
            df = pd.read_csv(file, skiprows=header_row)
            # Clean empty rows
            if 'Total Students' in df.columns:
                df = df.dropna(subset=['Total Students']).reset_index(drop=True)
                if not df.empty:
                    try:
                        # Extract values (assuming first row of data is correct)
                        total_students = pd.to_numeric(df['Total Students'].iloc[0], errors='coerce')
                        
                        # Handle Pass Rate (could be 0.8 or 80)
                        # Check various potential column names for Pass Rate
                        pass_col = [c for c in df.columns if "Pass Rate" in c]
                        if pass_col:
                            raw_rate = df[pass_col[0]].iloc[0]
                            pass_rate = pd.to_numeric(raw_rate, errors='coerce')
                            if pd.notna(pass_rate):
                                if pass_rate <= 1.0: 
                                    pass_rate *= 100
                                return total_students, pass_rate
                    except Exception as e:
                         # Log internal error but don't crash
                         print(f"Error extracting values in dashboard: {e}")
                         pass

    except Exception as e:
        st.error(f"Error parsing Dashboard {file.name}: {e}")
    return 0, 0

def parse_plo(file):
    """Extracts PLO scores from Table 3 CSV."""
    plo_scores = {}
    try:
        content = file.getvalue().decode('utf-8', errors='ignore')
        lines = content.split('\n')
        
        # Find header row with PLO labels
        header_row = -1
        for i, line in enumerate(lines):
            if "PLO 1" in line or "PLO1" in line:
                header_row = i
                break
                
        if header_row != -1:
            file.seek(0) # Reset file pointer
            df = pd.read_csv(file, skiprows=header_row)
            
            # Find the "Achievement" or "Average" row
            # Sometimes it's labeled 'Achievement', sometimes 'Average'
            target_row = None
            
            # Check if dataframe is empty or columns are missing
            if df.empty or df.shape[1] < 2:
                 return plo_scores

            # Search in the first column for the label
            # We iterate through the first column to find the row index
            first_col_name = df.columns[0]
            
            for idx, row_val in df[first_col_name].items():
                if isinstance(row_val, str) and ("Achievement" in row_val or "Average" in row_val):
                    target_row = df.iloc[idx]
                    break
            
            if target_row is not None:
                # Iterate columns to find PLO headers and matching values
                for col in df.columns:
                    if "PLO" in str(col):
                        clean_plo = str(col).strip() # e.g., "PLO 1"
                        try:
                            val = pd.to_numeric(target_row[col], errors='coerce')
                            if pd.notna(val) and val > 0: # Only count if assessed
                                if val <= 1.0: val *= 100 # Normalize to %
                                plo_scores[clean_plo] = val
                        except:
                            pass
    except Exception as e:
        st.error(f"Error parsing PLO {file.name}: {e}")
    return plo_scores

# --- Main App Interface ---

st.title("📊 ESPAR Report Generator")
st.markdown("""
**Instructions:**
1. Drag and drop all your semester CSV files here (Dashboards, Table 3 PLO, CRR, etc.).
2. The app will group them by course code and generate the report text.
""")

uploaded_files = st.file_uploader("Upload CSV Files", accept_multiple_files=True, type=['csv'])

if uploaded_files:
    # Data Containers
    course_data = {} # { 'CODE': {'pass_rate': 80, 'students': 10, 'plo': {}} }
    
    # 1. Process Files
    for uploaded_file in uploaded_files:
        code = extract_course_code(uploaded_file.name)
        if code not in course_data:
            course_data[code] = {'students': 0, 'pass_rate': 0, 'plo': {}, 'has_dashboard': False}
            
        # Identify File Type
        fname = uploaded_file.name
        
        if "Dashboard" in fname or "CRR" in fname:
            students, rate = parse_dashboard(uploaded_file)
            # Dashboard is authoritative for pass rates
            if students > 0 or rate > 0:
                course_data[code]['students'] = students
                course_data[code]['pass_rate'] = rate
                course_data[code]['has_dashboard'] = True
                
        elif "Table 3" in fname or "PLO" in fname:
            # Re-read file pointer
            uploaded_file.seek(0)
            plos = parse_plo(uploaded_file)
            course_data[code]['plo'] = plos

    # 2. Aggregation Logic
    df_courses = []
    all_plo_scores = {} # { 'PLO 1': [80, 90, ...], ... }
    
    total_students_cohort = 0
    passed_students_cohort = 0
    
    for code, data in course_data.items():
        # Calculate stats
        # Ensure we have numeric data before calculation
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
            'Fail Rate': 100 - rate if rate <= 100 else 0, # Sanity check
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
        # Sanity check for negative sizes
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
        
        # High Failure Logic
        high_fail_df = results_df[results_df['Fail Rate'] > 15]
        
    # --- REPORT GENERATION SECTION ---
    
    st.markdown("---")
    st.header("📝 Generated Report Text")
    st.info("Copy and paste these sections directly into your ESPAR Word document.")

    # Calculate PLO Averages
    plo_averages = {k: sum(v)/len(v) for k, v in all_plo_scores.items()}
    sorted_plos = sorted(plo_averages.items(), key=lambda x: x[0])
    
    # Identify Strengths/Weaknesses
    strength_text = "Students showed consistent performance across core modules."
    weakness_text = "No critical failure rates observed."
    
    # Strength: Highest PLO or 100% Pass courses
    if plo_averages:
        best_plo = max(plo_averages, key=plo_averages.get)
        strength_text = f"Students performed best in **{best_plo}** (Average: {plo_averages[best_plo]:.1f}%), indicating strong achievement in this domain."
        
    full_pass_courses = results_df[results_df['Pass Rate'] == 100]['Course Code'].tolist()
    if full_pass_courses:
        strength_text += f" Additionally, a **100% pass rate** was recorded in subjects such as {', '.join(full_pass_courses[:3])}."

    # Weakness: High Failure
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
    st.info("Waiting for files... Please upload CSVs.")