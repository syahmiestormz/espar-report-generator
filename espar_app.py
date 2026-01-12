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

def get_smart_recommendation(issue_text, course_name=""):
    """
    Acts as a 'Smart AI' to generate pedagogical actions based on keywords 
    in the issue text or course context.
    """
    text = (str(issue_text) + " " + str(course_name)).lower()
    
    recommendations = {
        "attendance": "Implement strict attendance monitoring and issue warning letters to chronic absentees.",
        "late": "Review submission deadlines and enforce strict penalties for late submissions to encourage discipline.",
        "submit": "Review submission deadlines and enforce strict penalties for late submissions to encourage discipline.",
        "theory": "Introduce more interactive visual aids and real-world case studies to explain abstract theoretical concepts.",
        "concept": "Introduce more interactive visual aids and real-world case studies to explain abstract theoretical concepts.",
        "calculation": "Conduct remedial drills focusing specifically on step-by-step calculation methods.",
        "math": "Conduct remedial drills focusing specifically on step-by-step calculation methods.",
        "programming": "Organize 'Code Clinics' where students can get one-on-one debugging help from seniors or lecturers.",
        "coding": "Organize 'Code Clinics' where students can get one-on-one debugging help from seniors or lecturers.",
        "drawing": "Host extra studio sessions with live demonstrations to improve technique application.",
        "sketching": "Host extra studio sessions with live demonstrations to improve technique application.",
        "design": "Incorporate critique sessions (critique) earlier in the semester to provide formative feedback.",
        "visual": "Provide more examples of high-quality visual analysis to guide student expectations.",
        "communication": "Integrate mandatory presentation components in assessments to build confidence and skills.",
        "english": "Encourage usage of English in class discussions and recommend support workshops.",
        "group": "Implement a peer-evaluation mechanism to ensure fair contribution in group projects.",
        "project": "Break down the final project into smaller milestones to monitor progress more effectively.",
        "software": "Conduct specific lab tutorials focusing on software tools and shortcuts.",
        "basic": "Conduct 'Back-to-Basics' revision classes to strengthen fundamental understanding."
    }
    
    for key, action in recommendations.items():
        if key in text:
            return action
            
    # Default generic professional response if no keywords match
    return "Conduct focused revision classes targeting the specific weak topics identified in the assessment."

def extract_course_code(filename):
    """Attempts to extract a pattern like DMIM1033 from filename."""
    match = re.search(r'([A-Z]{4}\d{4})', filename)
    if match:
        return match.group(1)
    return "Unknown_Course"

def find_val_in_df(df, keyword):
    """Searches a dataframe for a keyword and returns the cell coordinates."""
    for r_idx, row in df.iterrows():
        row_str = row.astype(str).str.cat(sep=' ')
        if keyword in row_str:
            return r_idx
    return -1

def extract_dashboard_metrics(df):
    """Extracts stats from a dataframe that looks like the Dashboard."""
    total_students = 0
    pass_rate = 0
    
    header_row = find_val_in_df(df, "Total Students")
    
    if header_row != -1:
        try:
            df.columns = df.iloc[header_row]
            df = df.iloc[header_row+1:].reset_index(drop=True)
            df = df.dropna(subset=['Total Students'])
            
            if not df.empty:
                val_stud = df['Total Students'].iloc[0]
                total_students = pd.to_numeric(val_stud, errors='coerce')
                
                pass_cols = [c for c in df.columns if "Pass Rate" in str(c)]
                if pass_cols:
                    val_rate = df[pass_cols[0]].iloc[0]
                    raw_rate = pd.to_numeric(val_rate, errors='coerce')
                    if pd.notna(raw_rate):
                        if raw_rate <= 1.0: 
                            pass_rate = raw_rate * 100
                        else:
                            pass_rate = raw_rate
                            
                return total_students, pass_rate
        except Exception as e:
            print(f"Error extracting metrics: {e}")
            
    return 0, 0

def extract_plo_metrics(df):
    """Extracts PLO scores from a dataframe that looks like Table 3."""
    plo_scores = {}
    
    header_row = -1
    for r_idx, row in df.iterrows():
        row_str = row.astype(str).str.cat(sep=' ')
        if "PLO 1" in row_str or "PLO1" in row_str:
            header_row = r_idx
            break
            
    if header_row != -1:
        try:
            df.columns = df.iloc[header_row]
            df = df.iloc[header_row+1:].reset_index(drop=True)
            
            target_row = None
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
            print(f"Error extracting PLO: {e}")
            
    return plo_scores

def extract_cqi_issues(df):
    """
    Extracts user-entered CQI issues from Table 2 - CLO Analysis.
    Only includes items where status is FAIL or Score < 50%.
    """
    cqi_list = [] # [{'issue': '...', 'action': '...', 'evidence': '...'}, ...]
    
    # Locate header row containing 'Issue' and 'Suggestion'
    header_row = -1
    for r_idx, row in df.iterrows():
        row_str = row.astype(str).str.cat(sep=' ')
        if "Issue" in row_str and "Suggestion" in row_str:
            header_row = r_idx
            break
            
    if header_row != -1:
        try:
            df.columns = df.iloc[header_row]
            df = df.iloc[header_row+1:].reset_index(drop=True)
            
            # Identify columns
            col_issue = [c for c in df.columns if "Issue" in str(c)][0]
            col_action = [c for c in df.columns if "Suggestion" in str(c)][0]
            col_evidence = [c for c in df.columns if "Audit" in str(c) or "Evidence" in str(c)]
            col_evidence = col_evidence[0] if col_evidence else None
            
            # Identify Pass/Fail or Score columns for filtering
            col_status = [c for c in df.columns if "Pass" in str(c) or "Met" in str(c)]
            col_score = [c for c in df.columns if "%" in str(c) or "Score" in str(c)]
            
            # Iterate rows to find non-empty issues
            for idx, row in df.iterrows():
                
                # --- FILTER LOGIC: Only process if FAILED or < 50% ---
                is_fail = False
                
                # Check 1: Explicit "Fail" or "No" in status column
                if col_status:
                    status_val = str(row[col_status[0]]).lower()
                    if "fail" in status_val or "no" in status_val:
                        is_fail = True
                        
                # Check 2: Score < 50% (0.5)
                if not is_fail and col_score:
                    try:
                        score_val = pd.to_numeric(row[col_score[0]], errors='coerce')
                        # Handle 0.45 (45%) or 45 (45%)
                        if score_val <= 1.0: score_val *= 100
                        if score_val < 50:
                            is_fail = True
                    except:
                        pass
                
                # Only proceed if it is a failure case
                if is_fail:
                    issue_text = str(row[col_issue]) if pd.notna(row[col_issue]) else ""
                    action_text = str(row[col_action]) if pd.notna(row[col_action]) else ""
                    
                    # Check if meaningful text (not '0', 'nan', or empty)
                    if len(issue_text) > 3 and issue_text != "0" and issue_text.lower() != "nan":
                        
                        # === INTELLIGENT AUTO-FILL ===
                        if len(action_text) < 4 or action_text == "0":
                            action_text = get_smart_recommendation(issue_text)
                        
                        evidence = str(row[col_evidence]) if col_evidence and pd.notna(row[col_evidence]) else "Course Audit Report"
                        cqi_list.append({
                            'issue': issue_text,
                            'action': action_text,
                            'evidence': evidence
                        })
                    
        except Exception as e:
            print(f"Error extracting CQI: {e}")
            
    return cqi_list

# --- Main App Interface ---

st.title("📊 ESPAR Report Generator")
st.markdown("""
**Instructions:**
1. Drag and drop all your semester **Excel (.xlsx)** or **CSV** files here.
2. The app will group them by course code and generate the report text.
""")

uploaded_files = st.file_uploader("Upload Course Files", accept_multiple_files=True, type=['csv', 'xlsx'])

if uploaded_files:
    # Data Containers
    course_data = {} 
    
    # 1. Process Files
    for uploaded_file in uploaded_files:
        code = extract_course_code(uploaded_file.name)
        if code not in course_data:
            course_data[code] = {'students': 0, 'pass_rate': 0, 'plo': {}, 'cqi': [], 'has_dashboard': False}
            
        fname = uploaded_file.name
        
        # === EXCEL WORKBOOK LOGIC ===
        if fname.endswith('.xlsx'):
            try:
                # Read ALL sheets
                xls = pd.read_excel(uploaded_file, sheet_name=None, header=None)
                
                # 1. Dashboard
                dash_df = None
                for sheet_name, sheet_df in xls.items():
                    if "Dashboard" in sheet_name or "CRR" in sheet_name:
                        dash_df = sheet_df
                        break
                if dash_df is not None:
                    studs, rate = extract_dashboard_metrics(dash_df)
                    if studs > 0:
                        course_data[code]['students'] = studs
                        course_data[code]['pass_rate'] = rate
                        course_data[code]['has_dashboard'] = True
                        
                # 2. PLO
                plo_df = None
                for sheet_name, sheet_df in xls.items():
                    if "Table 3" in sheet_name or "PLO" in sheet_name:
                        plo_df = sheet_df
                        break
                if plo_df is not None:
                    plos = extract_plo_metrics(plo_df)
                    if plos:
                        course_data[code]['plo'].update(plos)

                # 3. CQI (Table 2)
                cqi_df = None
                for sheet_name, sheet_df in xls.items():
                    if "Table 2" in sheet_name or "CLO" in sheet_name:
                        cqi_df = sheet_df
                        break
                if cqi_df is not None:
                    cqi_items = extract_cqi_issues(cqi_df)
                    if cqi_items:
                        course_data[code]['cqi'].extend(cqi_items)

            except Exception as e:
                st.error(f"Error reading Excel file {fname}: {e}")

        # === CSV LOGIC ===
        else: 
            if "Dashboard" in fname or "CRR" in fname:
                try:
                    df = pd.read_csv(uploaded_file, header=None)
                    studs, rate = extract_dashboard_metrics(df)
                    if studs > 0:
                        course_data[code]['students'] = studs
                        course_data[code]['pass_rate'] = rate
                        course_data[code]['has_dashboard'] = True
                except: pass
            elif "Table 3" in fname or "PLO" in fname:
                try:
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, header=None)
                    plos = extract_plo_metrics(df)
                    if plos:
                        course_data[code]['plo'].update(plos)
                except: pass
            elif "Table 2" in fname:
                try:
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, header=None)
                    cqi_items = extract_cqi_issues(df)
                    if cqi_items:
                        course_data[code]['cqi'].extend(cqi_items)
                except: pass

    # 2. Aggregation Logic
    df_courses = []
    all_plo_scores = {} 
    
    total_students_cohort = 0
    passed_students_cohort = 0
    
    for code, data in course_data.items():
        studs = pd.to_numeric(data.get('students', 0), errors='coerce')
        if pd.isna(studs): studs = 0
        
        rate = pd.to_numeric(data.get('pass_rate', 0), errors='coerce')
        if pd.isna(rate): rate = 0

        passed = round(studs * (rate / 100))
        total_students_cohort += studs
        passed_students_cohort += passed
        
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

    results_df = pd.DataFrame(df_courses)
    
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
        if not results_df.empty:
            st.dataframe(results_df.style.format({"Pass Rate": "{:.1f}%", "Fail Rate": "{:.1f}%"}), use_container_width=True)
            high_fail_df = results_df[results_df['Fail Rate'] > 15]
        else:
            high_fail_df = pd.DataFrame()
        
    # --- REPORT GENERATION SECTION ---
    
    st.markdown("---")
    st.header("📝 Generated Report Text")
    st.info("Copy and paste these sections directly into your ESPAR Word document.")

    plo_averages = {k: sum(v)/len(v) for k, v in all_plo_scores.items()}
    sorted_plos = sorted(plo_averages.items(), key=lambda x: x[0])
    
    # Text Generation Logic
    strength_text = "Students showed consistent performance across core modules."
    weakness_text = "No critical failure rates observed."
    
    if plo_averages:
        best_plo = max(plo_averages, key=plo_averages.get)
        strength_text = f"Students performed best in **{best_plo}** (Average: {plo_averages[best_plo]:.1f}%), indicating strong achievement in this domain."
        
    if not results_df.empty:
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

    # 1.0 & 3.1 & 3.2 (Standard Sections)
    st.subheader("1.0 EXECUTIVE SUMMARY")
    exec_summary = f"""
* **Overall Status:** Satisfactory. The cohort achieved an **Overall Pass Rate of {overall_pass_rate:.1f}%**.
* **Key Strength:** {strength_text}
* **Key Weakness:** {weakness_text}
"""
    st.text_area("1.0 Executive Summary", value=exec_summary, height=150)

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

    # 4.0 Strategic CQI Action Plan (SMART AI VERSION + FILTERED)
    st.subheader("4.0 STRATEGIC CQI ACTION PLAN (AI-Assisted)")
    
    cqi_rows = ""
    
    # 1. Use user-entered CQI from Excel first (FILTERED BY FAIL)
    has_specific_cqi = False
    for code, data in course_data.items():
        if data['cqi']:
            for item in data['cqi']:
                cqi_rows += f"| {item['issue']} ({code}) | {item['action']} | In Progress | {item['evidence']} |\n"
            has_specific_cqi = True
            
    # 2. Fallback if no specific entries found (AUTO-GENERATE using Smart Logic)
    if not has_specific_cqi:
        if not high_fail_df.empty:
            for _, row in high_fail_df.iterrows():
                # Try to guess context from code or just generic
                action = get_smart_recommendation("theory calculation", row['Course Code'])
                cqi_rows += f"| High Failure Rate in {row['Course Code']} ({row['Fail Rate']:.1f}%) | {action} | Completed | Attendance List (Appendix A) |\n"
        
        if plo_averages:
            worst_plo = min(plo_averages, key=plo_averages.get)
            if plo_averages[worst_plo] < 50:
                action = get_smart_recommendation(worst_plo)
                cqi_rows += f"| Low Performance in {worst_plo} (Avg: {plo_averages[worst_plo]:.1f}%) | {action} | In Progress | New Course Outline (Appendix B) |\n"
            
    if not cqi_rows:
        cqi_rows = "| No critical failures observed. | Maintain current teaching strategies. | Completed | Semester Report |\n"

    cqi_plan = f"""
| Issue / Weakness Identified | Action Taken | Status | Evidence Reference (Bukti) |
| :--- | :--- | :--- | :--- |
{cqi_rows}
"""
    st.text_area("4.0 Strategic CQI Action Plan", value=cqi_plan, height=250)

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
