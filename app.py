import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import statsmodels.api as sm
import urllib.request
import json
from streamlit_oauth import OAuth2Component

# ==========================================
# PAGE CONFIGURATION & CUSTOM CSS
# ==========================================
# This MUST be the first Streamlit command
st.set_page_config(page_title="Impact Analytics Dashboard", layout="wide", initial_sidebar_state="expanded")

# Custom CSS for better KPI cards and UI polishing
st.markdown("""
<style>
    [data-testid="stMetricValue"] {
        font-size: 2rem;
        font-weight: 700;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 24px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: transparent;
        border-radius: 4px 4px 0px 0px;
        padding-top: 10px;
        padding-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🔒 AUTHENTICATION GATEKEEPER
# ==========================================
# Securely fetch credentials from Streamlit Secrets
try:
    CLIENT_ID = st.secrets["GOOGLE_CLIENT_ID"]
    CLIENT_SECRET = st.secrets["GOOGLE_CLIENT_SECRET"]
except FileNotFoundError:
    st.error("Missing `.streamlit/secrets.toml` file or Streamlit Cloud Secrets. Please ensure your Google Client ID and Secret are configured.")
    st.stop()

AUTHORIZE_URL = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN_URL = "https://oauth2.googleapis.com/token"
REVOKE_TOKEN_URL = "https://oauth2.googleapis.com/revoke"

# Initialize session states
if "logged_in_email" not in st.session_state:
    st.session_state["logged_in_email"] = None
if "user_first_name" not in st.session_state:
    st.session_state["user_first_name"] = "User"

# If not logged in, show ONLY the login screen and STOP execution
if not st.session_state["logged_in_email"]:
    col1, col2, col3 = st.columns()
    with col2:
        st.write("") # Spacing
        st.write("")
        try:
            st.image("evidyaloka_logo.png", width=300)
        except:
            st.empty()
            
        st.markdown("<h2 style='text-align: center; color: #0094c9;'>Staff Analytics Portal</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center;'>Please sign in with your @evidyaloka.org email to access the dashboard.</p>", unsafe_allow_html=True)
        st.markdown("---")
        
        # Create OAuth Object
        oauth2 = OAuth2Component(CLIENT_ID, CLIENT_SECRET, AUTHORIZE_URL, TOKEN_URL, REVOKE_TOKEN_URL)
        
        # Create Login Button
        result = oauth2.authorize_button(
            name="Sign in with Google",
            icon="https://upload.wikimedia.org/wikipedia/commons/5/53/Google_%22G%22_Logo.svg",
            redirect_uri="https://ev-assessments.streamlit.app", 
            scope="openid email profile",
            key="google_login",
            use_container_width=True
        )
        
        if result and "token" in result:
            # Safely extract the ID token
            id_token = result["token"]["id_token"]
            
            # Failsafe: If the server somehow returns the token wrapped in a list, extract the string
            if isinstance(id_token, list):
                id_token = id_token
                
            # Bulletproof approach: Ask Google's official endpoint to decode and validate the token
            verify_url = f"https://oauth2.googleapis.com/tokeninfo?id_token={id_token}"
            
            try:
                with urllib.request.urlopen(verify_url) as response:
                    user_info = json.loads(response.read().decode())
                    
                # Save BOTH the email and the First Name (given_name)
                st.session_state["logged_in_email"] = user_info.get("email") 
                st.session_state["user_first_name"] = user_info.get("given_name", "User")
                st.rerun()
                
            except Exception as e:
                st.error(f"⚠️ Error verifying login with Google: {e}")
                st.stop()
            
    # CRITICAL: This stops the rest of your dashboard and sidebar from loading for unauthorized users!
    st.stop()

# ==========================================
# 🚀 MAIN DASHBOARD (ONLY REACHED IF LOGGED IN)
# ==========================================
# Because this is placed AFTER st.stop(), it will never show on the login screen
st.title("📈 Impact Analytics Dashboard")
st.markdown("<p style='color: gray; font-size: 1.1em;'>Comprehensive Baseline vs. Endline Performance Assessment</p>", unsafe_allow_html=True)

# ==========================================
# SIDEBAR: LOGO, USER INFO & FILTERS
# ==========================================
with st.sidebar:
    try:
        st.image("evidyaloka_logo.png", width=273)
    except:
        st.warning("⚠️ Logo not found.")
    
    # Use the first name grabbed from Google!
    st.success(f"👤 **Logged in as:** {st.session_state['user_first_name']}")
    
    if st.button("Sign Out", use_container_width=True):
        # Clear the states and reload the page
        st.session_state["logged_in_email"] = None
        st.session_state["user_first_name"] = "User"
        st.rerun()
    st.markdown("---")

# ==========================================
# DATA LOADING ENGINE
# ==========================================
@st.cache_data
def load_and_prep_data(file_path):
    common_cols = ['State', 'Centre Name', 'Donor', 'Subject', 'Grade', 'Student ID', 'Gender', 'Total Marks', 'Obtained Marks', 'Category', 'Academic Year']
    dfs_to_concat = []

    try:
        xls = pd.ExcelFile(file_path)
        sheet_names = xls.sheet_names
        
        # Smart defaults (auto-detect if named appropriately)
        base_sheet = 'Baseline' if 'Baseline' in sheet_names else 0
        end_sheet = 'Endline' if 'Endline' in sheet_names else (1 if len(sheet_names) > 1 else 0)

        df_base = pd.read_excel(file_path, sheet_name=base_sheet)
        if 'Rubrics' in df_base.columns: df_base.rename(columns={'Rubrics': 'Category'}, inplace=True)
        df_base['Academic Year'] = 'Baseline'
        dfs_to_concat.append(df_base[[c for c in common_cols if c in df_base.columns]])

        df_end = pd.read_excel(file_path, sheet_name=end_sheet)
        if 'Rubrics' in df_end.columns: df_end.rename(columns={'Rubrics': 'Category'}, inplace=True)
        df_end['Academic Year'] = 'Endline'
        dfs_to_concat.append(df_end[[c for c in common_cols if c in df_end.columns]])
        
    except Exception as e:
        st.error(f"Error reading the Excel file: {e}")
        return pd.DataFrame()

    if not dfs_to_concat:
        return pd.DataFrame()

    df_combined = pd.concat(dfs_to_concat, ignore_index=True)

    # Added 'Category' to the stripping loop
    for col in ['State', 'Centre Name', 'Donor', 'Subject', 'Student ID', 'Gender', 'Category']:
        if col in df_combined.columns:
            df_combined[col] = df_combined[col].astype(str).str.strip()

    # Standardize Gender formatting to catch ALL variations (boy, BOY, bOy -> Boy)
    if 'Gender' in df_combined.columns:
        df_combined['Gender'] = df_combined['Gender'].astype(str).str.strip().str.title()

    if 'Grade' in df_combined.columns:
        df_combined['Grade'] = df_combined['Grade'].astype(str).str.replace(r'\.0$', '', regex=True)

    # Enforce R.I.S.E ordering
    if 'Category' in df_combined.columns:
        rise_order = ["Reviving", "Initiating", "Shaping", "Evolving"]
        df_combined['Category'] = pd.Categorical(df_combined['Category'], categories=rise_order, ordered=True)

    # Ensure numeric
    df_combined['Obtained Marks'] = pd.to_numeric(df_combined['Obtained Marks'], errors='coerce')

    return df_combined

# Define custom color palette for R.I.S.E Categories
COLOR_MAP = {'Baseline': '#636EFA', 'Endline': '#00CC96'}
RISE_COLORS = {"Reviving": "#f27c48", "Initiating": "#0094c9", "Shaping": "#00964d", "Evolving": "#ed1c2d"}

DATA_FILE = "BL-EL-AY-25-26-Final-AllSubjects.xlsx"

# Strictly look for the local file
if os.path.exists(DATA_FILE):
    with st.spinner('Loading and crunching numbers...'):
        df = load_and_prep_data(DATA_FILE)

    if not df.empty:
        with st.sidebar:
            st.header("🎯 Global Filters")
            
            # 1. State Filter (Displays Full Names from pristine df)
            states = ["All"] + sorted(df['State'].dropna().unique().tolist())
            selected_states = st.selectbox("Select State", states, index=0)
            
            # Pre-filter for next dropdown
            df_state_filtered = df.copy()
            if selected_states != "All": 
                df_state_filtered = df_state_filtered[df_state_filtered['State'] == selected_states]

            # 2. Donor Filter (Dependent on State)
            donors = ["All"] + sorted(df_state_filtered['Donor'].dropna().unique().tolist())
            selected_donors = st.selectbox("Select Donor", donors, index=0)
            
            # Pre-filter for next dropdown
            df_donor_filtered = df_state_filtered.copy()
            if selected_donors != "All":
                df_donor_filtered = df_donor_filtered[df_donor_filtered['Donor'] == selected_donors]

            # 3. Centre Filter (Dependent on State & Donor)
            centres = ["All"] + sorted(df_donor_filtered['Centre Name'].dropna().unique().tolist())
            selected_centres = st.selectbox("Select Centre", centres, index=0)
            
            # Pre-filter for next dropdown
            df_centre_filtered = df_donor_filtered.copy()
            if selected_centres != "All":
                df_centre_filtered = df_centre_filtered[df_centre_filtered['Centre Name'] == selected_centres]

            # 4. Subject Filter (Dependent on Centre)
            subjects = ["All"] + sorted(df_centre_filtered['Subject'].dropna().unique().tolist())
            selected_subjects = st.selectbox("Select Subject", subjects, index=0)

            # Pre-filter for next dropdown
            df_subject_filtered = df_centre_filtered.copy()
            if selected_subjects != "All":
                df_subject_filtered = df_subject_filtered[df_subject_filtered['Subject'] == selected_subjects]

            # 5. Grade Filter (Dependent on Subject, Multi-select)
            grades = sorted(df_subject_filtered['Grade'].dropna().unique().tolist())
            selected_grades = st.multiselect("Select Grade(s)", options=grades, default=grades)

            # Pre-filter for next dropdown
            df_grade_filtered = df_subject_filtered.copy()
            if selected_grades:
                df_grade_filtered = df_grade_filtered[df_grade_filtered['Grade'].isin(selected_grades)]
            else:
                df_grade_filtered = df_grade_filtered.iloc[0:0] 

            # 6. Gender Filter (Dependent on Grade, Multi-select)
            if 'Gender' in df_grade_filtered.columns:
                valid_genders = df_grade_filtered[~df_grade_filtered['Gender'].str.lower().isin(['nan', 'none', 'null', ''])].copy()
                genders = sorted(valid_genders['Gender'].unique().tolist())
                if genders:
                    selected_genders = st.multiselect("Select Gender(s)", options=genders, default=genders)
                    filtered_df = df_grade_filtered[df_grade_filtered['Gender'].isin(selected_genders)].copy()
                else:
                    filtered_df = df_grade_filtered.copy()
            else:
                filtered_df = df_grade_filtered.copy()

        if filtered_df.empty:
            st.warning("⚠️ No data available for the selected filters. Please adjust your criteria.")
        else:
            # ==========================================
            # DASHBOARD TABS
            # ==========================================
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                "📊 Executive Summary", 
                "📚 Subject Deep-Dive", 
                "🗺️ Geographic View",
                "🧑‍🎓 Student-Level Impact",
                "🚻 Gender Analysis",
                "📉 RTM Analysis"
            ])

            base_df = filtered_df[filtered_df['Academic Year'] == 'Baseline']
            end_df = filtered_df[filtered_df['Academic Year'] == 'Endline']

            # ------------------------------------------
            # TAB 1: EXECUTIVE SUMMARY
            # ------------------------------------------
            with tab1:
                st.markdown("### 🚀 High-Level Metrics")
                
                # Expanded to 5 columns to fit SD
                kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
                
                # Calculate Matched Students for the KPI
                if not base_df.empty and not end_df.empty and 'Student ID' in df.columns:
                    matched_students = len(pd.merge(base_df[['Student ID', 'Subject']].dropna(), end_df[['Student ID', 'Subject']].dropna(), on=['Student ID', 'Subject']))
                else:
                    matched_students = 0
                
                avg_base = base_df['Obtained Marks'].mean() if not base_df.empty else None
                avg_end = end_df['Obtained Marks'].mean() if not end_df.empty else None
                
                sd_base = base_df['Obtained Marks'].std() if not base_df.empty and len(base_df) > 1 else None
                sd_end = end_df['Obtained Marks'].std() if not end_df.empty and len(end_df) > 1 else None

                kpi1.metric("Matched Students", f"{matched_students:,}")
                
                if avg_base is not None and avg_end is not None:
                    kpi2.metric("Baseline Mean Score", f"{avg_base:.2f}")
                    kpi3.metric("Endline Mean Score", f"{avg_end:.2f}", delta=f"{avg_end - avg_base:.2f}")
                    
                    if sd_base is not None and sd_end is not None:
                        kpi4.metric("Endline Score SD", f"{sd_end:.2f}", delta=f"{sd_end - sd_base:.2f}", delta_color="inverse")
                    else:
                        kpi4.metric("Endline Score SD", "N/A")
                    
                    # Calculate % of students in "Evolving" (top category)
                    base_evolve = len(base_df[base_df['Category'] == 'Evolving']) / len(base_df) * 100 if len(base_df) > 0 else 0
                    end_evolve = len(end_df[end_df['Category'] == 'Evolving']) / len(end_df) * 100 if len(end_df) > 0 else 0
                    kpi5.metric("Students in 'Evolving'", f"{end_evolve:.1f}%", delta=f"{end_evolve - base_evolve:.1f}%")
                elif avg_base is not None:
                    kpi2.metric("Baseline Mean Score", f"{avg_base:.2f}")
                    kpi3.metric("Endline Mean Score", "N/A")
                    kpi4.metric("Endline Score SD", "N/A")
                    kpi5.metric("Data Status", "Awaiting Endline")
                else:
                    kpi2.metric("Baseline Mean Score", "N/A")
                    kpi3.metric("Endline Mean Score", f"{avg_end:.2f}")
                    kpi4.metric("Endline Score SD", f"{sd_end:.2f}" if sd_end else "N/A")
                    kpi5.metric("Data Status", "Endline Only")
                    
                st.info("**💡 Understanding Standard Deviation (SD):** SD measures the spread of student scores. A **decrease** (green delta) in SD means scores are becoming more clustered and consistent, indicating that the learning gap between high and low performers is closing. An **increase** (red delta) means the gap is widening.")

                st.markdown("---")
                
                col_a, col_b = st.columns(2)
                with col_a:
                    st.markdown("#### 📈 Score Distribution (Box Plot)")
                    st.caption("Visualizes the spread of scores, median, and outliers.")
                    fig_box = px.box(filtered_df, x="Academic Year", y="Obtained Marks", color="Academic Year", 
                                     color_discrete_map=COLOR_MAP, points="all")
                    fig_box.update_layout(showlegend=False, margin=dict(l=0, r=0, t=30))
                    st.plotly_chart(fig_box, width='stretch')

                with col_b:
                    st.markdown("#### 🧬 R.I.S.E Category Shift")
                    st.caption("Proportional breakdown of performance categories.")
                    cat_counts = filtered_df.groupby(['Academic Year', 'Category']).size().reset_index(name='Count')
                    cat_counts['Percentage'] = cat_counts.groupby('Academic Year')['Count'].transform(lambda x: x / x.sum() * 100)
                    cat_counts = cat_counts.sort_values(['Academic Year', 'Category'])
                    
                    # Grouped side-by-side bar chart
                    fig_rise = px.bar(cat_counts, x="Category", y="Percentage", color="Academic Year", 
                                      text=cat_counts['Percentage'].apply(lambda x: f'{x:.1f}%' if not pd.isna(x) else ''),
                                      color_discrete_map=COLOR_MAP,
                                      category_orders={"Category": ["Reviving", "Initiating", "Shaping", "Evolving"]})
                    fig_rise.update_layout(barmode='group', margin=dict(l=0, r=0, t=30),
                                           legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, title=""))
                    st.plotly_chart(fig_rise, width='stretch')


            # ------------------------------------------
            # TAB 2: SUBJECT DEEP-DIVE
            # ------------------------------------------
            with tab2:
                st.markdown("### 📚 Subject & Grade Performance (R.I.S.E. Distribution)")
                
                def get_stacked_data(df_subset):
                    if df_subset.empty or 'Grade' not in df_subset.columns:
                        return pd.DataFrame()
                    grouped = df_subset.groupby(['Grade', 'Category']).size().reset_index(name='Count')
                    grouped['Percentage'] = grouped.groupby('Grade')['Count'].transform(lambda x: x / x.sum() * 100)
                    return grouped
                
                base_stacked = get_stacked_data(base_df)
                end_stacked = get_stacked_data(end_df)
                
                sub_col1, sub_col2 = st.columns(2)
                
                with sub_col1:
                    st.markdown("#### Baseline R.I.S.E by Grade")
                    if not base_stacked.empty:
                        fig_base_grade = px.bar(base_stacked, x="Grade", y="Percentage", color="Category",
                                                color_discrete_map=RISE_COLORS, 
                                                text=base_stacked['Percentage'].apply(lambda x: f'{x:.1f}%' if x > 5 else ''),
                                                category_orders={"Category": ["Reviving", "Initiating", "Shaping", "Evolving"]})
                        fig_base_grade.update_layout(barmode='stack', yaxis_title="% of Students", margin=dict(l=0, r=0, t=30),
                                                     legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, title=""))
                        st.plotly_chart(fig_base_grade, width='stretch')
                    else:
                        st.info("No Baseline data available.")
                        
                with sub_col2:
                    st.markdown("#### Endline R.I.S.E by Grade")
                    if not end_stacked.empty:
                        fig_end_grade = px.bar(end_stacked, x="Grade", y="Percentage", color="Category",
                                               color_discrete_map=RISE_COLORS, 
                                               text=end_stacked['Percentage'].apply(lambda x: f'{x:.1f}%' if x > 5 else ''),
                                               category_orders={"Category": ["Reviving", "Initiating", "Shaping", "Evolving"]})
                        fig_end_grade.update_layout(barmode='stack', yaxis_title="% of Students", margin=dict(l=0, r=0, t=30),
                                                    legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, title=""))
                        st.plotly_chart(fig_end_grade, width='stretch')
                    else:
                        st.info("No Endline data available.")
                
                st.markdown("---")
                st.markdown("#### 🧠 Automated Insights")
                
                if not base_stacked.empty and not end_stacked.empty:
                    try:
                        base_piv = base_stacked.pivot(index='Grade', columns='Category', values='Percentage').fillna(0)
                        end_piv = end_stacked.pivot(index='Grade', columns='Category', values='Percentage').fillna(0)
                        
                        for cat in ["Reviving", "Initiating", "Shaping", "Evolving"]:
                            if cat not in base_piv.columns: base_piv[cat] = 0
                            if cat not in end_piv.columns: end_piv[cat] = 0
                            
                        common_grades = base_piv.index.intersection(end_piv.index)
                        
                        if len(common_grades) > 0:
                            diff_piv = end_piv.loc[common_grades] - base_piv.loc[common_grades]
                            
                            best_evo_grade = diff_piv['Evolving'].idxmax()
                            best_evo_val = diff_piv['Evolving'].max()
                            
                            best_rev_grade = diff_piv['Reviving'].idxmin()
                            best_rev_val = diff_piv['Reviving'].min()
                            
                            if best_evo_val > 0:
                                st.success(f"📈 **Top Excellence Growth:** Grade **{best_evo_grade}** saw the highest shift into the 'Evolving' category, increasing its top-tier share by **{best_evo_val:+.1f}** percentage points from Baseline to Endline.")
                            else:
                                st.warning("⚠️ **Excellence Alert:** No grade saw an increase in the 'Evolving' category percentage.")
                            
                            if best_rev_val < 0:
                                st.success(f"📉 **Highest Risk Reduction:** Grade **{best_rev_grade}** had the most successful intervention for struggling students, reducing its 'Reviving' (lowest tier) population by **{abs(best_rev_val):.1f}** percentage points.")
                            else:
                                st.warning("⚠️ **Risk Alert:** No grade successfully reduced their share of students in the 'Reviving' category.")
                        else:
                            st.info("Insufficient overlapping grades between Baseline and Endline to generate comparative insights.")
                    except Exception as e:
                        st.info("Not enough data variance to generate automated insights.")
                else:
                    st.info("Awaiting both Baseline and Endline data to generate comparative insights.")


            # ------------------------------------------
            # TAB 3: GEOGRAPHIC VIEW
            # ------------------------------------------
            with tab3:
                st.markdown("### 🗺️ Geographic & Centre Analysis")
                
                st.markdown("#### State-wise Performance Comparison (R.I.S.E %)")
                
                if not filtered_df.empty:
                    state_cat = filtered_df.groupby(['State', 'Academic Year', 'Category']).size().reset_index(name='Count')
                    state_cat['Percentage'] = state_cat.groupby(['State', 'Academic Year'])['Count'].transform(lambda x: x / x.sum() * 100)
                    
                    # Apply Abbreviation to the Academic Year column for this specific chart
                    state_cat['Period'] = state_cat['Academic Year'].map({'Baseline': 'B', 'Endline': 'E'})
                    
                    # Dynamic State Abbreviation (Initials for multi-word, first 3 letters for single-word)
                    def abbreviate_state(state_name):
                        words = str(state_name).split()
                        if len(words) > 1:
                            return "".join([w.upper() for w in words])
                        return str(state_name)[:3].upper()
                        
                    # Create a specific column for the facet layout so the main 'State' is intact for hover
                    state_cat['State Abbr'] = state_cat['State'].apply(abbreviate_state)
                    
                    fig_state = px.bar(state_cat, x="Period", y="Percentage", color="Category", facet_col="State Abbr",
                                       hover_data={"State": True, "State Abbr": False, "Period": False, "Academic Year": True},
                                       color_discrete_map=RISE_COLORS,
                                       text=state_cat['Percentage'].apply(lambda x: f'{x:.1f}%' if not pd.isna(x) and x > 5 else ''),
                                       category_orders={"Category": ["Reviving", "Initiating", "Shaping", "Evolving"],
                                                        "Period": ["B", "E"]})
                    
                    fig_state.update_layout(barmode='stack', yaxis_title="% of Students", margin=dict(l=0, r=0, t=40),
                                            legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, title=""))
                    fig_state.update_xaxes(title_text='')
                    fig_state.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1]))
                    
                    st.plotly_chart(fig_state, width='stretch')
                else:
                    st.info("No data available for State comparison.")
                
                st.markdown("---")

                st.markdown("#### Top 10 Centres (Sorted by % Evolving)")
                
                if not filtered_df.empty:
                    center_cat = filtered_df.groupby(['Centre Name', 'Category']).size().reset_index(name='Count')
                    center_cat['Percentage'] = center_cat.groupby('Centre Name')['Count'].transform(lambda x: x / x.sum() * 100)
                    
                    center_piv = center_cat.pivot(index='Centre Name', columns='Category', values='Percentage').fillna(0)
                    
                    for cat in ["Reviving", "Initiating", "Shaping", "Evolving"]:
                        if cat not in center_piv.columns:
                            center_piv[cat] = 0
                            
                    center_piv_sorted = center_piv.sort_values(
                        by=['Evolving', 'Shaping', 'Initiating', 'Reviving'], 
                        ascending=[False, False, False, False]
                    ).head(10)
                    
                    center_piv_sorted = center_piv_sorted.iloc[::-1]
                    
                    top_centres_long = center_piv_sorted.reset_index().melt(
                        id_vars='Centre Name', 
                        value_vars=["Reviving", "Initiating", "Shaping", "Evolving"], 
                        var_name='Category', 
                        value_name='Percentage'
                    )
                    
                    fig_top_centres = px.bar(top_centres_long, x="Percentage", y="Centre Name", color="Category", 
                                             orientation='h', color_discrete_map=RISE_COLORS,
                                             text=top_centres_long['Percentage'].apply(lambda x: f'{x:.1f}%' if x > 5 else ''),
                                             category_orders={"Category": ["Reviving", "Initiating", "Shaping", "Evolving"]})
                    
                    fig_top_centres.update_layout(barmode='stack', xaxis_title="% of Students", yaxis_title="", margin=dict(l=0, r=0, t=30),
                                                  legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, title=""))
                    st.plotly_chart(fig_top_centres, width='stretch')
                else:
                    st.info("No data available for Top Centres.")


            # ------------------------------------------
            # TAB 4: STUDENT-LEVEL IMPACT
            # ------------------------------------------
            with tab4:
                st.markdown("### 🧑‍🎓 Student-Level Impact (Matched Cohort)")
                st.markdown("Tracking individual student growth by matching their Baseline and Endline records.")
                
                if not base_df.empty and not end_df.empty and 'Student ID' in df.columns:
                    base_clean = base_df[['Student ID', 'Subject', 'Obtained Marks', 'Category']].dropna(subset=['Student ID'])
                    end_clean = end_df[['Student ID', 'Subject', 'Obtained Marks', 'Category']].dropna(subset=['Student ID'])
                    
                    # Duplicate protection added here
                    base_clean = base_clean.drop_duplicates(subset=['Student ID', 'Subject'])
                    end_clean = end_clean.drop_duplicates(subset=['Student ID', 'Subject'])
                    
                    paired_df = pd.merge(base_clean, end_clean, on=['Student ID', 'Subject'], suffixes=('_BL', '_EL'))
                    
                    if not paired_df.empty:
                        paired_df['Score Delta'] = paired_df['Obtained Marks_EL'] - paired_df['Obtained Marks_BL']
                        mean_change = paired_df['Score Delta'].mean()
                        
                        total_paired = len(paired_df)
                        positive_pct = (len(paired_df[paired_df['Score Delta'] > 0]) / total_paired) * 100
                        neutral_pct = (len(paired_df[paired_df['Score Delta'] == 0]) / total_paired) * 100
                        negative_pct = (len(paired_df[paired_df['Score Delta'] < 0]) / total_paired) * 100
                        
                        st.markdown("---")
                        met_col1, met_col2, met_col3, met_col4, met_col5 = st.columns(5)
                        met_col1.metric("Matched Students", f"{total_paired:,}")
                        met_col2.metric("Avg Score Change", f"{mean_change:+.2f}")
                        met_col3.metric("Students (+ Score)", f"{positive_pct:.1f}%")
                        met_col4.metric("Students (No Change)", f"{neutral_pct:.1f}%")
                        met_col5.metric("Students (- Score)", f"{negative_pct:.1f}%")
                        st.markdown("---")
                        
                        st.markdown("#### 🔄 Category Transition Matrix")
                        st.caption("Read rows left-to-right to see student mobility. **Background colors represent transition status:** <span style='color:#82E0AA; font-weight:bold;'>Green (Upward Transition)</span>, <span style='color:#A9A9A9; font-weight:bold;'>Grey (No Transition)</span>, and <span style='color:#FF7F7F; font-weight:bold;'>Red (Downward Transition)</span>.", unsafe_allow_html=True)
                        
                        transition = pd.crosstab(paired_df['Category_BL'], paired_df['Category_EL'], normalize='index') * 100
                        
                        cat_order = ["Reviving", "Initiating", "Shaping", "Evolving"]
                        transition = transition.reindex(index=cat_order, columns=cat_order, fill_value=0)
                        
                        direction_matrix = pd.DataFrame(index=cat_order, columns=cat_order)
                        for i, bl in enumerate(cat_order):
                            for j, el in enumerate(cat_order):
                                if i == j: direction_matrix.loc[bl, el] = 0
                                elif j > i: direction_matrix.loc[bl, el] = 1
                                else: direction_matrix.loc[bl, el] = -1
                                    
                        direction_matrix = direction_matrix.astype(float)
                        
                        fig_heat = px.imshow(direction_matrix, 
                                             labels=dict(x="Endline Category", y="Baseline Category", color="Transition Type"),
                                             x=transition.columns, y=transition.index,
                                             color_continuous_scale=["#FF7F7F", "#F2F4F7", "#82E0AA"]) 
                        
                        text_matrix = transition.map(lambda x: f"{x:.1f}%")
                        fig_heat.update_traces(text=text_matrix, texttemplate="%{text}", 
                                               hovertemplate="Baseline: %{y}<br>Endline: %{x}<br>Students: %{text}<extra></extra>")
                        
                        fig_heat.update_coloraxes(showscale=False)
                        fig_heat.update_layout(margin=dict(l=0, r=0, t=30, b=0), height=500)
                        
                        col1, col2, col3 = st.columns(3)
                        with col2:
                            st.plotly_chart(fig_heat, width='stretch')

                    else:
                        st.warning("⚠️ Could not find matching 'Student ID' and 'Subject' between the Baseline and Endline datasets.")
                else:
                    st.info("⚠️ Both Baseline and Endline datasets with a valid 'Student ID' column are required for this analysis.")

            # ------------------------------------------
            # TAB 5: GENDER-WISE ANALYSIS
            # ------------------------------------------
            with tab5:
                st.markdown("### 🚻 Gender-Wise Performance")
                
                if 'Gender' in filtered_df.columns:
                    # Filter out any lingering null/blank string artifacts
                    gdf = filtered_df[~filtered_df['Gender'].str.lower().isin(['nan', 'none', 'null', ''])].copy()
                    
                    if not gdf.empty:
                        # 1. High-Level Metrics
                        st.markdown("#### 🏆 Endline Average Score Snapshot")
                        
                        g_base = gdf[gdf['Academic Year'] == 'Baseline']
                        g_end = gdf[gdf['Academic Year'] == 'Endline']
                        
                        genders_present = sorted(gdf['Gender'].unique())
                        cols = st.columns(max(len(genders_present), 2)) # Ensure at least 2 columns for layout
                        
                        for i, g in enumerate(genders_present):
                            with cols[i]:
                                b_mean = g_base[g_base['Gender'] == g]['Obtained Marks'].mean() if not g_base.empty else None
                                e_mean = g_end[g_end['Gender'] == g]['Obtained Marks'].mean() if not g_end.empty else None
                                
                                if b_mean is not None and e_mean is not None:
                                    st.metric(f"{g} - Endline Avg", f"{e_mean:.2f}", delta=f"{e_mean - b_mean:.2f}")
                                elif e_mean is not None:
                                    st.metric(f"{g} - Endline Avg", f"{e_mean:.2f}")
                                elif b_mean is not None:
                                    st.metric(f"{g} - Baseline Avg", f"{b_mean:.2f}")
                                    
                        st.markdown("---")
                        
                        # 2. Charts
                        gen_col1, gen_col2 = st.columns(2)
                        
                        with gen_col1:
                            st.markdown("#### 📈 Average Score Trend")
                            st.caption("Direct comparison of mean scores by gender.")
                            
                            avg_gen = gdf.groupby(['Gender', 'Academic Year'])['Obtained Marks'].mean().reset_index()
                            fig_gen_avg = px.bar(avg_gen, x="Gender", y="Obtained Marks", color="Academic Year", barmode="group",
                                                 color_discrete_map=COLOR_MAP, text_auto='.2f')
                            fig_gen_avg.update_layout(yaxis_title="Average Marks", margin=dict(l=0, r=0, t=30),
                                                      legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, title=""))
                            st.plotly_chart(fig_gen_avg, width='stretch')
                            
                        with gen_col2:
                            st.markdown("#### 🧬 R.I.S.E Category Shift")
                            st.caption("Proportional breakdown of performance tiers by gender.")
                            
                            gen_cat = gdf.groupby(['Gender', 'Academic Year', 'Category']).size().reset_index(name='Count')
                            gen_cat['Percentage'] = gen_cat.groupby(['Gender', 'Academic Year'])['Count'].transform(lambda x: x / x.sum() * 100)
                            
                            fig_gen_rise = px.bar(gen_cat, x="Academic Year", y="Percentage", color="Category", facet_col="Gender",
                                                  color_discrete_map=RISE_COLORS,
                                                  text=gen_cat['Percentage'].apply(lambda x: f'{x:.1f}%' if not pd.isna(x) and x > 5 else ''),
                                                  category_orders={"Category": ["Reviving", "Initiating", "Shaping", "Evolving"],
                                                                   "Academic Year": ["Baseline", "Endline"]})
                            fig_gen_rise.update_layout(barmode='stack', yaxis_title="% of Students", margin=dict(l=0, r=0, t=40),
                                                       legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, title=""))
                            fig_gen_rise.update_xaxes(title_text='')
                            fig_gen_rise.for_each_annotation(lambda a: a.update(text=a.text.split("=")[-1]))
                            st.plotly_chart(fig_gen_rise, width='stretch')
                            
                    else:
                        st.info("No valid gender data available in the current filtered selection.")
                else:
                    st.warning("⚠️ 'Gender' column is missing from the uploaded dataset. Please ensure your files include a column labeled 'Gender' with values like 'Boy' or 'Girl'.")

            # ------------------------------------------
            # TAB 6: RTM ANALYSIS
            # ------------------------------------------
            with tab6:
                st.markdown("### 📉 Regression to the Mean (RTM) Analysis")
                
                if not base_df.empty and not end_df.empty and 'Student ID' in df.columns:
                    # Clean and merge data specifically for RTM (Requires numeric marks)
                    base_rtm = base_df[['Student ID', 'Subject', 'Obtained Marks']].dropna(subset=['Student ID', 'Obtained Marks'])
                    end_rtm = end_df[['Student ID', 'Subject', 'Obtained Marks']].dropna(subset=['Student ID', 'Obtained Marks'])
                    
                    # Duplicate protection added here
                    base_rtm = base_rtm.drop_duplicates(subset=['Student ID', 'Subject'])
                    end_rtm = end_rtm.drop_duplicates(subset=['Student ID', 'Subject'])
                    
                    rtm_df = pd.merge(base_rtm, end_rtm, on=['Student ID', 'Subject'], suffixes=('_BL', '_EL'))
                    
                    if not rtm_df.empty:
                        st.markdown("---")
                        # --- NORMALIZATION TOGGLE ---
                        normalize_rtm = st.checkbox("⚙️ Normalize scores (Z-scores) before analysis", value=False, help="Standardizes scores so the Baseline and Endline both have a mean of 0 and standard deviation of 1. This prevents differences in test difficulty or variance from skewing the RTM analysis.")
                        
                        if normalize_rtm:
                            # Convert raw scores to Z-scores
                            rtm_df['Obtained Marks_BL'] = (rtm_df['Obtained Marks_BL'] - rtm_df['Obtained Marks_BL'].mean()) / rtm_df['Obtained Marks_BL'].std()
                            rtm_df['Obtained Marks_EL'] = (rtm_df['Obtained Marks_EL'] - rtm_df['Obtained Marks_EL'].mean()) / rtm_df['Obtained Marks_EL'].std()
                            
                        # Calculate Delta (will naturally use raw or Z-scores based on the toggle)
                        rtm_df['Score Delta'] = rtm_df['Obtained Marks_EL'] - rtm_df['Obtained Marks_BL']
                        
                        # --- KPI METRICS (TOP OF TAB) ---
                        # 1. Calculate Statistics for KPIs
                        correlation = rtm_df['Obtained Marks_BL'].corr(rtm_df['Score Delta'])
                        variance = rtm_df['Obtained Marks_BL'].var()
                        covariance = rtm_df['Obtained Marks_BL'].cov(rtm_df['Score Delta'])
                        slope = covariance / variance if variance != 0 and not pd.isna(variance) else 0.0
                        
                        # Calculate Intercept early for the validation section below
                        intercept = rtm_df['Score Delta'].mean() - (slope * rtm_df['Obtained Marks_BL'].mean()) if not pd.isna(slope) else 0.0
                        
                        # 2. Calculate Improving vs Declining Percentages
                        total_rtm = len(rtm_df)
                        improving_pct = (len(rtm_df[rtm_df['Score Delta'] > 0]) / total_rtm) * 100
                        declining_pct = (len(rtm_df[rtm_df['Score Delta'] < 0]) / total_rtm) * 100
                        
                        # 3. Determine Interpretation Tag
                        if slope <= -0.3:
                            rtm_tag = "Strong RTM detected"
                        elif slope <= -0.1:
                            rtm_tag = "Moderate RTM"
                        elif slope < 0:
                            rtm_tag = "Minimal RTM"
                        else:
                            rtm_tag = "No RTM detected"
                            
                        # 4. Render Top KPIs
                        kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
                        kpi_col1.metric("Correlation (r)", f"{correlation:.3f}" if not pd.isna(correlation) else "N/A")
                        kpi_col2.metric("Regression Slope", f"{slope:.3f}")
                        kpi_col3.metric("Improving vs Declining", f"{improving_pct:.1f}% / {declining_pct:.1f}%")
                        kpi_col4.metric("Interpretation", rtm_tag)
                        
                        # 5. Interpretation Box
                        if slope <= -0.1:
                            st.warning("💡 **Important Insight:** Part of the observed improvement may be due to statistical regression to the mean rather than pure intervention impact. Students who scored exceptionally low on the baseline naturally tend to score closer to the average on the endline.")
                        else:
                            st.success("💡 **Important Insight:** Observed improvements are less likely driven by RTM. The growth seen across the cohort is more likely attributable to the actual impact of the educational intervention.")
                            
                        st.markdown("---")
                        
                        # --- SCATTER PLOT ---
                        st.markdown("#### Core RTM View (Scatter Plot)")
                        st.caption("**Interpretation:** A **negative slope** (trendline going down) suggests the RTM effect is present — meaning students with lower baseline scores tended to improve more, while those with higher baseline scores improved less or dropped.")
                        
                        # A. Scatter Plot with OLS Trendline
                        fig_rtm = px.scatter(
                            rtm_df, 
                            x="Obtained Marks_BL", 
                            y="Score Delta", 
                            trendline="ols",
                            trendline_color_override="red",
                            opacity=0.6,
                            color_discrete_sequence=["#636EFA"],
                            labels={
                                "Obtained Marks_BL": "Baseline Score (Z-Score)" if normalize_rtm else "Baseline Score",
                                "Score Delta": "Score Delta (Z-Score)" if normalize_rtm else "Score Delta (Endline - Baseline)"
                            }
                        )
                        
                        # Add horizontal line at y = 0
                        fig_rtm.add_hline(y=0, line_dash="dash", line_color="black", annotation_text="No Change (Delta = 0)", annotation_position="bottom right")
                        
                        fig_rtm.update_layout(margin=dict(l=0, r=0, t=30))
                        st.plotly_chart(fig_rtm, width='stretch')
                        
                        st.markdown("---")
                        
                        # --- BINNED ANALYSIS ---
                        st.markdown("#### Binned Analysis (Quintiles)")
                        st.caption("**Interpretation:** Students are grouped into 5 equal-sized bins based on their initial Baseline scores. If RTM is present, the chart will typically show a 'staircase' pattern: large positive deltas on the left (lowest scorers improving the most) and smaller or negative deltas on the right (highest scorers plateauing or declining).")
                        
                        try:
                            # Create quintiles (5 bins) based on Baseline scores. 
                            # 'duplicates=drop' prevents errors if many students have the exact same score.
                            rtm_df['BL_Quintile'] = pd.qcut(rtm_df['Obtained Marks_BL'], q=5, duplicates='drop')
                            
                            # Calculate metrics for each bin
                            binned_stats = rtm_df.groupby('BL_Quintile', observed=False).agg(
                                Avg_BL_Score=('Obtained Marks_BL', 'mean'),
                                Avg_Score_Delta=('Score Delta', 'mean'),
                                Student_Count=('Student ID', 'count')
                            ).reset_index()
                            
                            # Convert categorical intervals to strings for clean X-axis labels
                            binned_stats['BL_Quintile_Str'] = binned_stats['BL_Quintile'].astype(str)
                            binned_stats = binned_stats.sort_values('Avg_BL_Score')
                            
                            # Plotly Bar Chart with diverging colors (Green for positive delta, Red for negative)
                            fig_binned = px.bar(
                                binned_stats, 
                                x='BL_Quintile_Str', 
                                y='Avg_Score_Delta',
                                text=binned_stats['Avg_Score_Delta'].apply(lambda x: f"{x:+.2f}"),
                                color='Avg_Score_Delta',
                                color_continuous_scale=px.colors.diverging.RdYlGn,
                                color_continuous_midpoint=0,
                                labels={
                                    "BL_Quintile_Str": "Baseline Score Range (Quintiles)",
                                    "Avg_Score_Delta": "Average Score Delta"
                                },
                                hover_data={
                                    "Student_Count": True, 
                                    "Avg_BL_Score": ':.2f',
                                    "Avg_Score_Delta": ':.2f'
                                }
                            )
                            
                            fig_binned.add_hline(y=0, line_dash="solid", line_color="black", line_width=1)
                            fig_binned.update_traces(textposition='outside')
                            fig_binned.update_layout(margin=dict(l=0, r=0, t=30, b=40), coloraxis_showscale=False)
                            
                            st.plotly_chart(fig_binned, width='stretch')
                            
                        except ValueError:
                            st.info("Not enough variance in Baseline scores to generate quintile bins for this selection.")
                            
                        st.markdown("---")
                        
                        # --- STATISTICAL VALIDATION ---
                        st.markdown("#### 🧮 Statistical Validation")
                        st.caption("Quantifying the strength of the Regression to the Mean effect using linear regression.")
                            
                        # 3. Calculate R-squared (For simple linear regression, R^2 = r^2)
                        r_squared = correlation ** 2 if not pd.isna(correlation) else 0.0
                        
                        # Display the metrics
                        stat_col1, stat_col2, stat_col3 = st.columns(3)
                        stat_col1.metric("Correlation (r)", f"{correlation:.3f}" if not pd.isna(correlation) else "N/A")
                        stat_col2.metric("Regression Slope (b)", f"{slope:.3f}")
                        stat_col3.metric("R-squared (R²)", f"{r_squared:.3f}")
                        
                        # Provide a dynamic text interpretation based on the slope
                        st.markdown("**Mathematical Interpretation:**")
                        equation_str = f"**Score Delta = {intercept:.2f} + ({slope:.2f} * Baseline)**"
                        
                        if slope < -0.3:
                            st.success(f"✔️ **Strong RTM Effect Confirmed:** The strongly negative slope and equation {equation_str} prove that students with lower baselines experienced significantly higher growth.")
                        elif slope < -0.1:
                            st.info(f"ℹ️ **Moderate RTM Effect:** The negative slope and equation {equation_str} indicate a moderate Regression to the Mean pattern.")
                        elif slope < 0:
                            st.warning(f"⚠️ **Weak RTM Effect:** The slope is very close to zero. {equation_str}")
                        else:
                            st.error(f"❌ **No RTM Effect Detected:** The slope is positive ({slope:.2f}), meaning higher-scoring students actually improved *more* than lower-scoring students.")
                            
                    else:
                        st.warning("⚠️ Could not find matching 'Student ID' and 'Subject' for RTM analysis.")
                else:
                    st.info("⚠️ Both Baseline and Endline datasets with a valid 'Student ID' column are required for this analysis.")

else:
    # Empty State Error Message text adjusted to ask for the local file instead of an upload
    st.error(f"⚠️ Data file not found! Please ensure `{DATA_FILE}` is placed in the same folder as this script to populate the dashboard.")
    st.image("https://cdn-icons-png.flaticon.com/512/7264/7264168.png", width=150)
