import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import statsmodels.api as sm
import urllib.request
import json
import io
from streamlit_oauth import OAuth2Component

# ==========================================
# PAGE CONFIGURATION & CUSTOM CSS
# ==========================================
st.set_page_config(page_title="Impact Analytics Dashboard", layout="wide", initial_sidebar_state="expanded")

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
try:
    CLIENT_ID = st.secrets["GOOGLE_CLIENT_ID"]
    CLIENT_SECRET = st.secrets["GOOGLE_CLIENT_SECRET"]
except FileNotFoundError:
    st.error("Missing `.streamlit/secrets.toml` file or Streamlit Cloud Secrets. Please ensure your Google Client ID and Secret are configured.")
    st.stop()

AUTHORIZE_URL = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN_URL = "https://oauth2.googleapis.com/token"
REVOKE_TOKEN_URL = "https://oauth2.googleapis.com/revoke"

if "logged_in_email" not in st.session_state:
    st.session_state["logged_in_email"] = None
if "user_first_name" not in st.session_state:
    st.session_state["user_first_name"] = "User"

if not st.session_state["logged_in_email"]:
    col1, col2, col3 = st.columns(3)
    with col2:
        st.write("")
        st.write("")
        try:
            st.image("evidyaloka_logo.png", width=320)
        except:
            st.empty()
            
        st.markdown("<h2 style='text-align: center; color: #0094c9;'>Student Analytics Portal</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center;'>Please sign in with your @evidyaloka.org email to access the dashboard.</p>", unsafe_allow_html=True)
        st.markdown("---")
        
        oauth2 = OAuth2Component(CLIENT_ID, CLIENT_SECRET, AUTHORIZE_URL, TOKEN_URL, REVOKE_TOKEN_URL)
        
        result = oauth2.authorize_button(
            name="Sign in with Google",
            icon="https://upload.wikimedia.org/wikipedia/commons/5/53/Google_%22G%22_Logo.svg",
            redirect_uri="https://ev-assessments.streamlit.app", 
            scope="openid email profile",
            key="google_login",
            use_container_width=True
        )
        
        if result and "token" in result:
            id_token = result["token"]["id_token"]
            if isinstance(id_token, list):
                id_token = id_token
            verify_url = f"https://oauth2.googleapis.com/tokeninfo?id_token={id_token}"
            try:
                with urllib.request.urlopen(verify_url) as response:
                    user_info = json.loads(response.read().decode())
                st.session_state["logged_in_email"] = user_info.get("email") 
                st.session_state["user_first_name"] = user_info.get("given_name", "User")
                st.rerun()
            except Exception as e:
                st.error(f"⚠️ Error verifying login with Google: {e}")
                st.stop()
            
    st.stop()


# ==========================================
# 🏠 APP ROUTER / HOMEPAGE GATEKEEPER
# ==========================================
if "current_page" not in st.session_state:
    st.session_state["current_page"] = "home"

if st.session_state["current_page"] == "home":
    st.write("")
    st.write("")
    st.title(f"👋 Welcome, {st.session_state['user_first_name']}!")
    st.markdown("<p style='color: gray; font-size: 1.1em;'>Select an application below to continue.</p>", unsafe_allow_html=True)
    st.markdown("---")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.markdown("<h1 style='text-align: center; font-size: 4rem;'>📈</h1>", unsafe_allow_html=True)
        if st.button("Impact Analytics Dashboard", use_container_width=True):
            st.session_state["current_page"] = "dashboard"
            st.rerun()
            
    with col2:
        st.markdown("<h1 style='text-align: center; font-size: 4rem;'>🏛️</h1>", unsafe_allow_html=True)
        if st.button("Longitudinal Analysis", use_container_width=True):
            st.session_state["current_page"] = "longitudinal"
            st.rerun()
            
    st.stop()


# ==========================================
# 🏛️ LONGITUDINAL ANALYSIS MODULE
# ==========================================
if st.session_state["current_page"] == "longitudinal":
    st.title("🏛️ Strategic Longitudinal Analysis")
    st.markdown("<p style='color: gray; font-size: 1.1em;'>Year-over-Year Trajectories, Equity Tracking, and Strategic Insights (AY 24-25 vs AY 25-26)</p>", unsafe_allow_html=True)
    st.markdown("---")

    @st.cache_data
    def load_multi_year_data(file_24, file_25):
        def clean_sheet(df, year, period):
            if df.empty: return pd.DataFrame()
            cols = ['State', 'Centre Name', 'Donor', 'Subject', 'Grade', 'Student ID', 'Gender', 'Obtained Marks', 'Rubrics', 'Category']
            df_clean = df[[c for c in cols if c in df.columns]].copy()
            if 'Rubrics' in df_clean.columns: df_clean.rename(columns={'Rubrics': 'Category'}, inplace=True)
            df_clean['Academic Year'] = year
            df_clean['Period'] = period
            df_clean['Timepoint'] = f"{year} {period}"
            df_clean['Obtained Marks'] = pd.to_numeric(df_clean['Obtained Marks'], errors='coerce')
            df_clean['Student ID'] = df_clean['Student ID'].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            for col in ['State', 'Centre Name', 'Donor', 'Subject', 'Gender', 'Category']:
                if col in df_clean.columns:
                    df_clean[col] = df_clean[col].astype(str).str.strip()
            if 'Gender' in df_clean.columns:
                df_clean['Gender'] = df_clean['Gender'].str.title()
            if 'Grade' in df_clean.columns:
                df_clean['Grade'] = df_clean['Grade'].astype(str).str.replace(r'\.0$', '', regex=True)
            return df_clean.dropna(subset=['Student ID', 'Obtained Marks'])

        try:
            xls_24 = pd.ExcelFile(file_24)
            df_24_bl = clean_sheet(pd.read_excel(file_24, sheet_name=0), 'AY24-25', 'Baseline')
            df_24_el = clean_sheet(pd.read_excel(file_24, sheet_name=1 if len(xls_24.sheet_names) > 1 else 0), 'AY24-25', 'Endline')
            xls_25 = pd.ExcelFile(file_25)
            df_25_bl = clean_sheet(pd.read_excel(file_25, sheet_name=0), 'AY25-26', 'Baseline')
            df_25_el = clean_sheet(pd.read_excel(file_25, sheet_name=1 if len(xls_25.sheet_names) > 1 else 0), 'AY25-26', 'Endline')
            combined = pd.concat([df_24_bl, df_24_el, df_25_bl, df_25_el], ignore_index=True)
            return combined
        except Exception as e:
            return None

    FILE_24 = "EL-BL-Data-AY-24-25.xlsx"
    FILE_25 = "BL-EL-AY-25-26-Final-AllSubjects.xlsx"
    
    filtered_df_long = pd.DataFrame()

    if os.path.exists(FILE_24) and os.path.exists(FILE_25):
        with st.spinner('Synthesizing Multi-Year Intelligence...'):
            df_long = load_multi_year_data(FILE_24, FILE_25)
            
        with st.sidebar:
            try:
                st.image("evidyaloka_logo.png", width=273)
            except:
                st.warning("⚠️ Logo not found.")
            
            st.success(f"👤 **Logged in as:** {st.session_state['user_first_name']}")
            
            nav_col1, nav_col2 = st.columns(2)
            with nav_col1:
                if st.button("🏠 Home", use_container_width=True, key="nav_home_long"):
                    st.session_state["current_page"] = "home"
                    st.rerun()
            with nav_col2:
                if st.button("Sign Out", use_container_width=True, key="signout_long"):
                    st.session_state["logged_in_email"] = None
                    st.session_state["user_first_name"] = "User"
                    st.session_state["current_page"] = "home"
                    st.rerun()
                    
            st.markdown("---")
            st.info("💡 **Module Note:** This section analyzes the overlap between AY 24-25 and AY 25-26 to track long-term strategic growth.")
            
            if df_long is not None and not df_long.empty:
                st.header("🎯 Global Filters")
                states = ["All"] + sorted([str(x) for x in df_long['State'].dropna().unique()]) if 'State' in df_long.columns else ["All"]
                selected_states = st.selectbox("Select State", states, index=0, key="state_long")
                df_state_filtered = df_long.copy()
                if selected_states != "All": 
                    df_state_filtered = df_state_filtered[df_state_filtered['State'].astype(str) == selected_states]

                donors = ["All"] + sorted([str(x) for x in df_state_filtered['Donor'].dropna().unique()]) if 'Donor' in df_state_filtered.columns else ["All"]
                selected_donors = st.selectbox("Select Donor", donors, index=0, key="donor_long")
                df_donor_filtered = df_state_filtered.copy()
                if selected_donors != "All":
                    df_donor_filtered = df_donor_filtered[df_donor_filtered['Donor'].astype(str) == selected_donors]

                centres = ["All"] + sorted([str(x) for x in df_donor_filtered['Centre Name'].dropna().unique()]) if 'Centre Name' in df_donor_filtered.columns else ["All"]
                selected_centres = st.selectbox("Select Centre", centres, index=0, key="centre_long")
                df_centre_filtered = df_donor_filtered.copy()
                if selected_centres != "All":
                    df_centre_filtered = df_centre_filtered[df_centre_filtered['Centre Name'].astype(str) == selected_centres]

                subjects = ["All"] + sorted([str(x) for x in df_centre_filtered['Subject'].dropna().unique()]) if 'Subject' in df_centre_filtered.columns else ["All"]
                selected_subjects = st.selectbox("Select Subject", subjects, index=0, key="subject_long")
                df_subject_filtered = df_centre_filtered.copy()
                if selected_subjects != "All":
                    df_subject_filtered = df_subject_filtered[df_subject_filtered['Subject'].astype(str) == selected_subjects]

                df_base_year = df_subject_filtered[df_subject_filtered['Academic Year'] == 'AY24-25']
                grades = sorted([str(x) for x in df_base_year['Grade'].dropna().unique()]) if 'Grade' in df_base_year.columns else []
                selected_grades = st.multiselect("Select AY 24-25 Grade (Cohort Tracking)", options=grades, default=grades, key="grade_long", help="Select the student's grade in AY 24-25. The dashboard will automatically track them into their promoted grade for AY 25-26.")

                df_grade_filtered = df_subject_filtered.copy()
                if selected_grades:
                    cohort_ids = df_base_year[df_base_year['Grade'].astype(str).isin(selected_grades)]['Student ID'].unique()
                    df_grade_filtered = df_grade_filtered[df_grade_filtered['Student ID'].isin(cohort_ids)]
                else:
                    df_grade_filtered = df_grade_filtered.iloc[0:0] 

                if 'Gender' in df_grade_filtered.columns:
                    valid_genders = df_grade_filtered.dropna(subset=['Gender'])
                    valid_genders = valid_genders[~valid_genders['Gender'].astype(str).str.lower().isin(['nan', 'none', 'null', ''])]
                    genders = sorted([str(x) for x in valid_genders['Gender'].unique()])
                    if genders:
                        selected_genders = st.multiselect("Select Gender(s)", options=genders, default=genders, key="gender_long")
                        filtered_df_long = df_grade_filtered[df_grade_filtered['Gender'].astype(str).isin(selected_genders)].copy()
                    else:
                        filtered_df_long = df_grade_filtered.copy()
                else:
                    filtered_df_long = df_grade_filtered.copy()

        if df_long is not None and not df_long.empty:
            if filtered_df_long.empty:
                st.warning("⚠️ No data available for the selected filters. Please adjust your criteria.")
            else:
                df_el24 = filtered_df_long[filtered_df_long['Timepoint'] == 'AY24-25 Endline']
                df_el25 = filtered_df_long[filtered_df_long['Timepoint'] == 'AY25-26 Endline']
                retained_students = set(df_el24['Student ID']).intersection(set(df_el25['Student ID']))
                
                if len(retained_students) == 0:
                    st.error("No overlapping Student IDs found between AY 24-25 and AY 25-26 for the current filters. Cannot perform longitudinal analysis.")
                    st.stop()
                    
                df_ret_24 = df_el24[df_el24['Student ID'].isin(retained_students)]
                df_ret_25 = df_el25[df_el25['Student ID'].isin(retained_students)]
                
                mig_tab, sub_tab, gen_tab, mic_tab = st.tabs([
                    "📊 Overall Health (Migration)", 
                    "📚 Subject Efficacy", 
                    "🚻 Gender Equity",
                    "🧑‍🎓 Single Student"
                ])
                
                with mig_tab:
                    st.markdown("### 🧱 Structural Tier Migration (Retained Cohort)")
                    col_m1, col_m2 = st.columns([1.5, 1])
                    with col_m1:
                        if 'Category' in df_ret_24.columns:
                            cat_24 = df_ret_24['Category'].value_counts(normalize=True).reset_index()
                            cat_24.columns = ['Category', 'Percentage']
                            cat_24['Year'] = 'AY 24-25 (Endline)'
                            cat_25 = df_ret_25['Category'].value_counts(normalize=True).reset_index()
                            cat_25.columns = ['Category', 'Percentage']
                            cat_25['Year'] = 'AY 25-26 (Endline)'
                            cat_df = pd.concat([cat_24, cat_25])
                            cat_df['Percentage'] = cat_df['Percentage'] * 100
                            fig_cat = px.bar(cat_df, x="Year", y="Percentage", color="Category", 
                                             color_discrete_map={"Reviving": "#f27c48", "Initiating": "#0094c9", "Shaping": "#00964d", "Evolving": "#ed1c2d"},
                                             text=cat_df['Percentage'].apply(lambda x: f'{x:.1f}%'),
                                             category_orders={"Category": ["Reviving", "Initiating", "Shaping", "Evolving"], "Year": ["AY 24-25 (Endline)", "AY 25-26 (Endline)"]})
                            fig_cat.update_layout(barmode='stack', xaxis_title="", yaxis_title="% of Cohort", 
                                                  plot_bgcolor="rgba(0,0,0,0)", margin=dict(l=0, r=0, t=30),
                                                  legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, title=""))
                            st.plotly_chart(fig_cat, width='stretch')
                    with col_m2:
                        st.info("**What this means:**\nThis chart strips away the noise and looks only at students who stayed with us for two full years. It shows how the 'shape' of their performance changed. A successful program will show the red/orange sections ('Reviving'/'Initiating') shrinking as students migrate upward into green sections ('Shaping'/'Evolving').")
                        try:
                            rev_24 = cat_24[cat_24['Category'] == 'Reviving']['Percentage'].values * 100 if not cat_24[cat_24['Category'] == 'Reviving'].empty else 0
                            rev_25 = cat_25[cat_25['Category'] == 'Reviving']['Percentage'].values * 100 if not cat_25[cat_25['Category'] == 'Reviving'].empty else 0
                            rev_diff = rev_25 - rev_24
                            st.success(f"**🔍 Key Insight:**\nThe proportion of critically struggling students ('Reviving') changed by **{rev_diff:+.1f}%** Year-over-Year.")
                            if rev_diff < 0:
                                st.markdown("**💡 Strategic Suggestion:**\nExcellent progress. The base is shrinking. Keep investing in the current foundational remedial strategies as they are actively pulling students out of the danger zone.")
                            else:
                                st.markdown("**💡 Strategic Suggestion:**\nThe struggling cohort is stagnating or growing. We need to audit our Tier-1 interventions. Consider implementing highly targeted, small-group tutoring specifically for the 'Reviving' students, as current general methods aren't lifting them.")
                        except:
                            st.write("Insufficient category data for insights.")

                with sub_tab:
                    st.markdown("### 📈 YoY Subject Trajectory (Slopegraph)")
                    col_s1, col_s2 = st.columns([1.5, 1])
                    with col_s1:
                        subj_24 = df_ret_24.groupby('Subject')['Obtained Marks'].mean().reset_index()
                        subj_24['Year'] = 'AY 24-25'
                        subj_25 = df_ret_25.groupby('Subject')['Obtained Marks'].mean().reset_index()
                        subj_25['Year'] = 'AY 25-26'
                        slope_df = pd.concat([subj_24, subj_25])
                        fig_slope = px.line(slope_df, x="Year", y="Obtained Marks", color="Subject", markers=True,
                                            line_shape="linear", text="Obtained Marks")
                        fig_slope.update_traces(textposition="top center", texttemplate='%{text:.1f}', marker=dict(size=10))
                        fig_slope.update_layout(xaxis_title="", yaxis_title="Average Score", showlegend=True, 
                                                plot_bgcolor="rgba(0,0,0,0)", margin=dict(l=0, r=0, t=30))
                        fig_slope.update_xaxes(showgrid=False, linecolor='black')
                        fig_slope.update_yaxes(showgrid=True, gridcolor='lightgrey', zeroline=False)
                        st.plotly_chart(fig_slope, width='stretch')
                    with col_s2:
                        st.info("**What this means:**\nThis 'Slopegraph' visualizes momentum. Instead of looking at a single point in time, the steepness and direction of the lines immediately reveal which subjects are improving and which are backsliding over a 12-month period.")
                        try:
                            growth_df = pd.merge(subj_24, subj_25, on='Subject', suffixes=('_24', '_25'))
                            growth_df['Delta'] = growth_df['Obtained Marks_25'] - growth_df['Obtained Marks_24']
                            best_sub = growth_df.loc[growth_df['Delta'].idxmax()]
                            worst_sub = growth_df.loc[growth_df['Delta'].idxmin()]
                            st.success(f"**🔍 Key Insight:**\n**{best_sub['Subject']}** is the strongest performer, growing by {best_sub['Delta']:+.2f} points. Conversely, **{worst_sub['Subject']}** showed the weakest momentum ({worst_sub['Delta']:+.2f} points).")
                            st.markdown(f"**💡 Strategic Suggestion:**\nConduct a 'Bright Spot Analysis' on the {best_sub['Subject']} curriculum and teaching methodologies from this past year. Identify the core drivers of that success and cross-pollinate those techniques into the {worst_sub['Subject']} planning for the upcoming semester.")
                        except:
                            st.write("Insufficient subject overlap for comparative insights.")

                with gen_tab:
                    st.markdown("### 🚻 Gender Equity Tracking")
                    if 'Gender' in df_ret_24.columns and 'Gender' in df_ret_25.columns:
                        col_g1, col_g2 = st.columns([1.5, 1])
                        with col_g1:
                            g24 = df_ret_24[df_ret_24['Gender'].isin(['Boy', 'Girl'])].groupby('Gender')['Obtained Marks'].mean().reset_index()
                            g24['Year'] = 'AY 24-25'
                            g25 = df_ret_25[df_ret_25['Gender'].isin(['Boy', 'Girl'])].groupby('Gender')['Obtained Marks'].mean().reset_index()
                            g25['Year'] = 'AY 25-26'
                            gen_df = pd.concat([g24, g25])
                            fig_gen = px.bar(gen_df, x="Year", y="Obtained Marks", color="Gender", barmode='group',
                                             color_discrete_map={"Boy": "#636EFA", "Girl": "#EF553B"},
                                             text=gen_df['Obtained Marks'].apply(lambda x: f'{x:.2f}'))
                            fig_gen.update_layout(xaxis_title="", yaxis_title="Average Score", 
                                                  plot_bgcolor="rgba(0,0,0,0)", margin=dict(l=0, r=0, t=30),
                                                  legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, title=""))
                            st.plotly_chart(fig_gen, width='stretch')
                        with col_g2:
                            st.info("**What this means:**\nThis metric tracks if our educational impact is equitable. It ensures that as the overall average rises, one gender isn't being left behind.")
                            try:
                                gap_24 = abs(g24[g24['Gender'] == 'Girl']['Obtained Marks'].values - g24[g24['Gender'] == 'Boy']['Obtained Marks'].values)
                                gap_25 = abs(g25[g25['Gender'] == 'Girl']['Obtained Marks'].values - g25[g25['Gender'] == 'Boy']['Obtained Marks'].values)
                                st.success(f"**🔍 Key Insight:**\nThe performance gap between boys and girls was **{gap_24:.2f} points** last year, and is now **{gap_25:.2f} points** this year.")
                                if gap_25 < gap_24:
                                    st.markdown("**💡 Strategic Suggestion:**\nThe equity gap is closing. Current inclusive classroom engagement strategies are working. Maintain the standard.")
                                else:
                                    st.markdown("**💡 Strategic Suggestion:**\nThe equity gap is widening. We recommend auditing classroom interactions or assessment biases. Consider implementing targeted mentorship or gender-specific encouragement initiatives in underperforming demographics.")
                            except:
                                st.write("Insufficient gender data to calculate gaps.")
                    else:
                        st.warning("Gender data is not consistently available across both academic years to perform this analysis.")

                with mic_tab:
                    st.markdown("### 🔎 Deep-Dive: Individual Trajectory")
                    all_retained_ids = sorted([str(x) for x in retained_students])
                    selected_student = st.selectbox("Search Student ID (Retained Cohort Only)", options=["Select an ID..."] + all_retained_ids, key="student_search_long")
                    if selected_student != "Select an ID...":
                        student_data = filtered_df_long[filtered_df_long['Student ID'].astype(str) == selected_student].copy()
                        if not student_data.empty:
                            time_order = ['AY24-25 Baseline', 'AY24-25 Endline', 'AY25-26 Baseline', 'AY25-26 Endline']
                            student_data['Timepoint'] = pd.Categorical(student_data['Timepoint'], categories=time_order, ordered=True)
                            student_data = student_data.sort_values('Timepoint')
                            col_ind1, col_ind2 = st.columns(2)
                            with col_ind1:
                                fig_ind = px.line(student_data, x="Timepoint", y="Obtained Marks", color="Subject", 
                                                  markers=True, line_shape="spline", text="Obtained Marks")
                                fig_ind.update_traces(textposition="top center", marker=dict(size=12))
                                fig_ind.update_layout(
                                    xaxis_title="", yaxis_title="Score", 
                                    yaxis=dict(range=[0, max(student_data['Obtained Marks'].max() + 1, 11)]), 
                                    plot_bgcolor="rgba(248, 249, 250, 0.5)", 
                                    margin=dict(l=0, r=0, t=30), hovermode="x unified"
                                )
                                fig_ind.update_xaxes(showgrid=True, gridcolor='lightgrey')
                                fig_ind.update_yaxes(showgrid=True, gridcolor='lightgrey', zeroline=True)
                                st.plotly_chart(fig_ind, width='stretch')
                            with col_ind2:
                                st.info("**What this means:**\nThis isolates the ultimate metric of success: *did this specific human being learn and grow over two years?* It reveals seasonal learning loss (drops between Endline and the next Baseline) and overall subject mastery.")
                                st.markdown("**💡 Case Management Next Steps:**")
                                st.markdown("1. **Check for Summer Slide:** Look at the gap between *AY24-25 Endline* and *AY25-26 Baseline*. If there is a sharp drop, this student suffers from severe retention loss during breaks.")
                                st.markdown("2. **Subject Variance:** If one line is consistently lower than the rest, flag this student for specific subject-level remediation rather than general tutoring.")
                            st.markdown("**Underlying Records**")
                            st.dataframe(student_data[['Academic Year', 'Period', 'Subject', 'Obtained Marks', 'Category']].reset_index(drop=True), use_container_width=True)
                        else:
                            st.warning("No data found for this Student ID.")
                    
    else:
        st.error(f"⚠️ **Longitudinal Data Missing!** \n\nTo view this module, both `{FILE_24}` and `{FILE_25}` must be placed in the same folder as this script.")

    st.stop()


# ==========================================
# 🚀 MAIN DASHBOARD
# ==========================================
st.title("📈 Impact Analytics Dashboard")
st.markdown("<p style='color: gray; font-size: 1.1em;'>Comprehensive Baseline vs. Endline Performance Assessment</p>", unsafe_allow_html=True)

with st.sidebar:
    try:
        st.image("evidyaloka_logo.png", width=273)
    except:
        st.warning("⚠️ Logo not found.")
    
    st.success(f"👤 **Logged in as:** {st.session_state['user_first_name']}")
    
    nav_col1, nav_col2 = st.columns(2)
    with nav_col1:
        if st.button("🏠 Home", use_container_width=True, key="nav_home_main"):
            st.session_state["current_page"] = "home"
            st.rerun()
    with nav_col2:
        if st.button("Sign Out", use_container_width=True, key="signout_main"):
            st.session_state["logged_in_email"] = None
            st.session_state["user_first_name"] = "User"
            st.session_state["current_page"] = "home"
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

    for col in ['State', 'Centre Name', 'Donor', 'Subject', 'Student ID', 'Gender', 'Category']:
        if col in df_combined.columns:
            df_combined[col] = df_combined[col].astype(str).str.strip()

    if 'Gender' in df_combined.columns:
        df_combined['Gender'] = df_combined['Gender'].astype(str).str.strip().str.title()

    if 'Grade' in df_combined.columns:
        df_combined['Grade'] = df_combined['Grade'].astype(str).str.replace(r'\.0$', '', regex=True)

    if 'Category' in df_combined.columns:
        rise_order = ["Reviving", "Initiating", "Shaping", "Evolving"]
        df_combined['Category'] = pd.Categorical(df_combined['Category'], categories=rise_order, ordered=True)

    df_combined['Obtained Marks'] = pd.to_numeric(df_combined['Obtained Marks'], errors='coerce')

    return df_combined

COLOR_MAP = {'Baseline': '#636EFA', 'Endline': '#00CC96'}
RISE_COLORS = {"Reviving": "#f27c48", "Initiating": "#0094c9", "Shaping": "#00964d", "Evolving": "#ed1c2d"}

DATA_FILE = "BL-EL-AY-25-26-Final-AllSubjects.xlsx"

if os.path.exists(DATA_FILE):
    with st.spinner('Loading and crunching numbers...'):
        df = load_and_prep_data(DATA_FILE)

    if not df.empty:
        with st.sidebar:
            st.header("🎯 Global Filters")
            
            states = ["All"] + sorted([str(x) for x in df['State'].dropna().unique()])
            selected_states = st.selectbox("Select State", states, index=0)
            df_state_filtered = df.copy()
            if selected_states != "All": 
                df_state_filtered = df_state_filtered[df_state_filtered['State'].astype(str) == selected_states]

            donors = ["All"] + sorted([str(x) for x in df_state_filtered['Donor'].dropna().unique()])
            selected_donors = st.selectbox("Select Donor", donors, index=0)
            df_donor_filtered = df_state_filtered.copy()
            if selected_donors != "All":
                df_donor_filtered = df_donor_filtered[df_donor_filtered['Donor'].astype(str) == selected_donors]

            centres = ["All"] + sorted([str(x) for x in df_donor_filtered['Centre Name'].dropna().unique()])
            selected_centres = st.selectbox("Select Centre", centres, index=0)
            df_centre_filtered = df_donor_filtered.copy()
            if selected_centres != "All":
                df_centre_filtered = df_centre_filtered[df_centre_filtered['Centre Name'].astype(str) == selected_centres]

            subjects = ["All"] + sorted([str(x) for x in df_centre_filtered['Subject'].dropna().unique()])
            selected_subjects = st.selectbox("Select Subject", subjects, index=0)
            df_subject_filtered = df_centre_filtered.copy()
            if selected_subjects != "All":
                df_subject_filtered = df_subject_filtered[df_subject_filtered['Subject'].astype(str) == selected_subjects]

            grades = sorted([str(x) for x in df_subject_filtered['Grade'].dropna().unique()])
            selected_grades = st.multiselect("Select Grade(s)", options=grades, default=grades)
            df_grade_filtered = df_subject_filtered.copy()
            if selected_grades:
                df_grade_filtered = df_grade_filtered[df_grade_filtered['Grade'].astype(str).isin(selected_grades)]
            else:
                df_grade_filtered = df_grade_filtered.iloc[0:0] 

            if 'Gender' in df_grade_filtered.columns:
                valid_genders = df_grade_filtered.dropna(subset=['Gender'])
                valid_genders = valid_genders[~valid_genders['Gender'].astype(str).str.lower().isin(['nan', 'none', 'null', ''])]
                genders = sorted([str(x) for x in valid_genders['Gender'].unique()])
                if genders:
                    selected_genders = st.multiselect("Select Gender(s)", options=genders, default=genders)
                    filtered_df = df_grade_filtered[df_grade_filtered['Gender'].astype(str).isin(selected_genders)].copy()
                else:
                    filtered_df = df_grade_filtered.copy()
            else:
                filtered_df = df_grade_filtered.copy()

        if filtered_df.empty:
            st.warning("⚠️ No data available for the selected filters. Please adjust your criteria.")
        else:
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
                kpi1, kpi2, kpi3, kpi4, kpi5 = st.columns(5)
                
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
                    state_cat['Period'] = state_cat['Academic Year'].map({'Baseline': 'B', 'Endline': 'E'})
                    
                    def abbreviate_state(state_name):
                        words = str(state_name).split()
                        if len(words) > 1:
                            return "".join([w.upper() for w in words])
                        return str(state_name)[:3].upper()
                        
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
                    gdf = filtered_df[~filtered_df['Gender'].astype(str).str.lower().isin(['nan', 'none', 'null', ''])].copy()
                    if not gdf.empty:
                        st.markdown("#### 🏆 Endline Average Score Snapshot")
                        g_base = gdf[gdf['Academic Year'] == 'Baseline']
                        g_end = gdf[gdf['Academic Year'] == 'Endline']
                        genders_present = sorted([str(x) for x in gdf['Gender'].dropna().unique()])
                        cols = st.columns(max(len(genders_present), 2))
                        for i, g in enumerate(genders_present):
                            with cols[i]:
                                b_mean = g_base[g_base['Gender'].astype(str) == g]['Obtained Marks'].mean() if not g_base.empty else None
                                e_mean = g_end[g_end['Gender'].astype(str) == g]['Obtained Marks'].mean() if not g_end.empty else None
                                if b_mean is not None and e_mean is not None:
                                    st.metric(f"{g} - Endline Avg", f"{e_mean:.2f}", delta=f"{e_mean - b_mean:.2f}")
                                elif e_mean is not None:
                                    st.metric(f"{g} - Endline Avg", f"{e_mean:.2f}")
                                elif b_mean is not None:
                                    st.metric(f"{g} - Baseline Avg", f"{b_mean:.2f}")
                        st.markdown("---")
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
                    base_rtm = base_df[['Student ID', 'Subject', 'Obtained Marks']].dropna(subset=['Student ID', 'Obtained Marks'])
                    end_rtm = end_df[['Student ID', 'Subject', 'Obtained Marks']].dropna(subset=['Student ID', 'Obtained Marks'])
                    base_rtm = base_rtm.drop_duplicates(subset=['Student ID', 'Subject'])
                    end_rtm = end_rtm.drop_duplicates(subset=['Student ID', 'Subject'])
                    rtm_df = pd.merge(base_rtm, end_rtm, on=['Student ID', 'Subject'], suffixes=('_BL', '_EL'))
                    
                    if not rtm_df.empty:
                        st.markdown("---")
                        normalize_rtm = st.checkbox("⚙️ Normalize scores (Z-scores) before analysis", value=False, help="Standardizes scores so the Baseline and Endline both have a mean of 0 and standard deviation of 1. This prevents differences in test difficulty or variance from skewing the RTM analysis.")
                        if normalize_rtm:
                            rtm_df['Obtained Marks_BL'] = (rtm_df['Obtained Marks_BL'] - rtm_df['Obtained Marks_BL'].mean()) / rtm_df['Obtained Marks_BL'].std()
                            rtm_df['Obtained Marks_EL'] = (rtm_df['Obtained Marks_EL'] - rtm_df['Obtained Marks_EL'].mean()) / rtm_df['Obtained Marks_EL'].std()
                        rtm_df['Score Delta'] = rtm_df['Obtained Marks_EL'] - rtm_df['Obtained Marks_BL']
                        
                        correlation = rtm_df['Obtained Marks_BL'].corr(rtm_df['Score Delta'])
                        variance = rtm_df['Obtained Marks_BL'].var()
                        covariance = rtm_df['Obtained Marks_BL'].cov(rtm_df['Score Delta'])
                        slope = covariance / variance if variance != 0 and not pd.isna(variance) else 0.0
                        intercept = rtm_df['Score Delta'].mean() - (slope * rtm_df['Obtained Marks_BL'].mean()) if not pd.isna(slope) else 0.0
                        total_rtm = len(rtm_df)
                        improving_pct = (len(rtm_df[rtm_df['Score Delta'] > 0]) / total_rtm) * 100
                        declining_pct = (len(rtm_df[rtm_df['Score Delta'] < 0]) / total_rtm) * 100
                        
                        if slope <= -0.3:
                            rtm_tag = "Strong RTM detected"
                        elif slope <= -0.1:
                            rtm_tag = "Moderate RTM"
                        elif slope < 0:
                            rtm_tag = "Minimal RTM"
                        else:
                            rtm_tag = "No RTM detected"
                            
                        kpi_col1, kpi_col2, kpi_col3, kpi_col4 = st.columns(4)
                        kpi_col1.metric("Correlation (r)", f"{correlation:.3f}" if not pd.isna(correlation) else "N/A")
                        kpi_col2.metric("Regression Slope", f"{slope:.3f}")
                        kpi_col3.metric("Improving vs Declining", f"{improving_pct:.1f}% / {declining_pct:.1f}%")
                        kpi_col4.metric("Interpretation", rtm_tag)
                        
                        if slope <= -0.1:
                            st.warning("💡 **Important Insight:** Part of the observed improvement may be due to statistical regression to the mean rather than pure intervention impact. Students who scored exceptionally low on the baseline naturally tend to score closer to the average on the endline.")
                        else:
                            st.success("💡 **Important Insight:** Observed improvements are less likely driven by RTM. The growth seen across the cohort is more likely attributable to the actual impact of the educational intervention.")
                        st.markdown("---")
                        
                        st.markdown("#### Core RTM View (Scatter Plot)")
                        st.caption("**Interpretation:** A **negative slope** (trendline going down) suggests the RTM effect is present — meaning students with lower baseline scores tended to improve more, while those with higher baseline scores improved less or dropped.")
                        fig_rtm = px.scatter(
                            rtm_df, x="Obtained Marks_BL", y="Score Delta", 
                            trendline="ols", trendline_color_override="red", opacity=0.6,
                            color_discrete_sequence=["#636EFA"],
                            labels={
                                "Obtained Marks_BL": "Baseline Score (Z-Score)" if normalize_rtm else "Baseline Score",
                                "Score Delta": "Score Delta (Z-Score)" if normalize_rtm else "Score Delta (Endline - Baseline)"
                            }
                        )
                        fig_rtm.add_hline(y=0, line_dash="dash", line_color="black", annotation_text="No Change (Delta = 0)", annotation_position="bottom right")
                        fig_rtm.update_layout(margin=dict(l=0, r=0, t=30))
                        st.plotly_chart(fig_rtm, width='stretch')
                        st.markdown("---")
                        
                        st.markdown("#### Binned Analysis (Quintiles)")
                        st.caption("**Interpretation:** Students are grouped into 5 equal-sized bins based on their initial Baseline scores. If RTM is present, the chart will typically show a 'staircase' pattern: large positive deltas on the left (lowest scorers improving the most) and smaller or negative deltas on the right (highest scorers plateauing or declining).")
                        try:
                            rtm_df['BL_Quintile'] = pd.qcut(rtm_df['Obtained Marks_BL'], q=5, duplicates='drop')
                            binned_stats = rtm_df.groupby('BL_Quintile', observed=False).agg(
                                Avg_BL_Score=('Obtained Marks_BL', 'mean'),
                                Avg_Score_Delta=('Score Delta', 'mean'),
                                Student_Count=('Student ID', 'count')
                            ).reset_index()
                            binned_stats['BL_Quintile_Str'] = binned_stats['BL_Quintile'].astype(str)
                            binned_stats = binned_stats.sort_values('Avg_BL_Score')
                            fig_binned = px.bar(
                                binned_stats, x='BL_Quintile_Str', y='Avg_Score_Delta',
                                text=binned_stats['Avg_Score_Delta'].apply(lambda x: f"{x:+.2f}"),
                                color='Avg_Score_Delta',
                                color_continuous_scale=px.colors.diverging.RdYlGn,
                                color_continuous_midpoint=0,
                                labels={"BL_Quintile_Str": "Baseline Score Range (Quintiles)", "Avg_Score_Delta": "Average Score Delta"},
                                hover_data={"Student_Count": True, "Avg_BL_Score": ':.2f', "Avg_Score_Delta": ':.2f'}
                            )
                            fig_binned.add_hline(y=0, line_dash="solid", line_color="black", line_width=1)
                            fig_binned.update_traces(textposition='outside')
                            fig_binned.update_layout(margin=dict(l=0, r=0, t=30, b=40), coloraxis_showscale=False)
                            st.plotly_chart(fig_binned, width='stretch')
                        except ValueError:
                            st.info("Not enough variance in Baseline scores to generate quintile bins for this selection.")
                        st.markdown("---")
                        
                        st.markdown("#### 🧮 Statistical Validation")
                        st.caption("Quantifying the strength of the Regression to the Mean effect using linear regression.")
                        r_squared = correlation ** 2 if not pd.isna(correlation) else 0.0
                        stat_col1, stat_col2, stat_col3 = st.columns(3)
                        stat_col1.metric("Correlation (r)", f"{correlation:.3f}" if not pd.isna(correlation) else "N/A")
                        stat_col2.metric("Regression Slope (b)", f"{slope:.3f}")
                        stat_col3.metric("R-squared (R²)", f"{r_squared:.3f}")
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

            # ==========================================
            # 📥 DRM REPORT GENERATION (PPTX)
            # ==========================================
            with st.sidebar:
                st.markdown("---")
                st.markdown("### 📄 DRM Compliance Report")
                if selected_donors != "All":
                    report_name = f"AY25-26_Impact_Report_{selected_donors.replace(' ', '_')}.pptx"
                    
                    if st.button(f"⚙️ Prepare PPTX for {selected_donors}", use_container_width=True):
                        with st.spinner("Compiling charts and generating presentation..."):
                            try:
                                from pptx import Presentation
                                from pptx.util import Inches, Pt
                                import io

                                prs = Presentation()

                                # ── Helper: abbreviate state name to initials ─────────────
                                def state_code(s):
                                    words = str(s).split()
                                    return "".join(w[0].upper() for w in words) if len(words) > 1 else str(s)[:2].upper()

                                # ── Helper: render fig to PNG with axes visible ────────────
                                def fig_to_png(fig):
                                    """Apply axis visibility fixes and return BytesIO PNG."""
                                    fig.update_yaxes(
                                        showline=True, linecolor="black", linewidth=1,
                                        showticklabels=True, ticks="outside"
                                    )
                                    fig.update_xaxes(
                                        showline=True, linecolor="black", linewidth=1,
                                        showticklabels=True, ticks="outside"
                                    )
                                    img_stream = io.BytesIO()
                                    fig.write_image(img_stream, format="png", engine="kaleido", width=1000, height=550)
                                    img_stream.seek(0)
                                    return img_stream

                                # ── Helper: add a chart slide (Title Only = layout 5) ─────
                                def add_chart_slide(fig, title_text):
                                    chart_slide = prs.slides.add_slide(prs.slide_layouts[5])
                                    chart_slide.shapes.title.text = title_text
                                    img_stream = fig_to_png(fig)
                                    chart_slide.shapes.add_picture(img_stream, Inches(0.5), Inches(1.5), width=Inches(9))

                                # ── SLIDE 1: Title Slide ──────────────────────────────────
                                slide1 = prs.slides.add_slide(prs.slide_layouts[0])
                                slide1.shapes.title.text = "AY 25-26 Impact Report"
                                slide1.placeholders[1].text = (
                                    f"Donor: {selected_donors}\n"
                                    f"Generated automatically via Streamlit"
                                )

                                # ── SLIDE 2: Executive Summary ────────────────────────────
                                slide2 = prs.slides.add_slide(prs.slide_layouts[1])
                                slide2.shapes.title.text = "Executive Summary"
                                tf = slide2.placeholders[1].text_frame
                                tf.word_wrap = True

                                num_schools = filtered_df['Centre Name'].nunique()
                                subjects_assessed = ", ".join(sorted(filtered_df['Subject'].dropna().unique()))

                                # States as initials/codes
                                states_in_data = sorted(filtered_df['State'].dropna().unique())
                                states_str = ", ".join(state_code(s) for s in states_in_data)

                                tf.text = f"States covered: {states_str}"

                                p_schools = tf.add_paragraph()
                                p_schools.text = f"Total Centres Impacted: {num_schools}"

                                p_subj_hdr = tf.add_paragraph()
                                p_subj_hdr.text = f"Subjects Assessed: {subjects_assessed}"

                                # Subject-wise Endline student distribution (replaces Grade-wise)
                                p_dist = tf.add_paragraph()
                                p_dist.text = "Subject-wise Student Distribution (Endline):"
                                if not end_df.empty and 'Subject' in end_df.columns:
                                    subj_counts = (
                                        end_df.drop_duplicates(subset=['Student ID', 'Subject'])
                                        ['Subject'].value_counts().sort_index()
                                    )
                                    for subj, count in subj_counts.items():
                                        p_s = tf.add_paragraph()
                                        p_s.text = f"{subj}: {count} students"
                                        p_s.level = 1
                                else:
                                    p_na = tf.add_paragraph()
                                    p_na.text = "No endline data available."
                                    p_na.level = 1

                                # ── SLIDE 3: Aggregated R.I.S.E shift chart ───────────────
                                if 'fig_rise' in locals():
                                    add_chart_slide(fig_rise, "Overall R.I.S.E Category Shift (BL vs EL)")

                                # ── SLIDES: Subject-wise BL & EL R.I.S.E by Grade ─────────
                                # For each subject generates 2 slides: BL then EL
                                all_subjects = sorted(filtered_df['Subject'].dropna().unique()) \
                                    if 'Subject' in filtered_df.columns else []

                                for subj in all_subjects:
                                    subj_base = base_df[base_df['Subject'].astype(str) == subj] \
                                        if 'Subject' in base_df.columns else pd.DataFrame()
                                    subj_end  = end_df[end_df['Subject'].astype(str) == subj] \
                                        if 'Subject' in end_df.columns else pd.DataFrame()

                                    # Baseline R.I.S.E by Grade — this subject
                                    if not subj_base.empty and 'Grade' in subj_base.columns:
                                        grp_bl = subj_base.groupby(['Grade', 'Category']).size().reset_index(name='Count')
                                        grp_bl['Percentage'] = grp_bl.groupby('Grade')['Count'].transform(
                                            lambda x: x / x.sum() * 100
                                        )
                                        fig_bl_subj = px.bar(
                                            grp_bl, x="Grade", y="Percentage", color="Category",
                                            color_discrete_map=RISE_COLORS,
                                            text=grp_bl['Percentage'].apply(lambda x: f'{x:.1f}%' if x > 5 else ''),
                                            category_orders={"Category": ["Reviving", "Initiating", "Shaping", "Evolving"]}
                                        )
                                        fig_bl_subj.update_layout(
                                            barmode='stack', yaxis_title="% of Students",
                                            margin=dict(l=0, r=0, t=50), showlegend=True,
                                            legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, title="")
                                        )
                                        add_chart_slide(fig_bl_subj, f"Baseline R.I.S.E by Grade — {subj}")

                                    # Endline R.I.S.E by Grade — this subject
                                    if not subj_end.empty and 'Grade' in subj_end.columns:
                                        grp_el = subj_end.groupby(['Grade', 'Category']).size().reset_index(name='Count')
                                        grp_el['Percentage'] = grp_el.groupby('Grade')['Count'].transform(
                                            lambda x: x / x.sum() * 100
                                        )
                                        fig_el_subj = px.bar(
                                            grp_el, x="Grade", y="Percentage", color="Category",
                                            color_discrete_map=RISE_COLORS,
                                            text=grp_el['Percentage'].apply(lambda x: f'{x:.1f}%' if x > 5 else ''),
                                            category_orders={"Category": ["Reviving", "Initiating", "Shaping", "Evolving"]}
                                        )
                                        fig_el_subj.update_layout(
                                            barmode='stack', yaxis_title="% of Students",
                                            margin=dict(l=0, r=0, t=50), showlegend=True,
                                            legend=dict(orientation="h", yanchor="top", y=-0.2, xanchor="center", x=0.5, title="")
                                        )
                                        add_chart_slide(fig_el_subj, f"Endline R.I.S.E by Grade — {subj}")

                                # ── Remaining summary charts ──────────────────────────────
                                if 'fig_box' in locals():
                                    add_chart_slide(fig_box, "Score Distribution (Box Plot)")
                                if 'fig_gen_avg' in locals():
                                    add_chart_slide(fig_gen_avg, "Average Score Trend by Gender")

                                # ── Save ─────────────────────────────────────────────────
                                ppt_stream = io.BytesIO()
                                prs.save(ppt_stream)
                                ppt_stream.seek(0)
                                st.session_state['ready_ppt'] = ppt_stream.getvalue()
                                st.session_state['ready_ppt_donor'] = selected_donors
                                st.success("Report prepared successfully!")

                            except ImportError:
                                st.error("⚠️ Missing required libraries! Please run: pip install python-pptx kaleido")
                            except Exception as e:
                                st.error(f"⚠️ Error preparing presentation: {e}")
                                
                    if 'ready_ppt' in st.session_state and st.session_state.get('ready_ppt_donor') == selected_donors:
                        st.download_button(
                            label="⬇️ Download Presentation",
                            data=st.session_state['ready_ppt'],
                            file_name=report_name,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
                else:
                    st.info("💡 Select a specific Donor from the global filters to enable the DRM Report generator.")

else:
    st.error(f"⚠️ Data file not found! Please ensure `{DATA_FILE}` is placed in the same folder as this script to populate the dashboard.")
    st.image("https://cdn-icons-png.flaticon.com/512/7264/7264168.png", width=150)
