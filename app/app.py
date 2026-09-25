import os
import sys
import shutil
from datetime import date, datetime

# =========================================================
# PROJECT PATH
# =========================================================

PROJECT_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..")
)

if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

# =========================================================
# THIRD-PARTY IMPORTS
# =========================================================

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import silhouette_score
from sklearn.decomposition import PCA

# =========================================================
# LOCAL IMPORTS
# =========================================================

from src.data_preprocessing import FEATURES

from src.auth_db import (
    init_db,
    authenticate,
    create_user,
    list_users,
    get_user_by_username,
    save_snapshot,
    get_student_history,
    add_intervention,
    list_interventions,
)

from src.excel_reports import dataframe_to_xlsx


# =========================================================
# PAGE CONFIG
# =========================================================

st.set_page_config(
    page_title="EduPulse AI | Student Intelligence",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# PATHS
# =========================================================

BASE_DIR = PROJECT_ROOT

DATA_PATH = os.path.join(
    BASE_DIR,
    "data",
    "raw",
    "student_performance.csv",
)


# =========================================================
# DATA LOADING
# =========================================================

@st.cache_data(show_spinner=False)
def load_dataset():

    if not os.path.exists(DATA_PATH):
        st.error(
            "Academic dataset was not found.\n\n"
            f"Expected file:\n{DATA_PATH}"
        )
        st.stop()

    # -----------------------------------------------------
    # Read CSV explicitly as UTF-8
    # -----------------------------------------------------
    try:
        df = pd.read_csv(
            DATA_PATH,
            encoding="utf-8",
            dtype={"Student_ID": "string"},
        )

    except UnicodeDecodeError:
        st.error(
            "The academic CSV file is not encoded as valid UTF-8."
        )
        st.stop()

    except Exception as exc:
        st.error(
            f"Could not read the academic dataset:\n{exc}"
        )
        st.stop()

    # -----------------------------------------------------
    # Required Student_ID column
    # -----------------------------------------------------
    if "Student_ID" not in df.columns:
        st.error(
            "Dataset must contain a 'Student_ID' column."
        )
        st.stop()

    # -----------------------------------------------------
    # Required ML features
    # -----------------------------------------------------
    missing_features = [
        feature
        for feature in FEATURES
        if feature not in df.columns
    ]

    if missing_features:
        st.error(
            "Dataset is missing required columns:\n"
            + ", ".join(missing_features)
        )
        st.stop()

    # -----------------------------------------------------
    # Clean Student IDs
    # -----------------------------------------------------
    df["Student_ID"] = (
        df["Student_ID"]
        .astype("string")
        .str.strip()
    )

    # Convert pandas StringDtype to normal Python strings.
    df["Student_ID"] = df["Student_ID"].fillna("").astype(str)

    # -----------------------------------------------------
    # Detect invalid characters BEFORE Streamlit rendering
    # -----------------------------------------------------
    invalid_id_mask = df["Student_ID"].str.contains(
        r"[\x00-\x08\x0B\x0C\x0E-\x1F\x7F�]",
        regex=True,
        na=False,
    )

    if invalid_id_mask.any():

        invalid_ids = df.loc[
            invalid_id_mask,
            "Student_ID"
        ].tolist()

        st.error(
            "Invalid characters were detected in Student_ID."
        )

        st.write("Affected Student IDs:")
        st.write(invalid_ids)

        st.stop()

    # -----------------------------------------------------
    # Validate duplicate / empty Student IDs
    # -----------------------------------------------------
    empty_ids = df["Student_ID"].eq("")

    if empty_ids.any():
        st.error(
            f"{int(empty_ids.sum())} student record(s) "
            "have an empty Student_ID."
        )
        st.stop()

    duplicate_ids = df[
        df["Student_ID"].duplicated(keep=False)
    ]["Student_ID"].tolist()

    if duplicate_ids:
        st.error(
            "Duplicate Student_ID values were detected."
        )
        st.write(
            sorted(set(duplicate_ids))
        )
        st.stop()

    # -----------------------------------------------------
    # Convert ML features to numeric
    # -----------------------------------------------------
    df[FEATURES] = (
        df[FEATURES]
        .apply(
            pd.to_numeric,
            errors="coerce",
        )
    )

    # -----------------------------------------------------
    # Check missing / invalid numeric values
    # -----------------------------------------------------
    missing_before_fill = (
        df[FEATURES]
        .isna()
        .sum()
    )

    total_invalid = int(
        missing_before_fill.sum()
    )

    if total_invalid > 0:

        # Fill invalid/missing values using
        # the corresponding feature median.
        for feature in FEATURES:

            median_value = df[feature].median()

            if pd.isna(median_value):
                st.error(
                    f"Feature '{feature}' contains "
                    "no valid numeric values."
                )
                st.stop()

            df[feature] = (
                df[feature]
                .fillna(median_value)
            )

    # -----------------------------------------------------
    # Final numeric validation
    # -----------------------------------------------------
    if not np.isfinite(
        df[FEATURES].to_numpy(
            dtype=float
        )
    ).all():

        st.error(
            "The dataset contains invalid "
            "infinite numeric values."
        )
        st.stop()

    # -----------------------------------------------------
    # Final serialization safety check
    # -----------------------------------------------------
    # Convert object/string columns into clean strings
    # before Streamlit/Plotly receives the dataframe.
    for column in df.columns:

        if (
            df[column].dtype == "object"
            or str(df[column].dtype) == "string"
        ):
            df[column] = (
                df[column]
                .fillna("")
                .astype(str)
            )

    # -----------------------------------------------------
    # Reset index
    # -----------------------------------------------------
    df = df.reset_index(drop=True)

    return df


def refresh_dataset():

    st.cache_data.clear()

    st.rerun()

def refresh_dataset():
    st.cache_data.clear()
    st.rerun()


# =========================================================
# LOAD DATA + INITIALIZE DATABASE
# =========================================================

df = load_dataset()

student_ids = (
    df["Student_ID"]
    .dropna()
    .astype(str)
    .tolist()
)

init_db(student_ids)


# =========================================================
# SESSION STATE
# =========================================================

SESSION_DEFAULTS = {
    "logged_in": False,
    "role": None,
    "username": None,
    "user": None,
}

for key, default in SESSION_DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = default


# =========================================================
# DATABASE COMPATIBILITY HELPERS
# =========================================================
#
# These wrappers use the functions that actually exist in
# your current auth_db.py:
#
# authenticate()
# create_user()
# =========================================================


def authenticate_account(username, password, role):
    """
    Authenticate against the existing SQLite auth_db API.

    The current backend exposes `authenticate`, not the old
    `authenticate_user` function.
    """

    try:
        # Current backend API.
        return authenticate(username, password, role)
    except TypeError:
        # Compatibility fallback if the backend's authenticate
        # function accepts only username/password.
        try:
            account = authenticate(username, password)

            if account and account.get("role") == role:
                return account

            return None

        except Exception:
            return None

    except Exception:
        return None


def create_student_account_compat(
    username,
    password,
    student_id,
    full_name,
    email="",
    created_by=None,
):
    """
    Create a Student account using the existing create_user()
    function.

    The helper supports common versions of the current
    create_user API without requiring a second auth_db file.
    """

    username = str(username).strip()
    password = str(password)
    student_id = str(student_id).strip()
    full_name = str(full_name).strip()
    email = str(email or "").strip()

    if not username:
        return False, "Username is required."

    if len(password) < 6:
        return False, "Password must contain at least 6 characters."

    if not student_id:
        return False, "Student ID is required."

    if not full_name:
        return False, "Full name is required."

    try:
        # Most complete version of the current API.
        result = create_user(
            username=username,
            password=password,
            role="Student",
            student_id=student_id,
            full_name=full_name,
            email=email,
            created_by=created_by,
        )

        if isinstance(result, tuple):
            return result

        if isinstance(result, bool):
            return (
                result,
                "Student account created successfully."
                if result
                else "Could not create student account.",
            )

        if isinstance(result, dict):
            return (
                True,
                "Student account created successfully.",
            )

        return True, "Student account created successfully."

    except TypeError:
        # Try positional form used by some versions.
        try:
            result = create_user(
                username,
                password,
                "Student",
                student_id,
                full_name,
                email,
                created_by,
            )

            if isinstance(result, tuple):
                return result

            if isinstance(result, bool):
                return (
                    result,
                    "Student account created successfully."
                    if result
                    else "Could not create student account.",
                )

            return True, "Student account created successfully."

        except TypeError:
            pass

        # Try a minimal API.
        try:
            result = create_user(
                username,
                password,
                "Student",
                student_id,
                full_name,
                email,
            )

            if isinstance(result, tuple):
                return result

            if isinstance(result, bool):
                return (
                    result,
                    "Student account created successfully."
                    if result
                    else "Could not create student account.",
                )

            return True, "Student account created successfully."

        except Exception as exc:
            return False, f"Could not create account: {exc}"

    except Exception as exc:
        return False, f"Could not create account: {exc}"


# =========================================================
# LOGIN SCREEN
# =========================================================

def login_screen():

    st.markdown(
        """
        <div style="
            text-align:center;
            padding:35px 0 15px;
        ">
            <h1>🎓 EduPulse AI</h1>
            <p style="
                font-size:18px;
                opacity:0.8;
            ">
                Academic Intelligence & Early Warning Platform
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    _, center, _ = st.columns([1, 2, 1])

    with center:

        role = st.radio(
            "Select Portal",
            ["Student", "Teacher"],
            horizontal=True,
        )

        username = st.text_input(
            "Username",
            placeholder="Enter username",
        )

        password = st.text_input(
            "Password",
            type="password",
            placeholder="Enter password",
        )

        if st.button(
            "🚀 Secure Login",
            use_container_width=True,
            type="primary",
        ):

            selected_role = (
                "Student"
                if role == "Student"
                else "Teacher"
            )

            account = authenticate_account(
                username.strip(),
                password,
                selected_role,
            )

            # Admin can enter through Teacher portal.
            if account is None and role == "Teacher":
                account = authenticate_account(
                    username.strip(),
                    password,
                    "Admin",
                )

            if account:

                st.session_state.logged_in = True
                st.session_state.role = account.get("role")
                st.session_state.username = account.get("username")
                st.session_state.user = account

                st.rerun()

            else:

                st.error(
                    "❌ Invalid username or password "
                    "for the selected portal."
                )

        st.divider()

        st.markdown("#### 🧪 Demo Accounts")

        st.info(
            "👨‍🎓 **Student:** `student001` / `student123`\n\n"
            "👨‍🏫 **Teacher:** `teacher001` / `teacher123`\n\n"
            "🛡️ **Admin:** `admin` / `admin123`"
        )

        # =================================================
        # STUDENT SELF REGISTRATION
        # =================================================

        with st.expander("📝 Student Self-Registration"):

            name = st.text_input(
                "Full name",
                key="reg_name",
            )

            sid = st.text_input(
                "Student ID",
                key="reg_sid",
            )

            new_username = st.text_input(
                "New username",
                key="reg_user",
            )

            new_password = st.text_input(
                "New password",
                type="password",
                key="reg_pw",
            )

            email = st.text_input(
                "Email (optional)",
                key="reg_email",
            )

            if st.button(
                "Create Student Account",
                use_container_width=True,
            ):

                valid_ids = set(
                    df["Student_ID"].astype(str)
                )

                sid_clean = sid.strip()

                if sid_clean not in valid_ids:

                    st.error(
                        "Student ID was not found "
                        "in the academic dataset."
                    )

                elif len(new_password) < 6:

                    st.error(
                        "Password must contain at least "
                        "6 characters."
                    )

                elif not name.strip():

                    st.error(
                        "Full name is required."
                    )

                elif not new_username.strip():

                    st.error(
                        "Username is required."
                    )

                else:

                    ok, message = create_student_account_compat(
                        new_username,
                        new_password,
                        sid_clean,
                        name.strip(),
                        email.strip(),
                        "self-registration",
                    )

                    if ok:
                        st.success(message)
                    else:
                        st.error(message)


# =========================================================
# SHOW LOGIN OR STOP
# =========================================================

if not st.session_state.logged_in:

    login_screen()

    st.stop()


# =========================================================
# ML / SCORING FUNCTIONS
# =========================================================

@st.cache_data(show_spinner=False)
def cluster_data(k):

    scaler = StandardScaler()

    X = scaler.fit_transform(
        df[FEATURES]
    )

    model = KMeans(
        n_clusters=k,
        n_init=20,
        random_state=42,
    )

    labels = model.fit_predict(X)

    clustered_df = df.copy()

    clustered_df["Cluster"] = labels

    means = clustered_df.groupby(
        "Cluster"
    )[FEATURES].mean()

    performance_score = (
        means["Attendance"] * 0.20
        + means["Assignment_Score"] * 0.25
        + means["Internal_Marks"] * 0.25
        + means["Previous_GPA"] * 10 * 0.30
    )

    ordered_clusters = list(
        performance_score
        .sort_values()
        .index
    )

    names = [
        "Needs Improvement",
        "Below Average",
        "Average",
        "Good Performer",
        "High Performer",
        "Excellent",
        "Outstanding",
        "Elite",
    ]

    mapping = {
        cluster: names[
            min(index, len(names) - 1)
        ]
        for index, cluster
        in enumerate(ordered_clusters)
    }

    clustered_df["Performance_Group"] = (
        clustered_df["Cluster"]
        .map(mapping)
    )

    return clustered_df, mapping


def academic_score(row):

    score = (
        row["Attendance"] * 0.20
        + row["Assignment_Score"] * 0.25
        + row["Internal_Marks"] * 0.25
        + row["Previous_GPA"] * 10 * 0.30
    )

    return float(
        np.clip(score, 0, 100)
    )


def risk_label(row):

    score = academic_score(row)

    if (
        score < 60
        or row["Attendance"] < 65
    ):
        return "🔴 High Risk"

    if (
        score < 75
        or row["Attendance"] < 75
    ):
        return "🟠 Watch"

    return "🟢 On Track"


def warning_probability(row):
    """
    Explainable early-warning score.

    IMPORTANT:
    This is a heuristic score, not a trained probability model.
    """

    deficits = [
        max(
            0,
            65 - row["Attendance"],
        ) / 65,

        max(
            0,
            60 - row["Assignment_Score"],
        ) / 60,

        max(
            0,
            60 - row["Internal_Marks"],
        ) / 60,

        max(
            0,
            6.5 - row["Previous_GPA"],
        ) / 6.5,

        max(
            0,
            2 - row["Study_Hours"],
        ) / 2,
    ]

    base = (
        100
        - academic_score(row)
    )

    probability = np.clip(
        base * 0.55
        + sum(deficits) * 12,
        0,
        100,
    )

    return float(probability)


def risk_reasons(row):

    reasons = []

    if row["Attendance"] < 65:
        reasons.append(
            f"Attendance is {row['Attendance']:.1f}%"
        )

    if row["Assignment_Score"] < 60:
        reasons.append(
            f"Assignment score is "
            f"{row['Assignment_Score']:.1f}"
        )

    if row["Internal_Marks"] < 60:
        reasons.append(
            f"Internal marks are "
            f"{row['Internal_Marks']:.1f}"
        )

    if row["Previous_GPA"] < 6.5:
        reasons.append(
            f"GPA is {row['Previous_GPA']:.2f}"
        )

    if row["Study_Hours"] < 2:
        reasons.append(
            f"Study time is "
            f"{row['Study_Hours']:.1f} hrs"
        )

    return (
        reasons
        or ["No major warning indicators detected"]
    )


def action_for(row):

    risk = row.get(
        "Risk",
        risk_label(row),
    )

    if risk.startswith("🔴"):
        return (
            "Schedule intervention + "
            "attendance monitoring + "
            "weekly follow-up"
        )

    if risk.startswith("🟠"):
        return (
            "Targeted feedback + "
            "monitor next evaluation"
        )

    return "Continue regular mentoring"


def styled_status(status):

    if "High Risk" in status:
        return "🔴 HIGH RISK"

    if "Watch" in status:
        return "🟠 WATCH"

    return "🟢 ON TRACK"


# =========================================================
# SIDEBAR
# =========================================================

with st.sidebar:

    st.markdown("## 🎓 EduPulse AI")

    st.caption(
        f"{st.session_state.role} Portal • "
        f"{st.session_state.username}"
    )

    st.divider()

    if st.session_state.role in (
        "Teacher",
        "Admin",
    ):

        st.markdown(
            "**Teacher Command Center**"
        )

        st.caption(
            "Monitor • Analyze • "
            "Intervene • Report"
        )

    else:

        st.markdown(
            "**Student Success Center**"
        )

        st.caption(
            "Track • Improve • Grow"
        )

    st.divider()

    if st.button(
        "🚪 Logout",
        use_container_width=True,
    ):

        st.session_state.logged_in = False
        st.session_state.role = None
        st.session_state.username = None
        st.session_state.user = None

        st.rerun()


# =========================================================
# STUDENT PORTAL
# =========================================================

if st.session_state.role == "Student":

    account = (
        st.session_state.user
        or get_user_by_username(
            st.session_state.username
        )
    )

    sid = ""

    if account:
        sid = str(
            account.get("student_id") or ""
        ).strip()

    valid_student_ids = set(
        df["Student_ID"].astype(str)
    )

    if (
        not sid
        or sid not in valid_student_ids
    ):

        st.error(
            "Your account is not linked to "
            "a valid Student ID. "
            "Please contact a teacher/admin."
        )

        st.stop()

    student_rows = df[
        df["Student_ID"] == sid
    ]

    if student_rows.empty:

        st.error(
            "Student academic record could not "
            "be found."
        )

        st.stop()

    student = student_rows.iloc[0]

    clustered, _ = cluster_data(3)

    cluster_rows = clustered[
        clustered["Student_ID"] == sid
    ]

    if cluster_rows.empty:

        st.error(
            "Student clustering record could "
            "not be generated."
        )

        st.stop()

    cluster_student = cluster_rows.iloc[0]

    score = academic_score(student)

    risk = risk_label(student)

    warning_score = warning_probability(student)

    history = pd.DataFrame(
        get_student_history(sid)
    )

    interventions = pd.DataFrame(
        list_interventions(sid)
    )

    # =====================================================
    # HEADER
    # =====================================================

    st.title(
        "👨‍🎓 Student Success Center"
    )

    st.caption(
        f"Welcome, "
        f"{account.get('full_name') or st.session_state.username}"
        f" • Student ID: {sid}"
    )

    # =====================================================
    # KPI CARDS
    # =====================================================

    a, b, c, d, e = st.columns(5)

    a.metric(
        "Academic Index",
        f"{score:.1f}/100",
    )

    b.metric(
        "Attendance",
        f"{student['Attendance']:.1f}%",
    )

    c.metric(
        "Previous GPA",
        f"{student['Previous_GPA']:.2f}",
    )

    d.metric(
        "Early-Warning Score",
        f"{warning_score:.0f}%",
    )

    e.metric(
        "Status",
        styled_status(risk),
    )

    st.divider()

    tabs = st.tabs(
        [
            "📊 My Dashboard",
            "📈 My Trends",
            "🤖 Improvement Plan",
            "📝 Teacher Feedback",
            "📥 My Report",
        ]
    )

    # =====================================================
    # STUDENT DASHBOARD
    # =====================================================

    with tabs[0]:

        left, right = st.columns(2)

        with left:

            st.subheader(
                "📈 Performance Radar"
            )

            values = [
                student["Study_Hours"] * 10,
                student["Attendance"],
                student["Assignment_Score"],
                student["Internal_Marks"],
                student["Previous_GPA"] * 10,
            ]

            categories = [
                "Study",
                "Attendance",
                "Assignments",
                "Internal",
                "GPA",
            ]

            radar_fig = go.Figure(
                go.Scatterpolar(
                    r=values + [values[0]],
                    theta=categories + [categories[0]],
                    fill="toself",
                    name="My Profile",
                )
            )

            radar_fig.update_layout(
                polar=dict(
                    radialaxis=dict(
                        range=[0, 100]
                    )
                ),
                height=420,
                showlegend=False,
            )

            st.plotly_chart(
                radar_fig,
                use_container_width=True,
            )

        with right:

            st.subheader(
                "🎯 Indicator Breakdown"
            )

            indicators = pd.DataFrame(
                {
                    "Indicator": [
                        "Study Hours",
                        "Attendance",
                        "Assignment",
                        "Internal Marks",
                        "GPA",
                    ],
                    "Score": [
                        student["Study_Hours"] * 10,
                        student["Attendance"],
                        student["Assignment_Score"],
                        student["Internal_Marks"],
                        student["Previous_GPA"] * 10,
                    ],
                }
            )

            fig = px.bar(
                indicators,
                x="Indicator",
                y="Score",
                text="Score",
                title="Current Academic Indicators",
            )

            fig.update_yaxes(
                range=[0, 100]
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
            )

        st.info(
            f"**Your K-Means segment:** "
            f"{cluster_student['Performance_Group']} "
            f"• Cluster {cluster_student['Cluster']}"
        )

    # =====================================================
    # STUDENT TRENDS
    # =====================================================

    with tabs[1]:

        st.subheader(
            "📈 Historical Performance"
        )

        if not history.empty:

            if "captured_at" in history.columns:

                history["captured_at"] = pd.to_datetime(
                    history["captured_at"],
                    errors="coerce",
                )

            trend_options = [
                "academic_index",
                "attendance",
                "assignment_score",
                "internal_marks",
                "previous_gpa",
            ]

            available_metrics = [
                metric
                for metric in trend_options
                if metric in history.columns
            ]

            if available_metrics:

                metric = st.selectbox(
                    "Trend metric",
                    available_metrics,
                )

                fig = px.line(
                    history,
                    x="captured_at",
                    y=metric,
                    markers=True,
                    title=(
                        metric
                        .replace("_", " ")
                        .title()
                        + " Trend"
                    ),
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True,
                )

                st.caption(
                    "Trend history is created when "
                    "a teacher captures an academic "
                    "snapshot after an update."
                )

            else:

                st.warning(
                    "Historical data exists, but no "
                    "supported trend metrics were found."
                )

        else:

            st.info(
                "No historical snapshots yet. "
                "Your teacher can capture a snapshot "
                "after each evaluation/update."
            )

    # =====================================================
    # IMPROVEMENT PLAN
    # =====================================================

    with tabs[2]:

        st.subheader(
            "🤖 Personalized Improvement Plan"
        )

        reasons = risk_reasons(student)

        has_warning = False

        if student["Attendance"] < 75:

            has_warning = True

            st.warning(
                "📅 **Attendance:** Target at least "
                "75% and avoid consecutive absences."
            )

        if student["Study_Hours"] < 2:

            has_warning = True

            st.warning(
                "📚 **Study:** Build a consistent "
                "daily study block of 2+ hours."
            )

        if student["Assignment_Score"] < 70:

            has_warning = True

            st.warning(
                "📝 **Assignments:** Complete every "
                "assignment and target 70+."
            )

        if student["Internal_Marks"] < 70:

            has_warning = True

            st.warning(
                "✍️ **Internals:** Focus on revision, "
                "previous questions and weekly practice."
            )

        if student["Previous_GPA"] < 7:

            has_warning = True

            st.warning(
                "🎯 **GPA:** Prioritize weak subjects "
                "and meet your mentor regularly."
            )

        if not has_warning:

            st.success(
                "🌟 Excellent consistency. Maintain "
                "your current routine and help peers "
                "where possible."
            )

        st.write(
            "**Current warning indicators:**"
        )

        for reason in reasons:

            st.write(
                f"• {reason}"
            )

    # =====================================================
    # TEACHER FEEDBACK
    # =====================================================

    with tabs[3]:

        st.subheader(
            "📝 Teacher Feedback & Interventions"
        )

        if interventions.empty:

            st.info(
                "No teacher interventions recorded yet."
            )

        else:

            columns = [
                "reason",
                "action",
                "notes",
                "follow_up_date",
                "status",
                "created_at",
            ]

            available = [
                column
                for column in columns
                if column in interventions.columns
            ]

            st.dataframe(
                interventions[available],
                use_container_width=True,
                hide_index=True,
            )

    # =====================================================
    # STUDENT REPORT
    # =====================================================

    with tabs[4]:

        st.subheader(
            "📥 Personal Excel Report"
        )

        report_df = pd.DataFrame(
            [student]
        ).copy()

        report_df["Academic_Index"] = score
        report_df["Risk"] = risk
        report_df["Early_Warning_Score"] = warning_score
        report_df["Performance_Group"] = (
            cluster_student["Performance_Group"]
        )

        intervention_report = (
            interventions
            if not interventions.empty
            else pd.DataFrame(
                columns=[
                    "student_id",
                    "reason",
                    "action",
                    "notes",
                    "follow_up_date",
                    "status",
                ]
            )
        )

        workbook = dataframe_to_xlsx(
            {
                "My Performance": report_df,
                "Interventions": intervention_report,
            }
        )

        st.download_button(
            "⬇️ Download My .xlsx Report",
            workbook,
            f"{sid}_EduPulse_Report.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# =========================================================
# TEACHER / ADMIN PORTAL
# =========================================================

else:

    clustered, mapping = cluster_data(3)

    clustered["Academic_Index"] = clustered.apply(
        academic_score,
        axis=1,
    )

    clustered["Risk"] = clustered.apply(
        risk_label,
        axis=1,
    )

    clustered["Early_Warning_Score"] = clustered.apply(
        warning_probability,
        axis=1,
    )

    clustered["Teacher_Action"] = clustered.apply(
        action_for,
        axis=1,
    )

    # =====================================================
    # HEADER
    # =====================================================

    st.title(
        "🧑‍🏫 Teacher Intelligence Dashboard"
    )

    st.caption(
        "Monitor • Segment • "
        "Analyze Risk • Intervene • Report"
    )

    # =====================================================
    # ANALYTICS CONTROLS
    # =====================================================

    with st.expander(
        "⚙️ Live Analytics Controls",
        expanded=True,
    ):

        c1, c2, c3, c4 = st.columns(4)

        with c1:

            k = st.slider(
                "Clusters (K)",
                2,
                8,
                3,
            )

        with c2:

            min_attendance = st.slider(
                "Minimum attendance",
                40,
                100,
                40,
            )

        with c3:

            risk_filter = st.multiselect(
                "Risk levels",
                [
                    "🔴 High Risk",
                    "🟠 Watch",
                    "🟢 On Track",
                ],
                default=[
                    "🔴 High Risk",
                    "🟠 Watch",
                    "🟢 On Track",
                ],
            )

        with c4:

            initial_groups = sorted(
                clustered["Performance_Group"]
                .dropna()
                .unique()
                .tolist()
            )

            group_filter = st.multiselect(
                "Performance groups",
                initial_groups,
                default=initial_groups,
            )

    # =====================================================
    # RECALCULATE CLUSTERS
    # =====================================================

    clustered, mapping = cluster_data(k)

    clustered["Academic_Index"] = clustered.apply(
        academic_score,
        axis=1,
    )

    clustered["Risk"] = clustered.apply(
        risk_label,
        axis=1,
    )

    clustered["Early_Warning_Score"] = clustered.apply(
        warning_probability,
        axis=1,
    )

    clustered["Teacher_Action"] = clustered.apply(
        action_for,
        axis=1,
    )

    view = clustered[
        (clustered["Attendance"] >= min_attendance)
        & clustered["Risk"].isin(risk_filter)
        & clustered["Performance_Group"].isin(group_filter)
    ].copy()

    high_risk = clustered[
        clustered["Risk"].str.startswith("🔴")
    ].copy()

    watch = clustered[
        clustered["Risk"].str.startswith("🟠")
    ].copy()

    on_track = clustered[
        clustered["Risk"].str.startswith("🟢")
    ].copy()

    # =====================================================
    # KPI CARDS
    # =====================================================

    a, b, c, d, e = st.columns(5)

    a.metric(
        "Total Students",
        len(clustered),
    )

    b.metric(
        "🟢 On Track",
        len(on_track),
    )

    c.metric(
        "🟠 Watch",
        len(watch),
    )

    d.metric(
        "🔴 High Risk",
        len(high_risk),
    )

    e.metric(
        "Avg GPA",
        f"{clustered['Previous_GPA'].mean():.2f}",
    )

    st.divider()

    # =====================================================
    # TEACHER TABS
    # =====================================================

    tabs = st.tabs(
        [
            "🏠 Command Center",
            "📊 Live Overview",
            "🚨 Early Warning",
            "🔎 Student Explorer",
            "🧠 Analytics Lab",
            "📝 Interventions",
            "👥 Accounts",
            "📥 Reports",
            "🗂️ Data Management",
        ]
    )

    # =====================================================
    # COMMAND CENTER
    # =====================================================

    with tabs[0]:

        left, right = st.columns(
            [1.3, 1]
        )

        with left:

            st.subheader(
                "🚨 Priority Students"
            )

            priority = (
                high_risk
                .sort_values(
                    "Early_Warning_Score",
                    ascending=False,
                )
                .head(10)
            )

            if priority.empty:

                st.success(
                    "No high-risk students detected."
                )

            else:

                priority_display = priority[
                    [
                        "Student_ID",
                        "Attendance",
                        "Previous_GPA",
                        "Academic_Index",
                        "Early_Warning_Score",
                    ]
                ].copy()

                priority_display[
                    "Early_Warning_Score"
                ] = (
                    priority_display[
                        "Early_Warning_Score"
                    ]
                    .round(0)
                    .astype(int)
                    .astype(str)
                    + "%"
                )

                st.dataframe(
                    priority_display,
                    use_container_width=True,
                    hide_index=True,
                )

        with right:

            st.subheader(
                "📌 Risk Distribution"
            )

            risk_counts = (
                clustered["Risk"]
                .value_counts()
                .rename_axis("Risk")
                .reset_index(name="Students")
            )

            fig = px.pie(
                risk_counts,
                names="Risk",
                values="Students",
                hole=0.45,
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
            )

        left, right = st.columns(2)

        with left:

            st.subheader(
                "📉 Students Needing Attention"
            )

            weak = pd.DataFrame(
                {
                    "Indicator": [
                        "Attendance <65%",
                        "Assignment <60",
                        "Internal <60",
                        "GPA <6.5",
                        "Study <2h",
                    ],
                    "Students": [
                        (
                            clustered["Attendance"] < 65
                        ).sum(),

                        (
                            clustered["Assignment_Score"] < 60
                        ).sum(),

                        (
                            clustered["Internal_Marks"] < 60
                        ).sum(),

                        (
                            clustered["Previous_GPA"] < 6.5
                        ).sum(),

                        (
                            clustered["Study_Hours"] < 2
                        ).sum(),
                    ],
                }
            )

            fig = px.bar(
                weak,
                x="Indicator",
                y="Students",
                text="Students",
                title="Weak-Performance Indicators",
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
            )

        with right:

            st.subheader(
                "📊 Class Average Profile"
            )

            average_profile = pd.DataFrame(
                {
                    "Indicator": [
                        "Attendance",
                        "Assignments",
                        "Internals",
                        "GPA",
                    ],
                    "Average": [
                        clustered["Attendance"].mean(),
                        clustered["Assignment_Score"].mean(),
                        clustered["Internal_Marks"].mean(),
                        clustered["Previous_GPA"].mean() * 10,
                    ],
                }
            )

            fig = px.bar(
                average_profile,
                x="Indicator",
                y="Average",
                text="Average",
                title="Class Average Profile",
            )

            fig.update_yaxes(
                range=[0, 100]
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
            )

    # =====================================================
    # LIVE OVERVIEW
    # =====================================================

    with tabs[1]:

        left, right = st.columns(2)

        with left:

            counts = (
                view["Performance_Group"]
                .value_counts()
                .rename_axis("Group")
                .reset_index(name="Students")
            )

            fig = px.bar(
                counts,
                x="Group",
                y="Students",
                color="Group",
                text="Students",
                title="Live Performance Segmentation",
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
            )

        with right:

            fig = px.scatter(
                view,
                x="Study_Hours",
                y="Previous_GPA",
                color="Performance_Group",
                size="Attendance",
                hover_name="Student_ID",
                hover_data=[
                    "Assignment_Score",
                    "Internal_Marks",
                    "Risk",
                ],
                title="Study Hours vs GPA",
                trendline="ols",
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
            )

        left, right = st.columns(2)

        with left:

            fig = px.histogram(
                view,
                x="Attendance",
                color="Risk",
                nbins=20,
                title="Attendance Distribution",
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
            )

        with right:

            fig = px.box(
                view,
                x="Performance_Group",
                y="Academic_Index",
                color="Performance_Group",
                title="Academic Index by Group",
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
            )

    # =====================================================
    # EARLY WARNING
    # =====================================================

    with tabs[2]:

        st.subheader(
            "🚨 Early Warning System"
        )

        st.caption(
            "The Early-Warning Score is an explainable "
            "heuristic derived from academic deficits. "
            "It is not presented as a trained probability model."
        )

        risk_view = (
            clustered
            .sort_values(
                "Early_Warning_Score",
                ascending=False,
            )
        )

        warning_columns = [
            "Student_ID",
            "Study_Hours",
            "Attendance",
            "Assignment_Score",
            "Internal_Marks",
            "Previous_GPA",
            "Academic_Index",
            "Early_Warning_Score",
            "Performance_Group",
            "Risk",
            "Teacher_Action",
        ]

        st.dataframe(
            risk_view[warning_columns],
            use_container_width=True,
            height=520,
        )

        warning_sid = st.selectbox(
            "Inspect warning reasons",
            clustered["Student_ID"].tolist(),
            key="warning_sid",
        )

        warning_row = clustered[
            clustered["Student_ID"]
            == warning_sid
        ].iloc[0]

        st.write(
            f"### {warning_sid} • "
            f"{warning_row['Risk']} • "
            f"{warning_row['Early_Warning_Score']:.0f}% warning score"
        )

        for reason in risk_reasons(
            warning_row
        ):

            st.write(
                "• " + reason
            )

    # =====================================================
    # STUDENT EXPLORER
    # =====================================================

    with tabs[3]:

        st.subheader(
            "🔎 Individual Student Explorer"
        )

        explorer_sid = st.selectbox(
            "Select Student",
            clustered["Student_ID"].tolist(),
            key="explorer_sid",
        )

        explorer_row = clustered[
            clustered["Student_ID"]
            == explorer_sid
        ].iloc[0]

        a, b, c, d, e = st.columns(5)

        a.metric(
            "Academic Index",
            f"{explorer_row['Academic_Index']:.1f}",
        )

        b.metric(
            "Attendance",
            f"{explorer_row['Attendance']:.1f}%",
        )

        c.metric(
            "GPA",
            f"{explorer_row['Previous_GPA']:.2f}",
        )

        d.metric(
            "Warning Score",
            f"{explorer_row['Early_Warning_Score']:.0f}%",
        )

        e.metric(
            "Risk",
            styled_status(
                explorer_row["Risk"]
            ),
        )

        left, right = st.columns(2)

        with left:

            radar = pd.DataFrame(
                {
                    "Indicator": [
                        "Attendance",
                        "Assignments",
                        "Internal",
                        "GPA",
                    ],
                    "Score": [
                        explorer_row["Attendance"],
                        explorer_row["Assignment_Score"],
                        explorer_row["Internal_Marks"],
                        explorer_row["Previous_GPA"] * 10,
                    ],
                }
            )

            fig = px.line_polar(
                radar,
                r="Score",
                theta="Indicator",
                line_close=True,
                title=(
                    f"{explorer_sid} "
                    "Performance Profile"
                ),
            )

            fig.update_layout(
                polar=dict(
                    radialaxis=dict(
                        range=[0, 100]
                    )
                )
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
            )

        with right:

            student_history = pd.DataFrame(
                get_student_history(
                    explorer_sid
                )
            )

            if not student_history.empty:

                if "captured_at" in student_history.columns:

                    student_history[
                        "captured_at"
                    ] = pd.to_datetime(
                        student_history["captured_at"],
                        errors="coerce",
                    )

                if "academic_index" in student_history.columns:

                    fig = px.line(
                        student_history,
                        x="captured_at",
                        y="academic_index",
                        markers=True,
                        title="Historical Academic Index",
                    )

                    st.plotly_chart(
                        fig,
                        use_container_width=True,
                    )

                else:

                    st.info(
                        "Historical snapshots exist, "
                        "but Academic Index is unavailable."
                    )

            else:

                st.info(
                    "No snapshots yet for this student."
                )

        st.info(
            "**Teacher action:** "
            + action_for(explorer_row)
        )

    # =====================================================
    # ANALYTICS LAB
    # =====================================================

    with tabs[4]:

        st.subheader(
            "🧠 Analytics Lab"
        )

        st.markdown(
            "#### Cluster Quality"
        )

        quality_rows = []

        X = StandardScaler().fit_transform(
            df[FEATURES]
        )

        max_k = min(
            8,
            len(df) - 1,
        )

        if max_k >= 2:

            for kk in range(
                2,
                max_k + 1,
            ):

                labels = KMeans(
                    n_clusters=kk,
                    n_init=20,
                    random_state=42,
                ).fit_predict(X)

                quality_rows.append(
                    {
                        "K": kk,
                        "Silhouette Score": silhouette_score(
                            X,
                            labels,
                        ),
                    }
                )

        quality_df = pd.DataFrame(
            quality_rows
        )

        if not quality_df.empty:

            left, right = st.columns(2)

            with left:

                fig = px.line(
                    quality_df,
                    x="K",
                    y="Silhouette Score",
                    markers=True,
                    title="Silhouette Score by K",
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True,
                )

            with right:

                fig = px.line(
                    quality_df,
                    x="K",
                    y="Silhouette Score",
                    markers=True,
                    title="Optimal K Explorer",
                )

                st.plotly_chart(
                    fig,
                    use_container_width=True,
                )

                best_k_row = quality_df.loc[
                    quality_df["Silhouette Score"].idxmax()
                ]

                st.success(
                    f"Best evaluated K: "
                    f"**{int(best_k_row['K'])}** "
                    f"with Silhouette Score "
                    f"**{best_k_row['Silhouette Score']:.3f}**"
                )

        st.markdown(
            "#### Feature Relationship Explorer"
        )

        f1, f2 = st.columns(2)

        with f1:

            x_feature = st.selectbox(
                "X Feature",
                FEATURES,
                index=0,
                key="analytics_x_feature",
            )

        with f2:

            y_feature = st.selectbox(
                "Y Feature",
                FEATURES,
                index=min(
                    4,
                    len(FEATURES) - 1,
                ),
                key="analytics_y_feature",
            )

        if not view.empty:

            fig = px.scatter(
                view,
                x=x_feature,
                y=y_feature,
                color="Performance_Group",
                hover_name="Student_ID",
                hover_data=[
                    "Attendance",
                    "Previous_GPA",
                    "Academic_Index",
                ],
                trendline="ols",
                title=(
                    f"{x_feature.replace('_', ' ')} "
                    f"vs "
                    f"{y_feature.replace('_', ' ')}"
                ),
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
            )

        correlation = view[
            FEATURES
        ].corr().round(2)

        if not correlation.empty:

            st.plotly_chart(
                px.imshow(
                    correlation,
                    text_auto=True,
                    title="Feature Correlation Heatmap",
                    aspect="auto",
                ),
                use_container_width=True,
            )

        # =================================================
        # PCA
        # =================================================

        if len(view) >= 2:

            st.markdown(
                "#### PCA Student Landscape"
            )

            pca_scaler = StandardScaler()

            z = pca_scaler.fit_transform(
                view[FEATURES]
            )

            pca = PCA(
                n_components=2
            )

            components = pca.fit_transform(z)

            pca_df = view[
                [
                    "Student_ID",
                    "Performance_Group",
                    "Risk",
                ]
            ].copy()

            pca_df["PC1"] = components[:, 0]
            pca_df["PC2"] = components[:, 1]

            explained_variance = (
                pca.explained_variance_ratio_.sum()
            )

            fig = px.scatter(
                pca_df,
                x="PC1",
                y="PC2",
                color="Performance_Group",
                hover_name="Student_ID",
                hover_data=["Risk"],
                title=(
                    "PCA Student Landscape • "
                    f"Explained Variance "
                    f"{explained_variance:.1%}"
                ),
            )

            st.plotly_chart(
                fig,
                use_container_width=True,
            )

        else:

            st.info(
                "At least two students are required "
                "for PCA visualization."
            )

    # =====================================================
    # INTERVENTIONS
    # =====================================================

    with tabs[5]:

        st.subheader(
            "📝 Intervention Management"
        )

        selected_sid = st.selectbox(
            "Student",
            clustered["Student_ID"].tolist(),
            key="intervention_sid",
        )

        selected_row = clustered[
            clustered["Student_ID"]
            == selected_sid
        ].iloc[0]

        st.info(
            f"Current status: "
            f"**{selected_row['Risk']}** • "
            f"{selected_row['Academic_Index']:.1f}/100"
        )

        with st.form(
            "intervention_form"
        ):

            default_reasons = []

            if selected_row["Attendance"] < 65:
                default_reasons.append(
                    "Low Attendance"
                )

            if selected_row["Assignment_Score"] < 60:
                default_reasons.append(
                    "Low Assignment Scores"
                )

            if selected_row["Internal_Marks"] < 60:
                default_reasons.append(
                    "Low Internal Marks"
                )

            if selected_row["Previous_GPA"] < 6.5:
                default_reasons.append(
                    "Low GPA"
                )

            if selected_row["Study_Hours"] < 2:
                default_reasons.append(
                    "Low Study Hours"
                )

            reasons = st.multiselect(
                "Reason",
                [
                    "Low Attendance",
                    "Low Assignment Scores",
                    "Low Internal Marks",
                    "Low GPA",
                    "Low Study Hours",
                    "Performance Decline",
                ],
                default=default_reasons,
            )

            action = st.selectbox(
                "Action",
                [
                    "Counseling",
                    "Parent Contact",
                    "Extra Classes",
                    "Study Plan",
                    "Attendance Monitoring",
                    "Mentoring",
                    "Targeted Feedback",
                ],
            )

            notes = st.text_area(
                "Teacher notes"
            )

            follow_up = st.date_input(
                "Follow-up date",
                value=date.today(),
            )

            intervention_status = st.selectbox(
                "Status",
                [
                    "Follow-up Required",
                    "In Progress",
                    "Resolved",
                ],
            )

            submitted = st.form_submit_button(
                "💾 Save Intervention",
                type="primary",
                use_container_width=True,
            )

            if submitted:

                if not reasons:

                    st.error(
                        "Select at least one reason."
                    )

                else:

                    try:

                        add_intervention(
                            selected_sid,
                            ", ".join(reasons),
                            action,
                            notes,
                            follow_up.isoformat(),
                            intervention_status,
                            st.session_state.username,
                        )

                        st.success(
                            "Intervention saved successfully."
                        )

                        st.rerun()

                    except Exception as exc:

                        st.error(
                            f"Could not save intervention: {exc}"
                        )

        all_interventions = pd.DataFrame(
            list_interventions(
                selected_sid
            )
        )

        if not all_interventions.empty:

            intervention_columns = [
                "reason",
                "action",
                "notes",
                "follow_up_date",
                "status",
                "teacher_username",
                "created_at",
            ]

            available_columns = [
                column
                for column in intervention_columns
                if column in all_interventions.columns
            ]

            st.dataframe(
                all_interventions[
                    available_columns
                ],
                use_container_width=True,
                hide_index=True,
            )

    # =====================================================
    # ACCOUNT MANAGEMENT
    # =====================================================

    with tabs[6]:

        st.subheader(
            "👥 Account Management"
        )

        st.caption(
            "Teachers can create student login "
            "accounts linked to academic records."
        )

        with st.form(
            "teacher_create_student"
        ):

            account_sid = st.selectbox(
                "Student ID",
                clustered["Student_ID"].tolist(),
            )

            full_name = st.text_input(
                "Student full name"
            )

            account_username = st.text_input(
                "Username"
            )

            account_password = st.text_input(
                "Temporary password",
                type="password",
            )

            account_email = st.text_input(
                "Email (optional)"
            )

            create_account = st.form_submit_button(
                "➕ Create Student Account",
                type="primary",
            )

            if create_account:

                if len(account_password) < 6:

                    st.error(
                        "Password must contain at least "
                        "6 characters."
                    )

                elif not full_name.strip():

                    st.error(
                        "Student full name is required."
                    )

                else:

                    ok, message = create_student_account_compat(
                        account_username,
                        account_password,
                        account_sid,
                        full_name,
                        account_email,
                        st.session_state.username,
                    )

                    if ok:
                        st.success(message)
                    else:
                        st.error(message)

        users = pd.DataFrame(
            list_users("Student")
        )

        if not users.empty:

            display_columns = [
                "username",
                "full_name",
                "student_id",
                "email",
                "active",
                "created_by",
                "created_at",
            ]

            available_columns = [
                column
                for column in display_columns
                if column in users.columns
            ]

            display = users[
                available_columns
            ].copy()

            if "active" in display.columns:

                display["active"] = display[
                    "active"
                ].map(
                    {
                        1: "Active",
                        0: "Disabled",
                        True: "Active",
                        False: "Disabled",
                    }
                )

            st.dataframe(
                display,
                use_container_width=True,
                hide_index=True,
            )

        else:

            st.info(
                "No student accounts found."
            )

    # =====================================================
    # REPORTS
    # =====================================================

    with tabs[7]:

        st.subheader(
            "📥 Professional Excel Reports"
        )

        summary = pd.DataFrame(
            {
                "Metric": [
                    "Total Students",
                    "On Track",
                    "Watch",
                    "High Risk",
                    "Average GPA",
                    "Average Attendance",
                    "Average Academic Index",
                ],
                "Value": [
                    len(clustered),
                    len(on_track),
                    len(watch),
                    len(high_risk),
                    clustered[
                        "Previous_GPA"
                    ].mean(),
                    clustered[
                        "Attendance"
                    ].mean(),
                    clustered[
                        "Academic_Index"
                    ].mean(),
                ],
            }
        )

        all_interventions = pd.DataFrame(
            list_interventions()
        )

        report_sheets = {
            "Executive Summary": summary,
            "All Students": clustered,
            "High Risk": high_risk,
            "Watch List": watch,
            "Top Performers": (
                clustered
                .sort_values(
                    "Academic_Index",
                    ascending=False,
                )
                .head(20)
            ),
            "Cluster Analysis": (
                clustered
                .groupby(
                    [
                        "Cluster",
                        "Performance_Group",
                    ],
                    as_index=False,
                )[
                    [
                        "Academic_Index",
                        "Attendance",
                        "Previous_GPA",
                    ]
                ]
                .mean()
            ),
        }

        if not all_interventions.empty:

            report_sheets[
                "Interventions"
            ] = all_interventions

        workbook = dataframe_to_xlsx(
            report_sheets
        )

        st.download_button(
            "📊 Download Complete EduPulse Excel Workbook",
            workbook,
            "EduPulse_Academic_Intelligence_Report.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.markdown(
            """
            **Workbook includes:**

            Executive Summary • All Students •
            High Risk • Watch List • Top Performers •
            Cluster Analysis • Intervention History
            """
        )

    # =====================================================
    # DATA MANAGEMENT
    # =====================================================

    with tabs[8]:

        st.subheader(
            "🗂️ Academic Data Management"
        )

        st.warning(
            "Upload a CSV only when you have an "
            "updated academic dataset. Required "
            "columns must match the project schema."
        )

        uploaded = st.file_uploader(
            "Upload updated student_performance.csv",
            type=["csv"],
        )

        if uploaded:

            try:

                new_df = pd.read_csv(
                    uploaded
                )

                required_columns = [
                    "Student_ID"
                ] + FEATURES

                missing_columns = [
                    column
                    for column in required_columns
                    if column not in new_df.columns
                ]

                if missing_columns:

                    st.error(
                        "Missing required columns: "
                        + ", ".join(
                            missing_columns
                        )
                    )

                else:

                    new_df["Student_ID"] = (
                        new_df["Student_ID"]
                        .astype(str)
                        .str.strip()
                    )

                    new_df[FEATURES] = (
                        new_df[FEATURES]
                        .apply(
                            pd.to_numeric,
                            errors="coerce",
                        )
                    )

                    invalid_numeric = new_df[
                        FEATURES
                    ].isna().sum()

                    invalid_total = int(
                        invalid_numeric.sum()
                    )

                    if invalid_total > 0:

                        st.warning(
                            f"The uploaded dataset contains "
                            f"{invalid_total} missing/invalid "
                            "numeric values. They will be "
                            "handled during dataset loading."
                        )

                    st.success(
                        f"Validated dataset: "
                        f"{len(new_df)} students, "
                        f"{len(new_df.columns)} columns."
                    )

                    st.dataframe(
                        new_df.head(10),
                        use_container_width=True,
                        hide_index=True,
                    )

                    if st.button(
                        "💾 Replace Academic Dataset",
                        type="primary",
                    ):

                        backup_path = (
                            DATA_PATH
                            + ".backup"
                        )

                        if os.path.exists(
                            DATA_PATH
                        ):

                            shutil.copy2(
                                DATA_PATH,
                                backup_path,
                            )

                        new_df.to_csv(
                            DATA_PATH,
                            index=False,
                        )

                        st.success(
                            "Dataset updated successfully. "
                            "Dashboard will reload with "
                            "the new records."
                        )

                        refresh_dataset()

            except Exception as exc:

                st.error(
                    "Could not read the uploaded CSV: "
                    f"{exc}"
                )

        st.divider()

        st.write(
            "**Current dataset:**",
            len(df),
            "students",
        )

        if os.path.exists(DATA_PATH):

            modified_time = datetime.fromtimestamp(
                os.path.getmtime(DATA_PATH)
            )

            st.write(
                "**Last local file modification:**",
                modified_time.strftime(
                    "%d %b %Y, %I:%M %p"
                ),
            )

        # =================================================
        # SNAPSHOT
        # =================================================

        if st.button(
            "📸 Capture Current Academic Snapshot",
            use_container_width=True,
        ):

            try:

                snapshot = clustered.copy()

                captured = save_snapshot(
                    snapshot,
                    st.session_state.username,
                )

                st.success(
                    f"Snapshot captured at {captured}. "
                    "Students can now see trend history."
                )

            except Exception as exc:

                st.error(
                    f"Could not capture snapshot: {exc}"
                )
