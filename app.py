import calendar
from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

# -------------------- Configuration --------------------
st.set_page_config(
    page_title="Student Attendance Management",
    page_icon="🎓",
    layout="wide",
)

DATA_FILE = Path("attendance_data.xlsx")

STUDENT_COLUMNS = ["Roll Number", "Student Name", "Class", "Section"]
ATTENDANCE_COLUMNS = [
    "Date", "Roll Number", "Student Name", "Class", "Section", "Status"
]
STATUSES = ["Present", "Absent"]


# -------------------- Excel helpers --------------------
def load_sheet(sheet_name: str, columns: list[str]) -> pd.DataFrame:
    """Load a sheet from the Excel workbook, or return an empty table."""
    if not DATA_FILE.exists():
        return pd.DataFrame(columns=columns)

    try:
        df = pd.read_excel(DATA_FILE, sheet_name=sheet_name, dtype=str)
        for column in columns:
            if column not in df.columns:
                df[column] = ""
        return df[columns].fillna("")
    except (ValueError, FileNotFoundError, OSError):
        return pd.DataFrame(columns=columns)


def save_workbook(students_df: pd.DataFrame, attendance_df: pd.DataFrame) -> None:
    """Save both tables in one Excel workbook."""
    with pd.ExcelWriter(DATA_FILE, engine="openpyxl") as writer:
        students_df[STUDENT_COLUMNS].to_excel(
            writer, sheet_name="Students", index=False
        )
        attendance_df[ATTENDANCE_COLUMNS].to_excel(
            writer, sheet_name="Attendance", index=False
        )


def excel_bytes(df: pd.DataFrame, sheet_name: str = "Report") -> bytes:
    """Create an Excel download in memory."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return buffer.getvalue()


def clean_text(value) -> str:
    return str(value).strip()


# Load the saved data whenever the app runs.
students = load_sheet("Students", STUDENT_COLUMNS)
attendance = load_sheet("Attendance", ATTENDANCE_COLUMNS)

# Normalize saved dates and values.
if not attendance.empty:
    attendance["Date"] = attendance["Date"].astype(str).str[:10]
    attendance["Status"] = attendance["Status"].astype(str).str.title()
    attendance = attendance[attendance["Status"].isin(STATUSES)].copy()

# -------------------- Page header --------------------
st.title("🎓 Student Attendance Management System")
st.caption("Register students, record daily attendance, and create monthly Excel reports.")

if DATA_FILE.exists():
    st.caption(f"Excel workbook: `{DATA_FILE.resolve()}`")

tab_students, tab_attendance, tab_reports = st.tabs(
    ["👩‍🎓 Students", "✅ Mark Attendance", "📊 Monthly Reports"]
)

# -------------------- Students tab --------------------
with tab_students:
    st.subheader("Add a student")

    with st.form("add_student_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            roll_number = st.text_input("Roll Number", placeholder="e.g. 101")
            student_name = st.text_input("Student Name", placeholder="e.g. Ravi Kumar")
        with c2:
            class_name = st.text_input("Class", placeholder="e.g. CSE-2")
            section = st.text_input("Section", placeholder="e.g. A")

        add_student = st.form_submit_button("➕ Add Student", type="primary")

    if add_student:
        roll_number = clean_text(roll_number)
        student_name = clean_text(student_name)
        class_name = clean_text(class_name)
        section = clean_text(section)

        if not all([roll_number, student_name, class_name, section]):
            st.error("Please fill in all four fields.")
        elif students["Roll Number"].astype(str).str.strip().str.casefold().eq(
            roll_number.casefold()
        ).any():
            st.error("That roll number already exists. Please use a unique roll number.")
        else:
            new_student = pd.DataFrame([{
                "Roll Number": roll_number,
                "Student Name": student_name,
                "Class": class_name,
                "Section": section,
            }])
            students = pd.concat([students, new_student], ignore_index=True)
            save_workbook(students, attendance)
            st.success(f"{student_name} was added and saved to {DATA_FILE.name}.")
            st.rerun()

    st.divider()
    st.subheader("Registered students")
    if students.empty:
        st.info("No students yet. Use the form above to add your first student.")
    else:
        st.metric("Total students", len(students))
        st.dataframe(
            students.sort_values(["Class", "Roll Number"]),
            use_container_width=True,
            hide_index=True,
        )
        st.download_button(
            "⬇️ Download student list",
            data=excel_bytes(students, "Students"),
            file_name="students.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

# -------------------- Mark Attendance tab --------------------
with tab_attendance:
    st.subheader("Record daily attendance")

    if students.empty:
        st.info("First, add students in the Students tab.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            attendance_date = st.date_input("Attendance date", value=date.today())
        with col2:
            classes = sorted(students["Class"].dropna().astype(str).unique().tolist())
            chosen_class = st.selectbox("Choose class", classes)

        class_students = students[students["Class"] == chosen_class].copy()
        class_students = class_students.sort_values("Roll Number").reset_index(drop=True)

        previous_records = attendance[
            (attendance["Date"].astype(str) == attendance_date.isoformat())
            & (attendance["Class"].astype(str) == chosen_class)
        ]

        initial_statuses = []
        for _, student in class_students.iterrows():
            prior = previous_records[
                previous_records["Roll Number"].astype(str)
                == str(student["Roll Number"])
            ]
            initial_statuses.append(
                prior.iloc[0]["Status"] if not prior.empty else "Present"
            )

        editable = class_students[["Roll Number", "Student Name", "Section"]].copy()
        editable["Status"] = initial_statuses

        st.write("Choose Present or Absent for each student. New students default to Present; check each row before saving.")
        edited = st.data_editor(
            editable,
            use_container_width=True,
            hide_index=True,
            disabled=["Roll Number", "Student Name", "Section"],
            column_config={
                "Status": st.column_config.SelectboxColumn(
                    "Attendance Status",
                    options=STATUSES,
                    required=True,
                )
            },
            key=f"attendance_editor_{attendance_date.isoformat()}_{chosen_class}",
        )

        present_count = int((edited["Status"] == "Present").sum())
        absent_count = int((edited["Status"] == "Absent").sum())
        metric1, metric2, metric3 = st.columns(3)
        metric1.metric("Students", len(edited))
        metric2.metric("Present", present_count)
        metric3.metric("Absent", absent_count)

        if st.button("💾 Save Attendance", type="primary", key="save_attendance"):
            if not edited["Status"].isin(STATUSES).all():
                st.error("Please select Present or Absent for every student.")
            else:
                new_records = pd.DataFrame([
                    {
                        "Date": attendance_date.isoformat(),
                        "Roll Number": str(row["Roll Number"]),
                        "Student Name": str(row["Student Name"]),
                        "Class": chosen_class,
                        "Section": str(row["Section"]),
                        "Status": str(row["Status"]),
                    }
                    for _, row in edited.iterrows()
                ])

                # Re-saving a date/class updates its records instead of duplicating them.
                if not attendance.empty:
                    keep = ~(
                        (attendance["Date"].astype(str) == attendance_date.isoformat())
                        & (attendance["Class"].astype(str) == chosen_class)
                    )
                    attendance = attendance[keep].copy()

                attendance = pd.concat(
                    [attendance, new_records], ignore_index=True
                )
                save_workbook(students, attendance)
                st.success(
                    f"Attendance for {chosen_class} on {attendance_date.isoformat()} "
                    f"was saved to {DATA_FILE.name}."
                )
                st.rerun()

# -------------------- Monthly Reports tab --------------------
with tab_reports:
    st.subheader("Monthly attendance report")

    if attendance.empty:
        st.info("No attendance records found. Mark and save attendance first.")
    else:
        # Use recorded months plus the current month.
        recorded_months = set(
            attendance["Date"].astype(str).str.slice(0, 7).dropna().tolist()
        )
        recorded_months.add(date.today().strftime("%Y-%m"))
        month_options = sorted(recorded_months, reverse=True)

        c1, c2 = st.columns(2)
        with c1:
            chosen_month = st.selectbox("Choose month", month_options)
        with c2:
            report_classes = ["All classes"] + sorted(
                attendance["Class"].dropna().astype(str).unique().tolist()
            )
            report_class = st.selectbox("Choose class", report_classes)

        month_data = attendance[
            attendance["Date"].astype(str).str.startswith(chosen_month)
        ].copy()

        if report_class != "All classes":
            month_data = month_data[month_data["Class"] == report_class]

        if month_data.empty:
            st.warning("No attendance records for this month and class.")
        else:
            total_records = len(month_data)
            total_present = int((month_data["Status"] == "Present").sum())
            total_absent = int((month_data["Status"] == "Absent").sum())
            overall_percentage = (
                total_present / total_records * 100 if total_records else 0
            )

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Attendance records", total_records)
            m2.metric("Present", total_present)
            m3.metric("Absent", total_absent)
            m4.metric("Overall attendance", f"{overall_percentage:.1f}%")

            # One row per student. Percentage is based on days with saved records.
            summary = (
                month_data.assign(
                    Present_Count=(month_data["Status"] == "Present").astype(int),
                    Absent_Count=(month_data["Status"] == "Absent").astype(int),
                )
                .groupby(
                    ["Roll Number", "Student Name", "Class", "Section"],
                    dropna=False,
                )
                .agg(
                    Days_Recorded=("Date", "nunique"),
                    Present_Days=("Present_Count", "sum"),
                    Absent_Days=("Absent_Count", "sum"),
                )
                .reset_index()
            )
            summary["Attendance Percentage"] = (
                summary["Present_Days"] / summary["Days_Recorded"] * 100
            ).round(2)
            summary = summary.sort_values(["Class", "Roll Number"])

            st.markdown("#### Student-wise attendance percentage")
            st.caption(
                "Formula: Present days ÷ days with saved attendance records × 100. "
                "Only dates on which attendance was saved are counted."
            )
            st.dataframe(
                summary,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Attendance Percentage": st.column_config.NumberColumn(
                        format="%.2f%%"
                    )
                },
            )

            st.markdown("#### Daily attendance details")
            st.dataframe(
                month_data.sort_values(["Date", "Class", "Roll Number"]),
                use_container_width=True,
                hide_index=True,
            )

            st.download_button(
                "⬇️ Download monthly report (Excel)",
                data=excel_bytes(summary, "Monthly Summary"),
                file_name=f"attendance_summary_{chosen_month}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
            st.download_button(
                "⬇️ Download detailed attendance (Excel)",
                data=excel_bytes(
                    month_data.sort_values(["Date", "Class", "Roll Number"]),
                    "Attendance Details",
                ),
                file_name=f"attendance_details_{chosen_month}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

st.divider()
st.caption(
    "Reminder: Keep this project folder and attendance_data.xlsx safe. "
    "Back up the workbook regularly. Do not edit the workbook while the app is running."
)
