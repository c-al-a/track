import streamlit as st
import pandas as pd
import numpy as np
from datetime import timedelta
from openpyxl import load_workbook

FILE_PATH = "tracker.xlsx"

# =========================
# LOAD NAME MANAGER LISTS
# =========================
def load_named_range(name):
    wb = load_workbook(FILE_PATH, data_only=True)
    dn = wb.defined_names[name]

    values = []
    for sheet, coord in dn.destinations:
        ws = wb[sheet]
        for row in ws[coord]:
            for cell in row:
                if cell.value is not None:
                    values.append(cell.value)
    return values

@st.cache_data
def load_lists():
    names = load_named_range("NameList")
    codes = load_named_range("CodeList")
    teams = load_named_range("TeamClientList")
    holidays = set(pd.to_datetime(load_named_range("Holidays")).date)
    return names, codes, teams, holidays

# =========================
# LOAD SHEETS
# =========================
@st.cache_data
def load_data():
    movement = pd.read_excel(FILE_PATH, sheet_name="MOVEMENT LOG")
    lifecycle = pd.read_excel(FILE_PATH, sheet_name="NEW HIRE_PROMOTION_RESIGNATION")
    resources = pd.read_excel(FILE_PATH, sheet_name="RESOURCES")
    return movement, lifecycle, resources

# =========================
# SESSION HANDLING
# =========================
def normalize_sessions(df):
    session_map = {"AM": 0, "PM": 1}

    df["Start Session"] = (
        df["Start Session"].astype(str).str.strip().str.upper().map(session_map)
    )
    df["End Session"] = (
        df["End Session"].astype(str).str.strip().str.upper().map(session_map)
    )

    df["Start Session"] = df["Start Session"].fillna(0).astype(int)
    df["End Session"] = df["End Session"].fillna(1).astype(int)

    df["Start Date"] = pd.to_datetime(df["Start Date"])
    df["End Date"] = pd.to_datetime(df["End Date"])

    df["start_dt"] = df["Start Date"] + pd.to_timedelta(df["Start Session"] * 12, unit="h")
    df["end_dt"] = df["End Date"] + pd.to_timedelta(df["End Session"] * 12, unit="h")

    return df

# =========================
# BUILD TIMELINE
# =========================
def build_timeline(start_date, end_date):
    days = pd.date_range(start_date, end_date, freq="D")
    timeline = []
    for d in days:
        timeline.append(d)
        timeline.append(d + pd.Timedelta(hours=12))
    return timeline

# =========================
# CURRENT DESIGNATION
# =========================
def get_current_designation(name, lifecycle, ref_date):
    df = lifecycle[lifecycle["Name"] == name].copy()
    df["Effective Date"] = pd.to_datetime(df["Effective Date"])

    df = df[df["Effective Date"] <= ref_date]

    if df.empty:
        return None

    df = df.sort_values("Effective Date")
    return df.iloc[-1]["Designation"]

# =========================
# CURRENT CORE TEAM
# =========================
def get_core_team(name, movement, ref_dt):
    df = movement[movement["Name"] == name].copy()
    df = normalize_sessions(df)

    df = df[df["start_dt"] <= ref_dt]

    if df.empty:
        return None

    df = df.sort_values("start_dt")
    return df.iloc[-1]["Destination Team"]

# =========================
# TIMELINE PER EMPLOYEE
# =========================
def build_employee_timeline(name, movement, timeline):
    df = movement[movement["Name"] == name].copy()
    df = normalize_sessions(df)

    result = []

    for t in timeline:
        team = None

        for _, row in df.iterrows():
            if row["start_dt"] <= t <= row["end_dt"]:
                team = row["Destination Team"]

        result.append(team)

    return result

# =========================
# COLOR MAP
# =========================
def generate_colors(values):
    palette = [
        "#1f77b4","#ff7f0e","#2ca02c","#d62728",
        "#9467bd","#8c564b","#e377c2","#7f7f7f"
    ]
    unique = list(set(values))
    return {v: palette[i % len(palette)] for i, v in enumerate(unique)}

# =========================
# RENDER GRID
# =========================
def render_grid(df, timeline, holidays):

    styled = pd.DataFrame(index=df.index, columns=timeline)

    for r in df.index:
        for t in timeline:
            day = t.date()

            if day.weekday() >= 5 or day in holidays:
                styled.loc[r, t] = "background-color: black"
            else:
                val = df.loc[r, t]
                color = df.attrs["colors"].get(val, "#ffffff")
                styled.loc[r, t] = f"background-color: {color}"

    st.dataframe(df.style.apply(lambda _: styled, axis=None), use_container_width=True)

# =========================
# MAIN
# =========================
movement, lifecycle, resources = load_data()
names, codes, teams, holidays = load_lists()

st.title("Workforce Timeline")

mode = st.selectbox("Mode", ["A", "B", "C"])

start_date = st.date_input("Start Date")
end_date = st.date_input("End Date")

timeline = build_timeline(start_date, end_date)

# =========================
# A: EMPLOYEE VIEW
# =========================
if mode == "A":
    name_input = st.text_input("Name (partial)")

    matched = [n for n in names if name_input.lower() in n.lower()]

    data = {}
    all_vals = []

    for emp in matched:
        vals = build_employee_timeline(emp, movement, timeline)
        data[emp] = vals
        all_vals.extend([v for v in vals if v])

    if data:
        df = pd.DataFrame(data, index=timeline).T
        df.attrs["colors"] = generate_colors(all_vals)

        render_grid(df, timeline, holidays)

        info = []
        for emp in matched:
            designation = get_current_designation(emp, lifecycle, end_date)
            core = get_core_team(emp, movement, timeline[-1])
            role = resources.loc[resources["Name"] == emp, "CURRENT ROLE"]

            role = role.iloc[0] if len(role) else None

            info.append([emp, designation, role, core])

        st.write(pd.DataFrame(info, columns=["Name","Designation","Role","Core Team"]))

# =========================
# B: CORE TEAM VIEW
# =========================
elif mode == "B":
    team = st.text_input("Team (exact)")

    ref_dt = timeline[-1]

    members = []
    for n in names:
        if get_core_team(n, movement, ref_dt) == team:
            members.append(n)

    def build(group):
        data = {}
        vals = []

        for emp in group:
            row = build_employee_timeline(emp, movement, timeline)
            data[emp] = row
            vals.extend([v for v in row if v])

        df = pd.DataFrame(data, index=timeline).T
        df.attrs["colors"] = generate_colors(vals)
        return df

    pl = [n for n in members if resources.loc[resources["Name"] == n, "CURRENT ROLE"].iloc[0] == "PL"]
    pa = [n for n in members if resources.loc[resources["Name"] == n, "CURRENT ROLE"].iloc[0] == "PA"]

    if pl:
        st.subheader("PL")
        render_grid(build(pl), timeline, holidays)

    if pa:
        st.subheader("PA")
        render_grid(build(pa), timeline, holidays)

# =========================
# C: PARTICIPATION VIEW
# =========================
elif mode == "C":
    team = st.text_input("Team (exact)")

    def build(group):
        data = {}

        for emp in group:
            row = build_employee_timeline(emp, movement, timeline)
            data[emp] = [1 if v == team else 0 for v in row]

        return pd.DataFrame(data, index=timeline).T

    involved = movement[movement["Destination Team"] == team]["Name"].unique()

    pl = [n for n in involved if resources.loc[resources["Name"] == n, "CURRENT ROLE"].iloc[0] == "PL"]
    pa = [n for n in involved if resources.loc[resources["Name"] == n, "CURRENT ROLE"].iloc[0] == "PA"]

    if pl:
        st.subheader("PL Participation")
        st.line_chart(build(pl).T)

    if pa:
        st.subheader("PA Participation")
        st.line_chart(build(pa).T)