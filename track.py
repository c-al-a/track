import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import io

def add_download_buttons(df, prefix="timeline"):
    """Adds CSV and Excel download buttons for a given DataFrame"""
    
    # CSV
    csv_buffer = io.StringIO()
    df.to_csv(csv_buffer, index=True)
    st.download_button(
        label=f"Download {prefix} as CSV",
        data=csv_buffer.getvalue().encode(),
        file_name=f"{prefix}.csv",
        mime="text/csv"
    )

    # Excel
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=True, sheet_name="Timeline")
        writer.save()
    st.download_button(
        label=f"Download {prefix} as Excel",
        data=excel_buffer.getvalue(),
        file_name=f"{prefix}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

FILE_PATH = "Resource Tracker_New (3).xlsm"

st.markdown(
    """
    <style>
    /* Reduce horizontal margin between sidebar and main page */
    /* Adjust the main container that sits next to the sidebar */
    .block-container {
        margin-left: 0.5rem;   /* default is bigger, e.g., 3rem */
        margin-right: 0.5rem;  /* optional: reduce right margin too */
    }
    </style>
    """,
    unsafe_allow_html=True
)

@st.cache_data
def load_data():
    movement = pd.read_excel(FILE_PATH, sheet_name="MOVEMENT LOG")
    lifecycle = pd.read_excel(FILE_PATH, sheet_name="NEW HIRE_PROMOTION_RESIGNATION")
    resources = pd.read_excel(FILE_PATH, sheet_name="RESOURCES")
    leaves = pd.read_excel(FILE_PATH, sheet_name="LEAVES")

    return movement, lifecycle, resources, leaves


def load_named_range(name):
    wb = load_workbook(FILE_PATH, data_only=True)
    rng = wb.defined_names[name]

    values = []
    for sheet, coord in rng.destinations:
        ws = wb[sheet]
        for row in ws[coord]:
            if isinstance(row, tuple):
                values.extend([c.value for c in row])
            else:
                values.append(row.value)

    return values


@st.cache_data
def load_lists():
    names = load_named_range("NameList")
    holidays = set(pd.to_datetime(load_named_range("Holidays")).date)
    team_clients_raw = load_named_range("TeamClientList")
    team_clients = sorted([
        t for t in team_clients_raw
        if t is not None and str(t).strip() != ""
        ])
    return names, holidays, team_clients


# =========================
# PREPROCESS LEAVES INTO DICTIONARY
# =========================
@st.cache_data
def build_leave_dict(leaves_df):
    leaves_df = prepare_leaves(leaves_df)
    
    leave_dict = {}
    
    for _, row in leaves_df.iterrows():
        name = row["NAME"]
        date = row["DATE"].date()
        schedule = str(row["SCHEDULE"]).upper()  # "WHOLE DAY", "AM", "PM"

        # Initialize key if not exists
        if (name, date) not in leave_dict:
            leave_dict[(name, date)] = set()
        
        if schedule == "WHOLE DAY":
            leave_dict[(name, date)] = {"AM", "PM"}  # mark both sessions
        else:
            leave_dict[(name, date)].add(schedule)

    return leave_dict


# =========================
# CLEAN VALUES
# =========================
def clean_team(val):
    if val is None or pd.isna(val) or str(val).strip() == "":
        return "-"
    return str(val)


# =========================
# PREPARE MOVEMENT
# =========================
def prepare_movement(df):
    df = df.copy()

    df["START DATE"] = pd.to_datetime(df["START DATE"])
    df["END DATE"] = pd.to_datetime(df["END DATE"])

    # normalize sessions
    df["START SESSION"] = df["START SESSION"].fillna("AM")
    df["END SESSION"] = df["END SESSION"].fillna("PM")

    df["END DATE"] = df["END DATE"].fillna(pd.Timestamp("2100-01-01"))

    df["DESTINATION TEAM/CLIENT"] = (
        df["DESTINATION TEAM/CLIENT"].apply(clean_team)
    )

    return df

# =========================
# PREPARE LEAVES
# =========================
def prepare_leaves(df):
    df = df.copy()

    df["DATE"] = pd.to_datetime(df["DATE"])

    return df


# =========================
# SPLIT CORE / TEMP
# =========================
def split_core_temp(df):
    core = df[df["ASSIGNMENT TYPE"] == "PERMANENT"].copy()
    temp = df[df["ASSIGNMENT TYPE"] == "TEMPORARY"].copy()
    return core, temp
    
def get_all_core_teams(movement):
    movement = prepare_movement(movement)
    core_df, _ = split_core_temp(movement)

    return sorted(core_df["DESTINATION TEAM/CLIENT"].dropna().unique())

# print(get_all_core_teams)


# =========================
# BUILD TEMP INTERVALS
# =========================
def build_temp_intervals(temp_rows):

    temp_rows = temp_rows.sort_values(
        ["START DATE", "START SESSION"]
    )

    intervals = []

    for _, row in temp_rows.iterrows():

        intervals.append({
            "start_date": row["START DATE"],
            "start_session": row["START SESSION"],
            "end_date": row["END DATE"],
            "end_session": row["END SESSION"],
            "team": clean_team(row["DESTINATION TEAM/CLIENT"])
        })

    return intervals

SESSION_ORDER = {
    "AM": 0,
    "PM": 1
}


def session_key(date, session):
    return (pd.Timestamp(date), SESSION_ORDER[session])


def is_session_in_range(
    current_date,
    current_session,
    start_date,
    start_session,
    end_date,
    end_session
):

    current = session_key(current_date, current_session)
    start = session_key(start_date, start_session)
    end = session_key(end_date, end_session)

    return start <= current <= end


# CHECK LEAVE

# =========================
# CHECK LEAVE (FAST)
# =========================
def is_on_leave(name, date, session, leave_dict):
    """
    name: employee name
    date: datetime.date or pd.Timestamp
    session: "AM" or "PM"
    leave_dict: dict from build_leave_dict
    """
    sessions = leave_dict.get((name, date.date()), set())
    return session in sessions


# =========================
# RESOLVE EMPLOYEE TIMELINE
# =========================
def resolve_employee(name, movement, timeline, holidays, leave_dict):
    movement = prepare_movement(movement)
    core_df, temp_df = split_core_temp(movement)

    core_rows = core_df[core_df["NAME"] == name]
    temp_rows = temp_df[temp_df["NAME"] == name]

    if core_rows.empty:
        return ["-"] * len(timeline)

    temp_intervals = build_temp_intervals(temp_rows)
    result = []

    for t_date, t_session in timeline:
        active = None

        # Check TEMP intervals first
        for interval in temp_intervals:
            if is_session_in_range(
                t_date,
                t_session,
                interval["start_date"],
                interval["start_session"],
                interval["end_date"],
                interval["end_session"]
            ):
                active = interval["team"]
                break

        # LEAVE
        if is_on_leave(name, t_date, t_session, leave_dict):
            result.append("LEAVE")

        # WEEKENDS / HOLIDAYS
        elif t_date.weekday() >= 5 or t_date.date() in holidays:
            result.append("NON-WORK DAY")

        # TEMP ASSIGNMENT
        elif active:
            result.append(active)

        # CORE TEAM (only if active for this date)
        else:
            # filter core rows active on this timeline date
            active_core = core_rows[
                (core_rows["START DATE"] <= t_date) &
                (core_rows["END DATE"] >= t_date)
            ]

            if not active_core.empty:
                core_team = clean_team(active_core.iloc[-1]["DESTINATION TEAM/CLIENT"])
            else:
                core_team = "-"

            result.append(core_team)

    return result


# =========================
# CURRENT CORE TEAM
# =========================
def get_current_core_team(name, movement):
    """
    Returns the core team for an employee as of today.
    Only considers assignments active today.
    """
    movement = prepare_movement(movement)
    core_df, _ = split_core_temp(movement)
    core_rows = core_df[core_df["NAME"] == name]

    if core_rows.empty:
        return "-"

    today = pd.Timestamp.today().normalize()

    # Only consider rows active today
    active_rows = core_rows[(core_rows["START DATE"] <= today) & (core_rows["END DATE"] >= today)]

    if not active_rows.empty:
        return clean_team(active_rows.iloc[-1]["DESTINATION TEAM/CLIENT"])
    else:
        return "-"


# =========================
# TIMELINE
# =========================
def build_timeline(start_date, end_date):

    dates = pd.date_range(start_date, end_date, freq="D")

    timeline = []

    for d in dates:
        timeline.append((d, "AM"))
        timeline.append((d, "PM"))

    return timeline


# =========================
# COLORS
# =========================
def generate_color_map(values):

    color_map = {
        "NON-WORK DAY": "#444444",
        "LEAVE": "#555555",
        "-": "#ffffff"
    }

    palette = [
        "#dbeafe",
        "#ffedd5",
        "#dcfce7",
        "#fee2e2",
        "#ede9fe",
        "#f3e8e2",
        "#fce7f3",
        "#e5e7eb",
        "#fef9c3",
        "#cffafe",
        "#fde68a",
        "#ddd6fe",
        "#bfdbfe",
        "#fecaca",
        "#bbf7d0"
    ]

    unique = sorted(set(values))

    team_values = [
        v for v in unique
        if v not in color_map
    ]

    for i, v in enumerate(team_values):
        color_map[v] = palette[i % len(palette)]

    return color_map


# =========================
# HTML TABLE
# =========================
def render_colored_table(df, color_map):

    html = """
    <style>
        table {
            border-collapse: separate;
            font-size: 12px;
        }

        td, th {
            border-left: 1px solid #ddd;
            border-right: 1px solid #ddd;
            border-top: 2px solid #999;
            border-bottom: 2px solid #999;
            padding: 4px;
            text-align: center;
            
            max-width: 137px;
            white-space: normal;
            word-wrap: break-word;
        }

        td {
            box-shadow:
                inset 0 2px 0 black
        }
        
        th {
            position: sticky;
            top: 0;
            background: white;
            z-index: 100;
        }
        
        table {
            border-collapse: collapse;
            font-size: 12px;
            width: max-content;
        }
        
        th {
            background-clip: padding-box;
            
            border-right: 3px solid black !important;
        }

        td.employee-col {
            position: sticky;
            left: 0;
            background: white;
            z-index: 10;
            text-align: left;
            
            

            box-shadow:
                inset 2px 0 0 black,
                inset 0 2px 0 black;
                inset 0 -2px 0 black;
        }
        
        th.employee-col {
            position: sticky;
            top: 0;
            left: 0;
            z-index: 200;
            background: white;
            

            box-shadow:
                inset 3px 0 0 black,
                inset -3px 0 0 black,
                inset 0 -3px 0 black;
        }

        .date-divider {
            border-right: 3px solid black !important;
        }
    </style>
    <table>
    """

    # =========================
    # HEADER ROW 1 (DATES)
    # =========================

    html += "<tr>"
    html += "<th rowspan='2' class='employee-col'>Name</th>"

    columns = list(df.columns)

    for i in range(0, len(columns), 2):

        date_val, _ = columns[i]

        html += (
            f"<th colspan='2' "
            f"style='text-align:center; vertical-align:middle;'>"
            f"{date_val.strftime('%a')}<br>"
            f"{date_val.month}/{date_val.day}"
            f"</th>"
            )

    html += "</tr>"

    # =========================
    # HEADER ROW 2 (AM/PM)
    # =========================

    html += "<tr>"

    for i, (date_val, session) in enumerate(columns):

        extra_class = ""
        align = "right" if session == "AM" else "left"

        if i % 2 == 0:
            extra_class = ""
        else:
            extra_class = "date-divider"

        html += (
            f"<th class='{extra_class}' "
            f"style='text-align:{align};'>"
            f"{session}"
            f"</th>"
        )

    html += "</tr>"

    # =========================
    # BODY
    # =========================

    for i, (idx, row) in enumerate(df.iterrows(), start=1):
        html += (
            f"<tr>"
            f"<td class='employee-col'><b style='font-size:10px'>[{i}]</b> {idx}</td>"
        )

        for i, val in enumerate(row):

            val = clean_team(val)

            color = color_map.get(val, "#ffffff")

            text_color = (
                "white"
                if val in ["LEAVE", "NON-WORK DAY"]
                else "black"
            )

            extra_class = ""

            if i % 2 == 1:
                extra_class = "date-divider"
                
            date_val, session = columns[i]
            
            tooltip = f"{date_val.strftime('%a')} {date_val.month}/{date_val.day}"

            html += (
                f"<td class='{extra_class}' "
                f"style='background-color:{color};color:{text_color}' "
                f"title='{tooltip}'>"
                f"{val}"
                f"</td>"
            )

        html += "</tr>"

    html += "</table>"

    st.markdown(html, unsafe_allow_html=True)


# =========================
# LEGEND
# =========================
def render_legend(color_map):
    st.sidebar.subheader("Legend")
    
    for key, color in color_map.items():
        st.sidebar.markdown(
            f"""
            <div style="
                display:flex;
                align-items:center;
                margin-bottom:2px;  /* reduce spacing */
            ">
                <div style="
                    width:10px;
                    height:14px;
                    background:{color};
                    margin-right:8px;
                "></div>
                {key}
            </div>
            """,
            unsafe_allow_html=True
        )


# =========================
# MAIN APP
# =========================
movement, lifecycle, resources, leaves = load_data()
names, holidays, team_clients = load_lists()

st.title("Prod Team Timeline")

mode = st.selectbox("Search by:", ["Employee", "Core Team/Client", "Assignment Team/Client"])

start_date = st.date_input(
    "START DATE",
    key="main_start_date"
    )

end_date = st.date_input(
    "END DATE",
    key="main_end_date"
    )

leave_dict = build_leave_dict(leaves)
timeline = build_timeline(start_date, end_date)

# =========================
# MODE: EMPLOYEE VIEW
# =========================
if mode == "Employee":
    name_input = st.text_input("Employee Search")

    matched = [
        n for n in names
        if isinstance(n, str) and name_input.lower() in n.lower()
    ]

    data = {}

    for emp in matched:
        data[emp] = resolve_employee(emp, movement, timeline, holidays, leave_dict)

    if data:
        df = pd.DataFrame(data, index=timeline).T
        df = df.sort_index()

        all_values = []
        for row in data.values():
            all_values.extend([v for v in row])

        color_map = generate_color_map(all_values)

        st.subheader("Employee Timeline")
        render_colored_table(df, color_map)
        render_legend(color_map)

# =========================
# MODE: CORE TEAM/CLIENT
# =========================
elif mode == "Core Team/Client":
    team_options = get_all_core_teams(movement)
    team_input = st.selectbox("Team/Client Name", team_options)

    if team_input:

        team_input = team_input.strip()

        # =========================
        # FIND EMPLOYEES BY CORE TEAM
        # =========================
        matched = []

        for emp in names:
            if not isinstance(emp, str):
                continue

            core_team = get_current_core_team(emp, movement)

            if core_team == team_input:
                matched.append(emp)

        # st.subheader(f"Team Timeline: {team_input}")

        if not matched:
            st.warning("No employees found for this team/client.")
        else:

            # =========================
            # SPLIT BY ROLE (PL / PA)
            # =========================
            pl_list = []
            pa_list = []

            for emp in matched:
                role = resources.loc[resources["NAME"] == emp, "CURRENT PROD ROLE"]
                role = role.iloc[0] if not role.empty else None

                if role == "PL":
                    pl_list.append(emp)
                elif role == "PA":
                    pa_list.append(emp)

            # =========================
            # FUNCTION TO BUILD TABLE
            # =========================
            def build_table(group):
                data = {}

                for emp in group:
                    data[emp] = resolve_employee(emp, movement, timeline, holidays, leave_dict)

                if not data:
                    return None, None

                df = pd.DataFrame(data, index=timeline).T
                df = df.sort_index()

                all_values = []
                for row in data.values():
                    all_values.extend([v for v in row])

                color_map = generate_color_map(all_values)

                return df, color_map

            # =========================
            # PL TABLE
            # =========================
            if pl_list:
                st.subheader("[" + team_input + "]" + " Core " + "PL")

                df_pl, color_map_pl = build_table(pl_list)

                if df_pl is not None:
                    render_colored_table(df_pl, color_map_pl)
                    render_legend(color_map_pl)
                    
                    # add_download_buttons(df_pl, prefix=f"{team_input}_PL")

            # =========================
            # PA TABLE
            # =========================
            if pa_list:
                st.subheader("[" + team_input + "]" + " Core " + "PA")

                df_pa, color_map_pa = build_table(pa_list)

                if df_pa is not None:
                    render_colored_table(df_pa, color_map_pa)
                    render_legend(color_map_pa)
                    
                    # add_download_buttons(df_pa, prefix=f"{team_input}_PA")

            # =========================
            # SUMMARY LIST (OPTIONAL)
            # =========================
            st.subheader("Employees in Team")

            st.dataframe(pd.DataFrame(matched, columns=["EMPLOYEES"]))

elif mode == "Assignment Team/Client":

    # =========================
    # INPUT
    # =========================
    team_options = sorted(
        prepare_movement(movement)["DESTINATION TEAM/CLIENT"].dropna().unique()
        )
    team_input2 = st.selectbox("Team/Client Name", team_options)

    # start_date = st.date_input("START DATE")
    # end_date = st.date_input("END DATE")

    window_start = pd.to_datetime(start_date)
    window_end = pd.to_datetime(end_date)

    # =========================
    # FIND EMPLOYEES WITH ANY OVERLAP
    # =========================
    matched = []

    for emp in names:
        if not isinstance(emp, str):
            continue

        emp_movement = movement[movement["NAME"] == emp]
        if emp_movement.empty:
            continue

        emp_movement = prepare_movement(emp_movement)

        team_rows = emp_movement[
            emp_movement["DESTINATION TEAM/CLIENT"] == team_input2
        ]

        if team_rows.empty:
            continue

        # check overlap
        for _, row in team_rows.iterrows():
            start = row["START DATE"]
            end = row["END DATE"]

            if start <= window_end and end >= window_start:
                matched.append(emp)
                break

    matched = sorted(
        set(matched),
        key=lambda emp: (
            str(get_current_core_team(emp, movement)),  # group by core team/client
            emp                                  # then sort employee name
            )
        )

    # =========================
    # SPLIT BY ROLE
    # =========================
    pl_list = []
    pa_list = []

    for emp in matched:
        role = resources.loc[resources["NAME"] == emp, "CURRENT PROD ROLE"]
        role = role.iloc[0] if not role.empty else None

        if role == "PL":
            pl_list.append(emp)
        elif role == "PA":
            pa_list.append(emp)

    
    def build_table(group):

        employee_rows = []

        for emp in group:

            core_team = get_current_core_team(emp, movement)

            employee_rows.append({
                "emp": emp,
                "core_team": core_team
            })

        # SORT
        employee_rows = sorted(
            employee_rows,
            key=lambda x: (
                str(x["core_team"]).lower(),
                x["emp"].lower()
            )
        )

        data = {}

        for row in employee_rows:

            emp = row["emp"]
            core_team = row["core_team"]

            display_name = f"{core_team} | {emp}"

            data[display_name] = resolve_employee(
                emp,
                movement,
                timeline,
                holidays,
                leave_dict
            )

        if not data:
            return None, None

        df = pd.DataFrame(data, index=timeline).T

        all_values = []

        for row in data.values():
            all_values.extend(row)

        color_map = generate_color_map(all_values)

        return df, color_map

    if not matched:
        st.warning("No employees found for this team in the selected date range.")
    else:

        if pl_list:
            st.subheader("[" + team_input2 + "]" + " Core and Assigned " + "PL")
            df_pl, color_map_pl = build_table(pl_list)
            if df_pl is not None:
                render_colored_table(df_pl, color_map_pl)
                render_legend(color_map_pl)
                
                # add_download_buttons(df_pl, prefix=f"{team_input2}_PL")

        if pa_list:
            st.subheader("[" + team_input2 + "]" + " Core and Assigned " + "PA")
            df_pa, color_map_pa = build_table(pa_list)
            if df_pa is not None:
                render_colored_table(df_pa, color_map_pa)
                render_legend(color_map_pa)
                # add_download_buttons(df_pa, prefix=f"{team_input2}_PA")

        st.subheader("Employees in Team")
        st.dataframe(pd.DataFrame(sorted(matched), columns=["EMPLOYEES"]))
