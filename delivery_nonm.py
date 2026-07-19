import pandas as pd

filename_delivery = input("Enter the Delivery Details file name: ")
delivery_lookup = pd.read_excel(filename_delivery)

filename_timesheet_3weeks = input("Enter the Timesheet (long) file name: ")
timesheet_3weeks = pd.read_excel(filename_timesheet_3weeks)

filename_timesheet_1week = input("Enter the Timesheet (delivery period) file name: ")
timesheet_1week = pd.read_excel(filename_timesheet_1week)

filename_resource_tracker = input("Enter the Resource Tracker file name: ")
resource_tracker = pd.read_excel(filename_resource_tracker, sheet_name= "RESOURCES")
resource_tracker = resource_tracker[[
    "EMPLOYEE CODE", "CORE TEAM/CLIENT"
]]

client_sbu_lookup = (
    delivery_lookup[["Client Name", "SBU Name"]].drop_duplicates()
)

delivery = timesheet_1week[
    timesheet_1week["Contract #"].notna()
    ].copy()

delivery = delivery[
    delivery["Type of work_Timesheet"].isin([
        "Sense Check", "SenseCheck"
    ])
]

delivery = delivery.merge(
    client_sbu_lookup, on = "Client Name", how = "left"
)

delivery = delivery[
    delivery["SBU Name"] != "NTrust - Manila"
]

delivery["Timesheet Date"] = pd.to_datetime(
    delivery["Timesheet Date"]
)

delivery = (
    delivery.sort_values("Timesheet Date")
    .groupby("Contract #", as_index = False).last()
)

delivery = delivery.rename(
    columns = {"Timesheet Date": "Delivered Date"}
)

delivery = delivery[[
    "Client Name", "Project Name", "Contract #", "Type of Work", "Delivered Date", "SBU Name"
]]

def role_summary(role, suffix):
    role_df = timesheet_3weeks[
        timesheet_3weeks["Role Name"].fillna("") == role
    ].copy()

    role_df["Timesheet Date"] = pd.to_datetime(role_df["Timesheet Date"])

    duration = (
        role_df.groupby("Contract #", as_index = False)
        .agg({"Duration in Minutes": "sum"})
        .rename(columns = {
            "Duration in Minutes": f"Duration in Minutes_{suffix}"
        })
    )

    resource = (
        role_df.sort_values("Timesheet Date").
        groupby("Contract #", as_index = False).last()
    )

    resource = resource[[
        "Contract #", "Resource Name", "Emp . ID"
    ]]

    resource = resource.rename(columns={
        "Resource Name": role,
        "Emp . ID": f"{suffix}_ID"
    })

    summary = resource.merge(
        duration,
        on = "Contract #",
        how = "outer"
    )

    summary[f"Duration in Hours_{suffix}"] = (
        summary[f"Duration in Minutes_{suffix}"]/60
    ).round(2)

    return summary

reviewer = role_summary("Reviewer", "Reviewer")
abstractor = role_summary("Abstractor", "Abstractor")
deliverytl = role_summary("Delivery TL", "TL")

summary = reviewer.merge(
    abstractor, on = "Contract #", how = "outer"
)

summary = summary.merge(
    deliverytl, on = "Contract #", how = "outer"
)

## delivery = delivery[
##    delivery["SBU Name"] == "NTrust - Manila"
##    ].copy()

result = delivery.merge(
    summary, on = "Contract #", how = "left"
)
########### left right inner or outer?

reviewer_cols = resource_tracker[[
    "EMPLOYEE CODE", "CORE TEAM/CLIENT"
]].rename(columns = {
    "CORE TEAM/CLIENT": "Reviewer_Team"
})

abstractor_cols = resource_tracker[[
    "EMPLOYEE CODE", "CORE TEAM/CLIENT"
]].rename(columns = {
    "CORE TEAM/CLIENT": "Abstractor_Team"
})

result = result.merge(
    reviewer_cols,
    left_on = "Reviewer_ID",
    right_on = "EMPLOYEE CODE",
    how = "left"
).drop(columns = "EMPLOYEE CODE")

result = result.merge(
    abstractor_cols,
    left_on = "Abstractor_ID",
    right_on = "EMPLOYEE CODE",
    how = "left"
).drop(columns = "EMPLOYEE CODE")

result = result[[
    "Client Name", "Project Name", "Contract #", "Type of Work", "Reviewer", "Reviewer_ID", "Reviewer_Team", "Duration in Minutes_Reviewer", "Duration in Hours_Reviewer", "Abstractor", "Abstractor_ID", "Abstractor_Team", "Duration in Minutes_Abstractor", "Duration in Hours_Abstractor", "Delivery TL", "TL_ID", "Delivered Date", "SBU Name"
]]

output_filename = input("Enter the desired output file name: ")
if not output_filename.lower().endswith(".xlsx"): output_filename += ".xlsx"
result.to_excel(output_filename, index = False)
