import pandas as pd

filename_delivery = input("Enter the Delivery Details file name: ")
delivery = pd.read_excel(filename_delivery)

filename_timesheet = input("Enter the Timesheet file name: ")
timesheet = pd.read_excel(filename_timesheet)

filename_resource_tracker = input("Enter the Resource Tracker file name: ")
resource_tracker = pd.read_excel(filename_resource_tracker, sheet_name= "RESOURCES")
resource_tracker = resource_tracker[[
    "EMPLOYEE CODE", "CORE TEAM/CLIENT"
]]

def role_summary(role, suffix):
    role_df = timesheet[
        timesheet["Role Name"].fillna("") == role
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

result = delivery.merge(
    summary, on = "Contract #", how = "inner"
)
########### inner or outer?

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
