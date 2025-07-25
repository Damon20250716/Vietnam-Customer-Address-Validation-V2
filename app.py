
import streamlit as st
import pandas as pd
import difflib
from io import BytesIO

st.set_page_config(page_title="Vietnam Address Validation Tool", layout="wide")

def compare_address_parts(old_line, new_line):
    if pd.isna(old_line) or pd.isna(new_line):
        return False
    old_line = str(old_line).lower().strip()
    new_line = str(new_line).lower().strip()
    return old_line == new_line

def highlight_mismatches(row):
    mismatches = []
    if not compare_address_parts(row["Address Line 1"], row["New Line 1"]):
        mismatches.append("Address No. / Industrial Park")
    if not compare_address_parts(row["Address Line 2"], row["New Line 2"]):
        mismatches.append("Street Name")
    if not compare_address_parts(row["Address Line 3"], row["New Line 3"]):
        mismatches.append("Ward/Commune")
    return ", ".join(mismatches) if mismatches else ""

st.title("Vietnam Customer Address Validation Tool")

forms_file = st.file_uploader("Upload Microsoft Forms Response File", type=["xlsx"])
system_file = st.file_uploader("Upload UPS System Address File", type=["xlsx"])
template_file = st.file_uploader("Upload Batch Upload Template", type=["xlsx"])

if forms_file and system_file and template_file:
    forms_df = pd.read_excel(forms_file)
    system_df = pd.read_excel(system_file)
    template_df = pd.read_excel(template_file)

    # Rename columns for consistency
    forms_df.rename(columns={
        "Account Number": "Account",
        "New Pick Up Address Line 1 (Address No., Industrial Park Name, etc)": "New Line 1",
        "New Pick Up Address Line 2 (Street Name)": "New Line 2",
        "New Pick Up Address Line 3 (Ward/Commune)": "New Line 3",
        "Is Your New Billing Address the Same as Your Pickup and Delivery Address?": "Same for All"
    }, inplace=True)

    system_df.rename(columns={
        "Address Line 1": "Address Line 1",
        "Address Line 2": "Address Line 2",
        "Address Line 3": "Address Line 3",
        "Account Number": "Account",
        "Address Type": "Address Type"
    }, inplace=True)

    # Merge based on account
    merged = pd.merge(system_df, forms_df, on="Account", how="left")

    # Validation rules
    merged["Mismatch Reason"] = merged.apply(highlight_mismatches, axis=1)

    matched_df = merged[merged["Mismatch Reason"] == ""].copy()
    unmatched_df = merged[merged["Mismatch Reason"] != ""].copy()

    # Prepare matched file for batch upload using template
    batch_upload = template_df.copy()
    batch_rows = []
    for _, row in matched_df.iterrows():
        if row["Same for All"].lower() == "yes":
            address_types = ["01"]  # All
        else:
            address_types = ["02", "03", "13"]  # Pickup, Billing, Delivery
        for code in address_types:
            batch_row = batch_upload.iloc[0].copy()
            batch_row["Account Number"] = row["Account"]
            batch_row["Address Type"] = code
            batch_row["Address Line 1"] = row["New Line 1"]
            batch_row["Address Line 2"] = row["New Line 2"]
            batch_row["Address Line 3"] = row["New Line 3"]
            batch_rows.append(batch_row)
    final_batch_df = pd.DataFrame(batch_rows)

    # Save results to Excel in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        final_batch_df.to_excel(writer, index=False, sheet_name="Matched For Upload")
        unmatched_df.to_excel(writer, index=False, sheet_name="Unmatched With Reasons")
    st.download_button("Download Validation Results", output.getvalue(), file_name="validation_results.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
