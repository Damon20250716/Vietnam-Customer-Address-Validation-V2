
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Vietnam Address Validation Tool", layout="wide")

st.title("🇻🇳 Vietnam Customer Address Validation Tool")

uploaded_form_file = st.file_uploader("Upload Microsoft Forms Responses (Excel)", type=["xlsx"])
uploaded_sys_file = st.file_uploader("Upload System Address Records (Excel)", type=["xlsx"])

if uploaded_form_file and uploaded_sys_file:
    form_df = pd.read_excel(uploaded_form_file)
    sys_df = pd.read_excel(uploaded_sys_file)

    st.success("Files uploaded successfully!")

    form_df.columns = form_df.columns.str.strip()
    sys_df.columns = sys_df.columns.str.strip()

    # Rename form columns for consistency
    form_df = form_df.rename(columns={
        "Account Number": "Account Number",
        "New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only": "New Line 1",
        "New Address Line 2 (Street Name)-In English Only": "New Line 2",
        "New Address Line 3 (Ward/Commune)-In English Only": "New Line 3",
        "City / Province": "New City/Province"
    })

    # Rule 1: Address No. and Street Name must remain the same
    merged = sys_df.merge(form_df, on="Account Number", how="left")

    def compare_line(old, new):
        return str(old).strip().lower() == str(new).strip().lower()

    merged["Line 1 Match"] = merged.apply(lambda row: compare_line(row["Address Line 1"], row["New Line 1"]), axis=1)
    merged["Line 2 Match"] = merged.apply(lambda row: compare_line(row["Address Line 2"], row["New Line 2"]), axis=1)
    merged["Line 3 Changed"] = merged.apply(lambda row: not compare_line(row["Address Line 3"], row["New Line 3"]), axis=1)
    merged["City Changed"] = merged.apply(lambda row: not compare_line(row["City / Province"], row["New City/Province"]), axis=1)

    merged["Address Consistency"] = merged.apply(
        lambda row: "Valid" if row["Line 1 Match"] and row["Line 2 Match"] else "Invalid - Base Address Changed", axis=1
    )

    # Rule 2: Number of pickup addresses must match
    form_pickup_count = form_df.groupby("Account Number").size().reset_index(name="Form Pickup Count")
    system_pickup_count = sys_df[sys_df["Address Type"] == "02"].groupby("Account Number").size().reset_index(name="System Pickup Count")

    pickup_check = pd.merge(system_pickup_count, form_pickup_count, on="Account Number", how="left")
    pickup_check["Pickup Count Match"] = pickup_check["System Pickup Count"] == pickup_check["Form Pickup Count"]

    merged = pd.merge(merged, pickup_check[["Account Number", "Pickup Count Match"]], on="Account Number", how="left")

    merged["Validation Result"] = merged.apply(
        lambda row: "Mismatch - Pickup Count" if not row["Pickup Count Match"]
        else row["Address Consistency"], axis=1
    )

    st.subheader("Validation Results")
    st.dataframe(merged[[
        "Account Number", "Address Type", "Address Line 1", "New Line 1", "Line 1 Match",
        "Address Line 2", "New Line 2", "Line 2 Match",
        "Address Line 3", "New Line 3", "Line 3 Changed",
        "City / Province", "New City/Province", "City Changed",
        "Validation Result"
    ]])

    def convert_df(df):
        return df.to_excel(index=False, engine='openpyxl')

    st.download_button(
        label="📥 Download Validation Results",
        data=convert_df(merged),
        file_name="validation_results.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
