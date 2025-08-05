import streamlit as st
import pandas as pd
from io import BytesIO
import unicodedata  # For removing Vietnamese tones

# Remove Vietnamese tones from text
def remove_tones(text):
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFD', text)
    text = ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')
    return text

# Normalize string columns: lowercase, strip spaces, remove tones
def normalize_col(col):
    return col.astype(str).str.lower().str.strip().apply(remove_tones)

# Match address line 1 & 2 (number + street), ignoring case, spaces and tones
def address_match(new_line1, new_line2, old_line1, old_line2):
    # Remove tones and normalize all compared strings
    nl1 = remove_tones(str(new_line1).strip().lower())
    nl2 = remove_tones(str(new_line2).strip().lower())
    ol1 = remove_tones(str(old_line1).strip().lower())
    ol2 = remove_tones(str(old_line2).strip().lower())
    return (nl1 == ol1) and (nl2 == ol2)

# Main validation and processing function
def process_files(forms_df, ups_df):
    matched_rows = []
    unmatched_rows = []
    upload_template_rows = []

    # Normalize Account Number for matching (no tone removal here, just lowercase & strip)
    ups_df['Account Number_norm'] = ups_df['Account Number'].astype(str).str.lower().str.strip()
    forms_df['Account Number_norm'] = forms_df['Account Number'].astype(str).str.lower().str.strip()

    # Group UPS data by normalized account number for easy lookup
    ups_grouped = ups_df.groupby('Account Number_norm')

    # Track processed Forms rows by index to find unmatched later
    processed_form_indices = set()

    for idx, form_row in forms_df.iterrows():
        acc_norm = form_row['Account Number_norm']
        is_same_billing = str(form_row.get("Is Your New Billing Address the Same as Your Pickup and Delivery Address?", "")).strip().lower()

        if acc_norm not in ups_grouped.groups:
            # No UPS data for this account, unmatched with reason
            unmatched_dict = form_row.to_dict()
            unmatched_dict['Unmatched Reason'] = "Account Number not found in UPS data"
            unmatched_rows.append(unmatched_dict)
            continue

        ups_acc_df = ups_grouped.get_group(acc_norm)

        # Count pickup addresses in UPS system for this account
        ups_pickup_count = (ups_acc_df['Address Type'] == '02').sum()

        if is_same_billing == "yes":
            # Single "All" address type 01 from Forms
            new_addr1 = form_row["New Address Line 1 (Address No., Industrial Park Name, etc)-In English Only"]
            new_addr2 = form_row["New Address Line 2 (Street Name)-In English Only"]
            new_addr3 = form_row["New Address Line 3 (Ward/Commune)-In English Only"]
            city = form_row["City / Province"]
            email = form_row.get("Please Provide Your Email Address-In English Only", "")
            contact = form_row.get("Full Name of Contact-In English Only", "")
            phone = form_row.get("Contact Phone Number", "")

            # Check if this address matches any UPS address lines (Line1+Line2) with tone removal
            matched_in_ups = False
            ups_row_for_template = None
            for _, ups_row in ups_acc_df.iterrows():
                if address_match(new_addr1, new_addr2, ups_row["Address Line 1"], ups_row["Address Line 2"]):
                    matched_in_ups = True
                    ups_row_for_template = ups_row
                    break

            if matched_in_ups:
                # Add tone-free address lines to matched output
                matched_dict = form_row.to_dict()
                matched_dict["New Address Line 1 (Tone-free)"] = remove_tones(new_addr1)
                matched_dict["New Address Line 2 (Tone-free)"] = remove_tones(new_addr2)
                matched_dict["New Address Line 3 (Tone-free)"] = remove_tones(new_addr3)
                matched_rows.append(matched_dict)
                processed_form_indices.add(idx)

                # Upload template requires 3 rows with codes 1, 2, 6
                for code in ["1", "2", "6"]:
                    upload_template_rows.append({
                        "AC_NUM": form_row["Account Number"],
                        "AC_Address_Type": code,
                        "AC_Name": ups_row_for_template["AC_Name"],
                        "Address_Line1": remove_tones(new_addr1),
                        "Address_Line2": remove_tones(new_addr2),
                        "City": city,
                        "Postal_Code": ups_row_for_template["Postal_Code"],
                        "Country_Code": ups_row_for_template["Country_Code"],
                        "Attention_Name": contact,
                        "Address_Line22": remove_tones(new_addr3),
                        "Address_Country_Code": ups_row_for_template["Address_Country_Code"]
                    })
            else:
                unmatched_dict = form_row.to_dict()
                unmatched_dict['Unmatched Reason'] = "Billing address not matched in UPS system"
                unmatched_rows.append(unmatched_dict)

        else:
            # When "No" for billing same as pickup/delivery,
            # Process billing, delivery, and up to 3 pickup addresses separately.

            # Billing address fields
            billing_addr1 = form_row.get("New Billing Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            billing_addr2 = form_row.get("New Billing Address Line 2 (Street Name)-In English Only", "")
            billing_addr3 = form_row.get("New Billing Address Line 3 (Ward/Commune)-In English Only", "")
            billing_city = form_row.get("New Billing City / Province", "")

            # Delivery address fields
            delivery_addr1 = form_row.get("New Delivery Address Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
            delivery_addr2 = form_row.get("New Delivery Address Line 2 (Street Name)-In English Only", "")
            delivery_addr3 = form_row.get("New Delivery Address Line 3 (Ward/Commune)-In English Only", "")
            delivery_city = form_row.get("New Delivery City / Province", "")

            # Number of pickup addresses in form
            pickup_num = 0
            try:
                pickup_num = int(form_row.get("How Many Pick Up Address Do You Have?", 0))
            except:
                pickup_num = 0
            if pickup_num > 3:
                pickup_num = 3  # max 3 pickups

            # Check pickup addresses from the Form: First, Second, Third New Pick Up Address
            pickup_addrs = []
            for i in range(1, pickup_num + 1):
                prefix = ["First", "Second", "Third"][i-1] + " New Pick Up Address"
                pu_addr1 = form_row.get(f"{prefix} Line 1 (Address No., Industrial Park Name, etc)-In English Only", "")
                pu_addr2 = form_row.get(f"{prefix} Line 2 (Street Name)-In English Only", "")
                pu_addr3 = form_row.get(f"{prefix} Line 3 (Ward/Commune)-In English Only", "")
                pu_city = form_row.get(f"{prefix} City / Province", "")
                pickup_addrs.append((pu_addr1, pu_addr2, pu_addr3, pu_city))

            # Validate address number and street for each type: must exist in UPS system for the account
            def check_address_in_ups(addr1, addr2, addr_type_code):
                for _, ups_row in ups_acc_df.iterrows():
                    if ups_row["Address Type"] == addr_type_code:
                        if address_match(addr1, addr2, ups_row["Address Line 1"], ups_row["Address Line 2"]):
                            return ups_row
                return None

            # Flags to track matching for billing/delivery/pickups
            billing_match = check_address_in_ups(billing_addr1, billing_addr2, "03")
            delivery_match = check_address_in_ups(delivery_addr1, delivery_addr2, "13")

            pickup_matches = []
            if len(pickup_addrs) != ups_pickup_count:
                # Number of pickup addresses mismatch => unmatched
                unmatched_dict = form_row.to_dict()
                unmatched_dict['Unmatched Reason'] = f"Pickup address count mismatch: Forms={len(pickup_addrs)}, UPS={ups_pickup_count}"
                unmatched_rows.append(unmatched_dict)
                continue
            else:
                # Check each pickup address
                for pu_addr in pickup_addrs:
                    match = check_address_in_ups(pu_addr[0], pu_addr[1], "02")
                    if match is None:
                        unmatched_dict = form_row.to_dict()
                        unmatched_dict['Unmatched Reason'] = f"Pickup address not matched: {pu_addr[0]}, {pu_addr[1]}"
                        unmatched_rows.append(unmatched_dict)
                        break
                    else:
                        pickup_matches.append(match)
                else:
                    # All pickups matched
                    processed_form_indices.add(idx)
                    # Add tone-free address lines in matched output for billing, delivery, pickups
                    matched_dict = form_row.to_dict()
                    matched_dict["New Billing Address Line 1 (Tone-free)"] = remove_tones(billing_addr1)
                    matched_dict["New Billing Address Line 2 (Tone-free)"] = remove_tones(billing_addr2)
                    matched_dict["New Billing Address Line 3 (Tone-free)"] = remove_tones(billing_addr3)
                    matched_dict["New Delivery Address Line 1 (Tone-free)"] = remove_tones(delivery_addr1)
                    matched_dict["New Delivery Address Line 2 (Tone-free)"] = remove_tones(delivery_addr2)
                    matched_dict["New Delivery Address Line 3 (Tone-free)"] = remove_tones(delivery_addr3)
                    for i, pu_addr in enumerate(pickup_addrs, 1):
                        matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 1 (Tone-free)"] = remove_tones(pu_addr[0])
                        matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 2 (Tone-free)"] = remove_tones(pu_addr[1])
                        matched_dict[f"{['First', 'Second', 'Third'][i-1]} New Pick Up Address Line 3 (Tone-free)"] = remove_tones(pu_addr[2])
                    matched_rows.append(matched_dict)

                    # Add pickup addresses to upload template
                    for pu_addr in pickup_addrs:
                        upload_template_rows.append({
                            "AC_NUM": form_row["Account Number"],
                            "AC_Address_Type": "02",
                            "AC_Name": ups_acc_df["AC_Name"].values[0],  # pick any AC_Name from UPS group
                            "Address_Line1": remove_tones(pu_addr[0]),
                            "Address_Line2": remove_tones(pu_addr[1]),
                            "City": pu_addr[3],
                            "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                            "Country_Code": ups_acc_df["Country_Code"].values[0],
                            "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                            "Address_Line22": remove_tones(pu_addr[2]),
                            "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
                        })

                    # Add billing address (expand to 3 codes 1, 2, 6)
                    for code in ["1", "2", "6"]:
                        upload_template_rows.append({
                            "AC_NUM": form_row["Account Number"],
                            "AC_Address_Type": code,
                            "AC_Name": ups_acc_df["AC_Name"].values[0],
                            "Address_Line1": remove_tones(billing_addr1),
                            "Address_Line2": remove_tones(billing_addr2),
                            "City": billing_city,
                            "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                            "Country_Code": ups_acc_df["Country_Code"].values[0],
                            "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                            "Address_Line22": remove_tones(billing_addr3),
                            "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
                        })

                    # Add delivery address
                    upload_template_rows.append({
                        "AC_NUM": form_row["Account Number"],
                        "AC_Address_Type": "13",
                        "AC_Name": ups_acc_df["AC_Name"].values[0],
                        "Address_Line1": remove_tones(delivery_addr1),
                        "Address_Line2": remove_tones(delivery_addr2),
                        "City": delivery_city,
                        "Postal_Code": ups_acc_df["Postal_Code"].values[0],
                        "Country_Code": ups_acc_df["Country_Code"].values[0],
                        "Attention_Name": form_row.get("Full Name of Contact-In English Only", ""),
                        "Address_Line22": remove_tones(delivery_addr3),
                        "Address_Country_Code": ups_acc_df["Address_Country_Code"].values[0]
                    })

    # Add Forms rows that never matched into unmatched with reason
    unmatched_not_processed = forms_df.loc[~forms_df.index.isin(processed_form_indices)]
    for _, row in unmatched_not_processed.iterrows():
        unmatched_dict = row.to_dict()
        unmatched_dict['Unmatched Reason'] = "No matching address found or not processed"
        unmatched_rows.append(unmatched_dict)

    # Convert lists of dicts back to DataFrames
    matched_df = pd.DataFrame(matched_rows)
    unmatched_df = pd.DataFrame(unmatched_rows)
    upload_template_df = pd.DataFrame(upload_template_rows)

    return matched_df, unmatched_df, upload_template_df

# --- Streamlit UI ---
def main():
    st.set_page_config(page_title="Vietnam Address Validation Tool", layout="wide")
    st.title("🇻🇳 Vietnam Address Validation Tool")
    st.write("Upload Microsoft Forms response file and UPS system address file to validate and generate upload template.")

    forms_file = st.file_uploader("Upload Microsoft Forms Response File (.xlsx)", type=["xlsx"])
    ups_file = st.file_uploader("Upload UPS System Address File (.xlsx)", type=["xlsx"])

    if forms_file and ups_file:
        with st.spinner("Processing files..."):
            forms_df = pd.read_excel(forms_file)
            ups_df = pd.read_excel(ups_file)

            matched_df, unmatched_df, upload_template_df = process_files(forms_df, ups_df)

            st.success(f"✅ Completed: {len(matched_df)} matched, {len(unmatched_df)} unmatched.")

            def to_excel_bytes(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False)
                return output.getvalue()

            if not matched_df.empty:
                st.download_button(
                    label="Download Matched Records",
                    data=to_excel_bytes(matched_df),
                    file_name="matched_records.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            if not unmatched_df.empty:
                st.download_button(
                    label="Download Unmatched Records",
                    data=to_excel_bytes(unmatched_df),
                    file_name="unmatched_records.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            if not upload_template_df.empty:
                st.download_button(
                    label="Download Upload Template",
                    data=to_excel_bytes(upload_template_df),
                    file_name="upload_template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )


if __name__ == "__main__":
    main()
