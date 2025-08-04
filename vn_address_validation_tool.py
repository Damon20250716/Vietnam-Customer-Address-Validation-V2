
import streamlit as st
import pandas as pd
import unicodedata

# Remove Vietnamese tones
def remove_tones(text):
    if not isinstance(text, str):
        return text
    text = unicodedata.normalize('NFD', text)
    return ''.join(ch for ch in text if unicodedata.category(ch) != 'Mn')

# Normalize string (lowercase, stripped, no tones)
def normalize(text):
    if not isinstance(text, str):
        return text
    return remove_tones(text.strip().lower())

# Check if street and number are same (before the first comma)
def match_address_base(old, new):
    if pd.isna(old) or pd.isna(new):
        return False
    old_base = normalize(old.split(",")[0])
    new_base = normalize(new.split(",")[0])
    return old_base == new_base

# Streamlit UI
st.title("Vietnam Address Validation Tool")
forms_file = st.file_uploader("Upload Microsoft Forms Response", type=["xlsx"])
ups_file = st.file_uploader("Upload UPS System Data", type=["xlsx"])

if forms_file and ups_file:
    forms_df = pd.read_excel(forms_file)
    ups_df = pd.read_excel(ups_file)

    # Strip tone marks and normalize Forms addresses for matching
    for col in ['New Address Line 1', 'New Address Line 2', 'New Address Line 3']:
        if col in forms_df.columns:
            forms_df[f"{col}_tonefree"] = forms_df[col].apply(remove_tones)

    matched_rows = []
    unmatched_rows = []

    for _, row in forms_df.iterrows():
        account = row['Account Number']
        same_all = row.get('Is Your New Billing Address the Same as Your Pickup and Delivery Address?', '').strip().lower() == 'yes'
        new1 = row.get('New Address Line 1', '')
        new2 = row.get('New Address Line 2', '')
        new3 = row.get('New Address Line 3', '')

        new_full = ', '.join(filter(None, [new1, new2, new3]))
        new_base = normalize(new1.split(",")[0])  # only Address Line 1 checked

        ups_account_rows = ups_df[ups_df['Account Number'] == account]
        ups_pickups = ups_account_rows[ups_account_rows['Address Type Code'] == 2]
        match_found = False
        reason = ""

        if same_all:
            # Check all address types with 01
            for _, sys_row in ups_pickups.iterrows():
                old1 = sys_row.get('Address Line 1', '')
                if match_address_base(old1, new1):
                    matched_rows.append({
                        'Account Number': account,
                        'Address Type': '01',
                        'New Address Line 1': remove_tones(new1),
                        'New Address Line 2': remove_tones(new2),
                        'New Address Line 3': remove_tones(new3)
                    })
                    match_found = True
                    break
            if not match_found:
                reason = "Address Number or Street Name Changed"

        else:
            # Validate number of pickup addresses
            form_pickups = int(row.get('Number of Pickup Addresses', 0))
            if form_pickups != len(ups_pickups):
                reason = "Pickup Address Count Mismatch"
            else:
                for _, sys_row in ups_pickups.iterrows():
                    old1 = sys_row.get('Address Line 1', '')
                    if match_address_base(old1, new1):
                        matched_rows.append({
                            'Account Number': account,
                            'Address Type': '02',
                            'New Address Line 1': remove_tones(new1),
                            'New Address Line 2': remove_tones(new2),
                            'New Address Line 3': remove_tones(new3)
                        })
                        match_found = True
                        break
                if not match_found:
                    reason = "Address Number or Street Name Changed"

        if not match_found:
            unmatched = row.to_dict()
            unmatched['Reason'] = reason
            unmatched_rows.append(unmatched)

    matched_df = pd.DataFrame(matched_rows)
    unmatched_df = pd.DataFrame(unmatched_rows)

    st.success(f"{len(matched_df)} matched, {len(unmatched_df)} unmatched")
    st.download_button("Download Matched", matched_df.to_csv(index=False), file_name="matched.csv")
    st.download_button("Download Unmatched", unmatched_df.to_csv(index=False), file_name="unmatched.csv")
