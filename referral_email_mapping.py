import streamlit as st

def get_referral_email(referral_manager: str) -> str:
    if not referral_manager:
        return ""
    name = str(referral_manager).strip()
    try:
        referral_map = st.secrets.get("referral_email", {})
        return referral_map.get(name, "")
    except Exception:
        return ""

# Usage:
# referral_manager = extracted_data.get("ReferralManager", "")
# row["ReferralManager"] = referral_manager
# row["ReferralEmail"] = get_referral_email(referral_manager)
