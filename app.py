import io
import pandas as pd
import streamlit as st
import edis_core as core

st.set_page_config(page_title="EDIS Checker", page_icon="📊", layout="wide")
st.title("📊 EDIS Checker")
st.caption("Upload today's CM1 file and Meroshare transfer report to see who has not transferred their shares. Trades are matched on Contract No + Client Code + Scrip + Quantity.")

c1, c2, c3 = st.columns(3)
cm_file = c1.file_uploader("1. CM1 file (xls / xlsx / csv)", type=["xls", "xlsx", "csv"])
mero_file = c2.file_uploader("2. Meroshare transfer report (xls / xlsx / csv)", type=["csv", "xls", "xlsx"])
contact_file = c3.file_uploader("3. Commission report (name, code, mobile)", type=["xls", "xlsx", "csv"])

if not (cm_file and mero_file):
    st.info("Please upload the CM1 file and the Meroshare transfer report to continue.")
    st.stop()

try:
    cm1 = core.prepare_cm1(cm_file)
    mero = core.prepare_mero(mero_file)
except Exception as e:
    st.error(f"Could not read the files: {e}")
    st.stop()

# ---- contacts (optional) ----
contacts = None
if contact_file:
    try:
        contacts = core.prepare_contacts(contact_file)
        st.caption(f"Contact list loaded: {len(contacts)} clients.")
    except Exception as e:
        st.warning(f"Contact file could not be read ({e}). Showing client codes only.")

s = core.build_summary(cm1, mero)
fmt = lambda x: f"{int(x):,}"

st.divider()
m1, m2, m3 = st.columns(3)
m1.metric("Total quantity to be received (CM1)", fmt(s["total_req"]))
m2.metric("Quantity received till now (Meroshare)", fmt(s["total_rec"]))
m3.metric("Difference", fmt(s["diff"]))

if s["diff"] == 0 and s["detail"].empty:
    st.success("✅ EDIS is completed.")
elif s["detail"].empty:
    st.warning("Every CM1 entry has been received, but the totals differ. Check the extra transfers section below.")
else:
    st.error(f"❌ EDIS is NOT completed. {fmt(s['diff'])} units are still to be received.")

if contacts is not None:
    pend_ids = set(s["status"][s["status"] != "done"].index)
    missing = sorted(pend_ids - set(contacts.client))
    if missing:
        st.warning(f"{len(missing)} client(s) with pending EDIS are not in the contact list: {', '.join(missing)}")

not_done = core.client_table(s, "not_done", contacts)
partial = core.client_table(s, "partial", contacts)

st.subheader(f"🔴 Not transferred at all ({len(not_done)} clients)")
if not_done.empty:
    st.write("None.")
else:
    st.dataframe(not_done, use_container_width=True, hide_index=True)

st.subheader(f"🟠 Partially transferred ({len(partial)} clients)")
if partial.empty:
    st.write("None.")
else:
    st.dataframe(partial, use_container_width=True, hide_index=True)

# Scrip-wise pending detail
det = s["detail"].rename(columns={"client": "Client Code", "scrip": "Scrip", "required": "Required",
                                  "received": "Received", "pending": "Pending"})
det = core.with_contacts(det.rename(columns={"Client Code": "client"}), contacts).rename(columns={"client": "Client Code"})
det = det[["Client Code", "name", "phone", "Scrip", "Required", "Received", "Pending"]].rename(
    columns={"name": "Name", "phone": "Phone"}).sort_values(["Client Code", "Scrip"])
with st.expander("Scrip-wise pending detail"):
    st.dataframe(det, use_container_width=True, hide_index=True)

extra = s["extra"]
if not extra.empty:
    st.subheader("⚠️ Extra transfers (received but not matching any CM1 trade)")
    st.dataframe(extra.rename(columns={"contract": "Contract No", "client": "Client Code",
                                       "scrip": "Scrip", "extra": "Quantity"}),
                 use_container_width=True, hide_index=True)

# Excel download
buf = io.BytesIO()
with pd.ExcelWriter(buf, engine="openpyxl") as w:
    pd.DataFrame({"Item": ["Total to be received", "Received till now", "Difference"],
                  "Quantity": [s["total_req"], s["total_rec"], s["diff"]]}).to_excel(w, sheet_name="Summary", index=False)
    not_done.to_excel(w, sheet_name="Not transferred", index=False)
    partial.to_excel(w, sheet_name="Partial", index=False)
    det.to_excel(w, sheet_name="Scrip detail", index=False)
st.download_button("⬇️ Download report as Excel", buf.getvalue(), "EDIS_report.xlsx",
                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
