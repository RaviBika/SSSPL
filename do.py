import streamlit as st
import pandas as pd
import requests
from io import BytesIO

# ----------------------------
# Helper Functions
# ----------------------------
def fetch_do_list(url, username, password):
    resp = requests.get(url, auth=(username, password))
    resp.raise_for_status()
    return resp.json().get("value", [])


def fetch_do_containers(url, username, password):
    resp = requests.get(url, auth=(username, password))
    resp.raise_for_status()
    return resp.json().get("value", [])


def process_data(excel_df, do_list, do_containers):
    results = []

    # Detect container column
    possible_cols = [c for c in excel_df.columns if "container" in c.lower()]
    if possible_cols:
        container_col = possible_cols[0]
    else:
        container_col = st.selectbox("Select Container Column", excel_df.columns)

    for _, row in excel_df.iterrows():
        container = str(row[container_col]).strip()
        match = next((c for c in do_containers if c.get("Container_No") == container), None)

        if match:
            doc_no = match.get("Document_No")
            do_info = next((d for d in do_list if d.get("DO_No") == doc_no), None)
            if do_info:
                do_date = do_info.get("DO_Date")
                do_date_fmt = pd.to_datetime(do_date).strftime("%d-%m-%Y") if do_date else None
                results.append({
                    "Excel Container": container,
                    "Matched DO No": doc_no,
                    "DO Date": do_date_fmt,
                    "Branch": do_info.get("Branch"),
                    "Status": "Matched ✅"
                })
            else:
                results.append({
                    "Excel Container": container,
                    "Matched DO No": "-",
                    "DO Date": None,
                    "Branch": None,
                    "Status": "Container Found, DO Missing ⚠️"
                })
        else:
            results.append({
                "Excel Container": container,
                "Matched DO No": "-",
                "DO Date": None,
                "Branch": None,
                "Status": "Not Found ❌"
            })
    return pd.DataFrame(results)


def prepare_export_excel(df, item_code, rate, gst_code, hsn, tds):
    output = []
    for _, row in df.iterrows():
        if row.get("Status") == "Matched ✅":  # only export matched
            output.append({
                "Item No.": item_code,
                "Doc Type": "DO Export",
                "Container No.": row["Excel Container"],
                "DO/SO No.": row["Matched DO No"],
                "Quantity": 1,
                "Direct UnitCost Excl. VAT": rate,
                "GST Group Code": gst_code,
                "HSN/SAC Code": hsn,
                "TDS Section Code": tds
            })
    return pd.DataFrame(output)


# ----------------------------
# Streamlit UI
# ----------------------------
st.set_page_config(layout="wide")

st.title("🚢 DO - Container Matching Helper")

left_col, right_col = st.columns([1, 3])

with left_col:
    st.subheader("🔑 BC Login")
    bc_user = st.text_input("BC Username")
    bc_pass = st.text_input("BC Password", type="password")

    st.markdown("### Upload Vendor Excel")
    excel_file = st.file_uploader("Upload Excel with Containers", type=["xlsx"])

    st.markdown("### Export Configuration")
    item_code = st.text_input("Item Code", value="ITEM001")
    item_rate = st.number_input("Rate", value=0.0)
    gst_code = st.text_input("GST Group Code", value="GST18")
    hsn_code = st.text_input("HSN/SAC Code", value="HSN00001")
    tds_section = st.text_input("TDS Section Code", value="TDS194C")

    recalc_btn = st.button("🔄 Recalculate")


with right_col:
    st.subheader("📊 Results")

    # Recalculation trigger
    if recalc_btn and excel_file and bc_user and bc_pass:
        df_excel = pd.read_excel(excel_file)

        # Fetch BC Data
        do_list = fetch_do_list(
            "http://111.90.175.54:61048/SSSPL/ODataV4/Company('SSSPL%202025')/ExportDoList",
            bc_user, bc_pass
        )
        do_containers = fetch_do_containers(
            "http://111.90.175.54:61048/SSSPL/ODataV4/Company('SSSPL%202025')/DoList",
            bc_user, bc_pass
        )

        st.session_state["matched_df"] = process_data(df_excel, do_list, do_containers)

    # Display results as long as we have them stored
    if "matched_df" in st.session_state:
        matched_df = st.session_state["matched_df"]

        # Metrics
        total_excel = matched_df.shape[0]
        total_matched = matched_df[matched_df["Status"] == "Matched ✅"].shape[0]
        total_unmatched = matched_df[matched_df["Status"] != "Matched ✅"].shape[0]

        k1, k2, k3 = st.columns(3)
        k1.metric("📦 Total in Excel", total_excel)
        k2.metric("✅ Matched in BC", total_matched)
        k3.metric("❌ Unmatched", total_unmatched)

        st.markdown("### 🔍 Detailed View")
        view_choice = st.radio("Filter containers:", ["All", "Only Matched", "Only Unmatched"], horizontal=True)

        if view_choice == "Only Matched":
            view_df = matched_df[matched_df["Status"] == "Matched ✅"]
        elif view_choice == "Only Unmatched":
            view_df = matched_df[matched_df["Status"] != "Matched ✅"]
        else:
            view_df = matched_df

        # Clean native table
        st.dataframe(
            view_df.style.set_properties(**{
                'text-align': 'center',
                'background-color': '#f9f9f9'
            }),
            height=600,
            use_container_width=True
        )

        # Export matched only
        export_df = prepare_export_excel(matched_df, item_code, item_rate, gst_code, hsn_code, tds_section)
        towrite = BytesIO()
        export_df.to_excel(towrite, index=False, sheet_name="Export")
        towrite.seek(0)

        st.download_button("📥 Download Excel", data=towrite,
                           file_name="Processed_DO.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
