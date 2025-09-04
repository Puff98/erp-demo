# streamlit_erp.py
import streamlit as st
import pandas as pd
import os
from pathlib import Path
from datetime import datetime
from io import BytesIO

# ---------- CONFIG ----------
DEFAULT_EXPORT_DIR = "./exports"
MASTER_FILE = "masters.xlsx"   # saved in the same folder as the app (or you can put in EXPORT_DIR)

# ---------- UTILITIES ----------
def ensure_dir(p: str):
    Path(p).mkdir(parents=True, exist_ok=True)

def month_filename_for(date_obj: datetime, export_dir: str):
    return os.path.join(export_dir, f"{date_obj.year}-{date_obj.month:02d}.xlsx")

def read_sheet_from_file(filepath: str, sheet: str):
    try:
        return pd.read_excel(filepath, sheet_name=sheet, engine="openpyxl")
    except Exception:
        return pd.DataFrame()

def write_workbook_sheets(filepath: str, sheets_dict: dict):
    """
    sheets_dict: {sheet_name: dataframe}
    Overwrites file with provided sheets (creates file if missing).
    """
    with pd.ExcelWriter(filepath, engine="openpyxl", mode="w") as writer:
        for name, df in sheets_dict.items():
            df.to_excel(writer, sheet_name=name, index=False)

def append_row_to_month_file(export_dir: str, date_obj: datetime, sheet_name: str, row: dict):
    ensure_dir(export_dir)
    filepath = month_filename_for(date_obj, export_dir)
    # Load existing sheets
    sheets = {}
    if os.path.exists(filepath):
        try:
            xls = pd.ExcelFile(filepath, engine="openpyxl")
            for s in xls.sheet_names:
                sheets[s] = pd.read_excel(filepath, sheet_name=s, engine="openpyxl")
        except Exception:
            sheets = {}
    # Append to target sheet
    if sheet_name in sheets and isinstance(sheets[sheet_name], pd.DataFrame) and not sheets[sheet_name].empty:
        sheets[sheet_name] = pd.concat([sheets[sheet_name], pd.DataFrame([row])], ignore_index=True)
    else:
        sheets[sheet_name] = pd.DataFrame([row])
    # Ensure the other sheet exists with correct columns (optional)
    other = "Inward" if sheet_name == "Outward" else "Outward"
    if other not in sheets:
        sheets[other] = pd.DataFrame(columns=[])  # will create empty sheet if not present
    write_workbook_sheets(filepath, sheets)
    return filepath

def list_export_files(export_dir: str):
    ensure_dir(export_dir)
    files = [f for f in sorted(os.listdir(export_dir)) if f.endswith(".xlsx")]
    return files

def read_all_monthly(export_dir: str):
    """Read all month files and return combined inward/outward dataframes"""
    ensure_dir(export_dir)
    inward_dfs = []
    outward_dfs = []
    for f in list_export_files(export_dir):
        path = os.path.join(export_dir, f)
        try:
            xls = pd.ExcelFile(path, engine="openpyxl")
            if "Inward" in xls.sheet_names:
                df = pd.read_excel(path, sheet_name="Inward", engine="openpyxl")
                inward_dfs.append(df)
            if "Outward" in xls.sheet_names:
                df = pd.read_excel(path, sheet_name="Outward", engine="openpyxl")
                outward_dfs.append(df)
        except Exception:
            continue
    inward_all = pd.concat(inward_dfs, ignore_index=True) if inward_dfs else pd.DataFrame()
    outward_all = pd.concat(outward_dfs, ignore_index=True) if outward_dfs else pd.DataFrame()
    return inward_all, outward_all

# ---------- MASTER (Customers / Items) Persistence ----------
def load_masters(export_dir: str):
    # Master file located next to app; uses MASTER_FILE name. We keep masters separate so they persist.
    if os.path.exists(MASTER_FILE):
        try:
            customers = pd.read_excel(MASTER_FILE, sheet_name="Customers", engine="openpyxl")
        except Exception:
            customers = pd.DataFrame(columns=["name", "gst_no", "address", "mobile", "email"])
        try:
            items = pd.read_excel(MASTER_FILE, sheet_name="Items", engine="openpyxl")
        except Exception:
            items = pd.DataFrame(columns=["name", "hsn_code", "material"])
    else:
        customers = pd.DataFrame(columns=["name", "gst_no", "address", "mobile", "email"])
        items = pd.DataFrame(columns=["name", "hsn_code", "material"])
    return customers, items

def save_masters(customers_df: pd.DataFrame, items_df: pd.DataFrame):
    sheets = {"Customers": customers_df, "Items": items_df}
    with pd.ExcelWriter(MASTER_FILE, engine="openpyxl", mode="w") as writer:
        customers_df.to_excel(writer, sheet_name="Customers", index=False)
        items_df.to_excel(writer, sheet_name="Items", index=False)

# ---------- STREAMLIT UI ----------
st.set_page_config(page_title="ERP Demo (Streamlit)", layout="wide")
st.title("ERP Demo — Streamlit (Excel storage)")

# Sidebar: export folder path
st.sidebar.header("Settings")
if "EXPORT_DIR" not in st.session_state:
    st.session_state.EXPORT_DIR = DEFAULT_EXPORT_DIR
export_dir_input = st.sidebar.text_input("Export folder path", value=st.session_state.EXPORT_DIR)
if st.sidebar.button("Set Export Path"):
    st.session_state.EXPORT_DIR = export_dir_input.strip() or DEFAULT_EXPORT_DIR
    ensure_dir(st.session_state.EXPORT_DIR)
    st.sidebar.success(f"Export folder set to: {st.session_state.EXPORT_DIR}")

st.sidebar.markdown("**Current Export Folder:**")
st.sidebar.code(st.session_state.EXPORT_DIR)

# Load masters
customers_df, items_df = load_masters(st.session_state.EXPORT_DIR)

tab1, tab2, tab3, tab4, tab5 = st.tabs(["Master", "Inward", "Outward", "Overall", "Export"])

# ---------- MASTER TAB ----------
with tab1:
    st.header("Master Data (Customers & Items)")
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Add Customer")
        with st.form("add_customer", clear_on_submit=True):
            name = st.text_input("Customer name", key="cust_name")
            gst_no = st.text_input("GST No")
            address = st.text_area("Address")
            mobile = st.text_input("Mobile")
            email = st.text_input("Email")
            submitted = st.form_submit_button("Add Customer")
            if submitted:
                if not name.strip():
                    st.warning("Customer name is required")
                else:
                    new = {"name": name.strip(), "gst_no": gst_no.strip(), "address": address.strip(), "mobile": mobile.strip(), "email": email.strip()}
                    customers_df = pd.concat([customers_df, pd.DataFrame([new])], ignore_index=True)
                    save_masters(customers_df, items_df)
                    st.success("Customer added")

        st.markdown("**Customers**")
        st.dataframe(customers_df.reset_index(drop=True))

        # delete customer
        if not customers_df.empty:
            del_idx = st.number_input("Delete customer row (index)", min_value=0, max_value=len(customers_df)-1, value=0)
            if st.button("Delete Customer"):
                customers_df = customers_df.drop(customers_df.index[del_idx]).reset_index(drop=True)
                save_masters(customers_df, items_df)
                st.success("Deleted customer")

    with col2:
        st.subheader("Add Item")
        with st.form("add_item", clear_on_submit=True):
            item_name = st.text_input("Item name", key="item_name")
            hsn = st.text_input("HSN Code")
            material = st.text_input("Material (optional)")
            it_sub = st.form_submit_button("Add Item")
            if it_sub:
                if not item_name.strip():
                    st.warning("Item name required")
                else:
                    new = {"name": item_name.strip(), "hsn_code": hsn.strip(), "material": material.strip()}
                    items_df = pd.concat([items_df, pd.DataFrame([new])], ignore_index=True)
                    save_masters(customers_df, items_df)
                    st.success("Item added")

        st.markdown("**Items**")
        st.dataframe(items_df.reset_index(drop=True))

        # delete item
        if not items_df.empty:
            del_idx2 = st.number_input("Delete item row (index)", min_value=0, max_value=len(items_df)-1, value=0, key="del_item_idx")
            if st.button("Delete Item"):
                items_df = items_df.drop(items_df.index[del_idx2]).reset_index(drop=True)
                save_masters(customers_df, items_df)
                st.success("Deleted item")

# ---------- INWARD TAB ----------
with tab2:
    st.header("Inward Entry")
    # reload masters to ensure latest
    customers_df, items_df = load_masters(st.session_state.EXPORT_DIR)
    cust_list = customers_df["name"].tolist() if not customers_df.empty else []
    item_list = items_df["name"].tolist() if not items_df.empty else []

    if not cust_list or not item_list:
        st.info("Add Customers and Items in Master tab first.")
    else:
        with st.form("inward_form", clear_on_submit=True):
            entry_date = st.date_input("Entry Date", value=datetime.utcnow().date())
            cust = st.selectbox("Customer", options=cust_list)
            item = st.selectbox("Item", options=item_list)
            dc_no_cust = st.text_input("DC No (Customer)")
            qty = st.number_input("Qty / Nos", min_value=0.0, step=1.0)
            rate = st.number_input("Rate", min_value=0.0, step=0.01)
            amt = qty * rate
            st.markdown(f"**Amount = {qty} × {rate} = {amt}**")
            saved = st.form_submit_button("Save Inward")
            if saved:
                # build row
                # find item/gst details
                cust_row = customers_df[customers_df["name"] == cust].iloc[0].to_dict()
                item_row = items_df[items_df["name"] == item].iloc[0].to_dict()
                row = {
                    "entry_date": entry_date.strftime("%Y-%m-%d"),
                    "customer_name": cust_row.get("name", ""),
                    "customer_gst": cust_row.get("gst_no", ""),
                    "item_name": ite_
