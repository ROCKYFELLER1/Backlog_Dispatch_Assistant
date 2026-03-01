import streamlit as st
import pandas as pd
import openpyxl
import requests
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request
from io import BytesIO


# PAGE CONFIG


st.set_page_config(
    page_title="OFT AI Backlog Assistant",
    layout="wide",
)


# GOOGLE AUTHENTICATION


scope = ["https://www.googleapis.com/auth/drive.readonly"]

credentials = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=scope,
)

file_id = "11u-AeuFdRbRgl-l6Wk2JaCraSreBc9Cz5oP-KRG31G8"


# HEADER


st.markdown(
    """
    <div style="
        background:#228B22;
        padding:25px;
        border-radius:10px;
        text-align:center;
        color:white;
        font-size:22px;
        font-weight:bold;
        margin-bottom:25px;">
        OFT Backlog & Dispatch Assistant
    </div>
    """,
    unsafe_allow_html=True,
)


# LOAD DATA



@st.cache_data(show_spinner=True)
def load_data():

    authed_credentials = credentials.with_scopes(scope)
    authed_credentials.refresh(Request())
    access_token = authed_credentials.token

    url = f"https://www.googleapis.com/drive/v3/files/{file_id}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    headers = {"Authorization": f"Bearer {access_token}"}

    response = requests.get(url, headers=headers)

    if response.status_code != 200:
        st.error("Drive API Error")
        st.stop()

    file_bytes = BytesIO(response.content)

    df_raw = pd.read_excel(file_bytes, engine="openpyxl")

    df_raw.columns = df_raw.columns.str.strip()

    df_raw["SOLDTO"] = df_raw["SOLDTO"].astype(str).str.strip()
    df_raw["City"] = df_raw["City"].astype(str).str.strip()
    df_raw["Type"] = df_raw["Type"].astype(str).str.strip()
    df_raw["Incoterm"] = df_raw["Incoterm"].astype(str).str.strip()
    df_raw["Status Summary"] = df_raw["Status Summary"].str.lower().str.strip()

    df_raw["ORDERED_QUANTITY"] = pd.to_numeric(
        df_raw["ORDERED_QUANTITY"], errors="coerce"
    ).fillna(0)

    group_cols = ["SOLDTO", "City", "Type", "Incoterm"]

    backlog = (
        df_raw[df_raw["Status Summary"] == "backlog"]
        .groupby(group_cols)["ORDERED_QUANTITY"]
        .sum()
        .reset_index()
        .rename(columns={"ORDERED_QUANTITY": "Backlog"})
    )

    mtd = (
        df_raw[df_raw["Status Summary"] == "dispatched"]
        .groupby(group_cols)["ORDERED_QUANTITY"]
        .sum()
        .reset_index()
        .rename(columns={"ORDERED_QUANTITY": "MTD"})
    )

    df = pd.merge(backlog, mtd, on=group_cols, how="outer").fillna(0)

    return df


df = load_data()

customers = sorted(df["SOLDTO"].unique())


# SIDEBAR FILTERS


st.sidebar.header("Filters")

selected_customer = st.sidebar.selectbox(
    "Select Customer",
    customers,
    index=None,
)

if selected_customer:
    customer_df = df[df["SOLDTO"] == selected_customer]

    city_options = sorted(customer_df["City"].unique())
    city_options.insert(0, "All Cities")

    type_options = sorted(customer_df["Type"].unique())
    type_options.insert(0, "All Types")
else:
    city_options = []
    type_options = []

selected_city = st.sidebar.selectbox(
    "Select City",
    city_options,
    index=0 if city_options else None,
)

selected_type = st.sidebar.selectbox(
    "Select Type",
    type_options,
    index=0 if type_options else None,
)

question_type = st.sidebar.selectbox(
    "Select Metric",
    ["Backlog", "MTD"],
)

fetch = st.sidebar.button("Fetch Result")

if st.sidebar.button("Clear Search"):
    st.session_state.clear()
    st.rerun()

if st.sidebar.button("Reset App"):
    st.cache_data.clear()
    st.session_state.clear()
    st.rerun()


# RESULT DISPLAY


if fetch:

    if not selected_customer:
        st.warning("Please select a customer.")
        st.stop()

    filtered_df = df[df["SOLDTO"] == selected_customer]

    if selected_city and selected_city != "All Cities":
        filtered_df = filtered_df[filtered_df["City"] == selected_city]

    if selected_type and selected_type != "All Types":
        filtered_df = filtered_df[filtered_df["Type"] == selected_type]

    if filtered_df.empty:
        st.error("No data found.")
        st.stop()

    breakdown = filtered_df.groupby(["City", "Type", "Incoterm"], as_index=False).agg(
        {"Backlog": "sum", "MTD": "sum"}
    )

    if question_type == "Backlog":

        total_backlog = breakdown["Backlog"].sum()

        st.markdown("## Backlog Summary")
        st.metric("Total Backlog", f"{total_backlog:,.2f}")

        st.dataframe(
            breakdown[["City", "Type", "Incoterm", "Backlog"]],
            use_container_width=True,
            hide_index=True,
        )

    else:

        total_mtd = breakdown["MTD"].sum()

        st.markdown("## MTD Summary")
        st.metric("Total MTD", f"{total_mtd:,.2f}")

        st.dataframe(
            breakdown[["City", "Type", "Incoterm", "MTD"]],
            use_container_width=True,
            hide_index=True,
        )
