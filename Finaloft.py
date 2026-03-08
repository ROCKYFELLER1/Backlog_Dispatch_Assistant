import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import requests
import certifi
import time
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request
from io import BytesIO

# -------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------

st.set_page_config(page_title="OFT Backlog & Dispatch Assistant", layout="wide")

# -------------------------------------------------
# LOGIN SYSTEM
# -------------------------------------------------

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

USERS = st.secrets["users"]


if not st.session_state["authenticated"]:
    st.title("🔐 OFT Dashboard Login")

    email = st.text_input("Email")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if email in USERS and USERS[email] == password:
            st.session_state["authenticated"] = True
            st.success("Login successful")
            st.rerun()
        else:
            st.error("Invalid email or password")

    st.stop()

# -------------------------------------------------
# GOOGLE AUTHENTICATION
# -------------------------------------------------

scope = ["https://www.googleapis.com/auth/drive.readonly"]

credentials = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=scope,
)

file_id = "11u-AeuFdRbRgl-l6Wk2JaCraSreBc9Cz5oP-KRG31G8"

# -------------------------------------------------
# STYLE
# -------------------------------------------------

st.markdown(
    """
<style>
.main {background-color:#0e1a2b;color:white;}
.block-container {padding-top:2rem;max-width:1400px;}
.filter-box{background:#14263d;padding:20px;border-radius:15px;margin-bottom:25px;}
.card{background:#14263d;padding:18px;border-radius:15px;margin-bottom:18px;height:170px;}
.card-normal{border:1px solid #00D100;}
.card-critical{border:1px solid #ff4d4d;}
.badge{padding:4px 10px;border-radius:10px;font-size:10px;display:inline-block;margin-top:6px;}
.normal{background:#00D100;}
.critical{background:#ff4d4d;}
</style>
""",
    unsafe_allow_html=True,
)

# -------------------------------------------------
# GREETING
# -------------------------------------------------

hour = datetime.now().hour

if hour < 12:
    greeting = "Good Morning,"
elif hour < 17:
    greeting = "Good Afternoon,"
else:
    greeting = "Good Evening,"

st.markdown(
    f"<div style='font-size:18px;margin-bottom:10px;margin-top:12px;'>{greeting}</div>",
    unsafe_allow_html=True,
)

# -------------------------------------------------
# HEADER
# -------------------------------------------------

st.markdown(
    """
<div style="
background:#228B22;
padding:25px;
border-radius:10px;
text-align:center;
color:white;
font-size:28px;
font-weight:bold;
margin-bottom:25px;">
🚚 OFT Backlog & Dispatch Assistant
</div>
""",
    unsafe_allow_html=True,
)

# -------------------------------------------------
# LOAD DATA
# -------------------------------------------------


@st.cache_data(ttl=600, show_spinner=True)
def load_data():

    authed_credentials = credentials.with_scopes(scope)
    authed_credentials.refresh(Request())
    access_token = authed_credentials.token

    url = f"https://www.googleapis.com/drive/v3/files/{file_id}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

    headers = {"Authorization": f"Bearer {access_token}"}

    response = None
    for attempt in range(3):
        try:
            response = requests.get(
                url, headers=headers, timeout=30, verify=certifi.where()
            )
            if response.status_code == 200:
                break
        except requests.exceptions.SSLError:
            time.sleep(2)
        except requests.exceptions.RequestException:
            time.sleep(2)

    if response is None or response.status_code != 200:
        st.error("Drive API Error")
        st.stop()

    file_bytes = BytesIO(response.content)

    df = pd.read_excel(file_bytes, engine="openpyxl")

    df = df.dropna(how="all")

    df.columns = df.columns.str.strip()

    df["SOLDTO"] = df["SOLDTO"].astype(str).str.strip()
    df = df[(df["SOLDTO"] != "") & (df["SOLDTO"] != "nan")]

    numeric_cols = [
        "Backlog",
        "TARGET",
        "ORDERED_QUANTITY",
        "Order_in_New",
        "Order_in_Pool",
    ]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["LOADING_TS"] = pd.to_datetime(df["LOADING_TS"], errors="coerce")
    df = df[df["LOADING_TS"].notna()]
    df["LOADING_DATE"] = df["LOADING_TS"].dt.date

    return df


# -------------------------------------------------
# LOAD DATAFRAME
# -------------------------------------------------

df = load_data()
today = datetime.now().date()

# -------------------------------------------------
# SIDEBAR
# -------------------------------------------------

if st.sidebar.button("Logout"):
    st.session_state["authenticated"] = False
    st.rerun()

# -------------------------------------------------
# DATE FILTER
# -------------------------------------------------

st.sidebar.header("MTD Date Range")

min_date = df["LOADING_DATE"].min()
max_date = df["LOADING_DATE"].max()

start_date = st.sidebar.date_input("Start Date", min_date)
end_date = st.sidebar.date_input("End Date", max_date)

if start_date > end_date:
    st.sidebar.error("End Date must be after Start Date")
    st.stop()

# -------------------------------------------------
# CUSTOMER FILTER
# -------------------------------------------------

customers = sorted(df["SOLDTO"].unique())

selected_customer = st.sidebar.selectbox(
    "Customer Search", ["Select Customer"] + customers
)

# -------------------------------------------------
# ACTION BUTTONS
# -------------------------------------------------

st.sidebar.markdown("### Actions")

fetch_clicked = st.sidebar.button("Fetch Results")
refresh_clicked = st.sidebar.button("🔄 Refresh")

if refresh_clicked:
    st.cache_data.clear()
    if "summary_loaded" in st.session_state:
        del st.session_state["summary_loaded"]
    st.rerun()

if selected_customer == "Select Customer":
    st.info("Please select a customer from the sidebar.")
    st.stop()

# -------------------------------------------------
# LOAD RESULTS
# -------------------------------------------------

if fetch_clicked or "summary_loaded" in st.session_state:

    st.session_state["summary_loaded"] = True

    df = df[(df["LOADING_DATE"] >= start_date) & (df["LOADING_DATE"] <= end_date)]
    df = df[df["SOLDTO"] == selected_customer]

    today_dispatch_value = df[df["LOADING_DATE"] == today]["ORDERED_QUANTITY"].sum()
    past_dispatch_value = df[df["LOADING_DATE"] < today]["ORDERED_QUANTITY"].sum()

    summary = (
        df.groupby(["SOLDTO", "Incoterm", "City", "Type", "Region"])
        .agg(
            Backlog=("Backlog", "first"),
            Target=("TARGET", "first"),
            Order_New=("Order_in_New", "first"),
            Order_Pool=("Order_in_Pool", "first"),
            Dispatch=(
                "ORDERED_QUANTITY",
                lambda x: x[df.loc[x.index, "LOADING_DATE"] < today].sum(),
            ),
        )
        .reset_index()
    )

    summary["Coverage"] = summary["Backlog"] / summary["Target"].replace(0, np.nan)

    alerts = []
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    for _, row in summary.iterrows():

        backlog = row["Backlog"]
        target = row["Target"]

        # backlog but no target
        if target == 0 and backlog > 0:
            alerts.append(
                {
                    "Time": timestamp,
                    "Customer": row["SOLDTO"],
                    "City": row["City"],
                    "Severity": "BLUE",
                    "Message": "Backlog available but no target defined",
                    "Reason": "Customer has backlog orders but no target assigned",
                }
            )

        if backlog == 0:
            alerts.append(
                {
                    "Time": timestamp,
                    "Customer": row["SOLDTO"],
                    "City": row["City"],
                    "Severity": "RED",
                    "Message": "Backlog critically low",
                    "Reason": "Backlog is currently at zero",
                }
            )

        elif backlog < target:
            alerts.append(
                {
                    "Time": timestamp,
                    "Customer": row["SOLDTO"],
                    "City": row["City"],
                    "Severity": "YELLOW",
                    "Message": "Backlog not healthy",
                    "Reason": "Backlog is less than the target available",
                }
            )

        else:
            alerts.append(
                {
                    "Time": timestamp,
                    "Customer": row["SOLDTO"],
                    "City": row["City"],
                    "Severity": "GREEN",
                    "Message": "Backlog healthy",
                    "Reason": "Backlog is more than the target available",
                }
            )

        if row["Order_New"] > 0:
            alerts.append(
                {
                    "Time": timestamp,
                    "Customer": row["SOLDTO"],
                    "City": row["City"],
                    "Severity": "YELLOW",
                    "Message": "Orders stuck in NEW",
                }
            )

        if row["Order_Pool"] > 0:
            alerts.append(
                {
                    "Time": timestamp,
                    "Customer": row["SOLDTO"],
                    "City": row["City"],
                    "Severity": "BLUE",
                    "Message": "Orders waiting in POOL",
                }
            )

    alerts_df = pd.DataFrame(alerts)

    st.markdown('<div class="filter-box">', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    city_options = ["All Cities"] + sorted(summary["City"].unique())
    type_options = ["All Types"] + sorted(summary["Type"].unique())

    with col1:
        selected_city = st.selectbox("City", city_options)

    with col2:
        selected_type = st.selectbox("Type", type_options)

    with col3:
        metric = st.selectbox(
            "Metric", ["Backlog", "Dispatch", "Order_New", "Order_Pool"]
        )

    st.markdown("</div>", unsafe_allow_html=True)

    filtered = summary.copy()

    if selected_city != "All Cities":
        filtered = filtered[filtered["City"] == selected_city]

    if selected_type != "All Types":
        filtered = filtered[filtered["Type"] == selected_type]

    left, right = st.columns([3, 1])

    with left:

        if metric in ["Backlog", "Order_New", "Order_Pool"]:
            total_value = filtered[metric].iloc[0] if not filtered.empty else 0
        else:
            total_value = filtered[metric].sum() if not filtered.empty else 0

        target_value = filtered["Target"].iloc[0] if not filtered.empty else 0

        k1, k2, k3, k4 = st.columns(4)

        with k1:
            st.metric(metric, f"{total_value:,.0f}")

        with k2:
            st.metric("Target", f"{target_value:,.0f}")

        with k3:
            st.metric("MTD Dispatch", f"{past_dispatch_value:,.0f}")

        with k4:
            st.metric("Today Dispatch", f"{today_dispatch_value:,.0f}")

        cols = st.columns(3)

        for i, row in filtered.iterrows():

            col = cols[i % 3]
            value = row[metric]

            card_class = "card card-normal" if value > 0 else "card card-critical"
            badge_class = "normal" if value > 0 else "critical"

            with col:

                st.markdown(
                    f"""
                <div class="{card_class}">
                <div style="font-weight:bold;">{row['City']}</div>
                <div>Type: {row['Type']}</div>
                <div>Incoterm: {row['Incoterm']}</div>

                <div style="margin-top:8px;font-size:20px;">
                {value:,.0f}
                </div>

                <div class="badge {badge_class}">
                {metric}
                </div>
                </div>
                """,
                    unsafe_allow_html=True,
                )

    with right:

        if not alerts_df.empty:

            st.subheader("🚨 Live Alerts")

            alerts_df = alerts_df.drop_duplicates()
            alerts_df = alerts_df.sort_values("Time", ascending=False)

            for _, alert in alerts_df.iterrows():

                message = f"""
Customer: {alert['Customer']}

City: {alert['City']}

{alert['Message']}

Reason:
{alert.get('Reason','')}
"""

                if alert["Severity"] == "RED":
                    st.error(message)

                elif alert["Severity"] == "YELLOW":
                    st.warning(message)

                elif alert["Severity"] == "GREEN":
                    st.success(message)

                else:
                    st.info(message)
