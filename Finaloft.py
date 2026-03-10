import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import openpyxl
import requests
import certifi
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request
from io import BytesIO

# -------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------

st.set_page_config(page_title="OFT Backlog & Dispatch Assistant", layout="wide")

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
.card{
    background:#14263d;
    padding:18px;
    border-radius:15px;
    margin-bottom:18px;
    min-height:260px;
    width:100%;
    overflow-x:auto;
}
.card-normal{border:3px solid #00D100;}
.card-critical{border:3px solid #ff0000;}
.real-table{
    width:100%;
    margin-top:0;
    border-collapse:collapse;
    table-layout:auto;
    text-align:center;
    font-size:12px;
    color:white;
}
.real-table th{
    border:1px solid #3b4f6b;
    padding:6px 4px;
    background:#1a2f4a;
    color:white;
}
.real-table td{
    border:1px solid #3b4f6b;
    padding:8px 4px;
    color:white;
    font-weight:bold;
}
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


@st.cache_data(show_spinner=True)
def load_data():
    authed_credentials = credentials.with_scopes(scope)
    authed_credentials.refresh(Request())
    access_token = authed_credentials.token

    url = f"https://www.googleapis.com/drive/v3/files/{file_id}/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    headers = {"Authorization": f"Bearer {access_token}"}

    session = requests.Session()
    retries = Retry(
        total=5,
        connect=5,
        read=5,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
    )
    session.mount("https://", HTTPAdapter(max_retries=retries))

    response = session.get(
        url,
        headers=headers,
        timeout=60,
        verify=certifi.where(),
    )

    if response.status_code != 200:
        st.error(f"Drive API Error: {response.status_code} - {response.text}")
        st.stop()

    file_bytes = BytesIO(response.content)
    df = pd.read_excel(file_bytes, engine="openpyxl")

    df = df.dropna(how="all")
    df.columns = df.columns.str.strip()

    required_cols = [
        "SOLDTO",
        "LOADING_TS",
        "Backlog",
        "TARGET",
        "ORDERED_QUANTITY",
        "Order_in_New",
        "Order_in_Pool",
        "Incoterm",
        "City",
        "Type",
        "Region",
    ]

    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        st.error(f"Missing required columns: {', '.join(missing_cols)}")
        st.stop()

    df["SOLDTO"] = df["SOLDTO"].astype(str).str.strip()
    df = df[(df["SOLDTO"] != "") & (df["SOLDTO"].str.lower() != "nan")]

    numeric_cols = [
        "Backlog",
        "TARGET",
        "ORDERED_QUANTITY",
        "Order_in_New",
        "Order_in_Pool",
    ]

    for col in numeric_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["LOADING_TS"] = pd.to_datetime(df["LOADING_TS"], errors="coerce")
    df = df[df["LOADING_TS"].notna()].copy()
    df["LOADING_DATE"] = df["LOADING_TS"].dt.date

    return df


# -------------------------------------------------
# LOAD DATAFRAME
# -------------------------------------------------

try:
    df = load_data()
except Exception as e:
    st.error(f"Unable to load data: {e}")
    st.stop()

today_ts = pd.Timestamp.now(tz="Africa/Lagos").normalize().tz_localize(None)
today = today_ts.date()
month_start_ts = today_ts.replace(day=1)

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

    customer_df = df[df["SOLDTO"] == selected_customer].copy()

    if customer_df.empty:
        st.warning("No data available for the selected customer.")
        st.stop()

    filtered_df = customer_df[
        (customer_df["LOADING_DATE"] >= start_date)
        & (customer_df["LOADING_DATE"] <= end_date)
    ].copy()

    if filtered_df.empty:
        st.warning("No data available for the selected customer and date range.")
        st.stop()

    # -------------------------------------------------
    # OVERALL KPI VALUES
    # -------------------------------------------------

    total_backlog_value = customer_df["Backlog"].max()
    total_target_value = customer_df["TARGET"].max()
    total_order_new_value = customer_df["Order_in_New"].max()
    total_order_pool_value = customer_df["Order_in_Pool"].max()

    today_dispatch_value = customer_df[customer_df["LOADING_DATE"] == today][
        "ORDERED_QUANTITY"
    ].sum()

    mtd_dispatch_value = customer_df[
        (customer_df["LOADING_DATE"] >= month_start)
        & (customer_df["LOADING_DATE"] <= today)
    ]["ORDERED_QUANTITY"].sum()

    # -------------------------------------------------
    # SUMMARY FOR ALERTS
    # -------------------------------------------------

    summary = (
        filtered_df.groupby(
            ["SOLDTO", "Incoterm", "City", "Type", "Region"], dropna=False
        )
        .agg(
            Backlog=("Backlog", "max"),
            Target=("TARGET", "max"),
            Order_New=("Order_in_New", "max"),
            Order_Pool=("Order_in_Pool", "max"),
            Dispatch=(
                "ORDERED_QUANTITY",
                lambda x: x[
                    (filtered_df.loc[x.index, "LOADING_DATE"] >= month_start)
                    & (filtered_df.loc[x.index, "LOADING_DATE"] <= today)
                ].sum(),
            ),
        )
        .reset_index()
    )

    summary["Coverage"] = summary["Backlog"] / summary["Target"].replace(0, np.nan)

    # -------------------------------------------------
    # CARD BREAKDOWN
    # -------------------------------------------------

    card_base = filtered_df.groupby(
        ["City", "Type", "Incoterm"], dropna=False, as_index=False
    ).agg(
        Ordered_Qty=("ORDERED_QUANTITY", "sum"),
        Dispatch=(
            "ORDERED_QUANTITY",
            lambda x: x[
                (filtered_df.loc[x.index, "LOADING_DATE"] >= month_start)
                & (filtered_df.loc[x.index, "LOADING_DATE"] <= today)
            ].sum(),
        ),
    )

    total_ordered_qty = card_base["Ordered_Qty"].sum()

    if total_ordered_qty > 0:
        card_base["Share"] = card_base["Ordered_Qty"] / total_ordered_qty
    else:
        card_base["Share"] = 0

    card_base["Backlog"] = card_base["Share"] * total_backlog_value
    card_base["Order_New"] = card_base["Share"] * total_order_new_value
    card_base["Order_Pool"] = card_base["Share"] * total_order_pool_value
    card_base["Target"] = card_base["Share"] * total_target_value

    # -------------------------------------------------
    # QUANTITY BUCKETS INSIDE EACH CARD
    # -------------------------------------------------

    qty_buckets = [45, 40, 30, 20, 15, 10]

    bucket_source = filtered_df.groupby(
        ["City", "Type", "Incoterm", "ORDERED_QUANTITY"],
        dropna=False,
        as_index=False,
    ).agg(Bucket_Qty_Total=("ORDERED_QUANTITY", "sum"))

    bucket_source = bucket_source[bucket_source["ORDERED_QUANTITY"].isin(qty_buckets)]

    card_bucket_map = {}

    for _, card_row in card_base.iterrows():
        city = card_row["City"]
        typ = card_row["Type"]
        incoterm = card_row["Incoterm"]

        temp = bucket_source[
            (bucket_source["City"] == city)
            & (bucket_source["Type"] == typ)
            & (bucket_source["Incoterm"] == incoterm)
        ].copy()

        bucket_totals = {b: 0 for b in qty_buckets}

        for _, r in temp.iterrows():
            bucket_totals[int(r["ORDERED_QUANTITY"])] = r["Bucket_Qty_Total"]

        bucket_sum = sum(bucket_totals.values())

        backlog_bucket = {b: 0 for b in qty_buckets}
        dispatch_bucket = {b: 0 for b in qty_buckets}
        order_new_bucket = {b: 0 for b in qty_buckets}
        order_pool_bucket = {b: 0 for b in qty_buckets}

        for b in qty_buckets:
            share = bucket_totals[b] / bucket_sum if bucket_sum > 0 else 0
            backlog_bucket[b] = round(card_row["Backlog"] * share, 0)
            dispatch_bucket[b] = round(card_row["Dispatch"] * share, 0)
            order_new_bucket[b] = round(card_row["Order_New"] * share, 0)
            order_pool_bucket[b] = round(card_row["Order_Pool"] * share, 0)

        card_bucket_map[(city, typ, incoterm)] = {
            "Backlog": backlog_bucket,
            "Dispatch": dispatch_bucket,
            "Order_New": order_new_bucket,
            "Order_Pool": order_pool_bucket,
        }

    # -------------------------------------------------
    # ALERTS
    # -------------------------------------------------

    alerts = []
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    overall_backlog = total_backlog_value
    overall_target = total_target_value

    if overall_target == 0 and overall_backlog > 0:
        alerts.append(
            {
                "Time": timestamp,
                "Customer": selected_customer,
                "City": "All Cities",
                "Severity": "BLUE",
                "Message": "Backlog available but no target defined",
                "Reason": "Customer has backlog orders but no target assigned",
            }
        )

    if overall_backlog == 0:
        alerts.append(
            {
                "Time": timestamp,
                "Customer": selected_customer,
                "City": "All Cities",
                "Severity": "RED",
                "Message": "Backlog critically low",
                "Reason": "Backlog is currently at zero",
            }
        )
    elif overall_backlog < overall_target:
        alerts.append(
            {
                "Time": timestamp,
                "Customer": selected_customer,
                "City": "All Cities",
                "Severity": "RED",
                "Message": "Backlog not healthy",
                "Reason": "Backlog is less than the target available",
            }
        )
    else:
        alerts.append(
            {
                "Time": timestamp,
                "Customer": selected_customer,
                "City": "All Cities",
                "Severity": "GREEN",
                "Message": "Backlog healthy",
                "Reason": "Backlog is equal to or more than the target available",
            }
        )

    for _, row in summary.iterrows():
        if row["Order_New"] > 0:
            alerts.append(
                {
                    "Time": timestamp,
                    "Customer": row["SOLDTO"],
                    "City": row["City"],
                    "Severity": "YELLOW",
                    "Message": "Orders stuck in NEW",
                    "Reason": "",
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
                    "Reason": "",
                }
            )

    alerts_df = pd.DataFrame(alerts)

    # -------------------------------------------------
    # FILTERS
    # -------------------------------------------------

    st.markdown('<div class="filter-box">', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    city_options = ["All Cities"] + sorted(
        summary["City"].dropna().astype(str).unique()
    )
    type_options = ["All Types"] + sorted(summary["Type"].dropna().astype(str).unique())

    with col1:
        selected_city = st.selectbox("City", city_options)

    with col2:
        selected_type = st.selectbox("Type", type_options)

    with col3:
        metric = st.selectbox(
            "Metric", ["Backlog", "Dispatch", "Order_New", "Order_Pool"]
        )

    st.markdown("</div>", unsafe_allow_html=True)

    filtered_cards = card_base.copy()

    if selected_city != "All Cities":
        filtered_cards = filtered_cards[
            filtered_cards["City"].astype(str) == selected_city
        ]

    if selected_type != "All Types":
        filtered_cards = filtered_cards[
            filtered_cards["Type"].astype(str) == selected_type
        ]

    # -------------------------------------------------
    # LAYOUT
    # -------------------------------------------------

    left, right = st.columns([3, 1])

    with left:
        k1, k2, k3, k4, k5, k6 = st.columns(6)

        with k1:
            st.metric("Total Backlog", f"{total_backlog_value:,.0f}T")

        with k2:
            st.metric("Target", f"{total_target_value:,.0f}T")

        with k3:
            st.metric("MTD Dispatch", f"{mtd_dispatch_value:,.0f}T")

        with k4:
            st.metric("Today Dispatch", f"{today_dispatch_value:,.0f}T")

        with k5:
            st.metric("Order in New", f"{total_order_new_value:,.0f}T")

        with k6:
            st.metric("Order in Pool", f"{total_order_pool_value:,.0f}T")

        cols = st.columns(2)

        metric_label_map = {
            "Backlog": "Allocated Backlog",
            "Dispatch": "MTD Dispatch",
            "Order_New": "Allocated Order in New",
            "Order_Pool": "Allocated Order in Pool",
        }

        if filtered_cards.empty:
            st.info("No card data available for the selected filters.")
        else:
            for i, row in filtered_cards.iterrows():
                col = cols[i % 2]

                city_val = row["City"] if pd.notna(row["City"]) else "N/A"
                type_val = row["Type"] if pd.notna(row["Type"]) else "N/A"
                incoterm_val = row["Incoterm"] if pd.notna(row["Incoterm"]) else "N/A"

                if metric == "Backlog":
                    card_class = (
                        "card card-critical"
                        if (row["Backlog"] == 0 or row["Backlog"] < row["Target"])
                        else "card card-normal"
                    )
                else:
                    card_class = (
                        "card card-normal" if row[metric] > 0 else "card card-critical"
                    )

                bucket_values = card_bucket_map.get(
                    (row["City"], row["Type"], row["Incoterm"]),
                    {metric: {b: 0 for b in qty_buckets}},
                )[metric]

                header_html = "".join([f"<th>{b}T</th>" for b in qty_buckets])
                value_html = "".join(
                    [
                        (
                            f"<td>{bucket_values[b]:,.0f}</td>"
                            if bucket_values[b] != 0
                            else "<td>0</td>"
                        )
                        for b in qty_buckets
                    ]
                )

                card_html = f"""<div class="{card_class}">
<div style="font-weight:bold;">{city_val}</div>
<div style="margin-top:6px;">Type: {type_val}</div>
<div>Incoterm: {incoterm_val}</div>
<div style="margin-top:12px;font-size:16px;font-weight:bold;">
{metric_label_map[metric]}: {row[metric]:,.0f}T
</div>
<div style="margin-top:12px;">
<table class="real-table">
<tr>{header_html}</tr>
<tr>{value_html}</tr>
</table>
</div>
</div>"""

                with col:
                    st.markdown(card_html, unsafe_allow_html=True)

    with right:
        if not alerts_df.empty:
            st.subheader("🚨 Live Alerts")

            alerts_df = alerts_df.drop_duplicates()
            alerts_df = alerts_df.sort_values("Time", ascending=False)

            for _, alert in alerts_df.iterrows():
                reason_text = alert.get("Reason", "")
                message = f"""
Customer: {alert['Customer']}

City: {alert['City']}

{alert['Message']}

Reason:
{reason_text}
"""

                if alert["Severity"] == "RED":
                    st.error(message)
                elif alert["Severity"] == "YELLOW":
                    st.warning(message)
                elif alert["Severity"] == "GREEN":
                    st.success(message)
                else:
                    st.info(message)
