import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import openpyxl
import requests
import certifi
import time
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request
from io import BytesIO
from http.client import RemoteDisconnected
from requests.exceptions import ConnectionError, ChunkedEncodingError, SSLError

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

    retries = Retry(
        total=5,
        connect=5,
        read=5,
        backoff_factor=2,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
        raise_on_status=False,
    )

    last_error = None

    for attempt in range(3):
        session = requests.Session()
        session.mount("https://", HTTPAdapter(max_retries=retries))
        session.headers.update(headers)

        try:
            response = session.get(
                url,
                timeout=(30, 120),
                verify=certifi.where(),
                stream=False,
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
                "Status Summary",
            ]

            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                st.error(f"Missing required columns: {', '.join(missing_cols)}")
                st.stop()

            df["SOLDTO"] = df["SOLDTO"].astype(str).str.strip()
            df = df[(df["SOLDTO"] != "") & (df["SOLDTO"].str.lower() != "nan")]

            df["Status Summary"] = df["Status Summary"].astype(str).str.strip()

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
            df["LOADING_TS"] = df["LOADING_TS"].dt.normalize()
            df["LOADING_DATE"] = df["LOADING_TS"].dt.date

            return df

        except (
            RemoteDisconnected,
            ConnectionError,
            ChunkedEncodingError,
            SSLError,
        ) as e:
            last_error = e
            time.sleep(2 * (attempt + 1))

        except Exception as e:
            last_error = e
            time.sleep(2 * (attempt + 1))

        finally:
            session.close()

    raise Exception(f"Google Drive connection failed after retries: {last_error}")


def get_snapshot_value(series):
    valid = series.dropna()
    if valid.empty:
        return 0
    return valid.max()


def allocate_snapshot_to_buckets(total_value, base_bucket_totals):
    bucket_sum = sum(base_bucket_totals.values())

    if bucket_sum == 0:
        return {b: 0 for b in base_bucket_totals.keys()}

    allocated = {}
    running_total = 0
    bucket_keys = list(base_bucket_totals.keys())

    for i, b in enumerate(bucket_keys):
        if i < len(bucket_keys) - 1:
            share = base_bucket_totals[b] / bucket_sum
            allocated[b] = round(total_value * share, 0)
            running_total += allocated[b]
        else:
            allocated[b] = round(total_value - running_total, 0)

    return allocated


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

    total_backlog_value = get_snapshot_value(customer_df["Backlog"])
    total_target_value = get_snapshot_value(customer_df["TARGET"])
    total_order_new_value = get_snapshot_value(customer_df["Order_in_New"])
    total_order_pool_value = get_snapshot_value(customer_df["Order_in_Pool"])

    # -------------------------------------------------
    # DISPATCH DATA, ONLY STATUS SUMMARY = DISPATCHED
    # -------------------------------------------------

    dispatched_df = filtered_df[
        filtered_df["Status Summary"].astype(str).str.strip().str.upper()
        == "DISPATCHED"
    ].copy()

    if dispatched_df.empty:
        report_today_ts = filtered_df["LOADING_TS"].max()
        report_today = report_today_ts.date()
        report_month_start_ts = report_today_ts.replace(day=1)
        today_dispatch_value = 0
        mtd_dispatch_value = 0
        today_dispatched_df = dispatched_df.copy()
        mtd_dispatched_df = dispatched_df.copy()
    else:
        report_today_ts = dispatched_df["LOADING_TS"].max()
        report_today = report_today_ts.date()
        report_month_start_ts = report_today_ts.replace(day=1)

        today_dispatched_df = dispatched_df[
            dispatched_df["LOADING_DATE"] == report_today
        ].copy()

        mtd_dispatched_df = dispatched_df[
            (dispatched_df["LOADING_TS"] >= report_month_start_ts)
            & (dispatched_df["LOADING_TS"] <= report_today_ts)
        ].copy()

        today_dispatch_value = today_dispatched_df["ORDERED_QUANTITY"].sum()
        mtd_dispatch_value = mtd_dispatched_df["ORDERED_QUANTITY"].sum()

    summary = (
        filtered_df.groupby(
            ["SOLDTO", "Incoterm", "City", "Type", "Region"], dropna=False
        )
        .agg(
            Backlog=("Backlog", "max"),
            Target=("TARGET", "max"),
            Order_New=("Order_in_New", "max"),
            Order_Pool=("Order_in_Pool", "max"),
        )
        .reset_index()
    )

    dispatch_summary = (
        mtd_dispatched_df.groupby(
            ["SOLDTO", "Incoterm", "City", "Type", "Region"], dropna=False
        )["ORDERED_QUANTITY"]
        .sum()
        .reset_index(name="Dispatch")
    )

    summary = summary.merge(
        dispatch_summary,
        on=["SOLDTO", "Incoterm", "City", "Type", "Region"],
        how="left",
    )
    summary["Dispatch"] = summary["Dispatch"].fillna(0)
    summary["Coverage"] = summary["Backlog"] / summary["Target"].replace(0, np.nan)

    backlog_card_base = filtered_df.groupby(
        ["City", "Type", "Incoterm"], dropna=False, as_index=False
    ).agg(Backlog=("ORDERED_QUANTITY", "sum"))

    dispatch_card_base = mtd_dispatched_df.groupby(
        ["City", "Type", "Incoterm"], dropna=False, as_index=False
    ).agg(Dispatch=("ORDERED_QUANTITY", "sum"))

    card_base = backlog_card_base.merge(
        dispatch_card_base,
        on=["City", "Type", "Incoterm"],
        how="outer",
    ).fillna(0)

    total_actual_backlog_qty = card_base["Backlog"].sum()

    if total_actual_backlog_qty > 0:
        card_base["Share"] = card_base["Backlog"] / total_actual_backlog_qty
    else:
        card_base["Share"] = 0

    card_base["Target"] = card_base["Share"] * total_target_value
    card_base["Order_New"] = card_base["Share"] * total_order_new_value
    card_base["Order_Pool"] = card_base["Share"] * total_order_pool_value

    qty_buckets = [45, 40, 30, 20, 15, 10]
    card_bucket_map = {}

    for _, card_row in card_base.iterrows():
        city = card_row["City"]
        typ = card_row["Type"]
        incoterm = card_row["Incoterm"]

        card_scope = filtered_df[
            (filtered_df["City"] == city)
            & (filtered_df["Type"] == typ)
            & (filtered_df["Incoterm"] == incoterm)
        ].copy()

        dispatch_scope = mtd_dispatched_df[
            (mtd_dispatched_df["City"] == city)
            & (mtd_dispatched_df["Type"] == typ)
            & (mtd_dispatched_df["Incoterm"] == incoterm)
        ].copy()

        actual_backlog_buckets = {b: 0 for b in qty_buckets}
        actual_dispatch_buckets = {b: 0 for b in qty_buckets}

        backlog_bucket_df = (
            card_scope[card_scope["ORDERED_QUANTITY"].isin(qty_buckets)]
            .groupby("ORDERED_QUANTITY", dropna=False, as_index=False)
            .agg(Bucket_Qty_Total=("ORDERED_QUANTITY", "sum"))
        )

        dispatch_bucket_df = (
            dispatch_scope[dispatch_scope["ORDERED_QUANTITY"].isin(qty_buckets)]
            .groupby("ORDERED_QUANTITY", dropna=False, as_index=False)
            .agg(Bucket_Qty_Total=("ORDERED_QUANTITY", "sum"))
        )

        for _, r in backlog_bucket_df.iterrows():
            actual_backlog_buckets[int(r["ORDERED_QUANTITY"])] = r["Bucket_Qty_Total"]

        for _, r in dispatch_bucket_df.iterrows():
            actual_dispatch_buckets[int(r["ORDERED_QUANTITY"])] = r["Bucket_Qty_Total"]

        order_new_buckets = allocate_snapshot_to_buckets(
            card_row["Order_New"], actual_backlog_buckets
        )
        order_pool_buckets = allocate_snapshot_to_buckets(
            card_row["Order_Pool"], actual_backlog_buckets
        )

        card_bucket_map[(city, typ, incoterm)] = {
            "Backlog": actual_backlog_buckets,
            "Dispatch": actual_dispatch_buckets,
            "Order_New": order_new_buckets,
            "Order_Pool": order_pool_buckets,
        }

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
                "Reason": "Customer has backlog orders but no target",
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

    if metric == "Dispatch":
        filtered_cards = filtered_cards[filtered_cards["Dispatch"] > 0]

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
            "Backlog": "Backlog Quantity",
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
                        if (overall_backlog == 0 or overall_backlog < overall_target)
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
