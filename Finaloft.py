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
import http.client
from requests.exceptions import ConnectionError, ChunkedEncodingError, SSLError
from streamlit_autorefresh import st_autorefresh
import base64
from pathlib import Path


# PAGE CONFIG


st.set_page_config(page_title="OFT Backlog & Dispatch Assistant", layout="wide")
st_autorefresh(interval=600000, key="data_refresh")


# GOOGLE AUTHENTICATION


scope = ["https://www.googleapis.com/auth/drive.readonly"]

credentials = Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=scope,
)

file_id = "11u-AeuFdRbRgl-l6Wk2JaCraSreBc9Cz5oP-KRG31G8"

st.markdown(
    """
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    """,
    unsafe_allow_html=True,
)


# STYLE


st.markdown(
    """
    <style>
    .main {
        background-color: #0e1a2b;
        color: white;
    }

    .block-container {
        max-width: 1400px;
        padding-top: 1rem;
        padding-left: 1rem;
        padding-right: 1rem;
        padding-bottom: 2rem;
    }

    .greeting-bar {
        background: #1a2f4a;
        padding: 14px 18px;
        border-radius: 12px;
        margin-bottom: 12px;
        font-size: 18px;
        font-weight: 500;
        line-height: 1.2;
        border: 1px solid #2e4a6b;
    }

    .header-bar {
        background: linear-gradient(90deg, #228B22, #1e7d1e);
        padding: 18px 20px;
        border-radius: 14px;
        margin-bottom: 18px;
        box-shadow: 0 4px 8px rgba(0,0,0,0.35);
        border: 1px solid #2e8b57;
    }

    .header-wrap {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 12px;
    }

    .header-title {
        color: white;
        font-weight: bold;
        font-size: clamp(18px, 2.2vw, 32px);
        text-align: center;
        flex: 1;
    }

    .header-logo {
        width: 70px;
        max-width: 100%;
        height: auto;
    }

    .kpi-box {
        background: #12263a;
        padding: 12px;
        border-radius: 12px;
        margin-bottom: 18px;
        border: 1px solid #223a57;
    }

    div[data-testid="stMetric"] {
        background: #14263d;
        border: 1px solid #223a57;
        padding: 10px;
        border-radius: 12px;
    }

    div[data-testid="stMetricLabel"] {
        font-size: 12px !important;
    }

    div[data-testid="stMetricValue"] {
        font-size: 22px !important;
        font-weight: bold;
    }

    div[data-testid="stMetricDelta"] {
        font-size: 12px !important;
    }

    .card {
        background: #14263d;
        padding: 18px;
        border-radius: 15px;
        margin-bottom: 18px;
        min-height: 240px;
        width: 100%;
        overflow-x: auto;
    }

    .card-normal {
        border: 3px solid #00D100;
    }

    .card-critical {
        border: 3px solid #ff0000;
    }

    .real-table {
        width: 100%;
        border-collapse: collapse;
        text-align: center;
        font-size: 12px;
        color: white;
    }

    .real-table th {
        border: 1px solid #3b4f6b;
        padding: 6px 4px;
        background: #1a2f4a;
    }

    .real-table td {
        border: 1px solid #3b4f6b;
        padding: 8px 4px;
        font-weight: bold;
    }

    .alert-box {
        background: #12263a;
        padding: 14px;
        border-radius: 12px;
        border: 1px solid #223a57;
        margin-bottom: 18px;
    }

    @media (max-width: 1024px) {
        .block-container {
            padding-left: 0.8rem;
            padding-right: 0.8rem;
        }

        .header-logo {
            width: 55px;
        }

        .greeting-bar {
            font-size: 16px;
        }
    }

    @media (max-width: 768px) {
        .header-wrap {
            flex-direction: column;
            text-align: center;
        }

        .header-logo {
            width: 50px;
        }

        .header-title {
            font-size: 20px;
        }

        .greeting-bar {
            font-size: 15px;
            padding: 12px;
        }

        div[data-testid="stMetricValue"] {
            font-size: 18px !important;
        }
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# GREETING


hour = datetime.now().hour

if hour < 12:
    greeting = "Good Morning,"
elif hour < 17:
    greeting = "Good Afternoon,"
else:
    greeting = "Good Evening,"


# HEADER


def get_base64_logo(path):
    with open(path, "rb") as f:
        data = f.read()
    return base64.b64encode(data).decode()


left_logo_path = Path("logo/lafarge.jpeg")

right_logo_path = Path("logo/Huaxin.jpeg")

if left_logo_path.exists() and right_logo_path.exists():
    left_logo = get_base64_logo(left_logo_path)
    right_logo = get_base64_logo(right_logo_path)

    st.markdown(
        f"""
        <div class="greeting-bar">
            {greeting}
        </div>

        <div class="header-bar">
            <div class="header-wrap">
                <img src="data:image/png;base64,{left_logo}" class="header-logo">
                <div class="header-title">🚚 OFT Backlog & Dispatch Assistant</div>
                <img src="data:image/png;base64,{right_logo}" class="header-logo">
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        f"""
        <div class="greeting-bar">
            {greeting}
        </div>

        <div class="header-bar">
            <div class="header-wrap">
                <div class="header-title">🚚 OFT Backlog & Dispatch Assistant</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# LOAD DATA


@st.cache_data(show_spinner=True, ttl=300)
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
                timeout=(10, 60),
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
                "SHIPPING POINTS",
                "DAILY DISPATCH",
            ]

            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                st.error(f"Missing required columns: {', '.join(missing_cols)}")
                st.stop()

            df["SOLDTO"] = df["SOLDTO"].astype(str).str.strip()
            df = df[(df["SOLDTO"] != "") & (df["SOLDTO"].str.lower() != "nan")]

            df["Status Summary"] = df["Status Summary"].astype(str).str.strip()
            df["SHIPPING POINTS"] = df["SHIPPING POINTS"].astype(str).str.strip()

            numeric_cols = [
                "Backlog",
                "TARGET",
                "ORDERED_QUANTITY",
                "Order_in_New",
                "Order_in_Pool",
                "DAILY DISPATCH",
            ]

            for col in numeric_cols:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

            df["LOADING_TS"] = pd.to_datetime(
                df["LOADING_TS"], format="mixed", errors="coerce"
            )
            df = df[df["LOADING_TS"].notna()].copy()
            df["LOADING_TS"] = df["LOADING_TS"].dt.normalize()
            df["LOADING_DATE"] = df["LOADING_TS"].dt.date

            return df

        except (
            http.client.RemoteDisconnected,
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

st.caption(f"🕒 Last data refresh: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

today_ts = pd.Timestamp.now(tz="Africa/Lagos").normalize().tz_localize(None)
today = today_ts.date()
today_data_exists = (df["LOADING_DATE"] == today).any()
month_start_ts = today_ts.replace(day=1)

# -------------------------------------------------
# SIDEBAR FILTERS
# -------------------------------------------------

st.sidebar.header("MTD Date Range")

date_series = pd.to_datetime(df["LOADING_DATE"], errors="coerce")

if date_series.notna().any():
    min_date = date_series.min().date()
    max_date = date_series.max().date()
else:
    st.error("No valid LOADING_DATE found in data.")
    st.stop()


df = df.dropna(subset=["LOADING_DATE"])

# --- FIX: keep default dates within available data range ---
safe_today = today

if safe_today < min_date:
    safe_today = min_date
elif safe_today > max_date:
    safe_today = max_date
safe_month_start = min(max(month_start_ts.date(), min_date), max_date)

start_date = st.sidebar.date_input(
    "Start Date", safe_month_start, min_value=min_date, max_value=max_date
)

end_date = st.sidebar.date_input(
    "End Date", safe_today, min_value=min_date, max_value=max_date
)

if start_date > end_date:
    st.sidebar.error("End Date must be after Start Date")
    st.stop()

customers = sorted(df["SOLDTO"].unique())

selected_customer = st.sidebar.selectbox(
    "Customer Search", ["Select Customer"] + customers
)

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

    filtered_df["SHIPPING POINTS"] = (
        filtered_df["SHIPPING POINTS"].astype(str).str.strip()
    )

    total_backlog_value = get_snapshot_value(customer_df["Backlog"])
    total_target_value = get_snapshot_value(customer_df["TARGET"])
    total_order_new_value = get_snapshot_value(customer_df["Order_in_New"])
    total_order_pool_value = get_snapshot_value(customer_df["Order_in_Pool"])

    # SHIPPING POINT ORDER IN POOL BREAKDOWN
    shipping_pool_base = (
        filtered_df.groupby("SHIPPING POINTS", dropna=False)["ORDERED_QUANTITY"]
        .sum()
        .reset_index(name="Backlog_Share")
    )

    total_backlog_qty_sp = shipping_pool_base["Backlog_Share"].sum()

    if total_backlog_qty_sp > 0:
        shipping_pool_base["Share"] = (
            shipping_pool_base["Backlog_Share"] / total_backlog_qty_sp
        )
    else:
        shipping_pool_base["Share"] = 0

    shipping_pool_base["Order_in_Pool"] = (
        shipping_pool_base["Share"] * total_order_pool_value
    )

    shipping_pool_df = shipping_pool_base[
        ["SHIPPING POINTS", "Order_in_Pool"]
    ].sort_values("Order_in_Pool", ascending=False)

    # DISPATCH DATA
    dispatched_df = filtered_df[
        filtered_df["Status Summary"].astype(str).str.strip().str.upper()
        == "DISPATCHED"
    ].copy()

    report_today_ts = today_ts
    report_today = today
    report_month_start_ts = month_start_ts

    today_dispatched_df = dispatched_df[
        dispatched_df["LOADING_DATE"] == report_today
    ].copy()

    mtd_dispatched_df = dispatched_df[
        (dispatched_df["LOADING_TS"] >= report_month_start_ts)
        & (dispatched_df["LOADING_TS"] <= report_today_ts)
    ].copy()

    today_dispatch_value = today_dispatched_df.drop_duplicates(
        subset=["LOADING_TS", "SOLDTO"]
    )["DAILY DISPATCH"].sum()

    today_dispatch_by_type = (
        today_dispatched_df.drop_duplicates(subset=["LOADING_TS", "SOLDTO", "Type"])
        .groupby("Type", dropna=False)["DAILY DISPATCH"]
        .sum()
        .reset_index()
    )

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
    ).agg(Base_Qty=("ORDERED_QUANTITY", "sum"))

    total_base_qty = backlog_card_base["Base_Qty"].sum()

    if total_base_qty > 0:
        backlog_card_base["Share"] = backlog_card_base["Base_Qty"] / total_base_qty
    else:
        backlog_card_base["Share"] = 0

    backlog_card_base["Backlog"] = backlog_card_base["Share"] * total_backlog_value

    # Optional: clean rounding so totals match exactly
    backlog_card_base["Backlog"] = backlog_card_base["Backlog"].round(0)

    difference = total_backlog_value - backlog_card_base["Backlog"].sum()
    if len(backlog_card_base) > 0:
        backlog_card_base.loc[backlog_card_base.index[-1], "Backlog"] += difference

    dispatch_card_base = mtd_dispatched_df.groupby(
        ["City", "Type", "Incoterm"], dropna=False, as_index=False
    ).agg(Dispatch=("ORDERED_QUANTITY", "sum"))

    card_base = backlog_card_base.merge(
        dispatch_card_base,
        on=["City", "Type", "Incoterm"],
        how="outer",
    ).fillna(0)

    order_pool_snapshot = (
        customer_df.groupby(["City", "Type", "Incoterm"], dropna=False)["Order_in_Pool"]
        .max()
        .reset_index()
    )

    card_base = card_base.merge(
        order_pool_snapshot, on=["City", "Type", "Incoterm"], how="left"
    )

    card_base["Order_in_Pool"] = card_base["Order_in_Pool"].fillna(0)

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
                "Message": "Backlog is available but no target",
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
                "Reason": "Backlog is more than the target available",
            }
        )

    for _, row in shipping_pool_df.iterrows():
        if row["Order_in_Pool"] > 0:
            alerts.append(
                {
                    "Time": timestamp,
                    "Customer": selected_customer,
                    "City": row["SHIPPING POINTS"],
                    "Severity": "BLUE",
                    "Message": f"Orders waiting in POOL: {row['Order_in_Pool']:,.0f}T",
                    "Reason": "Orders currently sitting in shipping point pool",
                }
            )

    alerts_df = pd.DataFrame(alerts)

    city_options = ["All Cities"] + sorted(
        summary["City"].dropna().astype(str).unique()
    )

    type_options = ["All Types"] + sorted(summary["Type"].dropna().astype(str).unique())

    shipping_point_options = ["All Shipping Points"] + sorted(
        filtered_df["SHIPPING POINTS"].dropna().astype(str).str.strip().unique()
    )

    # -------------------------------------------------
    # TOP FILTERS IN MAIN PAGE
    # -------------------------------------------------

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        selected_city = st.selectbox("City", city_options)

    with col2:
        selected_type = st.selectbox("Type", type_options)

    with col3:
        selected_shipping_point = st.selectbox("Shipping Point", shipping_point_options)

    with col4:
        metric = st.selectbox(
            "Metric", ["Backlog", "Dispatch", "Order_New", "Order_Pool"]
        )

    filtered_cards = card_base.copy()

    if selected_shipping_point == "All Shipping Points":
        selected_shipping_pool_value = shipping_pool_df["Order_in_Pool"].sum()
    else:
        selected_shipping_pool_value = shipping_pool_df[
            shipping_pool_df["SHIPPING POINTS"].astype(str).str.strip()
            == selected_shipping_point
        ]["Order_in_Pool"].sum()

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

    if selected_shipping_point != "All Shipping Points":
        alerts_df = alerts_df[
            alerts_df["City"].astype(str).str.strip() == selected_shipping_point
        ]
    st.markdown("### 🧠 Data Health Check")

    if df.empty:
        st.error("❌ No data loaded from source.")
    else:
        latest_data_date = df["LOADING_DATE"].max()

        if not today_data_exists:
            st.warning(
                f"⚠️ No data available for today ({today}). Latest data is {latest_data_date}"
            )
        else:
            st.success(f"✅ Data is up to date for today ({today})")

    latest_timestamp = df["LOADING_TS"].max()

    st.info(f"📡 Latest record timestamp: {latest_timestamp}")
    # -------------------------------------------------
    # KPI SECTION
    # -------------------------------------------------

    st.markdown('<div class="kpi-box">', unsafe_allow_html=True)

    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.metric("Total Backlog", f"{total_backlog_value:,.0f}T")
    with k2:
        st.metric("Target", f"{total_target_value:,.0f}T")
    with k3:
        st.metric("MTD Dispatch", f"{mtd_dispatch_value:,.0f}T")
    with k4:
        st.metric("Today Dispatch", f"{today_dispatch_value:,.0f}T")

    k5, k6, k7 = st.columns(3)
    with k5:
        st.metric("Order in New", f"{total_order_new_value:,.0f}T")
    with k6:
        st.metric("Order in Pool", f"{total_order_pool_value:,.0f}T")
    with k7:
        st.metric("Shipping Point Pool", f"{selected_shipping_pool_value:,.0f}T")

    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("### 📦 Today Dispatch by Order Type")

    if not today_dispatch_by_type.empty:
        for _, row in today_dispatch_by_type.iterrows():
            st.write(f"{row['Type']}: {row['DAILY DISPATCH']:,.0f}T")
    else:
        st.info("No dispatch recorded for today.")

    # -------------------------------------------------
    # MAIN CONTENT
    # -------------------------------------------------

    left, right = st.columns([5, 3])

    with left:
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

                card_html = f"""
                <div class="{card_class}">
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
                </div>
                """

                with col:
                    st.markdown(card_html, unsafe_allow_html=True)

    with right:
        st.markdown('<div class="alert-box">', unsafe_allow_html=True)
        st.subheader("🚨 Live Alerts")

        if not alerts_df.empty:
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
        else:
            st.info("No live alerts for the selected filters.")

        st.markdown("</div>", unsafe_allow_html=True)
