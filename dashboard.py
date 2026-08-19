"""
Outlook Email Dashboard
A Streamlit dashboard that visualizes email data extracted from Outlook.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
from outlook_extractor import extract_emails, export_to_csv, get_available_mailboxes

# Page config
st.set_page_config(
    page_title="Outlook Email Dashboard",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom styling
st.markdown("""
<style>
    .metric-card {
        background: linear-gradient(135deg, #1e3a5f 0%, #2d5a87 100%);
        padding: 20px;
        border-radius: 10px;
        color: white;
        text-align: center;
    }
    .stMetric > div {
        background-color: #0e1117;
        padding: 15px;
        border-radius: 8px;
        border: 1px solid #262730;
    }
</style>
""", unsafe_allow_html=True)


@st.cache_data(ttl=300, show_spinner="Extracting emails from Outlook...")
def load_data(max_per_folder, days_back, selected_mailboxes):
    """Load and cache email data."""
    mailboxes = selected_mailboxes if selected_mailboxes else None
    df = extract_emails(max_per_folder=max_per_folder, days_back=days_back, mailboxes=mailboxes)
    return df


@st.cache_data(ttl=600, show_spinner="Discovering mailboxes...")
def load_mailboxes():
    """Get available mailbox names."""
    return get_available_mailboxes()


def main():
    # Sidebar controls
    st.sidebar.title("📧 Mail Dashboard")
    st.sidebar.markdown("---")

    st.sidebar.subheader("Mailboxes")
    try:
        available_mailboxes = load_mailboxes()
    except Exception:
        available_mailboxes = []

    selected_mailboxes = st.sidebar.multiselect(
        "Select mailboxes to scan",
        available_mailboxes,
        default=available_mailboxes,
        help="Includes your personal mailbox and all shared mailboxes"
    )

    st.sidebar.markdown("---")
    st.sidebar.subheader("Settings")
    days_back = st.sidebar.slider("Days to look back", 7, 365, 90, step=7)
    max_per_folder = st.sidebar.slider("Max emails per folder", 100, 2000, 500, step=100)

    if st.sidebar.button("🔄 Refresh Data", use_container_width=True):
        st.cache_data.clear()

    # Load data
    try:
        df = load_data(max_per_folder, days_back, tuple(selected_mailboxes))
    except ConnectionError as e:
        st.error(f"❌ {e}")
        st.info("Make sure Microsoft Outlook is open and running.")
        return
    except Exception as e:
        st.error(f"❌ Error loading emails: {e}")
        return

    if df.empty:
        st.warning("No emails found for the selected time period.")
        return

    # Sidebar filters
    st.sidebar.markdown("---")
    st.sidebar.subheader("Filters")

    # Mailbox filter (for viewing)
    if "Mailbox" in df.columns:
        mailbox_options = ["All"] + sorted(df["Mailbox"].unique().tolist())
        selected_mailbox_filter = st.sidebar.selectbox("Mailbox", mailbox_options)
    else:
        selected_mailbox_filter = "All"

    # Folder filter
    folders = ["All"] + sorted(df["FolderName"].unique().tolist())
    selected_folder = st.sidebar.selectbox("Folder", folders)

    # Sender filter
    top_senders = df["Sender"].value_counts().head(20).index.tolist()
    selected_senders = st.sidebar.multiselect("Filter by Sender", top_senders)

    # Apply filters
    filtered_df = df.copy()
    if selected_mailbox_filter != "All" and "Mailbox" in df.columns:
        filtered_df = filtered_df[filtered_df["Mailbox"] == selected_mailbox_filter]
    if selected_folder != "All":
        filtered_df = filtered_df[filtered_df["FolderName"] == selected_folder]
    if selected_senders:
        filtered_df = filtered_df[filtered_df["Sender"].isin(selected_senders)]

    # Export option
    st.sidebar.markdown("---")
    if st.sidebar.button("📥 Export to CSV", use_container_width=True):
        path = export_to_csv(filtered_df)
        st.sidebar.success(f"Saved to: {path}")

    # ===== MAIN CONTENT =====
    st.title("📊 Outlook Email Dashboard")
    st.caption(f"Showing data from the last {days_back} days • {len(filtered_df):,} emails")

    # ----- KPI Metrics Row -----
    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        st.metric("Total Emails", f"{len(filtered_df):,}")
    with col2:
        unread = filtered_df[~filtered_df["IsRead"]].shape[0]
        st.metric("Unread", f"{unread:,}")
    with col3:
        with_attachments = filtered_df[filtered_df["HasAttachments"]].shape[0]
        st.metric("With Attachments", f"{with_attachments:,}")
    with col4:
        high_importance = filtered_df[filtered_df["Importance"] == 2].shape[0]
        st.metric("High Importance", f"{high_importance:,}")
    with col5:
        unique_senders = filtered_df["Sender"].nunique()
        st.metric("Unique Senders", f"{unique_senders:,}")

    st.markdown("---")

    # ----- Row 1: Email Volume Over Time + By Folder -----
    col_left, col_right = st.columns([2, 1])

    with col_left:
        st.subheader("📈 Email Volume Over Time")
        daily_counts = filtered_df.groupby("Date").size().reset_index(name="Count")
        daily_counts["Date"] = pd.to_datetime(daily_counts["Date"])

        fig_volume = px.area(
            daily_counts,
            x="Date",
            y="Count",
            title="",
            color_discrete_sequence=["#4fc3f7"],
        )
        fig_volume.update_layout(
            xaxis_title="",
            yaxis_title="Emails",
            template="plotly_dark",
            height=350,
            margin=dict(l=20, r=20, t=20, b=20),
        )
        st.plotly_chart(fig_volume, use_container_width=True)

    with col_right:
        st.subheader("📁 By Folder")
        folder_counts = filtered_df["FolderName"].value_counts().reset_index()
        folder_counts.columns = ["Folder", "Count"]

        fig_folder = px.pie(
            folder_counts,
            values="Count",
            names="Folder",
            color_discrete_sequence=px.colors.qualitative.Set2,
        )
        fig_folder.update_layout(
            template="plotly_dark",
            height=350,
            margin=dict(l=20, r=20, t=20, b=20),
            showlegend=True,
        )
        st.plotly_chart(fig_folder, use_container_width=True)

    # ----- Row 2: Top Senders + Day of Week -----
    col_left2, col_right2 = st.columns(2)

    with col_left2:
        st.subheader("👤 Top 15 Senders")
        top_sender_df = (
            filtered_df["Sender"]
            .value_counts()
            .head(15)
            .reset_index()
        )
        top_sender_df.columns = ["Sender", "Count"]

        fig_senders = px.bar(
            top_sender_df,
            x="Count",
            y="Sender",
            orientation="h",
            color="Count",
            color_continuous_scale="Blues",
        )
        fig_senders.update_layout(
            template="plotly_dark",
            height=400,
            margin=dict(l=20, r=20, t=20, b=20),
            yaxis=dict(autorange="reversed"),
            showlegend=False,
            coloraxis_showscale=False,
        )
        st.plotly_chart(fig_senders, use_container_width=True)

    with col_right2:
        st.subheader("📅 Emails by Day of Week")
        day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        day_counts = filtered_df["DayOfWeek"].value_counts().reindex(day_order, fill_value=0).reset_index()
        day_counts.columns = ["Day", "Count"]

        fig_days = px.bar(
            day_counts,
            x="Day",
            y="Count",
            color="Count",
            color_continuous_scale="Oranges",
        )
        fig_days.update_layout(
            template="plotly_dark",
            height=400,
            margin=dict(l=20, r=20, t=20, b=20),
            showlegend=False,
            coloraxis_showscale=False,
            xaxis_title="",
            yaxis_title="Emails",
        )
        st.plotly_chart(fig_days, use_container_width=True)

    # ----- Row 3: Hourly Heatmap + Importance -----
    col_left3, col_right3 = st.columns([2, 1])

    with col_left3:
        st.subheader("🕐 Email Activity Heatmap (Hour × Day)")
        heatmap_data = filtered_df.groupby(["DayOfWeek", "Hour"]).size().reset_index(name="Count")
        heatmap_pivot = heatmap_data.pivot_table(
            index="DayOfWeek", columns="Hour", values="Count", fill_value=0
        )
        # Reorder days
        heatmap_pivot = heatmap_pivot.reindex(day_order)

        fig_heatmap = px.imshow(
            heatmap_pivot,
            labels=dict(x="Hour of Day", y="Day of Week", color="Emails"),
            color_continuous_scale="YlOrRd",
            aspect="auto",
        )
        fig_heatmap.update_layout(
            template="plotly_dark",
            height=350,
            margin=dict(l=20, r=20, t=20, b=20),
        )
        st.plotly_chart(fig_heatmap, use_container_width=True)

    with col_right3:
        st.subheader("⚡ By Importance")
        importance_counts = filtered_df["ImportanceLabel"].value_counts().reset_index()
        importance_counts.columns = ["Importance", "Count"]

        colors = {"Low": "#4fc3f7", "Normal": "#81c784", "High": "#ef5350"}
        fig_importance = px.bar(
            importance_counts,
            x="Importance",
            y="Count",
            color="Importance",
            color_discrete_map=colors,
        )
        fig_importance.update_layout(
            template="plotly_dark",
            height=350,
            margin=dict(l=20, r=20, t=20, b=20),
            showlegend=False,
            xaxis_title="",
            yaxis_title="Count",
        )
        st.plotly_chart(fig_importance, use_container_width=True)

    # ----- Row 4: Weekly trend + Read/Unread -----
    col_left4, col_right4 = st.columns([2, 1])

    with col_left4:
        st.subheader("📊 Weekly Email Trend")
        weekly = filtered_df.groupby("Month").size().reset_index(name="Count")

        fig_weekly = px.bar(
            weekly,
            x="Month",
            y="Count",
            color_discrete_sequence=["#7c4dff"],
        )
        fig_weekly.update_layout(
            template="plotly_dark",
            height=300,
            margin=dict(l=20, r=20, t=20, b=20),
            xaxis_title="",
            yaxis_title="Emails",
        )
        st.plotly_chart(fig_weekly, use_container_width=True)

    with col_right4:
        st.subheader("📬 Read vs Unread")
        read_counts = filtered_df["IsRead"].value_counts().reset_index()
        read_counts.columns = ["Status", "Count"]
        read_counts["Status"] = read_counts["Status"].map({True: "Read", False: "Unread"})

        fig_read = px.pie(
            read_counts,
            values="Count",
            names="Status",
            color="Status",
            color_discrete_map={"Read": "#81c784", "Unread": "#ef5350"},
            hole=0.4,
        )
        fig_read.update_layout(
            template="plotly_dark",
            height=300,
            margin=dict(l=20, r=20, t=20, b=20),
        )
        st.plotly_chart(fig_read, use_container_width=True)

    # ----- Row 5: Recent Emails Table -----
    st.markdown("---")
    st.subheader("📋 Recent Emails")

    display_cols = ["Subject", "Sender", "ReceivedTime", "Mailbox", "FolderName", "IsRead", "HasAttachments", "ImportanceLabel"]
    available_cols = [c for c in display_cols if c in filtered_df.columns]
    display_df = filtered_df[available_cols].head(50).copy()
    col_names = {"Subject": "Subject", "Sender": "Sender", "ReceivedTime": "Received", "Mailbox": "Mailbox", "FolderName": "Folder", "IsRead": "Read", "HasAttachments": "Attachments", "ImportanceLabel": "Importance"}
    display_df.columns = [col_names.get(c, c) for c in available_cols]

    # Format the dataframe
    st.dataframe(
        display_df,
        use_container_width=True,
        height=400,
        column_config={
            "Subject": st.column_config.TextColumn("Subject", width="large"),
            "Sender": st.column_config.TextColumn("Sender", width="medium"),
            "Received": st.column_config.DatetimeColumn("Received", format="MMM DD, YYYY HH:mm"),
            "Read": st.column_config.CheckboxColumn("Read"),
            "Attachments": st.column_config.CheckboxColumn("📎"),
            "Importance": st.column_config.TextColumn("Priority"),
        },
    )


if __name__ == "__main__":
    main()
