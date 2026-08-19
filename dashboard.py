"""
Outlook Email Data Extractor Dashboard
Extracts structured data (Date, Brand, Cartons, Units, Adhoc, Comments, PO)
from Outlook emails using a local Ollama LLM.
"""

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from outlook_extractor import extract_emails, export_to_csv, get_available_mailboxes
from llm_extractor import extract_batch, check_ollama_available, TARGET_FIELDS, DEFAULT_MODEL

# Page config
st.set_page_config(
    page_title="Email Data Extractor",
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="expanded",
)


@st.cache_data(ttl=300, show_spinner=False)
def load_emails(max_per_folder, start_date, end_date, selected_mailboxes, folder_scope):
    """Extract raw emails (with body + attachments) from Outlook."""
    mailboxes = selected_mailboxes if selected_mailboxes else None
    folders = list(folder_scope) if folder_scope else None
    df = extract_emails(
        max_per_folder=max_per_folder,
        start_date=start_date,
        end_date=end_date,
        mailboxes=mailboxes,
        folder_scope=folders,
        include_body=True,
    )
    return df


@st.cache_data(ttl=600, show_spinner=False)
def load_mailboxes():
    """Get available mailbox names."""
    return get_available_mailboxes()


def main():
    st.sidebar.title("📧 Email Data Extractor")
    st.sidebar.markdown("---")

    # ---- Ollama status ----
    ok, msg = check_ollama_available()
    if ok:
        st.sidebar.success(f"🤖 Ollama ready ({DEFAULT_MODEL})")
    else:
        st.sidebar.error(f"🤖 Ollama: {msg}")

    # ---- Mailboxes ----
    st.sidebar.subheader("Mailboxes")
    try:
        available_mailboxes = load_mailboxes()
    except Exception:
        available_mailboxes = []

    col_sel1, col_sel2 = st.sidebar.columns(2)
    with col_sel1:
        if st.button("Select All", use_container_width=True):
            st.session_state["selected_mailboxes"] = available_mailboxes
    with col_sel2:
        if st.button("Deselect All", use_container_width=True):
            st.session_state["selected_mailboxes"] = []

    if "selected_mailboxes" not in st.session_state:
        st.session_state["selected_mailboxes"] = available_mailboxes

    selected_mailboxes = st.sidebar.multiselect(
        "Select mailboxes to scan",
        available_mailboxes,
        default=st.session_state["selected_mailboxes"],
        key="mailbox_selector",
    )

    st.sidebar.markdown("---")
    st.sidebar.subheader("Settings")

    # Date range picker
    today = datetime.now().date()
    default_start = today - timedelta(days=30)
    date_range = st.sidebar.date_input(
        "Date range",
        value=(default_start, today),
        max_value=today,
        help="Select start and end date for email extraction"
    )
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start_date, end_date = date_range
    else:
        start_date = date_range[0] if isinstance(date_range, tuple) else date_range
        end_date = today

    max_per_folder = st.sidebar.slider("Max emails per folder", 10, 500, 100, step=10)

    folder_scope = st.sidebar.multiselect(
        "Folders to scan",
        ["Inbox", "Sent Items", "Deleted Items", "Junk Email", "Archive", "Drafts"],
        default=["Inbox"],
    )

    # Sender keyword filter (to narrow down which emails to process)
    sender_filter = st.sidebar.text_input(
        "Sender contains (optional)",
        help="Only process emails whose sender/email contains this text. Leave blank for all."
    )
    subject_filter = st.sidebar.text_input(
        "Subject contains (optional)",
        help="Only process emails whose subject contains this text. Leave blank for all."
    )

    st.sidebar.markdown("---")

    extract_clicked = st.sidebar.button("🚀 Extract Data", use_container_width=True, type="primary")

    if st.sidebar.button("🔄 Clear Cache", use_container_width=True):
        st.cache_data.clear()
        for key in ["extracted_data", "raw_emails"]:
            st.session_state.pop(key, None)
        st.rerun()

    # ===== MAIN =====
    st.title("📊 Email Data Extractor")
    st.caption("Extracts Date, Brand, Cartons, Units, Adhoc, Comments, PO from your Outlook emails using a local AI model.")

    if not ok:
        st.error(f"⚠️ Ollama is not available: {msg}")
        st.info("Make sure Ollama is installed and running, and the model is pulled (`ollama pull llama3.1`).")
        return

    if extract_clicked:
        if not selected_mailboxes:
            st.warning("Please select at least one mailbox.")
            return

        # Step 1: pull raw emails
        with st.spinner("Reading emails from Outlook..."):
            try:
                raw_df = load_emails(
                    max_per_folder, str(start_date), str(end_date),
                    tuple(selected_mailboxes), tuple(folder_scope)
                )
            except ConnectionError as e:
                st.error(f"❌ {e}")
                st.info("Make sure Microsoft Outlook is open and running.")
                return
            except Exception as e:
                st.error(f"❌ Error reading emails: {e}")
                return

        if raw_df.empty:
            st.warning("No emails found for the selected filters.")
            return

        # Apply optional sender/subject filters
        if sender_filter:
            mask = raw_df["Sender"].str.contains(sender_filter, case=False, na=False) | \
                   raw_df["SenderEmail"].str.contains(sender_filter, case=False, na=False)
            raw_df = raw_df[mask]
        if subject_filter:
            raw_df = raw_df[raw_df["Subject"].str.contains(subject_filter, case=False, na=False)]

        if raw_df.empty:
            st.warning("No emails matched your sender/subject filters.")
            return

        st.info(f"Found {len(raw_df)} emails. Analyzing with AI to extract structured data...")

        # Step 2: run LLM extraction with progress
        emails_list = raw_df.to_dict("records")
        progress_bar = st.progress(0.0)
        status = st.empty()

        def on_progress(current, total):
            progress_bar.progress(current / total)
            status.text(f"Analyzing email {current} of {total}...")

        results = extract_batch(emails_list, model=DEFAULT_MODEL, progress_callback=on_progress)
        progress_bar.empty()
        status.empty()

        result_df = pd.DataFrame(results)
        st.session_state["extracted_data"] = result_df

    # ===== Display results =====
    result_df = st.session_state.get("extracted_data")

    if result_df is None:
        st.info("👈 Configure your filters in the sidebar, then click **Extract Data** to begin.")
        return

    if result_df.empty:
        st.warning("No data extracted.")
        return

    st.success(f"✅ Extracted data from {len(result_df)} emails")

    # Show the extracted structured data - main focus
    st.subheader("Extracted Data")

    # Column order: extracted fields first, then metadata
    display_order = TARGET_FIELDS + ["Subject", "Sender", "Mailbox", "FolderName", "ReceivedTime"]
    display_order = [c for c in display_order if c in result_df.columns]
    display_df = result_df[display_order]

    st.dataframe(display_df, use_container_width=True, height=500)

    # Export
    csv = display_df.to_csv(index=False, encoding="utf-8-sig")
    st.download_button(
        "📥 Download as CSV",
        data=csv,
        file_name=f"email_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        mime="text/csv",
        use_container_width=False,
    )

    # Summary of fields found
    st.markdown("---")
    st.subheader("Extraction Summary")
    cols = st.columns(len(TARGET_FIELDS))
    for i, field in enumerate(TARGET_FIELDS):
        with cols[i]:
            non_empty = result_df[field].astype(str).str.strip().replace("", pd.NA).notna().sum()
            st.metric(field, f"{non_empty}/{len(result_df)}")


if __name__ == "__main__":
    main()
