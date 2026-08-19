"""
Outlook Email Extractor
Connects to the local Outlook desktop app via win32com and extracts email data.
"""

import win32com.client
import pythoncom
import pywintypes
import pandas as pd
from datetime import datetime, timedelta, timezone
import os


def _com_datetime_to_iso(dt_val):
    """
    Convert a pywintypes.datetime (COM) to an ISO format string.
    Returns None if conversion fails.
    """
    if dt_val is None:
        return None
    try:
        # pywintypes.datetime is a subclass of datetime.datetime
        # Convert to a plain datetime by extracting components
        return datetime(
            year=dt_val.year,
            month=dt_val.month,
            day=dt_val.day,
            hour=dt_val.hour,
            minute=dt_val.minute,
            second=dt_val.second,
            tzinfo=timezone.utc,
        ).isoformat()
    except Exception:
        try:
            # Fallback: just use string representation
            return str(dt_val)
        except Exception:
            return None


def get_outlook_app():
    """Connect to the running Outlook application."""
    try:
        pythoncom.CoInitialize()
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        return namespace
    except Exception as e:
        raise ConnectionError(
            f"Could not connect to Outlook. Make sure Outlook is running.\nError: {e}"
        )


def get_folder_emails(folder, max_emails=300, mailbox_name="", start_date=None, end_date=None):
    """Extract emails from a specific Outlook folder using date restriction for speed."""
    emails = []

    # Use Outlook's Restrict to filter at COM level (much faster than iterating all)
    try:
        restrict_parts = []
        if start_date:
            restrict_parts.append(f"[ReceivedTime] >= '{start_date}'")
        if end_date:
            # Add one day to end_date to include emails from that entire day
            end_dt = datetime.strptime(str(end_date), "%Y-%m-%d") + timedelta(days=1)
            restrict_parts.append(f"[ReceivedTime] < '{end_dt.strftime('%m/%d/%Y')}'")

        if restrict_parts:
            restrict_str = " AND ".join(restrict_parts)
            items = folder.Items.Restrict(restrict_str)
        else:
            items = folder.Items

        items.Sort("[ReceivedTime]", True)
    except Exception:
        # Fallback: no restriction
        items = folder.Items
        items.Sort("[ReceivedTime]", True)

    count = 0
    for item in items:
        if count >= max_emails:
            break
        try:
            # Only process mail items (skip meeting requests, etc.)
            if item.Class == 43:  # olMail
                # Safely get datetime values and convert immediately to ISO strings
                try:
                    received = _com_datetime_to_iso(item.ReceivedTime)
                except Exception:
                    received = None

                try:
                    sent = _com_datetime_to_iso(item.SentOn)
                except Exception:
                    sent = None

                email_data = {
                    "Subject": str(getattr(item, "Subject", "") or "(No Subject)"),
                    "Sender": str(getattr(item, "SenderName", "") or "Unknown"),
                    "SenderEmail": str(getattr(item, "SenderEmailAddress", "") or ""),
                    "ReceivedTime": received,
                    "SentOn": sent,
                    "To": str(getattr(item, "To", "") or ""),
                    "CC": str(getattr(item, "CC", "") or ""),
                    "Size": getattr(item, "Size", 0) or 0,
                    "HasAttachments": item.Attachments.Count > 0,
                    "AttachmentCount": item.Attachments.Count,
                    "IsRead": not bool(getattr(item, "UnRead", True)),
                    "Importance": getattr(item, "Importance", 1),
                    "Categories": str(getattr(item, "Categories", "") or ""),
                    "FolderName": folder.Name,
                    "Mailbox": mailbox_name,
                    "ConversationTopic": str(getattr(item, "ConversationTopic", "") or ""),
                }
                emails.append(email_data)
                count += 1
        except Exception:
            # Skip items that can't be read (encrypted, corrupted, etc.)
            continue

    return emails


def get_all_stores(namespace):
    """Get all mailbox stores (personal + shared) from Outlook."""
    stores = []
    try:
        for store in namespace.Stores:
            try:
                store_name = store.DisplayName
                root = store.GetRootFolder()
                stores.append({"name": store_name, "root": root})
            except Exception:
                continue
    except Exception:
        pass
    return stores


def get_inbox_folders(root_folder, mailbox_name, folder_scope=None, depth=0):
    """Get mail folders from a mailbox store root filtered by folder_scope."""
    folders = []
    target_names = set(folder_scope) if folder_scope else {"Inbox", "Sent Items"}

    try:
        for subfolder in root_folder.Folders:
            try:
                if subfolder.Name in target_names:
                    folders.append({"folder": subfolder, "mailbox": mailbox_name})
                # Also check one level deeper for shared mailboxes with different structures
                if depth < 1:
                    for sub in subfolder.Folders:
                        if sub.Name in target_names:
                            folders.append({"folder": sub, "mailbox": mailbox_name})
            except Exception:
                continue
    except Exception:
        pass

    return folders


def _collect_folders(folder, folder_list, depth=0):
    """Recursively collect subfolders."""
    if depth > 3:  # Limit recursion depth
        return
    try:
        folder_list.append(folder)
        for subfolder in folder.Folders:
            _collect_folders(subfolder, folder_list, depth + 1)
    except Exception:
        pass


def extract_emails(max_per_folder=300, start_date=None, end_date=None, mailboxes=None, folder_scope=None):
    """
    Main extraction function. Scans all mailbox stores including shared mailboxes.
    
    Args:
        max_per_folder: Maximum emails to extract per folder
        start_date: Start date string (YYYY-MM-DD). None = 30 days ago.
        end_date: End date string (YYYY-MM-DD). None = today.
        mailboxes: List of mailbox names to scan. None = all mailboxes.
        folder_scope: List of folder names to scan (e.g. ["Inbox", "Sent Items"]). None = Inbox + Sent.
    
    Returns:
        pandas DataFrame with all extracted email data
    """
    namespace = get_outlook_app()
    all_emails = []

    # Parse dates
    if start_date:
        start_dt = datetime.strptime(str(start_date), "%Y-%m-%d")
    else:
        start_dt = datetime.now() - timedelta(days=30)

    if end_date:
        end_dt = datetime.strptime(str(end_date), "%Y-%m-%d")
    else:
        end_dt = datetime.now()

    # Format for Outlook Restrict filter (MM/DD/YYYY)
    restrict_start = start_dt.strftime("%m/%d/%Y")
    restrict_end = (end_dt + timedelta(days=1)).strftime("%m/%d/%Y")

    # Get all stores (personal + shared mailboxes)
    stores = get_all_stores(namespace)

    for store in stores:
        store_name = store["name"]

        # Filter by selected mailboxes if specified
        if mailboxes and store_name not in mailboxes:
            continue

        root = store["root"]

        # Get folders from this store based on folder_scope
        folder_entries = get_inbox_folders(root, store_name, folder_scope)

        # If no standard folders found, try the root's immediate subfolders
        if not folder_entries:
            target = set(folder_scope) if folder_scope else {"Inbox", "Sent Items"}
            try:
                for subfolder in root.Folders:
                    if subfolder.Name in target:
                        folder_entries.append({"folder": subfolder, "mailbox": store_name})
            except Exception:
                continue

        for entry in folder_entries:
            try:
                emails = get_folder_emails(entry["folder"], max_per_folder, entry["mailbox"], restrict_start, restrict_end)
                all_emails.extend(emails)
            except Exception:
                continue

    if not all_emails:
        return pd.DataFrame()

    df = pd.DataFrame(all_emails)

    # Convert datetime columns from ISO strings to pandas datetime
    df["ReceivedTime"] = pd.to_datetime(df["ReceivedTime"], errors="coerce", utc=True)
    df["SentOn"] = pd.to_datetime(df["SentOn"], errors="coerce", utc=True)

    # Drop rows where ReceivedTime is null (can't do anything useful with them)
    df = df.dropna(subset=["ReceivedTime"])

    if df.empty:
        return pd.DataFrame()

    # Filter by date range (belt and suspenders — Restrict should have handled this)
    start_aware = pd.Timestamp(start_dt, tz="UTC")
    end_aware = pd.Timestamp(end_dt + timedelta(days=1), tz="UTC")
    df = df[(df["ReceivedTime"] >= start_aware) & (df["ReceivedTime"] < end_aware)]

    if df.empty:
        return pd.DataFrame()

    # Add derived columns
    df["Date"] = df["ReceivedTime"].dt.date
    df["Hour"] = df["ReceivedTime"].dt.hour
    df["DayOfWeek"] = df["ReceivedTime"].dt.day_name()
    df["Week"] = df["ReceivedTime"].dt.isocalendar().week.astype(int)
    df["Month"] = df["ReceivedTime"].dt.tz_localize(None).dt.to_period("M").astype(str)

    # Importance labels
    importance_map = {0: "Low", 1: "Normal", 2: "High"}
    df["ImportanceLabel"] = df["Importance"].map(importance_map)

    # Sort by received time
    df = df.sort_values("ReceivedTime", ascending=False).reset_index(drop=True)

    return df


def _scan_folder_by_name(folder, target_names, all_emails, max_per_folder, cutoff_date, depth=0):
    """Recursively scan folders matching target names."""
    if depth > 3:
        return
    try:
        if folder.Name in target_names:
            emails = get_folder_emails(folder, max_per_folder)
            all_emails.extend(emails)
        for subfolder in folder.Folders:
            _scan_folder_by_name(subfolder, target_names, all_emails, max_per_folder, cutoff_date, depth + 1)
    except Exception:
        pass


def get_available_mailboxes():
    """Return a list of all mailbox store names available in Outlook."""
    namespace = get_outlook_app()
    mailboxes = []
    try:
        for store in namespace.Stores:
            try:
                mailboxes.append(store.DisplayName)
            except Exception:
                continue
    except Exception:
        pass
    return mailboxes


def export_to_csv(df, filename="outlook_emails.csv"):
    """Export DataFrame to CSV."""
    filepath = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
    df.to_csv(filepath, index=False, encoding="utf-8-sig")
    return filepath


if __name__ == "__main__":
    print("Extracting emails from Outlook...")
    df = extract_emails(max_per_folder=200, start_date="2026-07-20", end_date="2026-08-19")
    print(f"Extracted {len(df)} emails")
    if not df.empty:
        path = export_to_csv(df)
        print(f"Saved to: {path}")
        print(df[["Subject", "Sender", "ReceivedTime", "FolderName"]].head(10))
