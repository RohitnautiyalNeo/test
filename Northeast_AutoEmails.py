import pandas as pd
import numpy as np
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import datetime
import tempfile
import os
import re


# === 1. READ FILES ===
raw_data = pd.read_csv(r"D:\SFTP_Root\files\DailyPendingCases_PHDash.csv",low_memory=False)
config = pd.read_excel("D:\SFTP_Root\Python files\ReportRequests.xlsx")
cases_in_flight_data = pd.read_csv(r"D:\SFTP_Root\files\InflightCases.csv",low_memory=False)

log_messages = []


def send_mail(receiver_email, subject, content):
    """
    Send an email via Outlook/Microsoft SMTP.
    
    Parameters:
    - receiver_email: str or list of recipient emails
    - subject: str, email subject
    - content: str, email body text
    """
    # --- Configure your email ---
    sender_email = "automated@insidepfa.com"
    password = "EzQhK59fg@?S+SLB"
    smtp_server = "smtp.office365.com"
    smtp_port = 587
    bcc_list = ["rohit.nautiyal@insidepfa.com"]

    # Create email message
    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = receiver_email if isinstance(receiver_email, str) else ", ".join(receiver_email)
    msg["Subject"] = subject
    msg["Bcc"] = ", ".join(bcc_list) 

    # Default email body
    body_text = "Hi,\n\nPlease find attached the requested report.\n\nThankyou\n"
    msg.attach(MIMEText(body_text, "plain"))

    # If content is a DataFrame → attach as Excel file
    temp_path = None
    safe_name = re.sub(r'[\\/*?:"<>|]', "", subject)
    if isinstance(content, pd.DataFrame):
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            temp_path = tmp.name
            content.to_excel(temp_path, index=False)
        
        with open(temp_path, "rb") as f:
            part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition","attachment",filename=f"{safe_name}.xlsx")
            msg.attach(part)
    else:
        # if just text
        msg.attach(MIMEText(str(content), "plain"))

    # Send email
    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            server.login(sender_email, password)
            server.send_message(msg)
        print(f" Email sent to {receiver_email} successfully!")
    except Exception as e:
        print(f" Error sending email: {e}")
    finally:
        if temp_path and os.path.exists(temp_path):
            os.remove(temp_path)
        
def parse_charge_range(text):
    """Return (low, high) numeric range from text like '$300,000 - $749,999.99'."""
    if pd.isna(text):
        return (0, np.inf)
    text = text.replace(',', '')
    numbers = [float(x.replace('$', '')) for x in re.findall(r'\$?([\d,]+(?:\.\d+)?)', text)]
    
    if "greater" in text.lower():
        return numbers[0], np.inf
    elif "-" in text:
        return numbers[0], numbers[1]
    elif "<" in text:
        return 0, numbers[0]
    else:
        return 0, np.inf

def is_today_in_frequency(freq_string):
    """
    Simplest version:
    - Splits by '-'
    - Matches today's weekday against the list (M, T, W, Th, F, S, Su)
    """
    day_map = {"M": 0, "T": 1, "W": 2, "Th": 3, "F": 4, "S": 5, "Su": 6}
    today_num = datetime.datetime.today().weekday()
    
    # Split by '-' and normalize capitalization
    parts = [p.strip().capitalize() for p in freq_string.split('-')]
    
    # Map each part to weekday number if valid
    weekdays = [day_map[p] for p in parts if p in day_map]
    
    return today_num in weekdays

def parse_days_condition(text):
    """Extract numeric threshold from 'Greater than 3 days'."""
    if pd.isna(text):
        return 0
    match = re.search(r'(\d+)', text)
    if match:
        return int(match.group(1))
    return 0

def get_filtered_data(raw_data, client, enable_status, assigned_to, charge_balance, days_since_last_note, freq_report, report_name):
    """
    Filter raw_data based on config parameters.
    Returns filtered DataFrame.
    Supports multiple comma-separated enable statuses.
    """
    # Convert enable_status string into a list of lowercase trimmed values
    enable_status_list = [s.strip().lower() for s in enable_status.split(",")]

    # --- Clean relevant columns ---
    raw_data["Current Status"] = raw_data["Current Status"].fillna("").str.strip().str.lower()
    raw_data["Assigned To"] = raw_data["Assigned To"].fillna("").str.strip()
    raw_data["Hospital"] = raw_data["Hospital"].fillna("").str.strip()
    raw_data["Days_since_last_Touched"] = pd.to_numeric(raw_data["Days_since_last_Touched"], errors="coerce")
    raw_data["Total Charges"] = pd.to_numeric(raw_data["Total Charges"], errors="coerce")

    # Skip if enable_status_list is empty
    if not enable_status_list:
        message=f"Skipping {report_name} — Enable Status not set\n"
        print(f"Skipping {report_name} — Enable Status not set")
        log_messages.append(message)
        return pd.DataFrame()

    # Parse other filters
    low, high = parse_charge_range(charge_balance)
    days_threshold = parse_days_condition(days_since_last_note)

    # --- Status filter condition ---
    if "all" in enable_status_list:
        status_filter = True
    else:
        status_filter = raw_data["Current Status"].isin(enable_status_list)

    # --- Days_since_last_Touched filter condition ---
    if str(days_since_last_note).strip().upper() == "ALL":
        Days_since_last_Touched_filter = True
    else:
        Days_since_last_Touched_filter = (raw_data["Days_since_last_Touched"] > days_threshold)

    # --- Apply filters ---
    filtered = raw_data[
        status_filter &
        (raw_data["Assigned To"].str.contains(assigned_to.strip(), case=False, na=False)) &
        (raw_data["Hospital"].str.contains(client.strip(), case=False, na=False)) &
        (raw_data["Total Charges"].between(low, high)) &
        Days_since_last_Touched_filter
    ]
    message=f"{report_name}: {len(filtered)} records found.\n"
    log_messages.append(message)
    print(f"{report_name}: {len(filtered)} records found.")
    return filtered

def get_filtered_data_cases_in_flight(cases_data, client, enable_status, assigned_to, charge_balance, days_since_last_note, freq_report, report_name,rvp):
    """
    Filter Cases_in_flight data based on config parameters.

    Parameters:
    - cases_data: DataFrame (Cases_in_flight dataset)
    - client: str, hospital or client name filter
    - assigned_to: str, assigned staff name filter
    - charge_balance: str, charge range filter
    - days_since_last_note: str, e.g. 'Greater than 3 days'
    - freq_report: str, frequency rule (ignored in filtering)
    - report_name: str, used for logging
    
    Returns:
    - Filtered DataFrame
    """
    # Convert enable_status string into a list of lowercase trimmed values
    enable_status_list = [s.strip().lower() for s in enable_status.split(",")]
    
    # --- Clean relevant columns ---
    cases_data["Current Status"] = cases_data["Current Status"].fillna("").str.strip().str.lower()
    cases_data["Assigned To"] = cases_data["Assigned To"].fillna("").str.strip()
    cases_data["Hospital"] = cases_data["Hospital"].fillna("").str.strip()
    cases_data["Days_since_last_Touched"] = pd.to_numeric(cases_data["Days_since_last_Touched"], errors="coerce")
    cases_data["Total Charges"] = pd.to_numeric(cases_data["Total Charges"], errors="coerce")

     # --------- SPECIAL CASE FOR RVP MARSHALL ---------
    if str(rvp).strip().lower() == "marshall":
        assignee_filter = cases_data["CaseOwner"].str.contains(assigned_to.strip(), case=False, na=False)
        print(f"RVP = Marshall → using CaseOwner filter")
        message=f"RVP = Marshall → using CaseOwner filter\n"
        log_messages.append(message)
    else:
        assignee_filter = cases_data["Assigned To"].str.contains(assigned_to.strip(), case=False, na=False)


    # --- Parse filters ---
    low, high = parse_charge_range(charge_balance)
    days_threshold = parse_days_condition(days_since_last_note)

    # --- Status filter condition ---
    if "all" in enable_status_list:
        status_filter = True
    else:
        status_filter = cases_data["Current Status"].isin(enable_status_list)

    # --- Days_since_last_Touched filter condition ---
    if str(days_since_last_note).strip().upper() == "ALL":
        Days_since_last_Touched_filter = True
    else:
        Days_since_last_Touched_filter = (cases_data["Days_since_last_Touched"] > days_threshold)

     # --- Hospital filter condition ---
    if str(client).strip().upper() == "ALL":
        client_filter = True
    else:
        client_filter = (cases_data["Hospital"].str.contains(client.strip(), case=False, na=False))

    # --- Apply filters ---
    filtered = cases_data[
        status_filter &
        assignee_filter &
        client_filter  &
        (cases_data["Total Charges"].between(low, high)) &
         Days_since_last_Touched_filter
    ]
    # --- Sort by Total Charges descending ---
    filtered = filtered.sort_values(by="Total Charges", ascending=False)

    print(f"{report_name}: {len(filtered)} records found in Cases_in_flight.")
    message=f"{report_name}: {len(filtered)} records found in Cases_in_flight.\n"
    log_messages.append(message)
    return filtered


for idx, row in config.iterrows():
    client = row['Client']
    enable_status = row['Enable Status']
    assigned_to = row['Case assigned to']
    charge_balance = row['Charge balance']
    days_since_last_note = row['Days since last note']
    freq_report = row['Frequecy of report']
    report_name = row['Report name']
    Email_distribution = row['Email distribution']
    rvp=row['RVP']

    # --- Choose dataset ---
    if re.search(r'cases[_\s]*in[_\s]*flight', report_name, re.IGNORECASE):
        print(f"Using Cases_in_flight data for report: {report_name}")
        message=f"Using Cases_in_flight data for report: {report_name}\n"
        log_messages.append(message)
        if not is_today_in_frequency(freq_report):
            print(f"Skipping {report_name} — not scheduled for today ({freq_report})")
            message=f"Skipping {report_name} — not scheduled for today ({freq_report})\n"
            log_messages.append(message)
            continue

        filtered = get_filtered_data_cases_in_flight(
            cases_data=cases_in_flight_data,
            client=client,
            enable_status=enable_status,
            assigned_to=assigned_to,
            charge_balance=charge_balance,
            days_since_last_note=days_since_last_note,
            freq_report=freq_report,
            report_name=report_name,
            rvp=rvp
        )

    else:
        print(f"Using DailyPendingCases data for report: {report_name}")
        message=f"Using DailyPendingCases data for report: {report_name}\n"
        log_messages.append(message)
        if not is_today_in_frequency(freq_report):
            print(f"Skipping {report_name} — not scheduled for today ({freq_report})")
            message=f"Skipping {report_name} — not scheduled for today ({freq_report})\n"
            log_messages.append(message)
            continue

        filtered = get_filtered_data(
            raw_data=raw_data,
            client=client,
            enable_status=enable_status,
            assigned_to=assigned_to,
            charge_balance=charge_balance,
            days_since_last_note=days_since_last_note,
            freq_report=freq_report,
            report_name=report_name
        )

    # --- Output ---
    if not filtered.empty:
        print(f"\n=== {report_name} ===")
        print(filtered)
        Email_distribution = Email_distribution.strip()
        send_mail(Email_distribution, report_name, filtered)
        message=f"\n=== {report_name} ===\n"
        log_messages.append(message)
        print("--------------------------------------------------------------------------")
    else:
        print(f"No records found for {report_name}")
        message=f"No records found for {report_name}\n"
        log_messages.append(message)
        print("--------------------------------------------------------------------------")
    log_messages.append("--------------------------------------------------------------------------")

log_body = "\n".join(log_messages)
send_mail('prone@insidepfa.com', "Northeast AutoEmail ran successfully",log_body)
send_mail('rohit.nautiyal@insidepfa.com', "Northeast AutoEmail ran successfully",log_body)


