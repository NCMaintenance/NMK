import streamlit as st
import pypff
import openpyxl
from openpyxl.utils.exceptions import IllegalCharacterError
import hashlib
import tempfile
import os
import pandas as pd
from datetime import datetime

# --- Configuration ---
st.set_page_config(page_title="PST to Excel Converter", page_icon="📧", layout="wide")

# --- Helper Functions ---

def clean_string(text):
    """Removes characters that are illegal in Excel cells."""
    if not text:
        return ""
    # Filter out control characters except newlines/tabs
    return "".join(c for c in str(text) if c in ['\n', '\r', '\t'] or c >= ' ')

def format_date_time(timestamp):
    """Converts a timestamp to UK Date (DD/MM/YYYY) and Time (HH:MM:SS) strings."""
    if not timestamp:
        return None, None
    try:
        if isinstance(timestamp, datetime):
            return timestamp.strftime("%d/%m/%Y"), timestamp.strftime("%H:%M:%S")
        
        # Handle string inputs or other formats if pypff returns them
        dt_str = str(timestamp)
        # Attempt generic parsing if it's a string
        try:
            dt = datetime.strptime(dt_str, "%Y-%m-%d %H:%M:%S.%f")
        except ValueError:
            return str(timestamp), ""
            
        return dt.strftime("%d/%m/%Y"), dt.strftime("%H:%M:%S")
    except Exception:
        return str(timestamp), ""

def get_recipients(message):
    """Iterates through recipients and categorises them into To, CC, and BCC."""
    to_list = []
    cc_list = []
    bcc_list = []

    try:
        count = message.get_number_of_recipients()
        for i in range(count):
            recipient = message.get_recipient(i)
            name = recipient.get_name()
            email = recipient.get_email_address()
            
            if name and email:
                entry = f"{name} <{email}>"
            elif email:
                entry = email
            elif name:
                entry = name
            else:
                continue

            r_type = recipient.get_type()
            # 1: To, 2: CC, 3: BCC
            if r_type == 1:
                to_list.append(entry)
            elif r_type == 2:
                cc_list.append(entry)
            elif r_type == 3:
                bcc_list.append(entry)
                
    except Exception:
        pass

    return "; ".join(to_list), "; ".join(cc_list), "; ".join(bcc_list)

def generate_signature(subject, date_str, time_str, body):
    """Creates a unique hash for the email based on its content."""
    unique_string = f"{subject}|{date_str}|{time_str}|{body}"
    return hashlib.md5(unique_string.encode('utf-8', errors='ignore')).hexdigest()

def process_folder(folder, rows_list, current_path, seen_emails, progress_bar, status_text):
    """Recursively processes folders and appends data to a list."""
    
    # Get current folder name
    folder_name = folder.get_name()
    if not folder_name:
        folder_name = "Root"
    
    # Build path
    if current_path:
        full_path = f"{current_path} > {folder_name}"
    else:
        full_path = folder_name

    # Update UI status
    status_text.text(f"Scanning: {full_path}...")

    # Process messages
    for message in folder.sub_messages:
        try:
            subject = clean_string(message.get_subject())
            body = clean_string(message.get_plain_text_body())
            
            date_obj = message.get_delivery_time()
            if not date_obj:
                date_obj = message.get_client_submit_time()
            
            date_str, time_str = format_date_time(date_obj)
            to_str, cc_str, bcc_str = get_recipients(message)

            email_signature = generate_signature(subject, date_str, time_str, body)
            
            if email_signature in seen_emails:
                is_duplicate = "Yes"
            else:
                is_duplicate = "No"
                seen_emails.add(email_signature)

            rows_list.append([
                full_path,
                subject,
                date_str,
                time_str,
                clean_string(to_str),
                clean_string(cc_str),
                clean_string(bcc_str),
                is_duplicate,
                body
            ])

        except Exception:
            continue

    # Recurse into sub-folders
    for sub_folder in folder.sub_folders:
        process_folder(sub_folder, rows_list, full_path, seen_emails, progress_bar, status_text)

# --- Main App Interface ---

def main():
    st.title("📧 PST to Excel Converter")
    st.markdown("""
    Upload a Microsoft Outlook `.pst` file to extract all emails into an Excel spreadsheet.
    Includes automatic duplicate detection and UK date formatting.
    """)

    uploaded_file = st.file_uploader("Choose a PST file", type=["pst"])

    if uploaded_file is not None:
        st.info("File uploaded successfully. Processing may take a while depending on file size.")
        
        if st.button("Start Extraction"):
            # Create a temporary file to store the uploaded PST
            # pypff requires a file path, it cannot read from memory RAM directly
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pst") as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_file_path = tmp_file.name

            try:
                # Initialize tracking variables
                seen_emails = set()
                all_rows = []
                
                # UI Elements for feedback
                progress_bar = st.progress(0)
                status_text = st.empty()

                # Open PST
                pst = pypff.file()
                pst.open(tmp_file_path)
                root = pst.get_root_folder()
                
                status_text.text("File opened. Starting directory traversal...")

                # Process
                process_folder(root, all_rows, "", seen_emails, progress_bar, status_text)
                
                # Close PST
                pst.close()
                progress_bar.progress(100)
                status_text.text("Extraction Complete! Generating Excel file...")

                # Convert to DataFrame for easy Excel export
                headers = ['Folder Path', 'Subject', 'Date', 'Time', 'To', 'CC', 'BCC', 'Is Duplicate', 'Body']
                df = pd.DataFrame(all_rows, columns=headers)

                # Show preview
                st.subheader("Data Preview")
                st.dataframe(df.head(50))
                
                # Convert DF to Excel in memory
                # We use a buffer to avoid saving another temp file
                from io import BytesIO
                output = BytesIO()
                
                # Use ExcelWriter to handle formatting if needed, but standard to_excel is fine
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Emails')
                
                excel_data = output.getvalue()

                # Download Button
                st.success(f"Processed {len(all_rows)} emails.")
                st.download_button(
                    label="📥 Download Excel File",
                    data=excel_data,
                    file_name=f"extracted_emails_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"An error occurred during processing: {e}")
            
            finally:
                # Cleanup: Remove the temporary PST file from disk
                if os.path.exists(tmp_file_path):
                    try:
                        os.unlink(tmp_file_path)
                    except Exception:
                        pass

if __name__ == "__main__":
    main()
