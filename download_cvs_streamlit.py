import streamlit as st
import imaplib
import email
from email.header import decode_header
import os
import fitz  # PyMuPDF
import zipfile
import shutil  # To clean up the directory
from datetime import datetime, timedelta

# Define constants
DEFAULT_KEYWORDS = []
DEFAULT_REQUIRED_TERMS = ["experience", "education", "skills"]
MAX_PAGES = 3
attachments_folder = ''

def check_pdf_criteria(pdf_path, profession_keywords, required_terms, max_pages):
    """Check if the PDF meets all criteria: profession-related keywords, page count, and contains all required terms."""
    try:
        doc = fitz.open(pdf_path)
        text = ""
        num_pages = len(doc)
        for page in doc:
            text += page.get_text()
        doc.close()

        # If no keywords specified, skip the keyword check
        contains_keywords = any(keyword.lower() in text.lower() for keyword in profession_keywords) if profession_keywords else True
        is_page_count_valid = num_pages <= max_pages
        contains_all_required_terms = all(term.lower() in text.lower() for term in required_terms)
        
        return contains_keywords and is_page_count_valid and contains_all_required_terms
    except Exception as e:
        st.error(f"Error processing {pdf_path}: {e}")
        return False

def fetch_and_process_emails(username, password, profession_keywords, required_terms, email_limit=None, start_date=None, end_date=None):
    """Fetch emails, process attachments, and check PDF criteria."""
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(username, password)
    imap.select("inbox")

    # Create search criteria based on date range or default to "ALL"
    search_criteria = "ALL"
    if start_date and end_date:
        start_date_str = start_date.strftime("%d-%b-%Y")
        # Add one day to end_date for including emails on the end date
        end_date_plus_one_str = (end_date + timedelta(days=1)).strftime("%d-%b-%Y")
        search_criteria = f'(SINCE "{start_date_str}" BEFORE "{end_date_plus_one_str}")'

    # Search for emails
    status, messages = imap.search(None, search_criteria)
    
    # Handle cases where no messages are found or search failed
    if status != "OK" or not messages[0]:
        imap.close()
        imap.logout()
        return None, [], [], []

    messages = messages[0].split(b' ')
    messages.reverse()

    global attachments_folder
    attachments_folder = "email_attachments"
    if not os.path.exists(attachments_folder):
        os.makedirs(attachments_folder)

    saved_files = []
    deleted_files = []
    email_info = []

    processed_count = 0  # Counter for processed emails

    for mail in messages:
        if email_limit and processed_count >= email_limit:  # Check if the limit is reached
            break

        res, msg = imap.fetch(mail, "(RFC822)")
        
        # Handle cases where FETCH command fails
        if res != "OK":
            st.error(f"Failed to fetch email with ID: {mail}")
            continue
        
        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])
                subject, encoding = decode_header(msg["Subject"])[0]
                if isinstance(subject, bytes):
                    subject = subject.decode(encoding if encoding else "utf-8")
                date = msg.get("Date")
                if date:
                    date = email.utils.parsedate_to_datetime(date)

                email_info.append((subject, date))

                if msg.is_multipart():
                    for part in msg.walk():
                        if part.get_content_disposition() == "attachment":
                            filename = part.get_filename()
                            if filename and filename.lower().endswith(".pdf"):
                                filepath = os.path.join(attachments_folder, filename)
                                with open(filepath, "wb") as f:
                                    f.write(part.get_payload(decode=True))
                                
                                if check_pdf_criteria(filepath, profession_keywords, required_terms, MAX_PAGES):
                                    saved_files.append(filepath)
                                else:
                                    os.remove(filepath)
                                    deleted_files.append(filename)
        
        processed_count += 1  # Increment the counter after processing an email
    
    imap.close()
    imap.logout()

    # Create a zip file of saved PDFs only if there are files to save
    if saved_files:
        zip_filename = "cv_files.zip"
        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            for file in saved_files:
                zipf.write(file, os.path.basename(file))
        return zip_filename, saved_files, deleted_files, email_info
    else:
        return None, saved_files, deleted_files, email_info

# Streamlit UI
st.title("CV and Resume Filter")

username = st.text_input("Email Username")
password = st.text_input("Email Password", type="password")
keywords_input = st.text_area("Enter Keywords (comma-separated)", value="")
required_terms_input = st.text_area("Enter Required Terms for CV (comma-separated)", value=", ".join(DEFAULT_REQUIRED_TERMS))

filter_option = st.selectbox("Select Filtering Option", options=[None, "By Number of Emails", "By Date Range"], format_func=lambda x: "Select an option" if x is None else x)

if filter_option == "By Number of Emails":
    email_limit = st.number_input("Enter the number of emails to process", min_value=1, value=10, step=1)
    start_date, end_date = None, None
elif filter_option == "By Date Range":
    today = datetime.today()
    default_end_date = today
    default_start_date = today - timedelta(days=10)
    start_date = st.date_input("Start Date", value=default_start_date)
    end_date = st.date_input("End Date", value=default_end_date)
    email_limit = None
else:
    email_limit, start_date, end_date = None, None, None

zip_filename = ''
if st.button("Fetch and Process Emails"):
    profession_keywords = [keyword.strip() for keyword in keywords_input.split(",") if keyword.strip()]
    required_terms = [term.strip() for term in required_terms_input.split(",")]
    
    with st.spinner("Processing emails..."):
        zip_filename, saved_files, deleted_files, email_info = fetch_and_process_emails(username, password, profession_keywords, required_terms, email_limit=email_limit, start_date=start_date, end_date=end_date)
        
        if email_info:
            st.subheader("Email Information")
            for subject, date in email_info:
                st.write(f"Subject: {subject}")
                st.write(f"Date: {date}")
                st.write("")

        if zip_filename:  # Only if there are saved files
            st.success(f"Process completed!")
            
            # Provide a download link for the zip file
            with open(zip_filename, "rb") as f:
                st.download_button("Download CVs as Zip", f, file_name=zip_filename)

            # Cleanup
            os.remove(zip_filename)
            for file in saved_files:
                os.remove(file)
            if os.path.exists(attachments_folder):
                shutil.rmtree(attachments_folder)
        else:
            st.warning("No CVs matched the criteria.")
