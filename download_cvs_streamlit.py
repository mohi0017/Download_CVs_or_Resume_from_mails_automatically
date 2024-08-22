import streamlit as st
import imaplib
import email
from email.header import decode_header
import os
import fitz  # PyMuPDF
import zipfile

# Define constants
DEFAULT_KEYWORDS = []  # Empty by default
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
        
        contains_keywords = any(keyword.lower() in text.lower() for keyword in profession_keywords)
        is_page_count_valid = num_pages <= max_pages
        contains_all_required_terms = all(term.lower() in text.lower() for term in required_terms)
        
        return contains_keywords and is_page_count_valid and contains_all_required_terms
    except Exception as e:
        st.error(f"Error processing {pdf_path}: {e}")
        return False

def fetch_and_process_emails(username, password, profession_keywords, required_terms):
    """Fetch emails, process attachments, and check PDF criteria."""
    imap = imaplib.IMAP4_SSL("imap.gmail.com")
    imap.login(username, password)
    imap.select("inbox")

    status, messages = imap.search(None, "ALL")
    messages = messages[0].split(b' ')
    messages.reverse()

    global attachments_folder 
    attachments_folder = "email_attachments"
    if not os.path.exists(attachments_folder):
        os.makedirs(attachments_folder)

    saved_files = []
    deleted_files = []

    for mail in messages:
        res, msg = imap.fetch(mail, "(RFC822)")
        for response in msg:
            if isinstance(response, tuple):
                msg = email.message_from_bytes(response[1])
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
    
    imap.close()
    imap.logout()

    # Create a zip file of saved PDFs
    zip_filename = "cv_files.zip"
    with zipfile.ZipFile(zip_filename, 'w') as zipf:
        for file in saved_files:
            zipf.write(file, os.path.basename(file))
    
    return zip_filename, saved_files, deleted_files

# Streamlit UI
st.title("CV and Resume Filter")

username = st.text_input("Email Username")
password = st.text_input("Email Password", type="password")
keywords_input = st.text_area("Enter Keywords (comma-separated)", value="")
required_terms_input = st.text_area("Enter Required Terms for CV (comma-separated)", value=", ".join(DEFAULT_REQUIRED_TERMS))
zip_filename = ''
if st.button("Fetch and Process Emails"):
    profession_keywords = [keyword.strip() for keyword in keywords_input.split(",") if keyword.strip()]
    required_terms = [term.strip() for term in required_terms_input.split(",")]
    
    with st.spinner("Processing emails..."):
        zip_filename, saved_files, deleted_files = fetch_and_process_emails(username, password, profession_keywords, required_terms)
        
        st.success(f"Process completed!")
        
        # Provide a download link for the zip file
        with open(zip_filename, "rb") as f:
            st.download_button("Download CVs as Zip", f, file_name=zip_filename)

    # Optionally, clean up zip file and attachments folder
    os.remove(zip_filename)
    for file in saved_files:
        os.remove(file)
    if os.path.exists(attachments_folder):
        os.rmdir(attachments_folder)
