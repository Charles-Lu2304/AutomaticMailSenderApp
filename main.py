import streamlit as st
import gspread
from google.oauth2.service_account import Credentials as ServiceCredentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
from email.mime.text import MIMEText
import base64
import json
import time
import smtplib
from email.mime.multipart import MIMEMultipart
import pandas as pd
import io
import re

# Page configuration
st.set_page_config(page_title="Gmail Auto-Sender", page_icon="📧", layout="wide")

# Initialize session state
if 'available_columns' not in st.session_state:
    st.session_state.available_columns = ["Email"]
if 'loaded_data' not in st.session_state:
    st.session_state.loaded_data = None
if 'email_column' not in st.session_state:
    st.session_state.email_column = 'email'
if 'data_source_key' not in st.session_state:
    st.session_state.data_source_key = None

st.title("📧 Automatic Email Sender")
st.markdown("---")

# Sidebar: Authentication settings
with st.sidebar:
    st.header("🔐 Data Source Settings")
    
    # Data source selection
    data_source = st.radio(
        "Data Source",
        ["Google Sheets", "Excel File (Local Upload)", "Google Drive Excel"],
        help="Choose the source of your recipient list"
    )
    
    st.markdown("---")
    
    # Google Sheets section
    if data_source == "Google Sheets":
        st.subheader("1. Spreadsheet Authentication")
        st.caption("Use Service Account")
        sheets_credentials_json = st.text_area(
            "Service Account JSON",
            height=150,
            help="Paste the JSON key from your service account for reading spreadsheets",
            key="sheets_creds"
        )
        
        # Display service account email
        if sheets_credentials_json:
            try:
                creds_dict = json.loads(sheets_credentials_json)
                service_email = creds_dict.get('client_email', '')
                if service_email:
                    st.success(f"✅ Service Account: `{service_email}`")
                    st.info("💡 Share the spreadsheet with this email address")
            except:
                pass
        
        st.markdown("---")
        st.subheader("Spreadsheet Settings")
        spreadsheet_url = st.text_input(
            "Spreadsheet URL",
            help="Enter the URL of the Google Spreadsheet containing the recipient list"
        )
        sheet_name = st.text_input("Sheet Name", value="Sheet1", key="sheets_sheet_name")
    elif data_source == "Excel File (Local Upload)":
        # Excel File section
        st.subheader("Excel File Settings")
        excel_file = st.file_uploader(
            "Upload Excel File",
            type=['xlsx', 'xls'],
            help="Upload an Excel file (.xlsx or .xls) containing the recipient list",
            key="excel_file"
        )
        sheet_name = st.text_input("Sheet Name", value="Sheet1", key="excel_sheet_name")
        
        # Initialize variables for Google Sheets (to avoid errors)
        sheets_credentials_json = ""
        spreadsheet_url = ""
    else:
        # Google Drive Excel section
        st.subheader("1. Google Drive Authentication")
        st.caption("Use Service Account")
        sheets_credentials_json = st.text_area(
            "Service Account JSON",
            height=150,
            help="Paste the JSON key from your service account for accessing Google Drive",
            key="drive_creds"
        )
        
        # Display service account email
        if sheets_credentials_json:
            try:
                creds_dict = json.loads(sheets_credentials_json)
                service_email = creds_dict.get('client_email', '')
                if service_email:
                    st.success(f"✅ Service Account: `{service_email}`")
                    st.info("💡 Share the Google Drive file with this email address")
            except:
                pass
        
        st.markdown("---")
        st.subheader("Google Drive File Settings")
        drive_file_url = st.text_input(
            "Google Drive File URL",
            help="Enter the URL of the Excel file stored in Google Drive",
            placeholder="https://drive.google.com/file/d/FILE_ID/view"
        )
        sheet_name = st.text_input("Sheet Name", value="Sheet1", key="drive_sheet_name")
        
        # Initialize variables for other sources
        spreadsheet_url = ""
        excel_file = None

# Main area
col1, col2 = st.columns([1, 1])

with col1:
    st.header("📋 Email Template")
    
    template_source = st.radio(
        "Template Source",
        ["Manual Input", "Upload Files"],
        horizontal=True
    )
    
    if template_source == "Manual Input":
        subject_template = st.text_input(
            "Subject Template",
            value="Hello {{name}}",
            help="Use {{variable_name}} format for placeholders"
        )
        
        body_template = st.text_area(
            "Body Template",
            value="""Dear {{name}},

Thank you for your interest.

{{message}}

Best regards,""",
            height=300,
            help="Use {{variable_name}} format for placeholders"
        )
    else:
        st.info("📁 Upload text files containing your email templates. Use {{variable_name}} format for placeholders.")
        
        subject_file = st.file_uploader(
            "Upload Subject Template File",
            type=['txt'],
            help="Upload a .txt file containing the subject template",
            key="subject_file"
        )
        
        body_file = st.file_uploader(
            "Upload Body Template File",
            type=['txt'],
            help="Upload a .txt file containing the body template",
            key="body_file"
        )
        
        subject_template = ""
        body_template = ""
        
        if subject_file is not None:
            subject_template = subject_file.read().decode('utf-8')
            st.success(f"✅ Subject template loaded ({len(subject_template)} characters)")
            with st.expander("Preview Subject Template"):
                st.text(subject_template)
        else:
            st.warning("⚠️ Please upload a subject template file")
        
        if body_file is not None:
            body_template = body_file.read().decode('utf-8')
            st.success(f"✅ Body template loaded ({len(body_template)} characters)")
            with st.expander("Preview Body Template"):
                st.text(body_template)
        else:
            st.warning("⚠️ Please upload a body template file")

with col2:
    st.header("⚙️ Sending Settings")
    sender_email = st.text_input(
        "Sender Email Address",
        help="Gmail address that will be used to send emails"
    )
    
    test_mode = st.checkbox("Test Mode (Don't actually send)", value=True)
    delay_seconds = st.slider("Delay between emails (seconds)", min_value=1, max_value=10, value=2)
    
    st.info("💡 Your spreadsheet should contain the following columns:\n- email: Recipient email address\n- name: Recipient name\n- Other variables used in templates")

# Extract file ID from Google Drive URL
def extract_file_id_from_url(url):
    """Extract file ID from various Google Drive URL formats"""
    if not url:
        return None
    
    # Pattern 1: /file/d/{FILE_ID}/
    pattern1 = r'/file/d/([a-zA-Z0-9_-]+)'
    match = re.search(pattern1, url)
    if match:
        return match.group(1)
    
    # Pattern 2: ?id={FILE_ID}
    pattern2 = r'[?&]id=([a-zA-Z0-9_-]+)'
    match = re.search(pattern2, url)
    if match:
        return match.group(1)
    
    # Pattern 3: /d/{FILE_ID}/
    pattern3 = r'/d/([a-zA-Z0-9_-]+)'
    match = re.search(pattern3, url)
    if match:
        return match.group(1)
    
    return None

# Load Google Drive Excel data
def load_google_drive_excel(creds_json, file_url, sheet_name):
    """Load Excel data from Google Drive"""
    try:
        # Validate JSON
        creds_dict = json.loads(creds_json)
        
        # Extract file ID from URL
        file_id = extract_file_id_from_url(file_url)
        if not file_id:
            st.error("❌ Invalid Google Drive URL")
            st.info("💡 Please use a valid Google Drive file URL (e.g., https://drive.google.com/file/d/FILE_ID/view)")
            return None
        
        # Required scopes
        scopes = ['https://www.googleapis.com/auth/drive.readonly']
        
        # Create credentials
        creds = ServiceCredentials.from_service_account_info(creds_dict, scopes=scopes)
        
        # Build Drive service
        service = build('drive', 'v3', credentials=creds)
        
        # Download file
        request = service.files().get_media(fileId=file_id)
        file_content = io.BytesIO()
        downloader = MediaIoBaseDownload(file_content, request)
        
        done = False
        while not done:
            status, done = downloader.next_chunk()
        
        # Reset file pointer to beginning
        file_content.seek(0)
        
        # Read Excel file with pandas
        df = pd.read_excel(file_content, sheet_name=sheet_name, engine='openpyxl')
        
        # Check if dataframe is empty
        if df.empty:
            st.error("❌ Excel file is empty")
            st.info("💡 Please ensure the Excel file contains data")
            return None
        
        # Convert to list of dictionaries
        data = df.to_dict('records')
        
        # Convert NaN values to empty strings
        for record in data:
            for key, value in record.items():
                if pd.isna(value):
                    record[key] = ""
        
        return data
        
    except json.JSONDecodeError as e:
        st.error(f"❌ JSON Format Error: {str(e)}")
        st.info("💡 Please check that your service account JSON is in the correct format")
        return None
    except HttpError as e:
        if e.resp.status == 404:
            st.error("❌ File not found")
            st.info("💡 Check that the file ID is correct and the file exists")
        elif e.resp.status == 403:
            st.error("❌ Access permission denied")
            st.warning("🔧 To fix this issue:")
            st.markdown("""
            1. Open the file in Google Drive
            2. Click the "Share" button
            3. Add the service account email address (shown above)
            4. Set permission to "Viewer" and send
            """)
        else:
            st.error(f"❌ Google Drive API Error: {str(e)}")
        return None
    except ValueError as e:
        error_msg = str(e)
        if "Worksheet named" in error_msg:
            st.error(f"❌ Sheet '{sheet_name}' not found")
            st.info("💡 Check that the sheet name is correct (case-sensitive)")
        else:
            st.error(f"❌ Excel Format Error: {error_msg}")
            st.info("💡 Please check that your Excel file is in the correct format")
        return None
    except Exception as e:
        st.error(f"❌ Unexpected Error: {type(e).__name__}: {str(e)}")
        st.info("💡 Please review your settings based on the error details")
        return None

# Load Excel data
def load_excel_data(excel_file, sheet_name):
    """Load data from uploaded Excel file"""
    try:
        # Read Excel file
        df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
        
        # Check if dataframe is empty
        if df.empty:
            st.error("❌ Excel file is empty")
            st.info("💡 Please ensure the Excel file contains data")
            return None
        
        # Convert to list of dictionaries (same format as Google Sheets)
        data = df.to_dict('records')
        
        # Convert NaN values to empty strings for consistency
        for record in data:
            for key, value in record.items():
                if pd.isna(value):
                    record[key] = ""
        
        return data
        
    except ValueError as e:
        error_msg = str(e)
        if "Worksheet named" in error_msg:
            st.error(f"❌ Sheet '{sheet_name}' not found")
            st.info("💡 Check that the sheet name is correct (case-sensitive)")
        else:
            st.error(f"❌ Excel Format Error: {error_msg}")
            st.info("💡 Please check that your Excel file is in the correct format")
        return None
    except Exception as e:
        st.error(f"❌ Excel Reading Error: {type(e).__name__}: {str(e)}")
        st.info("💡 Please review your Excel file and try again")
        return None

# Load spreadsheet data
def load_spreadsheet_data(creds_json, sheet_url, sheet_name):
    try:
        # Validate JSON
        creds_dict = json.loads(creds_json)
        
        # Required scopes
        scopes = ['https://www.googleapis.com/auth/spreadsheets.readonly']
        
        # Create credentials
        creds = ServiceCredentials.from_service_account_info(creds_dict, scopes=scopes)
        
        # Connect to spreadsheet
        client = gspread.authorize(creds)
        
        # Open spreadsheet from URL
        spreadsheet = client.open_by_url(sheet_url)
        
        # Get worksheet
        worksheet = spreadsheet.worksheet(sheet_name)
        
        # Get data
        data = worksheet.get_all_records()
        
        return data
        
    except json.JSONDecodeError as e:
        st.error(f"❌ JSON Format Error: {str(e)}")
        st.info("💡 Please check that your service account JSON is in the correct format")
        return None
    except gspread.exceptions.SpreadsheetNotFound:
        st.error("❌ Spreadsheet not found")
        st.info("💡 Check that the URL is correct and the service account has access")
        return None
    except PermissionError as e:
        st.error("❌ Access permission denied")
        st.warning("🔧 To fix this issue:")
        st.markdown("""
        1. Open the spreadsheet
        2. Click the "Share" button in the top right
        3. Add the service account email address (shown above)
        4. Set permission to "Viewer" and send
        """)
        return None
    except gspread.exceptions.WorksheetNotFound:
        st.error(f"❌ Sheet '{sheet_name}' not found")
        st.info("💡 Check that the sheet name is correct (case-sensitive)")
        return None
    except gspread.exceptions.APIError as e:
        st.error(f"❌ Google API Error: {str(e)}")
        st.info("💡 Make sure Google Sheets API is enabled")
        return None
    except Exception as e:
        st.error(f"❌ Unexpected Error: {type(e).__name__}: {str(e)}")
        st.info("💡 Please review your settings based on the error details")
        return None

# Email sending function
def send_email_simple(to, subject, body, sender_email, app_password):
    """Send email using App Password"""
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = to
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, app_password)
        server.send_message(msg)
        server.quit()
        return True
    except smtplib.SMTPAuthenticationError as e:
        error_msg = str(e)
        if "Application-specific password required" in error_msg or "InvalidSecondFactor" in error_msg:
            st.error(f"❌ Authentication Error ({to}): App password required")
            st.warning("⚠️ Regular Gmail passwords cannot be used. Enable 2-step verification and generate an app password.")
        else:
            st.error(f"❌ Authentication Error ({to}): {error_msg}")
        return False
    except Exception as e:
        st.error(f"❌ Sending Error ({to}): {str(e)}")
        return False

# Template processing function
def apply_template(template, data):
    result = template
    for key, value in data.items():
        placeholder = f"{{{{{key}}}}}"
        result = result.replace(placeholder, str(value))
    return result

# Data preview
if st.button("📊 Preview Data", type="secondary"):
    if data_source == "Google Sheets":
        if sheets_credentials_json and spreadsheet_url:
            with st.spinner("Loading data..."):
                data = load_spreadsheet_data(sheets_credentials_json, spreadsheet_url, sheet_name)
                if data:
                    # Save to session state
                    st.session_state.available_columns = list(data[0].keys()) if data else ["email"]
                    st.session_state.loaded_data = data
                    st.session_state.data_source_key = data_source
        else:
            st.warning("Please enter authentication credentials and spreadsheet URL")
    elif data_source == "Excel File (Local Upload)":
        if excel_file is not None:
            with st.spinner("Loading data..."):
                data = load_excel_data(excel_file, sheet_name)
                if data:
                    # Save to session state
                    st.session_state.available_columns = list(data[0].keys()) if data else ["email"]
                    st.session_state.loaded_data = data
                    st.session_state.data_source_key = data_source
        else:
            st.warning("Please upload an Excel file")
    else:  # Google Drive Excel
        if sheets_credentials_json and drive_file_url:
            with st.spinner("Loading data..."):
                data = load_google_drive_excel(sheets_credentials_json, drive_file_url, sheet_name)
                if data:
                    # Save to session state
                    st.session_state.available_columns = list(data[0].keys()) if data else ["email"]
                    st.session_state.loaded_data = data
                    st.session_state.data_source_key = data_source
        else:
            st.warning("Please enter authentication credentials and Google Drive file URL")

# Display loaded data if exists
if st.session_state.loaded_data is not None:
    st.success(f"✅ Loaded {len(st.session_state.loaded_data)} records")
    st.dataframe(st.session_state.loaded_data, use_container_width=True)

# Email column selection (always visible)
st.markdown("---")
st.subheader("📧 Email Column Selection")

# Set default index
default_index = 0
if 'email' in st.session_state.available_columns:
    default_index = st.session_state.available_columns.index('email')

email_column = st.selectbox(
    "Select the column containing email addresses",
    options=st.session_state.available_columns,
    index=default_index,
    key="email_column_selector",
    help="Choose which column contains the recipient email addresses"
)

st.session_state.email_column = email_column
st.info(f"💡 Using column '{email_column}' for email addresses")

st.markdown("---")

# Sending method selection
st.subheader("📤 Email Sending Setup")

st.error("⚠️ Important: Regular Gmail passwords cannot be used. You must generate an app password.")

st.info("""
💡 **How to Get Google App Password (5 minutes):**

1. **Enable 2-Step Verification**
   - Visit https://myaccount.google.com/security
   - Click "Signing in to Google" → "2-Step Verification"
   - Register your phone number and enable it

2. **Generate App Password**
   - Visit https://myaccount.google.com/apppasswords
   - Select app: "Mail"
   - Select device: "Other (Custom name)" → Enter any name
   - Click "Generate"
   - Copy the 16-digit password displayed (e.g., `abcd efgh ijkl mnop`)

3. **Paste it below**
   - Spaces will be automatically removed
""")

app_password = st.text_input(
    "App Password (16 digits)", 
    type="password", 
    help="Enter the 16-digit app password (spaces will be ignored)",
    placeholder="abcdefghijklmnop or abcd efgh ijkl mnop"
)

# Send execution
if st.button("📤 Send Emails", type="primary"):
    # Validate based on data source
    validation_error = False
    
    if data_source == "Google Sheets":
        if not sheets_credentials_json or not spreadsheet_url or not sender_email:
            st.error("Please fill in all required fields")
            validation_error = True
    elif data_source == "Excel File (Local Upload)":
        if excel_file is None or not sender_email:
            st.error("Please fill in all required fields")
            validation_error = True
    else:  # Google Drive Excel
        if not sheets_credentials_json or not drive_file_url or not sender_email:
            st.error("Please fill in all required fields")
            validation_error = True
    
    if not validation_error:
        if not app_password:
            st.error("Please enter your app password")
        elif not subject_template or not body_template:
            st.error("Please provide both subject and body templates")
        elif 'email_column' not in st.session_state or not st.session_state.email_column:
            st.error("Please preview data and select an email column first")
        else:
            with st.spinner("Preparing to send emails..."):
                # Check if we can reuse loaded data
                if ('loaded_data' in st.session_state and 
                    st.session_state.loaded_data is not None and
                    'data_source_key' in st.session_state and 
                    st.session_state.data_source_key == data_source):
                    data = st.session_state.loaded_data
                else:
                    # Reload data
                    if data_source == "Google Sheets":
                        data = load_spreadsheet_data(sheets_credentials_json, spreadsheet_url, sheet_name)
                    elif data_source == "Excel File (Local Upload)":
                        data = load_excel_data(excel_file, sheet_name)
                    else:  # Google Drive Excel
                        data = load_google_drive_excel(sheets_credentials_json, drive_file_url, sheet_name)
            
            if data:
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                success_count = 0
                fail_count = 0
                
                # Get selected email column
                email_col = st.session_state.get('email_column', 'email')
                
                for idx, row in enumerate(data):
                    if email_col not in row:
                        st.warning(f"⚠️ Row {idx+1}: Email address not found in column '{email_col}'")
                        fail_count += 1
                        continue
                    
                    recipient_email = row[email_col]
                    
                    # Check if email is empty
                    if not recipient_email or str(recipient_email).strip() == "":
                        st.warning(f"⚠️ Row {idx+1}: Email address is empty")
                        fail_count += 1
                        continue
                    
                    subject = apply_template(subject_template, row)
                    body = apply_template(body_template, row)
                    
                    status_text.text(f"Sending: {recipient_email} ({idx+1}/{len(data)})")
                    
                    if test_mode:
                        st.info(f"🧪 Test Mode: Simulating send to {recipient_email}")
                        with st.expander(f"Email Preview: {recipient_email}"):
                            st.write(f"**Subject:** {subject}")
                            st.write(f"**Body:**")
                            st.text(body)
                        success_count += 1
                    else:
                        # Remove spaces from app password
                        clean_password = app_password.replace(" ", "")
                        if send_email_simple(recipient_email, subject, body, sender_email, clean_password):
                            st.success(f"✅ Sent successfully: {recipient_email}")
                            success_count += 1
                        else:
                            fail_count += 1
                    
                    progress_bar.progress((idx + 1) / len(data))
                    time.sleep(delay_seconds)
                
                status_text.text("Complete")
                st.balloons()
                
                st.markdown("---")
                st.subheader("📊 Sending Results")
                col1, col2, col3 = st.columns(3)
                col1.metric("Total", len(data))
                col2.metric("Success", success_count)
                col3.metric("Failed", fail_count)

# Footer
st.markdown("---")
with st.expander("📖 How to Use"):
    st.markdown("""
    ### Setup Instructions
    
    #### Google Cloud Console Setup (for Spreadsheet)
    
    1. **Create a Google Cloud Project**
       - Visit https://console.cloud.google.com
       - Create a new project
    
    2. **Enable Google Sheets API**
       - In the project, go to "APIs & Services" → "Library"
       - Search for "Google Sheets API" and enable it
    
    3. **Create Service Account**
       - Go to "APIs & Services" → "Credentials"
       - Click "Create Credentials" → "Service Account"
       - Fill in the details and create
       - Click on the created service account
       - Go to "Keys" tab → "Add Key" → "Create new key" → "JSON"
       - Download the JSON file
    
    #### Get Google App Password (for Gmail Sending)
    
    1. **Enable 2-Step Verification**
       - Visit https://myaccount.google.com/security
       - Enable 2-Step Verification with your phone number
    
    2. **Generate App Password**
       - Visit https://myaccount.google.com/apppasswords
       - Generate a 16-digit app password
    
    #### Prepare Spreadsheet
    
    1. **Create a spreadsheet with recipient list**
       - Required column: `email`
       - Other columns for variables used in templates (e.g., `name`, `message`)
    
    2. **Share spreadsheet with service account**
       - Copy the `client_email` from the service account JSON
       - Share the spreadsheet with this email address (Viewer permission)
    
    #### Prepare Template Files (Optional)
    
    If using "Upload Files" mode:
    
    1. **Create subject template file** (e.g., `subject.txt`)
       ```
       Hello {{name}} - Special Offer
       ```
    
    2. **Create body template file** (e.g., `body.txt`)
       ```
       Dear {{name}},
       
       We are pleased to offer you:
       {{message}}
       
       Best regards,
       Support Team
       ```
    
    3. **Upload both files in the app**
       - Use plain text (.txt) files
       - Use {{variable_name}} format for placeholders
       - Variables must match column names in your spreadsheet
    
    #### Use the App
    
    1. Paste the service account JSON in the sidebar
    2. Enter the spreadsheet URL and sheet name
    3. Enter your app password
    4. Choose template source (Manual Input or Upload Files)
    5. Set up email templates (either type them or upload files)
    6. Preview data to confirm
    7. Test in Test Mode first, then send for real
    
    ### Important Notes
    - Gmail has sending limits (approximately 500 emails per day)
    - Set appropriate sending intervals for bulk emails
    - Always test in Test Mode first before actual sending
    """)