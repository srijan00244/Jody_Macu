import streamlit as st
import os
import json
import pandas as pd
import base64
import anthropic
import re
from io import BytesIO
import io
import shutil
import tempfile
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
SCOPES = ["https://www.googleapis.com/auth/drive"]

def check_password():
    """Returns True if the user entered the correct password."""
    def password_entered():
        if st.session_state["password"] == st.secrets["app_password"]:  # <-- your hardcoded password here
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Don't store password in session state
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # Show input for password
        st.text_input("Enter password:", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.text_input("Enter password:", type="password", on_change=password_entered, key="password")
        st.error("Incorrect password. Try again.")
        return False
    else:
        return True

def grade_to_points(grade):
    """Convert letter grade to numerical points."""
    grade = grade.upper().strip()
    # Handle plus/minus modifiers
    base_grade = grade[0]
    modifier = grade[1:] if len(grade) > 1 else ""
    # Base grade values
    grade_values = {
        'A': 4.0,
        'B': 3.0,
        'C': 2.0,
        'D': 1.0,
        'F': 0.0
    }
    
    # Get base value
    if base_grade not in grade_values:
        return None  # Handle non-standard grades like P, W, etc.
    
    base_value = grade_values[base_grade]
    # Apply modifiers
    if modifier == '+' and base_grade != 'A':  # A+ is still 4.0 at most schools
        base_value += 0.3
    elif modifier == '-':
        base_value -= 0.3
        
    return base_value

def analyze_pdf(pdf_data_bytes, user_prompt: str):
    # Initialize Anthropic client - consider using st.secrets for API key in production
    client = anthropic.Anthropic(api_key=st.secrets["anthropic_api_key"])
    # Encode PDF data
    pdf_data = base64.b64encode(pdf_data_bytes).decode("utf-8")
    messages_payload = [
        {
            "role": "user",
            "content": [
                {
                    "type": "document",
                    "source": {
                        "type": "base64",
                        "media_type": "application/pdf",
                        "data": pdf_data
                    }
                },
                {
                    "type": "text",
                    "text": user_prompt
                }
            ]
        }
    ]

    with st.spinner("Analyzing transcript... This may take a moment."):
        message = client.messages.create(
            model="claude-3-7-sonnet-latest",
            max_tokens=4000,
            messages=messages_payload
        )

    # Calculate and display token usage
    cache_creation_input_tokens = message.usage.cache_creation_input_tokens
    cache_read_input_tokens = message.usage.cache_read_input_tokens
    input_tokens = message.usage.input_tokens
    output_tokens = message.usage.output_tokens
    # Calculate pricing based on tokens usage (price per million tokens)
    base_input_cost = input_tokens * 3.00 / 1e6
    cache_writes_cost = cache_creation_input_tokens * 3.75 / 1e6
    cache_hits_cost = cache_read_input_tokens * 0.30 / 1e6
    output_cost = output_tokens * 15.00 / 1e6
    total_cost = base_input_cost + cache_writes_cost + cache_hits_cost + output_cost

    # Create token usage message for display in an expander
    token_usage = f"""
    **Tokens Used:** {input_tokens + output_tokens}
    
    **Pricing Breakdown:**
    - Base Input Cost: ${base_input_cost:.6f}
    - Cache Writes Cost: ${cache_writes_cost:.6f}
    - Cache Hits Cost: ${cache_hits_cost:.6f}
    - Output Cost: ${output_cost:.6f}
    - **Total Cost:** ${total_cost:.6f}
    """

    return message.content[0].text, token_usage

def extract_json(text):
    match = re.search(r'```json\n(.*?)\n```', text, re.DOTALL)

    if match:
        json_str = match.group(1).strip()
        try:
            return json.loads(json_str)
        except json.JSONDecodeError:
            st.error("Failed to parse JSON output from Claude.")
            return None
    st.error("Could not find JSON data in Claude's response.")
    return None

def post_process_transcript_data(json_data):
    """Post-process the JSON data to ensure credits are correctly calculated."""
    for term in json_data:
        for course in term.get("courses", []):
            # If credits are missing but points and grade are available
            if (not course.get("credits") or course["credits"] == "") and course.get("points") and course.get("grade"):
                grade_value = grade_to_points(course["grade"])
                if grade_value:  # Only calculate if we have a valid grade value
                    try:
                        points = float(course["points"])
                        course["credits"] = round(points / grade_value, 1)
                    except (ValueError, ZeroDivisionError):
                        # Handle cases where conversion fails
                        pass
    return json_data

def get_term_code(term):
    """Convert term name to code."""
    term = term.lower()
    if "spring" in term:
        return "TS"
    elif "fall" in term:
        return "TF"
    elif "summer" in term:
        return "TU"
    return ""

def display_transcript_data(json_data):
    """Display transcript data in tables grouped by term."""
    if not json_data:
        st.error("No data to display")
        return
        
    for term_data in json_data:
        term = term_data.get("term", "")
        year = term_data.get("year", "")
        term_code = get_term_code(term)
        # Create header for each term
        st.subheader(f"{term} - {year} [{term_code}]")
        courses = term_data.get("courses", [])
        if not courses:
            st.write("No courses found for this term")
            continue
            
        # Create DataFrame for this term's courses
        df = pd.DataFrame([
            {
                "Course Code": course.get("course_code", ""),
                "Division": course.get("division", ""),
                "Title": course.get("title", ""),
                "Short Title": course.get("short_title", ""),
                "Credit": course.get("credits", ""),
                "Grade": course.get("grade", "")
            }
            for course in courses
        ])
        
        # Display the data as a table
        st.table(df)

def show_feedback_dialog():
    """Show feedback dialog and validate input."""
    with st.form(key="feedback_form"):
        st.subheader("Feedback")
        feedback = st.text_area(
            "Please provide feedback on the transcript analysis results:",
            height=150
        )
        submit_button = st.form_submit_button(label="Submit Feedback")
        
        if submit_button:
            if not feedback.strip():
                st.error("Feedback cannot be empty. Please enter at least one character.")
                return False, None
            else:
                st.success("Thank you for your feedback!")
                return True, feedback
    return False, None

def save_pdf_to_drive(pdf_bytes: bytes, filename: str):
    temp_file = None
    temp_file_path = None
    try:
        # ==== CONFIGURATION ====
        SCOPES = ['https://www.googleapis.com/auth/drive']
        
        # ==== AUTHENTICATION ====
        credentials = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=SCOPES
        )
        drive_service = build('drive', 'v3', credentials=credentials)

        # Create a temporary file with a unique name
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp:
            temp.write(pdf_bytes)
            temp_file_path = temp.name
        
        # ==== FILE METADATA ====
        file_metadata = {
            'name': filename,
            'mimeType': 'application/pdf',
            'parents': ['1z_N8QcDkRLbMjqvDDZtO1UX3sxCzx2Os']
        }
        
        # ==== FILE UPLOAD ====
        media = MediaFileUpload(temp_file_path, mimetype='application/pdf', resumable=True)
        
        file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, name, webViewLink",
            supportsAllDrives=True
        ).execute()
        
        # Get the web view link (URL)
        file_url = file.get('webViewLink', '')
        
        # Give the system a moment before trying to delete the file
        import time
        time.sleep(0.5)
        
        return True, f"PDF uploaded successfully: {file.get('name')}", file_url
    
    except Exception as e:
        return False, f"Failed to save PDF to Google Drive: {str(e)}"
    
    finally:
        # Clean up in finally block to ensure it runs even if there's an exception
        if temp_file and os.path.exists(temp_file_path):
            try:
                # Try to close and delete the file
                os.close(os.open(temp_file_path, os.O_RDONLY))
                os.remove(temp_file_path)
            except Exception as e:
                # If deletion fails, log it but don't fail the whole operation
                print(f"Warning: Could not delete temporary file: {str(e)}")
                
def save_to_google_sheet(file_url, json_data, user_comment):
    try:
        # Import gspread
        import gspread
        # Log debugging information
        st.write(f"File URL: {file_url}")
        st.write(f"JSON data type: {type(json_data)}")
        st.write(f"User comment length: {len(user_comment)}")
        # Use service account info from secrets
        credentials = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=[
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
        )
        # Create gspread client
        gc = gspread.authorize(credentials)
        # Open the spreadsheet by ID
        spreadsheet_id = "15HvKDTzxiXueIGluwMQPKcZvess7QYSzda2yWmTZiwI"
        sheet = gc.open_by_key(spreadsheet_id).sheet1  # Using the first sheet
        # Convert JSON data to string
        json_str = json.dumps(json_data)
        # Prepare row data
        row_data = [file_url, json_str, user_comment]
        # Append the row to the sheet
        sheet.append_row(row_data)
        # Get the row number that was just added
        next_row = len(sheet.get_all_values())
        return True, f"Data saved to Google Sheet in row {next_row}"
    
    except Exception as e:
        return False, f"Failed to save to Google Sheet: {str(e)}"
                    
# Prompt template for Claude
PROMPT = """
# Transcript Data Extraction Prompt

## **Objective**
Extract the following information from the provided PDF transcript file.

## **Instructions**

### **Step 1: Check for a "Transcript Explanation" Page**
- If the document contains a "Transcript Explanation" page, refer to it before extracting any data.
- Use this page to correctly interpret the structure, grading system, and any special formatting rules in the transcript.

### **Step 2: Check for sections titled "TRANSFER CREDIT ACCEPTED BY THE INSTITUTION", "Transfer Coursework", "Transfer Credit", "Transferred Courses", or any similar wording that indicates transfer credits**.
- These are NOT part of the student's earned credits at this institution and must not be included in the extracted data.
- Do not extract courses from these sections even if they look like normal course listings.
- Only extract courses that were taken and completed **at the issuing institution**.

### **Step 3: Extract the Required Information**
For each term, extract the following details:

- **Term:** Identify the academic term (Fall, Summer, Spring).
- **Year:** Extract the 4-digit academic year.
- **Courses:** A list of courses within that term, with the following attributes:
  - **Course Code:** Extract exactly as shown under "COURSE."
  - **Division:** Determine the division based on the first digit(s) of the "Course Code":
    - **0xxx - 4xxx** → **UNDG (Undergraduate)**
    - **5xxx - 6xxx** → **GRAD (Graduate)**
  - **Title:** Extract exactly as shown under "COURSE TITLE."
  - **Short Title:** Provide a shortened version of the course title.
            - If the full title is already under 40 characters, use it as is.
            - If it's longer, create a meaningful short version (<= 40 characters) while preserving essential context.
  - **Credits:** 
            - If "CRED" or "CREDIT" column exists, extract directly from there.
            - If missing, calculate credits by dividing "GRADE POINTS" or "POINTS" by the numerical value of the grade.
            - Example: If Points = 12 and Grade = A (4.0), then Credits = 12/4 = 3.
  - **Grade:** Extract what is listed under "GRADE."
  - **Points:** Extract what is listed under "GRADE POINTS" or "POINTS" if available.

### **Step 4: Output Format**
Return the extracted data in the following **JSON structure**:

```json
[
  {
    "term": "Fall",
    "year": "2023",
    "courses": [
      {
        "course_code": "CS101",
        "division": "UNDG",
        "title": "Real-Time Text and voice output enabled traffic sign detection system using deep learning",
        "short_title: "Real-Time Traffic Sign Detection",
        "credits": 3,
        "grade": "A"
      },
      {
        "course_code": "MATH202",
        "division": "UNDG",
        "title": "Calculus II",
        "short_title": "Calculus II",
        "credit_hours": 4,
        "grade": "B+"
      }
    ]
  },
  {
    "term": "Spring",
    "year": "2024",
    "courses": [
      {
        "course_code": "MATH5001",
        "division": "GRAD",
        "title": "Advanced Calculus",
        "short_title": "Advanced Calculus",
        "credits": 4,
        "grade": "A-"
      }
    ]
  }
]

## **Additional Considerations**
- If "CRED" is missing, calculate credits using: CRED = Points/Grade where grade values are A=4.0, B=3.0, C=2.0, D=1.0, F=0.0
- Plus/minus modifiers adjust by 0.3 (e.g., A- = 3.7, B+ = 3.3)
- Ensure that each course is correctly associated with its respective term and year.
- Make sure to extract and include the "points" field in the output as it's needed for credit calculation.
- If any required information is missing from a course, leave the value as an empty string ("") rather than omitting the field.
"""
# Streamlit app
def main():
    st.set_page_config(page_title="Transcript Analyzer", layout="wide")
    st.title("🔍 Academic Transcript Analyzer")
    
    # Initialize session state variables if they don't exist
    if "pdf_processed" not in st.session_state:
        st.session_state["pdf_processed"] = False
    if "feedback_submitted" not in st.session_state:
        st.session_state["feedback_submitted"] = False
    if "uploaded_file_name" not in st.session_state:
        st.session_state["uploaded_file_name"] = None
    if "pdf_bytes" not in st.session_state:
        st.session_state["pdf_bytes"] = None
    if "drive_upload_status" not in st.session_state:
        st.session_state["drive_upload_status"] = None
    
    # Step 1: Ask for password
    if not check_password():
        st.warning("Please enter the password to access the transcript analyzer.")
        st.stop()  # Don't run the rest of the app until the correct password is entered

    # Step 2: Show app functionality after successful login
    st.success("Access granted. You may now upload and analyze transcripts.")
    
    # If feedback has not been submitted after processing a PDF, show the feedback dialog
    if st.session_state.get("pdf_processed", False) and not st.session_state.get("feedback_submitted", False):
        feedback_submitted, feedback_text = show_feedback_dialog()
        if feedback_submitted:
            st.session_state["feedback_submitted"] = True
            # After feedback is submitted, save the PDF to Google Drive
        if st.session_state.get("pdf_bytes") and st.session_state.get("uploaded_file_name"):
            success, message,file_url = save_pdf_to_drive(
                st.session_state["pdf_bytes"], 
                st.session_state["uploaded_file_name"]
            )
            
            if success:
                st.session_state["drive_upload_status"] = "success"
                st.success(f"PDF successfully saved to Google Drive!")
                
                sheet_success = False
                sheet_message = ""
                
                if "json_data" in st.session_state and file_url:
                    st.write("Attempting to save data to Google Sheet...")
                    sheet_success, sheet_message = save_to_google_sheet(
                        file_url, 
                        st.session_state["json_data"], 
                        feedback_text
                    )
                    
                if sheet_success:
                    st.success(sheet_message)
                else:
                    st.error(sheet_message)
                    st.error("Failed to save data to Google Sheet. Please check the logs for details.")
                if file_url:
                    st.markdown(f"[View the file in Google Drive]({file_url})")
            else:
                st.session_state["drive_upload_status"] = "error"
                st.error(f"Failed to save PDF to Google Drive: {message}")
    else:
        # Show upload status from previous submission if available
        if st.session_state.get("drive_upload_status") == "success":
            st.success("Previous PDF was successfully saved to Google Drive.")
            # Clear the status to avoid showing it repeatedly
            st.session_state["drive_upload_status"] = None
        
        st.write("Upload a PDF transcript to extract course information.")
        uploaded_file = st.file_uploader("Choose a transcript PDF file", type="pdf")
        
        if uploaded_file is not None:
            # Save uploaded file contents to session state
            pdf_bytes = uploaded_file.getvalue()
            st.session_state["pdf_bytes"] = pdf_bytes
            st.session_state["uploaded_file_name"] = uploaded_file.name
            
            # Process the transcript
            if st.button("Process Transcript"):
                # Call Claude API to analyze the PDF
                claude_response, token_usage = analyze_pdf(pdf_bytes, PROMPT)
                # Extract JSON from Claude's response
                json_data = extract_json(claude_response)
                # Post-process the data
                if json_data:
                    json_data = post_process_transcript_data(json_data)
                    st.session_state["json_data"] = json_data
                    # Display the data
                    st.success("Transcript processed successfully!")
                    # Add download button for JSON
                    st.download_button(
                        label="Download JSON Data",
                        data=json.dumps(json_data, indent=4),
                        file_name=f"{uploaded_file.name.split('.')[0]}_processed.json",
                        mime="application/json"
                    )
                    
                    # Display token usage details in an expander
                    with st.expander("API Token Usage Details"):
                        st.markdown(token_usage)
                    
                    # Display the transcript data in tables
                    display_transcript_data(json_data)
                    # Show raw JSON in an expander
                    with st.expander("View Raw JSON Data"):
                        st.json(json_data)
                    
                    # Set the state to show that a PDF has been processed
                    st.session_state["pdf_processed"] = True
                    st.session_state["feedback_submitted"] = False
                    # Show feedback dialog after displaying results
                    st.markdown("---")
                    show_feedback_dialog()
                else:
                    st.error("Failed to extract data from the transcript.")
                    st.write("Raw response from Claude:")
                    st.text(claude_response)

if __name__ == "__main__":
    main()
