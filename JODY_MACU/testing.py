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
from googleapiclient.http import MediaIoBaseDownload
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
    grade = grade.upper().strip()
    base_grade = grade[0]
    modifier = grade[1:] if len(grade) > 1 else ""
    grade_values = {
        'A': 4.0,
        'B': 3.0,
        'C': 2.0,
        'D': 1.0,
        'F': 0.0
    }
    
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
    client = anthropic.Anthropic(api_key=st.secrets["anthropic_api_key"])
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

    try:
        with st.spinner("Analyzing transcript... This may take a moment."):
            message = client.messages.create(
                model="claude-3-7-sonnet-latest",
                max_tokens=8000,
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
    
    except anthropic.APIStatusError as e:
        # Handle specific HTTP status codes
        if e.status_code == 529:
            st.error("⚠️ Claude is currently experiencing high demand. Please try again in a few minutes.")
        elif e.status_code == 429:
            st.error("⚠️ API rate limit exceeded. Please wait a moment before trying again.")
        elif e.status_code >= 500:
            st.error("⚠️ Claude service is temporarily unavailable. Please try again later.")
        else:
            st.error(f"⚠️ API Error: {str(e)}")
        return None, None
        
    except anthropic.APIConnectionError:
        st.error("⚠️ Connection to Claude API failed. Please check your internet connection and try again.")
        return None, None
        
    except anthropic.APITimeoutError:
        st.error("⚠️ The request to Claude timed out. This PDF may be too complex or the service is busy. Please try again later.")
        return None, None
        
    except anthropic.AuthenticationError:
        st.error("⚠️ Authentication to Claude API failed. Please contact the administrator to check API credentials.")
        return None, None
        
    except Exception as e:
        st.error(f"⚠️ An unexpected error occurred: {str(e)}")
        return None, None

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
    # Ensure json_data is a list and not empty before accessing elements
    if json_data and isinstance(json_data, list) and len(json_data) > 0:
        # Check if the first term has an institution field
        if "institution" in json_data[0]:
            institution_name = json_data[0].get("institution", "")
            # Propagate to all terms
            for term in json_data:
                if "institution" not in term:
                    term["institution"] = institution_name
                
    # Process each term's courses        
    for term in json_data or []:
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
def load_institution_mappings():
    """Load institution name to code mappings from Google Sheet."""
    try:
        import gspread
        credentials = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=[
                'https://www.googleapis.com/auth/spreadsheets.readonly',
                'https://www.googleapis.com/auth/drive.readonly'
            ]
        )
        
        gc = gspread.authorize(credentials)
        spreadsheet_id = "122e-sqnpQWkue_uGxLLrcc7nuwBWppzUeh9cdp6vpRY"
        
        try:
            spreadsheet = gc.open_by_key(spreadsheet_id)
        except Exception as e:
            st.error(f"Failed to open SchoolInstitutions spreadsheet: {str(e)}")
            return pd.DataFrame()
        
        worksheet = spreadsheet.sheet1  # Using the first sheet
        sheet_values = worksheet.get_all_values()
        if not sheet_values or len(sheet_values) <= 1:
            st.error("SchoolInstitutions sheet is empty or contains insufficient data")
            return pd.DataFrame()
        
        headers = sheet_values[0]
        data = sheet_values[1:]
        df = pd.DataFrame(data, columns=headers)
        
        # Ensure the required columns exist
        if "ORG_NAME" not in df.columns or "ORG_CDE" not in df.columns:
            st.error("Required columns 'ORG_NAME' or 'ORG_CDE' not found in SchoolInstitutions sheet")
            return pd.DataFrame()
        
        return df
        
    except Exception as e:
        st.error(f"Error loading institution mappings from Google Sheets: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return pd.DataFrame()
def match_institution_code(institution_name, institution_df):

    if institution_df.empty or not institution_name:
        return ""
    
    # Normalize institution names for matching
    institution_name = institution_name.lower().strip()
    institution_df['ORG_NAME_NORMALIZED'] = institution_df['ORG_NAME'].str.lower().str.strip()
    
    # Try exact match first
    exact_matches = institution_df[institution_df['ORG_NAME_NORMALIZED'] == institution_name]
    
    if not exact_matches.empty:
        org_code = exact_matches.iloc[0]['ORG_CDE']
    else:
        # Try fuzzy matching if no exact match
        import difflib
        best_matches = difflib.get_close_matches(
            institution_name, 
            institution_df['ORG_NAME_NORMALIZED'].tolist(),
            n=1,  # Get only the best match
            cutoff=0.7  # Require at least 70% similarity
        )
        
        if best_matches:
            best_match = best_matches[0]
            matching_rows = institution_df[institution_df['ORG_NAME_NORMALIZED'] == best_match]
            if not matching_rows.empty:
                org_code = matching_rows.iloc[0]['ORG_CDE']
            else:
                return ""
        else:
            return ""
    
    # Format org code to be 6 digits with leading zeros
    try:
        # Convert to integer to remove any leading zeros, then format to 6 digits
        org_code = str(int(org_code)).zfill(6)
        return org_code
    except (ValueError, TypeError):
        # If conversion fails (e.g., non-numeric code), return as is
        return org_code

def enrich_with_macu_data(json_data, macu_df, ceqmacu_df=None):
    if macu_df.empty:
        st.warning("No CEP mapping data available.")
        return json_data
    import re
    
    def normalize(text):
        if pd.isna(text) or text is None:
            return ""
        # Replace hyphens with spaces in the text
        text = str(text).strip().lower().replace('-', ' ')
        # Add space between letters and numbers for consistent matching
        text = re.sub(r'([a-zA-Z])(\d)', r'\1 \2', text)
        return text
    
    # Determine which academic year sheet to use for each term
    def get_academic_year_sheet(term, year):
        year = int(year) if year.isdigit() else 0
        term = term.lower()
        
        # Map the term and year to appropriate academic year
        if "fall" in term:
            # Fall term is in the first year of an academic year span
            academic_year = f"{year}-{year+1}"
        elif "spring" in term or "summer" in term:
            # Spring and Summer terms are in the second year of an academic year span
            academic_year = f"{year-1}-{year}"
        else:
            # Default case if term is unrecognized
            academic_year = f"{year}-{year+1}"
            
        return academic_year
    
    # Improved extract_course_code function
    def extract_course_code(combined_text):
        if pd.isna(combined_text) or combined_text is None:
            return ""
        
        # First try a more robust pattern that looks for a subject code followed by a course number
        # This captures patterns like "COMM 1313", "ENGL 101", "BIO 2010", etc.
        match = re.match(r'^([A-Za-z]+)\s*(\d+)', str(combined_text), re.IGNORECASE)
        if match:
            subject = match.group(1).strip()
            number = match.group(2).strip()
            return normalize(f"{subject} {number}")
        
        # Fallback to the original pattern
        match = re.match(r'^([A-Za-z0-9\s\.]+?)(?:\s{2,}|\s+[^A-Za-z0-9\s\.])', str(combined_text))
        if match:
            return normalize(match.group(1))
        else:
            # Final fallback: try to get the first word with numbers (likely the course code)
            words = str(combined_text).split()
            for i, word in enumerate(words):
                if any(c.isdigit() for c in word) and i > 0:
                    return normalize(f"{words[i-1]} {word}")  # Subject code + course number
            
            # If nothing else works, just take the first two words if available
            if len(words) >= 2:
                return normalize(f"{words[0]} {words[1]}")
            return normalize(str(combined_text).split()[0]) if words else ""
    
    # Create a more efficient structure for course code lookup
    # Use the CombineTitleCode column for matching
    combine_column = 'CombineTitleCode'
    if combine_column not in macu_df.columns:
        # Look for alternative columns that might contain the combined data
        potential_columns = ['Combine'] 
        for col in potential_columns:
            if col in macu_df.columns:
                combine_column = col
                break
        else:
            st.error("No suitable column found for combined course code and title matching")
            return json_data
    
    # Create normalized columns for matching
    macu_df['combine_normalized'] = macu_df[combine_column].apply(normalize)
    macu_df['common_code_normalized'] = macu_df['CommonCode'].apply(normalize)
    macu_df['course_code_extracted'] = macu_df[combine_column].apply(extract_course_code)
    
    # Create a column with just the course code for secondary matching
    if 'CourseCode' in macu_df.columns:
        macu_df['course_code_normalized'] = macu_df['CourseCode'].apply(normalize)
    
    # Create course code lookup dictionary for faster matching
    course_code_lookup = {}
    
    # Create filtered dataframes for each academic year
    academic_year_dfs = {}
    available_sheets = ['2020-2021', '2021-2022', '2022-2023', '2023-2024', '2024-2025', '2025-2026']
    
    for sheet_name in available_sheets:
        sheet_df = macu_df[macu_df['source_sheet'] == sheet_name].copy()
        academic_year_dfs[sheet_name] = sheet_df
        
        # Create a lookup dictionary for course codes in this sheet
        for _, row in sheet_df.iterrows():
            code = row['course_code_extracted']
            if code and code not in course_code_lookup:
                course_code_lookup[code] = sheet_name
    
    # Create a specific dataframe for MACU institution rows for the second lookup
    macu_institution_df = macu_df[macu_df['Institution'] == 'MACU'].copy()
    
    # Phase 2: Setup for CEQMACU data
    ceqmacu_available = False
    if ceqmacu_df is not None and not ceqmacu_df.empty:
        ceqmacu_available = True
        ceqmacu_df['send_course_code_normalized'] = ceqmacu_df['SendCourse1CourseCode'].apply(normalize)
    
    # Count variables for tracking matches
    total_courses = 0
    cep_matches = 0
    macu_matches = 0
    ceqmacu_matches = 0
    sheet_matches = {'2020-2021': 0, '2021-2022': 0, '2022-2023': 0, '2023-2024': 0, '2024-2025': 0, '2025-2026': 0}
    older_courses = 0  # Count courses older than our available data
    
    for term in json_data:
        term_name = term.get("term", "")
        year = term.get("year", "")
        year_int = int(year) if year.isdigit() else 0
        academic_year = get_academic_year_sheet(term_name, year)
        
        # Flag to mark terms older than our available data
        is_old_term = False
        earliest_year = 2020  # Earliest year in our available sheets
        
        # Check if term is before our earliest data
        if "fall" in term_name.lower():
            if year_int < earliest_year:
                is_old_term = True
        elif "spring" in term_name.lower() or "summer" in term_name.lower():
            if year_int <= earliest_year:  # For spring/summer 2020, academic year would be 2019-2020 which we don't have
                is_old_term = True
                
        # Get the appropriate academic year dataframe
        current_academic_year_df = academic_year_dfs.get(academic_year, pd.DataFrame())
        
        for course in term.get("courses", []):
            total_courses += 1
            
            # Initialize match flags
            course["cep_match"] = False
            course["ceqmacu_match"] = False
            course["macu_division"] = ""
            course_code = course.get("course_code", "")
            title = course.get("title", "")
            combined_text = f"{course_code} {title}"
            combined_normalized = normalize(combined_text)
            course_code_normalized = normalize(course_code)
            course["CombineTitleCode"] = combined_text
            course["term_academic_year"] = academic_year
            
            # DEBUG: Log the course being processed
            # st.write(f"Processing course: {course_code} - {title}")
            
            # Mark courses from older terms explicitly
            if is_old_term:
                older_courses += 1
                course["older_than_data"] = True
                # For older terms, skip CEP matching and try CEQMACU directly
                cep_match_found = False
                
                # Add a note to indicate why no match was found in CEP
                course["data_from"] = ""
                course["no_match_reason"] = f"Term ({term_name} {year}) is before earliest available data (2020-2021)"
            else:
                course["older_than_data"] = False
                cep_match_found = False
                
                # MATCH METHOD 1: Try to find an exact match by course code only in the current academic year
                if not current_academic_year_df.empty:
                    # Print normalized course code for debugging
                    # st.write(f"Looking for course code: {course_code_normalized}")
                    
                    # First try an exact course code match
                    # Using both original and normalized course codes to increase matching chances
                    matching_rows = current_academic_year_df[
                        (current_academic_year_df['course_code_extracted'] == course_code_normalized) |
                        (current_academic_year_df['course_code_extracted'] == course_code.lower().strip())
                    ]
                    
                    if not matching_rows.empty:
                        # We found a matching course in the expected academic year sheet by course code
                        match = matching_rows.iloc[0]
                        common_code = normalize(match.get('CommonCode', ''))
                        course["cep_match"] = True
                        course["common_code"] = common_code
                        course["source_sheet"] = academic_year
                        course["matched_on"] = "course_code_exact"
                        cep_matches += 1
                        sheet_matches[academic_year] += 1
                        # Find the MACU course with the same CommonCode
                        if common_code:
                            # Look for rows where Institution = "MACU" and CommonCode matches
                            macu_matches_df = macu_institution_df[macu_institution_df['common_code_normalized'] == common_code]
                            if not macu_matches_df.empty:
                                # Found a MACU equivalent
                                macu_match = macu_matches_df.iloc[0]
                                course["macu_course_code"] = macu_match.get('CourseCode', '').replace(' ', '')
                                course["macu_course_title"] = macu_match.get('CommonCourseTitle', '')
                                course["macu_credits"] = course.get("credits", "")
                                course["data_from"] = "CEP"
                                course["macu_division"] = "C" if course.get("division", "") == "UNDG" else ""
                                macu_matches += 1
                                cep_match_found = True
                            else:
                                # Common code exists but no MACU institution match was found
                                course["data_from"] = " "
                                course["no_match_reason"] = "Common code found but no matching MACU course"
                                cep_match_found = True  # We did find a CEP match, just not a MACU match
                
                # If no match by course code, try the combined text approach for the current academic year
                if not cep_match_found and not current_academic_year_df.empty:
                    matching_rows = current_academic_year_df[current_academic_year_df['combine_normalized'] == combined_normalized]
                    if not matching_rows.empty:
                        # Found a matching course by combined text
                        match = matching_rows.iloc[0]
                        common_code = normalize(match.get('CommonCode', ''))
                        course["cep_match"] = True
                        course["common_code"] = common_code
                        course["source_sheet"] = academic_year
                        course["matched_on"] = "combined_text_exact"
                        cep_matches += 1
                        sheet_matches[academic_year] += 1
                        
                        # Find the MACU course with the same CommonCode
                        if common_code:
                            macu_matches_df = macu_institution_df[macu_institution_df['common_code_normalized'] == common_code]
                            
                            if not macu_matches_df.empty:
                                macu_match = macu_matches_df.iloc[0]
                                course["macu_course_code"] = macu_match.get('CourseCode', '').replace(' ', '')
                                course["macu_course_title"] = macu_match.get('CommonCourseTitle', '')
                                course["macu_credits"] = course.get("credits", "")
                                course["data_from"] = "CEP"
                                course["macu_division"] = "C" if course.get("division", "") == "UNDG" else ""
                                macu_matches += 1
                                cep_match_found = True
                            else:
                                course["data_from"] = ""
                                course["no_match_reason"] = "Common code found but no matching MACU course"
                                cep_match_found = True
                
                # If no match in the current academic year, try other sheets by course code first
                if not cep_match_found:
                    # Sort available sheets to try the closest years first
                    # For example, if academic_year is "2023-2024", try "2022-2023" before "2020-2021"
                    try:
                        target_year = int(academic_year.split('-')[0])
                        sorted_sheets = sorted(available_sheets, 
                                           key=lambda x: abs(int(x.split('-')[0]) - target_year))
                    except (ValueError, IndexError):
                        # If parsing fails, use the default order
                        sorted_sheets = available_sheets
                    
                    for sheet_name in sorted_sheets:
                        # Skip if it's the same as the current academic year we already checked
                        if sheet_name == academic_year:
                            continue
                            
                        sheet_df = academic_year_dfs.get(sheet_name, pd.DataFrame())
                        if sheet_df.empty:
                            continue
                        
                        # First try to match by course code
                        matching_rows = sheet_df[
                            (sheet_df['course_code_extracted'] == course_code_normalized) |
                            (sheet_df['course_code_extracted'] == course_code.lower().strip())
                        ]
                        match_type = "course_code_exact_different_year"
                        
                        # If no match by course code, try combined text
                        if matching_rows.empty:
                            matching_rows = sheet_df[sheet_df['combine_normalized'] == combined_normalized]
                            match_type = "combined_text_exact_different_year"
                        
                        if not matching_rows.empty:
                            # Found a match in another sheet
                            match = matching_rows.iloc[0]
                            common_code = normalize(match.get('CommonCode', ''))
                            course["cep_match"] = True
                            course["common_code"] = common_code
                            course["source_sheet"] = sheet_name  # Use the actual sheet where match was found
                            course["matched_on"] = match_type
                            cep_matches += 1
                            sheet_matches[sheet_name] += 1
                            # Find the MACU course with the same CommonCode
                            if common_code:
                                macu_matches_df = macu_institution_df[macu_institution_df['common_code_normalized'] == common_code]
                                
                                if not macu_matches_df.empty:
                                    macu_match = macu_matches_df.iloc[0]
                                    course["macu_course_code"] = macu_match.get('CourseCode', '').replace(' ', '')
                                    course["macu_course_title"] = macu_match.get('CommonCourseTitle', '')
                                    course["macu_credits"] = course.get("credits", "")
                                    course["data_from"] = "CEP"
                                    course["macu_division"] = "C" if course.get("division", "") == "UNDG" else ""
                                    macu_matches += 1
                                    cep_match_found = True
                                    break  # Exit the loop once match is found
                                else:
                                    course["data_from"] = "S"
                                    course["no_match_reason"] = "Common code found but no matching MACU course"
                                    cep_match_found = True
                                    break  # Exit the loop once match is found
            
            # MATCH METHOD 4: If no match in CEP data, try CEQMACU data
            if not cep_match_found and ceqmacu_available:
                # Try to match by exact course code first
                ceqmacu_matches_df = ceqmacu_df[
                    (ceqmacu_df['send_course_code_normalized'] == course_code_normalized) |
                    (ceqmacu_df['send_course_code_normalized'] == course_code.lower().strip())
                ]
                
                if not ceqmacu_matches_df.empty:
                    valid_year_matches = []
                    
                    for _, row in ceqmacu_matches_df.iterrows():
                        try:
                            low_year = int(row.get('SendEditionLowYear', 0))
                            if int(year) >= low_year:
                                valid_year_matches.append(row)
                        except (ValueError, TypeError):
                            # If year conversion fails, include the row anyway
                            valid_year_matches.append(row)
                    
                    # If we have valid matches, use the first one
                    if valid_year_matches:
                        match = valid_year_matches[0]
                        course["ceqmacu_match"] = True
                        course["macu_course_code"] = match.get('ReceiveCourse1CourseCode', '').replace(' ', '')
                        course["macu_course_title"] = match.get('ReceiveCourse1CourseTitle', '')
                        course["macu_credits"] = match.get('ReceiveCourse1Units', '')
                        course["data_from"] = "CEQMACU"
                        course["matched_on"] = "ceqmacu_course_code"
                        # Add MACU Division
                        course["macu_division"] = "C" if course.get("division", "") == "UNDG" else ""
                        ceqmacu_matches += 1
                    elif is_old_term:
                        # If this is an old term and we couldn't find a match in CEQMACU either
                        course["no_match_reason"] = f"Term ({term_name} {year}) is before earliest available data (2020-2021) and no CEQMACU match found"
            
            # Add "NO_MATCH" for data_from if we didn't find any match
            if not course.get("data_from"):
                course["data_from"] = " "
                # If no explicit reason was set, add a generic one
                if not course.get("no_match_reason"):
                    if is_old_term:
                        course["no_match_reason"] = f"Term ({term_name} {year}) is before earliest available data (2020-2021)"
                    else:
                        course["no_match_reason"] = "No matching course found in any available data source"
    
    # Add match statistics as metadata
    match_stats = {
        "total_courses": total_courses,
        "cep_matches": cep_matches,
        "macu_matches": macu_matches,
        "ceqmacu_matches": ceqmacu_matches,
        "older_courses": older_courses,
        "sheet_matches": sheet_matches
    }
    
    if json_data and len(json_data) > 0:
        json_data[0]["match_statistics"] = match_stats
    
    return json_data
def load_ceqmacu_mappings():
    try:
        import gspread
        credentials = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=[
                'https://www.googleapis.com/auth/spreadsheets.readonly',
                'https://www.googleapis.com/auth/drive.readonly'
            ]
        )
        
        gc = gspread.authorize(credentials)
        spreadsheet_id = "12CpxGQMyTa_cwyY0B-iomDgflD24kjYFYPLWljD6Jgo"
        try:
            spreadsheet = gc.open_by_key(spreadsheet_id)
            # Removed success message
        except Exception as e:
            st.error(f"Failed to open CEQMACU spreadsheet: {str(e)}")
            return pd.DataFrame()
        
        worksheet = spreadsheet.get_worksheet(0)  # Assuming data is in the first sheet
        sheet_values = worksheet.get_all_values()
        if not sheet_values or len(sheet_values) <= 1:
            st.warning(f"CEQMACU sheet is empty or contains insufficient data")
            return pd.DataFrame()
        
        headers = sheet_values[0]
        data = sheet_values[1:]
        df = pd.DataFrame(data, columns=headers)
        return df
        
    except Exception as e:
        st.error(f"Error loading CEQMACU mappings from Google Sheets: {str(e)}")
        import traceback
        st.error(traceback.format_exc())
        return pd.DataFrame()

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
    if not json_data or not isinstance(json_data, list) or len(json_data) == 0:
        st.error("No data to display")
        return
    
    # Use the institution dataframe from session state instead of loading it again
    institution_df = st.session_state.get("institution_df", pd.DataFrame())
    
    # Get institution name from the first term
    institution = json_data[0].get("institution", "")
    
    # Get institution code if institution name is available
    institution_code = ""
    if institution:
        institution_code = match_institution_code(institution, institution_df)
        
    # Display institution name and code at the top
    if institution:
        if institution_code:
            st.header(f"Institution: {institution} (Code: {institution_code})")
        else:
            st.header(f"Institution: {institution}")
        
    for term_data in json_data:
        term = term_data.get("term", "")
        year = term_data.get("year", "")
        term_code = get_term_code(term)
        st.subheader(f"{term} - {year} [{term_code}]")
        courses = term_data.get("courses", [])
        if not courses:
            st.write("No courses found for this term")
            continue
            
        df = pd.DataFrame([
            {
                "Course Code": course.get("course_code", ""),
                "Division": course.get("division", ""),
                "Title": course.get("title", ""),
                "Short Title": course.get("short_title", ""),
                "Credit": course.get("credits", ""),
                "Grade": course.get("grade", ""),
                "MACU Course Code": course.get("macu_course_code", ""),
                "MACU Course Title": course.get("macu_course_title", ""),
                "MACU Credits": course.get("macu_credits", ""),
                "MACU Division": course.get("macu_division", ""),
                "Data From": course.get("data_from", "")
            }
            for course in courses
        ])
        st.table(df)

def show_feedback_dialog():
    with st.form(key="feedback_form"):
        st.subheader("Feedback (Optional)")
        feedback = st.text_area(
            "Please provide feedback on the transcript analysis results:",
            height=150
        )
        col1, col2 = st.columns(2)
        with col1:
            submit_button = st.form_submit_button(label="Submit Feedback")
        with col2:
            skip_button = st.form_submit_button(label="Skip Feedback")
        
        if submit_button:
            if not feedback.strip():
                st.error("Feedback cannot be empty. Please enter at least one character or click 'Skip Feedback'.")
                return False, None
            else:
                st.success("Thank you for your feedback!")
                return True, feedback
        elif skip_button:
            return "skipped", None
    return False, None

def save_pdf_to_drive(pdf_bytes: bytes, filename: str):
    temp_file = None
    temp_file_path = None
    try:
        SCOPES = ['https://www.googleapis.com/auth/drive']
        credentials = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=SCOPES
        )
        drive_service = build('drive', 'v3', credentials=credentials)
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp:
            temp.write(pdf_bytes)
            temp_file_path = temp.name
        
        file_metadata = {
            'name': filename,
            'mimeType': 'application/pdf',
            'parents': ['1z_N8QcDkRLbMjqvDDZtO1UX3sxCzx2Os']
        }
        media = MediaFileUpload(temp_file_path, mimetype='application/pdf', resumable=True)
        
        file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, name, webViewLink",
            supportsAllDrives=True
        ).execute()
        file_url = file.get('webViewLink', '')
        import time
        time.sleep(0.5)
        return True, f"PDF uploaded successfully: {file.get('name')}", file_url
    
    except Exception as e:
        return False, f"Failed to save PDF to Google Drive: {str(e)}"
    
    finally:
        if temp_file and os.path.exists(temp_file_path):
            try:
                # Try to close and delete the file
                os.close(os.open(temp_file_path, os.O_RDONLY))
                os.remove(temp_file_path)
            except Exception as e:
                print(f"Warning: Could not delete temporary file: {str(e)}")
                
def save_to_google_sheet(file_url, json_data, user_comment):
    try:
        import gspread
        credentials = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=[
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
        )
        # Create gspread client
        gc = gspread.authorize(credentials)
        spreadsheet_id = "1n_jJ9Lq1lhNvQ6tWXZra4d4H_fLemXIqmHTyuWf4qEc"
        sheet = gc.open_by_key(spreadsheet_id).sheet1  # Using the first sheet
        json_str = json.dumps(json_data)
        row_data = [file_url, json_str, user_comment]
        sheet.append_row(row_data)
        next_row = len(sheet.get_all_values())
        return True, f"Data saved to Google Sheet in row {next_row}"
    
    except Exception as e:
        return False, f"Failed to save to Google Sheet: {str(e)}"

def load_macu_mappings_from_sheets():
    try:
        import gspread
        credentials = service_account.Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=[
                'https://www.googleapis.com/auth/spreadsheets.readonly',
                'https://www.googleapis.com/auth/drive.readonly'
            ]
        )

        gc = gspread.authorize(credentials)
        spreadsheet_id = "1p2_1E25dYfWWb2ugfsFSdDPss-ahzGBxaQ41YUkVRK4"
        try:
            spreadsheet = gc.open_by_key(spreadsheet_id)

        except Exception as e:
            st.error(f"Failed to open spreadsheet: {str(e)}")
            return pd.DataFrame()

        target_sheets = ["2020-2021", "2021-2022", "2022-2023", "2023-2024", "2024-2025","2025-2026"]
        all_data = pd.DataFrame()
        for sheet_name in target_sheets:
            try:
                try:
                    sheet = spreadsheet.worksheet(sheet_name)
                    # Removed debug output
                except gspread.exceptions.WorksheetNotFound:
                    st.warning(f"Sheet '{sheet_name}' not found in spreadsheet")
                    continue

                sheet_values = sheet.get_all_values()
                # Skip empty sheets
                if not sheet_values or len(sheet_values) <= 2:  # Need at least header row + column names + one data row
                    st.warning(f"Sheet '{sheet_name}' is empty or contains insufficient data")
                    continue
                headers = sheet_values[1]  # Second row as headers
                data = sheet_values[2:]
                df = pd.DataFrame(data, columns=headers)
                df['source_sheet'] = sheet_name
                # Append to the main DataFrame
                all_data = pd.concat([all_data, df], ignore_index=True)
                # Removed debug output
            except Exception as e:
                st.error(f"Error loading data from sheet {sheet_name}: {str(e)}")
                continue
        # Check if we got any data
        if all_data.empty:
            st.error("Failed to load any data from the spreadsheets")
            return pd.DataFrame()
        
        # Removed success message and columns listing
        return all_data
    except Exception as e:
        st.error(f"Error loading course mappings from Google Sheets: {str(e)}")
        return pd.DataFrame()
                    
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
First, extract the **Institution Name** from the transcript header or title section.
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
    "institution": "Langston University",
    "term": "Fall",
    "year": "2023",
    "courses": [
      {
        "course_code": "CS101",
        "division": "UNDG",
        "title": "Real-Time Text and voice output enabled traffic sign detection system using deep learning",
        "short_title": "Real-Time Traffic Sign Detection",
        "credits": 3,
        "grade": "A",
        "points": 12
      },
      {
        "course_code": "MATH202",
        "division": "UNDG",
        "title": "Calculus II",
        "short_title": "Calculus II",
        "credits": 4,
        "grade": "B+",
        "points": 13.2
      }
    ]
  },
  {
    "institution": "Langston University",
    "term": "Spring",
    "year": "2024",
    "courses": [
      {
        "course_code": "MATH5001",
        "division": "GRAD",
        "title": "Advanced Calculus",
        "short_title": "Advanced Calculus",
        "credits": 4,
        "grade": "A-",
        "points": 14.8
      }
    ]
  }
]
```

## **Additional Considerations**
- If "CRED" is missing, calculate credits using: CRED = Points/Grade where grade values are A=4.0, B=3.0, C=2.0, D=1.0, F=0.0
- Plus/minus modifiers adjust by 0.3 (e.g., A- = 3.7, B+ = 3.3)
- Ensure that each course is correctly associated with its respective term and year.
- Make sure to extract and include the "points" field in the output as it's needed for credit calculation.
- If any required information is missing from a course, leave the value as an empty string ("") rather than omitting the field.
- The institution name should be included at the term level in the JSON structure.
"""
def main():
    st.set_page_config(page_title="Transcript Analyzer", layout="wide")
    st.title("🔍 Academic Transcript Analyzer")
    
    # Initialize session state variables
    for key in ["pdf_processed", "feedback_submitted", "feedback_skipped", 
                "uploaded_file_name", "pdf_bytes", "drive_upload_status"]:
        if key not in st.session_state:
            st.session_state[key] = False if key in ["pdf_processed", "feedback_submitted", "feedback_skipped"] else None
    
    # Step 1: Ask for password
    if not check_password():
        st.warning("Please enter the password to access the transcript analyzer.")
        st.stop()  # Don't run the rest of the app until the correct password is entered

    st.success("Access granted. You may now upload and analyze transcripts.")
    
    # Show upload status from previous submission if available
    if st.session_state.get("drive_upload_status") == "success":
        st.success("Previous PDF was successfully saved to Google Drive.")
        # Clear the status to avoid showing it repeatedly
        st.session_state["drive_upload_status"] = None
    
    # Load institution mappings ONCE and store in session state
    if "institution_df" not in st.session_state:
        with st.spinner("Loading institution data..."):
            institution_df = load_institution_mappings()
            st.session_state["institution_df"] = institution_df
            if not institution_df.empty:
                st.success(f"Loaded institution mappings: {len(institution_df)} entries")
    
    # Always show the file uploader
    st.write("Upload a PDF transcript to extract course information.")
    uploaded_file = st.file_uploader("Choose a transcript PDF file", type="pdf")
    
    # Handle feedback dialog for previously processed PDF without blocking new uploads
    if st.session_state.get("pdf_processed", False) and not st.session_state.get("feedback_submitted", False) and not st.session_state.get("feedback_skipped", False):
        st.markdown("---")
        st.subheader("Feedback for Previous Analysis")
        feedback_result, feedback_text = show_feedback_dialog()
        
        if feedback_result == True:  # Feedback submitted
            st.session_state["feedback_submitted"] = True
            # After feedback is submitted, save the PDF to Google Drive
            if st.session_state.get("pdf_bytes") and st.session_state.get("uploaded_file_name"):
                success, message, file_url = save_pdf_to_drive(
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
        
        elif feedback_result == "skipped":  # Feedback skipped
            st.session_state["feedback_skipped"] = True
            st.info("Feedback skipped. You can process another transcript.")
            st.markdown("---")
    
    # Process the uploaded file (if any)
    if uploaded_file is not None:
        pdf_bytes = uploaded_file.getvalue()
        st.session_state["pdf_bytes"] = pdf_bytes
        st.session_state["uploaded_file_name"] = uploaded_file.name
        
        # Process the transcript button
        if st.button("Process Transcript"):
            # Reset feedback and processed states for new upload
            st.session_state["feedback_submitted"] = False
            st.session_state["feedback_skipped"] = False
            st.session_state["pdf_processed"] = False
            
            claude_response, token_usage = analyze_pdf(pdf_bytes, PROMPT)
            json_data = extract_json(claude_response)
            
            if json_data:
                json_data = post_process_transcript_data(json_data)
                macu_df = load_macu_mappings_from_sheets()
                ceqmacu_df = load_ceqmacu_mappings()
                json_data = enrich_with_macu_data(json_data, macu_df, ceqmacu_df)
                st.session_state["json_data"] = json_data
                
                st.success("Transcript processed successfully!")
                st.download_button(
                    label="Download JSON Data",
                    data=json.dumps(json_data, indent=4),
                    file_name=f"{uploaded_file.name.split('.')[0]}_processed.json",
                    mime="application/json"
                )
                
                # Display token usage details in an expander
                with st.expander("API Token Usage Details"):
                    st.markdown(token_usage)
                
                # Use the display function without passing institution_df
                # as it's now accessed from session state
                display_transcript_data(json_data)
                
                with st.expander("View Raw JSON Data"):
                    st.json(json_data)
                
                # Set the state to show that a PDF has been processed
                st.session_state["pdf_processed"] = True
            else:
                st.error("Failed to extract data from the transcript.")
                st.text(claude_response)
                
if __name__ == "__main__":
    main()
                
