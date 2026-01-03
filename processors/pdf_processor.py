"""
PDF Processor for NCRP Complaint Files
Extracts complaint and transaction data from PDF files
FINAL FIX: Label-anchored, non-greedy regex patterns
"""

import pdfplumber
import re
from datetime import datetime
from typing import List, Dict, Optional
from processors.normalizer import normalize_string, normalize_amount, normalize_complaint_id


def extract_text_from_pdf(filepath: str) -> str:
    """Extract all text from PDF file, handling all pages"""
    text = ""
    try:
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        raise Exception(f"Error reading PDF: {str(e)}")
    return text


def normalize_text(text: str) -> str:
    """
    Normalize PDF text line-by-line
    This ensures predictable line breaks for regex matching
    """
    # Replace \r with \n
    text = text.replace("\r", "\n")
    # Normalize multiple newlines to single newline
    text = re.sub(r"\n+", "\n", text)
    return text


def extract_section(text: str, start_marker: str, end_marker: str) -> str:
    """
    Extract text section between two markers
    Returns empty string if markers not found
    """
    try:
        # Case-insensitive search for start marker
        start_pattern = re.escape(start_marker)
        start_match = re.search(start_pattern, text, re.IGNORECASE)
        
        if not start_match:
            return ""
        
        start_pos = start_match.end()
        
        # Case-insensitive search for end marker after start
        end_pattern = re.escape(end_marker)
        end_match = re.search(end_pattern, text[start_pos:], re.IGNORECASE)
        
        if not end_match:
            # If end marker not found, take until end of text
            return text[start_pos:]
        
        end_pos = start_pos + end_match.start()
        return text[start_pos:end_pos]
    
    except Exception:
        return ""


def extract_field(pattern: str, text: str) -> Optional[str]:
    """
    Safe helper function to extract field using regex pattern
    NO re.DOTALL - matching must stop at line end
    """
    if not text or not pattern:
        return None
    
    try:
        match = re.search(pattern, text, re.IGNORECASE)
        if match and match.groups():
            result = match.group(1).strip()
            # Stop at newline if present
            if '\n' in result:
                result = result.split('\n')[0].strip()
            return result if result else None
    except Exception as e:
        pass
    
    return None


def parse_ncrp_date(date_str: str) -> str:
    """
    Parse NCRP date from DD/MM/YYYY format to YYYY-MM-DD
    NCRP dates are ALWAYS in DD/MM/YYYY format
    Returns "Not Available" if parsing fails
    """
    if not date_str:
        return "Not Available"
    
    date_str = date_str.strip()
    
    if not date_str:
        return "Not Available"
    
    try:
        # NCRP dates are ALWAYS in DD/MM/YYYY format
        dt = datetime.strptime(date_str, "%d/%m/%Y")
        return dt.strftime("%Y-%m-%d")
    except:
        return "Not Available"


def extract_complaint_id(text: str) -> str:
    """
    Extract Complaint ID (Acknowledgement Number)
    MUST capture only digits after the label - NOT transaction references
    """
    # Primary pattern: Acknowledgement Number with digits only
    pattern = r"Acknowledgement\s*Number\s*:\s*\n?\s*(\d{10,})"
    result = extract_field(pattern, text)
    if result:
        return result
    
    # Fallback: Acknowledgement Number without colon
    pattern = r"Acknowledgement\s*Number\s*\n\s*(\d{10,})"
    result = extract_field(pattern, text)
    if result:
        return result
    
    # Last resort: Complaint ID with digits only (not alphanumeric)
    pattern = r"Complaint\s*ID\s*:\s*\n?\s*(\d{10,})"
    result = extract_field(pattern, text)
    if result:
        return result
    
    return ""


def extract_category(text: str) -> str:
    """
    Extract Category of Complaint
    Captures only the immediate value on next line
    """
    pattern = r"Category\s+of\s+complaint\s*\n\s*([A-Za-z ]+)"
    result = extract_field(pattern, text)
    if result:
        # Stop at newline or next label
        result = result.split('\n')[0].strip()
        return result.title()
    
    return ""


def extract_sub_category(text: str) -> str:
    """
    Extract Sub Category of Complaint
    Captures only the immediate value on next line
    """
    pattern = r"Sub\s+Category\s+of\s+Complaint\s*\n\s*([A-Za-z ]+)"
    result = extract_field(pattern, text)
    if result:
        # Stop at newline or next label
        result = result.split('\n')[0].strip()
        return result.title()
    
    return ""


def extract_incident_date(text: str) -> str:
    """
    Extract Incident Date/Time
    Extract as RAW STRING in DD/MM/YYYY format
    NCRP dates are ALWAYS in DD/MM/YYYY format
    """
    # Extract DD/MM/YYYY format explicitly
    pattern = r"Incident\s+Date\/Time\s*\n\s*([\d]{2}/[\d]{2}/[\d]{4})"
    result = extract_field(pattern, text)
    if result:
        # Parse the date explicitly
        return parse_ncrp_date(result)
    
    # Fallback: Incident Date without time
    pattern = r"Incident\s+Date\s*\n\s*([\d]{2}/[\d]{2}/[\d]{4})"
    result = extract_field(pattern, text)
    if result:
        return parse_ncrp_date(result)
    
    return ""


def extract_complaint_date(text: str) -> str:
    """
    Extract Complaint Date
    Extract as RAW STRING in DD/MM/YYYY format
    NCRP dates are ALWAYS in DD/MM/YYYY format
    """
    # Extract DD/MM/YYYY format explicitly
    pattern = r"Complaint\s+Date\s*\n\s*([\d]{2}/[\d]{2}/[\d]{4})"
    result = extract_field(pattern, text)
    if result:
        # Parse the date explicitly
        return parse_ncrp_date(result)
    
    return ""


def extract_amount(text: str) -> float:
    """
    Extract Total Fraudulent Amount reported by complainant
    Must capture numeric value with commas and decimals
    """
    pattern = r"Total\s+Fraudulent\s+Amount\s+reported\s+by\s+complainant\s*:\s*([\d,\.]+)"
    result = extract_field(pattern, text)
    if result:
        try:
            amount_str = result.replace(',', '')
            amount = float(amount_str)
            return amount if amount > 0 else 0.0
        except:
            pass
    
    # Fallback: Look for amount with currency symbol after label
    pattern = r"Total\s+Fraudulent\s+Amount\s+reported\s+by\s+complainant\s*:\s*[₹Rs\.\s]*([\d,\.]+)"
    result = extract_field(pattern, text)
    if result:
        try:
            amount_str = result.replace(',', '')
            amount = float(amount_str)
            return amount if amount > 0 else 0.0
        except:
            pass
    
    return 0.0


def extract_district(text: str) -> str:
    """
    Extract District
    CRITICAL FIX: MUST capture only the next line text, not paragraphs
    """
    pattern = r"District\s*\n\s*([A-Za-z ]+)"
    result = extract_field(pattern, text)
    if result:
        # Stop at newline - only take first line
        result = result.split('\n')[0].strip()
        # Remove any trailing labels that might have been captured
        # Stop if we see common next labels
        stop_words = ['State', 'Amount', 'Complaint', 'Category', 'Sub']
        for word in stop_words:
            if word in result:
                result = result.split(word)[0].strip()
        return result
    
    return ""


def extract_state(text: str) -> str:
    """
    Extract State
    CRITICAL FIX: MUST capture only the next line text, not paragraphs
    """
    pattern = r"State\s*\n\s*([A-Za-z ]+)"
    result = extract_field(pattern, text)
    if result:
        # Stop at newline - only take first line
        result = result.split('\n')[0].strip()
        # Remove any trailing labels that might have been captured
        # Stop if we see common next labels
        stop_words = ['District', 'Amount', 'Complaint', 'Category', 'Sub']
        for word in stop_words:
            if word in result:
                result = result.split(word)[0].strip()
        return result
    
    return ""


def extract_transaction_ids(text: str) -> List[str]:
    """
    Extract transaction IDs/UTR numbers
    Only extract from transaction-specific sections, not from complaint ID area
    """
    transactions = []
    
    # Look for UTR/Transaction ID labels specifically
    patterns = [
        r"UTR\s*Number\s*:\s*\n?\s*([A-Z0-9]{8,})",
        r"Transaction\s*ID\s*:\s*\n?\s*([A-Z0-9]{8,})",
        r"Transaction\s*Reference\s*:\s*\n?\s*([A-Z0-9]{8,})",
    ]
    
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE)
        for match in matches:
            trans_id = str(match).strip()
            if len(trans_id) >= 8:
                transactions.append(trans_id)
    
    # Remove duplicates
    return list(set(transactions))


def extract_bank_platform_info(text: str) -> str:
    """Extract bank or platform information"""
    banks = ['SBI', 'HDFC', 'ICICI', 'Axis', 'Kotak', 'PNB', 'BOI', 'Canara', 
             'Paytm', 'PhonePe', 'GPay', 'Google Pay', 'UPI', 'NEFT', 'RTGS', 'IMPS',
             'Bank of India', 'State Bank', 'Punjab National Bank']
    
    found = []
    text_upper = text.upper()
    for bank in banks:
        if bank.upper() in text_upper:
            found.append(bank)
    
    return ", ".join(found) if found else ""


def extract_status(text: str) -> str:
    """Extract complaint status"""
    pattern = r"Status\s*:\s*\n?\s*([A-Za-z ]+)"
    result = extract_field(pattern, text)
    if result:
        # Stop at newline
        result = result.split('\n')[0].strip()
        return result.title()
    
    return "Pending"


def process_pdf(filepath: str) -> List[Dict]:
    """
    Process PDF file and extract NCRP complaint data
    Uses label-anchored, non-greedy regex patterns
    Returns list of complaint dictionaries with normalized data
    """
    try:
        text = extract_text_from_pdf(filepath)
        
        if not text or len(text.strip()) < 50:
            return []
        
        # STEP 1: Normalize text line-by-line (MANDATORY)
        text = normalize_text(text)
        
        # STEP 2: Isolate sections for targeted extraction
        # Extract complaint header section (between "Complaint Type" and "Complainant Details")
        header_text = extract_section(text, "Complaint Type", "Complainant Details")
        
        # Extract complainant/location section (between "Complainant Details" and "Suspect Details")
        location_text = extract_section(text, "Complainant Details", "Suspect Details")
        
        # STEP 3: Apply regex ONLY on correct sections
        # Extract from header_text: Complaint_ID, Category, Sub_Category
        complaint_id_raw = extract_complaint_id(header_text)
        
        # MANDATORY GUARDRAIL: Validate Complaint_ID before proceeding
        # A complaint row MUST be created ONLY IF:
        # - Acknowledgement Number exists
        # - Acknowledgement Number contains DIGITS ONLY
        # - Length >= 10
        if not complaint_id_raw or not complaint_id_raw.isdigit() or len(complaint_id_raw) < 10:
            # DO NOT create a complaint row - skip the record completely
            return []
        
        # Extract from header_text: Category, Sub_Category
        category = extract_category(header_text)
        sub_category = extract_sub_category(header_text)
        
        # Extract from location_text: District, State
        district = extract_district(location_text)
        state = extract_state(location_text)
        
        # Extract dates from FULL TEXT (dates may be anywhere in PDF)
        complaint_date = extract_complaint_date(text)
        incident_date = extract_incident_date(text)
        
        # Extract other fields from full text (may be in various sections)
        amount = extract_amount(text)
        transactions = extract_transaction_ids(text)
        bank_info = extract_bank_platform_info(text)
        status = extract_status(text)
        
        # Normalize Complaint_ID (already validated as digits only)
        complaint_id = normalize_complaint_id(complaint_id_raw)
        category = normalize_string(category) if category else "Not Available"
        sub_category = normalize_string(sub_category) if sub_category else "Not Available"
        
        # Dates are already parsed - use "Not Available" if missing (NO datetime.now fallback)
        complaint_date = complaint_date if complaint_date else "Not Available"
        incident_date = incident_date if incident_date else "Not Available"
        amount = normalize_amount(amount)
        district = normalize_string(district) if district else "Not Available"
        state = normalize_string(state) if state else "Not Available"
        status = normalize_string(status) if status else "Pending"
        bank_info = normalize_string(bank_info) if bank_info else "Not Available"
        
        # Complaint_ID is already validated and normalized
        # No need to generate fallback ID - validation ensures it exists
        
        # Create complaint record with all normalized fields
        complaint = {
            'Complaint_ID': complaint_id,
            'Complaint_Date': complaint_date,
            'Incident_Date': incident_date,
            'Category': category,
            'Sub_Category': sub_category,
            'District': district,
            'State': state,
            'Amount_Lost': amount,
            'Status': status,
            'Transaction_Count': len(transactions),
            'Transaction_IDs': transactions,
            'Bank_Platform_Info': bank_info
        }
        
        return [complaint]
    
    except Exception as e:
        raise Exception(f"Error processing PDF: {str(e)}")
