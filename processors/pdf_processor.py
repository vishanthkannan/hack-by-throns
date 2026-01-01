"""
PDF Processor for NCRP Complaint Files
Extracts complaint and transaction data from PDF files with robust extraction
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


def extract_field(pattern: str, text: str, flags: int = re.IGNORECASE | re.DOTALL) -> Optional[str]:
    """
    Reusable helper function to extract field using regex pattern
    Handles line breaks, missing colons, and spacing issues
    """
    if not text or not pattern:
        return None
    
    try:
        match = re.search(pattern, text, flags)
        if match:
            # Get the captured group (first group if multiple)
            if match.groups():
                result = match.group(1).strip()
            else:
                result = match.group(0).strip()
            
            # Clean up common issues
            result = re.sub(r'\s+', ' ', result)  # Normalize whitespace
            result = result.replace('\n', ' ').replace('\r', ' ')
            
            return result if result else None
    except Exception as e:
        pass
    
    return None


def parse_date(date_str: str) -> str:
    """Parse date string to YYYY-MM-DD format with multiple format support"""
    if not date_str or date_str.strip() == "":
        return ""
    
    date_str = date_str.strip()
    
    # Common date formats
    date_formats = [
        '%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d',
        '%d/%m/%y', '%d-%m-%y',
        '%d %B %Y', '%d %b %Y',
        '%B %d, %Y', '%b %d, %Y',
        '%d/%m/%Y %I:%M %p',  # With time
        '%d-%m-%Y %I:%M %p',
    ]
    
    for fmt in date_formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime('%Y-%m-%d')
        except:
            continue
    
    # Try regex patterns for flexible matching
    patterns = [
        r'(\d{1,2})[/-](\d{1,2})[/-](\d{4})',
        r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})',
        r'(\d{1,2})\s+(\w+)\s+(\d{4})',  # DD Month YYYY
    ]
    
    for pattern in patterns:
        match = re.search(pattern, date_str)
        if match:
            parts = match.groups()
            if len(parts) == 3:
                try:
                    if len(parts[2]) == 4:  # DD/MM/YYYY
                        dt = datetime(int(parts[2]), int(parts[1]), int(parts[0]))
                    elif len(parts[0]) == 4:  # YYYY/MM/DD
                        dt = datetime(int(parts[0]), int(parts[1]), int(parts[2]))
                    else:
                        continue
                    return dt.strftime('%Y-%m-%d')
                except:
                    continue
    
    return ""




def extract_complaint_id(text: str) -> str:
    """
    Extract complaint/acknowledgement number with robust patterns
    Handles cases where number appears on next line
    """
    # Strong patterns for Complaint ID / Acknowledgement Number
    patterns = [
        # With label and colon (may span lines)
        r'(?:Complaint\s+ID|Acknowledgment\s+Number|Acknowledgment\s+No|Ack\s+Number|Ack\s+No)[\s:]*[\n\r]?[\s:]*([A-Z0-9/-]{8,})',
        # Without colon, number on same or next line
        r'(?:Complaint\s+ID|Acknowledgment\s+Number|Ack\s+Number)[\s]*[\n\r]?\s*([A-Z0-9/-]{8,})',
        # Pattern: Letters-Digits-Digits (e.g., NCRP-2024-123456)
        r'([A-Z]{2,4}[-/]?\d{4,}[-/]?\d{4,})',
        # Long numeric IDs (10+ digits)
        r'\b(\d{10,})\b',
        # Alphanumeric patterns
        r'([A-Z]{2,}\d{6,})',
    ]
    
    for pattern in patterns:
        result = extract_field(pattern, text, re.IGNORECASE | re.DOTALL | re.MULTILINE)
        if result and len(result) >= 8:
            return result
    
    return ""


def extract_category(text: str) -> str:
    """
    Extract Category of Complaint
    Text may appear without colon
    """
    patterns = [
        r'(?:Category\s+of\s+Complaint|Complaint\s+Category|Category)[\s:]*[\n\r]?[\s:]*([A-Za-z\s]+?)(?:\n|Sub\s+Category|$)',
        r'(?:Category\s+of\s+Complaint|Complaint\s+Category)[\s]*[\n\r]?\s*([A-Za-z\s]{3,})',
        # Look for common categories in text
        r'\b(Fraud|Cyber\s+Crime|Financial\s+Crime|Online\s+Fraud|Banking\s+Fraud|UPI\s+Fraud|Phishing|Scam|Identity\s+Theft|Social\s+Media\s+Fraud)\b',
    ]
    
    for pattern in patterns:
        result = extract_field(pattern, text, re.IGNORECASE | re.DOTALL)
        if result:
            return result.title()
    
    return ""


def extract_sub_category(text: str) -> str:
    """
    Extract Sub Category of Complaint
    Often appears on next line without colon
    """
    patterns = [
        # With label, may span lines
        r'(?:Sub\s+Category\s+of\s+Complaint|Sub\s+Category|Subcategory)[\s:]*[\n\r]?[\s:]*([A-Za-z\s]+?)(?:\n|District|State|Amount|$)',
        # Without colon, on next line
        r'(?:Sub\s+Category\s+of\s+Complaint|Sub\s+Category)[\s]*[\n\r]?\s*([A-Za-z\s]{3,})',
        # Common sub-categories
        r'\b(UPI\s+Fraud|Online\s+Transaction\s+Fraud|Credit\s+Card\s+Fraud|Debit\s+Card\s+Fraud|Bank\s+Transfer\s+Fraud|Investment\s+Scam|Job\s+Scam|Loan\s+Scam)\b',
    ]
    
    for pattern in patterns:
        result = extract_field(pattern, text, re.IGNORECASE | re.DOTALL)
        if result:
            return result.title()
    
    return ""


def extract_complaint_date(text: str) -> str:
    """Extract Complaint Date"""
    patterns = [
        r'(?:Complaint\s+Date|Date\s+of\s+Complaint|Filed\s+Date)[\s:]*[\n\r]?[\s:]*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})',
        r'(?:Complaint\s+Date|Date\s+of\s+Complaint)[\s]*[\n\r]?\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})',
    ]
    
    for pattern in patterns:
        result = extract_field(pattern, text, re.IGNORECASE | re.DOTALL)
        if result:
            parsed = parse_date(result)
            if parsed:
                return parsed
    
    return ""


def extract_incident_date(text: str) -> str:
    """Extract Incident Date/Time (may include AM/PM and line breaks)"""
    patterns = [
        r'(?:Incident\s+Date|Date\s+of\s+Incident|Occurred\s+Date|Date\s+of\s+Occurrence)[\s:]*[\n\r]?[\s:]*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4}(?:\s+\d{1,2}:\d{2}(?:\s+[AP]M)?)?)',
        r'(?:Incident\s+Date|Date\s+of\s+Incident)[\s]*[\n\r]?\s*(\d{1,2}[/-]\d{1,2}[/-]\d{2,4})',
    ]
    
    for pattern in patterns:
        result = extract_field(pattern, text, re.IGNORECASE | re.DOTALL)
        if result:
            # Remove time portion if present
            date_part = result.split()[0] if ' ' in result else result
            parsed = parse_date(date_part)
            if parsed:
                return parsed
    
    return ""


def extract_amount(text: str) -> float:
    """
    Extract Total Fraudulent Amount reported by complainant
    Must capture numeric value with commas and decimals
    """
    # Strong patterns for amount extraction
    patterns = [
        # With label "Total Fraudulent Amount"
        r'(?:Total\s+Fraudulent\s+Amount|Amount\s+Lost|Fraudulent\s+Amount|Amount\s+Reported)[\s:]*[\n\r]?[\s:]*[₹Rs\.\s]*([\d,]+\.?\d*)',
        # Currency symbol patterns
        r'₹\s*([\d,]+\.?\d*)',
        r'Rs\.?\s*([\d,]+\.?\d*)',
        r'INR\s*([\d,]+\.?\d*)',
        # Amount with "rupees" or similar
        r'([\d,]+\.?\d*)\s*(?:rupees|Rs|₹|INR)',
    ]
    
    amounts = []
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE | re.DOTALL)
        for match in matches:
            try:
                amount_str = str(match).replace(',', '')
                amount = float(amount_str)
                if amount > 0:  # Only valid positive amounts
                    amounts.append(amount)
            except:
                continue
    
    # Return the largest amount found (most likely the total)
    return max(amounts) if amounts else 0.0


def extract_district(text: str) -> str:
    """Extract District"""
    patterns = [
        r'(?:District|City)[\s:]*[\n\r]?[\s:]*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)',
        r'(?:District|City)[\s]*[\n\r]?\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)',
    ]
    
    for pattern in patterns:
        result = extract_field(pattern, text, re.IGNORECASE)
        if result:
            return result
    
    return ""


def extract_state(text: str) -> str:
    """Extract State"""
    patterns = [
        r'(?:State|Province)[\s:]*[\n\r]?[\s:]*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)',
        r'(?:State|Province)[\s]*[\n\r]?\s*([A-Z][a-z]+(?:\s+[A-Z][a-z]+)*)',
    ]
    
    for pattern in patterns:
        result = extract_field(pattern, text, re.IGNORECASE)
        if result:
            return result
    
    return ""


def extract_transaction_ids(text: str) -> List[str]:
    """Extract transaction IDs/UTR numbers"""
    patterns = [
        r'(?:UTR|Transaction\s+ID|Transaction\s+Number|Txn\s+ID|Ref\s+Number)[\s:]*[\n\r]?[\s:]*([A-Z0-9]{8,})',
        r'([A-Z]{2,}\d{10,})',  # Alphanumeric patterns
        r'\b(\d{12,})\b',  # Long numeric IDs
    ]
    
    transactions = []
    for pattern in patterns:
        matches = re.findall(pattern, text, re.IGNORECASE | re.DOTALL)
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
    patterns = [
        r'(?:Status|Complaint\s+Status)[\s:]*[\n\r]?[\s:]*([A-Za-z\s]+?)(?:\n|$)',
    ]
    
    result = extract_field(patterns[0], text, re.IGNORECASE | re.DOTALL)
    if result:
        return result.title()
    
    return "Pending"


def process_pdf(filepath: str) -> List[Dict]:
    """
    Process PDF file and extract NCRP complaint data with robust extraction
    Returns list of complaint dictionaries with normalized data
    """
    try:
        text = extract_text_from_pdf(filepath)
        
        if not text or len(text.strip()) < 50:
            return []
        
        # Extract all fields using robust patterns
        complaint_id = extract_complaint_id(text)
        category = extract_category(text)
        sub_category = extract_sub_category(text)
        complaint_date = extract_complaint_date(text)
        incident_date = extract_incident_date(text)
        amount = extract_amount(text)
        district = extract_district(text)
        state = extract_state(text)
        transactions = extract_transaction_ids(text)
        bank_info = extract_bank_platform_info(text)
        status = extract_status(text)
        
        # Normalize all fields using shared normalizer
        complaint_id = normalize_complaint_id(complaint_id) if complaint_id else "Not Available"
        category = normalize_string(category) if category else "Not Available"
        sub_category = normalize_string(sub_category) if sub_category else "Not Available"
        complaint_date = complaint_date if complaint_date else datetime.now().strftime('%Y-%m-%d')
        incident_date = incident_date if incident_date else complaint_date
        amount = normalize_amount(amount)
        district = normalize_string(district) if district else "Not Available"
        state = normalize_string(state) if state else "Not Available"
        status = normalize_string(status) if status else "Pending"
        bank_info = normalize_string(bank_info) if bank_info else "Not Available"
        
        # Generate Complaint_ID if still missing
        if complaint_id == "Not Available":
            complaint_id = f"PDF_{datetime.now().strftime('%Y%m%d%H%M%S')}"
        
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
