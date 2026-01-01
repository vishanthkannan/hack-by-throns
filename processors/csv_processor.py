"""
CSV Processor for NCRP Complaint Files
Processes CSV exports from NCRP system
"""

import pandas as pd
from datetime import datetime
from typing import List, Dict
import re
from processors.normalizer import normalize_string, normalize_amount, normalize_complaint_id


def normalize_column_name(col: str) -> str:
    """Normalize column names to standard format"""
    col = str(col).strip()
    col_lower = col.lower()
    
    # Map common variations
    mappings = {
        'complaint id': 'Complaint_ID',
        'acknowledgement number': 'Complaint_ID',
        'ack number': 'Complaint_ID',
        'complaint number': 'Complaint_ID',
        'complaint date': 'Complaint_Date',
        'date of complaint': 'Complaint_Date',
        'filed date': 'Complaint_Date',
        'incident date': 'Incident_Date',
        'date of incident': 'Incident_Date',
        'occurred date': 'Incident_Date',
        'category': 'Category',
        'complaint category': 'Category',
        'sub category': 'Sub_Category',
        'subcategory': 'Sub_Category',
        'district': 'District',
        'state': 'State',
        'amount': 'Amount_Lost',
        'fraudulent amount': 'Amount_Lost',
        'total amount': 'Amount_Lost',
        'amount lost': 'Amount_Lost',
        'status': 'Status',
        'complaint status': 'Status',
        'transaction id': 'Transaction_IDs',
        'transaction': 'Transaction_IDs',
        'utr': 'Transaction_IDs',
        'bank': 'Bank_Platform_Info',
        'platform': 'Bank_Platform_Info',
    }
    
    return mappings.get(col_lower, col)


def parse_date(date_value) -> str:
    """Parse date to YYYY-MM-DD format"""
    if pd.isna(date_value) or date_value == "":
        return ""
    
    date_str = str(date_value).strip()
    
    if not date_str:
        return ""
    
    # Try pandas parsing first
    try:
        dt = pd.to_datetime(date_str)
        return dt.strftime('%Y-%m-%d')
    except:
        pass
    
    # Try manual parsing
    date_formats = [
        '%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d',
        '%d/%m/%y', '%d-%m-%y',
        '%d %B %Y', '%d %b %Y',
    ]
    
    for fmt in date_formats:
        try:
            dt = datetime.strptime(date_str, fmt)
            return dt.strftime('%Y-%m-%d')
        except:
            continue
    
    return ""


def parse_amount(amount_value) -> float:
    """Parse amount to float"""
    if pd.isna(amount_value):
        return 0.0
    
    amount_str = str(amount_value).strip()
    
    # Remove currency symbols and commas
    amount_str = re.sub(r'[₹,Rs\.\s]', '', amount_str)
    
    try:
        return float(amount_str)
    except:
        return 0.0


def extract_transactions(transaction_value) -> List[str]:
    """Extract transaction IDs from cell value"""
    if pd.isna(transaction_value):
        return []
    
    trans_str = str(transaction_value).strip()
    
    # Split by common delimiters
    transactions = re.split(r'[,;|\n]+', trans_str)
    
    # Clean and filter
    transactions = [t.strip() for t in transactions if len(t.strip()) >= 8]
    
    return transactions


def process_csv(filepath: str) -> List[Dict]:
    """
    Process CSV file and extract NCRP complaint data
    Returns list of complaint dictionaries
    """
    try:
        # Read CSV with flexible encoding
        encodings = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252']
        df = None
        
        for encoding in encodings:
            try:
                df = pd.read_csv(filepath, encoding=encoding)
                break
            except:
                continue
        
        if df is None:
            raise Exception("Could not read CSV file with any encoding")
        
        if df.empty:
            return []
        
        # Normalize column names
        df.columns = [normalize_column_name(col) for col in df.columns]
        
        complaints = []
        
        for idx, row in df.iterrows():
            complaint = {}
            
            # Extract standard fields
            complaint['Complaint_ID'] = str(row.get('Complaint_ID', '')).strip() if pd.notna(row.get('Complaint_ID')) else ""
            complaint['Complaint_Date'] = parse_date(row.get('Complaint_Date', ''))
            complaint['Incident_Date'] = parse_date(row.get('Incident_Date', ''))
            complaint['Category'] = str(row.get('Category', '')).strip() if pd.notna(row.get('Category')) else ""
            complaint['Sub_Category'] = str(row.get('Sub_Category', '')).strip() if pd.notna(row.get('Sub_Category')) else ""
            complaint['District'] = str(row.get('District', '')).strip() if pd.notna(row.get('District')) else ""
            complaint['State'] = str(row.get('State', '')).strip() if pd.notna(row.get('State')) else ""
            complaint['Amount_Lost'] = parse_amount(row.get('Amount_Lost', 0))
            complaint['Status'] = str(row.get('Status', 'Pending')).strip() if pd.notna(row.get('Status')) else "Pending"
            
            # Extract transactions
            transactions = extract_transactions(row.get('Transaction_IDs', ''))
            complaint['Transaction_Count'] = len(transactions)
            complaint['Transaction_IDs'] = transactions
            
            # Extract bank/platform info
            bank_info = str(row.get('Bank_Platform_Info', '')).strip() if pd.notna(row.get('Bank_Platform_Info')) else ""
            complaint['Bank_Platform_Info'] = bank_info
            
            # Normalize all fields
            complaint['Complaint_ID'] = normalize_complaint_id(complaint.get('Complaint_ID', ''))
            complaint['Category'] = normalize_string(complaint.get('Category', ''))
            complaint['Sub_Category'] = normalize_string(complaint.get('Sub_Category', ''))
            complaint['District'] = normalize_string(complaint.get('District', ''))
            complaint['State'] = normalize_string(complaint.get('State', ''))
            complaint['Amount_Lost'] = normalize_amount(complaint.get('Amount_Lost', 0))
            complaint['Status'] = normalize_string(complaint.get('Status', 'Pending'))
            complaint['Bank_Platform_Info'] = normalize_string(complaint.get('Bank_Platform_Info', ''))
            
            # Generate Complaint_ID if missing
            if complaint['Complaint_ID'] == "Not Available":
                complaint['Complaint_ID'] = f"CSV_{datetime.now().strftime('%Y%m%d%H%M%S')}_{idx}"
            
            # Set default dates if missing
            if not complaint.get('Complaint_Date', ''):
                complaint['Complaint_Date'] = datetime.now().strftime('%Y-%m-%d')
            
            if not complaint.get('Incident_Date', ''):
                complaint['Incident_Date'] = complaint['Complaint_Date']
            
            complaints.append(complaint)
        
        return complaints
    
    except Exception as e:
        raise Exception(f"Error processing CSV: {str(e)}")

