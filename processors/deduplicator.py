"""
Deduplicator Module
Handles intelligence features and master Excel file management
WITH SAFE EXCEL WRITING to prevent type inference and scientific notation
"""

import pandas as pd
import os
from datetime import datetime
from typing import List, Dict, Tuple
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
from processors.normalizer import normalize_string, normalize_amount, normalize_complaint_id


def calculate_data_quality_score(complaint: Dict) -> int:
    """
    Feature 1: Complaint Normalization Score
    Add 20 points for each:
    - Complaint ID present
    - Complaint Date present
    - Amount Lost present
    - District + State present
    - At least one transaction present
    """
    score = 0
    
    complaint_id = str(complaint.get('Complaint_ID', '')).strip()
    if complaint_id and complaint_id != "Not Available":
        score += 20
    
    complaint_date = str(complaint.get('Complaint_Date', '')).strip()
    if complaint_date and complaint_date != "Not Available":
        score += 20
    
    amount = complaint.get('Amount_Lost', 0)
    if isinstance(amount, (int, float)) and amount > 0:
        score += 20
    
    district = str(complaint.get('District', '')).strip()
    state = str(complaint.get('State', '')).strip()
    if district and district != "Not Available" and state and state != "Not Available":
        score += 20
    
    transaction_count = complaint.get('Transaction_Count', 0)
    if isinstance(transaction_count, (int, float)) and transaction_count > 0:
        score += 20
    
    return score


def calculate_investigation_readiness(complaint: Dict) -> str:
    """
    Feature 2: Investigation Readiness Flag
    Set Investigation_Ready = YES if:
    - Amount Lost > 0
    - At least one transaction ID exists
    - Bank / Platform info exists
    Else: NO
    """
    amount = complaint.get('Amount_Lost', 0)
    amount_ok = isinstance(amount, (int, float)) and amount > 0
    
    transaction_count = complaint.get('Transaction_Count', 0)
    transaction_ok = isinstance(transaction_count, (int, float)) and transaction_count > 0
    
    bank_info = str(complaint.get('Bank_Platform_Info', '')).strip()
    bank_ok = bank_info and bank_info != "Not Available"
    
    if amount_ok and transaction_ok and bank_ok:
        return "YES"
    else:
        return "NO"


def calculate_reporting_delay(complaint: Dict) -> Tuple[int, str]:
    """
    Feature 3: Reporting Delay Indicator
    Compute Reporting_Delay_Days = Complaint_Date - Incident_Date
    If delay > 7 days: DELAYED, Else: ON_TIME
    """
    complaint_date_str = str(complaint.get('Complaint_Date', '')).strip()
    incident_date_str = str(complaint.get('Incident_Date', '')).strip()
    
    if not complaint_date_str or complaint_date_str == "Not Available":
        return 0, "UNKNOWN"
    
    if not incident_date_str or incident_date_str == "Not Available":
        return 0, "UNKNOWN"
    
    try:
        complaint_date = pd.to_datetime(complaint_date_str)
        incident_date = pd.to_datetime(incident_date_str)
        
        delay_days = (complaint_date - incident_date).days
        
        if delay_days > 7:
            status = "DELAYED"
        else:
            status = "ON_TIME"
        
        return delay_days, status
    except:
        return 0, "UNKNOWN"


def calculate_transaction_pattern(complaint: Dict) -> str:
    """
    Feature 4: Transaction Pattern Flag
    - One transaction AND amount > ₹50,000 → SINGLE_LARGE
    - Multiple transactions AND each < ₹10,000 → MULTIPLE_SMALL
    - Else → MIXED
    """
    transaction_count = complaint.get('Transaction_Count', 0)
    amount = complaint.get('Amount_Lost', 0)
    
    if not isinstance(transaction_count, (int, float)):
        transaction_count = 0
    if not isinstance(amount, (int, float)):
        amount = 0
    
    if transaction_count == 1 and amount > 50000:
        return "SINGLE_LARGE"
    elif transaction_count > 1:
        # Check if each transaction is small
        avg_per_transaction = amount / transaction_count if transaction_count > 0 else 0
        if avg_per_transaction < 10000:
            return "MULTIPLE_SMALL"
        else:
            return "MIXED"
    else:
        return "MIXED"


def apply_intelligence_features(complaints: List[Dict]) -> List[Dict]:
    """Apply all intelligence features to complaint data"""
    enhanced_complaints = []
    
    for complaint in complaints:
        enhanced = complaint.copy()
        
        # Feature 1: Data Quality Score
        enhanced['Data_Quality_Score'] = calculate_data_quality_score(complaint)
        
        # Feature 2: Investigation Readiness
        enhanced['Investigation_Ready'] = calculate_investigation_readiness(complaint)
        
        # Feature 3: Reporting Delay
        delay_days, delay_status = calculate_reporting_delay(complaint)
        enhanced['Reporting_Delay_Days'] = delay_days
        enhanced['Reporting_Delay_Status'] = delay_status
        
        # Feature 4: Transaction Pattern
        enhanced['Transaction_Pattern'] = calculate_transaction_pattern(complaint)
        
        enhanced_complaints.append(enhanced)
    
    return enhanced_complaints


def safe_write_excel(df: pd.DataFrame, filepath: str):
    """
    Safely write DataFrame to Excel with explicit type formatting
    Prevents scientific notation and type inference issues
    """
    # First, write using pandas (creates the file structure)
    df.to_excel(filepath, index=False, engine='openpyxl')
    
    # Now open with openpyxl to apply explicit formatting
    wb = load_workbook(filepath)
    ws = wb.active
    
    # Define text style for string columns
    text_style = NamedStyle(name="text_style", number_format='@')
    
    # Columns that MUST be text (to prevent Excel auto-conversion)
    text_columns = ['Complaint_ID']
    
    # Find column indices
    header_row = 1
    for col_idx, header in enumerate(df.columns, start=1):
        col_letter = ws.cell(row=header_row, column=col_idx).column_letter
        
        # Apply text format to Complaint_ID column
        if header in text_columns:
            for row_idx in range(2, len(df) + 2):  # Start from row 2 (skip header)
                cell = ws.cell(row=row_idx, column=col_idx)
                # Force as string
                if cell.value is not None:
                    cell.value = str(cell.value)
                cell.number_format = '@'  # Text format
    
    # Format Amount_Lost column to prevent scientific notation
    if 'Amount_Lost' in df.columns:
        amount_col_idx = list(df.columns).index('Amount_Lost') + 1
        amount_col_letter = ws.cell(row=header_row, column=amount_col_idx).column_letter
        
        for row_idx in range(2, len(df) + 2):
            cell = ws.cell(row=row_idx, column=amount_col_idx)
            if cell.value is not None:
                try:
                    # Ensure it's a number, format as number with 2 decimals
                    amount = float(cell.value)
                    cell.value = amount
                    cell.number_format = '#,##0.00'  # Number format with commas
                except:
                    pass
    
    # Save the workbook
    wb.save(filepath)


def append_to_master_excel(complaints: List[Dict], source_file_type: str) -> Tuple[int, int]:
    """
    Append complaints to master Excel file with deduplication
    Uses safe Excel writing to prevent type inference and scientific notation
    Returns: (new_count, total_count)
    """
    master_file = 'output/ncrp_master.xlsx'
    
    # Apply intelligence features
    enhanced_complaints = apply_intelligence_features(complaints)
    
    # Prepare data for DataFrame with explicit normalization
    records = []
    for complaint in enhanced_complaints:
        # Normalize all fields before creating record
        complaint_id = normalize_complaint_id(complaint.get('Complaint_ID', ''))
        complaint_date = normalize_string(complaint.get('Complaint_Date', ''))
        category = normalize_string(complaint.get('Category', ''))
        sub_category = normalize_string(complaint.get('Sub_Category', ''))
        district = normalize_string(complaint.get('District', ''))
        state = normalize_string(complaint.get('State', ''))
        amount_lost = normalize_amount(complaint.get('Amount_Lost', 0))
        status = normalize_string(complaint.get('Status', 'Pending'))
        transaction_count = int(complaint.get('Transaction_Count', 0))
        
        record = {
            'Complaint_ID': complaint_id,  # FORCE STRING
            'Complaint_Date': complaint_date,
            'Category': category,
            'Sub_Category': sub_category,
            'District': district,
            'State': state,
            'Amount_Lost': amount_lost,  # FLOAT ONLY
            'Status': status,
            'Transaction_Count': transaction_count,
            'Data_Quality_Score': int(complaint.get('Data_Quality_Score', 0)),
            'Investigation_Ready': str(complaint.get('Investigation_Ready', 'NO')),
            'Reporting_Delay_Days': int(complaint.get('Reporting_Delay_Days', 0)),
            'Reporting_Delay_Status': str(complaint.get('Reporting_Delay_Status', 'UNKNOWN')),
            'Transaction_Pattern': str(complaint.get('Transaction_Pattern', 'MIXED')),
            'Source_File_Type': str(source_file_type).upper()
        }
        records.append(record)
    
    # Load existing master file if it exists
    existing_df = pd.DataFrame()
    if os.path.exists(master_file):
        try:
            # Read with dtype specification to prevent type inference
            existing_df = pd.read_excel(
                master_file,
                dtype={'Complaint_ID': str}  # Force Complaint_ID as string
            )
        except:
            existing_df = pd.DataFrame()
    
    # Get existing Complaint_IDs for deduplication
    existing_ids = set()
    if not existing_df.empty and 'Complaint_ID' in existing_df.columns:
        # Convert all to string for comparison
        existing_ids = set(existing_df['Complaint_ID'].astype(str).str.strip())
    
    # Filter out duplicates
    new_records = []
    for record in records:
        complaint_id = str(record['Complaint_ID']).strip()
        # Skip if empty or "Not Available" (but allow if it's a generated ID)
        if complaint_id and complaint_id not in existing_ids:
            new_records.append(record)
            existing_ids.add(complaint_id)
    
    new_count = len(new_records)
    
    # Combine existing and new records
    if new_records:
        new_df = pd.DataFrame(new_records)
        
        # Ensure Complaint_ID is string type in new DataFrame
        new_df['Complaint_ID'] = new_df['Complaint_ID'].astype(str)
        
        if existing_df.empty:
            combined_df = new_df
        else:
            # Ensure Complaint_ID is string type in existing DataFrame
            if 'Complaint_ID' in existing_df.columns:
                existing_df['Complaint_ID'] = existing_df['Complaint_ID'].astype(str)
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
    else:
        combined_df = existing_df
    
    # Ensure all required columns exist
    required_columns = [
        'Complaint_ID', 'Complaint_Date', 'Category', 'Sub_Category',
        'District', 'State', 'Amount_Lost', 'Status', 'Transaction_Count',
        'Data_Quality_Score', 'Investigation_Ready', 'Reporting_Delay_Days',
        'Reporting_Delay_Status', 'Transaction_Pattern', 'Source_File_Type'
    ]
    
    for col in required_columns:
        if col not in combined_df.columns:
            if col == 'Amount_Lost':
                combined_df[col] = 0.0
            elif col in ['Transaction_Count', 'Data_Quality_Score', 'Reporting_Delay_Days']:
                combined_df[col] = 0
            else:
                combined_df[col] = 'Not Available'
    
    # Reorder columns
    combined_df = combined_df[required_columns]
    
    # Ensure Complaint_ID is always string
    combined_df['Complaint_ID'] = combined_df['Complaint_ID'].astype(str)
    
    # Ensure Amount_Lost is always float (not scientific notation)
    combined_df['Amount_Lost'] = pd.to_numeric(combined_df['Amount_Lost'], errors='coerce').fillna(0.0)
    
    # Save to master Excel file using safe writing
    os.makedirs('output', exist_ok=True)
    safe_write_excel(combined_df, master_file)
    
    total_count = len(combined_df)
    
    return new_count, total_count
