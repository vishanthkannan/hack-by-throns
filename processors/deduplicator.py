"""
Deduplicator Module
Handles intelligence features and master Excel file management
WITH SAFE EXCEL WRITING to prevent type inference and scientific notation
FIXED SCHEMA APPROACH to prevent column mapping errors
"""

import pandas as pd
import os
from datetime import datetime
from typing import List, Dict, Tuple
from openpyxl import load_workbook
from openpyxl.styles import NamedStyle
from processors.normalizer import normalize_string, normalize_amount, normalize_complaint_id


# FIXED COLUMN SCHEMA - Order matters!
COLUMNS = [
    "Complaint_ID",
    "Complaint_Date",
    "Incident_Date",
    "Category",
    "Sub_Category",
    "District",
    "State",
    "Amount_Lost",
    "Status",
    "Transaction_Count",
    "Data_Quality_Score",
    "Investigation_Ready",
    "Reporting_Delay_Days",
    "Reporting_Delay_Status",
    "Transaction_Pattern",
    "Source_File_Type"
]


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


def build_row_from_complaint(complaint: Dict, source_file_type: str) -> Dict:
    """
    Build a row dictionary with ONLY the fixed schema columns
    This ensures correct mapping - no extra keys, no missing keys
    """
    # Extract and normalize all fields explicitly
    complaint_id = normalize_complaint_id(complaint.get('Complaint_ID', ''))
    complaint_date = normalize_string(complaint.get('Complaint_Date', ''))
    incident_date = normalize_string(complaint.get('Incident_Date', ''))
    category = normalize_string(complaint.get('Category', ''))
    sub_category = normalize_string(complaint.get('Sub_Category', ''))
    district = normalize_string(complaint.get('District', ''))
    state = normalize_string(complaint.get('State', ''))
    amount_lost = normalize_amount(complaint.get('Amount_Lost', 0))
    status = normalize_string(complaint.get('Status', 'Pending'))
    transaction_count = int(complaint.get('Transaction_Count', 0))
    data_quality_score = int(complaint.get('Data_Quality_Score', 0))
    investigation_ready = str(complaint.get('Investigation_Ready', 'NO'))
    reporting_delay_days = int(complaint.get('Reporting_Delay_Days', 0))
    reporting_delay_status = str(complaint.get('Reporting_Delay_Status', 'UNKNOWN'))
    transaction_pattern = str(complaint.get('Transaction_Pattern', 'MIXED'))
    source_file_type_str = str(source_file_type).upper()
    
    # Build row with EXACT schema - no extra keys
    row = {
        "Complaint_ID": complaint_id,
        "Complaint_Date": complaint_date,
        "Incident_Date": incident_date,
        "Category": category,
        "Sub_Category": sub_category,
        "District": district,
        "State": state,
        "Amount_Lost": amount_lost,
        "Status": status,
        "Transaction_Count": transaction_count,
        "Data_Quality_Score": data_quality_score,
        "Investigation_Ready": investigation_ready,
        "Reporting_Delay_Days": reporting_delay_days,
        "Reporting_Delay_Status": reporting_delay_status,
        "Transaction_Pattern": transaction_pattern,
        "Source_File_Type": source_file_type_str
    }
    
    return row


def safe_write_excel(df: pd.DataFrame, filepath: str):
    """
    Safely write DataFrame to Excel with explicit type formatting
    Prevents scientific notation and type inference issues
    Sets column widths and formats headers
    """
    # Ensure DataFrame has exactly the columns we expect, in the right order
    if list(df.columns) != COLUMNS:
        # Reorder and select only our columns
        df = df[[col for col in COLUMNS if col in df.columns]]
        # Add missing columns
        for col in COLUMNS:
            if col not in df.columns:
                if col == 'Amount_Lost':
                    df[col] = 0.0
                elif col in ['Transaction_Count', 'Data_Quality_Score', 'Reporting_Delay_Days']:
                    df[col] = 0
                else:
                    df[col] = 'Not Available'
        # Reorder to match COLUMNS exactly
        df = df[COLUMNS]
    
    # Force data types BEFORE writing
    df['Complaint_ID'] = df['Complaint_ID'].astype(str)
    df['District'] = df['District'].astype(str)
    df['State'] = df['State'].astype(str)
    df['Sub_Category'] = df['Sub_Category'].astype(str)
    df['Transaction_Pattern'] = df['Transaction_Pattern'].astype(str)
    df['Category'] = df['Category'].astype(str)
    df['Status'] = df['Status'].astype(str)
    df['Investigation_Ready'] = df['Investigation_Ready'].astype(str)
    df['Reporting_Delay_Status'] = df['Reporting_Delay_Status'].astype(str)
    df['Source_File_Type'] = df['Source_File_Type'].astype(str)
    
    df['Amount_Lost'] = pd.to_numeric(df['Amount_Lost'], errors='coerce').fillna(0.0).astype(float)
    df['Transaction_Count'] = pd.to_numeric(df['Transaction_Count'], errors='coerce').fillna(0).astype(int)
    df['Data_Quality_Score'] = pd.to_numeric(df['Data_Quality_Score'], errors='coerce').fillna(0).astype(int)
    df['Reporting_Delay_Days'] = pd.to_numeric(df['Reporting_Delay_Days'], errors='coerce').fillna(0).astype(int)
    
    # First, write using pandas (creates the file structure)
    df.to_excel(filepath, index=False, engine='openpyxl')
    
    # Now open with openpyxl to apply explicit formatting
    wb = load_workbook(filepath)
    ws = wb.active
    
    # Set column widths to prevent ########
    column_widths = {
        'A': 20,  # Complaint_ID
        'B': 15,  # Complaint_Date
        'C': 15,  # Incident_Date
        'D': 20,  # Category
        'E': 25,  # Sub_Category
        'F': 20,  # District
        'G': 20,  # State
        'H': 15,  # Amount_Lost
        'I': 15,  # Status
        'J': 18,  # Transaction_Count
        'K': 20,  # Data_Quality_Score
        'L': 20,  # Investigation_Ready
        'M': 22,  # Reporting_Delay_Days
        'N': 25,  # Reporting_Delay_Status
        'O': 20,  # Transaction_Pattern
        'P': 18   # Source_File_Type
    }
    
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width
    
    # Find column indices and apply formatting
    header_row = 1
    for col_idx, header in enumerate(COLUMNS, start=1):
        col_letter = ws.cell(row=header_row, column=col_idx).column_letter
        
        # Apply text format to Complaint_ID column
        if header == 'Complaint_ID':
            for row_idx in range(2, len(df) + 2):  # Start from row 2 (skip header)
                cell = ws.cell(row=row_idx, column=col_idx)
                # Force as string
                if cell.value is not None:
                    cell.value = str(cell.value)
                cell.number_format = '@'  # Text format
        
        # Format Amount_Lost column to prevent scientific notation
        elif header == 'Amount_Lost':
            for row_idx in range(2, len(df) + 2):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value is not None:
                    try:
                        # Ensure it's a number, format as number with 2 decimals
                        amount = float(cell.value)
                        cell.value = amount
                        cell.number_format = '#,##0.00'  # Number format with commas
                    except:
                        pass
    
    # Make header row bold and wrap text
    from openpyxl.styles import Font, Alignment
    for col_idx in range(1, len(COLUMNS) + 1):
        header_cell = ws.cell(row=header_row, column=col_idx)
        header_cell.font = Font(bold=True)
        header_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    # Save the workbook
    wb.save(filepath)


def append_to_master_excel(complaints: List[Dict], source_file_type: str) -> Tuple[int, int]:
    """
    Append complaints to master Excel file with deduplication
    Uses FIXED SCHEMA approach to prevent column mapping errors
    Returns: (new_count, total_count)
    """
    master_file = 'output/ncrp_master.xlsx'
    
    # Apply intelligence features
    enhanced_complaints = apply_intelligence_features(complaints)
    
    # Build rows using fixed schema
    new_rows = []
    for complaint in enhanced_complaints:
        row = build_row_from_complaint(complaint, source_file_type)
        new_rows.append(row)
    
    # Create DataFrame with EXPLICIT column order using fixed schema
    if new_rows:
        # Build DataFrame row by row using list comprehension with fixed schema
        data = [[row[col] for col in COLUMNS] for row in new_rows]
        new_df = pd.DataFrame(data, columns=COLUMNS)
    else:
        new_df = pd.DataFrame(columns=COLUMNS)
    
    # Load existing master file if it exists
    existing_df = pd.DataFrame(columns=COLUMNS)
    if os.path.exists(master_file):
        try:
            # Read with dtype specification to prevent type inference
            existing_df = pd.read_excel(
                master_file,
                dtype={'Complaint_ID': str}  # Force Complaint_ID as string
            )
            
            # Ensure existing DataFrame has correct columns
            if list(existing_df.columns) != COLUMNS:
                # Reorder and add missing columns
                for col in COLUMNS:
                    if col not in existing_df.columns:
                        if col == 'Amount_Lost':
                            existing_df[col] = 0.0
                        elif col in ['Transaction_Count', 'Data_Quality_Score', 'Reporting_Delay_Days']:
                            existing_df[col] = 0
                        else:
                            existing_df[col] = 'Not Available'
                # Reorder to match COLUMNS
                existing_df = existing_df[COLUMNS]
        except:
            existing_df = pd.DataFrame(columns=COLUMNS)
    
    # Get existing Complaint_IDs for deduplication
    existing_ids = set()
    if not existing_df.empty and 'Complaint_ID' in existing_df.columns:
        # Convert all to string for comparison
        existing_ids = set(existing_df['Complaint_ID'].astype(str).str.strip())
    
    # Filter out duplicates from new rows
    filtered_rows = []
    for row in new_rows:
        complaint_id = str(row['Complaint_ID']).strip()
        # Skip if empty or "Not Available" (but allow if it's a generated ID)
        if complaint_id and complaint_id != "Not Available" and complaint_id not in existing_ids:
            filtered_rows.append(row)
            existing_ids.add(complaint_id)
    
    new_count = len(filtered_rows)
    
    # Combine existing and new records
    if filtered_rows:
        # Build DataFrame with explicit schema
        data = [[row[col] for col in COLUMNS] for row in filtered_rows]
        filtered_df = pd.DataFrame(data, columns=COLUMNS)
        
        if existing_df.empty:
            combined_df = filtered_df
        else:
            combined_df = pd.concat([existing_df, filtered_df], ignore_index=True)
    else:
        combined_df = existing_df
    
    # Ensure we have a DataFrame (even if empty)
    if combined_df.empty:
        combined_df = pd.DataFrame(columns=COLUMNS)
    
    # Force data types one more time before saving
    combined_df['Complaint_ID'] = combined_df['Complaint_ID'].astype(str)
    combined_df['District'] = combined_df['District'].astype(str)
    combined_df['State'] = combined_df['State'].astype(str)
    combined_df['Sub_Category'] = combined_df['Sub_Category'].astype(str)
    combined_df['Transaction_Pattern'] = combined_df['Transaction_Pattern'].astype(str)
    combined_df['Category'] = combined_df['Category'].astype(str)
    combined_df['Status'] = combined_df['Status'].astype(str)
    combined_df['Investigation_Ready'] = combined_df['Investigation_Ready'].astype(str)
    combined_df['Reporting_Delay_Status'] = combined_df['Reporting_Delay_Status'].astype(str)
    combined_df['Source_File_Type'] = combined_df['Source_File_Type'].astype(str)
    
    combined_df['Amount_Lost'] = pd.to_numeric(combined_df['Amount_Lost'], errors='coerce').fillna(0.0).astype(float)
    combined_df['Transaction_Count'] = pd.to_numeric(combined_df['Transaction_Count'], errors='coerce').fillna(0).astype(int)
    combined_df['Data_Quality_Score'] = pd.to_numeric(combined_df['Data_Quality_Score'], errors='coerce').fillna(0).astype(int)
    combined_df['Reporting_Delay_Days'] = pd.to_numeric(combined_df['Reporting_Delay_Days'], errors='coerce').fillna(0).astype(int)
    
    # Save to master Excel file using safe writing
    os.makedirs('output', exist_ok=True)
    safe_write_excel(combined_df, master_file)
    
    total_count = len(combined_df)
    
    return new_count, total_count
