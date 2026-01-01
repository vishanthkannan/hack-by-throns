"""
Normalization Module
Provides functions to normalize data before storage
"""

import re
from typing import Any


def normalize_string(value: Any) -> str:
    """
    Normalize string values
    If value is None or empty → return "Not Available"
    Always return string
    """
    if value is None:
        return "Not Available"
    
    value_str = str(value).strip()
    
    if value_str == "" or value_str.lower() in ['nan', 'none', 'null', 'n/a', 'na', '']:
        return "Not Available"
    
    return value_str


def normalize_amount(value: Any) -> float:
    """
    Normalize amount values
    Remove commas, convert to float
    If missing or invalid → return 0.0
    """
    if value is None:
        return 0.0
    
    # If already a number
    if isinstance(value, (int, float)):
        return float(value)
    
    value_str = str(value).strip()
    
    if value_str == "" or value_str.lower() in ['nan', 'none', 'null', 'n/a', 'na']:
        return 0.0
    
    # Remove currency symbols, commas, and whitespace
    value_str = re.sub(r'[₹,Rs\.\s]', '', value_str, flags=re.IGNORECASE)
    
    try:
        return float(value_str)
    except:
        return 0.0


def normalize_complaint_id(value: Any) -> str:
    """
    Normalize Complaint ID - FORCE STRING to prevent Excel conversion
    """
    if value is None:
        return "Not Available"
    
    # Convert to string and ensure it's not treated as number
    value_str = str(value).strip()
    
    if value_str == "" or value_str.lower() in ['nan', 'none', 'null', 'n/a', 'na']:
        return "Not Available"
    
    # Remove any scientific notation artifacts
    if 'e+' in value_str.lower() or 'e-' in value_str.lower():
        # Try to recover from scientific notation
        try:
            num = float(value_str)
            # If it's a large number that was converted, keep as string
            if num > 1e10:
                return str(int(num))
        except:
            pass
    
    return value_str

