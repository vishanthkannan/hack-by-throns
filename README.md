# NCRP Complaint Automation & Intelligence System

A professional web application for processing NCRP (National Cybercrime Reporting Portal) complaint files with automated data extraction and intelligence features.

## Features

- **Multi-format Support**: Process PDF, CSV, and Excel files
- **Automated Data Extraction**: Extract complaint and transaction data from various file formats
- **Intelligence Features**:
  1. **Data Quality Score**: Normalization score (0-100) based on data completeness
  2. **Investigation Readiness**: Flag indicating if complaint is ready for investigation
  3. **Reporting Delay Indicator**: Calculate and flag delayed reports (>7 days)
  4. **Transaction Pattern Analysis**: Identify SINGLE_LARGE, MULTIPLE_SMALL, or MIXED patterns
- **Deduplication**: Automatically prevents duplicate entries in master Excel file
- **Clean Professional UI**: Government-style minimal interface

## Installation

1. Install Python 3.8 or higher

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

1. Start the Flask application:
```bash
python app.py
```

2. Open your browser and navigate to:
```
http://localhost:5000
```

3. Upload a NCRP complaint file (PDF, CSV, or Excel)

4. The system will:
   - Extract complaint data
   - Apply intelligence features
   - Append to master Excel file (`output/ncrp_master.xlsx`)
   - Skip duplicates automatically

## Project Structure

```
ncrp-auto-uploader/
│
├── app.py                 # Main Flask application
├── requirements.txt       # Python dependencies
│
├── uploads/              # Temporary file storage
├── output/               # Master Excel output
│   └── ncrp_master.xlsx
│
├── processors/           # Data processing modules
│   ├── pdf_processor.py
│   ├── csv_processor.py
│   ├── excel_processor.py
│   └── deduplicator.py
│
├── templates/            # Frontend templates
│   └── index.html
│
└── README.md
```

## Master Excel Output

The master Excel file (`output/ncrp_master.xlsx`) contains the following columns:

- Complaint_ID (Unique Key)
- Complaint_Date
- Category
- Sub_Category
- District
- State
- Amount_Lost
- Status
- Transaction_Count
- Data_Quality_Score
- Investigation_Ready
- Reporting_Delay_Days
- Reporting_Delay_Status
- Transaction_Pattern
- Source_File_Type

## Intelligence Features Details

### 1. Data Quality Score
Calculates a score (0-100) based on:
- Complaint ID present (20 points)
- Complaint Date present (20 points)
- Amount Lost present (20 points)
- District + State present (20 points)
- At least one transaction present (20 points)

### 2. Investigation Readiness
Flags as "YES" if:
- Amount Lost > 0
- At least one transaction ID exists
- Bank/Platform info exists

### 3. Reporting Delay Indicator
Calculates days between Incident Date and Complaint Date:
- > 7 days: DELAYED
- ≤ 7 days: ON_TIME

### 4. Transaction Pattern
- **SINGLE_LARGE**: One transaction AND amount > ₹50,000
- **MULTIPLE_SMALL**: Multiple transactions AND each < ₹10,000
- **MIXED**: All other cases

## Technical Stack

- **Backend**: Python + Flask
- **Frontend**: HTML + Bootstrap 5
- **Data Processing**: pandas, pdfplumber
- **Excel Handling**: openpyxl

## Notes

- Maximum file size: 16MB
- Files are automatically deleted after processing
- Duplicate complaints (based on Complaint_ID) are automatically skipped
- The system handles various date formats and data structures

## License

Prototype for Hackathon Demonstration

