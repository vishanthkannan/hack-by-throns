# NCRP Complaint Automation Tool

A lightweight Flask-based web application designed to process NCRP complaint files and automatically extract structured complaint information for quick review and analysis.

## Key Highlights
- Supports NCRP **PDF and CSV** complaint files  
- Automatically extracts complaint details using **rule-based and regex-driven logic**  
- Handles **multiple NCRP document layouts**  
- Detects and avoids **duplicate complaints** using Complaint ID  
- Displays extracted data in a clean, structured format  
- **Does not rely on Excel** for processing or storage  

## Core Technologies
- Python 3.11  
- Flask  
- pandas  
- pdfplumber  
- Gunicorn  
- Docker  

## System Workflow
1. Upload NCRP complaint files (PDF/CSV) via the web interface  
2. Extract key fields such as Complaint ID, dates, category, district, amount, and status  
3. Validate records and filter duplicates  
4. Present structured output for easy review and reporting  

## Scope & Limitations
- Works with **text-based PDFs and CSV files only**  
- Excel is **not used** for data processing or storage  
- OCR for scanned PDFs is not included in the current version  

## Future Enhancements
- OCR support for scanned NCRP documents  
- Database-backed storage and analytics  
- Advanced cybercrime pattern detection
