# Product Data Comparison System 3.0 - Project Documentation

## Overview
The Product Data Comparison System is a desktop application for comparing differences between data tables. Main features include:
- Database file management
- Multi-file manual comparison
- Automatic database comparison
- Difference report generation

## System Architecture
```
product-logger-py3.0/
├── main.py                # Program entry
├── requirements.txt       # Python dependencies
├── README.md              # Project introduction
├── PROJECT_DESCRIPTION.md # Project documentation
├── USER_MANUAL.md         # User manual
│
├── data/                  # Data storage directory
│   ├── xlsx/              # Database files (xlsx format)
│   ├── csv/               # Database files (csv format)
│   └── backup/            # Uploaded files backup
│
├── results/               # Comparison results
│   └── compare_reports/   # Manual comparison reports
│
├── ui/                    # User interface
│   └── main_window.py     # Main window implementation
│
├── logic/                 # Core logic
│   └── diff_logic.py      # Data comparison algorithm
│
└── utils/                 # Utility modules
    └── file_utils.py      # File handling utilities
```

## Module Descriptions

### 1. Database Management Module
- Supports uploading xlsx/csv format database files
- Automatically categorizes and stores files by format
- Password-protected upload functionality
- Automatic backup mechanism

### 2. Manual Comparison Module
- Supports comparing 2-3 data files simultaneously
- Automatically checks file format consistency
- Generates detailed difference reports (JSON format)
- Visual difference display

### 3. Database Comparison Module
- Single file comparison with database
- Smart matching based on product ID
- Automatic database updates
- Generates color-coded comparison reports

### 4. Account Management Module
- Password setup and recovery
- Email binding
- Administrator functions

## Technical Implementation
- Development language: Python 3.9+
- GUI framework: PyQt6
- Data processing: pandas/openpyxl
- Report formats: JSON/Excel
- Security mechanism: SHA256 encryption
