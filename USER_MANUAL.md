# Product Data Comparison System 3.0 - User Manual

## 1. Installation Guide

### System Requirements
- Operating System: Windows 10/11
- Python Version: 3.9+
- Memory: 4GB or more
- Disk Space: 500MB available space

### Installation Steps
1. Install Python 3.9+
   - Download installer from: https://www.python.org/downloads/
   - Check "Add Python to PATH" during installation

2. Install dependencies
```bash
pip install -r requirements.txt
```

3. Run the program
```bash
python main.py
```

## 2. Interface Overview

### Main Interface Layout
- Left functional area:
  - Database upload section
  - Manual comparison section (3 upload buttons)
  - Database comparison section
  - Account management section
- Right information display area:
  - Real-time operation logs
  - Comparison results display

## 3. Function Usage Instructions

### 1. Database Management
1. Click "Upload Database File" button
2. Select xlsx/csv format file
3. Enter password for verification (required for first-time setup)
4. System automatically categorizes and backs up files

### 2. Manual Comparison
1. Upload 2-3 comparison files (click upload buttons 1/2/3)
2. Click "Start Comparison" button
3. System automatically checks format consistency
4. View comparison results in the information area
5. Reports are automatically saved to results/compare_reports/

### 3. Database Comparison
1. Upload file to compare
2. Click "Compare with Database" button
3. System automatically matches product IDs
4. Differences are automatically updated to database
5. View comparison summary report

### 4. Account Management
- Change password: Requires old password verification
- Password recovery: Through email verification
- Change email: Requires password verification
- Admin settings: Requires admin password

## 4. Frequently Asked Questions

### Q1: What if file upload fails?
A: Check if file is being used by another program and verify format requirements

### Q2: What if I forgot my password?
A: Click "Forgot Password" button and reset through email verification

### Q3: Comparison results seem inaccurate?
A: Ensure comparison files have identical column structure and headers

### Q4: Program not responding?
A: Large files take time to process, please wait patiently

## 5. Important Notes
1. Backup important files in advance
2. Database files will automatically overwrite old versions
3. After lockout, wait 15 minutes or restart program
4. Regularly clean results directory is recommended
