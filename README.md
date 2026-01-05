# Transcript Automation Script – README

## Overview

This Python automation script logs into the SRMU PeopleSoft portal, downloads student unofficial transcript PDFs, extracts academic details, and stores everything in a structured Excel file.

The script is designed to be **safe, resumable, and reliable**, ensuring no data loss even if the process is interrupted.

---

## Features

- Automatic login to SRMU PeopleSoft
- Downloads Unofficial Transcript PDF
- Extracts:
  - Student Name
  - Subject-wise Grades
  - SGPA
  - CGPA
- Stores data in a single Excel file
- Tracks successful and failed roll numbers
- Incremental saving (no data loss)
- Resume support for failed roll numbers

---

## Requirements

### Python Version
- Python 3.9 or higher

### Required Packages

Install dependencies once:

```bash
pip install playwright pdfplumber openpyxl requests
python -m playwright install



project_folder/
│
├── result.py
├── README.md
├── Transcript_Data.xlsx
├── success_rolls.txt
├── failed_rolls.txt
└── downloads/
    ├── 202210101200001.pdf
    ├── 202210101200002.pdf
    └── ...
Configuration

Edit the following variables inside result.py:

Roll Numbers
ROLL_NUMBERS = [
    "202210101200008",
    "202210101200006",
    ...
]

Password (same for all students)
PASSWORD = "123"

Term Code
TERM = "2501"

Delay Between Students (recommended 2–5 seconds)
DELAY_BETWEEN_STUDENTS = 3

How to Run

From the project directory:

python result.py


The browser will open automatically and run visibly.

Output Files
1. Excel Output – Transcript_Data.xlsx

Columns:

Name

Course 1

Course 2

Course 3

SGPA

CGPA

Course columns are detected automatically from the first successful transcript.

Excel is saved after every student, so progress is never lost.

2. Success Log – success_rolls.txt

Contains roll numbers processed successfully:

202210101200001
202210101200002
202210101200003

3. Failure Log – failed_rolls.txt

Contains roll numbers that failed along with reason:

202210101200010 | Invalid or empty PDF
202210101200015 | No subjects found in PDF
202210101200020 | TimeoutError

Retry Failed Roll Numbers

To retry only failed students, update ROLL_NUMBERS as follows:

with open("failed_rolls.txt") as f:
    ROLL_NUMBERS = [line.split("|")[0].strip() for line in f]


Then run the script again:

python result.py

Safety & Reliability

Single browser instance reused

New session for each student

Automatic error handling

Incremental Excel saving

No duplicate entries

Works even if interrupted

Important Notes

Do not reduce delay below 1 second

Do not run multiple instances simultaneously

Avoid exceeding 150–200 logins per run

Use visible browser mode (headless=False) for stability

Common Issues
Excel file empty?

Headers are created from the first successful transcript

If first few students fail, headers will be created later automatically

PDF downloaded but no data?

Transcript format may have changed

Regex extraction may need update

Status

This script is:

Stable

Resumable

Scalable

Production-ready

For improvements such as parallel batches, retries, or dashboard integration, extend the script as needed.
