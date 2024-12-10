# interview_app_tk
Interview Question Application using Python Tkinter

# ITAlchemy Interview Application

Version: v1.0  
Author: Brendan Choi  
Date: December 1, 2024  

## Overview

The ITAlchemy Interview App is designed to assist interviewers by providing a structured way to prepare and display questions and answers during interviewing process. It ensures consistency and fairness in interviews through standardisation.

---

## Features

- Displays questions categorized by difficulty (Easy, Medium, Hard).
- Randomly selects questions for each interview session.
- Allows resetting of used questions.
- Provides a user-friendly interface with zoom functionality.
- Includes a logo for branding.

---

## Requirements

- Python (3.10+ recommended)
- Required Python libraries:
  - `pandas`
  - `openpyxl`
  - `tkinter`
  - `pyinstaller`
---

## Setup Instructions

1. Place the following files in the same directory:
   - `italchemy_interview_app_v1.0.py`
   - `interview_questions_v1.xlsx`
   - `italchemy_logo.png`

2. Compile the script using PyInstaller (optional):
   pyinstaller --onefile --windowed --add-data "interview_questions_v1.xlsx;." --add-data "italchemy_logo.png;." italchemy_interview_app_v1.0.py
