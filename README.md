# DataZenith

## Overview
This project implements a GUI-based data entry application using Python's tkinter library. The application allows users to enter and save personal and course information into an Excel workbook. It includes error handling for required fields and terms acceptance.

## Features
- **User Information Entry**: Capture first name, last name, title, age, and nationality.
- **Course Information Entry**: Record number of completed courses, number of semesters, and registration status.
- **Terms Acceptance**: User must accept terms before data submission.
- **Data Validation**: Ensures required fields are filled before saving data.
- **Excel Integration**: Data is saved to an Excel workbook (`data.xlsx`).

## Requirements
- Python 3.x
- `tkinter` (standard Python interface to the Tk GUI toolkit)
- `openpyxl` (library for reading and writing Excel files)

## Usage
1. **Setup**: Ensure Python and necessary libraries (`tkinter`, `openpyxl`) are installed.
2. **Execution**: Run `main.py` to start the GUI.
3. **Data Entry**: Fill in all required fields in the form.
4. **Saving Data**: Click "Enter Data" to save entered information to `data.xlsx`.
5. **Validation**: Error messages will prompt if required fields are missing or terms are not accepted.
6. **Clearing Fields**: Use the "Clear Fields" button to reset the form after successful data entry.

## File Structure
- `main.py`: Main Python script containing the GUI setup and logic.
- `README.md`: This file, providing an overview of the project.
- `data.xlsx`: Excel workbook where data is stored.

## Acknowledgments
- Built with Python `tkinter` and `openpyxl`.
- Inspired by practical GUI applications for data entry.
