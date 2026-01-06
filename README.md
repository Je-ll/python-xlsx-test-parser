# Python XLSX Test Parser

A Python utility to parse test cases from Excel (XLSX) files and convert them into formatted text output.

## Description

This script reads test case tables from Excel files and formats them into readable text blocks. It automatically detects test case tables with expected headers and extracts each test case into a standardized format.

## Requirements

- Python 3.7+
- openpyxl

## Installation

1. Clone this repository:
```bash
git clone https://github.com/Je-ll/python-xlsx-test-parser.git
cd python-xlsx-test-parser
```

2. Install the required dependency:
```bash
pip install openpyxl
```

## Usage

### Basic Usage

Parse all sheets in an XLSX file:
```bash
python main.py path/to/your/test-cases.xlsx
```

### Parse a Specific Sheet

To parse only a specific sheet:
```bash
python main.py path/to/your/test-cases.xlsx --sheet "Sheet1"
```

## Expected XLSX Format

The script looks for test case tables with the following column headers (case-insensitive):

- test case id
- requirements
- level
- description
- pre-conditions
- test steps
- test data
- expected result
- actual result
- status

**Note:** The script also tolerates the common typo "desciption" instead of "description".

### XLSX File Structure

- Multiple test case tables can exist in a single sheet
- Tables are separated by empty rows
- Each table must have the expected header row
- Only rows with a "test case id" value will be parsed

## Output Format

Each test case is formatted as:

```
Test Case ID: <value>
Requirements: <value>
Level: <value>
Description: <value>
Pre-Conditions: <value>
Test Steps: <value>
Test Data: <value>
Expected Result: <value>
```

Multiple test cases are separated by blank lines in the output.

## Example

Given an XLSX file with test cases, running:
```bash
python main.py my-tests.xlsx
```

Will output:
```
Test Case ID: <TC001>
Requirements: <REQ-123>
Level: <High>
Description: <Verify user login functionality>
Pre-Conditions: <User account exists>
Test Steps: <1. Open login page 2. Enter credentials 3. Click login>
Test Data: <username: test@example.com, password: test123>
Expected Result: <User successfully logged in and redirected to dashboard>


Test Case ID: <TC002>
Requirements: <REQ-124>
Level: <Medium>
Description: <Verify password reset>
Pre-Conditions: <N/A>
Test Steps: <1. Click forgot password>
Test Data: <N/A>
Expected Result: <Password reset email sent>
```

Note: Fields with empty or missing values will display as "N/A".

## Features

- Parses all worksheets in a workbook or a specific sheet
- Handles multiple test case tables within a single sheet
- Normalizes header names (case-insensitive matching)
- Tolerates common typos in headers
- Skips empty rows and invalid data
- Only extracts rows with valid test case IDs

## License

This project is provided as-is for personal use.
