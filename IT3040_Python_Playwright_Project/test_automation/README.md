# IT3040 Assignment 1 - Option 1 Automation Project

This project follows the lecturer's **Python Playwright** automation guideline for testing the PixelsSuite Chat Sinhala transliteration function.

## Files

```text
test_automation/
├── test_automation.py
├── Assignment 1 - Test cases.xlsx
├── testcases.json
├── requirements.txt
├── README.md
└── .gitignore
```

## Install prerequisites

Install Python 3.11 or 3.12.

Then open Command Prompt in the `test_automation` folder.

## Install dependencies

```bash
pip install -U pip
pip install playwright openpyxl
playwright install
```

Or:

```bash
pip install -r requirements.txt
playwright install
```

## Run the automation

Keep the Excel file closed before running.

```bash
python test_automation.py --excel "Assignment 1 - Test cases.xlsx" --url "https://www.pixelssuite.com/chat-translator" --wait-ms 5000 --type-delay-ms 80 --slow-mo-ms 200 --save-every 1 --keep-open
```

## What the script does

The script reads these columns from Excel:

- TC ID
- Input length type
- Input
- Expected output

Then it writes:

- Actual output
- Status

It does not overwrite:

- Singlish input types covered
- Evidence or rationale for the input type covered

## Important notes

- Do not open `Assignment 1 - Test cases.xlsx` while the script is running.
- If some test cases show `Pass`, replace those test cases with stronger negative cases and run again.
- If the website UI changes and the script cannot locate the input/output boxes, update the selector logic inside `test_automation.py`.

## Submission reminder

Before submitting, rename the main folder with your registration number, zip it, and upload it to CourseWeb.

Example:

```text
IT23709966/
└── test_automation/
    ├── test_automation.py
    ├── Assignment 1 - Test cases.xlsx
    ├── requirements.txt
    ├── README.md
    └── .gitignore
```
