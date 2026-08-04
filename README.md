# VTU Result Extraction and Analysis

This project automates the collection of VTU semester results from the university result portal, extracts subject-wise marks, calculates grades, SGPA, percentage, and class, and saves the output as an Excel report.

It is designed for batch processing of many students by generating USN ranges and processing each result automatically.

## What this project does

The workflow is as follows:

1. Read a list of USNs from a starting and ending USN.
2. Open the VTU result page using Selenium.
3. Enter the USN in the form.
4. Capture the CAPTCHA image from the page.
5. Use OCR (via Tesseract) to read the CAPTCHA.
6. Submit the form and wait for the result page.
7. Extract the student name, USN, subject codes, internal marks, external marks, and total marks.
8. Use credit-point data from JSON files to calculate SGPA and percentage.
9. Save the final result into an Excel workbook.

## Project purpose

This tool is useful when you need to:

- process results for a large group of students,
- generate semester-wise analysis files,
- export data into Excel for reporting or review,
- reduce the manual effort of visiting the VTU result portal repeatedly.

## Main files in this project

- [22scheme.py](22scheme.py) - result scraper used for one set of semester/branch data.
- [2ndsem.py](2ndsem.py) - variant for a second-semester result workflow.
- [5thsem.py](5thsem.py) - variant for a fifth-semester result workflow.
- [new22.py](new22.py) - another implementation with JSON export and Excel processing.
- [abhinew.py](abhinew.py) - another variant of the scraper pipeline.
- [credits.json](credits.json), [22credits.json](22credits.json), [5thcredits.json](5thcredits.json) - credit-point mapping files used for SGPA calculation.
- [requirements.txt](requirements.txt) - Python dependencies.
- [readme.txt](readme.txt) - short setup note.

## How the script works

### 1. USN range generation

The scripts ask for a starting USN and an ending USN. They then generate a range of USNs by changing the last three digits.

Example:

- start: 4CB21CS001
- end: 4CB21CS010

This creates:

- 4CB21CS001
- 4CB21CS002
- ...
- 4CB21CS010

### 2. Browser automation with Selenium

The scripts use Selenium to open the VTU result site and interact with the form fields.

The automation steps include:

- opening the page,
- filling in the USN,
- capturing the CAPTCHA image,
- solving it with OCR,
- submitting the form,
- scraping the result table.

### 3. CAPTCHA handling

The project uses OCR to read the image-based CAPTCHA.

The process is:

- take a screenshot of the CAPTCHA image,
- preprocess it with Pillow/OpenCV,
- send it to Tesseract OCR,
- clean the returned text and use it as the typed CAPTCHA.

### 4. Data processing and grade calculation

Once the subject marks are extracted, the scripts:

- assign grades based on marks,
- calculate grade points,
- use credit points from the JSON files,
- compute SGPA, percentage, and class.

### 5. Excel export

The final output is stored in an Excel file such as [results.xlsx](results.xlsx) or [vtu_results.xlsx](vtu_results.xlsx).

Each student row contains:

- USN,
- student name,
- subject-wise marks,
- grade and grade-point columns,
- SGPA,
- percentage,
- class.

## Setup instructions

### 1. Install Python dependencies

Run:

```bash
pip install -r requirements.txt
```

### 2. Install Tesseract OCR

This project depends on Tesseract for CAPTCHA recognition.

On Windows, install Tesseract from the official installer and note the installation path. The scripts currently reference a Windows path similar to:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

If your installation path is different, update it in the script.

### 3. Install Chrome and ChromeDriver

You need:

- Google Chrome installed,
- a matching ChromeDriver version.

Download ChromeDriver from the official Chrome for Testing site and make sure it matches your installed Chrome version.

Then update the ChromeDriver path in the scripts, for example:

```python
chrome_driver_path = r"C:\path\to\chromedriver.exe"
```

### 4. Prepare the credit JSON files

Before running, check the JSON files for the correct subject credit values for your branch/semester.

If the credits are wrong, the SGPA values will be wrong.

## How to use the project

### Option A: Run one of the semester scripts

Example:

```bash
python 2ndsem.py
```

Then enter:

- the starting USN,
- the ending USN.

The script will process the result range and create an Excel file.

### Option B: Run the newer variant

If you want to use the more recent variant:

```bash
python new22.py
```

The script will:

- collect results,
- save them in JSON,
- create a spreadsheet report.

## Important notes

- These scripts are tightly tied to the current structure of the VTU result page.
- The HTML/XPath locations may change if the website updates its layout.
- CAPTCHA recognition may fail sometimes and cause retries.
- The scripts use hard-coded local paths, so they may need to be adjusted for your machine.
- The project is not a polished production-grade application; it is a practical automation script for academic result processing.

## Troubleshooting

### ChromeDriver error

If you see a ChromeDriver-related error, check:

- the installed Chrome version,
- the ChromeDriver version,
- the path in the script.

### Tesseract not found

If OCR fails, verify the Tesseract installation path and update the script accordingly.

### CAPTCHA fails repeatedly

If the CAPTCHA is not being read correctly:

- try a different image preprocessing setup,
- improve the OCR settings,
- check whether the website changed its captcha image format.

### Result page not loading

If the page does not load or the required elements are not found:

- the website may have changed its DOM structure,
- the XPath may need to be updated,
- the connection may be slow or blocked.

## Example workflow

1. Install requirements.
2. Install Tesseract and ChromeDriver.
3. Edit the credit JSON file for the correct branch.
4. Choose the matching script.
5. Run the script.
6. Enter the USN range.
7. Wait for the Excel report to be generated.

## Summary

This repository is a result-scraping and analysis toolkit for VTU student results. It combines Selenium automation, OCR-based CAPTCHA solving, result parsing, grade computation, and Excel export into one workflow. It is a practical academic utility meant to save time when processing large batches of results.
