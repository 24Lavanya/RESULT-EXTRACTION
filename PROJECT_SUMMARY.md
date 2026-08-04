# Project Summary

## Overview

This project is a Python-based automation tool for extracting VTU examination results, processing the data, and generating structured academic reports. It is built around Selenium for browser automation and Tesseract OCR for CAPTCHA recognition.

## What the project is trying to solve

Manually checking VTU results for many students is repetitive and time-consuming. This project removes that effort by automating the process:

- it opens the VTU result portal,
- enters a student USN,
- solves the displayed CAPTCHA,
- fetches the marks page,
- extracts subject-wise results,
- calculates academic metrics,
- and saves everything into an Excel workbook.

## Main technical approach

### Browser automation

The scripts use Selenium WebDriver to interact with the VTU result page. This allows the program to:

- open the website,
- type into form fields,
- click buttons,
- inspect the response page,
- and extract the displayed result content.

### OCR-based CAPTCHA solving

Because the result system includes a CAPTCHA, the project uses Python imaging libraries and Tesseract OCR:

- the CAPTCHA image is captured from the page,
- it is preprocessed to improve readability,
- OCR is applied to convert the image to text,
- and the text is inserted into the captcha input field.

### Data processing

After the marks are extracted, the project processes them by:

- grouping data by student,
- organizing subject-level marks,
- assigning grades,
- calculating grade points,
- applying credit points,
- computing SGPA and percentage,
- and classifying the result into categories such as First Class or Distinction.

## Why there are several Python scripts

The repository contains multiple script variants rather than a single unified script. Each file appears to be a different version or adaptation for a particular use case:

- some scripts target different semester workflows,
- some have slightly different field extraction logic,
- some export to Excel in different layouts,
- and some are more experimental than others.

This means the repository is better understood as a collection of related automation experiments rather than a single polished app.

## Expected output

The project produces Excel files containing result summaries such as:

- USN,
- student name,
- subject codes,
- internal marks,
- external marks,
- total marks,
- grades,
- grade points,
- SGPA,
- percentage,
- and class.

## Dependencies used

The scripts rely mainly on:

- Selenium for browser automation,
- Pillow and OpenCV for image handling,
- Tesseract OCR for CAPTCHA reading,
- pandas for table handling and Excel generation,
- and open-source Python libraries required by the above.

## Strengths of the project

- Automates a repetitive academic task.
- Processes many students in a batch.
- Converts raw portal data into structured Excel reports.
- Provides a useful base for future improvements.

## Limitations

The project also has several practical limitations:

- it depends on the current HTML layout of the VTU result page,
- XPath selectors may break if the site changes,
- CAPTCHA recognition can be inconsistent,
- the scripts use hard-coded paths for ChromeDriver and Tesseract,
- and some files appear to be older or duplicate implementations.

## Recommended next improvements

To make the project more reliable and maintainable, the next steps could be:

1. create a single clean main script,
2. replace hard-coded paths with configurable settings,
3. add a proper configuration file,
4. improve CAPTCHA handling,
5. add logging and error handling,
6. and support more semesters and branches through a generic architecture.

## Final summary

This repository is a practical academic automation project that turns VTU result data into structured reports. It combines web scraping, OCR, grading logic, and spreadsheet generation in one workflow. The core idea is simple and useful: reduce manual effort by automatically collecting and analyzing student results.
