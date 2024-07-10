import time
import pytesseract
from PIL import Image
import cv2
import numpy as np
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoAlertPresentException, TimeoutException
import json

# Correct path to the ChromeDriver executable
chrome_driver_path = r"C:\Program Files\chromedriver-win64\chromedriver.exe"

# Initialize the Chrome WebDriver with the correct service
service = Service(executable_path=chrome_driver_path)

def assign_grade(internal_marks, external_marks, total_marks):
    total_marks = int(total_marks)
    external_marks = int(external_marks)
    internal_marks = int(internal_marks)
    if external_marks >= 18 and internal_marks >= 18:
        if 90 <= total_marks <= 100:
            return 'O'
        elif 80 <= total_marks <= 89:
            return 'A+'
        elif 70 <= total_marks <= 79:
            return 'A'
        elif 60 <= total_marks <= 69:
            return 'B+'
        elif 55 <= total_marks <= 59:
            return 'B'
        elif 50 <= total_marks <= 54:
            return 'C'
        elif 40 <= total_marks <= 49:
            return 'P'
    else:
        return 'F'

def grade_point(internal_marks, external_marks, total_marks):
    total_marks = int(total_marks)
    external_marks = int(external_marks)
    internal_marks = int(internal_marks)
    if external_marks >= 18 and internal_marks >= 18:
        if 90 <= total_marks <= 100:
            return 10
        elif 80 <= total_marks <= 89:
            return 9
        elif 70 <= total_marks <= 79:
            return 8
        elif 60 <= total_marks <= 69:
            return 7
        elif 55 <= total_marks <= 59:
            return 6
        elif 50 <= total_marks <= 54:
            return 5
        elif 40 <= total_marks <= 49:
            return 4
    else:
        return 0

def classify_sgpa(internal_marks, external_marks, percentage):
    if (any(mark < 18 for mark in external_marks)) or (any(mark < 18 for mark in internal_marks)):
        return "Fail"
    if percentage <= 60:
        return "Second Class"
    elif percentage <= 69:
        return "First Class"
    elif percentage >= 70:
        return "Distinction"

def process_and_save_data(data, filename, credit_points):
    formatted_data = []

    for (student_name, usn), group in pd.DataFrame(data).groupby(['USN', 'Student Name']):
        formatted_row = {'USN': usn.strip(': '), 'Student Name': student_name.strip(': ')}
        total_credits = 0
        weighted_sum = 0
        total_internal_marks = 0
        total_external_marks = 0
        total_max_marks = 0
        external_marks_list = []
        internal_marks_list = []

        for subject_code in sorted(group['Subject Code'].unique()):
            subject_group = group[group['Subject Code'] == subject_code]
            internal_marks = int(subject_group['Internal Marks'].iloc[0])
            external_marks = int(subject_group['External Marks'].iloc[0])
            total_marks = int(subject_group['Total Marks'].iloc[0])
            grade = assign_grade(internal_marks, external_marks, total_marks)
            gradepoint = grade_point(internal_marks, external_marks, total_marks)

            credit = int(credit_points.get(subject_code, 0))  # Default to 0 if not found
            total_credits += credit
            weighted_sum += gradepoint * credit

            formatted_row[f'{subject_code} CIE'] = internal_marks
            formatted_row[f'{subject_code} SEE'] = external_marks
            formatted_row[f'{subject_code} Total'] = total_marks
            formatted_row[f'{subject_code} Grade'] = grade
            formatted_row[f'{subject_code} Grade Point'] = gradepoint

            total_internal_marks += internal_marks
            total_external_marks += external_marks
            total_max_marks += 100  # Assuming each subject is out of 100 marks
            external_marks_list.append(external_marks)
            internal_marks_list.append(internal_marks)

        sgpa = weighted_sum / total_credits
        percentage = (sgpa - 0.75) * 10
        class_result = classify_sgpa(internal_marks_list, external_marks_list, percentage)

        formatted_row['SGPA'] = round(sgpa, 2)
        formatted_row['Percentage'] = round(percentage, 2)
        formatted_row['Class'] = class_result

        formatted_data.append(formatted_row)

    final_df = pd.DataFrame(formatted_data)

    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False, sheet_name='Sheet1', startrow=1)

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        columns = ['USN', 'Student Name']
        subject_codes = sorted({col.split(' ')[0] for col in final_df.columns if col.endswith('CIE')})
        for subject_code in subject_codes:
            columns.extend([subject_code, '', '', '', ''])
        columns.extend(['SGPA', 'Percentage', 'Class'])

        header1 = columns
        header2 = [''] * 2 + ['CIE', 'SEE', 'Total', 'Grade', 'Grade Point'] * len(subject_codes) + ['', '', '']

        for col_num, value in enumerate(header1):
            worksheet.write(0, col_num, value, workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True}))

        for col_num, value in enumerate(header2):
            worksheet.write(1, col_num, value, workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True}))

        for i in range(2, len(columns) - 3, 5):
            worksheet.merge_range(0, i, 0, i + 4, columns[i], workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True}))

    print(f"Data saved to {filename}")

def process_captcha(image_path):
    # Load the image
    img = cv2.imread(image_path)
        
    # Convert to grayscale
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
    # Define the bounds for the color to be masked
    lower = np.array([80])
    upper = np.array([125])
        
    # Create the mask and apply it
    mask = cv2.inRange(gray, lower, upper)
    img[mask != 0] = [0]
        
    # Save the intermediate image (semisolved)
    cv2.imwrite(r'./Captcha/semisolved.png', img)
        
    # Load the intermediate image
    img = Image.open(r'./Captcha/semisolved.png')
    pixels = img.load()
        
    # Change non-black pixels to white
    for i in range(img.size[0]):
        for j in range(img.size[1]):
            if pixels[i, j] != (0, 0, 0):
                pixels[i, j] = (255, 255, 255)
    
    # Save the final processed image (solved)
    img.save(r'./Captcha/solved.png')
        
    # Read the processed image for OCR using OpenCV
    img = cv2.imread(r'./Captcha/solved.png')
        
    # Perform OCR using pytesseract
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    config = r'--oem 1 --psm 8 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
    captcha_text = pytesseract.image_to_string(img, config=config)
        
    return captcha_text.replace(" ", "").replace("\n", "")

def fetch_and_process_data(usn_list, filename, credit_points):
    all_data = []
    driver = webdriver.Chrome(service=service)

    for usn in usn_list:
        print(f"Currently trying to grab the results of {usn}")
        repeat = True

        while repeat:
            driver.get("https://results.vtu.ac.in/DJcbcs24/index.php")
            element = driver.find_element(By.XPATH, """/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[1]/div/input""")
            element.send_keys(usn)
            
            captcha_element = driver.find_element(By.XPATH, """/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[2]/div[1]/img""")
            captcha_src = captcha_element.get_attribute('src')
            driver.get(captcha_src)

            with open(r'./Captcha/captcha.png', 'wb') as f:
                f.write(driver.find_element(By.XPATH, """/html/body/img""").screenshot_as_png)
                
            driver.back()
            captcha_text = process_captcha(r'./Captcha/captcha.png')
            print(f"The Captcha Text detected was: {captcha_text}")
            
            captcha_input_element = driver.find_element(By.XPATH, """/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[2]/div[2]/input""")
            captcha_input_element.send_keys(captcha_text)
            
            driver.find_element(By.XPATH, """/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[3]/button""").click()
            
            try:
                WebDriverWait(driver, 3).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                print(f"Alert Detected, meaning that Captcha and USN was invalid, retrying")
                alert.accept()
                repeat = True
            except (NoAlertPresentException, TimeoutException):
                repeat = False

        name = driver.find_element(By.XPATH, """/html/body/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]/div[1]/div/span""").text
        usn_number = driver.find_element(By.XPATH, """/html/body/div[2]/div[1]/div[2]/div[1]/div[1]/div[2]/div[2]/div/span""").text

        data = driver.find_elements(By.XPATH, """/html/body/div[2]/div[1]/div[2]/div[1]/div[3]/div/div/div""")
        individual_data = []

        for subject in data:
            rows = subject.find_elements(By.XPATH, "./div")
            if len(rows) >= 8:
                subject_code = rows[1].text
                subject_name = rows[2].text
                external_marks = rows[3].text
                internal_marks = rows[4].text
                total_marks = rows[5].text

                individual_data.append({
                    'USN': usn_number,
                    'Student Name': name,
                    'Subject Code': subject_code,
                    'Subject Name': subject_name,
                    'External Marks': external_marks,
                    'Internal Marks': internal_marks,
                    'Total Marks': total_marks
                })

        all_data.extend(individual_data)
        time.sleep(3)

    process_and_save_data(all_data, filename, credit_points)
    driver.quit()

def main():
    with open('usn_list.json') as f:
        usn_list = json.load(f)
        
    with open('course_credits.json') as f:
        credit_points = json.load(f)

    fetch_and_process_data(usn_list, "Results.xlsx", credit_points)

if __name__ == "__main__":
    main()
