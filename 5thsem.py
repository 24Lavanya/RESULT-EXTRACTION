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

def assign_grade(external_marks,total_marks):
    total_marks = int(total_marks)
    external_marks=int(external_marks)
    if external_marks>=18:
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
    

def grade_point(external_marks,total_marks):
    total_marks = int(total_marks)
    external_marks = int(external_marks)
    if external_marks>=18:
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

def classify_sgpa(external_marks, percentage):
    if any(mark < 18 for mark in external_marks):
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
        for subject_code in sorted(group['Subject Code'].unique()):
            subject_group = group[group['Subject Code'] == subject_code]
            internal_marks = int(subject_group['Internal Marks'].iloc[0])
            external_marks = int(subject_group['External Marks'].iloc[0])
            total_marks = int(subject_group['Total Marks'].iloc[0])
            grade = assign_grade(external_marks,total_marks)
            gradepoint = grade_point(external_marks,total_marks)


            
            credit = int(credit_points[subject_code])
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

        sgpa = weighted_sum / total_credits
        percentage = (sgpa - 0.75) * 10
        class_result = classify_sgpa(external_marks_list, percentage)

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
            worksheet.merge_range(0, i, 0, i+4, columns[i], workbook.add_format({'align': 'center', 'valign': 'vcenter', 'bold': True}))

    print(f"Data saved to {filename}")

def process_captcha(image_path):
    captcha = Image.open(image_path)
    captcha = captcha.convert("L")
    threshold = 128
    captcha = captcha.point(lambda p: p > threshold and 255)
    captcha.save("processed_captcha.png")
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    config = r'--oem 1 --psm 8 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
    captcha_text = pytesseract.image_to_string(captcha, config=config)
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
            
            captcha_element = driver.find_element(By.XPATH, """/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[2]/div[2]/img""") #CAPTCHA PATH IS CORRRECT IMAGE COPY FULL XPATH
            captcha_element.screenshot("captcha.png")

            captcha_text = process_captcha("captcha.png")
            print(f"Extracted Captcha Text: {captcha_text}")

            if len(captcha_text) != 6:
                print("Captcha text length is not 6, retrying...")
                continue

            captcha_input = driver.find_element(By.XPATH, """/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[2]/div[1]/input""")
            
            captcha_input.send_keys(captcha_text)

            submit_button = driver.find_element(By.XPATH, """//*[@id="submit"]""")
            submit_button.click()

            try:
                alert = driver.switch_to.alert
                if alert.text == "University Seat Number is not available or Invalid..!":
                    alert.accept()
                    repeat = False
                    break
                elif alert.text == "Invalid captcha code !!!":
                    alert.accept()
                    continue
            except NoAlertPresentException:
                pass

            try:
                element = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '/html/body/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[1]/td[2]'))  #SAME AS STUDENT USN CELL FULL XPATH
                )                                               
            except TimeoutException:
                continue
            
            usn_element = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[1]/td[2]')  

            stud_element = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[2]/td[2]') # PANEL BODY-->TABLE-->2ND TR MEIN 2ND TD JAHA NAME HAI
                                                        
            table_element = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div')  #DIVTABLEBODY


            # print("Stud",stud_element)
            # print("usn",usn_element)
            # print("table",table_element)                                          
            #                                    
            sub_elements = table_element.find_elements(By.XPATH, 'div')
            num_sub_elements = len(sub_elements)
            print("No of elements:",num_sub_elements) 
            stud_text = stud_element.text
            usn_text = usn_element.text
            subjects = []
            for i in range(2, num_sub_elements + 1):
                subject = {
                    'Student Name': stud_text,
                    'USN': usn_text,
                    'Subject Code': driver.find_element(By.XPATH, f'/html/body/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[{i}]/div[1]').text,
                    
                    'Internal Marks': driver.find_element(By.XPATH, f'/html/body/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[{i}]/div[3]').text,          #IN DIVTABLEROW WHERE THERE IS FIRST INTERNAL MARKS
                                                                                      
                    'External Marks': driver.find_element(By.XPATH, f'//*[@id="dataPrint"]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[{i}]/div[4]').text,              #IN DIVTABLEROW WHERE THERE IS FIRST External MARKS
                                                                  
                    'Total Marks': driver.find_element(By.XPATH, f'//*[@id="dataPrint"]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[{i}]/div[5]').text                  #IN DIVTABLEROW WHERE THERE IS FIRST total MARKS
                }
                subjects.append(subject)

            print(f"Extracted subjects for USN {usn}: {subjects}")
            all_data.extend(subjects)

            repeat = False

    driver.quit()

    if all_data:
        process_and_save_data(all_data, filename, credit_points)
    else:
        print("No data extracted, so no file created.")

def main():
    start_usn = input("Enter the starting USN (e.g., 4CB21CS001): ")
    end_usn = input("Enter the ending USN (e.g., 4CB21CS126): ")

    prefix = start_usn[:-3]
    start_number = int(start_usn[-3:])
    end_number = int(end_usn[-3:])

    usn_list = [f"{prefix}{str(i).zfill(3)}" for i in range(start_number, end_number + 1)]

    filename = 'results.xlsx'
    
    with open('5thcredits.json', 'r') as f:
        credit_points = json.load(f)

    fetch_and_process_data(usn_list, filename, credit_points)

    print("done")

if __name__ == "__main__":
    main()