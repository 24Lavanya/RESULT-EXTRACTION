import time
import pytesseract
from PIL import Image
import cv2
import numpy as np
import pandas as pd
from selenium import webdriver
import selenium
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoAlertPresentException, TimeoutException, InvalidArgumentException, NoSuchElementException, UnexpectedAlertPresentException
import json
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC    
chrome_driver_path = r"C:\Users\Lavanya\Downloads\chromedriver-win64 (1)\chromedriver-win64\chromedriver.exe"
service = Service(executable_path=chrome_driver_path)

def assign_grade(subject_code, internal_marks, external_marks, total_marks):
    total_marks = int(total_marks)
    external_marks = int(external_marks)
    internal_marks = int(internal_marks)

    if subject_code != "BSCK307" and subject_code not in ["BNSK359", "BPEK359", "BYOK359"]:
        if external_marks >= 18 and internal_marks >= 18:
            return grade_letter(total_marks)
        else:
            return 'F'
    else:
        if internal_marks >= 18:
            return grade_letter(total_marks)
        else:
            return 'F'

def grade_letter(total_marks):
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

def grade_point(subject_code, internal_marks, external_marks, total_marks):
    if assign_grade(subject_code, internal_marks, external_marks, total_marks) == 'F':
        return 0
    return calculate_grade_point(total_marks)

def calculate_grade_point(total_marks):
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

def classify_sgpa(percentage):
    if percentage >= 70:
        return "Distinction"
    elif 60 <= percentage < 70:
        return "First Class"
    elif 50 <= percentage < 60:
        return "Second Class"
    else:
        return "Fail"

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
            try:
                driver.get("https://results.vtu.ac.in/DJcbcs24/index.php")
                element = driver.find_element(By.XPATH, """/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[1]/div/input""")
                element.send_keys(usn)
                
                captcha_element = driver.find_element(By.XPATH, """/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[2]/div[2]/img""")
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

                  # Adjust timing as needed
                try:
                    alert = driver.switch_to.alert
                    alert.accept()
                except selenium.common.exceptions.NoAlertPresentException:
                    print("No alert present")

                

                try:
                    WebDriverWait(driver, 10).until(EC.alert_is_present())
                    alert = driver.switch_to.alert
                    alert.accept()
                except selenium.common.exceptions.NoAlertPresentException:
                     print("No alert present")


                usn_element = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[1]/td[2]')  
                stud_element = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[1]/div/table/tbody/tr[2]/td[2]')
                table_element = driver.find_element(By.XPATH, '/html/body/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div')

                stud_text = stud_element.text
                usn_text = usn_element.text
                subjects = []
                sub_elements = table_element.find_elements(By.XPATH, 'div')
                num_sub_elements = len(sub_elements)

                for i in range(2, num_sub_elements + 1):
                    subject = {
                        'Student Name': stud_text,
                        'USN': usn_text,
                        'Subject Code': driver.find_element(By.XPATH, f'/html/body/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[{i}]/div[1]').text,
                        'Subject Name': driver.find_element(By.XPATH, f'/html/body/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[{i}]/div[2]').text,
                        'External Marks': driver.find_element(By.XPATH, f'/html/body/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[{i}]/div[3]').text,
                        'Internal Marks': driver.find_element(By.XPATH, f'/html/body/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[{i}]/div[4]').text,
                        'Total Marks': driver.find_element(By.XPATH, f'/html/body/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[{i}]/div[5]').text,
                        'Result': driver.find_element(By.XPATH, f'/html/body/div[2]/div[2]/div[1]/div/div[2]/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[{i}]/div[6]').text
                    }

                    subject['Grade'] = assign_grade(subject['Subject Code'], subject['Internal Marks'], subject['External Marks'], subject['Total Marks'])
                    subject['Grade Point'] = grade_point(subject['Subject Code'], subject['Internal Marks'], subject['External Marks'], subject['Total Marks'])
                    subject['Credit Points'] = credit_points.get(subject['Subject Code'], 0) * subject['Grade Point']

                    subjects.append(subject)

                total_credits = sum(credit_points.get(subject['Subject Code'], 0) for subject in subjects)
                total_grade_points = sum(subject['Credit Points'] for subject in subjects)
                sgpa = round(total_grade_points / total_credits, 2)
                percentage = (sgpa - 0.75) * 10
                classification = classify_sgpa(percentage)

                all_data.append({
                    'Student Name': stud_text,
                    'USN': usn_text,
                    'SGPA': sgpa,
                    'Percentage': percentage,
                    'Classification': classification,
                    'Subjects': subjects
                })

                print(f"Data for USN {usn} processed successfully.")

                repeat = False
                break

            except (NoAlertPresentException, UnexpectedAlertPresentException) as e:
                print(f"Alert Present: {e}")
                alert = driver.switch_to.alert
                if "Invalid captcha code" in alert.text:
                    alert.accept()
                    print("Invalid captcha, retrying...")
                    continue
                else:
                    alert.accept()
                    print("Unexpected error occurred, moving to next USN...")
                    break

            except TimeoutException:
                print("Timeout occurred, retrying...")
                continue

            except InvalidArgumentException as e:
                print(f"Invalid Argument Exception: {e}")
                continue

            except NoSuchElementException as e:
                print(f"Element not found: {e}")
                repeat = False
                break

            except Exception as e:
                print(f"Unexpected Exception: {e}")
                continue

    driver.quit()

    # Saving the results to a JSON file
    with open(filename, 'w') as outfile:
        json.dump(all_data, outfile, indent=4)

    process_and_save_data(filename)

def process_and_save_data(filename):
    with open(filename, 'r') as infile:
        data = json.load(infile)

    all_subjects = set()
    for student in data:
        for subject in student['Subjects']:
            all_subjects.add(subject['Subject Code'])

    all_subjects = sorted(list(all_subjects))

    results_df = pd.DataFrame(columns=['Student Name', 'USN'] + all_subjects + ['SGPA', 'Percentage', 'Classification'])
    for student in data:
        row = {
            'Student Name': student['Student Name'],
            'USN': student['USN'],
            'SGPA': student['SGPA'],
            'Percentage': student['Percentage'],
            'Classification': student['Classification']
        }
        for subject in student['Subjects']:
            row[subject['Subject Code']] = subject['Grade']
        results_df = results_df.append(row, ignore_index=True)

    results_df.to_excel('results.xlsx', index=False)
    print("Results saved to 'vtu_results.xlsx'")




def main():
    start_usn = input("Enter the starting USN (e.g., 4CB22CS001): ")
    end_usn = input("Enter the ending USN (e.g., 4CB22CS126): ")

    prefix = start_usn[:-3]
    start_number = int(start_usn[-3:])
    end_number = int(end_usn[-3:])

    usn_list = [f"{prefix}{str(i).zfill(3)}" for i in range(start_number, end_number + 1)]

    filename = 'results.xlsx'
    
    with open('22credits.json', 'r') as f:
        credit_points = json.load(f)
    print(credit_points)

    fetch_and_process_data(usn_list, filename, credit_points)

if __name__ == "__main__":
    main()