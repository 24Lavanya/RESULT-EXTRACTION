import time
import pytesseract
from PIL import Image
import csv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service

# Correct path to the ChromeDriver executable
chrome_driver_path = r"C:\Program Files\chromedriver-win64\chromedriver.exe"

# Initialize the Chrome WebDriver with the correct service
service = Service(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service)

def process_captcha(image_path):
    # Load and process the captcha image
    captcha = Image.open(image_path)
    captcha = captcha.convert("L")  # Convert to grayscale
    threshold = 128
    captcha = captcha.point(lambda p: p > threshold and 255)  # Binarize the image
    captcha.save("processed_captcha.png")
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    config = r'--oem 1 --psm 8 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
    captcha_text = pytesseract.image_to_string(captcha, config=config)
    return captcha_text.replace(" ", "").replace("\n", "")

def web(usn):
    try:
        driver.get("https://results.vtu.ac.in/JJEcbcs22/index.php")
        element = driver.find_element(By.XPATH, """//*[@id="raj"]/div[1]/div/input""")
        element.send_keys(usn)

        # Capture captcha image
        captcha_element = driver.find_element(By.XPATH, """/html/body/div[2]/div[1]/div[2]/div/div[2]/form/div/div[2]/div[2]/div[2]/img""")
        captcha_element.screenshot("captcha.png")

        # Process the captcha image
        captcha_text = process_captcha("captcha.png")
        print(f"Extracted Captcha Text: {captcha_text}")

        if len(captcha_text) != 6:
            print("Captcha text length is not 6, retrying...")
            web(usn)
            return

        # Enter captcha text
        captcha_input = driver.find_element(By.XPATH, """//*[@id="raj"]/div[2]/div[1]/input""")
        captcha_input.send_keys(captcha_text)

        # Submit form
        submit_button = driver.find_element(By.XPATH, """//*[@id="submit"]""")
        submit_button.click()
    
        driver.switch_to.alert.accept()
        web(usn)
        return
    except Exception as e:
        print(f"No alert or other error: {e}")
        subject_code = []

        with open("result.csv", "a", newline='') as file:
            writer = csv.writer(file)

            marks = []
            for i in range(1, 10):
                try:
                    xpath = f"//*[@id='dataPrint']/div[2]/div/div/div[2]/div[1]/div/div/div[2]/div/div/div[2]/div/div[{i}]/div[5]"
                    subject_code = driver.find_element(By.XPATH, xpath).text
                    marks.append(subject_code)
                except:
                    break

            if marks:
                marks.insert(0, usn)
                writer.writerow(marks)
            return

def main():
    start_usn = input("Enter the starting USN (e.g., 4CB21CS001): ")
    end_usn = input("Enter the ending USN (e.g., 4CB21CS126): ")

    # Extract the prefix and numeric parts of the USNs
    prefix = start_usn[:-3]
    start_number = int(start_usn[-3:])
    end_number = int(end_usn[-3:])

    for i in range(start_number, end_number + 1):
        usn = f"{prefix}{str(i).zfill(3)}"
        print(usn)
        web(usn)

    print("done")

    # Close the browser
    driver.quit()

if __name__ == "__main__":
    main()
