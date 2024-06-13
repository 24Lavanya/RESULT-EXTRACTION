import time
import pytesseract
from PIL import Image
from io import BytesIO
import base64
import numpy as np
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import cv2

# Configure Tesseract path
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # Update this path

# Replace with the actual USN value
usn = "4CB21CS050"

# Correct path to the ChromeDriver executable
chrome_driver_path = r"C:\Program Files\chromedriver-win64\chromedriver.exe"

# Initialize the Chrome WebDriver with the correct service
service = Service(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service)

# Open the URL
driver.get("https://results.vtu.ac.in/JJEcbcs24/index.php")

# Wait for the page to load and the USN input field to be visible
usn_input = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, "lns")))

# Enter the USN
usn_input.send_keys(usn)

# Locate the CAPTCHA image element
captcha_img_element = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, "//img[contains(@src, '/captcha/vtu_captcha.php')]")))

# Save the screenshot of the webpage
driver.save_screenshot('full_screenshot.png')

# Load the screenshot and find the location of the CAPTCHA image
location = captcha_img_element.location
size = captcha_img_element.size

# Calculate the coordinates for cropping the image
left = location['x']
top = location['y']
right = location['x'] + size['width']
bottom = location['y'] + size['height']

# Open the screenshot using PIL
full_img = Image.open('a.png')

# Crop the CAPTCHA image from the screenshot
# captcha_img = full_img.crop((left, top, right, bottom))
# captcha_img.save('captcha_image.png')  # Save the cropped CAPTCHA image for debugging

# Use Tesseract to extract text from the CAPTCHA image
captcha_text = pytesseract.image_to_string(full_img).strip()
print(captcha_text)

# Re-open the original page to enter the CAPTCHA
driver.get("https://results.vtu.ac.in/JJEcbcs24/index.php")

# Re-enter the USN
usn_input = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, "lns")))
usn_input.send_keys(usn)

# Enter the CAPTCHA text
captcha_input = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, "captchacode")))
captcha_input.send_keys(captcha_text)
# Find and click the submit button
# submit_button = driver.find_element(By.ID, "submit")
# submit_button.click()

# Wait for the results page to load (you may need to adjust the wait time)
time.sleep(5)

# Parse the results page with BeautifulSoup
soup = BeautifulSoup(driver.page_source, 'html.parser')

# Close the browser
driver.quit()

# Process and print the results
# (You'll need to update this part based on the actual structure of the results page)
# print(soup.prettify())
