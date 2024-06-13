import time
import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
from io import BytesIO
import base64
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import cv2
import numpy as np

# Configure Tesseract path
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"  # Update this path

usn = "your_actual_usn_value"  # Replace with the actual USN value
url = "https://results.vtu.ac.in/JJEcbcs24/index.php"

# Correct path to the ChromeDriver executable
chrome_driver_path = r"C:\Program Files\chromedriver-win64\chromedriver.exe"

# Initialize the Chrome WebDriver with the correct service
service = Service(executable_path=chrome_driver_path)
driver = webdriver.Chrome(service=service)

# Open the URL
driver.get(url)

# Wait for the CAPTCHA image to load
captcha_element = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.CLASS_NAME, "col-md-4 img"))  # Update this selector based on the actual class of the captcha image
)

# Capture the CAPTCHA image as base64
captcha_base64 = captcha_element.screenshot_as_base64

# Convert base64 string to image
captcha_image = Image.open(BytesIO(base64.b64decode(captcha_base64)))

# Convert the image to grayscale
gray_image = captcha_image.convert('L')

# Increase the contrast of the image
enhancer = ImageEnhance.Contrast(gray_image)
enhanced_image = enhancer.enhance(2)

# Apply a filter to sharpen the image
sharpened_image = enhanced_image.filter(ImageFilter.SHARPEN)

# Convert the sharpened image to a numpy array
open_cv_image = np.array(sharpened_image)

# Apply binary threshold
threshold_value = 150  # Adjusted threshold value
_, binary_image = cv2.threshold(open_cv_image, threshold_value, 255, cv2.THRESH_BINARY_INV)

# Apply dilation to enhance the bold text
kernel = np.ones((2, 2), np.uint8)
dilated_image = cv2.dilate(binary_image, kernel, iterations=1)

# Apply erosion to remove noise
eroded_image = cv2.erode(dilated_image, kernel, iterations=1)

# Convert back to PIL image
processed_image_pil = Image.fromarray(eroded_image)

# Crop the image to focus more precisely on the CAPTCHA area
width, height = processed_image_pil.size
left = width * 0.35
top = height * 0.35
right = width * 0.65
bottom = height * 0.65
cropped_processed_image = processed_image_pil.crop((left, top, right, bottom))

# Use OCR to read the CAPTCHA text
captcha_text = pytesseract.image_to_string(cropped_processed_image, config='--psm 7')
captcha_text = captcha_text.strip()  # Clean up the OCR result

# Print the OCR result
print("CAPTCHA text:", captcha_text)

# Locate the input elements and send the USN and CAPTCHA
driver.find_element(By.CLASS_NAME, "form-control").send_keys(usn)
driver.find_element(By.NAME, "captchacode").send_keys(captcha_text)  # Update this selector based on the actual name of the captcha input field

# Optionally, submit the form
# driver.find_element(By.ID, "submit_button_id").click()  # Update with the actual ID of the submit button

time.sleep(10)
# Close the browser
driver.quit()
