time for time-related functions.
pytesseract for Optical Character Recognition (OCR) to read the captcha.
PIL (Pillow) to manipulate images.
csv to write data to a CSV file.
selenium to automate web browser interaction.
By: Selenium class for locating elements.
Service: Selenium class for managing the ChromeDriver service


CHROME DRIVER:
Specifies the path to the ChromeDriver executable and initializes the Chrome WebDriver with this path. The Service class is used to manage the ChromeDriver service, and webdriver.Chrome creates a new instance of the Chrome browser.

PROCESS_CAPTCHA:
captcha = Image.open(image_path): Opens the image file.
captcha.convert("L"): Converts the image to grayscale.
captcha.point(lambda p: p > threshold and 255): Binarizes the image (converts it to black and white) using a threshold.
captcha.save("processed_captcha.png"): Saves the processed image.
pytesseract.pytesseract.tesseract_cmd: Specifies the path to the Tesseract OCR executable.
config: Configuration for Tesseract OCR, specifying the character whitelist and the page segmentation mode.
pytesseract.image_to_string: Extracts text from the processed image using Tesseract OCR.
return captcha_text.replace(" ", "").replace("\n", ""): Cleans up the extracted text and returns it.


USN:Main Function: Handles user input and manages the range of USNs to process.
Takes the starting and ending USNs as input from the user.
Extracts the prefix and numeric parts of the USNs for iteration.