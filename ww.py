import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# === CONFIGURATION ===
YOUR_NUMBER = "7358639799"
EXCEL_FILE = "topcontacts.xlsx"
IMAGE_PATH = os.path.abspath("nrilf.jpeg")
MESSAGE_TEMPLATE = """Dear Valued Customer,

Warm greetings!

We are here to ensure your experience with us is smooth, rewarding, and to address any issues faced by you at the grass root level.

We invite you to refer your friends, relatives, and colleagues to us.

A detailed brochure and formal email will also be sent from:
@ 

Feel free to reach out to us directly for any assistance or queries. We will revert with solutions asap.

Warm regards,
www"""
WAIT_BETWEEN_MESSAGES = 15

# Verify image path
if not os.path.exists(IMAGE_PATH):
    raise FileNotFoundError(f"Image file not found: {IMAGE_PATH}")
print(f"✅ Image path: {IMAGE_PATH}")

# === LOAD CONTACTS ===
df = pd.read_excel(EXCEL_FILE)
df = df[df['phone number'].astype(str) != YOUR_NUMBER]  # Skip your own number
contacts = df[['Name', 'phone number']].values.tolist()

# === START SELENIUM ===
options = Options()
options.add_argument("--disable-notifications")
driver = webdriver.Chrome(options=options)
driver.get("https://web.whatsapp.com")
input("✅ Scan QR code and press Enter after WhatsApp Web is fully loaded...")

# Initialize WebDriverWait
wait = WebDriverWait(driver, 20)

try:
    for name, raw_number in contacts:
        # Check for stop.txt file to stop
        if os.path.exists("stop.txt"):
            print("⚠️ Stop requested by user (stop.txt found). Exiting...")
            break

        number = str(raw_number).strip()
        full_number = number  # Use number as-is from Excel
        print(f"➡️ Processing {full_number} ({name})")

        try:
            # Open chat via URL
            driver.get(f"https://web.whatsapp.com/send?phone={full_number}&text&app_absent=0")
            time.sleep(8)  # Initial page load

            # Wait for message input box
            msg_box = wait.until(EC.element_to_be_clickable((By.XPATH, '//div[@contenteditable="true"][@data-tab="10"]')))
            print("✅ Message box found")

            # Type message as a single block with Shift+Enter for line breaks
            lines = MESSAGE_TEMPLATE.split('\n')
            for i, line in enumerate(lines):
                if line.strip():  # Skip empty lines
                    msg_box.send_keys(line)
                    if i < len(lines) - 1:  # Add Shift+Enter except for the last line
                        msg_box.send_keys(Keys.SHIFT + Keys.ENTER)
                    time.sleep(0.1)  # Small delay for stability
            time.sleep(1)

            # Wait for attachment button
            attach_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'span[data-icon="plus-rounded"]')))
            print("✅ Attachment button found")
            attach_btn.click()
            time.sleep(1)

            # Wait for file input and upload image
            image_input = wait.until(EC.presence_of_element_located((By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]')))
            image_input.send_keys(IMAGE_PATH)
            print("✅ Image uploaded")
            time.sleep(2)  # Wait for image preview

            # Wait for send button
            try:
                send_btn = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, 'span[data-icon="wds-ic-send-filled"]')))
                print("✅ Send button found")
                send_btn.click()
            except TimeoutException:
                print("⚠️ Send button not found, trying Enter key")
                msg_box.send_keys(Keys.ENTER)
                time.sleep(1)

            print(f"✅ Message sent to {full_number}")

        except (TimeoutException, NoSuchElementException) as e:
            print(f"⚠️ Failed for {full_number}: {str(e)}")
            driver.save_screenshot(f"error_{full_number}.png")
            with open(f"error_{full_number}_source.html", "w", encoding="utf-8") as f:
                f.write(driver.page_source)
            continue
        except Exception as e:
            print(f"⚠️ Unexpected error for {full_number}: {str(e)}")
            driver.save_screenshot(f"error_{full_number}.png")
            continue

        time.sleep(WAIT_BETWEEN_MESSAGES)

finally:
    print("\n🎉 All messages processed or stopped.")
    driver.quit()