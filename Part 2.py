import sys
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

"""
Usage (from command prompt):

python auditram_email_agent.py your_email your_password "Your email subject" "Your email body text"

Example:
python auditram_email_agent.py test@gmail.com mypass123 "AuditRAM Test" "This is my assignment email"

The program will:
1. Login to your Gmail
2. Compose an email with the provided subject + body
3. Send it to scittest@auditram.com
"""

def main():
    if len(sys.argv) < 5:
        print("Usage: python auditram_email_agent.py <email> <password> <subject> <body>")
        sys.exit(1)

    email = sys.argv[1]
    password = sys.argv[2]
    subject = sys.argv[3]
    body = sys.argv[4]

    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)

    # Step 1: LOGIN
    driver.get("https://mail.google.com/")

    time.sleep(3)
    driver.find_element(By.ID, "identifierId").send_keys(email)
    driver.find_element(By.ID, "identifierId").send_keys(Keys.ENTER)

    time.sleep(3)
    driver.find_element(By.NAME, "password").send_keys(password)
    driver.find_element(By.NAME, "password").send_keys(Keys.ENTER)

    time.sleep(8)  # wait for inbox to load

    # Step 2: COMPOSE
    compose_btn = driver.find_element(By.XPATH, "//div[contains(text(),'Compose')]")
    compose_btn.click()

    time.sleep(3)

    to_box = driver.find_element(By.NAME, "to")
    to_box.send_keys("scittest@auditram.com")

    subject_box = driver.find_element(By.NAME, "subjectbox")
    subject_box.send_keys(subject)

    body_area = driver.find_element(By.XPATH, "//div[@aria-label='Message Body']")
    body_area.send_keys(body)

    # Step 3: SEND
    send_btn = driver.find_element(By.XPATH, "//div[text()='Send']")
    send_btn.click()

    print("Email sent successfully to scittest@auditram.com")

    time.sleep(5)
    driver.quit()


if __name__ == "__main__":
    main()
