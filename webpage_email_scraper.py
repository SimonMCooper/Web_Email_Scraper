import pandas as pd
import re
import time
from getpass import getpass
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import requests

# Prompt for credentials
username = input("Enter your username: ")
password = getpass("Enter your password: ")

# Load Excel file
df = pd.read_excel("file to be read", sheet_name=1, engine="openpyxl") #change file to be read to the name of the input file stored in th e same folder as this script.
entity_ids = df["PageId"].dropna().astype(int).unique()

# Set up Edge WebDriver
service = EdgeService("msedgedriver.exe")  # Ensure this is in the same folder or provide full path
options = EdgeOptions()
driver = webdriver.Edge(service=service, options=options)
wait = WebDriverWait(driver, 20)

# Step 1: Go to login page
driver.get("url") ##change to login page url

# Step 2: Enter username and password
wait.until(EC.presence_of_element_located((By.ID, "UserName"))).send_keys(username) #change username if named differently in html code
driver.find_element(By.ID, "Password").send_keys(password + Keys.RETURN) #change password if named differently in html code

# Step 3: Wait for OTP page and prompt for code - this can be removed if there is no multifactor authentication on the account
wait.until(EC.presence_of_element_located((By.ID, "Code"))) # change code if the OTP is different in the html code
otp = input("Enter the OTP sent to your device: ")
driver.find_element(By.ID, "Code").send_keys(otp + Keys.RETURN) # change code if the OTP is different in the html code

# Step 4: Wait for login to complete (adjust ID if needed)
wait.until(EC.presence_of_element_located((By.ID, "split-layout")))

# Step 5: Scrape emails from each page
results = []
for entity_id in entity_ids:
    url = f"url{entity_id}" #change url to the url before the page id
    driver.get(url)
    time.sleep(3)  # Wait for page to load

    # Get the fully rendered page source from Selenium
    page_source = driver.page_source
    soup = BeautifulSoup(page_source, "html.parser")

    emails = []
    for a_tag in soup.find_all("a", href=True):
        href = a_tag["href"]
        if href.startswith("mailto:"):
            email = href[7:].split("?")[0]
            emails.append(email)

    emails = list(set(emails))  # Remove duplicates
    results.append({"PageId": entity_id, "Emails": ", ".join(emails)})


# Step 6: Save results to CSV
output_df = pd.DataFrame(results)
output_df.to_csv("filename.csv", index=False) #change filename to required filename

print("✅ Scraping complete. Results saved to 'csv file'.")
driver.quit()
