from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
import time
from selenium.webdriver.chrome.options import Options

# Attach to an existing Chrome session
options = Options()
options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
driver = webdriver.Chrome(options=options)

# Load data from the spreadsheet
file_path = "/Users/andrewkunik/Downloads/citation data formated for python illiad requestor.xlsx"  # Replace with your actual file path
data = pd.read_excel(file_path)

# Open the ILLiad form page
url = "https://uflib.illiad.oclc.org/illiad/FUG/illiad.dll?Action=10&Form=22"  # Replace with your actual ILLiad form URL
driver.get(url)

# Log in if necessary (uncomment and modify if your ILLiad requires login)
# driver.find_element(By.ID, "username_field_id").send_keys("your_username")
# driver.find_element(By.ID, "password_field_id").send_keys("your_password")
# driver.find_element(By.ID, "login_button_id").click()

# Loop through each row in the spreadsheet
for index, row in data.iterrows():
    # Fill out the form fields using the mapped column data
    driver.find_element(By.NAME, "PhotoJournalTitle").send_keys(row["(Journal, Conference Proceedings, Anthology)"])
    driver.find_element(By.NAME, "PhotoArticleTitle").send_keys(row["Article Title"])
    driver.find_element(By.NAME, "PhotoArticleAuthor").send_keys(row["Author"])
    driver.find_element(By.NAME, "PhotoJournalVolume").send_keys(row["Volume"])
    driver.find_element(By.NAME, "PhotoJournalIssue").send_keys(row["Issue Number or Designation(I just guve them all of the number detailes that are at the end of the citation) "])
    driver.find_element(By.NAME, "PhotoJournalYear").send_keys(row["Year"])
    driver.find_element(By.NAME, "PhotoJournalInclusivePages").send_keys(row["pages"])
    
    # Optional fields
    if not pd.isna(row["ISSN"]):
        driver.find_element(By.NAME, "ISSN").send_keys(row["ISSN"])
    if not pd.isna(row["DOI"]):
        driver.find_element(By.NAME, "DOI").send_keys(row["DOI"])
    
    # Submit the form
    driver.find_element(By.ID, "buttonSubmitRequest").click() 

    # Wait for submission and reload the form for the next entry
    time.sleep(2)
    driver.get(url)

# Close the browser
driver.quit()
