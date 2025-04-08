import pandas as pd
import pickle
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

# Load Excel Data
df = pd.read_excel("database.xlsx")

# Fix column names
df.columns = df.columns.str.strip()

# Debugging: Print column names
print("Columns in DataFrame:", df.columns)

# Ensure correct column names
if "Section" not in df.columns or "Student_Name" not in df.columns:
    print("Error: Column names should be 'Section' and 'Student_Name'.")
    exit(1)

# Group students by Section (Team Name)
sections = df.groupby("Section")["Student_Name"].apply(list).to_dict()

# Set up Selenium WebDriver
driver = webdriver.Chrome()
driver.get("https://teams.microsoft.com/")

# Load cookies for auto-login
try:
    cookies = pickle.load(open("teams_cookies.pkl", "rb"))
    for cookie in cookies:
        driver.add_cookie(cookie)
    driver.refresh()
    print("Logged in using saved cookies!")
except Exception as e:
    print("Failed to load cookies. Please log in manually and re-save them.")
    driver.quit()
    exit(1)

# Navigate to "Create a Team" Page
driver.get("https://teams.microsoft.com/_#/conversations/new")

# Loop through Sections to create Teams and add members
for section, students in sections.items():
    team_name = section  # Section = Team Name
    
    try:
        time.sleep(5)
        create_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Create')]")
        create_button.click()

        time.sleep(3)
        from_scratch_button = driver.find_element(By.XPATH, "//button[contains(text(), 'From scratch')]")
        from_scratch_button.click()

        time.sleep(3)
        private_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Private')]")
        private_button.click()

        time.sleep(3)
        team_name_input = driver.find_element(By.XPATH, "//input[@placeholder='Enter team name']")
        team_name_input.send_keys(team_name)

        create_team_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Create team')]")
        create_team_button.click()

        time.sleep(10)  # Wait for team creation

        # Add Members
        for student in students:
            add_member_input = driver.find_element(By.XPATH, "//input[@placeholder='Add members']")
            add_member_input.send_keys(student)
            time.sleep(2)
            add_member_input.send_keys(Keys.RETURN)
            time.sleep(2)

        # Click "Skip" if no more members to add
        skip_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Skip')]")
        skip_button.click()

        print(f"Created Team: {team_name} and added members: {', '.join(students)}")
        time.sleep(5)

    except Exception as e:
        print(f"Error while creating team {team_name}: {e}")

print("All Teams Created Successfully!")
driver.quit()


