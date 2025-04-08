'''
from selenium import webdriver

driver = webdriver.Chrome()
driver.get("https://www.teams.microsoft.com")
print("ChromeDriver is working!")
driver.quit()
'''
'''
import pandas as pd

# Load Excel filec
file_path = "database.xlsx"  # Ensure correct path
df = pd.read_excel(file_path)  # Now 'df' is defined

# Strip spaces from column names
df.columns = df.columns.str.strip()

# Group by Section
sections = df.groupby("Section")["Student Name"].apply(list).to_dict()

# Print output
print(sections)
'''

from selenium import webdriver
from selenium.webdriver.chrome.service import Service

# Set path to the downloaded ChromeDriver
chrome_driver_path = "C:/Users/Acer/Documents/VS Code/CAI Project/chromedriver.exe"
service = Service(chrome_driver_path)

# Initialize WebDriver
driver = webdriver.Chrome(service=service)
