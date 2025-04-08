import pickle
from selenium import webdriver

# Set up Selenium WebDriver
driver = webdriver.Chrome()
driver.get("https://teams.microsoft.com/")

# Wait for manual login
input("Log in to Microsoft Teams manually, then press Enter to save cookies...")

# Save cookies
pickle.dump(driver.get_cookies(), open("teams_cookies.pkl", "wb"))
print("Cookies saved! Next time, the script will auto-login.")
driver.quit()
