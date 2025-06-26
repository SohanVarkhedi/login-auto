from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import time
import subprocess as sub

# Load Excel file
df = pd.read_excel("Santosh_Data_With_Status.xlsx")

# Fix dtype warning by casting status columns to string
df["Login Status"] = df.get("Login Status", "").astype(str)
df["Logout Status"] = df.get("Logout Status", "").astype(str)

# ✅ Use Chrome profile 'Profile 4' with visible window
options = Options()
#options.add_argument("--new-window")  # Opens tab properly
chrome_path="C:\Program Files\Google\Chrome\Application\chrome.exe"
process=sub.Popen([chrome_path])
chrome_pid=process.pid
print(f"launched w pid :{chrome_pid}")

#options.add_argument(r"user-data-dir=C:\Users\sohan\AppData\Local\Google\Chrome\User Data")
#options.add_argument("profile-directory=Profile 4")  # Your actual profile
#options.add_experimental_option("detach", True)  # Chrome stays open even after script ends

# Start browser
driver = webdriver.Chrome(options=options)

try:
    for index, row in df.iterrows():
        login_id = row['Code']
        password = "112233"

        # Step 1: Open login page
        driver.get("https://www.bluebirdservices.co.in/login.php")
        time.sleep(10)

        try:
            # Step 2: Fill login form
            driver.find_element(By.ID, "uid").clear()
            driver.find_element(By.ID, "uid").send_keys(login_id)
            driver.find_element(By.NAME, "pass").clear()
            driver.find_element(By.NAME, "pass").send_keys(password)
            driver.find_element(By.XPATH, "//input[@value='Sign In']").click()
            time.sleep(5)

            # Step 3: Check if login successful
            try:
                driver.find_element(By.XPATH, "//a[@href='logout.php']")
                df.at[index, "Login Status"] = "Success"
            except NoSuchElementException:
                df.at[index, "Login Status"] = "Failed"
                df.at[index, "Logout Status"] = "-"
                print(f"❌ Login failed for ID: {login_id}")
                break  # Stop if login fails

            # Step 4: Logout
            driver.find_element(By.XPATH, "//a[@href='logout.php']").click()
            df.at[index, "Logout Status"] = "Success"
            print(f"✅ Logged out for ID: {login_id}")
            time.sleep(2)

            # Step 5: Click Login again
            driver.find_element(By.XPATH, "//a[@href='login.php']").click()
            time.sleep(2)

        except Exception as e:
            print(f"⚠️ Error during login/logout for ID {login_id}: {e}")
            df.at[index, "Login Status"] = "Error"
            df.at[index, "Logout Status"] = "-"
            break

finally:
    # Save updated Excel file
    df.to_excel("Santosh_Data_With_Status.xlsx", index=False)
    print("✅ Script finished. Excel updated.")
