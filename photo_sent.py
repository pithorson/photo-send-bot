import os
import pandas as pd
from tkinter import filedialog
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import time
import tkinter as tk

# Set up ChromeOptions
options = Options()
options.add_argument("profile-directory=Profile 1")
options.add_argument("C:/Users/computer/OneDrive/Desktop/bot/college sent bot/name/chrome-data")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('useAutomationExtension', False)
driver = webdriver.Chrome(options=options)
driver.get("http://web.whatsapp.com")
time.sleep(10)  # Wait time to scan the QR code in seconds
input("code red")

def send_photos(phone_number, photo_paths):
    # Open chat window
    driver.get("https://web.whatsapp.com/send?phone={}".format(phone_number))

    try:
        # Wait for attachment button to load
        attachment_button = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//div[@title='Attach']")))
        attachment_button.click()

        # Select photo option
        photo_option = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located(
                (By.XPATH, "//input[@accept='image/*,video/mp4,video/3gpp,video/quicktime']")))

        for photo_path in photo_paths:
            photo_paths_str = "\n".join(photo_paths)

            # Send all photos at once
            photo_option.send_keys(photo_paths_str)

            # Wait for photos to load in chat
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, "//span[@data-icon='send']")))

            # Click send button
            send_button = driver.find_element(By.XPATH, "//span[@data-icon='send']")
            send_button.click()

            # Wait for a while before sending the next photo
            time.sleep(5)

    except Exception as e:
        print(f"Error: {e}")

# Prompt the user to select the Excel file
excel_file = filedialog.askopenfilename(initialdir="/", title="Select an Excel File", filetypes=(("Excel File", "*.xlsx"), ("All Files", "*.*")))

# Read the Excel file with names, numbers, and photo paths
df = pd.read_excel(excel_file)

# Create a tkinter window
root = tk.Tk()

# Hide the main window
root.withdraw()

# Ask the user to select the parent directory containing donor subdirectories
parent_directory = filedialog.askdirectory(title="Select Parent Directory")

# Check if a parent directory was selected
if parent_directory:
    # Iterate over donor subdirectories and send all photos to recipients
    for donor_directory in os.listdir(parent_directory):
        donor_directory_path = os.path.join(parent_directory, donor_directory)

        # Check if it's a subdirectory
        if os.path.isdir(donor_directory_path):
            # Check if the donor name exists in the Excel file
            if donor_directory.lower() in df["NAME"].str.lower().values:
                # Get the donor's phone number from the Excel file
                phone_number = str(df.loc[df["NAME"].str.lower() == donor_directory.lower(), "NUM"].values[0]).strip()

                # Check if phone number length is valid
                if len(phone_number) == 12:
                    # Get the list of all photo files in the donor subdirectory
                    photo_files = [os.path.join(donor_directory_path, file) for file in os.listdir(donor_directory_path)
                                   if file.lower().endswith((".jpg", ".jpeg", ".png", ".gif"))]

                    # Check if there are any photo files in the subfolder
                    if photo_files:
                        # Send all photos in the subfolder to the recipient
                        send_photos(phone_number, photo_files)
                        print(f"All {len(photo_files)} photos for {donor_directory} sent to {phone_number}")
                        time.sleep(6)  # Delay between sending each set of photos (adjust as needed)
                    else:
                        print(f"No photo files found in subfolder {donor_directory}")
