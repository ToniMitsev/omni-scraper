import time
import pandas as pd
from selenium import webdriver
from selenium.common import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import tkinter as tk
from tkinter import messagebox


def run_script():
    # Retrieving user inputs from the GUI fields
    omnilinx_username = username_entry.get()
    omnilinx_password = password_entry.get()

    # Check if the pages entry is valid
    try:
        total_pages = int(pages_entry.get())
    except ValueError:
        messagebox.showerror("Invalid input", "Please enter a valid number for pages.")
        return

    # Show message in GUI that script is running
    status_label.config(text="Starting extraction... Please wait.")
    root.update_idletasks()

    # Selenium WebDriver setup
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--start-maximized")  # This starts the browser maximized
    options.add_experimental_option("prefs", {
        "profile.default_content_setting_values.notifications": 2,  # Block notifications
        "profile.default_content_setting_values.geolocation": 1,  # Automatically allow geolocation
        "profile.default_content_setting_values.media_stream": 1,
        # Automatically allow media stream (camera/microphone)
    })
    options.add_experimental_option("detach", True)  # Keeps browser open after script ends

    driver = webdriver.Chrome(options=options, service=Service(ChromeDriverManager().install()))

    def login(driver, email, password):
        email_input = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located(
                (By.XPATH, '/html/body/div/div/div/div/div/div[1]/div[2]/form/div[1]/div/div[1]/div/input'))
        )
        email_input.send_keys(email)

        password_input = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located(
                (By.XPATH, '/html/body/div/div/div/div/div/div[1]/div[2]/form/div[2]/div/div[1]/div/input'))
        )
        password_input.send_keys(password)

        login_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, '/html/body/div/div/div/div/div/div[1]/div[2]/form/div[4]/button'))
        )
        login_button.click()

        WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH, '/html/body/div/div[1]/header/div/img')))

        driver.find_element(By.XPATH, '/html/body/div/div[1]/div[1]/div[5]/div/a').click()

    driver.get("https://app.omnilinx.com/")

    login(driver, omnilinx_username, omnilinx_password)
    time.sleep(10)

    # Handling potential popups
    try:
        WebDriverWait(driver, 10).until(EC.alert_is_present(),
                                        'Timed out waiting for PA creation confirmation popup to appear.')
        alert = driver.switch_to.alert
        alert.accept()
        print("Alert accepted")
    except TimeoutException:
        print("No alert")

    time.sleep(3)

    try:
        WebDriverWait(driver, 10).until(EC.alert_is_present(),
                                        'Timed out waiting for PA creation confirmation popup to appear.')
        alert = driver.switch_to.alert
        alert.accept()
        print("Alert accepted")
    except TimeoutException:
        print("No alert")

    driver.find_element(By.XPATH,
                        '/html/body/div/div[1]/main/div/div[1]/div/div/div[1]/div[2]/div/div/div[1]/div[3]/div/div[2]/div/div/div/div[1]/div[2]/div').click()
    driver.find_element(By.XPATH, '/html/body/div/div[57]/div/div[3]/div/div').click()

    time.sleep(10)

    all_highlighted_texts = []

    for i in range(total_pages):
        # Wait for the highlighted text elements to load
        highlighted_text_elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "highlighted-text"))
        )

        highlighted_texts = []

        for element in highlighted_text_elements:
            text = element.text.strip()
            if text and '@' in text:
                highlighted_texts.append(text)

        all_highlighted_texts.extend(highlighted_texts)

        # Update the status in the GUI
        status_label.config(text=f"Iteration {i + 1}: Found {len(highlighted_texts)} emails")
        root.update_idletasks()

        print(f"Iteration {i + 1}: {highlighted_texts}")

        try:
            next_button = driver.find_element(By.XPATH,
                                              '/html/body/div/div[1]/main/div/div[1]/div/div/div[1]/div[2]/div/div/div[1]/div[3]/div/div[1]/div/button[2]/span/i')
            next_button.click()
        except Exception as e:
            print(f"Error finding 'next' button: {e}")
            break  # Break the loop if the next button cannot be clicked

        time.sleep(5)

    print(f"All highlighted texts containing '@': {all_highlighted_texts}")

    # Convert the Final list into a pandas DataFrame
    df = pd.DataFrame(all_highlighted_texts, columns=["Data"])

    # Export to an Excel file
    output_file = "Omnilinx Exported Data.xlsx"
    df.to_excel(output_file, index=False, engine='openpyxl')

    # Show completion message in the GUI
    messagebox.showinfo("Success", f"Data has been exported to {output_file}")
    status_label.config(text="Extraction complete. Data saved to Excel.")
    root.update_idletasks()

    driver.close()

    print("***********************************")
    print("***                             ***")
    print("*** Moже да затворите прозореца ***")
    print("***                             ***")
    print("***********************************")


# Create the main window
root = tk.Tk()
root.title("Omnilinx Data Extractor")

# Create labels and input fields
tk.Label(root, text="Username:").grid(row=0, column=0, padx=10, pady=5)
username_entry = tk.Entry(root, width=30)
username_entry.grid(row=0, column=1, padx=10, pady=5)

tk.Label(root, text="Password:").grid(row=1, column=0, padx=10, pady=5)
password_entry = tk.Entry(root, width=30, show="*")
password_entry.grid(row=1, column=1, padx=10, pady=5)

tk.Label(root, text="Total Pages:").grid(row=2, column=0, padx=10, pady=5)
pages_entry = tk.Entry(root, width=30)
pages_entry.grid(row=2, column=1, padx=10, pady=5)

# Create the start button
run_button = tk.Button(root, text="Start Extraction", command=run_script)
run_button.grid(row=3, column=0, columnspan=2, pady=10)

# Status label to display progress
status_label = tk.Label(root, text="Status: Waiting for input...", fg="blue")
status_label.grid(row=4, column=0, columnspan=2, pady=10)

# Start the Tkinter event loop
root.mainloop()
