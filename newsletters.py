import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager
import time
import logging
import os
import ctypes
import threading
import random

# Disable Quick Edit mode in Windows Command Prompt
def disable_quick_edit():
    kernel32 = ctypes.windll.kernel32
    handle = kernel32.GetStdHandle(-10)
    mode = ctypes.c_uint()
    kernel32.GetConsoleMode(handle, ctypes.byref(mode))
    new_mode = mode.value & ~(0x0040)
    kernel32.SetConsoleMode(handle, new_mode)

disable_quick_edit()

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Global counters and stop flag
Subscribed, good, bad = 0, 0, 0
stop_flag = False  # Flag to stop the subscription process

# Newsletter URLs and selectors
newsletters = {
    "Forbes": {
        "url": "https://account.forbes.com/newsletters",
        "email": (By.XPATH, '/html/body/div[1]/div/div[2]/div[2]/input'),
        "button1": (By.XPATH, '/html/body/div[4]/div/div[2]/div[4]/div/div[1]/div[2]/button'),  # First button
        "button2": (By.XPATH, '/html/body/div[1]/div/div[3]/button'),  # Second button
        "confirmation": (By.XPATH, '//*[contains(text(), "You‚Äôre Subscribed!")]')
    },
    "Fragrance Foundation": {
        "url": "https://www.fragrancefoundation.fr/newsletter-adherents/",
        "email": (By.XPATH, '//*[@id="mce-EMAIL"]'),
        "button": (By.XPATH, '//*[@id="mc-embedded-subscribe"]'),
        "confirmation": (By.XPATH, '//*[contains(text(), "Thank you for subscribing!")]')
    },
    "Dayspring": {
        "url": "https://www.dayspring.com/",
        "email": (By.XPATH, '//*[@id="st-signup-footer-email"]'),
        "button": (By.XPATH, '//*[@id="sailthru-signup-footer"]'),
        "confirmation": (By.XPATH, '//*[contains(text(), "Your first email is on its way.")]')
    },
    "Lifehacker": {
        "url": "https://lifehacker.com/newsletters",
        "email": (By.XPATH, '//*[@id="email"]'),
        "button": (By.XPATH, '//*[@id="newsletter-form"]/button'),
        "confirmation": (By.XPATH, '//*[contains(text(), "Success!")]')
    },
    "Indigo": {
        "url": "https://www.indigo.ca/en-ca/",
        "email": (By.XPATH, '//*[@id="footercontent"]/div/div[2]/div[4]/div[1]/div[2]/form/div/input'),
        "button": (By.XPATH, '//*[@id="footercontent"]/div/div[2]/div[4]/div[1]/div[2]/form/div/span/button'),
        "confirmation": (By.XPATH, '//*[contains(text(), "Thanks for subscribing!")]')
    },
    "Kiplinger": {
        "url": "https://my.kiplinger.com/email/signup.php",
        "email": (By.XPATH, '//*[@id="kipEmail"]'),
        "button": (By.XPATH, '//*[@id="emailsubmit"]'),
        "confirmation": (By.XPATH, '//*[contains(text(), "Form submitted. Thank you for subscribing!")]')
    },
    "Goop": {
        "url": "https://goop.com/subscribe/",
        "email": (By.XPATH, '//*[@id="subscribe-inner"]/div[2]/form/div[1]/label/input'),
        "button": (By.XPATH, '//*[@id="subscribe-inner"]/div[2]/form/button'),
        "confirmation": (By.XPATH, '//*[contains(text(), "THANK YOU")]')
    },
    "Next Draft":{
        "url": "https://managingeditor.substack.com/",
        "email":(By.XPATH, '//*[@id="entry"]/div[1]/div/div/div/div/div[2]/div/div[1]/form/div[1]/div/input'),
        "button" :(By.XPATH, '//*[@id="subscribe-inner"]/div[2]/form/button'),
        "confirmation": (By.XPATH, '//*[contains(text(), "confirm")]')
    },
    "Well+Good":{
        "url": "https://www.wellandgood.com/",
        "email":(By.XPATH, '//*[@id=":R6namvfff5b:"]'),
        "button" :(By.XPATH, '/html/body/footer/div/div[1]/div[1]/div[1]/div/div[1]/form/div[1]/label/div/div/button'),
        "confirmation": (By.XPATH, '//*[contains(text(), "Got it, you\'ve been added to our email list")]')
    },
     "Vulture":{
        "url": "https://www.vulture.com/promo/sign-up-for-the-vulture-newsletter.html",
        "email":(By.XPATH, '//*[@id="columnSubscribeEmail-495"]'),
        "button" :(By.XPATH, '/html/body/main/div[1]/section/div/aside/div/div/div[1]/form/input[4]'),
        "confirmation": (By.XPATH, '//*[contains(text(), "Thanks, you\'re all set!You\'ll receive the next newsletter in your inbox.")]')
    },
    "Sidebar":{
        "url": "https://sidebar.io/",
        "email":(By.XPATH, '//*[@id="wrapper"]/div/div[2]/div[2]/div/div[1]/div[1]/div[3]/div/div/div/div/div/div/input'),
        "button" :(By.XPATH, '//*[@id="wrapper"]/div/div[2]/div[2]/div/div[1]/div[1]/div[3]/div/div/div/div/div/button'),
        "confirmation": (By.XPATH, '//*[contains(text(), "Thanks for subscribing!")]')
    },
    "Light Hub":{
        "url": "https://link.lithub.com/join/signup",
        "email":(By.XPATH, '//*[@id="content"]/form/label[2]/input'),
        "button" :(By.XPATH, '//*[@id="content"]/form/input[2]'),
        "confirmation": (By.XPATH, '//*[contains(text(), "Thank you for signing up for our newsletter.")]')
    }
    
}    
# Function to set up the Chrome driver
def setup_driver():
    # Path to the file where the ChromeDriver path will be saved
    chromedriver_path_file = "chromedriver_path.txt"

    # Check if the ChromeDriver path is already saved
    if os.path.exists(chromedriver_path_file):
        with open(chromedriver_path_file, "r") as file:
            chromedriver_path = file.read().strip()
    else:
        try:
            # Automatically download and install the correct ChromeDriver
            chromedriver_path = ChromeDriverManager().install()
            
            # Save the ChromeDriver path to a file
            with open(chromedriver_path_file, "w") as file:
                file.write(chromedriver_path)
        except Exception as e:
            logging.error(f"Failed to install ChromeDriver: {e}")
            raise

    # Set up the Chrome driver using the saved path
    service = Service(chromedriver_path)
    options = webdriver.ChromeOptions()
    return webdriver.Chrome(service=service, options=options)

# Semaphore to limit the number of concurrent threads
max_concurrent_windows = 5  # Adjust this value based on your system's capacity
semaphore = threading.Semaphore(max_concurrent_windows)

# Function to subscribe to a single newsletter
from selenium.webdriver.common.action_chains import ActionChains  # Add this import

# Function to simulate human-like typing
def human_type(element, text):
    for character in text:
        element.send_keys(character)
        time.sleep(random.uniform(0.1, 0.2))  # Random delay between keystrokes
        
# Function to subscribe to a single newsletter
def subscribe_to_newsletter(email, newsletter_name, log_text):
    global Subscribed, good, bad, stop_flag
    if stop_flag:  # Check if the stop flag is set
        log_text.insert(tk.END, f"Subscription process stopped by user for {newsletter_name}.\n")
        return

    with semaphore:  # Acquire a semaphore slot
        if stop_flag:  # Check again after acquiring the semaphore
            log_text.insert(tk.END, f"Subscription process stopped by user for {newsletter_name}.\n")
            return

        newsletter = newsletters[newsletter_name]
        driver = setup_driver()
        driver.get(newsletter["url"])
        
        try:
            # Handle Forbes-specific two-button process
            if newsletter_name == "Forbes":
                button1 = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(newsletter["button1"])
                )
                button1.click()
                time.sleep(2)

            # Fill in the email input field
            email_input = WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located(newsletter["email"])
            )
            email_input.clear()
            human_type(email_input,email)
            # Scroll to the button (if necessary)
            if newsletter_name == "Dayspring":
                subscribe_button = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located(newsletter["button"])
                )
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", subscribe_button)
                time.sleep(1)  # Wait for the page to scroll

            # Click the appropriate button
            if newsletter_name == "Forbes":
                button2 = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(newsletter["button2"])
                )
                button2.click()
            else:
                subscribe_button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable(newsletter["button"])
                )
                subscribe_button.click()
                time.sleep(2)
            # Wait for the confirmation message
            WebDriverWait(driver, 10).until(
                EC.visibility_of_element_located(newsletter["confirmation"])
            )
            log_text.insert(tk.END, f"Successfully subscribed {email} to {newsletter_name}\n")
            Subscribed += 1
            good += 1

        except TimeoutException:
            log_text.insert(tk.END, f"Timeout: Could not subscribe {email} to {newsletter_name}. The page took too long to load.\n")
            bad += 1

        except NoSuchElementException:
            log_text.insert(tk.END, f"Element not found: Could not subscribe {email} to {newsletter_name}. The required element was missing.\n")
            bad += 1

        except WebDriverException as e:
            # Extract the main error message from the exception
            error_message = str(e).split("\n")[0]  # Take only the first line of the error
            log_text.insert(tk.END, f"WebDriver Error: Could not subscribe {email} to {newsletter_name}. {error_message}\n")
            bad += 1

        except Exception as e:
            # Handle any other exceptions
            error_message = str(e).split("\n")[0]  # Take only the first line of the error
            log_text.insert(tk.END, f"Error: Could not subscribe {email} to {newsletter_name}. {error_message}\n")
            bad += 1

        finally:
            driver.quit()

# Function to subscribe to newsletters
def subscribe_to_newsletters(emails, selected_newsletters, log_text):
    global stop_flag
    threads = []

    for email in emails:
        if stop_flag:  # Check if the stop flag is set
            log_text.insert(tk.END, "Subscription process stopped by user.\n")
            break

        for newsletter_name in selected_newsletters:
            if stop_flag:  # Check if the stop flag is set
                log_text.insert(tk.END, "Subscription process stopped by user.\n")
                break

            # Start a new thread for each newsletter
            thread = threading.Thread(target=subscribe_to_newsletter, args=(email, newsletter_name, log_text))
            threads.append(thread)
            thread.start()

    # Wait for all threads to complete
    for thread in threads:
        thread.join()

    # Notify the user that the process is complete or stopped
    if stop_flag:
        messagebox.showinfo("Process Stopped", "The subscription process was stopped by the user.")
    else:
        messagebox.showinfo("Script Completed", "The subscription process is finished.\nCheck the log for details.")

# Function to start the subscription process
def start_subscription(file_path, selected_newsletters, log_text):
    def run_subscription():
        if file_path:
            try:
                with open(file_path, 'r') as file:
                    emails = [line.strip() for line in file if line.strip()]
            except FileNotFoundError:
                log_text.insert(tk.END, f"{file_path} file not found.\n")
            else:
                subscribe_to_newsletters(emails, selected_newsletters, log_text)
        else:
            log_text.insert(tk.END, "No file selected.\n")

    # Run the subscription process in a separate thread
    global subscription_thread
    subscription_thread = threading.Thread(target=run_subscription)
    subscription_thread.start()

# Function to stop the subscription process
def stop_subscription():
    global stop_flag
    stop_flag = True
    messagebox.showinfo("Process Stopped", "The subscription process will stop after the current email.")

# Tkinter GUI
# Tkinter GUI
class NewsletterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Newsletter Subscriber made by Anouar.opm53")
        self.root.geometry("12100x720+100+50")

        # File selection
        self.file_label = tk.Label(root, text="Select emails.txt file:")
        self.file_label.pack(pady=10)
        self.file_button = tk.Button(root, text="Browse", command=self.select_file)
        self.file_button.pack(pady=5)

        # Newsletter selection
        self.newsletter_label = tk.Label(root, text="Select newsletters to subscribe to:")
        self.newsletter_label.pack(pady=10)

        # Create a frame to hold the checkboxes
        self.checkbox_frame = tk.Frame(root)
        self.checkbox_frame.pack()

        self.newsletter_vars = {}
        for i, newsletter in enumerate(newsletters.keys()):
            var = tk.BooleanVar()
            self.newsletter_vars[newsletter] = var
            # Calculate row and column for grid placement
            row = i // 2  # Integer division to determine row
            col = i % 2   # Modulo to determine column (0 or 1)
            cb = tk.Checkbutton(self.checkbox_frame, text=newsletter, variable=var, anchor="w")
            cb.grid(row=row, column=col, sticky="w", padx=20, pady=5)  # Align checkboxes to the left

        # Log output
        self.log_label = tk.Label(root, text="Log Output:")
        self.log_label.pack(pady=10)
        self.log_text = scrolledtext.ScrolledText(root, height=10, width=70)
        self.log_text.pack(pady=10)

        # Start button
        self.start_button = tk.Button(root, text="Start Subscription", command=self.start)
        self.start_button.pack(pady=20)

        # Stop button
        self.stop_button = tk.Button(root, text="Stop Subscription", command=self.stop_subscription)
        self.stop_button.pack(pady=10)

    def select_file(self):
        file_path = filedialog.askopenfilename(title="Select emails.txt", filetypes=(("Text files", "*.txt"), ("All files", "*.*")))
        if file_path:
            self.file_label.config(text=f"Selected file: {file_path}")
            self.file_path = file_path

    def start(self):
        global stop_flag
        stop_flag = False  # Reset the stop flag
        selected_newsletters = [newsletter for newsletter, var in self.newsletter_vars.items() if var.get()]
        if not selected_newsletters:
            messagebox.showwarning("No Selection", "Please select at least one newsletter.")
            return
        if not hasattr(self, 'file_path'):
            messagebox.showwarning("No File", "Please select an emails.txt file.")
            return
        start_subscription(self.file_path, selected_newsletters, self.log_text)

    def stop_subscription(self):
        global stop_flag
        stop_flag = True
        messagebox.showinfo("Process Stopped", "The subscription process will stop after the current email.")

# Run the application
if __name__ == "__main__":
    root = tk.Tk()
    app = NewsletterApp(root)
    root.mainloop()