import tkinter as tk
from tkinter import messagebox, simpledialog
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import os
import json
#test

# Freee login URL
free_url = "https://accounts.secure.freee.co.jp/sessions/new?redirect_url=https%3A%2F%2Fp.secure.freee.co.jp%2Fusers%2Fafter_login%3Fhash%3D&service_name=payroll&sign_up_url=https%3A%2F%2Faccounts.secure.freee.co.jp%2Fsign_up%3Finitial_service%3Dpayroll%26login_url%3Dhttps%253A%252F%252Fp.secure.freee.co.jp%252F%26prefer%3Demail%26redirect_url%3Dhttps%253A%252F%252Fp.secure.freee.co.jp%252Factivation%26service_name%3Dpayroll%26verification_type%3Dcode"

# File to store credentials
credentials_file = "credentials.json"

# Function to load or ask for credentials
def get_credentials():
    if os.path.exists(credentials_file):
        use_saved = messagebox.askyesno("ログイン情報入力確認", "以前に入力したIDとパスワードを使用しますか？")
        if use_saved:
            with open(credentials_file, 'r') as f:
                credentials = json.load(f)
            return credentials.get("user_id"), credentials.get("user_password")
    
    # Prompt for new credentials if no file exists or user selects "No"
    user_id = simpledialog.askstring("ユーザーID", "IDを入力してください:")
    user_password = simpledialog.askstring("パスワード", "パスワードを入力してください:", show='*')
    
    # Save credentials to file
    with open(credentials_file, 'w') as f:
        json.dump({"user_id": user_id, "user_password": user_password}, f)
    
    return user_id, user_password

# Get user ID and password
user_id, user_password = get_credentials()

# Create window
window = tk.Tk()
window.title("freee出退勤打刻")
window.geometry("330x400")

# Get current date
current_date = datetime.now().strftime("%Y-%m-%d")

# Create labels
label = tk.Label(window, text="頑張りましょう")
label.pack(pady=5)

# Label to display the current date
date_label = tk.Label(window, text=current_date)
date_label.pack(pady=5)

# Initialize start time for break
break_start_time = None

# Functions for each action button
def on_ok():
    shukkin = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    times_label.config(text=times_label.cget("text") + f"出勤時刻: {shukkin}\n")
    
    options = Options()
    options.add_argument("--headless")  # ヘッドレスモードでバックグラウンド実行
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")

    driver = webdriver.Chrome(options=options)
    driver.get(free_url)
    time.sleep(3)

    login_field = driver.find_element(By.XPATH, '//*[@id="loginIdField"]')
    login_field.send_keys(user_id)
    password_field = driver.find_element(By.XPATH, '//*[@id="passwordField"]')
    password_field.send_keys(user_password)
    login_button = driver.find_element(By.XPATH, '//span[text()="ログイン"]')
    login_button.click()
    time.sleep(5)

    button = driver.find_element(By.XPATH, '//span[text()="出勤"]')
    button.click()
    time.sleep(5)

def on_start():
    global break_start_time
    break_start_time = datetime.now()
    start = break_start_time.strftime("%Y-%m-%d %H:%M:%S")
    
    times_label.config(text=times_label.cget("text") + f"休憩開始: {start}\n")
    check_break_time()
    window.after(3540000, on_kyuukei_end)

    options = Options()
    options.add_argument("--headless")  # ヘッドレスモードでバックグラウンド実行
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")

    driver = webdriver.Chrome(options=options)
    driver.get(free_url)
    time.sleep(3)

    login_field = driver.find_element(By.XPATH, '//*[@id="loginIdField"]')
    login_field.send_keys(user_id)
    password_field = driver.find_element(By.XPATH, '//*[@id="passwordField"]')
    password_field.send_keys(user_password)
    login_button = driver.find_element(By.XPATH, '//span[text()="ログイン"]')
    login_button.click()
    time.sleep(5)

    button2 = driver.find_element(By.XPATH, '//span[text()="休憩開始"]')
    button2.click()
    time.sleep(5)

def check_break_time():
    # Check if 50 minutes have passed since the break started
    if break_start_time:
        elapsed_time = datetime.now() - break_start_time
        if elapsed_time >= timedelta(minutes=50):
            # Change the color of the last entry in times_label to red
            times_label.config(fg="red")
        else:
            # Check again after 1 minute
            window.after(60000, check_break_time)

def on_end():
    global taikin
    taikin = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    options = Options()
    options.add_argument("--headless")  # ヘッドレスモードでバックグラウンド実行
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")

    driver = webdriver.Chrome(options=options)
    driver.get(free_url)
    time.sleep(3)

    login_field = driver.find_element(By.XPATH, '//*[@id="loginIdField"]')
    login_field.send_keys(user_id)
    password_field = driver.find_element(By.XPATH, '//*[@id="passwordField"]')
    password_field.send_keys(user_password)
    login_button = driver.find_element(By.XPATH, '//span[text()="ログイン"]')
    login_button.click()
    time.sleep(5)

    button1 = driver.find_element(By.XPATH, '//span[text()="退勤"]')
    button1.click()
    time.sleep(5)
    window.quit()

def on_kyuukei_end():
    global end
    end = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    times_label.config(text=times_label.cget("text") + f"休憩終了: {end}\n")

    options = Options()
    options.add_argument("--headless")  # ヘッドレスモードでバックグラウンド実行
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")

    driver = webdriver.Chrome(options=options)
    driver.get(free_url)
    time.sleep(3)

    login_field = driver.find_element(By.XPATH, '//*[@id="loginIdField"]')
    login_field.send_keys(user_id)
    password_field = driver.find_element(By.XPATH, '//*[@id="passwordField"]')
    password_field.send_keys(user_password)
    login_button = driver.find_element(By.XPATH, '//span[text()="ログイン"]')
    login_button.click()
    time.sleep(5)

    button3 = driver.find_element(By.XPATH, '//span[text()="休憩終了"]')
    button3.click()
    time.sleep(5)

def on_home():
    global home
    end = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    
    driver = webdriver.Chrome()
    driver.get(free_url)
    time.sleep(3)

    login_field = driver.find_element(By.XPATH, '//*[@id="loginIdField"]')
    login_field.send_keys(user_id)
    password_field = driver.find_element(By.XPATH, '//*[@id="passwordField"]')
    password_field.send_keys(user_password)
    login_button = driver.find_element(By.XPATH, '//span[text()="ログイン"]')
    login_button.click()

    while driver.window_handles:  # ブラウザウィンドウが存在するかを確認
            time.sleep(1) 
    


# Label to display times
times_label = tk.Label(window, text="", justify="left")
times_label.pack(pady=10)

# Create buttons
ok_button = tk.Button(window, text="出勤", command=on_ok)
ok_button.pack(side="left", padx=10, pady=10)

start_button = tk.Button(window, text="休憩開始", command=on_start)
start_button.pack(side="left", padx=10, pady=10)

home_button = tk.Button(window, text="ホーム", command=on_home)
home_button.pack(side="left", padx=10, pady=10)

end_button = tk.Button(window, text="退勤", command=on_end)
end_button.pack(side="right", padx=10, pady=10)

kyuukei_end_button = tk.Button(window, text="休憩終了", command=on_kyuukei_end)
kyuukei_end_button.pack(side="right", padx=10, pady=10)

# Display the window
window.mainloop()
