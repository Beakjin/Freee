import threading
import tkinter as tk
from tkinter import messagebox, simpledialog
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
import time
import os
import json

# Freee login URL
free_url = "https://accounts.secure.freee.co.jp/sessions/new?redirect_url=https%3A%2F%2Fp.secure.freee.co.jp%2Fusers%2Fafter_login%3Fhash%3D&service_name=payroll&sign_up_url=https%3A%2F%2Faccounts.secure.freee.co.jp%2Fsign_up%3Finitial_service%3Dpayroll%26login_url%3Dhttps%253A%252F%252Fp.secure.freee.co.jp%252F%26prefer%3Demail%26redirect_url%3Dhttps%253A%252F%252Fp.secure.freee.co.jp%252Factivation%26service_name%3Dpayroll%26verification_type%3Dcode"
Cliq_url = "https://accounts.zoho.jp/signin?servicename=ZohoChat&signupurl=https://www.zoho.com/cliq/signup.html"
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
    # 現在の出勤時刻を取得
    shukkin = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    window.after(500, on_Cliqwork_start)
    
    options = Options()
    options.add_argument("--headless")  # ヘッドレスモードでバックグラウンド実行
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")

    driver = webdriver.Chrome(options=options)
    driver.get(free_url)
    time.sleep(3)

    # ログイン処理
    login_field = driver.find_element(By.XPATH, '//*[@id="loginIdField"]')
    login_field.send_keys(user_id)
    password_field = driver.find_element(By.XPATH, '//*[@id="passwordField"]')
    password_field.send_keys(user_password)
    login_button = driver.find_element(By.XPATH, '//span[text()="ログイン"]')
    login_button.click()
    time.sleep(5)

    # 出勤ボタンをクリックするが、出勤ボタンがない場合は修正ボタンをクリックして値を取得
    try:
        # 出勤ボタンを探してクリック
        button = driver.find_element(By.XPATH, '//span[text()="出勤"]')
        button.click()
        time.sleep(5)

        times_label.config(text=times_label.cget("text") + f"出勤時刻: {shukkin}\n")
    except:
        # 出勤ボタンが見つからない場合、修正ボタンをクリック
        print("出勤ボタンが見つかりません。修正ボタンをクリックし、値を取得します。")
        try:
            # 修正ボタンをクリック
            fix_button = driver.find_element(By.XPATH, '//span[text()="修正"]')
            fix_button.click()
            time.sleep(8)

            # 値を取得
            element = driver.find_element(By.XPATH, '//span[@class="vb-tableListCell__text" and contains(text(), ":")]')
            startwork = element.text  # 値をstartwork変数に保存
            print("取得した値:", startwork)

            # 取得した時間部分をshukkinの時間部分と置き換える
            startwork_time = shukkin.split()[1]  # startworkから時間部分のみ取得（例えば、"08:00:00"）
            shukkin = shukkin.split()[0] + ' ' + startwork  # 日付部分はそのままで時間部分を置換

            # 時間が置き換えられた出勤時刻を表示
            times_label.config(text=times_label.cget("text") + f"出勤時刻: {shukkin}\n")

        except Exception as e:
            print("修正ボタンまたは指定のXPathが見つかりません:", e)

def on_start():
    global break_start_time
    break_start_time = datetime.now()
    start = break_start_time.strftime("%Y-%m-%d %H:%M:%S")
    
    times_label.config(text=times_label.cget("text") + f"休憩開始: {start}\n")
    window.after(1000, on_Cliqkyuukei_start)
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


def on_end():
    global taikin
    taikin = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    window.after(1000, on_Cliqwork_start)

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

def on_Cliqkyuukei_start():

    options = Options()
    options.add_argument("--headless")  # ヘッドレスモードでバックグラウンド実行
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    
    driver = webdriver.Chrome(options=options)
    driver.get(Cliq_url)
    time.sleep(3)

    login_field = driver.find_element(By.XPATH, '//*[@id="login_id"]')
    login_field.send_keys(user_id)
    next_button = driver.find_element(By.XPATH, '//span[text()="次へ"]')
    next_button.click()    
    time.sleep(5)
    password_field = driver.find_element(By.XPATH, '//*[@id="password"]')
    password_field.send_keys(user_password)
    time.sleep(5)
    login_button = driver.find_element(By.XPATH, '//*[@id="nextbtn"]/span')
    login_button.click()
    time.sleep(5)
    status_text = driver.find_element(By.XPATH, '//span[@class="meal-brk mR6"]')
    status_text.click()
    time.sleep(1)

def on_Cliqwork_start():

    options = Options()
    options.add_argument("--headless")  # ヘッドレスモードでバックグラウンド実行
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    
    driver = webdriver.Chrome(options=options)
    driver.get(Cliq_url)
    time.sleep(3)

    login_field = driver.find_element(By.XPATH, '//*[@id="login_id"]')
    login_field.send_keys(user_id)
    next_button = driver.find_element(By.XPATH, '//span[text()="次へ"]')
    next_button.click()    
    time.sleep(5)
    password_field = driver.find_element(By.XPATH, '//*[@id="password"]')
    password_field.send_keys(user_password)
    time.sleep(5)
    login_button = driver.find_element(By.XPATH, '//*[@id="nextbtn"]/span')
    login_button.click()
    time.sleep(5)
    status_text = driver.find_element(By.XPATH, '//span[@class="slider"]')
    status_text.click()
    time.sleep(1)

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

    window.after()
    
def show_description_window():
    # 新しいウィンドウ（Toplevel）を作成
    description_window = tk.Toplevel(window)
    description_window.title("説明文")
    description_window.geometry("500x400")

    # 説明文を表示する複数のLabelを設定
    description_label2 = tk.Label(description_window, text="ツールについて\n", wraplength=350, justify="left", fg="blue")
    description_label2.pack(pady=2)

    description_label1 = tk.Label(description_window, text="このボタンの機能は以下の通りです。\n"
                                         "出勤 : 出勤時にボタンを押すと打刻が開始されます。\n"
                                         "休憩開始 : 勤務時間内での休憩開始時間が登録されます。\n"
                                         "退勤 : 退勤時にボタンを押すと打刻が終了されます。\n"
                              , wraplength=350, justify="left")
    description_label1.pack(pady=2)

    description_label2 = tk.Label(description_window, text="注意 : 出勤時について\n", wraplength=350, justify="left", fg="red")
    description_label2.pack(pady=2)

    description_label3 = tk.Label(description_window, text="出勤時に、前回の出勤実績が残っている（退勤を押せていない）場合は\n"
                                    "Freeeの仕様上退勤ボタンのみが残り、出勤が押せないため\n"
                                    "手動で退勤処理を行った後、ツール上で出勤ボタンを押下してください。\n"
                            , wraplength=350, justify="left")
    description_label3.pack(pady=2)

    description_label4 = tk.Label(description_window, text="休憩時間について\n", wraplength=350, justify="left", fg="red")
    description_label4.pack(pady=2)

    description_label5 = tk.Label(description_window, text="出勤後5時間たっても、休憩開始をしていない場合休憩実施確認が表示されます。\n"
                                    "休憩をとっていないと毎月の給与精算処理が正常に行えないため必ず休憩処理をお願いします。\n"
                                    "（休憩確認表示時すでに実際に休憩をとっている場合は、～～を押してください。自動で休憩処理を行います。）\n"
                                    , wraplength=350, justify="left")
    description_label5.pack(pady=2)

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

description_button = tk.Button(window, text="説明を見る", command=show_description_window)
description_button.place(x=310, y=10, anchor="ne")


# Display the window
window.mainloop()
