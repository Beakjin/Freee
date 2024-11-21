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
Freeecredentials_file = "Freeecredentials.json"
Cliqcredentials_file = "Cliqcredentials.json"
# Function to load or ask for credentials
def get_credentials():
    if os.path.exists(Freeecredentials_file):
        use_saved = messagebox.askyesno("ログイン情報入力確認", "以前に入力したIDとパスワードを使用しますか？")
        if use_saved:
            with open(Freeecredentials_file, 'r') as f:
                credentials = json.load(f)
            return credentials.get("user_id"), credentials.get("user_password")
    
    # Prompt for new credentials if no file exists or user selects "No"
    user_id = simpledialog.askstring("Freee", "Freee ID")
    user_password = simpledialog.askstring("Freee", "Freee PW", show='*')
    
    # Save credentials to file
    with open(Freeecredentials_file, 'w') as f:
        json.dump({"user_id": user_id, "user_password": user_password}, f)
    
    return user_id, user_password

# Get user ID and password
user_id, user_password = get_credentials()

def get_credentials():
    if os.path.exists(Cliqcredentials_file):
        use_saved = messagebox.askyesno("ログイン情報入力確認", "以前に入力したIDとパスワードを使用しますか？")
        if use_saved:
            with open(Cliqcredentials_file, 'r') as f:
                credentials = json.load(f)
            return credentials.get("user_id"), credentials.get("user_password")
    
    # Prompt for new credentials if no file exists or user selects "No"
    Cliquser_id = simpledialog.askstring("Cliq", "Cliq ID")
    Cliquser_password = simpledialog.askstring("Cliq", "Cliq PW", show='*')
    
    # Save credentials to file
    with open(Cliqcredentials_file, 'w') as f:
        json.dump({"user_id": Cliquser_id, "user_password": Cliquser_password}, f)
    
    return Cliquser_id, Cliquser_password

# Get user ID and password
Cliquser_id, Cliquser_password = get_credentials()
print(f"Freee ユーザーID: {user_id}")
print(f"Freee ユーザーPW: {user_password}")
print(f"Cliq ユーザーID: {Cliquser_id}")
print(f"Cliq ユーザーPW: {Cliquser_password}")




# Create window
window = tk.Tk()
window.title("freee出退勤打刻")
window.geometry("400x400")

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


def check_work():
    

    try:
    

        today = datetime.now().strftime("%Y-%m-%d %H:%M")

        options = Options()
        options.add_argument("--headless")  # ヘッドレスモードでバックグラウンド実行
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-gpu")

        # WebDriverの起動
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

        # 出勤ボタンの有無を確認
        try:
            button = driver.find_element(By.XPATH, '//span[text()="出勤"]')
            print("出勤ボタンが見つかりました。")
            messagebox.showinfo("Freee", "Freeeの使用準備が完了しました。")
            return  # 出勤ボタンがある場合は何もしない
        except:
            print("出勤ボタンが見つかりません。修正ボタンをクリックします。")
            

        # 修正ボタンをクリック
        fix_button = driver.find_element(By.XPATH, '//span[text()="修正"]')
        fix_button.click()
        time.sleep(5)

        # 指定したXPathのデータを取得
        data1, data2, data3 = None, None, None
        try:
            data1 = driver.find_element(By.XPATH, '(//span[@class="vb-tableListCell__text" and contains(text(), ":")])[1]').text
            print(data1)
        except :
            print("データ1が見つかりませんでした。")
            data1="1"
            print(data1)
    
        try:
            data2 = driver.find_element(By.XPATH, '(//span[@class="vb-tableListCell__text" and contains(text(), ":")])[2]').text
            print(data2)
        except :
            print("データ2が見つかりませんでした。")
            data2="1"
            print(data2)
        
        try:
            data3 = driver.find_element(By.XPATH, '(//span[@class="vb-tableListCell__text" and contains(text(), ":")])[3]').text
            print(data3)
        except :
            print("データ3が見つかりませんでした。")
            data3="1"
            print(data3)
        
        
        today1 = today.split()[1]  # startworkから時間部分のみ取得（例えば、"08:00:00"）
        todayshu = today.split()[0] + ' ' + data1
        
        today1 = today.split()[1]  # startworkから時間部分のみ取得（例えば、"08:00:00"）
        todaykyustart = today.split()[0] + ' ' + data2

        today1 = today.split()[1]  # startworkから時間部分のみ取得（例えば、"08:00:00"）
        todaykyuend = today.split()[0] + ' ' + data3

        # times_labelにデータを表示
        text_content = ""
        if data1.strip() != '1':
           text_content += f"出勤時刻: {todayshu}\n"

        if data2.strip() != '1':
           text_content += f"休憩開始: {todaykyustart}\n"

        if data3.strip() != '1':
           text_content += f"休憩終了: {todaykyuend}"

        times_label.config(text=text_content)
        messagebox.showinfo("Freee", "Freeeの使用準備が完了しました。")

    except Exception as e:
        print("エラーが発生しました:", e)

    finally:
        # WebDriverを閉じる
        try:
            driver.quit()
        except:
            pass

messagebox.showinfo("Freee", "OKを押すとステータス確認を行います。しばらくお待ちください。")

# アプリ起動後すぐに処理を実行
window.after(100, check_work)


# Functions for each action button
def on_ok():
    # 現在の出勤時刻を取得
    shukkin = datetime.now().strftime("%Y-%m-%d %H:%M")
    
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
    start = break_start_time.strftime("%Y-%m-%d %H:%M")
    
    times_label.config(text=times_label.cget("text") + f"休憩開始: {start}\n")
    window.after(1000, on_Cliqkyuukei_start)
    window.after(3600000, on_kyuukei_end)

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
    taikin = datetime.now().strftime("%Y-%m-%d %H:%M")

    window.after(1000, on_Cliqwork_end)

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
    
def on_Cliqk_home():
    print(f"Cliq ユーザーID: {user_id}")
    print(f"Cliq ユーザーID: {user_password}")
    print(f"Cliq ユーザーID: {Cliquser_id}")
    print(f"Cliq ユーザーID: {Cliquser_password}")

    
    driver = webdriver.Chrome()
    driver.get(Cliq_url)
    time.sleep(3)
    
    login_field = driver.find_element(By.XPATH, '//*[@id="login_id"]')
    login_field.send_keys(Cliquser_id)
    next_button = driver.find_element(By.XPATH, '//span[text()="次へ"]')
    next_button.click()    
    time.sleep(5)
    password_field = driver.find_element(By.XPATH, '//*[@id="password"]')
    password_field.send_keys(Cliquser_password)
    time.sleep(5)
    login_button = driver.find_element(By.XPATH, '//*[@id="nextbtn"]/span')
    login_button.click()
    window.after()
    

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
    login_field.send_keys(Cliquser_id)
    next_button = driver.find_element(By.XPATH, '//span[text()="次へ"]')
    next_button.click()    
    time.sleep(5)
    password_field = driver.find_element(By.XPATH, '//*[@id="password"]')
    password_field.send_keys(Cliquser_password)
    time.sleep(5)
    login_button = driver.find_element(By.XPATH, '//*[@id="nextbtn"]/span')
    login_button.click()
    time.sleep(10)
    try:
        # 要素を探してクリック
        button = driver.find_element(By.XPATH, '//a[@class="succbutton" and text()="確認する"]')
        button.click()
        print("ボタンをクリックしました。")
    except:
        # 要素が見つからない場合はスキップ
        print("ボタンが見つかりませんでした。処理をスキップします。")
    time.sleep(5)
    max_retries = 5  # 最大試行回数
    attempts = 0
    while attempts < max_retries:
        try:
             continue_button = driver.find_element(By.XPATH, '//a[text()="同意する"]')
             continue_button.click()
             print("次のボタンがクリックされました。")
             break
        except:
             attempts += 1
             print("次のボタンが見つかりませんでした。")
             if attempts < max_retries:
                time.sleep(2)  # 2秒待機してから再試行
             else:
                 print("最大試行回数に達しました。処理を終了します。")

    time.sleep(8)
    status_text = driver.find_element(By.XPATH, '//span[@class="meal-brk mR6"]')
    status_text.click()
    time.sleep(1)

def on_Cliqkyuukei_end():

    options = Options()
    options.add_argument("--headless")  # ヘッドレスモードでバックグラウンド実行
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    
    driver = webdriver.Chrome(options=options)
    driver.get(Cliq_url)
    time.sleep(3)

    login_field = driver.find_element(By.XPATH, '//*[@id="login_id"]')
    login_field.send_keys(Cliquser_id)
    next_button = driver.find_element(By.XPATH, '//span[text()="次へ"]')
    next_button.click()    
    time.sleep(5)
    password_field = driver.find_element(By.XPATH, '//*[@id="password"]')
    password_field.send_keys(Cliquser_password)
    time.sleep(5)
    login_button = driver.find_element(By.XPATH, '//*[@id="nextbtn"]/span')
    login_button.click()
    time.sleep(10)
    try:
        # 要素を探してクリック
        button = driver.find_element(By.XPATH, '//a[@class="succbutton" and text()="確認する"]')
        button.click()
        print("ボタンをクリックしました。")
    except:
        # 要素が見つからない場合はスキップ
        print("ボタンが見つかりませんでした。処理をスキップします。")
    time.sleep(5)
    max_retries = 5  # 最大試行回数
    attempts = 0
    while attempts < max_retries:
        try:
             continue_button = driver.find_element(By.XPATH, '//a[text()="同意する"]')
             continue_button.click()
             print("次のボタンがクリックされました。")
             break
        except:
             attempts += 1
             print("次のボタンが見つかりませんでした。")
             if attempts < max_retries:
                time.sleep(2)  # 2秒待機してから再試行
             else:
                 print("最大試行回数に達しました。処理を終了します。")
    time.sleep(8)
    status_text = driver.find_element(By.XPATH, "//div[contains(@class, 'laptop-work')]")
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
    login_field.send_keys(Cliquser_id)
    next_button = driver.find_element(By.XPATH, '//span[text()="次へ"]')
    next_button.click()    
    time.sleep(5)
    password_field = driver.find_element(By.XPATH, '//*[@id="password"]')
    password_field.send_keys(Cliquser_password)
    time.sleep(5)
    login_button = driver.find_element(By.XPATH, '//*[@id="nextbtn"]/span')
    login_button.click()
    time.sleep(10)
    try:
        # 要素を探してクリック
        button = driver.find_element(By.XPATH, '//a[@class="succbutton" and text()="確認する"]')
        button.click()
        print("ボタンをクリックしました。")
    except:
        # 要素が見つからない場合はスキップ
        print("ボタンが見つかりませんでした。処理をスキップします。")

    max_retries = 5  # 最大試行回数
    attempts = 0
    while attempts < max_retries:
        try:
             continue_button = driver.find_element(By.XPATH, '//a[text()="同意する"]')
             continue_button.click()
             print("次のボタンがクリックされました。")
             break
        except:
             attempts += 1
             print("次のボタンが見つかりませんでした。")
             if attempts < max_retries:
                time.sleep(2)  # 2秒待機してから再試行
             else:
                 print("最大試行回数に達しました。処理を終了します。")
    
    time.sleep(8)
    status_text = driver.find_element(By.XPATH, '//span[@class="slider"]')
    status_text.click()
    time.sleep(1)

def on_Cliqwork_end():

    options = Options()
    options.add_argument("--headless")  # ヘッドレスモードでバックグラウンド実行
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    
    driver = webdriver.Chrome(options=options)
    driver.get(Cliq_url)
    time.sleep(3)
    
    login_field = driver.find_element(By.XPATH, '//*[@id="login_id"]')
    login_field.send_keys(Cliquser_id)
    next_button = driver.find_element(By.XPATH, '//span[text()="次へ"]')
    next_button.click()    
    time.sleep(5)
    password_field = driver.find_element(By.XPATH, '//*[@id="password"]')
    password_field.send_keys(Cliquser_password)
    time.sleep(5)
    login_button = driver.find_element(By.XPATH, '//*[@id="nextbtn"]/span')
    login_button.click()
    
    time.sleep(5)
    try:
        # 要素を探してクリック
        button = driver.find_element(By.XPATH, '//a[@class="succbutton" and text()="確認する"]')
        button.click()
        print("ボタンをクリックしました。")
    except:
        # 要素が見つからない場合はスキップ
        print("ボタンが見つかりませんでした。処理をスキップします。")

    time.sleep(5)
    max_retries = 5  # 最大試行回数
    attempts = 0
    while attempts < max_retries:
        try:
             continue_button = driver.find_element(By.XPATH, '//a[text()="同意する"]')
             continue_button.click()
             print("次のボタンがクリックされました。")
             break
        except:
             attempts += 1
             print("次のボタンが見つかりませんでした。")
             if attempts < max_retries:
                time.sleep(2)  # 2秒待機してから再試行
             else:
                 print("最大試行回数に達しました。処理を終了します。")
    time.sleep(5)

    time.sleep(5)
    status_text = driver.find_element(By.XPATH, '//span[@class="slider"]')
    status_text.click()
    time.sleep(1)
    messagebox.showinfo("Freee", "今日も一日お疲れ様でした。")
    time.sleep(1)
    window.quit()

def on_kyuukei_end():
    global end
    end = datetime.now().strftime("%Y-%m-%d %H:%M")
    times_label.config(text=times_label.cget("text") + f"休憩終了: {end}\n")
    #window.after(1000, on_Cliqkyuukei_end)


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
    end = datetime.now().strftime("%Y-%m-%d %H:%M")
    
    
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
                                         "ホーム : Freeeのホーム画面が表示されます。\n"
                                         "退勤 : 退勤時にボタンを押すと打刻が終了されます。\n"
                              , wraplength=350, justify="left")
    description_label1.pack(pady=2)

    description_label2 = tk.Label(description_window, text="Cliq連携", wraplength=350, justify="left", fg="red")
    description_label2.pack(pady=2)

    description_label3 = tk.Label(description_window, text="CliqとFreeeのアカウント情報を同一しておくと\n"
                                    "出勤・休憩開始・退勤の際に連携してステータスを変更します。\n"

                            , wraplength=350, justify="left")
    description_label3.pack(pady=2)

    description_label4 = tk.Label(description_window, text="休憩について\n", wraplength=350, justify="left", fg="red")
    description_label4.pack(pady=2)

    description_label5 = tk.Label(description_window, text="休憩時、ツールを開いたままにしておくと\n"
                                  "休憩開始60分後に自動で休憩終了を打刻します。\n"

                                    , wraplength=350, justify="left")
    description_label5.pack(pady=2)

# Label to display times
times_label = tk.Label(window, text="", justify="left")
times_label.pack(pady=10)

# Create buttons
ok_button = tk.Button(window, text="出勤", command=on_ok, width=5, height=1)
ok_button.pack(side="left", padx=10, pady=10)

start_button = tk.Button(window, text="休憩開始", command=on_start, width=7, height=1)
start_button.pack(side="left", padx=10, pady=10)

home_button = tk.Button(window, text="Freeeホーム", command=on_home, width=8, height=1)
home_button.pack(side="left", padx=15, pady=15)

kyuukei_end_button = tk.Button(window, text="Cliqホーム", command=on_Cliqk_home, width=8, height=1)
kyuukei_end_button.pack(side="left", padx=10, pady=10)

end_button = tk.Button(window, text="退勤", command=on_end, width=5, height=1)
end_button.pack(side="left", padx=10, pady=10)




#kyuukei_end_button = tk.Button(window, text="休憩終了", command=on_kyuukei_end, width=5, height=1)
#kyuukei_end_button.pack(side="right", padx=10, pady=10)

description_button = tk.Button(window, text="説明を見る", command=show_description_window)
description_button.place(x=350, y=10, anchor="ne")


# Display the window
window.mainloop()
