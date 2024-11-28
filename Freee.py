import tkinter as tk
from tkinter import messagebox, simpledialog
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time
import os
import json
from tkinter import font
import webbrowser
import sys

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
window.geometry("410x400")

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

# JSONファイルのパスを設定

startworkt_click_time = "startwork_click_time.json"
startrest_click_time = "startrest_click_time.json"
endrest_click_time = "endrest_click_time.json"
endwork_click_time = "endwork_click_time.json"


# JSONファイルを読み込む
# startwork
try:
    with open(startworkt_click_time, 'r', encoding='utf-8') as file:
        startwork_click_timedata = json.load(file)
    print("startwork:")
    print(json.dumps(startwork_click_timedata, ensure_ascii=False, indent=4))
except FileNotFoundError:
    print(f"{startworkt_click_time} が見つかりません。ファイルのパスを確認してください。")
except json.JSONDecodeError:
    print(f"{startworkt_click_time} は正しいJSON形式ではありません。")


# startrest
try:
    with open(startrest_click_time, 'r', encoding='utf-8') as file:
        startrest_click_timedata = json.load(file)
    print("startrest:")
    print(json.dumps(startrest_click_timedata, ensure_ascii=False, indent=4))
except FileNotFoundError:
    print(f"{startrest_click_time} が見つかりません。ファイルのパスを確認してください。")
except json.JSONDecodeError:
    print(f"{startrest_click_time} は正しいJSON形式ではありません。")


# endrest
try:
    with open(endrest_click_time, 'r', encoding='utf-8') as file:
        endrest_click_timedata = json.load(file)
    print("endrest:")
    print(json.dumps(endrest_click_timedata, ensure_ascii=False, indent=4))
except FileNotFoundError:
    print(f"{endrest_click_time} が見つかりません。ファイルのパスを確認してください。")
except json.JSONDecodeError:
    print(f"{endrest_click_time} は正しいJSON形式ではありません。")

# endwork
try:
    with open(endwork_click_time, 'r', encoding='utf-8') as file:
        endwork_click_timedata = json.load(file)
    print("endwork:")
    print(json.dumps(endwork_click_timedata, ensure_ascii=False, indent=4))
except FileNotFoundError:
    print(f"{endwork_click_time} が見つかりません。。")
except json.JSONDecodeError:
    print(f"{endwork_click_time} は正しいJSON形式ではありません。")




# Functions for each action button
def on_ok():
    new_window = tk.Toplevel(window)
    new_window.title("出勤時間を選択")
    new_window.geometry("300x100")

    def select_9am():
        print("9時出勤が選択されました")
        shukkin = datetime.now().strftime("%Y-%m-%d %H:%M")

      
        options = Options()
        options.add_argument("--headless")  # ヘッドレスモードでバックグラウンド実行
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
        try:
            elements = driver.find_elements(By.XPATH, '//span[text()="ログイン"]')
            if elements:
                root = tk.Tk()  # tkinterのウィンドウ作成
                root.withdraw()  # メインウィンドウを非表示にする
                messagebox.showerror("エラー", "Freeeにてログインが失敗しました。")
                root.destroy()  # ダイアログを閉じた後、ウィンドウを破
                return 
            else:
                print("'ログイン'要素は存在しません。正常に処理を続行します。")
        except Exception as e:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("例外エラー", f"エラーが発生しました: {str(e)}")
            root.destroy()
            return 
            
        try:
            button = driver.find_element(By.XPATH, '//span[text()="出勤"]')
            button.click()
        except:
            print("エラー: '出勤'ボタンが見つかりませんでした。処理をスキップします。")
            root = tk.Tk()  # tkinterのウィンドウ作成
            root.withdraw()  # メインウィンドウを非表示にする
            messagebox.showerror("エラー", "Freeeにて出勤ボタンが見つかりませんでした。")
            root.destroy() 
            return

        time.sleep(5)

        fix_button = driver.find_element(By.XPATH, '//span[text()="修正"]')
        fix_button.click()
        time.sleep(2)
        input_field = driver.find_element(By.XPATH, '//input[@name="time_clocks[0].datetime"]')
        time.sleep(2)
    
    # フィールドをクリアしてから値を入力
        input_field.send_keys(Keys.CONTROL + "a")  # 全選択 
        input_field.send_keys(Keys.DELETE)  # 選択部分を削除
        input_field.send_keys("900") 
        time.sleep(2)
        fix_button = driver.find_element(By.XPATH, '//span[text()="保存"]')
        fix_button.click()
        time.sleep(5)

        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        


            # JSONファイルに保存
        data = current_time
        startwork_click_time = "startwork_click_time.json"

            # ファイルに書き込み
        with open(startwork_click_time, 'w', encoding='utf-8') as file:
         json.dump(data, file, ensure_ascii=False, indent=4)


        times_label.config(text=times_label.cget("text") + f"出勤時刻: {shukkin}\n")
        print(f"修憩開始ボタンのクリック時刻を {startwork_click_time} に保存しました。")
        window.after(500, on_Cliqwork_start)
        new_window.destroy()
        

    def select_10am():
        print("10時出勤が選択されました")
        shukkin = datetime.now().strftime("%Y-%m-%d %H:%M")
        
    
        options = Options()
        options.add_argument("--headless")  # ヘッドレスモードでバックグラウンド実行
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
        try:
            elements = driver.find_elements(By.XPATH, '//span[text()="ログイン"]')
            if elements:
                root = tk.Tk()  # tkinterのウィンドウ作成
                root.withdraw()  # メインウィンドウを非表示にする
                messagebox.showerror("エラー", "Freeeにてログインが失敗しました。")
                root.destroy()  # ダイアログを閉じた後、ウィンドウを破
                return 
            else:
                print("'ログイン'要素は存在しません。正常に処理を続行します。")
        except Exception as e:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("例外エラー", f"エラーが発生しました: {str(e)}")
            root.destroy()
            return  
        
        try:
            button = driver.find_element(By.XPATH, '//span[text()="出勤"]')
            button.click()
        except:
            print("エラー: '出勤'ボタンが見つかりませんでした。処理をスキップします。")
            root = tk.Tk()  # tkinterのウィンドウ作成
            root.withdraw()  # メインウィンドウを非表示にする
            messagebox.showerror("エラー", "Freeeにて出勤ボタンが見つかりませんでした。")
            root.destroy() 
            return 

        time.sleep(5)

        fix_button = driver.find_element(By.XPATH, '//span[text()="修正"]')
        fix_button.click()
        time.sleep(2)
        input_field = driver.find_element(By.XPATH, '//input[@name="time_clocks[0].datetime"]')
        time.sleep(2)
    
    # フィールドをクリアしてから値を入力
        input_field.send_keys(Keys.CONTROL + "a")  # 全選択 
        input_field.send_keys(Keys.DELETE)  # 選択部分を削除
        input_field.send_keys("1000") 
        time.sleep(2)
        fix_button = driver.find_element(By.XPATH, '//span[text()="保存"]')
        fix_button.click()
        time.sleep(5)

        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # JSONファイルに保存
        data = current_time
        startwork_click_time = "startwork_click_time.json"

            # ファイルに書き込み
        with open(startwork_click_time, 'w', encoding='utf-8') as file:
         json.dump(data, file, ensure_ascii=False, indent=4)


        times_label.config(text=times_label.cget("text") + f"出勤時刻: {shukkin}\n")
        print(f"修憩開始ボタンのクリック時刻を {startwork_click_time} に保存しました。")
        window.after(500, on_Cliqwork_start)
        new_window.destroy()        
        

    btn_9am = tk.Button(new_window, text="9時出勤", command=select_9am, width=10, height=1)
    btn_9am.pack(pady=5)

    # 10時出勤ボタン
    btn_10am = tk.Button(new_window, text="10時出勤", command=select_10am, width=10, height=1)
    btn_10am.pack(pady=5)
    # 現在の出勤時刻を取得

def on_end():
    
    global taikin
    taikin = datetime.now().strftime("%Y-%m-%d %H:%M")

    

    options = Options()
    options.add_argument("--headless")  # ヘッドレスモードでバックグラウンド実行
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
    try:
            elements = driver.find_elements(By.XPATH, '//span[text()="ログイン"]')
            if elements:
                root = tk.Tk()  # tkinterのウィンドウ作成
                root.withdraw()  # メインウィンドウを非表示にする
                messagebox.showerror("エラー", "Freeeにてログインが失敗しました。")
                root.destroy()  # ダイアログを閉じた後、ウィンドウを破
                return 
            else:
                print("'ログイン'要素は存在しません。正常に処理を続行します。")
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("例外エラー", f"エラーが発生しました: {str(e)}")
        root.destroy()
        return 
    time.sleep(5)

    try:
        button1 = driver.find_element(By.XPATH, '//span[text()="退勤"]')
        button1.click()
        time.sleep(2)  # 少し待機してから確認
        try:
            driver.find_element(By.XPATH, '//span[text()="退勤"]')
            print("ボタンがまだ存在しています（クリックされていないか、遷移していない）。")
        except:
            print("ボタンは存在しません（クリックされたか、遷移しました）。")
        print("退勤打刻成功")
    except:
        print("エラー: '退勤'ボタンが見つかりませんでした。処理をスキップします。")
        root = tk.Tk()  # tkinterのウィンドウ作成
        root.withdraw()  # メインウィンドウを非表示にする
        messagebox.showerror("エラー", "Freeeにて出勤処理が失敗しました。")
        root.destroy() 
        return  
    try:
    # ファイルが存在するかチェック
       if os.path.exists(startworkt_click_time):
        os.remove(startworkt_click_time)  # ファイル削除
        print(f"{startworkt_click_time} は削除されました。")
       else:
        print(f"{startworkt_click_time} は存在しません。削除できません。")
    except Exception as e:
    # エラーが発生した場合、処理を無視して何もしない
       pass

    try:
    # ファイルが存在するかチェック
       if os.path.exists(startrest_click_time):
        os.remove(startrest_click_time)  # ファイル削除
        print(f"{startrest_click_time} は削除されました。")
       else:
        print(f"{startrest_click_time} は存在しません。削除できません。")
    except Exception as e:
    # エラーが発生した場合、処理を無視して何もしない
       pass
    window.after(1000, on_Cliqwork_end)

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
    next_button = driver.find_element(By.XPATH, '//*[@id="nextbtn"]/span')
    next_button.click()    
    time.sleep(5)
    password_field = driver.find_element(By.XPATH, '//*[@id="password"]')
    password_field.send_keys(Cliquser_password)
    time.sleep(5)
    login_button = driver.find_element(By.XPATH, '//*[@id="nextbtn"]/span') 

    login_button.click()
    time.sleep(5)
    max_retries = 3  # 最大試行回数
    attempts = 0
    try:
        # 要素を探してクリック
        button = driver.find_element(By.XPATH, '//a[@class="succbutton" and text()="確認する"]')
        button.click()
        print("ボタンをクリックしました。")
    except:
        # 要素が見つからない場合はスキップ
        print("ボタンが見つかりませんでした。処理をスキップします。")
    time.sleep(5)
    try:
        # 要素を探してクリック
        button = driver.find_element(By.XPATH, '//a[@id="continue_button"]')
        button.click()
        print("ボタンをクリックしました。")
    except:
        # 要素が見つからない場合はスキップ
        print("ボタンが見つかりませんでした。処理をスキップします。")
    time.sleep(5)
    while attempts < max_retries:
        try:
        # まず、status_text をクリックする処理
            status_text = driver.find_element(By.XPATH, '//a[text()="同意する"]')
            status_text.click()
            print("ステータスのテキストがクリックされました。")
            break
        except:
            try:
            # 現在のtryをこちらに移動
                continue_button = driver.find_element(By.XPATH, '//a[@class="succbutton" and text()="確認する"]')
                continue_button.click()
                print("次のボタンがクリックされました。")
                time.sleep(5)
                status_text = driver.find_element(By.XPATH, '//a[text()="同意する"]')
                status_text.click()
                break
            except:
                attempts += 1
                print("次のボタンが見つかりませんでした。")
                if attempts < max_retries:
                    time.sleep(2)  # 2秒待機してから再試行
                else:
                    print("最大試行回数に達しました。処理を終了します。")
                    root = tk.Tk()  # tkinterのウィンドウ作成
                    root.withdraw()  # メインウィンドウを非表示にする
                    messagebox.showerror("エラー", "Cliqログインに失敗しました。")
                    root.destroy() 
                    return 

    window.after()
    
def on_Cliqkyuukei_start():
    beakstart = datetime.now().strftime("%Y-%m-%d %H:%M")
    options = Options()
    options.add_argument("--headless")  # ヘッドレスモードでバックグラウンド実行
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    
    driver = webdriver.Chrome(options=options)
    driver.get(Cliq_url)
    time.sleep(3)

    login_field = driver.find_element(By.XPATH, '//*[@id="login_id"]')
    login_field.send_keys(Cliquser_id)
    next_button = driver.find_element(By.XPATH, '//*[@id="nextbtn"]/span')
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
    try:
        # 要素を探してクリック
        button = driver.find_element(By.XPATH, '//a[@id="continue_button"]')
        button.click()
        print("ボタンをクリックしました。")
    except:
        # 要素が見つからない場合はスキップ
        print("ボタンが見つかりませんでした。処理をスキップします。")
    time.sleep(5)
    max_retries = 3  # 最大試行回数
    attempts = 0

    while attempts < max_retries:
        try:
        # まず、status_text をクリックする処理
            status_text = driver.find_element(By.XPATH, '//span[@class="meal-brk mR6"]')
            status_text.click()
            print("ステータスのテキストがクリックされました。")
            break
        except:
            try:
            # 現在のtryをこちらに移動
                continue_button = driver.find_element(By.XPATH, '//a[text()="同意する"]')
                continue_button.click()
                print("次のボタンがクリックされました。")
                time.sleep(5)
                status_text = driver.find_element(By.XPATH, '//span[@class="meal-brk mR6"]')
                status_text.click()
                break
            except:
                attempts += 1
                print("次のボタンが見つかりませんでした。")
                if attempts < max_retries:
                    time.sleep(2)  # 2秒待機してから再試行
                else:
                    print("最大試行回数に達しました。処理を終了します。")
                    root = tk.Tk()  # tkinterのウィンドウ作成
                    root.withdraw()  # メインウィンドウを非表示にする
                    messagebox.showerror("エラー", "Cliq休憩開始に失敗しました。")
                    root.destroy()
                    return


    time.sleep(5)
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # JSONファイルに保存
    data = current_time
    startrest_click_time = "startrest_click_time.json"

            # ファイルに書き込み
    with open(startrest_click_time, 'w', encoding='utf-8') as file:
         json.dump(data, file, ensure_ascii=False, indent=4)
    times_label.config(text=times_label.cget("text") + f"休憩開始: {beakstart}\n")
    print(f"修憩開始ボタンのクリック時刻を {startrest_click_time} に保存しました。")

def on_Cliqwork_start():

    options = Options()
    options.add_argument("--headless")  # ヘッドレスモードでバックグラウンド実行
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    
    driver = webdriver.Chrome(options=options)
    driver.get(Cliq_url)
    time.sleep(3)
    
    login_field = driver.find_element(By.XPATH, '//*[@id="login_id"]')
    login_field.send_keys(Cliquser_id)
    next_button = driver.find_element(By.XPATH, '//*[@id="nextbtn"]/span')
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
    try:
        # 要素を探してクリック
        button = driver.find_element(By.XPATH, '//a[@id="continue_button"]')
        button.click()
        print("ボタンをクリックしました。")
    except:
        # 要素が見つからない場合はスキップ
        print("ボタンが見つかりませんでした。処理をスキップします。")
    time.sleep(5)
    max_retries = 3  # 最大試行回数
    attempts = 0

    while attempts < max_retries:
        try:
        # まず、status_text をクリックする処理
            status_text = driver.find_element(By.XPATH, '//span[@class="slider"]')
            status_text.click()
            print("ステータスのテキストがクリックされました。")
            break
        except:
            try:
            # 現在のtryをこちらに移動
                continue_button = driver.find_element(By.XPATH, '//a[text()="同意する"]')
                continue_button.click()
                print("次のボタンがクリックされました。")
                time.sleep(5)
                status_text = driver.find_element(By.XPATH, '//span[@class="slider"]')
                status_text.click()
                break
            except:
                attempts += 1
                print("次のボタンが見つかりませんでした。")
                if attempts < max_retries:
                    time.sleep(2)  # 2秒待機してから再試行
                else:
                    print("最大試行回数に達しました。処理を終了します。")
                    root = tk.Tk()  # tkinterのウィンドウ作成
                    root.withdraw()  # メインウィンドウを非表示にする
                    messagebox.showerror("エラー", "Cliq業務開始に失敗しました。")
                    root.destroy()
                    return 
    time.sleep(5)

def on_Cliqwork_end():

    options = Options()
    options.add_argument("--headless")  # ヘッドレスモードでバックグラウンド実行
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    
    driver = webdriver.Chrome(options=options)
    driver.get(Cliq_url)
    time.sleep(3)
    
    login_field = driver.find_element(By.XPATH, '//*[@id="login_id"]')
    login_field.send_keys(Cliquser_id)
    next_button = driver.find_element(By.XPATH, '//*[@id="nextbtn"]/span')
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
        time.sleep(5)
        
        status_text = driver.find_element(By.XPATH, '//span[@class="slider"]')
        status_text.click()
    except:
        # 要素が見つからない場合はスキップ
        print("ボタンが見つかりませんでした。処理をスキップします。")
    try:
        # 要素を探してクリック
        button = driver.find_element(By.XPATH, '//a[@id="continue_button"]')
        button.click()
        print("ボタンをクリックしました。")
    except:
        # 要素が見つからない場合はスキップ
        print("ボタンが見つかりませんでした。処理をスキップします。")
    time.sleep(5)
    time.sleep(5)
    max_retries = 3  # 最大試行回数
    attempts = 0

    while attempts < max_retries:
        try:
        # まず、status_text をクリックする処理
            status_text = driver.find_element(By.XPATH, '//span[@class="slider"]')
            status_text.click()
            print("ステータスのテキストがクリックされました。")
            break
        except:
            try:
            # 現在のtryをこちらに移動
                continue_button = driver.find_element(By.XPATH, '//a[text()="同意する"]')
                continue_button.click()
                time.sleep(5)
                status_text = driver.find_element(By.XPATH, '//span[@class="slider"]')
                status_text.click()
                print("次のボタンがクリックされました。")
                break
            except:
                attempts += 1
                print("次のボタンが見つかりませんでした。")
                if attempts < max_retries:
                    time.sleep(2)  # 2秒待機してから再試行
                else:
                    print("最大試行回数に達しました。処理を終了します。")
                    root = tk.Tk()  # tkinterのウィンドウ作成
                    root.withdraw()  # メインウィンドウを非表示にする
                    messagebox.showerror("エラー", "Cliq退勤処理に失敗しました。")
                    root.destroy()
                    return   
    time.sleep(3)
    messagebox.showinfo("Freee", "今日も一日お疲れ様でした。")
    time.sleep(1)
    window.quit()

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
    try:
        login_button = driver.find_element(By.XPATH, '//span[text()="ログイン"]')
        login_button.click()
    except:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("エラー", "ログインが見つかりませんでした。")
        root.destroy()
        return  
    time.sleep(5)
    try:
        elements = driver.find_elements(By.XPATH, '//span[text()="ログイン"]')

        if elements:
            root = tk.Tk()  # tkinterのウィンドウ作成
            root.withdraw()  # メインウィンドウを非表示にする
            messagebox.showerror("エラー", "Freeeにてログインが失敗しました。")
            root.destroy()  # ダイアログを閉じた後、ウィンドウを破 
            return  
        else:
            print("'ログイン'要素は存在しません。正常に処理を続行します。")
    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("例外エラー", f"エラーが発生しました: {str(e)}")
        root.destroy()

        
  

    window.after()
    
def show_description_window():
    # 新しいウィンドウ（Toplevel）を作成
    description_window = tk.Toplevel(window)
    description_window.title("説明文")
    description_window.geometry("500x500")


    # コピー機能

    # 説明文を表示する複数のLabelを設定
    description_label2 = tk.Label(description_window, text="ツールについて\n", wraplength=350, justify="left", fg="red")
    description_label2.pack(pady=2)

    description_label1 = tk.Label(
        description_window,
        text=(
            "このツールは、業務の効率化を目指して\n"
            "よく使われる作業を自動化または簡素化するための機能を集めたものです。\n"
            "\n"
            "出勤 : 出勤時にボタンを押すと打刻が開始されます。\n"
            "\n"
            "退勤 : 退勤時にボタンを押すと打刻が終了されます。\n"
            "\n"
            "Cliq休憩開始 : Cliqのステータスをお昼休憩に変更します。\n"
            "\n"
            "Freeeホーム : Freeeのホーム画面が表示されます。\n"
            "\n"
            "Cliqホーム : Cliqのホーム画面が表示されます。\n"
            
            
        ),
        wraplength=350,
        justify="left"
    )
    description_label1.pack(pady=2)

    """description_label3 = tk.Label(description_window, text="Cliq連携", wraplength=350, justify="left", fg="red")
    description_label3.pack(pady=2)

    description_label4 = tk.Label(
        description_window,
        text=(
            "CliqとFreeeのアカウント情報を同一しておくと\n"
            "出勤・休憩開始・退勤の際に連携してステータスを変更します。\n"
        ),
        wraplength=350,
        justify="left"
    )
    description_label4.pack(pady=2)"""

    description_label5 = tk.Label(description_window, text="お問い合わせ\n", wraplength=350, justify="left", fg="red")
    description_label5.pack(pady=2)

    description_label6 = tk.Label(
        description_window,
        text=(
            "ご使用中にお気づきの問題点や改善のご要望がございましたら、\n以下をクリックし、お送りください。\n"
        ),
        wraplength=350,
        justify="left"
    )
    description_label6.pack(pady=2)

    def open_link():
        """リンクをブラウザで開く"""
        url = "https://docs.google.com/forms/d/e/1FAIpQLSdaSTK_DpkQBCV8hezp7JWfhpXSzpjkxlmEf9GRTp4FIDReZw/viewform?vc=0&c=0&w=1&flr=0"
        webbrowser.open(url)

    # ハイパーリンク風のLabelを作成
    link_label = tk.Label(
        description_window,
        text="問い合わせ先ページ表示",
        fg="blue",
        cursor="hand2",
        wraplength=350,
        justify="left"
    )
    link_label.pack(pady=10)

    # Labelクリック時にリンクを開く
    link_label.bind("<Button-1>", lambda e: open_link())



    # ステータス表示用ラベル
    status_label = tk.Label(description_window, text="", fg="red")
    status_label.pack(pady=5)

    










# start work 
try:
    # startwork_click_timedata が辞書の場合
    if isinstance(startwork_click_timedata, dict):
        formatted_data = json.dumps(startwork_click_timedata, ensure_ascii=False, indent=4).replace('"', '')
    else:
        formatted_data = str(startwork_click_timedata).replace('"', '')



    record_date = datetime.strptime(formatted_data, "%Y-%m-%d %H:%M:%S").date()
    current_date = datetime.now()
    current_date_only = current_date.date()
    if record_date < current_date_only:
        breaktime = ""
    else:
        breaktime = text=f"出勤時刻:{formatted_data}"

    
    # ラベル作成
    break_label = tk.Label(window, text=breaktime)
    break_label.pack(pady=10)

except Exception as e:
    # エラーが発生した場合、処理を無視して何もしない
    pass

# break start 
try:
    # startrest_click_timedata が辞書の場合
    if isinstance(startrest_click_timedata, dict):
        formatted_data = json.dumps(startrest_click_timedata, ensure_ascii=False, indent=4).replace('"', '')
    else:
        formatted_data = str(startrest_click_timedata).replace('"', '')

    record_date = datetime.strptime(formatted_data, "%Y-%m-%d %H:%M:%S").date()
    current_date = datetime.now()
    current_date_only = current_date.date()
    if record_date < current_date_only:
        breakstarttime = ""
    else:
        breakstarttime = text=f"休憩開始:{formatted_data}"
        
    # ラベル作成
    break_label = tk.Label(window, text=breakstarttime)
    break_label.pack(pady=10)

except Exception as e:
    # エラーが発生した場合、処理を無視して何もしない
    pass

# break end

try:
    # endrest_click_timedata が辞書の場合
    if isinstance(endrest_click_timedata, dict):
        formatted_data = json.dumps(endrest_click_timedata, ensure_ascii=False, indent=4).replace('"', '')
    else:
        formatted_data = str(endrest_click_timedata).replace('"', '')

    record_date = datetime.strptime(formatted_data, "%Y-%m-%d %H:%M:%S").date()
    current_date = datetime.now()
    current_date_only = current_date.date()
    if record_date < current_date_only:
        breakendtime = ""
    else:
        breakendtime = text=f"休憩終了:{formatted_data}"

    # ラベル作成
    break_label = tk.Label(window, text=breakendtime)
    break_label.pack(pady=10)
except Exception as e:
    # エラーが発生した場合、処理を無視して何もしない
    pass
#  end work

try:
    if isinstance(startwork_click_timedata, dict):
        formatted_data = json.dumps(startwork_click_timedata, ensure_ascii=False, indent=4).replace('"', '')
    else:
        formatted_data = str(startwork_click_timedata).replace('"', '')



    record_date = datetime.strptime(formatted_data, "%Y-%m-%d %H:%M:%S").date()
    current_date = datetime.now()
    current_date_only = current_date.date()
    if record_date < current_date_only:
        if os.path.exists(startworkt_click_time):  # Cliqcredentials_fileが存在する場合
        # メッセージボックスで警告を表示
            messagebox.showwarning(
            "退勤処理未完了",  # メッセージボックスのタイトル
            "前回の退勤処理が行われていません。\nFreeeホームボタンクリックで前日の勤怠を完了させてください。"  # メッセージ
        )
            if os.path.exists(startworkt_click_time):
               os.remove(startworkt_click_time)  # ファイル削除
               print(f"{startworkt_click_time} は削除されました。")
            else:
               print(f"{startworkt_click_time} は存在しません。削除できません。")
        else:
            print("退勤処理済み")
    else:
        breaktime = text=f"出勤時刻:{formatted_data}"
    # ここで必要な処理を実行
    
except Exception as e:
    # エラーが発生した場合、処理を無視して何もしない
    pass
# Label to display times
times_label = tk.Label(window, text="", justify="left")
times_label.pack(pady=10)

# Create buttons
ok_button = tk.Button(window, text="出勤", command=on_ok, width=5, height=1)
ok_button.pack(side="left", padx=10, pady=10)

start_button = tk.Button(window, text="Cliq休憩開始", command=on_Cliqkyuukei_start, width=10, height=1)
start_button.pack(side="left", padx=10, pady=10)

end_button = tk.Button(window, text="退勤", command=on_end, width=5, height=1)
end_button.pack(side="left", padx=5, pady=5)

home_button = tk.Button(window, text="Freeeホーム", command=on_home, width=8, height=1)
home_button.pack(side="left", padx=15, pady=15)

kyuukei_end_button = tk.Button(window, text="Cliqホーム", command=on_Cliqk_home, width=11, height=1)
kyuukei_end_button.pack(side="left", padx=15, pady=15)






#kyuukei_end_button = tk.Button(window, text="休憩終了", command=on_kyuukei_end, width=5, height=1)
#kyuukei_end_button.pack(side="right", padx=10, pady=10)

description_button = tk.Button(window, text="説明を見る", command=show_description_window)
description_button.place(x=350, y=10, anchor="ne")


# Display the window
window.mainloop()
