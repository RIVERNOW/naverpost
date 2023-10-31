import time
from selenium.webdriver.chrome.options import Options
import tkinter as tk
from tkinter import filedialog
from selenium import webdriver
import pyautogui
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.alert import Alert
import pyperclip
import os
import threading
import openpyxl



# Get webdriver
co = Options()
# co.add_experimental_option('debuggerAddress', '127.0.0.1:9222')
driver = webdriver.Chrome(options=co)
driver.implicitly_wait(15)

term = int(input('시간 텀을 설정해 주세요(분 단위)'))
term = term * 60


root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

wb = openpyxl.load_workbook(file_path, data_only=True)
sheet = 'Sheet1'
load_ws = wb[sheet]


#어드민 권한으로 열기
# if sys.argv[-1] != 'asadmin':
#     script = os.path.abspath(sys.argv[0])
#     params = ' '.join([script] + sys.argv[1:] + ['asadmin'])
#     shell.ShellExecuteEx(lpVerb='runas', lpFile=sys.executable, lpParameters=params)
#     sys.exit(0)

# 새탭 안 열리게 방지
# driver.execute_script("window.open = function(url, name, features) { window.location.href = url; }")

# for window in driver.window_handles:
#     driver.switch_to.window(window)
#     if driver.current_url == "https://www.sellergram.co/step2":
#         driver.switch_to.window(window)
#         break


#########################################################################################
#아래 세 함수(closePopup, policyAgree, alertChecker)는 팝업 관리를 위함.
#멀티쓰레드로 각기 실행됨.
def closePopup():
    while True:
        try:
            elm = pyautogui.locateOnScreen('Img/5.png', grayscale=True, confidence=.8)
            if elm is not None:
                pyautogui.click(elm)

            elm = pyautogui.locateOnScreen('Img/7.png', grayscale=True, confidence=.8)
            if elm is not None:
                pyautogui.click(elm)
        except:
            # If no alert is present, continue with your code
            pass
        time.sleep(1)

def policyAgree():
    while True:
        try:
            elm = pyautogui.locateOnScreen('Img/6.png', grayscale=True, confidence=.8)
            if elm is not None:
                pyautogui.click(elm)
        except:
            # If no alert is present, continue with your code
            pass
        time.sleep(1)


def alertChecker():
    while True:
        try:
            # Check if an alert is present
            alert = Alert(driver)

            # If an alert is present, close it (accept it)
            alert.accept()
        except:
            # If no alert is present, continue with your code
            pass
        time.sleep(1)



def logout():
    driver.get('https://post.naver.com/navigator.naver')
    time.sleep(2)

    MY_a = driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/ul/li[4]/a')
    MY_a.click()
    time.sleep(1)

    logout_a = driver.find_element(By.XPATH, '//*[@id="footer"]/div/div[1]/a[1]')
    logout_a.click()
    time.sleep(3)



#네이버 포스트 포스팅 함수
def posting():
    #제목
    pyautogui.hotkey("tab", "up")
    pyperclip.copy(POSTING_TITLE)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1)

    # 사진 첨부
    elm = pyautogui.locateOnScreen('Img/2.png', grayscale=True, confidence=.8)
    if elm is not None:
        time.sleep(0.5)
        pyautogui.click(elm)
    time.sleep(1)

    elm = pyautogui.locateOnScreen('Img/3.png', grayscale=True, confidence=.9)
    if elm is not None:
        time.sleep(0.5)
        pyautogui.click(elm)
    time.sleep(3)


    r = f'posting_img/{POSTING_IMG}'
    pyautogui.typewrite(os.path.abspath(r))
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(5)
    pyautogui.click()
    time.sleep(1)

    #사진 설명 작성
    elm = pyautogui.locateOnScreen('Img/4.png', grayscale=True, confidence=.8)
    if elm is not None:
        time.sleep(0.5)
        pyautogui.click(elm)
        time.sleep(1)
        pyperclip.copy(POSTING_DES)
        pyautogui.hotkey("ctrl", "v")
    time.sleep(1)

    #내용 작성
    pyautogui.press('enter')
    pyperclip.copy(POSTING_CONTENTS)
    pyautogui.hotkey("ctrl", "v")

    submit_a = driver.find_element(By.XPATH, '//*[@id="se_top_publish_btn"]')
    submit_a.click()
    time.sleep(2)

    submit_btn = driver.find_element(By.XPATH, '//*[@id="se_top_publish_setting_layer"]/div/div[2]/div[3]/button[1]')
    submit_btn.click()
    time.sleep(7)

def login():
    ip = driver.find_element(By.XPATH, '//*[@id="switch_blind"]')
    if ip.text == 'on':
        ip_label = driver.find_element(By.XPATH, '//*[@id="login_keep_wrap"]/div[2]/span/label')
        ip_label.click()
    time.sleep(1)

    pyperclip.copy(ID)
    id_input = driver.find_element(By.XPATH, '//*[@id="id"]')
    id_input.send_keys(Keys.CONTROL, 'v')
    time.sleep(1)

    pyperclip.copy(PWD)
    pwd_input = driver.find_element(By.XPATH, '//*[@id="pw"]')
    pwd_input.send_keys(Keys.CONTROL, 'v')
    time.sleep(1)

    submit_btn = driver.find_element(By.XPATH, '//*[@id="log.login"]')
    submit_btn.click()
    time.sleep(3)

#네이버포스트 프로필 설정 함수
def profileSetting():

    profile_a = driver.find_element(By.XPATH, '//*[@id="bgImage"]/div[1]/div[1]/div/a')
    profile_a.click()
    time.sleep(1)

    profile_label = driver.find_element(By.XPATH, '//*[@id="change_profile"]/label')
    profile_label.click()
    time.sleep(3)

    r = f'profile_img/{PROFILE_IMG}'
    pyautogui.typewrite(os.path.abspath(r))
    time.sleep(1)
    pyautogui.press('enter')

    name_input = driver.find_element(By.XPATH, '//*[@id="penName"]')
    name_input.clear()
    name_input.send_keys(PROFILE_NICKNAME)
    time.sleep(1)

    title_input = driver.find_element(By.XPATH, '//*[@id="editorTitle"]')
    title_input.clear()
    title_input.send_keys(PROFILE_TITLE)
    time.sleep(1)

    intro_textarea = driver.find_element(By.XPATH, '//*[@id="introduceDesc"]')
    intro_textarea.click()
    time.sleep(1)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.press('backspace')
    time.sleep(1)
    pyperclip.copy(PROFILE_INTRO)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1)

    career_textarea = driver.find_element(By.XPATH, '//*[@id="activityDesc"]')
    career_textarea.click()
    time.sleep(1)
    pyautogui.hotkey("ctrl", "a")
    pyautogui.press('backspace')
    time.sleep(1)
    pyperclip.copy(PROFILE_CAREER)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1)

    save_a = driver.find_element(By.XPATH, '//*[@id="save"]')
    save_a.click()


# 메인쓰레드 전반적인 함수 관리 및 프로세스 관할
def mainThread():

        driver.maximize_window()

        for i in range(1, load_ws.max_row):
            i += 1
            global ID
            ID = load_ws['A' + str(i)].value
            global PWD
            PWD = load_ws['B' + str(i)].value
            global PROFILE_IMG
            PROFILE_IMG = load_ws['C' + str(i)].value
            global PROFILE_NICKNAME
            PROFILE_NICKNAME = load_ws['D' + str(i)].value
            global PROFILE_TITLE
            PROFILE_TITLE = load_ws['E' + str(i)].value
            global PROFILE_INTRO
            PROFILE_INTRO = load_ws['F' + str(i)].value
            global PROFILE_CAREER
            PROFILE_CAREER = load_ws['G' + str(i)].value
            global POSTING_TITLE
            POSTING_TITLE = load_ws['H' + str(i)].value
            global POSTING_IMG
            POSTING_IMG = load_ws['I' + str(i)].value
            global POSTING_DES
            POSTING_DES = load_ws['J' + str(i)].value
            global POSTING_CONTENTS
            POSTING_CONTENTS = load_ws['K' + str(i)].value
            try:
                driver.get('https://post.naver.com/navigator.naver')
                time.sleep(5)

                MY_a = driver.find_element(By.XPATH, '//*[@id="header"]/div[2]/ul/li[4]/a')
                MY_a.click()
                time.sleep(1)

                # 로그인 창으로
                toLogin_a = driver.find_element(By.XPATH, '//*[@id="cont"]/div/div/a[1]')
                toLogin_a.click()
                time.sleep(2)
                # 로그인 함수
                login()

                # 프로필 설정으로
                toSetting_a = driver.find_element(By.XPATH, '//*[@id="cont"]/div[1]/div[1]/div[1]/div[3]/div[1]/a')
                toSetting_a.click()
                time.sleep(2)
                # 프로필 함수
                profileSetting()

                # 포스팅으로
                toPost_a = driver.find_element(By.XPATH, '//*[@id="header"]/div[1]/a[3]')
                toPost_a.click()
                time.sleep(5)
                # 포스팅 함수
                posting()
                # 로그아웃
                logout()
            except Exception as e:
                print(e)
                pass
        time.sleep(int(term))
        mainThread()



t1 = threading.Thread(target=mainThread)
t1.start()
t2 = threading.Thread(target=alertChecker)
t2.start()
t3 = threading.Thread(target=closePopup)
t3.start()
t4 = threading.Thread(target=policyAgree)
t4.start()




