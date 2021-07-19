# ---------------------------------------------- Other
import os
from time import sleep
# ---------------------------------------------- tkinter window pop up
import tkinter as tk
import threading
global text

# ---------------------------------------------- Voice
from gtts import gTTS
from playsound import playsound
# ---------------------------------------------- gSheet
import gspread
from oauth2client.service_account import ServiceAccountCredentials
# from pprint import pprint
# from pyasn1.type.univ import Null
# ---------------------------------------------- fetch api from random web
from selenium import webdriver
# from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
# from bs4 import BeautifulSoup

class App(threading.Thread):

    def __init__(self):
        threading.Thread.__init__(self)
        self.start()

    def callback(self):
        self.root.quit()

    def run(self):
        self.root = tk.Tk()
        self.root.geometry("1000x40")
        self.root.title("กิจกรรมจับสายรหัส IT & INE KMUTNB PRI 2564")
        # self.root.configure(bg='#000')
        self.root.attributes('-topmost',True)
        self.root.protocol("WM_DELETE_WINDOW", self.callback)
        self.label = tk.Label(self.root, text="กิจกรรมจับสายรหัส IT & INE KMUTNB PRI 2564", fg="red4", font="bold_font")
        self.label.pack(pady=6)
        self.root.mainloop()

    def _quit(self):
        try :
            self.root.quit()     # stops mainloop
            self.root.destroy()
        except:
            pass

# -------------------------------------------- function for tk
def chageTextAlert(txt):
    app.label.config(text=txt)

def exitAlert() :
    app._quit()



# -------------------------------------------- get keyPass
def setResultKeyPass(value) :
    sheet.update_cell((n+1), 7, str(value))

# ------------------------------------------------------------------------ Sound area
# print(n)
def checkBranch(value) :
    if (str(value) == 'IT'):  return 'ไอที'
    else: return 'ไอเอ็นอี'

def playSound_info_Dek64(text_dek64) :
    speech_ = gTTS(text_dek64, lang="th", slow=True)
    speech_.save("dek64.mp3")
    playsound('dek64.mp3')
    sleep(2)
    os.remove('dek64.mp3')

def playSound_result(result) :
    speech_ = gTTS(result, lang="th", slow=False)
    speech_.save("result.mp3")
    playsound('result.mp3')
    sleep(2)
    os.remove('result.mp3')

# Variable set default
global n, find, tempDek64, tempP63
n = 0
find = ''
tempDek64 = ''
tempP63 = ''
other = ''
test = False
step_1 = True
step_2 = False

app = App()

# -------------------------------------------- Sheet Api
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',
        "https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("data_api.json", scope)

client = gspread.authorize(creds)
sheet = client.open("ข้อมูลสายรหัสน้อง 64").sheet1
data = sheet.get_all_records()

# variable for data sigle low
# row = sheet.row_values(3)
d_col_no = sheet.col_values(1)
d_col_nickname = sheet.col_values(4)
d_col_branch = sheet.col_values(5)
d_col_sec = sheet.col_values(6)

# -------------------------------------------- Selenium
PATH = "C:/bin/msedgedriver.exe"
driver_dek64 = webdriver.Edge(PATH)
driver_dek64.get('https://wheelofnames.com/whf-cgj')
driver_keyPass = webdriver.Edge(PATH)
driver_keyPass.get('https://wheelofnames.com/mrm-ncz')

# Close ads
sleep(1)
driver_dek64.find_element_by_xpath('/html/body/span/section/div/div[1]/div/div[1]/span/button').click()
driver_keyPass.find_element_by_xpath('/html/body/span/section/div/div[1]/div/div[1]/span/button').click()

# Test run first
driver_dek64.find_element_by_xpath('/html/body/span/section/div/div[2]/div').click()
sleep(2)
driver_keyPass.find_element_by_xpath('/html/body/span/section/div/div[2]/div').click()

while (True) :
    #   Random position dek64
    wait2 = WebDriverWait(driver_dek64, 20000)
    if step_1:
        if  (wait2.until(EC.element_to_be_clickable((By.XPATH, '/html/body/span/div[7]/div[2]/div/section/h1')))):
            find2 = driver_dek64.find_element_by_xpath('/html/body/span/div[7]/div[2]/div/section/h1')
            n = int(find2.text)
            if str(tempDek64) != str(n):
                step_1 = False
                step_2 = True
                tempDek64 = str(n)
                print()
                print(f">> สุ่มน้องได้ลำดับที่ {str(d_col_no[n])} สาขา {d_col_branch[n]} sec {d_col_sec[n]} \" {str(d_col_nickname[n])} \"")
                if test: chageTextAlert(f"STEP: 1 🎉      สุ่มน้องได้ลำดับที่ : {str(d_col_no[n])}      สาขา : {d_col_branch[n]}      sec : {d_col_sec[n]}      \" {str(d_col_nickname[n])} \"")
                playSound_info_Dek64(f"สุ่มน้องได้ ลำดับที่ {str(d_col_no[n])} สาขา {checkBranch(str(d_col_branch[n]))} sec {d_col_sec[n]} น้อง {str(d_col_nickname[n])}")
        else :
            break
    #   Random keyPass
    wait = WebDriverWait(driver_keyPass, 20000)
    if step_2:
        if (wait.until(EC.element_to_be_clickable((By.XPATH, '/html/body/span/div[7]/div[2]/div/section/h1')))) :
            find = driver_keyPass.find_element_by_xpath('/html/body/span/div[7]/div[2]/div/section/h1')
            if str(tempP63) != str(find.text):
                step_1 = True
                step_2 = False
                tempP63 = str(find.text)
                setResultKeyPass(str(find.text))
                print(f">> ลำดับที่ {d_col_no[n]}".ljust(14), f"\" {d_col_nickname[n]} \"".ljust(6), f"สายรหัส [{str(find.text)}]")
                if test: chageTextAlert(f"STEP: 2 🎉      ลำดับที่ : {d_col_no[n]}      สาขา : {d_col_branch[n]}      sec : {d_col_sec[n]}      \" {d_col_nickname[n]} \"      สายรหัส [  {str(find.text)}  ]")
                if (str(find.text).find('*') >= 0):
                    other = "ได้พี่รหัส 2 คน"
                playSound_result(f"ลำดับที่ {d_col_no[n]} น้อง {d_col_nickname[n]} {other} สายรหัสที่ได้ คือ {str(find.text)}".replace('*', ''))
        else :
            break
    test = True
    other = ''
    sleep(1)
