from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import soundfile
import speech_recognition as sr
import os
import random
import urllib
import pandas as pd
import numpy as np
import certifi
import math
from tkinter import *
from tkinter import messagebox
from datetime import datetime, timedelta
import ssl 
ssl._create_default_https_context = ssl._create_unverified_context

class Window(Frame):
    
    def __init__(self, master=None):
        Frame.__init__(self, master)        
        self.master = master
        self.pack(fill=BOTH, expand=1)

        label_frame3 = LabelFrame(self, text='ユーザー情報', height=110, width=200)
        label_frame3.place(x=15,y=10)
        self.username = Entry(self, bd=1)
        self.username.place(x=90,y=30,width=110)
        text6 = Label(self, text="ユーザーID")
        text6.place(x=20,y=30)
        self.pwd = Entry(self, bd=1)
        self.pwd.place(x=90,y=60,width=110)
        text7 = Label(self, text="パスワード")
        text7.place(x=20,y=60)
        self.sheet_url = Entry(self, bd=1)
        self.sheet_url.place(x=90,y=90,width=110)
        text8 = Label(self, text="シートURL")
        text8.place(x=20,y=90)
        
        self.var = IntVar()
        self.var.set(1)
        label_frame0 = LabelFrame(self, text="モード", height=50, width=200)
        label_frame0.place(x=15,y=130)
        r_button1 = Radiobutton(self, text="自動", variable=self.var, value=1)
        r_button1.place(x=20,y=147)
        r_button2 = Radiobutton(self, text="手動", variable=self.var, value=2)
        r_button2.place(x=120,y=147)
        
        
        label_frame1 = LabelFrame(self, text='出品時間の間隔', height=50, width=200)
        label_frame1.place(x=15,y=185)
        self.E1 = Entry(self, bd=1)
        self.E1.insert(0, 20)
        self.E1.place(x=20,y=205,width=50)
        text1 = Label(self, text="分から")
        text1.place(x=70,y=205)
        self.E2 = Entry(self, bd=1)
        self.E2.insert(0, 35)
        self.E2.place(x=120,y=205,width=50)
        text2 = Label(self, text="分まで")
        text2.place(x=170,y=205)
        
        label_frame2 = LabelFrame(self, text='ツール稼働時間', height=80, width=200)
        label_frame2.place(x=15,y=250)
        self.start_time = Entry(self, bd=1)
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M")
        self.start_time.insert(0, current_time)
        self.start_time.place(x=20,y=270,width=130)
        text4 = Label(self, text="開始時間")
        text4.place(x=150,y=270)
        self.end_time = Entry(self, bd=1)
        prevent_time = (datetime.now() + timedelta(hours=6)).strftime("%Y-%m-%d %H:%M")
        self.end_time.insert(0, prevent_time)
        self.end_time.place(x=20,y=300,width=130)
        text5 = Label(self, text="終了時間")
        text5.place(x=150,y=300)
        
        self.var1 = IntVar()
        self.var1.set(0)
        r_button1 = Radiobutton(self, text="今すぐ開始", variable=self.var1, value=1, command=self.activation)
        r_button1.place(x=20,y=340)
        r_button2 = Radiobutton(self, text="時間を指定", variable=self.var1, value=2, command=self.activation)
        r_button2.place(x=120,y=340)
        
        # create button, link it to clickExitButton()
        start_button = Button(self, text="開始", command=self.clickStartButton)
        start_button.place(x=20, y=375, width=80)
        close_button = Button(self, text="完了", command=self.clickExitButton)
        close_button.place(x=130, y=375, width=80)
    
    def activation(self):
        if self.var1.get() == 1:
            self.start_time.config(state="disabled")
            self.end_time.config(state="disabled")
        else:
            self.start_time.config(state="normal")
            self.end_time.config(state="normal")
        
    def clickStartButton(self):
        end_time = self.end_time.get()
        print('END Time is just like this: ', end_time)
        end_obj = datetime.strptime(end_time, "%Y-%m-%d %H:%M")
        if self.var1.get() == 2:
            start_time = self.start_time.get()
            start_obj = datetime.strptime(start_time, "%Y-%m-%d %H:%M")
            while True:
                current_obj = datetime.now()
                print(f"waiting...      {current_obj}")
                if (current_obj >= start_obj):
                    if start_obj <= end_obj:
                        break
                    else:
                        messagebox.showinfo("アラート", "稼働時間を正しく入力してください！")
                        break
        username = self.username.get()
        password = self.pwd.get()
        
        print(username)
        print(password)
        # Go to site
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        driver = webdriver.Chrome()
        statusbar.config(text="稼働を開始しました。")
        app.update()
        driver.get('https://gameclub.jp')

        # Press the sell button
        while TRUE:
            try:
                sell_button = driver.find_element("xpath", "//i[@class='fas fa-camera']")
                sell_button.click()
                break
            except:
                statusbar.config(text="起動中")
                app.update()
                

        statusbar.config(text="ログイン中")
        app.update()
        # Input email and password
        while TRUE:
            try:
                email_input = driver.find_element(By.NAME, "email")
                email_input.send_keys(username)
                break
            except:
                statusbar.config(text='起動中')
                app.update()
                
        time.sleep(2)
        pwd_input = driver.find_element(By.NAME, "password")
        pwd_input.send_keys(password)
        time.sleep(2)

        # Click the captcha button
        recaptcha = driver.find_element(By.XPATH, "//div[@class = 'g-recaptcha']")
        recaptcha.click()
        time.sleep(10)

        #Judge the status of check
        driver.switch_to.default_content()
        check_status = driver.find_element(By.XPATH, "//div[5]")
        status = check_status.value_of_css_property("visibility")
        print(status)

        if status != "hidden":
            statusbar.config(text="Captcha認証の突破中です…")
            app.update()
            
            # Get audio challenge
            driver.switch_to.default_content()
            frames=driver.find_element(By.XPATH, "//div[5]").find_elements(By.TAG_NAME, "iframe")
            driver.switch_to.frame(frames[0])
            driver.find_element(By.ID, "recaptcha-audio-button").click()

            # Click the play button
            driver.switch_to.default_content()   
            frames= driver.find_elements(By.TAG_NAME, "iframe")
            driver.switch_to.frame(frames[-1])
            time.sleep(2)
            try:
                driver.find_element(By.XPATH, "//div/div//div[3]/div/button").click()
            except:
                pass                
                
            #get the mp3 audio file
            src = driver.find_element(By.ID, "audio-source").get_attribute("src")
            print("[INFO] Audio src: %s"%src)   

            # download the mp3 audio file from the source
            file_path = os.path.join(os.getcwd(), "captcha1.wav")
            print(file_path)
            urllib.request.urlretrieve(src, file_path)

            data,samplerate=soundfile.read('captcha1.wav')
            soundfile.write('rmt.wav',data,samplerate, subtype='PCM_16')
            r=sr.Recognizer()
            while True:
                try:                    
                    with sr.AudioFile("rmt.wav") as source:
                        audio_data=r.record(source)
                        text=r.recognize_google(audio_data)
                        print(text)
                    time.sleep(2)
                    break
                except sr.UnknownValueError:
                    print("Speech recognition could not understand audio")
                    driver.close()
                except sr.RequestError as e:
                    print("Could not request results from Google Speech Recognition service")     
            
            time.sleep(5)
            
            driver.find_element(By.ID, "audio-response").send_keys(text)
            time.sleep(2)
            driver.find_element(By.ID, "recaptcha-verify-button").click()          

        time.sleep(2) 

        driver.switch_to.default_content()
        login_click = driver.find_element(By.XPATH,"//button[@class='btn btn-danger btn-registration']")
        login_click.click()

        statusbar.config(text="ログインが完了しました。")
        app.update()
        time.sleep(2)
        
        index = 0
        flag = True
        while flag:
            
            end_time = self.end_time.get()
            end_obj = datetime.strptime(end_time, "%Y-%m-%d %H:%M")
            current_obj = datetime.now()
            if current_obj >= end_obj:
                flag = False
                print("The tool is closed")
                break
            
            sheetid = self.sheet_url.get()
            substr = sheetid.split("edit")[0]                      
            url = ''.join([substr, 'gviz/tq?tqx=out:csv&sheet=gameclubWindows'])
            print(url)
            content=pd.read_csv(url)
            
            # mode setting
            if self.var.get() == 1:
                check_status = True
            else:
                check_status = content.iloc[index, 0]
            
            print("check status", check_status)
            # exhibit only when the check is activated   
            if not check_status:
                index += 1
                continue
            
            time.sleep(4)           
            
            
            statusbar.config(text="出品中です…")
            app.update()
            upload_Image = content.iloc[index,2]
            print(upload_Image)
            
            try:
                driver.find_element(By.XPATH, "//input[@type='file']").send_keys(upload_Image)
            except:
                pass
            
            while True:
                try:
                    search = driver.find_element(By.ID, "btn-search-title")
                    search.click()
                    break
                except:
                    statusbar.config(text="起動中")
                    app.update()
            
            time.sleep(10)
            
            while True:
                try:
                    search_title = driver.find_element(By.ID, "search-title-input")
                    search_title.send_keys(content.iloc[index,3])
                    break
                except:
                    statusbar.config(text="起動中")
                    app.update()                   
            
            time.sleep(5)
            while True:
                try:
                    item = driver.find_element(By.XPATH, "//div[@class='item']")
                    item.click()
                    break
                except:
                    statusbar.config(text="起動中")
                    app.update()
                    
            time.sleep(3)
            
            while True:
                try:
                    if content.iloc[index,4] == "代行":
                        radio = driver.find_element(By.ID, "account-type-id-40")
                    else:
                        radio = driver.find_element(By.ID, "account-type-id-10")
                    radio.click()
                    break
                except:
                    statusbar.config(text="起動中")
                    app.update()
            
            while True:
                try:
                    name = driver.find_element(By.NAME, "name")
                    name.send_keys(content.iloc[index,5])
                    break
                except:
                    statusbar.config(text="起動中")
                    app.update()
            
            while True:
                try:
                    detail = driver.find_element(By.NAME, "detail")
                    emoji_text = content.iloc[index,6]
                    driver.execute_script("arguments[0].value = arguments[1]", detail, emoji_text)
                    break
                except:
                    statusbar.config(text="起動中")
                    app.update()
            
            if not isinstance(content.iloc[index, 7], (int, float)):
                while True:
                    try:
                        notice = driver.find_element(By.NAME, "notice_information")
                        notice.send_keys(content.iloc[index,7])
                        break
                    except:
                        statusbar.config(text="起動中")
                        app.update()

            price_budget = int(content.iloc[index, 8])
            
            while True:
                try:
                    price = driver.find_element(By.NAME, "price")
                    price.send_keys(price_budget)
                    break
                except:
                    statusbar.config(text="起動中")
                    app.update()
            
            while True:
                try:
                    confirm_button = driver.find_element(By.ID, "btn-confirm")
                    confirm_button.click()
                    break
                except:
                    statusbar.config(text="起動中")
                    app.update()
            
            time.sleep(2)
            while True:
                try:
                    add_button = driver.find_element(By.ID, "btn-add")
                    add_button.click()
                    break
                except:
                    statusbar.config(text="起動中")
                    app.update()
                    
            statusbar.config(text="出品が完了しました。", fg="green")
            app.update()
            
            time.sleep(3)
            
            while True:
                try:
                    close_button = driver.find_element(By.XPATH, "//*[@id='content-wrapper']/div/div[2]/div[8]/div/div[1]/i")
                    close_button.click()
                    break
                except:
                    statusbar.config(text="起動中")
                    app.update()
            
            if self.var1.get() == 2:
                print('When you see me it means that you fail if it is auto mode!!!')
                current_obj = datetime.now()
                if current_obj >= end_obj:
                    flag = False
                    print("The tool is closed")
                    break
            else:
                pass

            time1 = int(self.E1.get()) * 60
            time2 = int(self.E2.get()) * 60
            wait_time = random.uniform(time1, time2)
            print(wait_time)
            
            if self.var1.get() == 2:
                print('When you see me it means that you fail if it is auto mode!!!')
                current_obj = datetime.now()
                wait_delta = timedelta(seconds=wait_time)
                if (current_obj+wait_delta) >= end_obj:
                    flag = False
                    print("The tool is closed")
                    break
            else:
                pass
            
            statusbar.config(text="待機中")
            app.update()
            time.sleep(wait_time)
            
            while True:
                try:
                    return_button = driver.find_element(By.XPATH, "//*[@id='content-wrapper']/header/div/div[2]/div[2]/a[2]")
                    return_button.click()
                    break
                except:
                    statusbar.config(text="起動中")
                    app.update()
                    
            print("iloc", content.iloc[index+1, 8])
            if pd.isnull(content.iloc[index+1, 8]):
                index = 0
            else:
                index = index + 1
            
            print("Bot is running at the bottom, the index number is: ", index)
            
        driver.quit()
        statusbar.config(text='止まれました。')
        app.update()
    def clickExitButton(self):
        exit()

root = Tk()
app = Window(root)
statusbar = Label(app, text="設定してください。", bd=1, relief=SUNKEN, anchor=E)
statusbar.pack(side=BOTTOM, fill=X)
root.wm_title("gameclub.jp")
root.geometry("230x430+1100+400")
root.resizable(False, False)
root.mainloop()