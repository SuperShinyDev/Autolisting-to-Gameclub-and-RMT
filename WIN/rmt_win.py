from numpy import nan
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import soundfile
import speech_recognition as sr
import os
import random
import urllib
import pandas as pd
import certifi
import math
import numpy as np
from datetime import datetime
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
        prevent_time = (datetime.now() + timedelta(hours=48)).strftime("%Y-%m-%d %H:%M")
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
        driver = webdriver.Chrome()
        statusbar.config(text="稼働を開始しました。")
        app.update()
        driver.get('https://rmt.club/')

        # Press the sell button
        sell_button = driver.find_element("xpath", "//img[@class= 'top_slide_pc']")
        sell_button.click()
        
        statusbar.config(text="ログイン中です…")
        app.update()
        # Input email and password
        email_input = driver.find_element(By.NAME, "data[User][mail]")
        email_input.send_keys(username)
        time.sleep(1)
        pwd_input = driver.find_element(By.NAME, "data[User][password]")
        pwd_input.send_keys(password)
        time.sleep(3)

        # Click the captcha button
        recaptcha = driver.find_element(By.XPATH, "//div[@class = 'g-recaptcha']")
        recaptcha.click()
        time.sleep(10)

        #Judge the status of check
        check_status = driver.find_element(By.XPATH, "//div[3]")
        status = check_status.value_of_css_property("visibility")
        print(status)

        if status != "hidden":
            statusbar.config(text="Captcha認証の突破中です…")
            app.update()            

            # Get audio challenge
            driver.switch_to.default_content()
            frames = driver.find_element(By.XPATH, "//div[3]/div[4]/iframe")
            driver.switch_to.frame(frames)
            driver.find_element(By.ID, "recaptcha-audio-button").click()

            # Click the play button
            driver.switch_to.default_content()   
            frames= driver.find_elements(By.TAG_NAME, "iframe")
            driver.switch_to.frame(frames[-1])
            time.sleep(2)
            while TRUE:
                try:
                    driver.find_element(By.XPATH, "//div/div//div[3]/div/button").click()
                    break
                except:
                    continue                
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
                    time.sleep(3)
                    break
                except sr.UnknownValueError:
                    print("Speech recognition could not understand audio")
                    driver.find_element(By.ID, "recaptcha-reload-button").click()
                except sr.RequestError as e:
                    print("Could not request results from Google Speech Recognition service; {0}".format(e))               
            
            time.sleep(2)
                        
            # Click the verify button
            driver.find_element(By.ID, "audio-response").send_keys(text)
            time.sleep(2)
            driver.find_element(By.ID, "recaptcha-verify-button").click()

        time.sleep(2)
        driver.switch_to.default_content()

        login_click = driver.find_element("xpath", "//input[@class='btn_type1 fade']")
        login_click.click()
        statusbar.config(text="ログインが完了しました。")
        app.update()

        index = 0
        flag = True
        while flag:
            if self.var1.get() == 2:
                current_obj = datetime.now()
                if current_obj >= end_obj:
                    flag = False
                    print("The tool is closed")
                    break
            else:
                pass
                
            sheetid = self.sheet_url.get()
            substr = sheetid.split("edit")[0]                      
            url = ''.join([substr, 'gviz/tq?tqx=out:csv&sheet=rmtWindows'])
            print(url)
            content=pd.read_csv(url)

            # mode setting
            if self.var.get() == 1:
                check_status = True
            else:
                check_status = content.iloc[index, 0]
            
            # exhibit only when the check is activated0   
            if not check_status:
                index += 1
                continue
            
            time.sleep(4)
            
            
            
            statusbar.config(text="出品中")
            app.update()
            
            while True:
                try:
                    display_button = driver.find_element("xpath", "//img[@class= 'top_slide_pc']")
                    time.sleep(1)
                    display_button.click()
                    break
                except:
                    statusbar.config(text="待機中")
                    app.update()
                    
            time.sleep(2)
            while True:
                try:
                    sale_button = driver.find_element("xpath", "//label[@for='DealRequest0']")
                    time.sleep(1)
                    sale_button.click()
                    break
                except:
                    statusbar.config(text="待機中")
                    app.update()
            
            try:
                game_name = driver.find_element(By.NAME, "data[Deal][game_title]")
                game_name.send_keys(content.iloc[index,2])
            except:
                statusbar.config(text="ゲム名が選択されませんでした。")
                app.update()
                time.sleep(1)
                statusbar.config(text="出品中です")
                app.update()
                pass
            
            time.sleep(3)
            
            try:
                publication_title = driver.find_element(By.NAME, "data[Deal][deal_title]")
                publication_title.send_keys(content.iloc[index,3])
            except:
                statusbar.config(text="タイトルが選択されませんでした。")
                app.update()
                time.sleep(1)
                statusbar.config(text="出品中です")
                app.update()
                pass
            
            try:
                tag = driver.find_element(By.NAME, "data[Deal][tag]")
                tag.send_keys(content.iloc[index,4])
            except:
                pass

            try:
                detail = driver.find_element(By.NAME, "data[Deal][info]")
                emoji_text = content.iloc[index,5]
                driver.execute_script("arguments[0].value = arguments[1]", detail, emoji_text)
            except:
                statusbar.config(text="テキストが選択されませんでした。")
                app.update()
                time.sleep(1)
                statusbar.config(text="出品中です...")
                app.update()
                pass  

            upload_Image = content.iloc[index,6]
            try:
                driver.find_element(By.XPATH, "//input[@type='file']").send_keys(upload_Image)
            except:
                statusbar.config(text="画像が選択されませんでした。")
                app.update()
                time.sleep(2)
                statusbar.config(text="出品中です")
                app.update()
                pass 
                      
            try:
                offer_user = driver.find_element(By.NAME, "data[Deal][user_name]")
                if type(content.iloc[index, 7]) == np.float64:
                    offer_user.send_keys('')
                else:
                    offer_user.send_keys(content.iloc[index,7])
                  
            except:
                statusbar.config(text="客様のIDが選択されませんでした。")
                app.update()
                time.sleep(2)
                statusbar.config(text="出品中です...")
                app.update()
                pass 
            
            price_budget = int(content.iloc[index,8])
            
            while True:
                try:
                    price = driver.find_element(By.NAME, "data[Deal][deal_price]")
                    price.send_keys(price_budget)
                    break
                except:
                    statusbar.config(text="待機中")
                    app.update()
            
            time.sleep(2)
            
            try:
                if len(driver.find_elements(By.XPATH, '//input[@name="deal_account_id[]" and @value="4"]')):
                    account_type = driver.find_element(By.XPATH, '//input[@name="deal_account_id[]" and @value="4"]')
                else:
                    account_type = driver.find_element(By.XPATH, '//input[@name="deal_account_id[]" and @value="1"]')
                account_type.click()
                time.sleep(2)
            except:
                statusbar.config(text="アカウントが選択されませんでした。")
                app.update()
                time.sleep(2)
                statusbar.config(text="出品中です...")
                app.update()
                pass

            while True:
                try:
                    confirm_button = driver.find_element(By.NAME, "smt_confirm")
                    confirm_button.click()
                    break
                except:
                    statusbar.config(text="待機中")
                    app.update()

            while True:
                try:
                    agree_button = driver.find_element(By.NAME, "data[Deal][agreement]")
                    agree_button.click()
                    break
                except:
                    statusbar.config(text="待機中")
                    app.update()

            while True:
                try:
                    finish_button = driver.find_element(By.NAME, "smt_finish")
                    finish_button.click()
                    statusbar.config(text="出品が完了しました。", fg="green")
                    app.update()
                    break
                except:
                    statusbar.config(text="待機中")
                    app.update()
            
            if self.var1.get() == 2:
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
                current_obj = datetime.now()
                wait_delta = timedelta(seconds=wait_time)
                if (current_obj+wait_delta) >= end_obj:
                    flag = False
                    print("The tool is closed")
                    break
            else:
                pass
            
            statusbar.config(text="待機中です…")
            app.update()
            time.sleep(wait_time)
            
            print("iloc", content.iloc[index+1,8])
            if pd.isnull(content.iloc[index+1, 8]):
                index = 0
            else:
                index = index + 1
            
            print("Our bot is running at the bottom, the index number is: ", index)
        driver.quit()
        statusbar.config(text="止まれました。")
        app.update()
        
    def clickExitButton(self):
        exit()

root = Tk()
app = Window(root)
statusbar = Label(app, text="設定してください。", bd=1, relief=SUNKEN, anchor=E)
statusbar.pack(side=BOTTOM, fill=X)
root.wm_title("rmt.club")
root.geometry("230x430+1100+400")
root.resizable(False, False)
root.mainloop()