#!/usr/bin/env python
# coding: utf-8

# In[92]:


from selenium import webdriver
from selenium.webdriver.common.keys import Keys #키보드 사용 위해
from selenium.webdriver.common.by import By #wait, contains 쓰기 위해
from selenium.webdriver.support.ui import WebDriverWait #wait, contains 쓰기 위해
from selenium.webdriver.support import expected_conditions as EC #wait, contains 쓰기 위해
from selenium.common.exceptions import TimeoutException #에러


# In[93]:


import time #break time 주려고


# In[94]:


import pyautogui #패키지 설치 필요 !pip3 install pyautogui 


# In[95]:


import pandas as pd #엑셀 불러 오려고 설치 필요 !pip3 install pandas
                     #!pip3 install xlrd
                     #!pip3 install openpyxl


# In[96]:


import tkinter as tk    #GUI창 띄우기 위해
from tkinter import filedialog, messagebox, ttk


# In[97]:


# initalise the tkinter GUI
root = tk.Tk()

root.geometry("400x200") # set the root dimensions
root.pack_propagate(False) # tells the root to not let the widgets inside it determine its size.
root.resizable(0, 0) # makes the root window fixed in size.


# Frame for open file dialog
file_frame = tk.LabelFrame(root, text="Open File")
file_frame.place(height=180, width=400)

# Buttons
button1 = tk.Button(file_frame, text="Browse A File", command=lambda: File_dialog())
button1.place(rely=0.65, relx=0.50)

button2 = tk.Button(file_frame, text="Load File", command=lambda: Load_excel_data())
button2.place(rely=0.65, relx=0.30)

# The file/file path text
label_file = ttk.Label(file_frame, text="No File Selected")
label_file.place(rely=0, relx=0)


def File_dialog():
    """This Function will open the file explorer and assign the chosen file path to label_file"""
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select A File",
                                          filetype=(("xlsx files", "*.xlsx"),("All Files", "*.*")))
    label_file["text"] = filename
    return None


def Load_excel_data():
    """If the file selected is valid, this will load the file into the Treeview"""
    file_path = label_file["text"]
    try:
        excel_filename = r"{}".format(file_path)
        global df
        if excel_filename[-4:] == ".csv":
            df = pd.read_csv(excel_filename)
        else:
            df = pd.read_excel(excel_filename)

    except ValueError:
        tk.messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        tk.messagebox.showerror("Information", f"No such file as {file_path}")
        return None


root.mainloop()


# In[98]:


driver = webdriver.Chrome()


# In[99]:


#링크 입력
driver.get("https://gsshop.workplace.com")
time.sleep(3)


# In[100]:


# 로그인 수동입력 위해
pyautogui.alert('로그인 후 확인을 눌러주세요.') #로그인 위해


# In[101]:


#엑셀 파일 불러오기
#참고 https://pandas.pydata.org/docs/user_guide/io.html#excel-files
#f=d
print(df.head()) 


# In[102]:


#엑셀 값을 list로 변환
f_list = df.values.tolist() #0부터 시작
print(f_list)


# In[106]:


driver.switch_to.window(driver.window_handles[0])

#키보드 이용하기 / 검색창에 있는 글 지우기
search = driver.find_element_by_xpath('//*[@id="mount_0_0"]/div/div[1]/div/div[2]/div/div[1]/div/div/div[1]/div/div/div[2]/div/div/div[1]/div/div/div/div/div/label/input')
search.send_keys(Keys.CONTROL + 'a');
search.send_keys(Keys.DELETE);


# In[107]:


###New 새창+이름 검색###
for i in range(len(f_list)) :
    try : 
        #새 탭
        driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])

        #검색창 입력
        driver.find_element_by_xpath('//*[@id="mount_0_0"]/div/div[1]/div/div[2]/div/div[1]/div/div/div[1]/div/div/div[2]/div/div/div[1]/div/div/div/div/div/label/input').send_keys(f_list[i])

        #프로필 클릭
        time.sleep(1) #timesleep 꼭 있어야됨
        button = WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, "//*[contains(@id,'1000')]")))
        button.click()
        #print(button.get_attribute('id'))
        
        #5개씩만 띄우기
        if((i % 5) == 4) : #나머지
            pyautogui.alert('계속하려면 클릭하세요.')     
        
        driver.switch_to.window(driver.window_handles[len(driver.window_handles)-1])
        
        if i != len(f_list) :
            driver.execute_script('window.open("https://gsshop.workplace.com");')

    except TimeoutException :
        # 위 에러 무시
        print('Timeout')
        pass


# In[ ]:


##결론 : 

