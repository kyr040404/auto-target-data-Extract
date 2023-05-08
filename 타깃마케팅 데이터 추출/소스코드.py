#!/usr/bin/env python
# coding: utf-8

# In[41]:


#!pip install selenium
#get_ipython().system('pip install pyautogui')
#get_ipython().system('pip install pyperclip')

import time
import pyautogui
import win32com.client
import pyperclip

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from win32api import GetSystemMetrics


# In[135]:


print("키워드를 입력하면 국내 [업체이름, 주소, 연락처] 정보를 추출합니다.")
print("ex) 파이썬 \nex) 서귀포시 시장 \nex) 월계동 중국집 \n")
search = input("키워드를 입력하세요 : ")


# In[136]:
#카카오지도 접속
driver = webdriver.Chrome()
driver.get('https://map.kakao.com/')
time.sleep(1)

#모니터 크기 확인 후 웹 위치 조정
x, y = pyautogui.size()
driver.set_window_position(x/2, 0)


# In[131]:


#검색
sel = '#search\.keyword\.query'
ele = driver.find_element(By.CSS_SELECTOR, sel)
ele.send_keys(search + '\n')
time.sleep(1)


# In[132]:


#카카오맵 가이드 닫기
sel_option_close = 'body > div.coach_layer.coach_layer_type1 > div > div > div > span'
ele_option_close = driver.find_element(By.CSS_SELECTOR, sel_option_close)
ele_option_close.click()
time.sleep(1) 


#리스트 생성
title_folder = []
addres_folder = []
number_folder = []

#정보 더보기 셀렉터
try:
    sel_more = '#info\.search\.place\.more'
    ele_more = driver.find_element(By.CSS_SELECTOR, sel_more)
    ele_more.click()
    time.sleep(1)
except: # 정보가 적을 경우 더보기란이 뜨지 않음. -> 오류 제거
    pass


# In[133]:


page_num = 2
count = 0
no_datar = 'no_data'

# 데이터 300개 이하 에러방지 try except문 
try:   
    for j in range(20): #페이지 이동 반복문 최대 300개   20 * 15
        for i in range(1,16): # 제목 출력 반복문
            
            # 이름 저장
            title = driver.find_element(By.XPATH, f'//*[@id="info.search.place.list"]/li[{i}]/div[3]/strong/a[2]').text
            title_folder.append(title) 
            
            
            # 주소 저장
            addres = driver.find_element(By.XPATH, f'//*[@id="info.search.place.list"]/li[{i}]/div[5]/div[2]/p[1]').text 
            addres_folder.append(addres) 
        
        
            # 번호가 있는지 확인 후 저장
            number = driver.find_element(By.XPATH, f'//*[@id="info.search.place.list"]/li[{i}]/div[5]/div[4]/span[1]').text 
            if(len(number) > 0):
                number_folder.append(number) # number 저장
            else:
                number_folder.append(" ") # r 저장

            
            # 출력
            count += 1 # 순번
            print(f"[{count}] {title} - {addres} - {number}")
            

        # 다음페이지 이동
        sel_nextpage = f'#info\.search\.page\.no{page_num}'    
        ele_nextpage = driver.find_element(By.CSS_SELECTOR, sel_nextpage)
        ele_nextpage.click()
        time.sleep(1)
        page_num += 1
            
            
        # 5페이지까지 이동 후 다음화살표 클릭
        if (page_num % 6 == 0):
            page_num = 2
            sel_nextarrow = '#info\.search\.page\.next'    
            ele_nextarrow = driver.find_element(By.CSS_SELECTOR, sel_nextarrow)
            ele_nextarrow.click()
            time.sleep(1)
                   
except:
    pass

driver.close()

print(f"\n [데이터 수집 완료] 총 {count}개의 데이터를 찾았습니다. (최대 300개)\n")
print("엑셀로 데이터 옮기는중...")



#액셀 실행 win32.com 사용
time.sleep(2)
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True
wb = excel.Workbooks.Add()
ws = wb.Worksheets("Sheet1")



#상호 주소 전화번호 출력
time.sleep(1)
print_excel = "상호"
pyperclip.copy(print_excel)
pyautogui.hotkey('ctrl', 'v')
time.sleep(0.1)
pyautogui.hotkey('right')
time.sleep(0.1)

print_excel = "주소"
pyperclip.copy(print_excel)
pyautogui.hotkey('ctrl', 'v')
time.sleep(0.1)
pyautogui.hotkey('right')
time.sleep(0.1)

print_excel = "전화번호"
pyperclip.copy(print_excel)
time.sleep(0.1)
pyautogui.hotkey('ctrl', 'v')
time.sleep(0.1)

pyautogui.hotkey('enter')
pyautogui.hotkey('left')
pyautogui.hotkey('left')


#엑셀 데이터 출력
for i in range(0, count-1, 2):
    print_excel = title_folder[i]# 1A
    pyperclip.copy(print_excel)
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.1)
    pyautogui.hotkey('right')
    
    
    print_excel = addres_folder[i]#1B
    pyperclip.copy(print_excel)
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.1)
    pyautogui.hotkey('right')
   

    print_excel = number_folder[i]#1C
    pyperclip.copy(print_excel)
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.1)
    pyautogui.hotkey('enter')

    
    print_excel = number_folder[i+1]#2C
    pyperclip.copy(print_excel)
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.1)
    pyautogui.hotkey('left')
    
    print_excel = addres_folder[i+1]#2B
    pyperclip.copy(print_excel)
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.1)
    pyautogui.hotkey('left')
    
    print_excel = title_folder[i+1]# 1A
    pyperclip.copy(print_excel)
    time.sleep(0.1)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.1)
    pyautogui.hotkey('enter')

#표 만들기
pyautogui.hotkey('up')
pyautogui.hotkey('ctrl', 't')
pyautogui.hotkey('enter')

#칸 늘리기
pyautogui.hotkey('ctrl', 'a')
time.sleep(1)
pyautogui.hotkey('alt','o')
time.sleep(1)
pyautogui.hotkey('c','a')
time.sleep(1)
pyautogui.hotkey('ctrl', 't')

#종료
time.sleep(1)
pyautogui.hotkey('up')

print("엑셀 옮기기 완료")


