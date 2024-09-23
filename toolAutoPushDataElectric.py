import openpyxl
from tkinter.ttk import *
from tkinter import *
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
from tkinter import filedialog
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import chromedriver_autoinstaller 
file_path = ''
global str_sheet  
global checkbox_vars
global checkbox_values
def login(tk,mk,link_url,xpath_file):
    global str_sheet
    global checkbox_vars
    driver = webdriver.Chrome()
    if link_url == 'https://pmis.npc.com.vn/PMIS_Web/tmsLogin.jsf':
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        driver = webdriver.Chrome(options=options)
        time.sleep(1)
        driver.get(link_url)
        time.sleep(2)
        user = driver.find_element(By.ID,"formMain:txtUserName")
        user.send_keys(tk)
        time.sleep(0.5)
        password = driver.find_element(By.ID,"formMain:txtPassWord")
        password.send_keys(mk)
        password.send_keys(Keys.ENTER)
        time.sleep(2)
        link_cbm = 'https://pmis.npc.com.vn/PMIS_Web/eam/cbm/cbmWork.jsf'
        driver.get(link_cbm)
        time.sleep(5)
        for i, var in enumerate(checkbox_vars):
            print(str_sheet[i],var.get())
            if str_sheet[i] == 'sheet_MC_22_35' and var.get():
                sheet_MC_22_35(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_MC_110' and var.get():
                sheet_MC_110(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_MBA' and var.get():
                sheet_MBA(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_MBA_TD' and var.get():
                sheet_MBA_TD(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_Tu' and var.get():
                sheet_Tu(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_Tu_22_35' and var.get():
                sheet_Tu_22_35(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_Tu_3_Pha' and var.get():
                sheet_Tu_3_Pha(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_TI' and var.get():
                sheet_TI(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_TI_35_22' and var.get():
                sheet_TI_35_22(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_DCL' and var.get():
                sheet_DCL(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_CSV' and var.get():
                sheet_CSV(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_CSV2' and var.get():
                sheet_CSV2(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_CAP_TONG' and var.get():
                sheet_CAP_TONG(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_AC_QUY' and var.get():
                sheet_AC_QUY(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_DAY_DAN' and var.get():
                sheet_DAY_DAN(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_SU_110' and var.get():
                sheet_SU_110(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_MC_110_Hgis' and var.get():
                sheet_MC_110_Hgis(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_MC35kV_Ngoai_Troi' and var.get():
                sheet_MC35kV_Ngoai_Troi(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_Tu_Bu_35' and var.get():
                sheet_Tu_Bu_35(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_TU_BU_22' and var.get():
                sheet_TU_BU_22(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_CUON_KHANG_22_35' and var.get():
                sheet_CUON_KHANG_22_35(xpath_file,driver,link_cbm)
        messagebox.showinfo("Thông Báo", "Hoàn Thành!!!")
    else:
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        options.add_argument("--disable-notifications")
        options.add_argument('--ignore-certificate-errors')
        driver = webdriver.Chrome(options=options)
        driver.get(link_url)
        time.sleep(2)
        user = driver.find_element(By.ID,"formMain:txtUserName")
        user.send_keys(tk)
        time.sleep(0.5)
        password = driver.find_element(By.ID,"formMain:txtPassWord")
        password.send_keys(mk)
        password.send_keys(Keys.ENTER)
        time.sleep(2)
        try:
            driver.find_element(By.XPATH,'/html/body/div/div[2]/button[3]').click()
        except:
            pass
        time.sleep(2)
        link_cbm = 'https://10.21.50.29/PMIS_Web/eam/cbm/cbmWork.jsf'
        driver.get(link_cbm)
        time.sleep(5)
        for i, var in enumerate(checkbox_vars):
            print(str_sheet[i],var.get())
            if str_sheet[i] == 'sheet_MC_22_35' and var.get():
                sheet_MC_22_35(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_MC_110' and var.get():
                sheet_MC_110(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_MBA' and var.get():
                sheet_MBA(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_MBA_TD' and var.get():
                sheet_MBA_TD(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_Tu' and var.get():
                sheet_Tu(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_Tu_22_35' and var.get():
                sheet_Tu_22_35(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_Tu_3_Pha' and var.get():
                sheet_Tu_3_Pha(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_TI' and var.get():
                sheet_TI(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_TI_35_22' and var.get():
                sheet_TI_35_22(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_DCL' and var.get():
                sheet_DCL(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_CSV' and var.get():
                sheet_CSV(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_CSV2' and var.get():
                sheet_CSV2(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_CAP_TONG' and var.get():
                sheet_CAP_TONG(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_AC_QUY' and var.get():
                sheet_AC_QUY(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_DAY_DAN' and var.get():
                sheet_DAY_DAN(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_SU_110' and var.get():
                sheet_SU_110(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_MC_110_Hgis' and var.get():
                sheet_MC_110_Hgis(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_MC35kV_Ngoai_Troi' and var.get():
                sheet_MC35kV_Ngoai_Troi(xpath_file,driver,link_cbm)
            elif str_sheet[i] == 'sheet_Tu_Bu_35' and var.get():
                sheet_Tu_Bu_35(xpath_file,driver,link_cbm)
        messagebox.showinfo("Thông Báo", "Hoàn Thành!!!")

def check_None(arr):
    #Check mang 1 chieu
    if any(element == 'None' for element in arr):
        return True
    else:
        return False
def sheet_Tu_Bu_35(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet Tụ Bù 35...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['Tụ bù 35']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 43
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 13 or col == 23 or col == 32 or col == 41 or col == 42 or col == 43 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2

        temp=''
        info.append(b)
        b = list()
    toltal_part = 6
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1 
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1 
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            #Handle
            for i in range(2,17):
                if i <= 8:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2])
                elif i>=11:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2-2])
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            for i in range(2,10):        
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[11]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,3):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][8]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[12]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][9]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    

        # Phan 3
        print('Phần 3')
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            for i in range(2,9):        
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[10]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][7]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[11]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][8]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu                                                                      
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break                 
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            for i in range(2,8):        
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][4][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[8]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][6]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[9]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][7]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[10]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][8]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu                                                                      
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break                 
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][5]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
        
            # end combobox
            time.sleep(1)
            #luu                                                                      
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break                 
        # Phan 6
        print('Phần 6')
        if info[row][6][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][6]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break

def sheet_MC_110_Hgis(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet MC 110 Hgis...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['MC 110 Hgis']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 40
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 30 or col == 35 or col == 38 or col == 39 or col == 40 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2

        temp=''
        info.append(b)
        b = list()
    toltal_part = 5
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            #Handle
            for i in range(2,55):
                if i <= 8:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2])
                elif i>=12 and i <= 14:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2-3])
                elif i==16:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2-3-1])
                elif i>=20 and i <= 22:   
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2-3-1-3])
                elif i==24:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2-3-1-3-1])
                elif i>=28 and i<=30:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2-3-1-3-1-3])
                elif i==32:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2-3-1-3-1-3-1])
                elif i>=36 and i<=38:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2-3-1-3-1-3-1-3])
                elif i==40:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2-3-1-3-1-3-1-3-1])
                elif i>=44 and i<=46:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2-3-1-3-1-3-1-3-1-3])
                elif i==48:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2-3-1-3-1-3-1-3-1-3-1])
                elif i>=52 and i<=54:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2-3-1-3-1-3-1-3-1-3-1-3])
                time.sleep(0.2)
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            for i in range(3,6):        
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-3])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[6]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,3):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                print(text, info[row][2][3],text == info[row][2][3])
                if(text == info[row][2][3]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][4]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    

        # Phan 3
        print('Phần 3')
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span[1]/input[1]'))).send_keys(info[row][3][0])
            time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[5]/td[2]/div/span/input[1]'))).send_keys(info[row][3][1])
            time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[6]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,3):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][2]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
           
            # end combobox
            time.sleep(1)
            #luu                                                                      
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break                 
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][4][0]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break

def sheet_MC35kV_Ngoai_Troi(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet MC35kV Ngoài Trời...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['MC35kV ngoài trời']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 30
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 12 or col == 25 or col == 28 or col == 29 or col == 30 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
        # start time -  end time
            if col == col_end:
                col_end += 2

        temp=''
        info.append(b)
        b = list()
    toltal_part = 5
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            for i in range(2,15):
                if i==11:
                    continue  
                if i>11:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-3])
                else:         
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2])
                time.sleep(0.2)
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            for i in range(2,12):   
                if i <= 7:     
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                    time.sleep(0.2)
                elif i>= 9:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-4])
                    time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[13]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][9]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[14]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][10]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[15]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][11]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[16]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[9]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][12]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[9]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    

        # Phan 3
        print('Phần 3')
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            for i in range(2,4):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,3):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][2]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
           
            # end combobox
            time.sleep(1)
            #luu                                                                      
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break                 
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][4][0]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
         
def sheet_MC_22_35(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet MC 22 35...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['MC 22-35']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
                break
        col_end = 43
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 13 or col == 23 or col == 32 or col == 41 or col == 42 or col == 43 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2
        temp=''
        info.append(b)
        b = list()
    toltal_part = 6
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,17):
                if i ==9 or i==10:
                    continue  
                if i>10:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-4])
                else:         
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2])
                time.sleep(0.2)
            time.sleep(1)
            #luu                                                                      
            print('at here')
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,10):        
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[11]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][8]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[12]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][9]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
        print('Phần 3')
        # Phan 3
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,9):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[10]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][7]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[11]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][8]):
                    driver.find_element(By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
            
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,8):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][4][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[8]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][4][6]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[9]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][7]):
                    driver.find_element(By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]').click()
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[10]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][8]):
                    driver.find_element(By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False :
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][5][0]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 6
        print('Phần 6')
        if info[row][6][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][6]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break

def sheet_MC_110(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet MC 110...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['MC 110']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 51
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 21 or col == 33 or col == 46 or col == 49 or col == 50 or col == 51 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2

        temp=''
        info.append(b)
        b = list()
    toltal_part = 6
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/span/input[1]'))).send_keys(info[row][1][0])
            time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/input[1]'))).send_keys(info[row][1][1])
            time.sleep(0.5)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][2]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            time.sleep(0.2)
            #end combobox
            #begin checkbox
            list_checkbox = info[row][1][3].split(',')
            count = 0
            for i in range(1,3):
                if(count == len(list_checkbox)): break
                try:
                    for j in range(2,7,2):
                            text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[5]/td[2]/div/table/tbody/tr[{i}]/td[{j}]/label'))).text
                            if(text == list_checkbox[count].strip()):
                                count +=1
                                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[5]/td[2]/div/table/tbody/tr[{i}]/td[{j-1}]/div/div[2]/span'))).click()
                            time.sleep(0.1)
                except:
                    break
            #end checkbox
            time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[6]/td[2]/div/span/input[1]'))).send_keys(info[row][1][4])
            #begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][5]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            time.sleep(0.2)
            #end combobox
            #begin checkbox
            list_checkbox = info[row][1][6].split(',')
            count = 0
            for i in range(1,5):
                if(count == len(list_checkbox)): break
                try:
                    for j in range(2,7,2):
                        text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[8]/td[2]/div/table/tbody/tr[{i}]/td[{j}]/label'))).text
                        if(text == list_checkbox[count].strip()):
                            count+=1
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[8]/td[2]/div/table/tbody/tr[{i}]/td[{j-1}]/div/div[2]/span'))).click()
                        time.sleep(0.1)
                except:
                    break   
            #end checkbox
            time.sleep(0.2)
            #begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[9]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][7]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            time.sleep(0.2)
            #end combobox
            #begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[10]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[9]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][8]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[9]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            time.sleep(0.2)
            #end combobox
            #begin checkbox
            for k in range(11,16):
                list_checkbox = info[row][1][k-2].split(',')
                count = 0
                for i in range(1,3):
                    if(count == len(list_checkbox)): break
                    try:
                        for j in range(2,7,2):
                            text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{k}]/td[2]/div/table/tbody/tr[{i}]/td[{j}]/label'))).text
                            if(text == list_checkbox[count].strip()):
                                count+=1
                                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{k}]/td[2]/div/table/tbody/tr[{i}]/td[{j-1}]/div/div[2]/span'))).click()
                            time.sleep(0.1)
                    except:
                        break
                time.sleep(0.3)
            time.sleep(1)
            #end checkbox
            #begin combobox
            for k in range(16,21):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{k}]/td[2]/div/span/button'))).click()
                time.sleep(1)
                for i in range(1,5):    
                    text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[{k-6}]/table/tbody/tr[{i}]/td/label'))).text
                    if(text == info[row][1][k-2]):
                        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[{k-6}]/table/tbody/tr[{i}]'))).click()
                        break
                    time.sleep(0.5)
            time.sleep(0.2)
            #end combobox 
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[21]/td[2]/div/textarea'))).send_keys(info[row][1][19])
            time.sleep(0.2)
            #begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[22]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[15]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][20]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[15]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            time.sleep(0.2)
            #end combobox
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        
        #Phan 2
        print('Phần 2')
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,15):
                if i==11:
                    continue  
                if i>11:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-3])
                else:         
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                time.sleep(0.2)
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 3')
        #Phan 3
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,11):        
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[15]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][9]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[16]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][10]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[17]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][11]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[18]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[9]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][12]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[9]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    

        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(3,5):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][4][i-3])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[5]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][2]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
           
            # end combobox
            time.sleep(1)
            #luu                                                                      
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break                 
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][5][0]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 6
        print('Phần 6')
        if info[row][6][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][6]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
           
def sheet_MBA(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet MBA ...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['MBA']
    info = list()
    temp = ''
    a = list()
    b = list()
    
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 143
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            # 13 phần
            if col ==0 or col == 2 or col == 50 or col == 78 or col == 81 or col == 87 or col == 90 or col == 94 or col == 104 or col == 127 or col == 136 or col == 137 or col == 139 or col == 143 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2
        temp=''
        info.append(b)
        b = list()
    toltal_part = 13
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1   
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1  
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            
            # fill
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][1][0])    
            time.sleep(0.2)
            time.sleep(1)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1 
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
               
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            
            for i in range(2,68):
                if i==5 or i==15 or i==20 or i==25 or i==30 or i==31 or i==37 or i==42 or i==47 or i==48 or i==50 or i==53 or i==54 or i==57 or i==58 or i==61 or i==62 or i==65:
                    continue
                if i < 5:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                elif i< 15:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-3])
                elif i< 20:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-4])
                elif i< 25:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-5])
                elif i< 30: #31
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-6])
                elif i< 37:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-8])
                elif i< 42:    
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-9])
                elif i< 47: #48
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-10])
                elif i< 50:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-12])
                elif i< 53: #54
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-13])
                elif i<57: #58
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-15])
                elif i<61: #62
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-17])
                elif i<65:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-19])
                else:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-20])
                time.sleep(0.2)
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)                                       
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
        print('Phần 3')
        # Phan 3
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,17):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
            # begin combobox
            count = 6
            for i in range(18,32):
                if  i == 22 or i == 27:
                    continue
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/button'))).click()
                time.sleep(1)
                if i<22:
                    for k in range(1,5):  
                        text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]/td/label'))).text
                        if(text == info[row][3][i-3]):
                            driver.find_element(By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]').click()
                            count +=1
                            break
                elif i<27:
                    for k in range(1,5):  
                        text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]/td/label'))).text
                        if(text == info[row][3][i-4]):
                            driver.find_element(By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]').click()
                            count +=1
                            break
                else:
                    for k in range(1,5):  
                        text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]/td/label'))).text
                        if(text == info[row][3][i-5]):
                            driver.find_element(By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]').click()
                            count +=1
                            break
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[33]/td[2]/div/span/button'))).click()
            for k in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]/td/label'))).text
                if(text == info[row][3][27]):
                    driver.find_element(By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]').click()
                    count +=1
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
            
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1   
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1  
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][4][0]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/input[1]'))).send_keys(info[row][4][1])
            time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[6]/td[2]/div/span/input[1]'))).send_keys(info[row][4][2])
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break 
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1  
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1   
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
           
            
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][5][0]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span[1]/input[1]'))).send_keys(info[row][5][1])
            time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[5]/td[2]/div/span[1]/input[1]'))).send_keys(info[row][5][2])
            time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/span[1]/input[1]'))).send_keys(info[row][5][3])
            time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[9]/td[2]/div/span[1]/input[1]'))).send_keys(info[row][5][4])
            time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[11]/td[2]/div/span[1]/input[1]'))).send_keys(info[row][5][5])
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break 
        print('Phần 6')
        # Phan 6
        if info[row][6][0].lower() == 'không nhập':
            pos +=1
            current_part +=1    
        elif check_None(info[row][6]) == False:
            pos+=1
            current_part +=1    
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/span[1]/input[1]'))).send_keys(info[row][6][0])
            time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span[1]/input[1]'))).send_keys(info[row][6][1])
            time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[5]/td[2]/div/span[1]/input[1]'))).send_keys(info[row][6][2])
            time.sleep(0.2)
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 7')
        # Phan 7
        if info[row][7][0].lower() == 'không nhập':
            pos +=1
            current_part +=1    
        elif check_None(info[row][7]) == False:
            pos+=1
            current_part +=1    
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
           
           
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/input[1]'))).send_keys(info[row][7][0])
            time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span[1]/input[1]'))).send_keys(info[row][7][1])
            time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[5]/td[2]/div/span[1]/input[1]'))).send_keys(info[row][7][2])
            time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/span[1]/input[1]'))).send_keys(info[row][7][2])
            time.sleep(0.2)
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        # Phan 8
        print('Phần 8')
        if info[row][8][0].lower() == 'không nhập':
            pos +=1
            current_part +=1    
        elif check_None(info[row][8]) == False:
            pos+=1
            current_part +=1    
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            
            for i in range(3,12):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][8][i-3])
                time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[15]/td[2]/div/span/input'))).send_keys(info[row][8][9])
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        
        print('Phần 9')
        #Phan 9
        if info[row][9][0].lower() == 'không nhập':
            pos +=1
            current_part +=1    
        elif check_None(info[row][9]) == False:
            pos +=1
            current_part +=1    
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
                
                # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            
            time.sleep(2)
            for i in range(2,31):
                if i==10 or i==12 or i==15 or i==20 or i==21 or i==22 or i==29:
                    continue
                if i < 10:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][9][i-2])
                elif i< 12:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][9][i-3])
                elif i< 15:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][9][i-4])
                elif i< 20:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][9][i-5])
                elif i< 29:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][9][i-8])
                else:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][9][i-9])
                time.sleep(0.2)
            time.sleep(1)
            #combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[32]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for k in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{k}]/td/label'))).text
                if(text == info[row][9][22]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{k}]').click()
                    break

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
        
        print('Phần 10')
        # Phan 10
        if info[row][10][0].lower() == 'không nhập':
            pos +=1
            current_part +=1    
        elif check_None(info[row][10]) == False:
            pos+=1
            current_part +=1    
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
           
            
            for i in range(2,8):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][10][i-2])
                time.sleep(0.2)
            # begin combobox

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[9]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for k in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{k}]/td/label'))).text
                if(text == info[row][10][6]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{k}]').click()
                    break
            # write
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[11]/td[2]/div/span/input[1]'))).send_keys(info[row][10][7])
            time.sleep(0.2)

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[14]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for k in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{k}]/td/label'))).text
                if(text == info[row][10][8]):
                    driver.find_element(By.XPATH,f'/html/body/div[7]/table/tbody/tr[{k}]').click()
                    break        
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        # Phan 11
        print('Phần 11')
        if info[row][11][0].lower() == 'không nhập':
            pos +=1
            current_part +=1    
        elif check_None(info[row][11]) == False:
            pos+=1
            current_part +=1    
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            time.sleep(2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][11][0]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
                
        # Phan 12 Theo dõi dòng ngắn mạch lớn nhât .........
        print('Phần 12')
        if info[row][12][0].lower() == 'không nhập':
            pos +=1
            current_part +=1    
        elif check_None(info[row][12]) == False:
            pos+=1
            current_part +=1    
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span[1]/input[1]'))).send_keys(info[row][12][0])
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[5]/td[2]/div/span[1]/input[1]'))).send_keys(info[row][12][1])
            time.sleep(0.5)
            
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
                
        # Phan 13 Tuổi thọ
        print('Phần 13')
        if info[row][13][0].lower() == 'không nhập':
            pos +=1
            current_part +=1    
        elif check_None(info[row][13]) == False:
            pos+=1
            current_part +=1    
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
              
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/input[1]'))).send_keys(info[row][13][0])
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/input[1]'))).send_keys(info[row][13][1])
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[6]/td[2]/div/span/input[1]'))).send_keys(info[row][13][2])
            time.sleep(0.5)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][13][3]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break

def sheet_MBA_TD(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet MBA TD...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['MBA TD']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 38
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 26 or col == 36 or col == 37 or col == 38 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2

        temp=''
        info.append(b)
        b = list()
    toltal_part = 4
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        print('Phần 1')
        #Phan 1
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,38):
                if i==13 or i==14 or i==18 or i==19 or i==24 or i==25 or i==29 or i==30 or i==34 or i==35:
                    continue
                if i < 13:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).clear()
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2])
                elif i< 18:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-4])
                elif i< 24:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-6])
                elif i< 29:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-8])
                elif i< 34:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-10])
                else:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-12])
                time.sleep(0.2)
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
        
        print('Phần 2')
        # Phan 2
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,11):
                if i == 5 or i == 7 or i == 9:
                    continue
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                time.sleep(0.2)
            # begin combobox
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[5]/td[2]/div/span/button'))).click()
            time.sleep(0.5)
            for k in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{k}]/td/label'))).text
                if(text == info[row][2][3]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{k}]').click()
                    break
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/span/button'))).click()
            time.sleep(0.5)
            for k in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{k}]/td/label'))).text
                if(text == info[row][2][5]):
                    driver.find_element(By.XPATH,f'/html/body/div[7]/table/tbody/tr[{k}]').click()
                    break

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[9]/td[2]/div/span/button'))).click()
            time.sleep(0.5)
            for k in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[8]/table/tbody/tr[{k}]/td/label'))).text
                if(text == info[row][2][7]):
                    driver.find_element(By.XPATH,f'/html/body/div[8]/table/tbody/tr[{k}]').click()
                    break

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[11]/td[2]/div/span/button'))).click()
            time.sleep(0.5)
            for k in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[9]/table/tbody/tr[{k}]/td/label'))).text
                if(text == info[row][2][9]):
                    driver.find_element(By.XPATH,f'/html/body/div[9]/table/tbody/tr[{k}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
        # Phan 3
        print('Phần 3')
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][3][0]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
                         
def sheet_Tu(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet TU ...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['Tu']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 26
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 2 or col == 7 or col == 16 or col == 23 or col == 25 or col == 26 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2

        temp=''
        info.append(b)
        b = list()
    toltal_part = 6
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][1][0])    
            time.sleep(0.2)
            time.sleep(1)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,6):        
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[6]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][4]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
        print('Phần 3')
        # Phan 3
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,11):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
            
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
            
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,7):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][4][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][4][5]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[8]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][6]):
                    driver.find_element(By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break 
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][5][0])
            time.sleep(0.5)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][5][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 6
        print('Phần 6')
        if info[row][6][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][6]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break

def sheet_Tu_3_Pha(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet Tu 3 Pha ...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['Tu 3 Pha']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 33
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 7 or col == 18 or col == 31 or col == 32 or col == 33 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2
        temp=''
        info.append(b)
        b = list()
    toltal_part = 5
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        print('Phần 1')
        #Phan 1
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,6):        
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[6]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][4]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][5]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[8]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][6]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
        print('Phần 2')
        # Phan 2
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,13):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                time.sleep(0.2)
            
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
            
        # Phan 3
        print('Phần 3')
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,9):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
            # begin combobox
            count = 6
            for i in range(9,15):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/button'))).click()
                time.sleep(1)
                for k in range(1,5):  
                    text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]/td/label'))).text
                    if(text == info[row][3][i-2]):
                        driver.find_element(By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]').click()
                        count +=1
                        break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break 
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][4][0]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break

def sheet_Tu_22_35(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet TU 22-35...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['Tu 22-35']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 33
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 2 or col == 9 or col == 16 or col == 23 or col == 30 or col == 32 or col == 33 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2

        temp=''
        info.append(b)
        b = list()
    toltal_part = 7
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][1][0])    
            time.sleep(0.2)
            time.sleep(1)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,9):        
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                time.sleep(0.2)
           
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
        
        print('Phần 3')
        #Phan 3
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,9):        
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
           
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break      
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,8):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][4][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[8]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][4][6]):                    
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
           
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break 
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,7):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][5][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][5][5]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[8]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][5][6]):
                    driver.find_element(By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break 
        # Phan 6
        print('Phần 6')
        if info[row][6][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][6]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][6][0])
            time.sleep(0.5)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][6][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 7
        print('Phần 7')
        if info[row][7][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][7]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break

def sheet_TI(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet TI ...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['TI']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 21
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 2 or col == 11 or col == 18 or col == 20 or col == 21 or col == col_end: 
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2

        temp=''
        info.append(b)
        b = list()
    toltal_part = 5
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][1][0])    
            time.sleep(0.2)
            time.sleep(1)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,11):        
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                time.sleep(0.2)
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    

        # Phan 3
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,7):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][5]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[8]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][6]):
                    driver.find_element(By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][4][0])
            time.sleep(0.5)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break

def sheet_TI_35_22(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet TI 35-22...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['TI 35-22']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 25
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 2 or col == 15 or col == 22 or col == 24 or col == 25 or col == col_end: 
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2

        temp=''
        info.append(b)
        b = list()
    toltal_part = 5
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][1][0])    
            time.sleep(0.2)
            time.sleep(1)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,18):
                if i == 13 or i == 14 or i == 15:
                    continue
                if i < 13:        
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                else:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-5])
                time.sleep(0.2)

            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    

        # Phan 3
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,7):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][5]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[8]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][6]):
                    driver.find_element(By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][4][0])
            time.sleep(0.5)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break

def sheet_DCL(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet DCL...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['DCL']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 31
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 2 or col == 15 or col == 28 or col == 30 or col == 31 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2

        temp=''
        info.append(b)
        b = list()
    toltal_part = 5
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][1][0])    
            time.sleep(0.2)
            time.sleep(1)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,15):        
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                time.sleep(0.2)
            
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    

        # Phan 3
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,9):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
            # begin combobox
            count = 6
            for i in range(9,15):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/button'))).click()
                time.sleep(1)
                for k in range(1,5):  
                    text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]/td/label'))).text
                    if(text == info[row][3][i-2]):
                        driver.find_element(By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]').click()
                        count +=1
                        break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break             
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][4][0])
            time.sleep(0.5)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
           
def sheet_CSV(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet CSV...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['CSV']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 29
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 3 or col == 9 or col == 21 or col == 27 or col == 28 or col == 29 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2

        temp=''
        info.append(b)
        b = list()
    toltal_part = 6
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][1][0])    
            time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/textarea'))).send_keys(info[row][1][1])    
            time.sleep(1)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][2]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,8):        
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                time.sleep(0.2)
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    

        # Phan 3
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,14):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
            
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,6):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][4][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[6]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][4][4]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][5]):
                    driver.find_element(By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break 
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/span[1]/input[1]'))).send_keys(info[row][5][0])

            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 6
        print('Phần 6')
        if info[row][6][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][6]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break

def sheet_CSV2(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet CSV 2...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['CSV 2']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 24
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 3 or col == 15 or col == 21 or col == 23 or col == 24 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2

        temp=''
        info.append(b)
        b = list()
    toltal_part = 5
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][1][0])    
            time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/textarea'))).send_keys(info[row][1][1])    
            time.sleep(1)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][2]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,14):        
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                time.sleep(0.2)
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    

        # Phan 3
        print('Phần 3')
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,4):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][3][2]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[5]/td[2]/div/span/input[1]'))).send_keys(info[row][3][3])
            time.sleep(0.2)

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[6]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][4]):
                    driver.find_element(By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]').click()
                    break
                    
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][5]):
                    driver.find_element(By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break 
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][4][0])
            time.sleep(0.5)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
           
def sheet_CAP_TONG(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet Cap Tong...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['Cap Tong']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 35
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 2 or col == 13 or col == 22 or col == 30 or col ==34  or col == 35 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2

        temp=''
        info.append(b)
        b = list()
    toltal_part = 6
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][1][0])    
            time.sleep(0.2)
            time.sleep(1)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,15):   
                if i ==6 or i==7 or i==12:
                    continue
                if i < 6:     
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                    time.sleep(0.2)
                elif i < 12:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-4])
                    time.sleep(0.2)
                else:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-5])
                    time.sleep(0.2)
                    
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[16]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][10]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
        print('Phần 3')
        # Phan 3
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,7):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][5]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[8]/td[2]/div/span/input[1]'))).send_keys(info[row][3][6])
            time.sleep(0.2)

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[9]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][7]):
                    driver.find_element(By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]').click()
                    break

            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[10]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][8]):
                    driver.find_element(By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
            
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,6):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][4][i-2])
                time.sleep(0.2)
            # begin combobox  
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[6]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][4]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/span/input[1]'))).send_keys(info[row][4][5])
            time.sleep(0.2)

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[8]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][6]):
                    driver.find_element(By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]').click()
                    break

            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[9]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][7]):
                    driver.find_element(By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break 
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][5][0])
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/textarea'))).send_keys(info[row][5][1])
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/textarea'))).send_keys(info[row][5][2])
            time.sleep(0.5)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[5]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][5][3]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 6
        print('Phần 6')
        if info[row][6][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][6]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break

def sheet_AC_QUY(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet Ac Quy...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['Ac Quy']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 13
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col == 0 or col == 2 or col == 7 or col == 10 or col == 12 or col == 13 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2

        temp=''
        info.append(b)
        b = list()
    toltal_part = 5
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][1][0])    
            time.sleep(0.2)
            time.sleep(1)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,6):        
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[6]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):    
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][4]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
        print('Phần 3')
        # Phan 3
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,4):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][2]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break              
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][4][0])
            time.sleep(0.5)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
           
def sheet_DAY_DAN(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet Day Dan...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['Day Dan']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 22
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 2 or col == 4 or col == 20 or col == 21 or col == 22 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2

        temp=''
        info.append(b)
        b = list()
    toltal_part = 5
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][1][0])    
            time.sleep(0.2)
            time.sleep(1)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/input'))).send_keys(info[row][1][0])    
            time.sleep(0.2)
            time.sleep(1)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
                
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
        print('Phần 3')
        # Phan 3
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,14):                                                             
                if i == 7 :
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/input'))).send_keys(info[row][3][i-2])
                else:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
            # begin combobox
            count = 6
            for i in range(17,21):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/button'))).click()
                time.sleep(1)
                for k in range(1,5):  
                    text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]/td/label'))).text
                    if(text == info[row][3][i-5]):
                        driver.find_element(By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]').click()
                        count +=1
                        break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break             
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][4][0]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
           
def sheet_SU_110(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet Su 110...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['Su 110']
    info = list()
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 15
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 13 or col == 14 or col == 15 or col == col_end:
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2

        temp=''
        info.append(b)
        b = list()
    toltal_part = 3
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,9):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2])
                time.sleep(0.2)
            # begin combobox
            count = 6
            for i in range(9,15):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/button'))).click()
                time.sleep(1)
                for k in range(1,5):  
                    text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]/td/label'))).text
                    if(text == info[row][1][i-2]):
                        driver.find_element(By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]').click()
                        count +=1
                        break
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break 
        # Phan 2
        print('Phần 2')
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5): 
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text 
                if(text == info[row][2][0]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 3
        print('Phần 3')
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break

def sheet_CUON_KHANG_22_35(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet CUỘN KHÁNG 22&35...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['CUỘN KHÁNG 22&35']
    info = list()   
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 38
        for col in range(0, sheet.max_column):
            print(sheet.max_column)
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 2 or col == 22 or col == 35 or col == 37 or col == 38 or col == col_end: 
                a = temp.strip().split('|')
                temp =''
                b.append(a)

            # start time -  end time
            if col == col_end:
                col_end += 2
                print(col,col_end)
        temp=''
        info.append(b)
        b = list()
    print(info)
    toltal_part = 5
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][1][0])    
            time.sleep(0.2)
            time.sleep(1)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][1][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,19):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                time.sleep(0.2)
                
            # begin combobox                                                                            
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[21]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][17]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[22]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][18]):
                    driver.find_element(By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]').click()
                    break
                
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[23]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][19]):
                    driver.find_element(By.XPATH,f'/html/body/div[8]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    

        # Phan 3
        print('Phần 3')
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,9):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][3][i-2])
                time.sleep(0.2)
            # begin combobox
            count = 6
            for i in range(9,15): 
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/button'))).click()
                time.sleep(1)
                for k in range(1,5):  
                    text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]/td/label'))).text
                    if(text == info[row][3][i-2]):
                        driver.find_element(By.XPATH,f'/html/body/div[{count}]/table/tbody/tr[{k}]').click()
                        count +=1
                        break
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)          
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][4][0])
            time.sleep(0.5)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)

            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break

def sheet_TU_BU_22(xpath_file,driver,link_cbm):
    print('Đang thực hiện sheet TU Bu 22...')
    ex = openpyxl.load_workbook(xpath_file)
    sheet = ex['TỤ BÙ 22']
    info = list()   
    temp = ''
    a = list()
    b = list()
    for row in range(3, sheet.max_row+1):
        if sheet[row][0].value is None:
            break
        col_end = 21
        for col in range(0, sheet.max_column):
            temp += str(sheet[row][col].value) + '|'
            if col ==0 or col == 8 or col == 15 or col == 17 or col == 19 or col == 21 or col == col_end: 
                a = temp.strip().split('|')
                temp =''
                b.append(a)
            # start time -  end time
            if col == col_end:
                col_end += 2
        temp=''
        info.append(b)
        b = list()
    toltal_part = 5
    for row in range(0,len(info)):
        current_part = 0
        print('Dòng',row+1)
        driver.get(link_cbm)
        time.sleep(1)
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
        time.sleep(0.5)
        
        driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
        time.sleep(1)
        while True:
            time.sleep(1)
            try:
                driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
            except:
                break
        
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
        time.sleep(2)
        pos = 1
        #Phan 1
        print('Phần 1')
        if info[row][1][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][1]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
                
            #fill a form
            for i in range(2,10):
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][1][i-2])
                time.sleep(0.2)
            
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break  
        print('Phần 2')
        #Phan 2 
        if info[row][2][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][2]) == False:
            pos +=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            
            for i in range(2,7):       
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[{i}]/td[2]/div/span/input[1]'))).send_keys(info[row][2][i-2])
                time.sleep(0.2)

            # begin combobox  
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[7]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):  
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][5]):
                    driver.find_element(By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]').click()
                    break

            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[8]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][2][6]):
                    driver.find_element(By.XPATH,f'/html/body/div[7]/table/tbody/tr[{i}]').click()
                    break
            # end combobox
            
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
                                                                                        
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    

        # Phan 3
        if info[row][3][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][3]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)  
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][3][0])
            time.sleep(0.5)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][3][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break    
        # Phan 4
        print('Phần 4')
        if info[row][4][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][4]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/textarea'))).send_keys(info[row][4][0])
            time.sleep(0.5)
            # begin combobox
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[4]/td[2]/div/span/button'))).click()
            time.sleep(1)
            for i in range(1,5):
                text = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]/td/label'))).text
                if(text == info[row][4][1]):
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/div[6]/table/tbody/tr[{i}]'))).click()
                    break
                time.sleep(0.5)
            # end combobox
            time.sleep(1)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
        # Phan 5
        print('Phần 5')
        if info[row][5][0].lower() == 'không nhập':
            pos +=1
            current_part +=1
        elif check_None(info[row][5]) == False:
            pos+=1
            current_part +=1
            #click plus
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    try:
                        try:  
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[8]/button'))).click()
                        except:
                            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[1]/button'))).click()
                    except:
                        break
            while True:
                time.sleep(1)
                try:
                    iframe = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrDlgCommonCBM"]')))
                    break
                except:
                    pass
            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe)
                    break
                except:
                    pass
            time.sleep(2)
            # start time - end time
            if check_None(info[row][current_part + toltal_part]) == False:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[1]/input'))).send_keys(info[row][current_part + toltal_part][0])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).clear()
                time.sleep(0.1)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[2]/span[2]/input'))).send_keys(info[row][current_part + toltal_part][1])
                time.sleep(0.2)
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[1]/td/span/table/tbody/tr[7]/td[1]/span[1]'))).click()
                
                time.sleep(0.2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[2]/td[2]/div/span/input[1]'))).send_keys(info[row][5][0])
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[2]/td[1]/div/div/table/tbody/tr[2]/td/span/table/tbody/tr[1]/td/div/div/table/tbody/tr[3]/td[2]/div/span/input[1]'))).send_keys(info[row][5][1])
            time.sleep(0.5)
            #luu
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/table/tbody/tr[1]/td/div/div[1]/button[4]'))).click()
            # take iframe
            while True:
                time.sleep(1)
                try:
                    iframe1 = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="ifrdlgChangeStatus"]')))
                    break
                except:
                    pass

            time.sleep(2)
            while True:
                time.sleep(1)
                try:
                    driver.switch_to.frame(iframe1)
                    break
                except:
                    pass
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'//*[@id="formDlg:j_idt29"]/div[3]'))).click()
                    break
                except:
                    pass
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/div[3]/div/ul/li[2]'))).click()
            time.sleep(0.5)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/center/button[2]'))).click()
            
            time.sleep(1)
            driver.get(link_cbm)
            time.sleep(1)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[3]/input'))).send_keys(info[row][0])
            time.sleep(0.5)
            
            driver.find_element(By.XPATH,'/html/body/form/div[4]/div[1]/div/table/tbody/tr/td[4]/button/span[1]').click()
            time.sleep(1)
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,'/html/body/form/div[4]/div[2]/div/div/div[2]/table/tbody/tr/td[2]'))).click()
            time.sleep(2)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH,f'/html/body/form/div[4]/div[3]/div/div[1]/div/div[1]/div/div[2]/table/tbody/tr[{pos}]/td[2]/button'))).click()
            while True:
                time.sleep(1)
                try:
                    driver.find_element(By.XPATH,'/html/body/div[3]/div').click()
                except:
                    break

def graphics():
    win = Tk()
    win.title("By Trương Tùng - Đội 110kV Lai Châu")
    win.geometry('800x500')
    label = Label(win,text='Phần Mềm Tự Động Hóa',font=('arial',15))
    label.place(x=85,y=10)
    global str_sheet
    global checkbox_vars
    global checkbox_values
    checkbox_vars = []
    str_sheet = []
    checkbox_values = []
    
    def setup_initial_values():
        global str_sheet
        global checkbox_vars
        global checkbox_values
        

        # Xóa các checkbox cũ nếu có
        for checkbox in checkbox_values:
            checkbox.destroy()
        checkbox_values.clear()
        checkbox_vars.clear()
        # Tùy thuộc vào giá trị được chọn từ Combobox, cập nhật danh sách str_sheet
        selected_index = address_link_text.current()
        selected_value = address_link_text['value'][selected_index]
        
        str_sheet = ['sheet_MC_22_35','sheet_MC_110','sheet_MBA','sheet_MBA_TD','sheet_Tu','sheet_Tu_22_35','sheet_Tu_3_Pha','sheet_TI','sheet_TI_35_22','sheet_DCL','sheet_CSV','sheet_CSV2','sheet_CAP_TONG','sheet_AC_QUY','sheet_DAY_DAN','sheet_SU_110','sheet_MC_110_Hgis','sheet_MC35kV_Ngoai_Troi','sheet_Tu_Bu_35'
                     ,'sheet_CUON_KHANG_22_35','sheet_TU_BU_22']

        # Tạo lại các checkbox với danh sách str_sheet mới
        parameter = 0
        count = 0
        for i, sheet in enumerate(str_sheet): 
            var = IntVar(value=0)
            checkbox = Checkbutton(win, text=sheet, variable=var)
            checkbox.place(x=400 + parameter, y=50 + count * 30)  # Điều chỉnh vị trí của mỗi checkbox
            count+=1
            if i%12==0 and i != 0: 
                parameter+=200
                count = 0
            checkbox_values.append(checkbox)  # Thêm checkbox vào danh sách checkbox_values
            checkbox_vars.append(var)

        
    def event(object):
        global file_path
        if object == bt_ctn:
            if len(str_sheet) != len(checkbox_vars):
                messagebox.showinfo("Thông Báo", "Đã xảy ra lỗi vui lòng liên hệ coder!!!")
            try:
                user_text.get(),password_text.get()
                login(user_text.get(),password_text.get(),address_link_text.get(),file_path)
            except Exception as e:
                print(e)
                messagebox.showinfo("Thông Báo", "Đã xảy ra lỗi vui lòng kiểm tra!!!")
        elif object == bt_xpath:
            file_path = filedialog.askopenfilename(title="Chọn file")
            label1 = Label(win,text= file_path.split('/')[-1],font=('arial',10))
            label1.place(x=120,y=80)
        elif object == bt_update:
            try:
                chromedriver_autoinstaller.install()
                messagebox.showinfo("Thông Báo", "Đã cập nhật chromedrive thành công")
            except:
                messagebox.showinfo("Thông Báo", "Đã cập nhật chromedrive thất bại")
    
        
    label_path = Label(win,text='Địa chỉ link excel: ',font=('arial',10))
    label_path.place(x=10,y=50)
    bt_xpath = Button(win,text='Chọn File',font=('arial',10),padx=10,command=lambda: event(bt_xpath))
    bt_xpath.place(x=120,y=50)
    label_note = Label(win,text='(Y/C: Đúng dạng quy định sẵn)',font=('arial',10))
    label_note.place(x=210,y=50)


    label_path1 = Label(win,text='Chọn đường dẫn: ',font=('arial',10))
    label_path1.place(x=10,y=110)
    address_link_text = Combobox(win, font=('arial',10),width=28)
    address_link_text['value'] = ('https://pmis.npc.com.vn/PMIS_Web/tmsLogin.jsf','https://10.21.50.29/PMIS_Web/tmsLogin.jsf?faces-redirect=true')
    address_link_text.current(0)
    address_link_text.place(x=120,y=110)
    address_link_text.bind('<<ComboboxSelected>>',setup_initial_values())
    address_link_text.bind('<<ComboboxSelected>>', lambda event: setup_initial_values())

    label_path2 = Label(win,text='Tài khoản: ',font=('arial',10))
    label_path2.place(x=10,y=140)
    user_text = Entry(win, font=('arial',10),width=30,borderwidth=2)
    user_text.place(x=120,y=140)

    label_path3 = Label(win,text='Mật khẩu: ',font=('arial',10))
    label_path3.place(x=10,y=170)
    password_text = Entry(win, font=('arial',10),width=30,borderwidth=2)
    password_text.place(x=120,y=170)

    bt_ctn = Button(win,text='Tiếp Tục',font=('arial',10),padx=25,command=lambda: event(bt_ctn))
    bt_ctn.place(x=150,y=210)
    
    bt_update = Button(win,text='Cập Nhật Chrome',font=('arial',10),padx=35,command=lambda: event(bt_update))
    bt_update.place(x=600,y=10)
    win.mainloop()

def main():
    graphics()

if __name__ == '__main__':
    main()
