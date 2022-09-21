import time
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# pip install selenium
# pip install webdriver-manager
# pip install pandas
# pip install bs4
# pip install openpyxl

# Create Table btn click
def table_add_btn(driver):
    driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div[2]/div/div[1]/div/div/div/div[2]/div[2]/div/button").click()

def table_ID_input(driver, name):
    driver.find_element(By.ID, "tblId").send_keys(name)

def table_Name_input(driver, name):
    driver.find_element(By.ID, "tblKorNm").send_keys(name)

# Check_alert >> Existence : 1, Non-Existence : 0 return
def Check_alert(driver):
    try:
        WebDriverWait(driver, 3).until(EC.alert_is_present())
        alert = driver.switch_to.alert
        # Close alert or...
        # alert.dismiss()

        if alert.text =="성공적으로 등록 되었습니다.":
            alert.accept()
            return 2
        # Accept alert
        alert.accept()
        return 1
    except:
        return 0

# Data Read
file_name_list = glob.glob(r"C:\Users\gonet\Desktop\공공데이터 업무\220920화\[서천]*.xlsx")
df = pd.read_excel(file_name_list[0], sheet_name=None)
df2 = df['테이블명']
df = df['컬럼명']

# Chrom Driver setting & Login
chrome_options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
driver.implicitly_wait(10)
driver.get("http://scdata.new.acego.net/ham/login.do")
driver.find_element(By.ID, "id").click()
driver.find_element(By.ID, "id").send_keys("sc04")
time.sleep(0.2)
driver.find_element(By.ID, "password").click()
driver.find_element(By.ID, "password").send_keys("sc04!@#")
time.sleep(0.2)
driver.find_element(By.XPATH, "//*[@id='login_btn']").click()
time.sleep(0.2)

# Access Management Window
driver.find_element(By.LINK_TEXT, "공공데이터 관리").click()
time.sleep(0.2)
driver.find_element(By.LINK_TEXT, "테이블 관리 마스터").click()
time.sleep(0.2)

# Auto Table DB Create ###############
#### Input Table Setting (User setting)
Table_num = []

#### Macro Run
for i in Table_num:
    Table = df[df["테이블순번"]==i]
    Table.index = range(len(Table))

    table_add_btn(driver)
    driver.find_element(By.ID, "tblId").click()
    driver.find_element(By.ID, "tblId").clear()
    table_ID_input(driver, Table.loc[0, "테이블명"]) ## Table ID input
    time.sleep(0.2)
    driver.find_element(By.ID, "tblKorNm").click()
    driver.find_element(By.ID, "tblKorNm").clear()
    table_Name_input(driver, df2.loc[i-1, "테이블 한글명"]) ## Table Name input
    time.sleep(0.2)
    driver.find_element(By.XPATH, "//*[@id='searchForm']/div/div[1]/div[1]/div[3]/button").click()
    time.sleep(0.2)
    if Check_alert(driver):
        print("테이블 아이디 중복확인 오류")
        quit()
    driver.find_element(By.XPATH, "//*[@id='searchForm']/div/div[2]/div/button").click()
    if Check_alert(driver) != 2:
        print("테이블 등록 실패!")
        quit()

    # Table Column Data Create
    driver.find_element(By.XPATH, "//*[@id='datatable']/tbody/tr[1]/td[7]/div/button[1]").click()
    time.sleep(0.2)

    for j in range(len(Table)):
        driver.find_element(By.XPATH, "/html/body/div[2]/div[1]/div[2]/div/div[1]/div/div/div/div[2]/div[2]/div/button[2]").click()
        time.sleep(0.2)
        driver.find_element(By.ID, "columnEngNm").click()
        if Table.loc[j, "공공데이터표준용어\n(영문 약어)"] == "-":
            driver.find_element(By.ID, "columnEngNm").send_keys(Table.loc[j, "항목명\n(영문)"])
        else:
            if sum(Table.loc[j, "공공데이터표준용어\n(영문 약어)"] == Table.loc[:, "공공데이터표준용어\n(영문 약어)"]) > 1:
                driver.find_element(By.ID, "columnEngNm").send_keys(Table.loc[j, "항목명\n(영문)"])
            else:
                driver.find_element(By.ID, "columnEngNm").send_keys(Table.loc[j, "공공데이터표준용어\n(영문 약어)"])
        time.sleep(0.2)
        driver.find_element(By.ID, "columnNm").click()
        driver.find_element(By.ID, "columnNm").send_keys(Table.loc[j, "항목명\n(국문)"])
        time.sleep(0.2)
        driver.find_element(By.ID, "columnCm").click()
        driver.find_element(By.ID, "columnCm").send_keys(Table.loc[j, "항목설명"])
        time.sleep(0.2)
        driver.find_element(By.ID, "columnTy").click()
        driver.find_element(By.ID, "columnTy").send_keys(Table.loc[j, "데이터타입"])
        time.sleep(0.2)
        driver.find_element(By.ID, "columnLt").click()
        driver.find_element(By.ID, "columnLt").send_keys(int(Table.loc[j, "데이터길이"]))
        time.sleep(0.2)
        driver.find_element(By.XPATH, "//*[@id='columnNullYn']").click()
        if j ==0: #### isNull Decision Condition
            driver.find_element(By.XPATH, "//*[@id='columnNullYn']/option[3]").click()
        else:
            driver.find_element(By.XPATH, "//*[@id='columnNullYn']/option[2]").click()
        driver.find_element(By.XPATH, "//*[@id='columnPk']").click()
        driver.find_element(By.XPATH, "//*[@id='columnPk']/option[3]").click()
        driver.find_element(By.XPATH, "//*[@id='columnIndex']").click()
        driver.find_element(By.XPATH, "//*[@id='columnIndex']/option[3]").click()
        driver.find_element(By.XPATH, "//*[@id='searchForm']/div/div[2]/div/button").click()
        Check_alert(driver)
        Check_alert(driver)
    driver.get("http://scdata.new.acego.net/mng/bdTblMaster/list.do")
    time.sleep(0.2)
