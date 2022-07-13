from selenium import webdriver
import time
import xlrd
from selenium.webdriver.common.by import By


def web_based_cal(link, f_name, l_name, password, confirm_password):
    driver.get(link)
    time.sleep(10)
    driver.find_element(By.XPATH, "//span[contains(.,'Sign Up with a Password')]").click()
    time.sleep(5)
    driver.find_element(By.ID, 'firstName').send_keys(f_name)
    time.sleep(5)
    driver.find_element(By.ID, 'lastName').send_keys(l_name)
    time.sleep(5)
    driver.find_element(By.ID, 'password').send_keys(password)
    time.sleep(5)
    driver.find_element(By.ID, 'confirm_password').send_keys(confirm_password)
    time.sleep(5)
    # driver.find_element(By.XPATH,'//button[@class=\'mgb-md mgt-md zm-button--primary zm-button--small zm-button\']').click()
    # time.sleep(5)


login = input(str("Enter your name.")).lower()
password = input(str("Enter your password.")).lower()

if login == "admin":
    if password == "limon@123":

        print("""
         .----------------.  .----------------.  .----------------.  .----------------. 
        | .--------------. || .--------------. || .--------------. || .--------------. |
        | |     _____    | || |     _____    | || | ____    ____ | || |   _____      | |
        | |    |_   _|   | || |    |_   _|   | || ||_   \  /   _|| || |  |_   _|     | |
        | |      | |     | || |      | |     | || |  |   \/   |  | || |    | |       | |
        | |      | |     | || |      | |     | || |  | |\  /| |  | || |    | |   _   | |
        | |     _| |_    | || |     _| |_    | || | _| |_\/_| |_ | || |   _| |__/ |  | |
        | |    |_____|   | || |    |_____|   | || ||_____||_____|| || |  |________|  | |
        | |              | || |              | || |              | || |              | |
        | '--------------' || '--------------' || '--------------' || '--------------' |
         '----------------'  '----------------'  '----------------'  '----------------' 
        """)

        print("[INFO] Zoom Account Creating .........".center(40))

        driver = webdriver.Chrome()
        driver.implicitly_wait(30)

        workbok = xlrd.open_workbook('username.xlsx')
        sheet = workbok.sheet_by_name('login')

        rowCount = sheet.nrows
        colCount = sheet.ncols
        # print(rowCount)
        # print(colCount)

        for curr_row in range(1,rowCount):
            urlValue = sheet.cell_value(curr_row,0)
            f_name = sheet.cell_value(curr_row,1)
            l_name = sheet.cell_value(curr_row,2)
            password = sheet.cell_value(curr_row,5)
            confirm_password = sheet.cell_value(curr_row,5)

            web_based_cal(urlValue,f_name,l_name, password, confirm_password)
            print(f"[INFO] {curr_row} {f_name} {l_name} User Zoom Account created.")

        print("[INFO] Zoom Account Created .........".center(40))
    else:
        print("[INFO] Incorrect Password".center(40))
else:
    print("[INFO] login incorrect.".center(40))

