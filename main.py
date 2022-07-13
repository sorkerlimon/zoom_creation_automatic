# from selenium import webdriver  # Add selenium web driver

# driver = webdriver.Chrome()  # Acces selenium class

# driver.get("https://webmail.daldirect.com/?_task=logout&_token=IiOpQiNHlu53LYspB8pk1HkRDkfiKfyU")  ## website accces
# driver.fullscreen_window() # Full screen mode
# driver.maximize_window() # Maximize mdde
# print(driver.title)  # Running website title
# print(driver.current_url)  # Current Website link will ber given
# print(driver.page_source)  # Current website source
# driver.close() # Current browser will be closed


############################## Web element acces ##############

from selenium import webdriver
import time
import xlrd
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions  as EC

#
# driver = webdriver.Chrome()
# driver.implicitly_wait(30)
# driver.get("https://webmail.daldirect.com/?_task=logout&_token=IiOpQiNHlu53LYspB8pk1HkRDkfiKfyU")
#
# username = driver.find_element(By.ID,'rcmloginuser')
# username.send_keys('security@capturedata.com')
#
#
# password = driver.find_element(By.ID,'rcmloginpwd')
# password.send_keys('qZ@n1h69WlGifptv')
#
# button = driver.find_element(By.ID,'rcmloginsubmit')
# button.submit()
#
# search_button = driver.find_element(By.ID,'mailsearchform')
# search_button.send_keys('Zoom account invitation')
# search_button.submit()
#
#
# zoom_account_invitation = driver.find_element(By.ID,'rcmrowMzUxOTM')
# zoom_account_invitation.click()
#
# try:
#     element = WebDriverWait(driver,10).until(
#         EC.presence_of_element_located(By.LINK_TEXT,"Activate Your Zoom Account")
#     )
#     element.click()
#     print("active link click")
#     driver.back()
# except:
#     driver.quit()







# cnt1 = "https://us02web.zoom.us/activate_help?code=8ctU4AaPOB3awxUYGJFW97KxbnnnKkuxvI5bkJRdnE4.AG.tQdoNXbY2WrrB_ILw3yd3lwOJ6kSPxAPskknBgdfi-HsWQ6vJqglTz6Tlb7J3D7ReBc3lR5cGuSfICmBj5dj94G2Eh_rLoV3YRyj6h4NmJdNPKYknuSPHn2MTW15qWaAd46x7oBOTXhE4V9ktbxJ2h4Z6FrWMsqVhXa8eQz7HJJ7IWItPF6VA5fL7WGJN8Fnwr7x9d5HSDPp2oBPzLiNwYTVefAFLkxwUHNdDsb-K68HRs4B5N-ko134VC067iVtCII.RtXMOkSX4SpZxz5VffP2eg.NVfeHLYunM83Yz2W&fr=hostinvite"
# acnt2 = "https://us02web.zoom.us/activate_help?code=8XyVfc7d7A5kumxYHjOLt7pZYjyBpK7-0UbJJb1F8gs.AG.-CByPPWCMeyuDTcjxEqw28C83XiJcSS3XqahBVLaDHeWqlDuSRH84aBu-Dvcdi_M22hdrtDicbVo4dK72ZKspBr9Nn3xgRjnuwveJ_B-Obg5sErTWwpLRSeqkiUlZkoD79leliCETLpUmdeIS0W0CzygJJ7yUgzKKhHZ7VOhXf9fM6JzpsLU8get0batz0SGJr7geI8Fxhx7ekuArPtCEiSFDZlFpkucjIFmnyvSSq_jRv-VJO6smpkUxPYWj-B-7nY.qfkG22szoVPZ2M7Xt2nxDA.EqBePfsnQDpHpuLI&fr=hostinvite"
# acnt3 = "https://us02web.zoom.us/activate_help?code=HETeIupF0j2dkPsurRmIVES3ExHCJazOzgoHW2roD4I.AG.ShS3--eeYcbB5TqQblpW5YpeSc8qpdOxhxeFFy9ZHT4eyP1VrzfG0rFnC5lACH41OX1ZXZNPY2wET3Q9dyguK0Ok2Agv9zbDqQeDDtS8Vlbq5MQF-aLk8JTjYKJMjHjIUvsv-O8wJVST3gfn2kQKt-bS49FME0PWTSgBLoyN1t2w8mR6J7X4zHELtk50YA5FDbhMsT3yChtATzJ1OvHB7jomLEGUpecv1nvaS5IAREKKeIKostyPcwCLOgEa1Nc6MQQ.QEnMtcK8a1RZ2dgPr7Vl1g.2NWLdaKw0m_1Eoe0&fr=hostinvite"

# print("Account are create...")
#
# acnt_actice = driver.get("https://us02web.zoom.us/activate_help?code=omaZQLK0hhzuJEjbzOUtNgEhHij17QAurM7m5WaEr5M.AG.-1lBlmtDDNw6D5D6ybIBCXoYa28ZytMlA4HqvieJNLpR8Qmr7l4nPKeBnoqJMYGwPjlXAyYKxRPS2VZdSrXJCr1cK_-3i5c02vP0hWwfDVnkKJAk6QzYq_5M2H9inLMuJtqJd7EIH4usNK8p7Jbxj7LV-2kMXpUwdOnFGNKeOToCcfz3d9Xwt2hVCDpVTGMn5pUxaWh0sRB_aSpWoLT375IHGL5ZeaisdDY_GCzQAtgBRz0hear2oCl4cmHCbTRuqEk.kbVZnkKRznnUE-s8cG5jJg.R9Cp2qHTxcFZhMoy&fr=hostinvite")
# time.sleep(10)
# driver.find_element(By.XPATH, "//span[contains(.,'Sign Up with a Password')]").click()
#
# search_box = driver.find_element(By.ID,'firstName')
# search_box.send_keys('Maruf')
# time.sleep(3)
#
# search_box = driver.find_element(By.ID,'lastName')
# search_box.send_keys('Hasan')
# time.sleep(3)
#
# search_box = driver.find_element(By.ID,'password')
# search_box.send_keys('limon@123N')
# time.sleep(3)
#
# search_box = driver.find_element(By.ID,'confirm_password')
# search_box.send_keys('limon@123N')
# time.sleep(5)

# submit_button = driver.find_element(By.CSS_SELECTOR,'input[type="button"]')
# submit_button = driver.find_element(By.XPATH,'//button[@class=\'mgb-md mgt-md zm-button--primary zm-button--small zm-button\']')
# submit_button.click()
# time.sleep(10)

# print("first account create done...")





driver = webdriver.Chrome()
driver.implicitly_wait(30)

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
    driver.find_element(By.XPATH,'//button[@class=\'mgb-md mgt-md zm-button--primary zm-button--small zm-button\']').click()
    time.sleep(5)


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


print("Zoom Account Creating .........".center(40))

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

print("Zoom Account Created .........".center(40))
    # print(f"URl : {urlValue}, First_Name : {f_name} , Last_Name : {l_name} Password : {password} Confirm PAssword : {confirm_password} ")
