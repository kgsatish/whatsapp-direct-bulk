# importing Pandas library
import pandas as pd
import urllib.parse
import argparse
import decimal
import getpass
 
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
import openpyxl

import time
from datetime import datetime
import os
import threading

def fn1():
    cwd = os.getcwd() 
    print ('[', cwd, ']')
    print (datetime.now())

def fn2():
    tst = ''
    lst = tst.split(',')
    print(tst)
    print(len(lst))
    print(lst[0].strip() is not None)

def fn3():
    num = 2
    for idx in range(0, num):
        print(idx)

    str = 'abc'
    str = 'xyz' + str
    print(str)

def fn4():
    df = pd.read_csv("successgyan-whatsapp.csv", encoding="utf-8")
    print(df['Message'][0])
    wf = open("successgyan-whatsapp.txt", "w", encoding="utf-8")
    for ind in df.index:
        print(df['Name'][ind], df['Contact'][ind], df['Message'][ind], file=wf)
        wf.write("%s, %s, %s", df['Name'][ind], df['Contact'][ind], df['Message'][ind])
    wf.close()

def fn5():
    message = df['Message'][0]
    print(message)
    phone = 916360908831
    qry = {}
    qry['phone'] = str(phone)
    qry['text'] = message
    print(qry)
    url = "https://api.whatsapp.com/send/?{}".format(urllib.parse.urlencode(qry))
    print(url)

def fn7():
    df = pd.read_csv("successgyan-whatsapp.csv", encoding="utf-8")
    message = df['Message'][0]
    print(message)
    msg_lines = message.splitlines()
    msg_lines[:] = [msg for msg in msg_lines if msg.strip()]
    print(msg_lines)
    for msg in msg_lines:
        if msg.strip():
            print(msg.strip())

def fn8():
    parser = argparse.ArgumentParser(description='Whatsapp Bulk Messaging with optional attachment (document, image, audio, video)')
    parser.add_argument('csv_name', help='Full path of CSV file name with Name, Contact and Message columns. Contact should be <country code><10-digit mobile number>', type=str)
    parser.add_argument('--attach', help='Full path of image/audio/video attachment', type=str, dest='attachment_path')
    parsed_args = parser.parse_args()
    args = vars(parsed_args)
    print (parsed_args)
    print (args)

def fn9():
    width = 15
    precision = 3
    value = decimal.Decimal("12.34567")
    print (f"result: {value:{width}.{precision}}")  # nested fields
    print (f'hello {width} \n')
    print (r'hello {width} \n')
    print (fr'hello {width} \n')
    username = getpass.getuser()
    print(fr'--user-data-dir=C:\Users\{username}\AppData\Local\Google\Chrome\User Data\WhatsApp')

def fn10():
    sheet_gid = ''
    sheet_id = ''
    sheet_name = ''
    url = f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx&gid={sheet_gid}'
    excel_data = pd.read_excel(url, sheet_name=sheet_name, engine='openpyxl')

def fn11():
    #excel_name='rajani.xlsx'
    #excel_data = pd.read_excel(excel_name, usecols=['SL NO', 'NAME', 'CONTACT DETAILS'], sheet_name='12th June', engine='openpyxl')
    #print (excel_data['SL NO'].tolist())
    #print (excel_data.head())

    path = 'rajani1.xlsx'
    df = pd.DataFrame([['SL NO', 'NAME', 'CONTACT DETAILS']])
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='test1', index=False, header=False, startrow=0)
    
    val1 = 11
    val2 = 'hello'
    val3 = 31
    cnt = 1
    df = pd.DataFrame([[val1, val2, val3]])
    with pd.ExcelWriter(path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name='test1', index=False, header=False, startrow=cnt)
    
    cnt += 1
    df = pd.DataFrame([[12, 22, 32]])
    with pd.ExcelWriter(path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name='test1', index=False, header=False, startrow=cnt)

    cnt += 1
    df = pd.DataFrame([[31, 32, 33]])
    with pd.ExcelWriter(path, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
        df.to_excel(writer, sheet_name='test1', index=False, header=False, startrow=cnt)

def fn111():
    path = 'rajani1.xlsx'
    with pd.ExcelWriter(path) as writer:
        writer.book = openpyxl.load_workbook(path)
        df.to_excel(writer, sheet_name='test2')
              
def fn12():
    options = webdriver.ChromeOptions()
    username = getpass.getuser()
    options.add_argument(fr'--user-data-dir=C:\Users\{username}\AppData\Local\Google\Chrome\User Data\WhatsApp')
    options.add_argument('--disable-web-security')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-dev-shm-usage')
    options.add_experimental_option("excludeSwitches", ["enable-logging"]) 
    driver = webdriver.Chrome(service=Service(fr'C:\Users\{username}\AppData\Local\Programs\Python\Python38\chromedriver.exe'), options=options)
    driver_wait = WebDriverWait(driver, 60)

    url="https://web.whatsapp.com/send?phone=911234567890&text=gundanna"
    driver.get(url)
    time.sleep(5)
    try:
        elements = driver.find_elements(By.XPATH, '//div[starts-with(@class, "_")]')
        for element in elements:
            if element.text == 'Phone number shared via url is invalid.':
                ok_button = driver.find_element(By.XPATH, '//div[@role = "button" and starts-with(@class, "_20C5O")]')
                print (ok_button)
                action = ActionChains(driver)
                action.move_to_element(ok_button).click().perform()
    except StaleElementReferenceException:
        pass

def print_cube(num):
    time.sleep(3)
    print("Cube: ", num * num * num, datetime.now())
    return 0

def print_square(num):
    time.sleep(3)
    print("Square: ", num * num, datetime.now())
    return 0

def fn13():
    # Illustrate the concept of threading
    # creating thread
    t1 = threading.Thread(target=print_square, args=(10,))
    t2 = threading.Thread(target=print_cube, args=(10,))

    # starting thread 1
    t1.start()
    # starting thread 2
    t2.start()

    # wait until thread 1 is completely executed
    t1.join()
    # wait until thread 2 is completely executed
    t2.join()

    # both threads completely executed
    print("Done!")

class TestMulti:
    """ Use with WebDriverWait to combine expected_conditions
        in an OR.
    """
    def __init__(self, *args):
        self.ecs = args
    def __call__(self, num):
        for fn in self.ecs:
            try:
                res = fn(num)
                if res:
                    return True
                    # Or return res if you need the element found
            except:
                pass

def fn14():
  return lambda a : a * 0 or a * 4

def fn15():
    lambdafunc = fn14()
    print(lambdafunc(11))

def fn16():
    options = webdriver.ChromeOptions()
    username = getpass.getuser()
    options.add_argument(fr'--user-data-dir=C:\Users\{username}\AppData\Local\Google\Chrome\User Data\WhatsApp')
    options.add_argument('--disable-web-security')
    options.add_argument('--disable-gpu')
    options.add_argument('--disable-dev-shm-usage')
    options.add_experimental_option("excludeSwitches", ["enable-logging"]) 
    driver = webdriver.Chrome(service=Service(fr'C:\Users\{username}\AppData\Local\Programs\Python\Python310\chromedriver.exe'), options=options)
    driver_wait = WebDriverWait(driver, 10)

    url="https://web.whatsapp.com/send?phone=919535246513&text=gundanna"
    driver.get(url)
    try:
        element = driver_wait.until(
            lambda driver:  driver.find_element(By.XPATH, 
            '//div[text()="Phone number shared via url is invalid."] | //span[@class="selectable-text copyable-text" and @data-lexical-text="true" and contains(text(),"gundanna")]'))
        
        time.sleep(0.5)

        print(element.tag_name)
        if element.tag_name == 'span':
            print('valid phone')
        else:
            print ('invalid phone')
    except Exception as exp:
        print(exp)
        
    driver.quit()
            
#TestMulti(print_square(10), print_cube(10))
fn11()