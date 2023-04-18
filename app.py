# Program to send bulk customized message with attachment (image etc..) through WhatsApp web application without
# adding contact

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import StaleElementReferenceException
from selenium.webdriver.common.by import By
import pandas as pd
import time
import config
import argparse
import urllib.parse
import os
from os.path import exists
import logging
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
from tktooltip import ToolTip
import pytz
import pymysql.cursors
import errno
from dateutil.relativedelta import relativedelta
import decimal
import debugpy


import os.path
import google.auth
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaFileUpload


# A class that encapsulates Whatsapp Message automation function and attributes
class WhatsappBulkMessage(object):
    exceptionList = {}
    def __init__(self, **kwargs):
        logging.info('Entered init')
        self.xlsFile = kwargs.get('xls_file')
        self.msgFile = kwargs.get('msg_file')
        self.msgInd = kwargs.get('msg_ind')
        self.imgFile = kwargs.get('img_file')
        self.imgPath = None
        if self.imgFile:
            self.imgPath = os.path.join(os.getcwd(), self.imgFile)
        self.excel_data = None
        self.driver = None
        self.driver_wait = None
        self.msg_text = None
        # failCnt keeps tacks of the records that couldn't be processed due to technical exceptions, to be retried
        self.failCnt = 0
        logging.info('Exited init')

    def perform_task(self):
        logging.info('Entered perform task')
        try:
            if not exists(self.xlsFile):
                logging.error("Excel file [%s] does not exist", self.xlsFile)
                raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), self.xlsFile)
            if not exists(self.msgFile):
                logging.error("Message file [%s] does not exist", self.msgFile)
                raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), self.msgFile)
            if self.imgFile and not exists(self.imgPath):
                logging.error("Image file [%s] does not exist", self.imgPath)
                raise FileNotFoundError(errno.ENOENT, os.strerror(errno.ENOENT), self.imgPath)
            logging.info('File exists checks done')
            self.initialize()
            self.read()
            self.process()
        except FileNotFoundError as exp:
            logging.error('Encountered FileNotFoundError %s', exp)
            self.exceptionList[config.FILE_NOT_FOUND] = exp
        except Exception as exp:
            logging.error('Encountered general exception %s', exp)
            self.exceptionList[config.UNCAUGHT_EXCEPTION] = exp
        finally:
            self.close()
            logging.info('Exited perform task')

    def initialize(self):
        # Load the chrome driver
        logging.info('Entered initialize')
        options = webdriver.ChromeOptions()
        options.add_argument(config.CHROME_PROFILE_PATH)
        options.add_argument('--disable-web-security')
        options.add_argument('--disable-gpu')
        options.add_argument('--disable-dev-shm-usage')
        options.add_experimental_option("excludeSwitches", ["enable-logging"])

        self.driver = webdriver.Chrome(service=Service(config.CHROME_DRIVER_PATH), options=options)
        self.driver_wait = WebDriverWait(self.driver, int(config.WAIT_TIME))
        logging.info('Exited initialize')

    def read(self):
        # Read data from excel
        # noinspection PyArgumentList
        logging.info('Entered read')
        self.excel_data = pd.read_excel(self.xlsFile, usecols=['SL NO', 'NAME', 'CONTACT DETAILS'],
                                        engine='openpyxl')
        # Read message from text file
        fp = open(self.msgFile, "r", encoding="utf-8")
        self.msg_text = fp.read()
        fp.close()
        logging.info('Exited read')

    def process(self):
        # Iterate excel rows till to finish
        msgInd = 0
        slNo = ''                
        name = ''
        contactNumber = ''
        
        respFile = 'resp_' + self.xlsFile
        df = pd.DataFrame([['SL NO', 'NAME', 'CONTACT DETAILS']])
        with pd.ExcelWriter(respFile, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, header=False, startrow=self.failCnt)

        logging.info('Started processing...')
        try:
            for idx in self.excel_data.index:
                slNo = str(self.excel_data['SL NO'][idx])
                name = str(self.excel_data['NAME'][idx])
                contactNumber = str(self.excel_data['CONTACT DETAILS'][idx])[0:12]

                logging.info('SL NO: [%s], msgInd: [%d], passed msgInd: [%s], idx: [%d]', slNo, msgInd, self.msgInd, idx)

                if self.msgInd != 'A':
                    logging.debug('After SL NO check - success')
                    if slNo == '1':
                        msgInd += 1
                        logging.debug('msgInd after increase: [%d]', msgInd)

                    # Get the right set of messages to work on
                    if msgInd != int(self.msgInd):
                        logging.info('Ignored: [%s]. [%s]', slNo, name)
                        continue

                logging.info('Considered: [%s]. [%s]', slNo, name)
                # Assign customized addressing only if there is message to be sent 
                if os.stat(self.msgFile).st_size == 0:
                    message = 'Hi ' + name + ', ' + config.IMAGE_TEXT
                else:
                    message = 'Hi ' + name + ',\n\n' + self.msg_text
                
                for mobilePhone in config.VALID_PHONES:
                    if (contactNumber.startswith(mobilePhone['ISD']) and len(contactNumber) != mobilePhone['length']):
                        logging.error('Invalid mobile number: [%s]', contactNumber)
                        self.exceptionList[config.INVALID_PHONE_NUM] = contactNumber
                        break
                
                # No need to go further if an invalid phone
                if self.exceptionList:
                    continue
                
                if len(contactNumber) == 10:
                    contactNumber = '91' + contactNumber
                    
                logging.info('Chrome driver: [%s]', config.CHROME_DRIVER_PATH)
                qry = {'phone': contactNumber, 'text': message}
                url = "https://web.whatsapp.com/send/?{}".format(urllib.parse.urlencode(qry))

                if config.TEST_MODE == 'Y':
                    logging.info('url: %s', url)
                    continue

                logging.debug('before get')
                self.driver.get(url)
                logging.debug('after get')
                
                try:
                    # Either 'invalid phone' alert or  whatsapp message screen depending on if the phone num is valid or not 
                    element = self.driver_wait.until(
                        lambda driver:  driver.find_element(By.XPATH, 
                        '//div[text()="' + config.INVALID_PHONE + '"] | //span[@class="selectable-text copyable-text" and @data-lexical-text="true" and contains(text(),"' + name + '")]'))
                except Exception as exp:
                    logging.error('Exception while processing %s', contactNumber)
                    logging.error(exp)
                    self.exceptionList[config.PROCESS_EXCEPTION] = exp
                    self.failCnt += 1
                    df = pd.DataFrame([[self.failCnt, name, contactNumber]])
                    with pd.ExcelWriter(respFile, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                        df.to_excel(writer, index=False, header=False, startrow=self.failCnt)
                    continue
                    
                # this sleep is required to make things work!    
                time.sleep(2)

                if element.tag_name == 'span':
                    logging.info('Phone number valid: %s', contactNumber)
                else:
                    logging.error('Phone number not in whatsapp: %s', contactNumber)
                    self.exceptionList[config.NOT_IN_WHATSAPP] = contactNumber
                    ok_button = self.driver.find_element(By.XPATH, '//div[@role="button"]')
                    action = ActionChains(self.driver)
                    action.move_to_element(ok_button).click().perform()
                    continue

                if self.imgPath:
                    attachment_button = self.driver_wait.until(
                        lambda driver: driver.find_element(By.XPATH, '//div[@title="Attach"]'))
                    attachment_button.click()
                    image_button = self.driver_wait.until(lambda driver: driver.find_element(
                        By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'))
                    image_button.send_keys(self.imgPath)
                    send_button = self.driver_wait.until(lambda driver: driver.find_element(
                        By.XPATH, '//div[@role="button" and @aria-label="Send"]'))
                    action = ActionChains(self.driver)
                    action.move_to_element(send_button).click().perform()
                    time.sleep(10)
                else:
                    message_box = self.driver_wait.until(lambda driver: driver.find_element(
                        By.XPATH, '//div[@title="Type a message" and @role="textbox"]'))
                    action = ActionChains(self.driver)
                    action.move_to_element(message_box).click()
                    action.send_keys(Keys.ENTER)
                    action.perform()
                    time.sleep(5)

                logging.info('Done for %s - %s', name, contactNumber)
        except StaleElementReferenceException as exp:
            logging.error('Stale element exception for SL NO: [%s], name: [%s], contactNumber: [%s], idx: [%d]', slNo, name, contactNumber, idx)
            self.exceptionList[config.STALE_ELEMENT] = exp
            self.failCnt += 1
            df = pd.DataFrame([[self.failCnt, name, contactNumber]])
            with pd.ExcelWriter(respFile, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                df.to_excel(writer, index=False, header=False, startrow=self.failCnt)
            pass
        except Exception as exp:
            logging.error('Following exception for SL NO: [%s], name: [%s], contactNumber: [%s], idx: [%d]', slNo, name, contactNumber, idx)
            logging.error(exp)
            self.exceptionList[config.PROCESS_EXCEPTION] = exp
            self.failCnt += 1
            df = pd.DataFrame([[self.failCnt, name, contactNumber]])
            with pd.ExcelWriter(respFile, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                df.to_excel(writer, index=False, header=False, startrow=self.failCnt)
            pass

    def close(self):
        # Close Chrome browser
        if self.driver:
            self.driver.quit()

# Main starts here
def select_file(fileStr, fileTypes, widget):
    fileName = fd.askopenfilenames(title='Open a file', initialdir='.', filetypes=((fileStr, fileTypes), ("all files","*.*")))
    widget.delete(0, tk.END)
    widget.insert(tk.END, fileName)

def clear_files(xlsEntry, txtEntry, imgEntry):
    xlsEntry.delete(0, tk.END)
    txtEntry.delete(0, tk.END)
    imgEntry.delete(0, tk.END)

def call_whatsapp(xlsEntry, txtEntry, imgEntry, resText):
    dirName = os.path.dirname(xlsEntry.get())
    xlsFile = os.path.basename(xlsEntry.get())
    msgFiles = txtEntry.get().split(' ')
    imgFiles = imgEntry.get().split(' ')

    logging.info('Dir name [%s]', dirName)
    logging.info('Excel file [%s]', xlsFile)
    logging.info('Text file(s) [%s]', msgFiles)
    logging.info('Image file(s) [%s]', imgFiles)
    
    for idx in range(0, len(msgFiles)):
        msgFile = os.path.basename(msgFiles[idx].strip())
        imgFile = ''

        if idx < len(imgFiles):
            imgFile = os.path.basename(imgFiles[idx].strip())

        logging.info('Set [%d]', idx+1)
        logging.info('Text file(s) [%s]', msgFile)
        logging.info('Image file(s) [%s]', imgFile)

        errorType = ''
        errorMsg = ''
        failVal = call_whatsapp_for_one(xlsFile, msgFile, imgFile, str(idx+1))
        if failVal:
            for errorType in failVal: 
                errorMsg = failVal[errorType]
            logging.info('Failure after processing [%s] [%s]', errorType, errorMsg)

        resText.delete('1.0', tk.END)    
        # check if all messages were sent successfully
        if errorType != '':
            dispTxt = 'Errors while processing set no ' + str(idx+1) + ':' + errorMsg
            resText.insert('end', dispTxt)
            break
    
    if errorType == '':
        resText.insert('end', 'Completed Successfully. Close this window')

def call_whatsapp_for_one(xlsFile, msgFile, imgFile, msgInd):
    failVal = {}

    logging.info('SEND_MSG is set to [%s]', config.SEND_MSG)
    if config.SEND_MSG == 'N':
        return failVal
        
    args = {}
    args['xls_file'] = xlsFile
    args['msg_file'] = msgFile
    args['img_file'] = imgFile
    args['msg_ind'] = msgInd

    logging.info('args [%s]', args)
    
    # send whatsapp messages
    whatsapp = WhatsappBulkMessage(**args)
    whatsapp.perform_task()
    lenExp = len(WhatsappBulkMessage.exceptionList)

    if lenExp != 0:
        logging.error('Exception occured [%d]. Handling in the caller', lenExp)
        # Exceptions have occurred
        failVal = WhatsappBulkMessage.exceptionList
        logging.error('Exceptions: [%s]', failVal)
    return failVal

class WhatsappDB:
    def __init__(self):
        logging.info('Entering WhatsappDB init')
        self.mydb = config.dbConnect
        
        self.createSessionTable()
        self.createEnrolmentTable()
        self.createEnrolmentHistoryTable()        
        self.createPaymentTable()
        self.createErrorTable()
        logging.info('Exiting WhatsappDB init')
    
    def createSessionTable(self):
        logging.info('Checking started for SESSION_DETAILS')
        with self.mydb.cursor() as cursor:
            sql = """SELECT count(table_name) cnt 
                FROM information_schema.tables 
                WHERE table_name = %s"""
            cursor.execute(sql, ('SESSION_DETAILS',))
            result = cursor.fetchone()
            
        if result['cnt'] == 1:
            logging.info('SESSION_DETAILS exists')
            return
        
        logging.info('SESSION_DETAILS does not exist. Creating...')
        with self.mydb.cursor() as cursor:
            sql = """CREATE TABLE SESSION_DETAILS (
                name VARCHAR(255), 
                mobile_num VARCHAR(20), 
                session_start_date_time DATETIME, 
                message VARCHAR(4096), 
                form_submit_time DATETIME(6), 
                msg_sent_time DATETIME(6), 
                num_hours INT, 
                recording_links VARCHAR(1024), 
                session_type VARCHAR(64), 
                processed CHAR(1), 
                unique index idx_session_mobfrm (mobile_num, form_submit_time))"""
            cursor.execute(sql)
        logging.info('SESSION_DETAILS created')

    def createPaymentTable(self):
        logging.info('Checking started for PAYMENT_DETAILS')
        with self.mydb.cursor() as cursor:
            sql = """SELECT count(table_name) cnt 
                FROM information_schema.tables 
                WHERE table_name = %s"""
            cursor.execute(sql, ('PAYMENT_DETAILS',))
            result = cursor.fetchone()
            
        if result['cnt'] == 1:
            logging.info('PAYMENT_DETAILS exists')
            return
        
        logging.info('PAYMENT_DETAILS does not exist. Creating...')
        with self.mydb.cursor() as cursor:
            sql = """CREATE TABLE PAYMENT_DETAILS (
                name VARCHAR(255), 
                mobile_num VARCHAR(20), 
                payment_amount INT, 
                payment_date DATE, 
                text VARCHAR(4096), 
                form_submit_time DATETIME(6), 
                msg_sent_time DATETIME(6), 
                processed CHAR(1), 
                unique index idx_payment_mobfrm (mobile_num, form_submit_time))"""
            cursor.execute(sql)
        logging.info('PAYMENT_DETAILS created')

    def createEnrolmentTable(self):
        logging.info('Checking started for ENROL_DETAILS')
        with self.mydb.cursor() as cursor:
            sql = """SELECT count(table_name) cnt 
                FROM information_schema.tables 
                WHERE table_name = %s"""
            cursor.execute(sql, ('ENROL_DETAILS',))
            result = cursor.fetchone()
                
        if result['cnt'] == 1:
            logging.info('ENROL_DETAILS exists')
            return
        
        logging.info('ENROL_DETAILS does not exist. Creating...')
        with self.mydb.cursor() as cursor:
            sql = """CREATE TABLE ENROL_DETAILS (
                name VARCHAR(255), 
                mobile_num VARCHAR(20), 
                email_id VARCHAR(100), 
                package VARCHAR(50), 
                birth_day DATE, 
                member_start_date DATE, 
                notes VARCHAR (255), 
                calendar_id VARCHAR(100), 
                event_id VARCHAR(100), 
                meet_link VARCHAR(255), 
                summary VARCHAR(255), 
                session_start_date_time DATETIME, 
                message VARCHAR(4096), 
                form_submit_time DATETIME(6), 
                total_payment INT, 
                remaining_payment INT, 
                payment_frequency VARCHAR(16), 
                payment_period INT, 
                msg_sent_time DATETIME(6), 
                processed CHAR(1), 
                total_sessions INT, 
                completed_sessions DECIMAL(3,1), 
                session_duration DECIMAL(3,1), 
                session_frequency VARCHAR(16), 
                unique index idx_enrol_name (name), 
                unique idx_enrol_mobile (mobile_num))"""
            cursor.execute(sql)
        logging.info('ENROL_DETAILS created')

    def createEnrolmentHistoryTable(self):
        logging.info('Checking started for ENROL_HISTORY')
        with self.mydb.cursor() as cursor:
            sql = """SELECT count(table_name) cnt 
                FROM information_schema.tables 
                WHERE table_name = %s"""
            cursor.execute(sql, ('ENROL_HISTORY',))
            result = cursor.fetchone()
                
        if result['cnt'] == 1:
            logging.info('ENROL_HISTORY exists')
            return
        
        logging.info('ENROL_HISTORY does not exist. Creating...')
        with self.mydb.cursor() as cursor:
            sql = """CREATE TABLE ENROL_HISTORY (
                name VARCHAR(255), 
                mobile_num VARCHAR(20), 
                email_id VARCHAR(100), 
                package VARCHAR(50), 
                birth_day DATE, 
                member_start_date DATE, 
                notes VARCHAR (255), 
                calendar_id VARCHAR(100), 
                event_id VARCHAR(100), 
                meet_link VARCHAR(255), 
                summary VARCHAR(255), 
                session_start_date_time DATETIME, 
                message VARCHAR(4096), 
                form_submit_time DATETIME(6), 
                total_payment INT, 
                remaining_payment INT, 
                payment_frequency VARCHAR(16), 
                payment_period INT, 
                msg_sent_time DATETIME(6), 
                processed CHAR(1), 
                total_sessions INT, 
                completed_sessions DECIMAL(3,1), 
                session_duration DECIMAL(3,1), 
                session_frequency VARCHAR(16), 
                unique index idx_enrol_name (name), 
                unique idx_enrol_mobile (mobile_num))"""
            cursor.execute(sql)
        logging.info('ENROL_HISTORY created')

    def createErrorTable(self):
        logging.info('Checking started for ERROR_DETAILS')
        with self.mydb.cursor() as cursor:
            sql = """SELECT count(table_name) cnt 
                FROM information_schema.tables 
                WHERE table_name = %s"""
            cursor.execute(sql, ('ERROR_DETAILS',))
            result = cursor.fetchone()
                
        if result['cnt'] == 1:
            logging.info('ERROR_DETAILS exists')
            return
        
        logging.info('ERROR_DETAILS does not exist. Creating...')
        with self.mydb.cursor() as cursor:
            sql = """CREATE TABLE ERROR_DETAILS (
                process VARCHAR(255), 
                process_date_time DATETIME, 
                name VARCHAR(255), 
                mobile_num VARCHAR(20), 
                error_type VARCHAR(2), 
                error_msg VARCHAR(4096), 
                error_cre_time DATETIME(6), 
                index err_mobile_idx (mobile_num))"""
            cursor.execute(sql)
        logging.info('ERROR_DETAILS created')
        
    def sessionInsert(self, sessionDict):
        with self.mydb.cursor() as cursor:
            sql = """INSERT INTO SESSION_DETAILS(
                    name, 
                    mobile_num, 
                    session_start_date_time, 
                    message, 
                    form_submit_time, 
                    msg_sent_time, 
                    num_hours, 
                    recording_links, 
                    session_type, 
                    processed
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""
    
            val = (
                sessionDict['name'], 
                sessionDict['mobileNum'], 
                sessionDict['sesStartDateTime'].strftime("%Y-%m-%d %H:%M"), 
                sessionDict['message'], 
                sessionDict['formSbmTime'].strftime("%Y-%m-%d %H:%M:%S.%f"), 
                sessionDict['msgSentTime'].strftime("%Y-%m-%d %H:%M:%S.%f"), 
                sessionDict['numHours'], 
                sessionDict['recordingLinks'], 
                sessionDict['sessionType'], 
                sessionDict['prcFlg']
            )
            cursor.execute(sql, val)
        logging.info("Session record inserted for [%s] [%s]", sessionDict['name'], sessionDict['mobileNum'])
    
    def sessionUpdate(self, sessionDict):
        with self.mydb.cursor() as cursor:
            sql = """UPDATE SESSION_DETAILS SET 
                name = %s, 
                session_start_date_time = %s, 
                message = %s, 
                msg_sent_time = %s, 
                num_hours = %s, 
                recording_links = %s, 
                session_type = %s, 
                processed = %s
                WHERE mobile_num = %s and form_submit_time = %s"""
            
            val = (
                sessionDict['name'], 
                sessionDict['sesStartDateTime'].strftime("%Y-%m-%d %H:%M"), 
                sessionDict['message'], 
                sessionDict['msgSentTime'].strftime("%Y-%m-%d %H:%M:%S.%f"), 
                sessionDict['numHours'], 
                sessionDict['recordingLinks'], 
                sessionDict['sessionType'], 
                sessionDict['prcFlg'],
                sessionDict['mobileNum'], 
                sessionDict['formSbmTime'].strftime("%Y-%m-%d %H:%M:%S.%f") 
            )
            cursor.execute(sql, val)
        logging.info("Session record updated for [%s] [%s]", sessionDict['name'], sessionDict['mobileNum'])
    
    def sessionSelect(self, mobileNum, formSbmTime):
        row = ''
        with self.mydb.cursor() as cntcursor:
            sql = """SELECT count(1) cnt 
                FROM SESSION_DETAILS 
                WHERE mobile_num = %s 
                AND form_submit_time = %s"""
                
            val = (mobileNum, formSbmTime.strftime("%Y-%m-%d %H:%M:%S.%f"))
            cntcursor.execute(sql, val)   
            if cntcursor.fetchone()['cnt'] != 0:
                logging.info("Session Record found [%s] ", mobileNum)
                with self.mydb.cursor() as rowcursor:
                    sql = """SELECT
                        name,
                        mobile_num mobileNum 
                        session_start_date_time sesStartDateTime, 
                        message, 
                        form_submit_time formSbmTime, 
                        msg_sent_time msgSentTime, 
                        num_hours numHours, 
                        recording_links recordingLinks, 
                        session_type sessionType, 
                        processed prcFlg
                        FROM SESSION_DETAILS
                        WHERE mobile_num = %s 
                        AND form_submit_time = %s"""
                    val = (mobileNum, formSbmTime.strftime("%Y-%m-%d %H:%M:%S.%f"))
                    rowcursor.execute(sql, val)
                    row = rowcursor.fetchone()
                logging.info("Session record selected [%s] ", mobileNum)
            else:
                logging.info("No Session Record found [%s] ", mobileNum)
                row = ''
        return row
        
    def paymentInsert(self, paymentDict):
        with self.mydb.cursor() as cursor:
            sql = """INSERT INTO PAYMENT_DETAILS(
                    name, 
                    mobile_num, 
                    payment_amount,
                    payment_date,
                    text, 
                    form_submit_time, 
                    msg_sent_time, 
                    processed
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s)"""
    
            val = (
                paymentDict['name'], 
                paymentDict['mobileNum'], 
                paymentDict['payAmt'], 
                paymentDict['payDate'].strftime("%Y-%m-%d"), 
                paymentDict['text'], 
                paymentDict['formSbmTime'].strftime("%Y-%m-%d %H:%M:%S.%f"), 
                paymentDict['msgSentTime'].strftime("%Y-%m-%d %H:%M:%S.%f"), 
                paymentDict['prcFlg']
            )
            cursor.execute(sql, val)
        logging.info("Payment record inserted for [%s] [%s]", paymentDict['name'], paymentDict['mobileNum'])
    
    def paymentUpdate(self, paymentDict):
        with self.mydb.cursor() as cursor:
            sql = """UPDATE PAYMENT_DETAILS SET 
                name = %s, 
                payment_amount = %s, 
                payment_date = %s, 
                text = %s, 
                msg_sent_time = %s, 
                processed = %s
                WHERE mobile_num = %s and form_submit_time = %s"""
            
            val = (
                paymentDict['name'], 
                paymentDict['payAmt'], 
                paymentDict['payDate'].strftime("%Y-%m-%d"), 
                paymentDict['text'], 
                paymentDict['msgSentTime'].strftime("%Y-%m-%d %H:%M:%S.%f"), 
                paymentDict['prcFlg'],
                paymentDict['mobileNum'], 
                paymentDict['formSbmTime'].strftime("%Y-%m-%d %H:%M:%S.%f") 
            )
            cursor.execute(sql, val)
        logging.info("Payment record updated for [%s] [%s]", paymentDict['name'], paymentDict['mobileNum'])

    def paymentSelect(self, mobileNum, formSbmTime):
        row = ''
        with self.mydb.cursor() as cntcursor:
            sql = """SELECT count(1) cnt 
                FROM PAYMENT_DETAILS 
                WHERE mobile_num = %s 
                AND form_submit_time = %s"""
                
            val = (mobileNum, formSbmTime.strftime("%Y-%m-%d %H:%M:%S.%f"))
            cntcursor.execute(sql, val)   
            if cntcursor.fetchone()['cnt'] != 0:
                logging.info("Payment Record found [%s] ", mobileNum)
                with self.mydb.cursor() as rowcursor:
                    sql = """SELECT
                        name,
                        mobile_num mobileNum, 
                        payment_amount payAmt, 
                        payment_date payDate, 
                        text, 
                        form_submit_time formSbmTime, 
                        msg_sent_time msgSentTime, 
                        processed prcFlg
                        FROM SESSION_DETAILS
                        WHERE mobile_num = %s 
                        AND form_submit_time = %s"""
                    val = (mobileNum, formSbmTime.strftime("%Y-%m-%d %H:%M:%S.%f"))
                    rowcursor.execute(sql, val)
                    row = rowcursor.fetchone()
                logging.info("Payment record selected [%s] ", mobileNum)
            else:
                logging.info("No Payment Record found [%s] ", mobileNum)
                row = ''
        return row
        
    def enrolInsert(self, enrolDict):
        with self.mydb.cursor() as cursor:
            sql = """INSERT INTO ENROL_DETAILS(
                    name, 
                    mobile_num, 
                    email_id, 
                    package, 
                    birth_day, 
                    member_start_date, 
                    notes, 
                    calendar_id, 
                    event_id, 
                    meet_link, 
                    summary, 
                    session_start_date_time, 
                    message, 
                    form_submit_time, 
                    total_payment, 
                    remaining_payment, 
                    payment_frequency, 
                    payment_period, 
                    msg_sent_time, 
                    processed,
                    total_sessions,
                    completed_sessions,
                    session_duration,
                    session_frequency
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""
            
            val = (
                enrolDict['name'], 
                enrolDict['mobileNum'], 
                enrolDict['emailId'], 
                enrolDict['package'], 
                enrolDict['birthDay'].strftime("%Y-%m-%d"), 
                enrolDict['memStartDate'].strftime("%Y-%m-%d"), 
                enrolDict['notes'], 
                enrolDict['calendarId'], 
                enrolDict['eventId'], 
                enrolDict['meetLink'], 
                enrolDict['summary'], 
                enrolDict['sesStartDateTime'].strftime("%Y-%m-%d %H:%M"), 
                enrolDict['message'], 
                enrolDict['formSbmTime'].strftime("%Y-%m-%d %H:%M:%S.%f"), 
                enrolDict['totPay'], 
                enrolDict['remPay'], 
                enrolDict['frqPay'], 
                enrolDict['numPay'], 
                enrolDict['msgSentTime'].strftime("%Y-%m-%d %H:%M:%S.%f"), 
                enrolDict['prcFlg'], 
                enrolDict['totalSessions'], 
                enrolDict['completedSessions'],
                enrolDict['sessionDuration'],
                enrolDict['frqSes']
            )
            cursor.execute(sql, val)
        logging.info("Enrolment inserted for [%s] [%s]", enrolDict['name'], enrolDict['mobileNum'])

    def enrolUpdate(self, enrolDict):
        with self.mydb.cursor() as cursor:
            sql = """UPDATE ENROL_DETAILS SET 
                name = %s, 
                email_id = %s, 
                package = %s, 
                birth_day = %s, 
                member_start_date = %s, 
                notes = %s, 
                calendar_id = %s, 
                event_id = %s, 
                meet_link = %s, 
                summary = %s, 
                session_start_date_time = %s, 
                message = %s, 
                form_submit_time  = %s, 
                total_payment  = %s, 
                remaining_payment  = %s, 
                payment_frequency  = %s, 
                payment_period  = %s, 
                msg_sent_time = %s, 
                processed = %s, 
                total_sessions = %s,
                completed_sessions = %s, 
                session_duration = %s,
                session_frequency = %s 
                WHERE mobile_num = %s"""
                
            val = (
                enrolDict['name'], 
                enrolDict['emailId'], 
                enrolDict['package'], 
                enrolDict['birthDay'].strftime("%Y-%m-%d"), 
                enrolDict['memStartDate'].strftime("%Y-%m-%d"), 
                enrolDict['notes'], 
                enrolDict['calendarId'], 
                enrolDict['eventId'], 
                enrolDict['meetLink'], 
                enrolDict['summary'], 
                enrolDict['sesStartDateTime'].strftime("%Y-%m-%d %H:%M"), 
                enrolDict['message'], 
                enrolDict['formSbmTime'].strftime("%Y-%m-%d %H:%M:%S.%f"), 
                enrolDict['totPay'], 
                enrolDict['remPay'], 
                enrolDict['frqPay'], 
                enrolDict['numPay'], 
                enrolDict['msgSentTime'].strftime("%Y-%m-%d %H:%M:%S.%f"), 
                enrolDict['prcFlg'], 
                enrolDict['totalSessions'],
                enrolDict['completedSessions'],
                enrolDict['sessionDuration'],
                enrolDict['frqSes'],
                enrolDict['mobileNum'] 
            )
            
            cursor.execute(sql, val)
        logging.info("Enrolment updated for [%s] [%s]", enrolDict['name'], enrolDict['mobileNum'])

    def enrolUpdateSessions(self, mobileNum, numHours):
        with self.mydb.cursor() as cursor:
            sql = """UPDATE ENROL_DETAILS SET 
                completed_sessions = completed_sessions + %s 
                WHERE mobile_num = %s"""
                
            val = (numHours, mobileNum,)
            
            cursor.execute(sql, val)
        logging.info("Enrolment sessions count updated for [%s]", mobileNum)

    def enrolUpdatePayments(self, mobileNum, payAmt):
        with self.mydb.cursor() as cursor:
            sql = """UPDATE ENROL_DETAILS SET 
                remaining_payment = remaining_payment - %s
                WHERE mobile_num = %s"""
                
            val = (payAmt, mobileNum)
            
            cursor.execute(sql, val)
        logging.info("Enrolment payments data updated for [%s]", mobileNum)

    def enrolSelect(self, mobileNum):
        with self.mydb.cursor() as cntcursor:
            sql = """SELECT count(1) cnt 
                FROM ENROL_DETAILS 
                WHERE mobile_num = %s"""
            val = (mobileNum,)
            cntcursor.execute(sql, val)   
            if cntcursor.fetchone()['cnt'] != 0:
                logging.info("Enrolment Record found [%s] ", mobileNum)
                with self.mydb.cursor() as rowcursor:
                    sql = """SELECT
                        name,
                        mobile_num mobileNum,
                        email_id emailId, 
                        package, 
                        birth_day birthDay, 
                        member_start_date memStartDate, 
                        notes, 
                        calendar_id calendarId, 
                        event_id eventId, 
                        meet_link meetLink,
                        summary, 
                        session_start_date_time sesStartDateTime,
                        message, 
                        form_submit_time formSbmTime, 
                        total_payment totPay, 
                        remaining_payment remPay, 
                        payment_frequency frqPay, 
                        payment_period numPay, 
                        msg_sent_time msgSentTime,
                        processed prcFlg, 
                        total_sessions totalSessions,
                        completed_sessions completedSessions,
                        session_duration sessionDuration,
                        session_frequency frqSes 
                        FROM ENROL_DETAILS 
                        WHERE mobile_num = %s"""
                    val = (mobileNum,)
                    rowcursor.execute(sql, val)   
                    row = rowcursor.fetchone()
                logging.info("Enrolment record selected [%s] ", mobileNum)
            else:
                logging.info("No Enrolment Record found [%s] ", mobileNum)
                row = ''
        return row

    def enrolHistoryInsert(self, enrolDict):
        with self.mydb.cursor() as cursor:
            sql = """INSERT INTO ENROL_HISTORY(
                    name, 
                    mobile_num, 
                    email_id, 
                    package, 
                    birth_day, 
                    member_start_date, 
                    notes, 
                    calendar_id, 
                    event_id, 
                    meet_link, 
                    summary, 
                    session_start_date_time, 
                    message, 
                    form_submit_time, 
                    total_payment, 
                    remaining_payment, 
                    payment_frequency, 
                    payment_period, 
                    msg_sent_time, 
                    processed,
                    total_sessions,
                    completed_sessions,
                    session_duration,
                    session_frequency
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"""
            
            val = (
                enrolDict['name'], 
                enrolDict['mobileNum'], 
                enrolDict['emailId'], 
                enrolDict['package'], 
                enrolDict['birthDay'].strftime("%Y-%m-%d"), 
                enrolDict['memStartDate'].strftime("%Y-%m-%d"), 
                enrolDict['notes'], 
                enrolDict['calendarId'], 
                enrolDict['eventId'], 
                enrolDict['meetLink'], 
                enrolDict['summary'], 
                enrolDict['sesStartDateTime'].strftime("%Y-%m-%d %H:%M"), 
                enrolDict['message'], 
                enrolDict['formSbmTime'].strftime("%Y-%m-%d %H:%M:%S.%f"), 
                enrolDict['totPay'], 
                enrolDict['remPay'], 
                enrolDict['frqPay'], 
                enrolDict['numPay'], 
                enrolDict['msgSentTime'].strftime("%Y-%m-%d %H:%M:%S.%f"), 
                enrolDict['prcFlg'], 
                enrolDict['totalSessions'], 
                enrolDict['completedSessions'],
                enrolDict['sessionDuration'],
                enrolDict['frqSes']
            )
            cursor.execute(sql, val)
        logging.info("Enrolment history inserted for [%s] [%s]", enrolDict['name'], enrolDict['mobileNum'])

    def errorInsert(self, errorDict):
        with self.mydb.cursor() as cursor:
            sql = """INSERT INTO ERROR_DETAILS(
                process, 
                process_date_time, 
                name, 
                mobile_num, 
                error_type, 
                error_msg, 
                error_cre_time) 
                VALUES (%s, %s, %s, %s, %s, %s, %s)"""
            val = (
                errorDict['process'], 
                errorDict['processDateTime'].strftime("%Y-%m-%d %H:%M:%S.%f"), 
                errorDict['name'], 
                errorDict['mobileNum'], 
                errorDict['errorType'], 
                errorDict['errorMsg'], 
                errorDict['errorCreTime'].strftime("%Y-%m-%d %H:%M:%S.%f"))
            cursor.execute(sql, val)
        logging.info("Error record inserted for [%s] [%s]", errorDict['name'], errorDict['mobileNum'])
    
    def close(self, commitFlg):
        logging.info("DB close [%s] ", commitFlg)
        if commitFlg == 'Y':
            self.mydb.commit()
        else:
            self.mydb.rollback()
        self.mydb.close()
            
def batch_process():
    def processSessions(service):
        logging.info('Entering processSessions')
        # Call the Sheets API
        sheet = service.spreadsheets()

        result = sheet.values().get(spreadsheetId=config.SESSION_REQ_ID, range=config.SESSION_REQ_RANGE).execute()
        values = result.get('values', [])

        if not values:
            logging.error('No data found.')
            return 1
        
        # If any data is found in response spreadsheet, it means google apps script still hasn't processed the previous response. Stop until that's done
        result = sheet.values().get(spreadsheetId=config.SESSION_RES_ID, range=config.SESSION_RES_RANGE).execute()
        respValues = result.get('values', [])

        if respValues:
            logging.error('Google apps script still has not processed the previous response. Stop until that is done.')
            return 1

        logging.info("values: [%s]", values)
        xlsFile = "whatsapp.xlsx"
        msgFile = "whatsapp.txt"
        
        whatsappDB = WhatsappDB()
        result = []
        
        for curColSet in values:
            # Check if this record is already processed
            name = curColSet[0]
            mobileNum = curColSet[1]
            sesStartDateTime = (datetime.fromisoformat(curColSet[2][:-1]+'000+00:00')).astimezone()
            partMessage = curColSet[3]
            recdFormSbmTimeStr = curColSet[4]
            formSbmTime = (datetime.fromisoformat(recdFormSbmTimeStr[:-1]+'000+00:00')).astimezone()
            numHours = round(float(curColSet[5]), 1)
            recordingLinks = curColSet[6]
            sessionType = curColSet[7]
            
            # Check if the person is enrolled
            enrolRec = whatsappDB.enrolSelect(mobileNum)
            if enrolRec == '':
                logging.error("Enrolment record itself does not exist: [%s] [%s] [%s]", name, mobileNum, formSbmTime)
                result.append([name, mobileNum, recdFormSbmTimeStr, 'N', name + ' is not yet enrolled'])
                continue
                
            logging.info(enrolRec)
            if sesStartDateTime == '':
                # Pick the default session_start_date_time
                sesStartDateTime = datetime.fromisoformat(enrolRec['sesStartDateTime'].strftime('%Y-%m-%dT%H:%i')) + timedelta(weeks=enrolRec['completedSessions'])
                
            logging.info("Next session date time: [%s]", sesStartDateTime)
            logging.info("Before fetch Session rec: [%s] [%s], start [%s] Form Submit: [%s]", name, mobileNum, sesStartDateTime, formSbmTime)
            
            existingRec = whatsappDB.sessionSelect(mobileNum, formSbmTime)
            if existingRec != '' and existingRec['prcFlg'] == 'Y':
                # Already processed successfully. Tell Google apps script to remove it 
                logging.info("Record exists and already processed successfully: [%s] [%s] [%s]", name, mobileNum, formSbmTime)
                result.append([name, mobileNum, recdFormSbmTimeStr, 'Y', 'success'])
                continue

            numSessions = decimal.Decimal(numHours) / enrolRec['sessionDuration']

            # Construct the full message
            message = "Thank you for attending the 1:1 session on " + datetime.strftime(sesStartDateTime, '%d %b %Y')
            message = message + " at " + datetime.strftime(sesStartDateTime, '%I:%M%p') + " IST. I hope you found it useful😊\n\n"
            
            remainSessions = enrolRec['totalSessions'] - (enrolRec['completedSessions'] + numSessions)
            match remainSessions * enrolRec['sessionDuration']:
                case 0:
                    endDate = datetime.strptime(enrolRec['memStartDate'], "%Y-%m-%d") + relativedelta(years=1)
                    message = message + "You have completed all your sessions. However, your membership for Rajani's BioHub Inner Circle will continue till " + endDate.strftime("%d %b, %Y")
                case 0.5|1:
                    message = message + "You have " + str(remainSessions) + " bonus session left. Let's make the best use of it 👍" 
                case 1.5:
                    message = message + "You have " + str(remainSessions) + " bonus sessions left. Let's make the best use of it 👍" 
                case 2:
                    message = message + "Your regular sessions are completed. We'll begin your bonus sessions from now on." 
                case _:
                    message = message + "The number of your remaining sessions: " + str(remainSessions)
            
            message = message + "\n\n"
            message = message + partMessage
            message = message + "PS: The session recording is given below: \n" + recordingLinks + "\n\n"
            
            logging.info("Message length: [%d]", len(message))
            
            df = pd.DataFrame([['SL NO', 'NAME', 'CONTACT DETAILS'], ['1', name, mobileNum]])
            with pd.ExcelWriter(xlsFile, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, header=False, startrow=0)

            with open(msgFile, "w", encoding="utf-8") as msgWriter:
                msgWriter.write(message)

            errorType = ''
            errorMsg = ''
            # Call whatsapp messaging for the given mobile number
            failVal = call_whatsapp_for_one(xlsFile, msgFile, '', 'A')
            if failVal:
                for errorType in failVal: 
                    errorMsg = failVal[errorType]
                logging.info('Failure after processing [%s] [%s]', errorType, errorMsg)

            msgSentTime = datetime.now()

            # Build the session row
            sessionDict = {
                'name': name, 
                'mobileNum': mobileNum, 
                'sesStartDateTime': sesStartDateTime, 
                'message': message, 
                'formSbmTime': formSbmTime, 
                'msgSentTime': msgSentTime, 
                'numHours': numHours, 
                'recordingLinks': recordingLinks, 
                'sessionType': sessionType, 
                'prcFlg': ''
            }
            
            # check if the message was sent successfully
            if errorType != '':
                logging.error('Errors while processing for %s. Check logs', mobileNum)
                # Build the error record
                errorDict = {
                    'process': 'session', 
                    'processDateTime': formSbmTime, 
                    'name': name, 
                    'mobileNum': mobileNum, 
                    'errorType': errorType, 
                    'errorMsg': errorMsg, 
                    'errorCreTime': msgSentTime
                }
                whatsappDB.errorInsert(errorDict); 
                sessionDict['prcFlg'] = 'N'
                if existingRec == '':
                    logging.info("No session record exists. So inserting a failure record for : [%s] [%s] [%s]", name, mobileNum, formSbmTime)
                    whatsappDB.sessionInsert(sessionDict)
                    logging.info('Updating completed sessions %s.', mobileNum)
                    whatsappDB.enrolUpdateSessions(mobileNum, numSessions)
                else:
                    logging.info("Session Record exists. So updating the failure record for : [%s] [%s] [%s]", name, mobileNum, formSbmTime)
                    whatsappDB.sessionUpdate(sessionDict)                    
                result.append([name, mobileNum, recdFormSbmTimeStr, 'N', errorMsg])
            else:
                logging.info('Completed Successfully for %s %d', mobileNum, len(message))
                sessionDict['prcFlg'] = 'Y'
                if existingRec == '':
                    logging.info("No session record exists. So inserting a success record for : [%s] [%s] [%s]", name, mobileNum, formSbmTime)
                    whatsappDB.sessionInsert(sessionDict)
                    logging.info('Updating completed sessions %s.', mobileNum)
                    whatsappDB.enrolUpdateSessions(mobileNum, numSessions)
                else:
                    logging.info("Session Record exists. So updating the record as successfully processed : [%s] [%s] [%s]", name, mobileNum, formSbmTime)
                    whatsappDB.sessionUpdate(sessionDict)                    
                result.append([name, mobileNum, recdFormSbmTimeStr, 'Y', 'success'])
        
        logging.info('Calling function removeProcessedSessions')
        body = {
            'values': result
        }
        updResp = sheet.values().update(spreadsheetId=config.SESSION_RES_ID, range=config.SESSION_RES_RANGE,
            valueInputOption='USER_ENTERED', body=body).execute()
        logging.info('%d cells updated.', updResp.get('updatedCells'))
        
        # Finally, remove these prosessed entries from request sheet in Google drive.
        request = {"function": "removeProcessedSessions"}
        scriptService = build('script', 'v1', credentials=creds)
        response = scriptService.scripts().run(scriptId=config.SCRIPT_ID, body=request).execute()
        logging.info('Returned from script call with %s', response)
        if 'error' in response:
            error = response['error']['details'][0]
            logging.error("Script error message: %s", error['errorMessage'])
            whatsappDB.close('N')
            return 1
        else:
            logging.info('script success message: %s', response['response'].get('result'))
            logging.info(response)
            whatsappDB.close('Y')
            return 0
        logging.info('Exiting processSessions')

    def processEnrolments(service):
        logging.info('Entering processEnrolments')
        # Call the Sheets API
        sheet = service.spreadsheets()

        # Check for new enrolments
        result = sheet.values().get(spreadsheetId=config.ENROLMENT_REQ_ID, range=config.ENROLMENT_REQ_RANGE).execute()
        enrolValues = result.get('values', [])
        if not enrolValues:
            logging.error('No new enrolments.')
            return 1
        
        # If any data is found in response spreadsheet, it means google apps script still hasn't processed the previous response. Stop until that's done
        result = sheet.values().get(spreadsheetId=config.ENROLMENT_RES_ID, range=config.ENROLMENT_RES_RANGE).execute()
        respValues = result.get('values', [])

        if respValues:
            logging.error('Google apps script still has not processed the previous enrolment response. Stop until that is done.')
            return 1

        logging.info("Enrolment values: %s", enrolValues)
        xlsFile = "whatsapp.xlsx"
        msgFile = "whatsapp.txt"
        
        whatsappDB = WhatsappDB()
        result = []
        
        for curColSet in enrolValues:
            logging.info(curColSet)
            logging.info(curColSet[0])
            recdFormSbmTimeStr = curColSet[0]
            formSbmTime = (datetime.fromisoformat(recdFormSbmTimeStr[:-1]+'000+00:00')).astimezone()
            logging.info(curColSet[1])
            name = curColSet[1]
            logging.info(curColSet[2])
            mobileNum = curColSet[2]
            logging.info(curColSet[3])
            emailId = curColSet[3]
            logging.info(curColSet[4])
            package = curColSet[4]
            logging.info(curColSet[5])
            birthDayStr = curColSet[5]
            # If birth day is below 1970, astimezone() does not work
            logging.info("birth day [%d]", int(birthDayStr[0:4]))
            if int(birthDayStr[0:4]) < 1970:
                diff = 1970 - int(birthDayStr[0:4])
                birthDayStr = '1970' + birthDayStr[4:]
                birthDay = (datetime.fromisoformat(birthDayStr[:-1]+'000+00:00')).astimezone() + relativedelta(years=-diff)
            else:
                birthDay = (datetime.fromisoformat(birthDayStr[:-1]+'000+00:00')).astimezone()
            logging.info(curColSet[6])
            memStartDate = (datetime.fromisoformat(curColSet[6][:-1]+'000+00:00')).astimezone()
            logging.info(curColSet[7])
            notes = curColSet[7]
            logging.info(curColSet[8])
            calendarId = curColSet[8]
            logging.info(curColSet[9])
            eventId = curColSet[9]
            logging.info(curColSet[10])
            meetLink = curColSet[10]
            logging.info(curColSet[11])
            summary = curColSet[11]
            logging.info(curColSet[12])
            if curColSet[12] != '':
                sesStartDateTime = (datetime.fromisoformat(curColSet[12][:-1]+'000+00:00')).astimezone()
            else:
                sesStartDateTime = memStartDate
            logging.info(curColSet[13])
            message = curColSet[13]            
            logging.info(curColSet[14])
            totPay = int(curColSet[14])            
            logging.info(curColSet[15])
            frqPay = curColSet[15]
            logging.info(curColSet[16])
            numPay = int(curColSet[16])
            logging.info(curColSet[17])
            totalSessions = int(curColSet[18])
            sessionDuration = round(float(curColSet[18]), 1)
            frqSes = curColSet[19]
            logging.info("Enrolment Record details: [%s] [%s], Form Submitted time: [%s]", name, mobileNum, formSbmTime)

            existingRec = whatsappDB.enrolSelect(mobileNum)
            if existingRec != '' and existingRec['package'] == package and existingRec['prcFlg'] == 'Y':
                # Already processed successfully. Tell Google apps script to remove it 
                logging.info("Enrolment exists and already processed successfully: [%s] [%s] [%s]", name, mobileNum, formSbmTime)
                result.append([name, mobileNum, recdFormSbmTimeStr, 'Y', 'success'])
                continue

            if existingRec != '' and existingRec['package'] != package and existingRec['prcFlg'] == 'Y':
                # This member is joining a new package. Move the existing record to history  
                logging.info("The member [%s] has joined a new package: [%s]", name, package)
                whatsappDB.enrolHistoryInsert(existingRec)

            df = pd.DataFrame([['SL NO', 'NAME', 'CONTACT DETAILS'], ['1', name, mobileNum]])
            with pd.ExcelWriter(xlsFile, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, header=False, startrow=0)

            with open(msgFile, "w", encoding="utf-8") as msgWriter:
                msgWriter.write(message)

            errorType = ''
            errorMsg = ''
            # Call whatsapp messaging for the given mobile number
            failVal = call_whatsapp_for_one(xlsFile, msgFile, '', 'A')
            if failVal:
                for errorType in failVal: 
                    errorMsg = failVal[errorType]
                logging.info('Failure after processing [%s] [%s]', errorType, errorMsg)

            msgSentTime = datetime.now()

            # Build the enrolment row
            enrolDict = {
                'name': name, 
                'mobileNum': mobileNum, 
                'emailId': emailId, 
                'package': package, 
                'birthDay': birthDay, 
                'memStartDate': memStartDate, 
                'notes': notes, 
                'calendarId': calendarId, 
                'eventId': eventId, 
                'meetLink': meetLink, 
                'summary': summary, 
                'sesStartDateTime': sesStartDateTime, 
                'message': message, 
                'formSbmTime': formSbmTime, 
                'totPay': totPay, 
                'remPay': totPay, 
                'frqPay': frqPay, 
                'numPay': numPay, 
                'totalSessions': totalSessions, 
                'completedSessions': 0, 
                'sessionDuration': sessionDuration, 
                'frqSes': frqSes, 
                'msgSentTime': msgSentTime, 
                'prcFlg': ''
            }
            
            # check if the message was sent successfully
            if errorType != '':
                logging.error('Errors while processing enrolment for %s. Check logs', mobileNum)
                errorDict = {
                    'process': 'enrolment', 
                    'processDateTime': formSbmTime, 
                    'name': name, 
                    'mobileNum': mobileNum, 
                    'errorType': errorType, 
                    'errorMsg': errorMsg, 
                    'errorCreTime': msgSentTime
                }
                whatsappDB.errorInsert(errorDict); 
                enrolDict['prcFlg'] = 'N'
                if existingRec == '':
                    logging.info("No enrolment record exists. So sending a failure record for : [%s] [%s] [%s]", name, mobileNum)
                    whatsappDB.enrolInsert(enrolDict)
                else:
                    logging.info("Enrolment Record exists. So updating the failure record for : [%s] [%s] [%s]", name, mobileNum, formSbmTime)
                    whatsappDB.enrolUpdate(enrolDict)
                result.append([name, mobileNum, recdFormSbmTimeStr, 'N', errorMsg])
            else:
                logging.info('Completed Successfully for %s', mobileNum)
                enrolDict['prcFlg'] = 'Y'
                if existingRec == '':
                    logging.info("No enrolment record exists. So inserting a success record for : [%s] [%s] [%s]", name, mobileNum, formSbmTime)
                    whatsappDB.enrolInsert(enrolDict)
                else:
                    logging.info("Enrolment Record exists. So updating as successfully processed : [%s] [%s] [%s]", name, mobileNum, formSbmTime)
                    whatsappDB.enrolUpdate(enrolDict)
                result.append([name, mobileNum, recdFormSbmTimeStr, 'Y', 'success'])
        
        logging.info('Calling function removeProcessedEnrolments')
        body = {
            'values': result
        }
        updResp = sheet.values().update(spreadsheetId=config.ENROLMENT_RES_ID, range=config.ENROLMENT_RES_RANGE,
            valueInputOption='USER_ENTERED', body=body).execute()
        logging.info('%d cells updated.', updResp.get('updatedCells'))
        
        # Finally, remove these prosessed entried from request sheet in Google drive.
        request = {"function": "removeProcessedEnrolments"}
        scriptService = build('script', 'v1', credentials=creds)
        response = scriptService.scripts().run(scriptId=config.SCRIPT_ID, body=request).execute()
        logging.info('Returned from script call with %s', response)
        if 'error' in response:
            error = response['error']['details'][0]
            logging.error("Script error message: %s", error['errorMessage'])
            whatsappDB.close('N')
            return 1
        else:
            logging.info("Script error message: %s", response['response'].get('result'))
            whatsappDB.close('Y')
            return 0
        logging.info('Exiting processEnrolments')

    def processPayments(service):
        logging.info('Entering processPayments')
        # Call the Sheets API
        sheet = service.spreadsheets()

        result = sheet.values().get(spreadsheetId=config.PAYMENT_REQ_ID, range=config.PAYMENT_REQ_RANGE).execute()
        values = result.get('values', [])

        if not values:
            logging.error('No data found.')
            return 1
        
        # If any data is found in response spreadsheet, it means google apps script still hasn't processed the previous response. Stop until that's done
        result = sheet.values().get(spreadsheetId=config.PAYMENT_RES_ID, range=config.PAYMENT_RES_RANGE).execute()
        respValues = result.get('values', [])

        if respValues:
            logging.error('Google apps script still has not processed the previous response. Stop until that is done.')
            return 1

        logging.info("values: [%s]", values)
        xlsFile = "whatsapp.xlsx"
        msgFile = "whatsapp.txt"
        
        whatsappDB = WhatsappDB()
        result = []
        
        for curColSet in values:
            # Check if this record is already processed
            name = curColSet[0]
            mobileNum = curColSet[1]
            payAmt = int(curColSet[2])
            payDate = (datetime.fromisoformat(curColSet[3][:-1]+'000+00:00')).astimezone()
            text = curColSet[4]
            recdFormSbmTimeStr = curColSet[5]
            formSbmTime = (datetime.fromisoformat(recdFormSbmTimeStr[:-1]+'000+00:00')).astimezone()
            
            # Check if the person is enrolled
            enrolRec = whatsappDB.enrolSelect(mobileNum)
            if enrolRec == '':
                logging.error("Enrolment record itself does not exist: [%s] [%s] [%s]", name, mobileNum, formSbmTime)
                result.append([name, mobileNum, recdFormSbmTimeStr, 'N', name + ' is not yet enrolled'])
                continue
                
            logging.info(enrolRec)
            
            existingRec = whatsappDB.paymentSelect(mobileNum, formSbmTime)
            if existingRec != '' and existingRec['prcFlg'] == 'Y':
                # Already processed successfully. Tell Google apps script to remove it 
                logging.info("Payment record exists and already processed successfully: [%s] [%s] [%s]", name, mobileNum, formSbmTime)
                result.append([name, mobileNum, recdFormSbmTimeStr, 'Y', 'success'])
                continue

            remPay = enrolRec['remPay'] - payAmt
            if remPay != 0:
                text = text + '\n' + 'Note: You have *₹' + str(remPay) + '* outstanding payment.'
            else: 
                text = text + '\n' + 'Note: You have no outstanding payment. Thank you!'

            df = pd.DataFrame([['SL NO', 'NAME', 'CONTACT DETAILS'], ['1', name, mobileNum]])
            with pd.ExcelWriter(xlsFile, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, header=False, startrow=0)
            
            with open(msgFile, "w", encoding="utf-8") as msgWriter:
                msgWriter.write(text)

            errorType = ''
            errorMsg = ''
            # Call whatsapp messaging for the given mobile number
            failVal = call_whatsapp_for_one(xlsFile, msgFile, '', 'A')
            if failVal:
                for errorType in failVal: 
                    errorMsg = failVal[errorType]
                logging.info('Failure after processing [%s] [%s]', errorType, errorMsg)

            msgSentTime = datetime.now()

            # Build the payment row
            paymentDict = {
                'name': name, 
                'mobileNum': mobileNum, 
                'payAmt': payAmt, 
                'payDate': payDate, 
                'text': text, 
                'formSbmTime': formSbmTime, 
                'msgSentTime': msgSentTime, 
                'prcFlg': ''
            }
            
            # check if the message was sent successfully
            if errorType != '':
                logging.error('Errors while processing for %s. Check logs', mobileNum)
                # Build the error record
                errorDict = {
                    'process': 'payment', 
                    'processDateTime': formSbmTime, 
                    'name': name, 
                    'mobileNum': mobileNum, 
                    'errorType': errorType, 
                    'errorMsg': errorMsg, 
                    'errorCreTime': msgSentTime
                }
                whatsappDB.errorInsert(errorDict); 
                paymentDict['prcFlg'] = 'N'
                if existingRec == '':
                    logging.info("No payment record exists. So inserting a failure record for : [%s] [%s] [%s]", name, mobileNum, formSbmTime)
                    whatsappDB.paymentInsert(paymentDict)
                    logging.info('Updating completed payment %s.', mobileNum)
                    whatsappDB.enrolUpdatePayments(mobileNum, payAmt)
                else:
                    logging.info("Payment Record exists. So updating the failure record for : [%s] [%s] [%s]", name, mobileNum, formSbmTime)
                    whatsappDB.paymentUpdate(paymentDict)                    
                result.append([name, mobileNum, recdFormSbmTimeStr, 'N', errorMsg])
            else:
                logging.info('Completed Successfully for %s %d', mobileNum, len(text))
                paymentDict['prcFlg'] = 'Y'
                if existingRec == '':
                    logging.info("No payment record exists. So inserting a success record for : [%s] [%s] [%s]", name, mobileNum, formSbmTime)
                    whatsappDB.paymentInsert(paymentDict)
                    logging.info('Updating completed payments %s.', mobileNum)
                    whatsappDB.enrolUpdatePayments(mobileNum, payAmt)
                else:
                    logging.info("Payment Record exists. So updating the record as successfully processed : [%s] [%s] [%s]", name, mobileNum, formSbmTime)
                    whatsappDB.paymentUpdate(paymentDict)                    
                result.append([name, mobileNum, recdFormSbmTimeStr, 'Y', 'success'])
        
        logging.info('Calling function removeProcessedPayments')
        body = {
            'values': result
        }
        updResp = sheet.values().update(spreadsheetId=config.PAYMENT_RES_ID, range=config.PAYMENT_RES_RANGE,
            valueInputOption='USER_ENTERED', body=body).execute()
        logging.info('%d cells updated.', updResp.get('updatedCells'))
        
        # Finally, remove these prosessed entries from request sheet in Google drive.
        request = {"function": "removeProcessedPayments"}
        scriptService = build('script', 'v1', credentials=creds)
        response = scriptService.scripts().run(scriptId=config.SCRIPT_ID, body=request).execute()
        logging.info('Returned from script call with %s', response)
        if 'error' in response:
            error = response['error']['details'][0]
            logging.error("Script error message: %s", error['errorMessage'])
            whatsappDB.close('N')
            return 1
        else:
            logging.info('script success message: %s', response['response'].get('result'))
            logging.info(response)
            whatsappDB.close('Y')
            return 0
        logging.info('Exiting processPayments')

    # Batch processing starts here
    args = {}
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', config.SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(config.CRED_FILE, config.SCOPES)
            creds = flow.run_local_server(port=0)

        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        if config.PROCESS_TYPE == 'E' or config.PROCESS_TYPE == 'A':
            processEnrolments(service)
        if config.PROCESS_TYPE == 'S' or config.PROCESS_TYPE == 'A':
            processSessions(service)
        if config.PROCESS_TYPE == 'P' or config.PROCESS_TYPE == 'A':
            processPayments(service)

    except HttpError as err:
        logging.error(err)
    except Exception as exp:
        logging.error(exp)

def interactive_process():
    win = tk.Tk()

    #Set the geometry
    win.geometry("640x320")
    win.title("WhatsApp Bulk Messenger")
    win.resizable(True, True)
    win.configure(bg='floralwhite')
    style = ttk.Style()

    # layout on the window
    win.columnconfigure(0, weight=1)
    win.rowconfigure(0, weight=1)
    win.rowconfigure(1, weight=1)

    inpFrame = ttk.LabelFrame(win, text='Give your inputs')
    inpFrame.grid(row=0, column=0)
    inpFrame['borderwidth'] = 5
    inpFrame['relief'] = 'raised'

    rspFrame = ttk.LabelFrame(win, text='Status of messaging')
    rspFrame.grid(row=1, column=0)

    rowNum = 0
    colNum = 0

    # Create label, entry and file dialog dutton for excel file
    xlsLabel=ttk.Label(inpFrame, text="Excel Contacts file")
    xlsLabel.grid(row=rowNum, column=colNum, padx=5, pady=5, sticky=tk.W)
    colNum += 1
    xlsEntry = ttk.Entry(inpFrame, width=60, textvariable=tk.StringVar(), )
    xlsEntry.grid(row=rowNum, column=colNum, padx=5, pady=5, sticky=tk.E)
    xlsTip = "Enter the excel file which has all the contacts to be messaged"
    ToolTip(xlsEntry, xlsTip)
    colNum += 1
    xlsButton = ttk.Button(inpFrame, text='...', width=3, command=lambda:select_file('Open excel file','*.xlsx',xlsEntry))
    xlsButton.grid(row=rowNum, column=colNum, padx=5, pady=5, sticky=tk.W)

    rowNum += 1
    colNum = 0
    # Create label, entry and file dialog dutton for message file
    txtLabel=ttk.Label(inpFrame, text="Whatsapp Message file")
    txtLabel.grid(row=rowNum, column=colNum, padx=5, pady=5, sticky=tk.W)
    colNum += 1
    txtEntry = ttk.Entry(inpFrame, width=60, textvariable=tk.StringVar())
    txtEntry.grid(row=rowNum, column=colNum, padx=5, pady=5, sticky=tk.E)
    txtTip = "Enter the text file(s) with message(s) to be sent. The number of files should match the number of sets of contacts in the excel file"
    ToolTip(txtEntry, txtTip)
    colNum += 1
    txtButton = ttk.Button(inpFrame, text='...', width=3, command=lambda:select_file('Open text file','*.txt',txtEntry))
    txtButton.grid(row=rowNum, column=colNum, padx=5, pady=5, sticky=tk.W)

    rowNum += 1
    colNum = 0
    # Create label, entry and file dialog dutton for image file
    imgLabel=ttk.Label(inpFrame, text="Whatsapp Image file")
    imgLabel.grid(row=rowNum, column=colNum, padx=5, pady=5, sticky=tk.W)
    colNum += 1
    imgEntry = ttk.Entry(inpFrame, width=60, textvariable=tk.StringVar())
    imgEntry.grid(row=rowNum, column=colNum, padx=5, pady=5, sticky=tk.E)
    imgTip = "Enter the image file(s) to be sent with message(s). The number of images should match the number of sets of contacts in the excel file"
    ToolTip(imgEntry, imgTip)
    colNum += 1
    imgButton = ttk.Button(inpFrame, text='...', width=3, command=lambda:select_file('Open image file','*.jpg *.png *.gif *.jpeg',imgEntry))
    imgButton.grid(row=rowNum, column=colNum, padx=5, pady=5, sticky=tk.W)

    rowNum += 1
    colNum = 2
    # Submit and clear buttons
    subButton = ttk.Button(inpFrame, text='Submit', command=lambda:call_whatsapp(xlsEntry, txtEntry, imgEntry, resText))
    subButton.grid(row=rowNum, column=colNum, padx=5, pady=5, sticky=tk.W)
    colNum -= 1
    clrButton = ttk.Button(inpFrame, text='Clear', command=lambda:clear_files(xlsEntry, txtEntry, imgEntry))
    clrButton.grid(row=rowNum, column=colNum, padx=5, pady=5, sticky=tk.E)

    # Now, layout of the response frame
    rowNum = 0
    colNum = 0
    # Create text box to display response
    resText=tk.Text(rspFrame, width=75, height=1)
    resText.grid(row=rowNum, column=colNum, columnspan=3, padx=5, pady=5, sticky=tk.W)
    resText.insert('end', 'Yet to start')

    win.mainloop()


parser = argparse.ArgumentParser()
parser.add_argument('mode', help='(O)nline or (B)atch', type=str)
parsedArgs = parser.parse_args()
args = vars(parsedArgs)
if args.get('mode') == 'O':
    interactive_process()
else:
    batch_process()