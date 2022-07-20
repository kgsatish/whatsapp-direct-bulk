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


class WhatsappBulkMessage(object):
    """
    A class that encapsulates Whatsapp Message automation function and attributes
    """

    def __init__(self, **kwargs):
        logging.info('Entered init')
        self.excel_file = kwargs.get('excel_file')
        if not exists(self.excel_file):
            logging.error("Excel file [%s] does not exist", self.excel_file)
        self.msg_file = kwargs.get('msg_file')
        if not exists(self.msg_file):
            logging.error("Message file [%s] does not exist", self.msg_file)
        self.msg_ind = kwargs.get('msg_ind')
        img_file = kwargs.get('img_file')
        self.img_path = None
        if img_file:
            self.img_path = os.path.join(os.getcwd(), img_file)
            if not exists(self.img_path):
                logging.error("Image file [%s] does not exist", self.img_path)
        self.excel_data = None
        self.driver = None
        self.driver_wait = None
        self.msg_text = None
        logging.info('Exited init')

    def perform_task(self):
        logging.info('Entered perform task')
        try:
            self.initialize()
            self.read()
            self.process()
        finally:
            self.close()
            logging.info('Exited perform task')

    def initialize(self):
        # Load the chrome driver
        options = webdriver.ChromeOptions()
        options.add_argument(config.CHROME_PROFILE_PATH)
        options.add_argument('--disable-web-security')
        options.add_argument('--disable-gpu')
        options.add_argument('--disable-dev-shm-usage')
        options.add_experimental_option("excludeSwitches", ["enable-logging"])

        self.driver = webdriver.Chrome(service=Service(config.CHROME_DRIVER_PATH), options=options)
        self.driver_wait = WebDriverWait(self.driver, config.WAIT_TIME)

    def read(self):
        # Read data from excel
        # noinspection PyArgumentList
        self.excel_data = pd.read_excel(self.excel_file, usecols=['SL NO', 'NAME', 'CONTACT DETAILS'],
                                        engine='openpyxl')
        # Read message from text file
        fp = open(self.msg_file, "r", encoding="utf-8")
        self.msg_text = fp.read()
        fp.close()

    def process(self):
        # Iterate excel rows till to finish
        msg_ind = 0
        sl_no = ''                
        name = ''
        contact_number = ''
        logging.info('Started processing...')
        try:
            for idx in self.excel_data.index:
                sl_no = str(self.excel_data['SL NO'][idx])
                name = str(self.excel_data['NAME'][idx])
                contact_number = str(self.excel_data['CONTACT DETAILS'][idx])[0:12]

                logging.info('SL NO: [%s], msg_ind: [%d], passed msg_ind: [%s], idx: [%d]', sl_no, msg_ind,
                              self.msg_ind, idx)

                if self.msg_ind != 'A':
                    logging.debug('After SL NO check - success')
                    if sl_no == '1':
                        msg_ind += 1
                        logging.debug('msg_ind after increase: [%d]', msg_ind)

                    # Get the right set of messages to work on
                    if msg_ind != int(self.msg_ind):
                        logging.info('Ignored: [%s]. [%s]', sl_no, name)
                        continue

                logging.info('Considered: [%s]. [%s]', sl_no, name)
                # Assign customized message
                message = 'Hi ' + name + ',\n\n' + self.msg_text

                if len(contact_number) == 10:
                    contact_number = '91' + contact_number
                    
                if (not contact_number.startswith('91')) or (contact_number.startswith('91') and not len(contact_number) == 12):
                        logging.error('Invalid mobile number: [%s]', contact_number)
                        continue
                
                logging.info('Chrome driver: [%s]', config.CHROME_PROFILE_PATH)
                qry = {'phone': contact_number, 'text': message}
                url = "https://web.whatsapp.com/send/?{}".format(urllib.parse.urlencode(qry))

                if config.TEST_MODE == 'Y':
                    logging.info('url: %s', url)
                    continue

                logging.debug('before get')
                self.driver.get(url)
                logging.debug('after get')
                if idx == 0:
                    time.sleep(20)
                else:
                    time.sleep(9)
                    
                logging.debug('after sleep')

                # Check if the whatsapp number is valid
                #valid = self.check_number(contact_number)
                #if valid == 'N':
                #    continue

                if self.img_path:
                    attachment_button = self.driver_wait.until(
                        lambda driver: driver.find_element(By.XPATH, '//div[@title="Attach"]'))
                    attachment_button.click()
                    image_button = self.driver_wait.until(lambda driver: driver.find_element(
                        By.XPATH, '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'))
                    image_button.send_keys(self.img_path)
                    time.sleep(3)
                    send_button = self.driver_wait.until(lambda driver: driver.find_element(
                        By.XPATH, '//div[@role="button" and @aria-label="Send"]'))
                    action = ActionChains(self.driver)
                    action.move_to_element(send_button).click().perform()
                    time.sleep(3)
                else:
                    message_box = self.driver_wait.until(lambda driver: driver.find_element(
                        By.XPATH, '//div[@title="Type a message" and @role="textbox"]'))
                    action = ActionChains(self.driver)
                    action.move_to_element(message_box).click()
                    action.send_keys(Keys.ENTER)
                    action.perform()
                    time.sleep(3)

                logging.info('Done for %s - %s', name, contact_number)
        except StaleElementReferenceException:
            logging.error('Stale element exception for SL NO: [%s], name: [%s], contact_number: [%s], idx: [%d]', sl_no, name, contact_number, idx)
            pass
        except Exception as exp:
            logging.error('Following exception for SL NO: [%s], name: [%s], contact_number: [%s], idx: [%d]', sl_no, name, contact_number, idx)
            logging.error(exp)
            pass

    def check_number(self, contact_number):
        try:
            elements = self.driver.find_elements(By.XPATH, '//div[starts-with(@class, "_")]')
            for element in elements:
                if element.text == config.INVALID_PHONE:
                    logging.error('Phone number not in whatsapp: %s', contact_number)
                    ok_button = self.driver.find_element(By.XPATH, '//div[@role="button"]')
                    action = ActionChains(self.driver)
                    action.move_to_element(ok_button).click().perform()
                    return 'N'
        except StaleElementReferenceException:
            logging.error("ignoring stale element for: ")
            attributes = self.driver.execute_script(
                'var items = {}; for (index = 0; index < arguments[0].attributes.length; ++index) { items['
                'arguments[0].attributes[index].name] = arguments[0].attributes[index].value }; return items;',
                element)
            logging.error(attributes)
            logging.error(element.tag_name)
            pass
        except Exception as exp:
            logging.error("Exception occurred while checking number for: ")
            logging.error(exp)
            pass

        logging.info('Phone number found in whatsapp: %s', contact_number)
        return 'Y'

    def close(self):
        # Close Chrome browser
        self.driver.quit()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Whatsapp Bulk Messaging with optional attachment (document, image, audio, video)',
        epilog="Make sure the above file to be attached is in the same folder as this python script")
    parser.add_argument('excel_file', help='Excel file with SL NO, NAME, CONTACT DETAILS columns. '
                                           'Phone num should be <country code><10-digit mobile number>', type=str)
    parser.add_argument('msg_file', help='Text file containing the message text to be sent', type=str)
    parser.add_argument('msg_ind', help='1 for first set, 2 for second set, 3 for third set, and so on, A for all',
                        type=str)
    parser.add_argument('--doc', help='Document to be attached', type=str, dest='doc_file')
    parser.add_argument('--img', help='Image to be attached', type=str, dest='img_file')
    parser.add_argument('--vid', help='Video to be attached', type=str, dest='vid_file')
    parser.add_argument('--aud', help='Audio to be attached', type=str, dest='aud_file')
    parsed_args = parser.parse_args()
    args = vars(parsed_args)
    whatsapp = WhatsappBulkMessage(**args)
    whatsapp.perform_task()
