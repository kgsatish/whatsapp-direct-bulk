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


class WhatsappBulkMessage(object):
    """
    A class that encapsulates Whatsapp Message automation function and attributes
    """

    def __init__(self, **kwargs):
        self.csv_name = kwargs.get('csv_name')
        self.img_path = kwargs.get('img_path')
        self.excel_data = None
        self.driver = None
        self.driver_wait = None

    def perform_task(self):
        try:
            self.initialize()
            self.read()
            self.process()
        finally:
            self.close()

    def initialize(self):
        # Load the chrome driver
        options = webdriver.ChromeOptions()
        options.add_argument(config.CHROME_PROFILE_PATH)
        options.add_argument('--disable-web-security')
        options.add_argument('--disable-gpu')
        options.add_argument('--disable-dev-shm-usage')
        options.add_experimental_option("excludeSwitches", ["enable-logging"]) 

        self.driver = webdriver.Chrome(service=Service(config.CHROME_DRIVER_PATH), options=options)
        self.driver_wait = WebDriverWait(self.driver, 60)

    def read(self):
        # Read data from excel
        self.excel_data = pd.read_csv(self.csv_name, encoding="utf-8")

    def process(self):
        # Iterate excel rows till to finish
        try:
            for ind in self.excel_data.index:
                # Assign customized message
                message = 'Hi ' + self.excel_data['Name'][ind] + ',\n\n' + self.excel_data['Message'][ind]

                # Send contact number in search box
                contact_number = str(self.excel_data['Contact'][ind])[0:12]

                if not len(contact_number) == 12 and not contact_number.startswith('91'):
                    print('Contact number has be <91><10-digit mobile number>: ' + str(contact_number))
                    continue
                
                qry = {'phone': contact_number, 'text': message}
                url = "https://web.whatsapp.com/send/?{}".format(urllib.parse.urlencode(qry))
                print (url)
                self.driver.get(url)
                time.sleep(7)

                # Check if the whatsapp number is valid
                valid = self.check_number(contact_number)
                if valid == 'N':
                    continue
                
                if self.img_path is not None:
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
                else:
                    message_box = self.driver_wait.until(lambda driver: driver.find_element(
                        By.XPATH, '//div[@title="Type a message" and @role="textbox"]'))
                    action = ActionChains(self.driver)
                    action.move_to_element(active).click()
                    action.send_keys(Keys.ENTER)
                    action.perform()
                    time.sleep(3)

                print('Done for ' + self.excel_data['Name'][ind] + '-' + str(contact_number))
        except Exception as exp:
            print(exp)

    def check_number(self, contact_number):
        try:
            elements = self.driver.find_elements(By.XPATH, '//div[starts-with(@class, "_")]')
            for element in elements:
                if element.text == config.INVALID_PHONE:
                    print ('Phone number not in whatsapp: ', contact_number)
                    ok_button = self.driver.find_element(By.XPATH, '//div[@role="button"]')
                    action = ActionChains(self.driver)
                    action.move_to_element(ok_button).click().perform()
                    return 'N'
        except StaleElementReferenceException:
            pass
        print ('Phone number found in whatsapp: ', contact_number)
        return 'Y'

    def close(self):
        # Close Chrome browser
        self.driver.quit()

if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description='Whatsapp Bulk Messaging with optional attachment (document, image, audio, video)')
    parser.add_argument('csv_name', help='Full path of CSV file name with Name, Contact and Message columns. '
                                         'Contact should be <country code><10-digit mobile number>', type=str)
    parser.add_argument('--doc', help='Full path of the document to be attached', type=str, dest='doc_path')
    parser.add_argument('--img', help='Full path of the image to be attached', type=str, dest='img_path')
    parser.add_argument('--vid', help='Full path of the video to be attached', type=str, dest='vid_path')
    parser.add_argument('--aud', help='Full path of the audio to be attached', type=str, dest='aud_path')
    parsed_args = parser.parse_args()
    args = vars(parsed_args)
    whatsapp = WhatsappBulkMessage(**args)
    whatsapp.perform_task()
