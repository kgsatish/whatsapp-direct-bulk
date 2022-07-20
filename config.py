import getpass
import platform
import logging

# Change this to 'N' to send messages, 'Y' to test without sending
TEST_MODE = 'N'
os_name = platform.system()
username = getpass.getuser()

if os_name == 'Linux':
    CHROME_PROFILE_PATH = f'--user-data-dir=/home/{username}/.config/google-chrome/WhatsApp'
    CHROME_DRIVER_PATH=r'/usr/local/bin/chromedriver'
elif os_name == 'Windows':
    CHROME_PROFILE_PATH = fr'--user-data-dir=C:\Users\{username}\AppData\Local\Google\Chrome\User Data\WhatsApp'
    CHROME_DRIVER_PATH=fr'C:\Users\{username}\AppData\Local\Programs\Python\Python310\chromedriver.exe'

INVALID_PHONE = r'Phone number shared via url is invalid.'
WAIT_TIME = 60

logging.basicConfig(
    filename='rajani.log', 
    encoding='utf-8',
    format='%(asctime)s %(levelname)-8s [%(filename)s:%(lineno)d]: %(message)s',
    level=logging.INFO,
    datefmt='%Y-%m-%d %H:%M:%S')
