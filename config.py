import getpass
import platform
import logging
import pymysql.cursors

dbConnect = pymysql.connect(host='192.168.29.122',
                        user='rajani',
                        password='SuHaan7112',
                        database='whatsapp',
                        cursorclass=pymysql.cursors.DictCursor)
                        
# Change this to 'N' to send messages, 'Y' to test without sending
TEST_MODE = 'N'
os_name = platform.system()
username = getpass.getuser()

if os_name == 'Linux':
    CHROME_PROFILE_PATH = f'--user-data-dir=/home/{username}/.config/google-chrome/WhatsApp'
    CHROME_DRIVER_PATH=r'/usr/local/bin/chromedriver'
elif os_name == 'Windows':
    CHROME_PROFILE_PATH = fr'--user-data-dir=C:\Users\{username}\AppData\Local\Google\Chrome\User Data\WhatsApp'
    CHROME_DRIVER_PATH=fr'C:\Users\{username}\AppData\Local\chromedriver.exe'

INVALID_PHONE = r'Phone number shared via url is invalid.'
WAIT_TIME = 120

CRED_FILE = r'credentials-biohub.json'
# Scope for reading the google spreadsheet
#SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/documents']
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/calendar']

# The ID and range of google spreadsheets with whatsapp message data.
SESSION_REQ_ID = '1pGBoBmOzQAezLFoGjTIcQn96mMPDZY8ke36RJuuw9Lg'
SESSION_RES_ID = '1x5HROWby06BLI8_trvS9eG5ILkd_mZxSYi9xBOaQoM4'
ENROLMENT_REQ_ID = '1pQqmZHYugvCQCC6Ww-OI7EL6JNTax4ezXleh17MvhNo'
ENROLMENT_RES_ID = '1o_Kr2MIe5HJmpSDWtycGUzSx8GxYMMdK36ky7ozLmMM'
PAYMENT_REQ_ID = '1GLMwELG536HlUvHHShsVf4wF58G1w54NhCUJ79SanE0'
PAYMENT_RES_ID = '1HYUS-V0jzc-_W0w4CAiks9kulfIf7QNbBOGYcCxQP_k'

SESSION_REQ_RANGE = 'Sheet1!A2:H'
SESSION_RES_RANGE = 'Sheet1!A2:E'
ENROLMENT_REQ_RANGE = 'Sheet1!A2:T'
ENROLMENT_RES_RANGE = 'Sheet1!A2:E'
PAYMENT_REQ_RANGE = 'Sheet1!A2:F'
PAYMENT_RES_RANGE = 'Sheet1!A2:E'

#Exceptiion list
FILE_NOT_FOUND = '01'
UNCAUGHT_EXCEPTION = '02'
INVALID_PHONE_NUM = '03'
PROCESS_EXCEPTION = '04'
NOT_IN_WHATSAPP = '05'
STALE_ELEMENT = '06'

VALID_PHONES = [
    {'ISD': '91', 'length': 12},
    {'ISD': '33', 'length': 11},
]

SEND_MSG = 'N'
PROCESS_TYPE = 'E'

SCRIPT_ID = 'AKfycbxsDXRGck-9xD6rmNM5MngRSA_sUe3bcmCrrtJ69K1VIL2e91KhExHKn2SZChs8JugZ'

logging.basicConfig(
    filename='app.log', 
    encoding='utf-8',
    format='%(asctime)s %(levelname)-8s [%(filename)s:%(lineno)d]: %(message)s',
    level=logging.INFO,
    datefmt='%Y-%m-%d %H:%M:%S')

IMAGE_TEXT = 'join today ☝️'