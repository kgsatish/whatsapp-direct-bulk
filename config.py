import getpass
import platform

os_name = platform.system()
username = getpass.getuser()

if os_name == 'Linux':
    CHROME_PROFILE_PATH = f'--user-data-dir=/home/{username}/.config/google-chrome/WhatsApp'
    CHROME_DRIVER_PATH=r'/usr/local/bin/chromedriver'
elif os_name == 'Windows':
    CHROME_PROFILE_PATH = fr'--user-data-dir=C:\Users\{username}\AppData\Local\Google\Chrome\User Data\WhatsApp'
    CHROME_DRIVER_PATH=fr'C:\Users\{username}\AppData\Local\Programs\Python\Python38\chromedriver.exe'
