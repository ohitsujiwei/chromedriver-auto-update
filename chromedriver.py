import os
import zipfile
import subprocess

import requests
import win32com.client

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver

CHROME_BROWSER_EXE_PATH = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

CHROME_DRIVER_BASE_URL = "https://googlechromelabs.github.io/chrome-for-testing"
CHROME_DRIVER_DOWNLOAD_URL = f"https://storage.googleapis.com/chrome-for-testing-public"
CHROME_DRIVER_DOWNLOAD_OS_VER = "win32"
CHROME_DRIVER_DOWNLOAD_ZIP_NAME = f"chromedriver-{CHROME_DRIVER_DOWNLOAD_OS_VER}.zip"

CHROME_DRIVER_FOLDER = os.path.dirname(__file__)

CHROME_DRIVER_EXE_PATH = f"{CHROME_DRIVER_FOLDER}\\chromedriver-{CHROME_DRIVER_DOWNLOAD_OS_VER}\\chromedriver.exe"
CHROME_DRIVER_ZIP_PATH = f"{CHROME_DRIVER_FOLDER}\\{CHROME_DRIVER_DOWNLOAD_ZIP_NAME}"


class ChromeDriver(object):
    def __init__(self, check_version_startup=True):
        super().__init__()
        if check_version_startup:
            self.update_driver()

    def get_driver_version(self):
        if not os.path.isfile(CHROME_DRIVER_EXE_PATH):
            raise FileNotFoundError(f"[{CHROME_DRIVER_EXE_PATH}] is not found.")
        version = subprocess.check_output(f"{CHROME_DRIVER_EXE_PATH} --version").decode().split()[1]
        return version

    def get_browser_version(self):
        if not os.path.isfile(CHROME_BROWSER_EXE_PATH):
            raise FileNotFoundError(f"[{CHROME_BROWSER_EXE_PATH}] is not found.")
        wincom_object = win32com.client.Dispatch("Scripting.FileSystemObject")
        version: str = wincom_object.GetFileVersion(CHROME_BROWSER_EXE_PATH)
        return version

    def update_driver(self):
        driver_version = self.get_driver_version()
        browser_version = self.get_browser_version()
        driver_version_major = driver_version.split(".")[0]
        browser_version_major = browser_version.split(".")[0]

        if driver_version_major != browser_version_major:
            get_available_version_url = f"{CHROME_DRIVER_BASE_URL}/LATEST_RELEASE_{browser_version_major}"
            available_version = requests.get(get_available_version_url, stream=True, timeout=300).content.decode()

            download_api = f"{CHROME_DRIVER_DOWNLOAD_URL}/{available_version}/{CHROME_DRIVER_DOWNLOAD_OS_VER}/{CHROME_DRIVER_DOWNLOAD_ZIP_NAME}"
            response = requests.get(download_api, stream=True, timeout=300)

            if response.status_code == 200:
                with open(CHROME_DRIVER_ZIP_PATH, "wb") as file:
                    file.write(response.content)
            else:
                raise Exception(f"[{response.status_code}] Download chrome driver failed.")

            with zipfile.ZipFile(CHROME_DRIVER_ZIP_PATH, "r") as zip_ref:
                zip_ref.extractall(CHROME_DRIVER_FOLDER)

    def create_driver(self):
        options = Options()
        options.add_argument("--disable-notifications")
        options.add_argument("--disable-plugins")
        options.add_argument("blink-settings=imagesEnabled=false")
        options.add_argument("--incognito")
        options.add_argument("--headless")
        # options.add_experimental_option("detach", True)

        service = Service(CHROME_DRIVER_EXE_PATH)

        driver: WebDriver = webdriver.Chrome(service=service, options=options)
        driver.set_page_load_timeout(300)
        return driver
