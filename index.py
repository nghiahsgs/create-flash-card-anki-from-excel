from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec

from time import sleep
import shutil
import openpyxl
class Ttsmp3():
    def __init__(self):
        chrome_options = webdriver.ChromeOptions()
        # prefs = {
        #     "profile.managed_default_content_settings.images": 2
        # }
        # chrome_options.add_experimental_option("prefs", prefs)
        self.driver = webdriver.Chrome(
            './chromedriver', chrome_options=chrome_options)
        
        self.default_folder_download_chrome = r'C:\Users\AnDongPc\Downloads'
    
    def set_sentence_eng(self, sentence_eng):
        self.sentence_eng = sentence_eng

    def set_sentence_index(self, sentence_index):
        self.sentence_index = sentence_index

    def downMp3FromText(self):
        self.driver.get('http://ttsmp3.com')
        sleep(5)
        text_area=self.driver.find_element_by_css_selector('#voicetext')
        text_area.send_keys(self.sentence_eng)
        sleep(5)

        btn_download = self.driver.find_element_by_css_selector(
            '#downloadenbutton')
        btn_download.click()
        sleep(10)
        #input('download')
        self.driver.get('chrome://downloads/')
        #css selector to get downlead file name
        #document.querySelector('downloads-manager').shadowRoot.querySelector('downloads-item').shadowRoot.querySelector('#file-link').innerText
        file_name=self.driver.execute_script('return document.querySelector("downloads-manager").shadowRoot.querySelector("downloads-item").shadowRoot.querySelector("#file-link").innerText')
        file_name = '%s/%s' % (self.default_folder_download_chrome, file_name)

        shutil.move(file_name,'audio/%s.mp3'%(self.sentence_index))

class ReadAndWriteExcel():
    def __init__(self, file_name):
        self.file_name=file_name

    def get_value_excel(self, cellname):
        wb = openpyxl.load_workbook(self.file_name)
        Sheet1 = wb['Sheet1']
        wb.close()
        return Sheet1[cellname].value

    def update_value_excel(self, cellname, value):
        wb = openpyxl.load_workbook(self.file_name)
        Sheet1 = wb['Sheet1']
        Sheet1[cellname].value = value
        wb.close()
        wb.save(self.file_name)

class Anki():
    def __init__(self):
        pass

    pyautogui.mouseDown(200, 200)
    pyautogui.mouseUp(200, 200)
    sleep(3)
    pyautogui.keyDown('a')
    pyautogui.keyUp('a')

#file_excel=ReadAndWriteExcel('list_words.xlsx')
#print(file_excel.get_value_excel('A1'))
    
# bot_mp3=Ttsmp3()
# bot_mp3.set_sentence_eng('i love you')
# bot_mp3.set_sentence_index(0)
# bot_mp3.downMp3FromText()
