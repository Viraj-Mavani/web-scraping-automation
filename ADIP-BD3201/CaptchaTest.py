import sys
import requests
import base64
import os
from typing import Self
import pytesseract
import imageio.v3 as iio
import cv2
from sqlite3 import Error
from bs4 import BeautifulSoup
import subprocess
import time
import chromedriver_autoinstaller
# import requests
# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.common.exceptions import NoSuchElementException,StaleElementReferenceException,TimeoutException
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.common.keys import Keys
# # from selenium.webdriver.common.action_chains import ActionChains
# from selenium.webdriver.support import expected_conditions as EC


# BasePath = os.getcwd()
BasePath = 'D:\\Projects\\CedarPython\\ADIP-BD3201'
# converted_image_path = BasePath + '\\Log\\ADIP-BD3201-Converted_img.png'
# og_image_path = BasePath + '\\Log\\nc_cap.dib'
og_image_path = BasePath + '\\Log\\Discord_Captcha.png'


def img2txt():
    # print("Resampling the Image")
    # image = iio.imread(og_image_path)
    # iio.imwrite(converted_image_path, image)

    img = cv2.imread(og_image_path)                                                 # import image data
    gray_img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)                                # convert to grayscale
    _, threshold_img = cv2.threshold(gray_img, 190, 255, cv2.THRESH_BINARY_INV)     # threshold image
    blurred_img = cv2.GaussianBlur(threshold_img, (5, 5), 0)                        # BLUR image
    # img to txt
    return pytesseract.image_to_string(blurred_img, config="-c tessedit_char_whitelist=0123456789 --psm 10")


def img_dl(image_path, img_url):
    response = requests.get(img_url, timeout=200, verify=False)
    # img_soup = BeautifulSoup(response.content, 'html.parser')
    # img_tag = img_soup.find('img', src=img_url)
    with open(image_path, 'wb') as handler:
        handler.write(response.content)


def captcha(obj):
    p_hash = obj.select_one('input[name="p_hash"]').get('value')
    captcha_img_url = 'https://app.roc.gov.bd/psp/nc_cap?p_hash=' + p_hash
    img_dl(og_image_path, captcha_img_url)
    captcha_text = img2txt()
    return p_hash, captcha_text.strip()


def restart_script():
    python = sys.executable
    subprocess.call([python] + sys.argv)


def request(payload):
    retry_attempts = 5
    retry_delay = 2
    while True:
        try:
            Retry = 1
            while Retry <= retry_attempts:
                try:
                    obj = requests.post(
                        Base_URL, data=payload, timeout=3600, verify=False)
                    break
                except Exception as e:
                    raise (e)
                    print(f"Error occurred in status for Home URL")
                    delay = retry_delay * (2 ** Retry)
                    print(f'Retrying in {delay} seconds...{Retry}')
                    time.sleep(delay)
                    Retry += 1
                    continue
            else:
                print('exception')
                os._exit(1)
            soup = BeautifulSoup(obj.content, 'html.parser')
            form_element = soup.find('form', action='nc_search')
            error_element = form_element.find(
                'b', string=' Incorrect Code- Please try again.')

            if error_element is None:
                soup.decompose()
                break
            else:
                p_hash = soup.select_one('input[name="p_hash"]').get('value')
                captcha_img_url = 'https://app.roc.gov.bd/psp/nc_cap?p_hash=' + p_hash
                img_dl(og_image_path, captcha_img_url)
                captcha_text = img2txt()
                captcha_value = captcha_text.strip()
                payload['p_captcha'] = captcha_value
                payload['p_hash'] = p_hash
                continue
        except:
            print('exception')
            time.sleep(200)
    return form_element


if __name__ == "__main__":
    try:
        Base_URL = 'https://app.roc.gov.bd/psp/nc_search'

        f_req = requests.get(Base_URL, timeout=200, verify=False)
        f_soup = BeautifulSoup(f_req.content, 'html.parser')
        p_hash, captcha_value = captcha(f_soup)
        
        if int(captcha_value)<60000:
            print(f'{captcha_value} --> Restarting!!')
            restart_script()
        else:
            print(f'{captcha_value}')

        letter = 'a'

        fields = {
            'entity_type': '1',
            'search_text': letter,
            'CB': '1',
            'p_captcha': captcha_value,
            'p_hash': p_hash,
            'result_type': '0',
            'p_entry_mode': '3',
            'page_no': '1'
        }
        form_element = request(fields)

        table2 = form_element.find('table', id='AutoNumber2')

        index_table = table2.find('table', id='AutoNumber3')
        try:
            page_element = index_table.find('font')
            page_text = page_element.get_text(strip=True)
            total_page = int(page_text.split()[-1])
        except:
            total_page = 1

        print(total_page)

        sl_element = table2.find('b', string='SL.')
        data_table_rows = sl_element.find_parents(
            'td')[1].find('table').find_all('tr')[1:]

        if total_page < 2 and len(data_table_rows) == 0:
            print(f'No Data for {letter}')
            success = True
        else:
            success = True
            if success:
                print('Complete ' + letter + ' for page: 1')
            else:
                print('Failed!! ' + letter + ' for page: 1')
            for index in range(2, total_page+1):
                p_hash, captcha_value = captcha(form_element)
                fields['search_text'] = letter
                fields['p_captcha'] = captcha_value
                fields['p_hash'] = p_hash
                fields['page_no'] = index
                form_element = request(fields)
                table2 = form_element.find('table', id='AutoNumber2')
                sl_element = table2.find('b', string='SL.')
                data_table_rows = sl_element.find_parents(
                    'td')[1].find('table').find_all('tr')[1:]
                success = True
                if success:
                    print(f'Complete {letter} for page: {str(index)}')
                else:
                    print(f'Failed!! {letter} for page: {str(index)}')

        # text = int(img2txt())
        # print(text)

        # if text<30000:
        #     restart_script()

    finally:
        print("###END###")
