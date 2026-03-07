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
import speech_recognition as sr
from pydub import AudioSegment
from pydub.silence import split_on_silence



# BasePath = os.getcwd()
BasePath = 'D:\\Projects\\CedarPython\\ADIP-BD3201'
# converted_image_path = BasePath + '\\Log\\ADIP-BD3201-Converted_img.png'
# og_image_path = BasePath + '\\Log\\nc_cap.dib'
og_image_path = BasePath + '\\Log\\Discord_Captcha.png'
audio_file_path = "C:\\Users\\devpl\\Downloads\\captcha_a.wav"  # Replace with the path to your .wav file


def recognize_audio_without_noise(file_path):
    # Load the audio file
    audio = AudioSegment.from_wav(file_path)

    # Split the audio into segments based on silence (assumed silence duration: 700ms)
    segments = split_on_silence(audio, min_silence_len=700, silence_thresh=-40)

    recognized_text = ""

    # Initialize a recognizer
    recognizer = sr.Recognizer()

    for segment in segments:
        # Export the segment to a temporary WAV file
        temp_file = "temp.wav"
        segment.export(temp_file, format="wav")

        # Recognize the speech in the temporary file
        with sr.AudioFile(temp_file) as source:
            audio_data = recognizer.record(source)
            try:
                text = recognizer.recognize_google(audio_data)
                recognized_text += text + " "
            except sr.UnknownValueError:
                pass  # If speech cannot be recognized in the segment, move to the next one
            except sr.RequestError as e:
                print("Google Speech Recognition service error: {0}".format(e))

        # Clean up the temporary file
        os.remove(temp_file)

    return recognized_text.strip()

if __name__ == "__main__":
    try:
        recognized_text = recognize_audio_without_noise(audio_file_path)
        print("Recognized Text:")
        print(recognized_text)
    finally:
        print("###END###")
