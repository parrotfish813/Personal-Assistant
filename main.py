import speech_recognition as sr
import pyttsx3
import pywhatkit
import datetime
import time
import pyjokes
import os
import win32com.client
import pyautogui

listener = sr.Recognizer()
engine = pyttsx3.init()
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

engine.say("Hello, I am your personal assistant")
engine.say("Bob")
engine.say("What can I do for you today?")
engine.runAndWait()
print('listening...')


def talk(text):
    engine.say(text)
    engine.runAndWait()


def take_command():
    try:
        with sr.Microphone() as source:
            voice = listener.listen(source)
            command = listener.recognize_google(voice)
            command = command.lower()
    except:
        pass

    return command


def run_brandon():
    command = take_command()

    if 'play' in command:
        song = command.replace('play', '')
        talk('Playing' + song)
        pywhatkit.playonyt(song)

    elif 'time' in command:
        time = datetime.datetime.now().strftime('%I:%M %p')
        talk('It is currently' + time)

    elif 'tell me a joke' in command:
        talk(pyjokes.get_joke())

    elif 'open outlook' in command:
        os.startfile("outlook")
        talk("Opening")
        talk("MICROSOFT OUTLOOK")

    elif 'close outlook' in command:
        os.killpg(17988, 'outlook.exe')

    elif 'open chrome' in command:
        os.startfile("chrome")
        talk("Opening")
        talk("GOOGLE CHROME")

    elif "open excel" in command:
        os.startfile("excel")
        talk("Opening")
        talk("MICROSOFT EXCEL")

    elif 'close excel' in command:
        talk("Closing")
        talk("MICROSOFT EXCEL")
        os.kill(10, "excel")

    elif 'volume up' in command:
        pyautogui.press("volumeup")

    elif 'volume down' in command:
        pyautogui.press("volumedown")

    elif 'mute' in command:
        pyautogui.press("volumemute")

    elif 'stop' in command:
        exit()

    else:
        talk("Hmmmmmm I don't know that one yet")

    '''
    elif 'alarm' in command:
        talk('What time would you like the alarm set')
        tt = command.replace('set alarm to ', '')
        tt = tt.replace('.', '')
        tt = tt.upper()
        import Alarm
        Alarm.alarm(tt)
    '''


while True:
    run_brandon()
