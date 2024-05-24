import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import random
import os
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL, CoInitialize
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from googletrans import Translator
import schedule
import re
import time

engine = pyttsx3.init('sapi5')
engine.setProperty('voice', engine.getProperty('voices')[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wish_me():
    hour = datetime.datetime.now().hour
    if 0 <= hour < 12:
        speak("Good Morning!")
    elif 12 <= hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak("I am Tom! Please tell me how may I help you.")

def take_command():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        audio = r.listen(source)
    try:
        print("Recognizing...")
        return r.recognize_google(audio, language='en-in').lower()
    except Exception:
        print("Say that again please...")
        return "None"

def set_volume(change_percentage, increase=True):
    CoInitialize()
    devices = AudioUtilities.GetSpeakers()
    volume = cast(devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None), POINTER(IAudioEndpointVolume))
    current_volume = volume.GetMasterVolumeLevelScalar()
    change_fraction = change_percentage / 100.0
    new_volume = min(1.0, current_volume + change_fraction) if increase else max(0.0, current_volume - change_fraction)
    volume.SetMasterVolumeLevelScalar(new_volume, None)
    action = "increased" if increase else "decreased"
    speak(f"Volume {action} by {change_percentage}%!")

def translate_text():
    speak("What text would you like to translate?")
    text_to_translate = take_command()
    speak("To which language would you like to translate?")
    target_language = take_command()
    translated_text = Translator().translate(text_to_translate, dest=target_language)
    speak(f"The translation to {target_language} is: {translated_text.text}")

def set_reminder():
    speak("What task would you like to set a reminder for?")
    task = take_command()
    speak("When would you like to be reminded? Please specify the time (in HH:MM format).")
    reminder_time = take_command()
    if re.match(r'^\d{2}:\d{2}$', reminder_time):
        schedule.every().day.at(reminder_time).do(lambda: speak(f"Reminder: {task}"))
    else:
        speak("Invalid time format. Please use HH:MM format.")

if __name__ == "__main__":
    wish_me()
    while True:
        query = take_command()

        if 'wikipedia' in query:
            speak("Searching Wikipedia...")
            results = wikipedia.summary(query.replace('wikipedia', ''), sentences=2)
            speak(f"According to Wikipedia: {results}")

        elif 'the time' in query:
            speak(f"Sir, the time is {datetime.datetime.now().strftime('%H:%M:%S')}")

        elif 'play music' in query:
            music_dir = 'D:\\songs'
            os.startfile(os.path.join(music_dir, random.choice(os.listdir(music_dir))))

        elif 'increase volume' in query:
            speak("Enter the percentage by which to increase the volume:")
            set_volume(float(take_command()), increase=True)

        elif 'decrease volume' in query:
            speak("Enter the percentage by which to decrease the volume:")
            set_volume(float(take_command()), increase=False)

        elif 'open website' in query:
            speak("Sure, please specify the website.")
            webbrowser.open(take_command())

        elif 'tell joke' in query:
            speak("You are a very funny joke in yourself. I laugh very hard looking at you every time HAHAHAHA")

        elif 'search web' in query:
            speak("Sure, what would you like to search for?")
            webbrowser.open(f"https://www.google.com/search?q={take_command()}")

        elif 'translate text' in query:
            translate_text()

        elif 'set reminder' in query:
            set_reminder()

        elif 'bye' in query:
            speak('Goodbye! Have a nice day!')
            break

        schedule.run_pending()
        time.sleep(1)
