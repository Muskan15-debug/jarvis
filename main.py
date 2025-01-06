import speech_recognition as sr
import win32com.client
import pyaudio
import pyttsx3
import datetime
import wikipedia
import webbrowser
import os

from wikipedia import languages

engine=pyttsx3.init('sapi5');
voices=engine.getProperty('voices')
engine.setProperty('voices',voices[0].id)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour=int(datetime.datetime.now().hour)
    if hour>=0 and hour<=12:
        speak("Good Morning!")
    elif hour>=12 and hour<=18:
        speak("Good Afternoon!")
    else:
        speak("Good evening!")

    speak("i am jarvis  how may i help you")

def takeCommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold=1
        audio=r.listen(source)

    try:
        print("Recognising...")
        query=r.recognize_google(audio,language='en-in')
        print(f"user said:{query}")

    except Exception as e:
        print("Say that again please...")
        return "None"
    return query

if __name__=="__main__":
    wishMe()
    while True:
        query= takeCommand().lower()

        if 'wikipedia' in query:
            speak('searching wikipedia')
            query=query.replace("Wikipedia", "")
            results = wikipedia.summary(query,sentences=2)
            speak("According to wikipedia")
            print(results)
            speak(results)
        elif 'open youtube' in query:
            webbrowser.open("youtube.com")
        elif 'open google' in query:
            webbrowser.open("google.com")
        elif 'open spotify' in query:
            webbrowser.open("spotify.com")
        elif 'the time' in query:
            strTime=datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"The time is {strTime}")
        elif 'open code' in query:
            codePath="C:\\Users\\parag dalsania\\Desktop\\Microsoft VS Code\\Code.exe"
            os.startfile(codePath)
