import win32com.client as wincl
import speech_recognition as sr
import pyaudio
import google
speak = wincl.Dispatch("SAPI.SpVoice")
speak.Speak("Hello Megha, how may i help ")


r = sr.Recognizer()
with sr.Microphone() as source:
    print("Speak:")
    audio = r.listen(source)
    text = r.recognize_google(audio)
    print(text)
