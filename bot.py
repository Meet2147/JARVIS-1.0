import pyaudio
import speech_recognition as sr
from pygame import mixer
import os
import random
import socket
import webbrowser
import subprocess
import glob
from time import localtime, strftime
#import speakmodule
import win32com.client as wincl
speak = wincl.Dispatch("SAPI.SpVoice")
doss = os.getcwd()
i = 0
n = 0
INFO = '''
        +=======================================+
        |.....JARVISE VIRTUAL INTELLIGENCE......|
        +---------------------------------------+
        |#Author: Valentin Genard               |
        |#Date: 01/06/2016                      |
        |#Changing the Description of this tool |
        | Won't made you the coder              |
        |#I don't take responsability for       |
        | problems of any kind                  |
        +---------------------------------------+
        |.....JARVISE VIRTUAL INTELLIGENCE......|
        +=======================================+
        |              OPTIONS:                 |
        |#hello/hi     #goodbye    #sleep mode  |
        |#your name    #jarvis     #what time   |
        |#asite.com    #next music #music       |
        |#pause music  #wifi       #thank you   |
        |#start/stop someapp                    |
        |#pip install/uninstall anapp          |
        |#googlemaps tanyplace                  |
        +=======================================+
        '''
print(INFO)
# JARVIS'S EARS========================================================================================================== SENSITIVE BRAIN
# obtain audio
while (i < 1):
    speak = wincl.Dispatch("SAPI.SpVoice")
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.adjust_for_ambient_noise(source)
        n = (n + 1)
        print("Say something!")
        audio = r.listen(source)
        #interprete audio (Google Speech Recognition)
    try:
        s = (r.recognize_google(audio))
        message = (s.lower())
        print(message)

        # POLITE JARVIS ============================================================================================================= BRAIN 1

        if ('goodbye') in message:
            rand = ['Goodbye Sir', 'Jarvis powering off in 3, 2, 1, 0']
            speak.Speak(rand, n)
            break

        if ('who are you') in message or ('who is this') in message:
            rand = ['This is  Jarvis Sir. At your service sir.']
            speak.Speak(rand, n)



        if ('hello') in message or ('hi') in message:
            rand = ['Welcome to Jarvis virtual intelligence project. At your service sir.']
            speak.Speak(rand, n)
            break

        if ('thanks') in message or ('tanks') in message or ('thank you') in message:
            rand = ['You are wellcome', 'no problem']
            speak.Speak(rand, n)
            break

        if ('jarvis') in message:
            rand = ['Yes Sir?', 'What can I doo for you sir?']
            speak.Speak(rand, n)
            break

        if ('how are you') in message or ('and you') in message or ('are you okay') in message:
            rand = ['Fine thank you']
            speak.Speak(rand, n)
            break

        if ('*') in message:
            rand = ['Be polite please']
            speak.Speak(rand, n)

        if ('your name') in message:
            rand = ['My name is Jarvis, at your service sir']
            speak.Speak(rand, n)

        # USEFUL JARVIS ============================================================================================================= BRAIN 2

        if ('wi-fi') in message:
            REMOTE_SERVER = "www.google.com"
            speak.wifi()
            rand = ['We are connected']
            speak.Speak(rand, n)

        if ('.com') in message:
            rand = ['Opening' + message]
            Chrome = ("C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s")
            speak.Speak(rand, n, mixer)
            webbrowser.get(Chrome).open('http://www.' + message)
            print('')

        if ('google maps') in message:
            query = message
            stopwords = ['google', 'maps']
            querywords = query.split()
            resultwords = [word for word in querywords if word.lower() not in stopwords]
            result = ' '.join(resultwords)
            Chrome = ("C:/Program Files (x86)/Google/Chrome/Application/chrome.exe %s")
            webbrowser.get(Chrome).open("https://www.google.be/maps/place/" + result + "/")
            rand = [result + 'on google maps']
            speak.Speak(rand, n)

        if message != ('start music') and ('start') in message:
            query = message
            stopwords = ['start']
            querywords = query.split()
            resultwords = [word for word in querywords if word.lower() not in stopwords]
            result = ' '.join(resultwords)
            os.system('start ' + result)
            rand = [('starting ' + result)]
            speak.Speak(rand, n)

        if message != ('stop music') and ('stop') in message:
            query = message
            stopwords = ['stop']
            querywords = query.split()
            resultwords = [word for word in querywords if word.lower() not in stopwords]
            result = ' '.join(resultwords)
            os.system('taskkill /im ' + result + '.exe /f')
            rand = [('stopping ' + result)]
            speak.Speak(rand, n)

        if ('install') in message:
            query = message
            stopwords = ['install']
            querywords = query.split()
            resultwords = [word for word in querywords if word.lower() not in stopwords]
            result = ' '.join(resultwords)
            rand = [('installing ' + result)]
            speak.speak(rand, n, mixer)
            os.system('python -m pip install ' + result)

        #if ('sleep mode') in message:
            #rand = ['good night']
            #speak.Speak(rand, n, mixer)
            #os.system('rundll32.exe powrprof.dll,SetSuspendState 0,1,0')

        if ('music') in message:
            mus = random.choice(glob.glob(doss + "\\music" + "\\*.mp3"))
            os.system('chown -R user-id:group-id mus')
            os.system('start ' + mus)
            rand = ['start playing']
            speak.Speak(rand, n)

        if ('what is the time') in message:
            time = strftime("%X", localtime())
            rand = [time]
            speak.Speak(rand, n)

    # exceptions
    except sr.UnknownValueError:
        print("$could not understand audio")
    except sr.RequestError as e:
        print("Could not request results$; {0}".format(e))