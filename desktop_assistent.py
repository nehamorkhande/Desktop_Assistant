import pyttsx3
import datetime
import wikipedia
import webbrowser
import speech_recognition as sr
import os
import subprocess
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from comtypes import CLSCTX_ALL
import screen_brightness_control as sbc
from comtypes import CLSCTX_ALL

engine=pyttsx3.init('sapi5')
voices=engine.getProperty('voices')
#print(voices)

engine.setProperty('voice',voices[0].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour=int(datetime.datetime.now().hour)
    if hour>=0 and hour<12:
        speak("Good Morning")
    
    elif hour>=12 and hour<17:
        speak("Good Afternoon")
    
    else:
        speak("Good Evening!")

    speak("I am Jarvis . Please tell me how may i help you")


def takeCommand():
    r=sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold=1
        audio=r.listen(source)

    try:
        print("Recognizing...")
        query=r.recognize_google(audio, language='en-in')
        print(f"User Said: {query}\n")
    
    except Exception as e:
       # print(e)
        print("Say that again please...")
        return "None"
    return query


if __name__=="__main__":
   wishMe()
   while True:
       query=takeCommand().lower()
       if 'wikipedia' in query:
           speak('Searching Wikipedi...')
           query=query.replace("wikipedia","")
           results=wikipedia.summary(query,sentences=2)
           speak("According to Wikipedia")
           print(results)
           speak(results)
        
       elif 'open youtube' in query:
           webbrowser.open("youtube.com")
        
       
       elif 'open linkedin' in query:
           webbrowser.open("linkedin.com")
        
       elif 'open google' in query:
           webbrowser.open("google.com")
        
       elif 'open stackoverflow' in query:
           webbrowser.open("stackoverflow.com")
        
       elif 'play music' in query:
           music_dir='C:\\Users\\neham\\Music\\mp3 song'
           songs=os.listdir(music_dir)
           print(songs)
           os.startfile(os.path.join(music_dir, songs[0]))
        
       elif 'the time' in query:
           strTime=datetime.datetime.now().strftime("%H:%M:%S")
           print(strTime)
           speak(f"sir, the time is {strTime}")
    
       elif 'open code' in query:  
            code_path = "C:\\Users\\neham\\AppData\\Local\\Programs\\Microsoft VS Code\\Code.exe"  
            try:
                 if os.path.exists(code_path): 
                   os.startfile(code_path)  
                 else:
                  print("Visual Studio Code executable not found at the specified path.")
            except Exception as e:
                 print(f"An error occurred while trying to open VS Code: {e}")
      
       elif 'open word' in query:  
            word_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"  
            try:
                 if os.path.exists(word_path): 
                   os.startfile(word_path)  
                 else:
                  print("Visual Studio Code executable not found at the specified path.")
            except Exception as e:
                 print(f"An error occurred while trying to open word: {e}")

       elif 'open power point' in query:  
            word_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE"  
            try:
                 if os.path.exists(word_path): 
                   os.startfile(word_path)  
                 else:
                  print("Visual Studio Code executable not found at the specified path.")
            except Exception as e:
                 print(f"An error occurred while trying to open power point: {e}")

       elif 'open excel' in query:  
            word_path = "C:\\Program Files\\Microsoft Office\\root\\Office16\EXCEL.EXE"  
            try:
                 if os.path.exists(word_path): 
                   os.startfile(word_path)  
                 else:
                  print("Visual Studio Code executable not found at the specified path.")
            except Exception as e:
                 print(f"An error occurred while trying to open power point: {e}")
    
       elif 'shutdown' in query:
            os.system("shutdown /s /t 1")
      
       elif 'restart' in query:
            os.system("shutdown /r /t 1")
        
       elif 'sleep' in query:
            subprocess.call("rundll32.exe powrprof.dll,SetSuspendState 0,1,0", shell=True)

       elif 'lock' in query:
            os.system("rundll32.exe user32.dll,LockWorkStation")
       
       elif 'increase volume' in query:
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = interface.QueryInterface(IAudioEndpointVolume)
            current_volume = volume.GetMasterVolumeLevelScalar()
            new_volume = min(current_volume + 0.1, 1.0)  # Increase volume by 10%
            volume.SetMasterVolumeLevelScalar(new_volume, None)
            speak(f"Volume increased to {int(new_volume * 100)} percent.")

       elif 'decrease volume' in query:
           devices = AudioUtilities.GetSpeakers()
           interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
           volume = interface.QueryInterface(IAudioEndpointVolume)
           current_volume = volume.GetMasterVolumeLevelScalar()
           new_volume = max(current_volume - 0.1, 0.0)  # Decrease volume by 10%
           volume.SetMasterVolumeLevelScalar(new_volume, None)
           speak(f"Volume decreased to {int(new_volume * 100)} percent.")
      
       elif 'increase brightness' in query:

            try:
                current_brightness = sbc.get_brightness(display=0)[0]  # Get current brightness
                new_brightness = min(current_brightness + 10, 100)  # Increase by 10%, max is 100%
                sbc.set_brightness(new_brightness, display=0)
                speak(f"Brightness increased to {new_brightness} percent.")
            except Exception as e:
                print(f"An error occurred: {e}")
                speak("Sorry, I couldn't increase the brightness.")
        
       elif 'decrease brightness' in query:
            try:
                current_brightness = sbc.get_brightness(display=0)[0]  # Get current brightness
                new_brightness = max(current_brightness - 10, 0)  # Decrease by 10%, min is 0%
                sbc.set_brightness(new_brightness, display=0)
                speak(f"Brightness decreased to {new_brightness} percent.")
            except Exception as e:
                print(f"An error occurred: {e}")
                speak("Sorry, I couldn't decrease the brightness.")
      
       elif 'exit' in query:
            print("Exiting...")
            break
        
       