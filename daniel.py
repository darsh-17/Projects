import pyttsx3
import datetime
import speech_recognition as sr
import wikipedia
import webbrowser
import os
import numpy as np
import pandas as pd

# Sapi5 is an API by Microsoft to help developers working with voice and speech recognization
engine = pyttsx3.init("sapi5")
voices = engine.getProperty("voices")
# print(voices])#Two voices one male one female
# Means we have chosen male voice which is at 0th index
engine.setProperty("voice", voices[0].id)


def speak(audio):
    engine.say(audio)
    engine.runAndWait()


def wishMe():
    hour = int(datetime.datetime.now().hour)
    if hour >= 4 and hour <= 12:
        speak("Good Morning")
    elif hour > 12 and hour <= 17:
        speak("Good Afternoon")
    else:
        speak("Good Evening")


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening....")
        # If nothing is spoken for 1 second, then daniel will consider that the phrase is completed
        r.pause_threshold = 1
        audio = r.listen(source, phrase_time_limit=8)

    try:
        print("Recognising...")
        # Uses google to understand/translate to text the input audio
        query = r.recognize_google(audio, language="en-in")
        print("User said ", query)

    except Exception as e:
        # print(e)
        print("Say that again please....")
        return "None"

    return query  # query simply contains the meaning of what we have said


if __name__ == "__main__":
    wishMe()
    speak("I am Daniel, your personalised virtual assistant. How may I help you?")

    # Actual coding logic
    if 1:
        query = takeCommand().lower()  # turn the query into lower case

        if "wikipedia" in query or "who is" in query or "tell me about" in query:
            speak("Searching Wikipedia")
            # To remove "wikipedia" from our query so that the actual thing is searched directly
            query = query.replace("wikipedia", "")
            # passing our query into wikipdedia module and telling its summary in 2 sentences
            results = wikipedia.summary(query, sentences=2)
            print(results)
            speak("According to Wikipedia")
            speak(results)

        elif "open youtube" in query:
            speak("Sure, what do you want to watch?")
            query2 = takeCommand().lower()
            webbrowser.open(
                f"https://www.youtube.com/results?search_query={query2}")

        elif "open google" in query:
            speak("Sure!")
            webbrowser.open("google.com")

        elif "play" in query:
            speak("Playing a song")
            music_dir = "C:\\Users\\DARSH\\Desktop\\Music"
            songs = os.listdir(music_dir)
            num = np.random.random_integers(0, (len(songs)-1))
            os.startfile(os.path.join(music_dir, songs[num]))

        elif "current time" in query:
            time = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"Sir, the time is {time}")

        elif "day" in query:
            day = datetime.datetime.now().strftime("%A")
            speak(f"The day is {day}")

        elif "visual studio" in query:
            speak("Opening Visual Studio Code")
            path = "C:\\Users\\DARSH\\AppData\\Local\Programs\\Microsoft VS Code\\Code.exe"
            os.startfile(path)

        elif "set reminder" in query or "set another reminder" in query or "set a reminder" in query:
            df = pd.DataFrame(columns=["Date", "Purpose"])            
            speak("Sure sir, tell me the date")
            query3 = takeCommand().lower()
            speak("Okay sir. Can you tell me what is the reminder")
            query4 = takeCommand().lower()
            print(query4)
            df.loc[len(df)]=[query3,query4]
            df.to_excel("C:/Users/DARSH/Desktop/Reminders.xlsx")
            speak("Reminder successfully set")
