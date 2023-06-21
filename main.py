import datetime
import webbrowser

import speech_recognition as sr
import win32com.client
import os
import openai
from config import apikey

speaker = win32com.client.Dispatch("SAPI.SpVoice")


chatStr = ""

def chat(query) :
    global chatStr
    print(chatStr)
    openai.api_key = apikey
    chatStr += f"Ayush: {query}\n Jarvis: "
    response = openai.Completion.create(
        model="text-davinci-003",
        prompt=chatStr,
        temperature=0.7,
        max_tokens=256,
        top_p=1,
        frequency_penalty=0,
        presence_penalty=0
    )
    speaker.Speak(response["choices"][0]["text"])
    chatStr += f"{response['choices'][0]['text']}\n"
    return response["choices"][0]["text"]


def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        r.pause_threshold = 0.6
        audio = r.listen(source)
        try:
            print("Recognizing...")
            query = r.recognize_google(audio, language="en-in")
            print(f"User said : {query}")
            return query

        except Exception as e:
            return "Some Error Occurred. Sorry from Jarvis"

if __name__ == '__main__':
    speaker.Speak("Say some thing")
    while True:
        print("Listening...")
        query = takeCommand()

        if "quit" in query.lower():
            speaker.Speak("Thank you for using Jarvis...")
            break

        sites = [["youtube", "https://www.youtube.com"],
                 ["wikipedia", "http://www.wikipedia.com"],
                 ["google", "http://www.google.com"]]

        for site in sites:
            if f"open {site[0]}" in query.lower():
                speaker.Speak(f"Opening {site[0]} sir...")
                webbrowser.open(site[1])

        if "play music" in query.lower():
            musicPath = "C:/Users/ayush/Videos/Music/Nomyn-Forever.mp3"
            os.startfile(musicPath)
        elif "reset chat" in query.lower():
            chatStr = ""
        else:
            print("Chatting..")
            chat(query)