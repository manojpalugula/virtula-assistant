import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia
import webbrowser
import os
import sys
import spotipy
from spotipy.oauth2 import SpotifyOAuth
from groq import Groq
import subprocess
import pyperclip
import time
import pyautogui
import re
import json
import shutil
from docx import Document
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
import requests
from config import (SPOTIPY_CLIENT_ID, SPOTIPY_CLIENT_SECRET, SPOTIPY_REDIRECT_URI,
                    GROQ_API_KEY, SCOPES, CREDENTIALS_FILE, TOKEN_FILE, WEATHER_API_KEY, NEWS_API_KEY)

# Initialize the text-to-speech engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

# Set up Spotify credentials
sp = spotipy.Spotify(auth_manager=SpotifyOAuth(
    client_id=SPOTIPY_CLIENT_ID,
    client_secret=SPOTIPY_CLIENT_SECRET,
    redirect_uri=SPOTIPY_REDIRECT_URI,
    scope="user-read-playback-state,user-modify-playback-state"
))

def authenticate_google_slides():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    return creds

creds = authenticate_google_slides()
service = build('slides', 'v1', credentials=creds)

def speak(audio):
    engine.say(audio)
    engine.runAndWait()

def wishMe():
    hour = datetime.datetime.now().hour
    if 0 <= hour < 12:
        speak("Good Morning!")
    elif 12 <= hour < 18:
        speak("Good Afternoon!")
    else:
        speak("Good Evening!")
    speak("I am Jarvis, your personal virtual assistant. Please tell me how may I help you")

def takeCommand():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        print("Listening...")
        r.pause_threshold = 1
        audio = r.listen(source)

    try:
        print("Recognizing...")
        query = r.recognize_google(audio, language='en-in')
        print(f"User said: {query}\n")
    except Exception as e:
        print("Say that again please...")
        return "None"
    return query

def playSpotifyTrack(track_name):
    results = sp.search(q=track_name, type='track', limit=1)
    if results['tracks']['items']:
        track = results['tracks']['items'][0]
        track_name = track['name']
        track_artists = ', '.join([artist['name'] for artist in track['artists']])
        track_uri = track['uri']
        track_id = track_uri.split(":")[-1]
        webbrowser.open(f"https://open.spotify.com/track/{track_id}")
        speak(f"Playing {track_name} by {track_artists} on Spotify")
        speak("Please press play in the browser if the song doesn't start automatically.")
    else:
        speak(f"Couldn't find {track_name} on Spotify")

def searchGoogle(query):
    url = f"https://www.google.com/search?q={query}"
    webbrowser.open(url)

def searchYouTube(query):
    url = f"https://www.youtube.com/results?search_query={query}"
    webbrowser.open(url)

def query_groq(query):
    client = Groq(api_key=GROQ_API_KEY)
    chat_completion = client.chat.completions.create(
        messages=[
            {
                "role": "user",
                "content": query,
            }
        ],
        model="llama3-8b-8192",
    )
    return chat_completion.choices[0].message.content

def add_slide(presentation_id, title, content):
    print("Adding slide...")

    requests = [
        {
            'createSlide': {
                'slideLayoutReference': {
                    'predefinedLayout': 'TITLE_AND_BODY'
                }
            }
        }
    ]

    response = service.presentations().batchUpdate(
        presentationId=presentation_id, body={'requests': requests}).execute()

    slide_id = response['replies'][0]['createSlide']['objectId']

    layouts = service.presentations().pages().get(
        presentationId=presentation_id, pageObjectId=slide_id).execute()

    title_id = None
    body_id = None

    for element in layouts.get('pageElements', []):
        if 'shape' in element:
            shape = element['shape']
            if 'placeholder' in shape:
                placeholder = shape['placeholder']
                if placeholder['type'] == 'TITLE':
                    title_id = element['objectId']
                elif placeholder['type'] == 'BODY':
                    body_id = element['objectId']

    if title_id and body_id:
        requests = [
            {
                'insertText': {
                    'objectId': title_id,
                    'text': title,
                    'insertionIndex': 0
                }
            },
            {
                'insertText': {
                    'objectId': body_id,
                    'text': content,
                    'insertionIndex': 0
                }
            }
        ]

        response = service.presentations().batchUpdate(
            presentationId=presentation_id, body={'requests': requests}).execute()

def create_presentation_with_content(title, content):
    presentation = service.presentations().create(body={'title': title}).execute()
    presentation_id = presentation['presentationId']
    add_slide(presentation_id, title, content)

def get_weather(city):
    base_url = f"http://api.openweathermap.org/data/2.5/weather?q={city}&appid={WEATHER_API_KEY}&units=metric"
    response = requests.get(base_url)
    data = response.json()
    if data['cod'] != '404':
        main = data['main']
        temperature = main['temp']
        weather_description = data['weather'][0]['description']
        return f"Temperature in {city} is {temperature}°C with {weather_description}."
    else:
        return f"City {city} not found."

def get_news(query):
    url = f'https://newsapi.org/v2/everything?q={query}&apiKey={NEWS_API_KEY}'
    response = requests.get(url)
    data = response.json()
    if data['status'] == 'ok':
        articles = data['articles']
        if articles:
            news = [f"{article['title']} - {article['source']['name']}" for article in articles[:5]]
            return news
        else:
            return ["No news articles found."]
    else:
        return ["Error retrieving news."]

if __name__ == "__main__":
    wishMe()
    while True:
        query = takeCommand().lower()

        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            searchYouTube(query)

        elif 'open google' in query:
            searchGoogle(query)

        elif 'play' in query and 'spotify' in query:
            track_name = query.replace("play", "").replace("on spotify", "").strip()
            playSpotifyTrack(track_name)

        elif 'create presentation' in query:
            speak('What is the title of the presentation?')
            title = takeCommand()
            speak('What is the content of the presentation?')
            content = takeCommand()
            create_presentation_with_content(title, content)
            speak('Presentation created successfully.')

        elif 'weather in' in query:
            city = query.split("in")[-1].strip()
            weather_info = get_weather(city)
            speak(weather_info)

        elif 'news about' in query:
            topic = query.split("about")[-1].strip()
            news_articles = get_news(topic)
            for article in news_articles:
                speak(article)

        elif 'exit' in query or 'quit' in query:
            speak("Goodbye!")
            break

        else:
            response = query_groq(query)
            speak(response)
