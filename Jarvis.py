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
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx import Presentation
from pptx.util import Inches
import textwrap
import requests

# Initialize the text-to-speech engine
engine = pyttsx3.init('sapi5')
voices = engine.getProperty('voices')
engine.setProperty('voice', voices[0].id)

# Set up Spotify credentials
sp = spotipy.Spotify(auth_manager=SpotifyOAuth(
    client_id="bc4b741fbf23492f939365b367a95119",
    client_secret="1db92c2b6116461c8b6063925b21974c",
    redirect_uri="http://localhost:8888/callback",
    scope="user-read-playback-state,user-modify-playback-state"
))

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

def query_groq(query, api_key):
    client = Groq(api_key=api_key)
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

PIXABAY_API_KEY = "50245343-f63fa43cdf38d8cb5f42a3c68"  # Get from https://pixabay.com/api/docs/

def get_pixabay_image(query):
    """Fetch relevant image from Pixabay"""
    try:
        url = f"https://pixabay.com/api/?key={PIXABAY_API_KEY}&q={query}&image_type=photo&per_page=3"
        response = requests.get(url).json()
        if response['hits']:
            return response['hits'][0]['webformatURL']  # Medium quality image
    except Exception as e:
        print(f"Pixabay error: {e}")
    return None

def download_image(image_url, filename="temp_img.jpg"):
    """Download image from URL"""
    try:
        response = requests.get(image_url)
        with open(filename, 'wb') as f:
            f.write(response.content)
        return filename
    except Exception as e:
        print(f"Image download failed: {e}")
        return None

# --- Enhanced Content Processing ---
def clean_groq_content(content):
    """Remove unwanted * markers and format content"""
    # Remove * characters but preserve bullet structure
    content = content.replace('*', '')
    # Convert markdown-like bullets to proper formatting
    content = re.sub(r'^\s*-\s+', '• ', content, flags=re.MULTILINE)
    content = re.sub(r'^\s*•\s+', '• ', content, flags=re.MULTILINE)
    # Remove excessive line breaks
    content = '\n'.join([x.strip() for x in content.split('\n') if x.strip()])
    return content

def split_into_slides(content, max_lines=6):
    """Split content into slide-sized chunks with smart wrapping"""
    paragraphs = [p for p in content.split('\n') if p.strip()]
    slides = []
    current_slide = []
    
    for para in paragraphs:
        wrapped = textwrap.wrap(para, width=80)  # 80 chars per line
        if len(current_slide) + len(wrapped) > max_lines:
            slides.append('\n'.join(current_slide))
            current_slide = []
        current_slide.extend(wrapped)
    
    if current_slide:
        slides.append('\n'.join(current_slide))
    
    return slides if slides else [content]

# --- Enhanced PowerPoint Creation ---
def create_ppt_file_with_content(file_path, topic, api_key):
    """Generate a structured PowerPoint with text on the left and scaled image on the right."""
    prompt = f"""
Create a PowerPoint presentation about "{topic}" with 3 to 5 slides.
Each slide must follow this format:

[SLIDE TITLE: Slide title here]
[CONTENT:
- Bullet point 1
- Bullet point 2
- Bullet point 3
]
[IMAGE KEYWORD: relevant image keyword]

Use '---' to separate each slide.
Stick exactly to the format.
"""
    raw_content = query_groq(prompt, api_key)
    slides = [slide.strip() for slide in raw_content.split('---') if slide.strip()]

    prs = Presentation()
    slide_width = prs.slide_width
    slide_height = prs.slide_height

    # Title Slide
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = topic
    title_slide.placeholders[1].text = "Created by Jarvis Assistant"

    for i, slide_content in enumerate(slides[:5]):
        title_match = re.search(r'\[SLIDE TITLE:\s*(.+?)\]', slide_content, re.IGNORECASE)
        content_match = re.search(r'\[CONTENT:(.+?)\](?:\[IMAGE|\Z)', slide_content, re.DOTALL | re.IGNORECASE)
        image_match = re.search(r'\[IMAGE KEYWORD:\s*(.+?)\]', slide_content, re.IGNORECASE)

        if not content_match or not content_match.group(1).strip():
            continue

        # Blank layout
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        # Title box
        title_text = title_match.group(1).strip() if title_match else f"{topic} - Slide {i+1}"
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.2), slide_width - Inches(1), Inches(1))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = f'"{title_text}"'
        p.font.size = Pt(32)
        p.font.bold = True
        p.font.color.rgb = RGBColor(59, 89, 152)
        p.alignment = PP_ALIGN.LEFT

        # Text on left
        left_box = slide.shapes.add_textbox(Inches(0.5), Inches(1.2), Inches(5.5), slide_height - Inches(2))
        tf = left_box.text_frame
        tf.word_wrap = True

        bullet_lines = [line.strip("-*• ") for line in content_match.group(1).strip().split("\n") if line.strip()]
        for line in bullet_lines:
            para = tf.add_paragraph()
            para.text = f"• {line}"
            para.level = 0
            para.font.size = Pt(20)

        # Image on right
        if image_match:
            img_keyword = image_match.group(1).strip()
            img_url = get_pixabay_image(img_keyword)
            if img_url:
                img_path = download_image(img_url)
                if img_path:
                    try:
                        # Smaller, better-aligned image
                        left = Inches(6.0)
                        top = Inches(1.8)
                        slide.shapes.add_picture(img_path, left, top, width=Inches(3.5))
                        os.remove(img_path)
                    except Exception as e:
                        print(f"Image error: {e}")

    # Thank you slide
    thank_slide = prs.slides.add_slide(prs.slide_layouts[5])
    txBox = thank_slide.shapes.add_textbox(Inches(1), Inches(2), prs.slide_width - Inches(2), Inches(1.5))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "Thank You!"
    p.font.size = Pt(48)
    p.font.color.rgb = RGBColor(59, 89, 152)
    p.alignment = PP_ALIGN.CENTER

    try:
        prs.save(file_path)
        speak(f"Presentation saved as {os.path.basename(file_path)}")
        open_file(file_path)
    except PermissionError:
        speak("Please close the file if it's already open and try again.")

def extract_code(response):
    """
    Extracts the code portion from the given response.
    """
    code_pattern = re.compile(r'```(.*?)```', re.DOTALL)
    match = code_pattern.search(response)
    if match:
        return match.group(1).strip()
    return response

def open_notepad_with_code(code):
    """
    Opens Notepad and writes the given code into it.
    """
    notepad_path = "notepad.exe"
    process = subprocess.Popen(notepad_path)
    time.sleep(1)
    pyperclip.copy(code)
    pyautogui.hotkey("ctrl", "v")

def create_text_file_with_content(file_path, topic, api_key):
    content = query_groq(topic, api_key)
    with open(file_path, 'w') as file:
        file.write(content)
    speak(f"Text file '{os.path.basename(file_path)}' created on Desktop")
    open_file(file_path)

def create_word_file_with_content(file_path, topic, api_key):
    content = query_groq(topic, api_key)
    doc = Document()
    doc.add_heading(topic, 0)
    doc.add_paragraph(content)
    doc.save(file_path)
    speak(f"Word file '{os.path.basename(file_path)}' created on Desktop")
    open_file(file_path)

def create_folder(folder_path):
    os.makedirs(folder_path, exist_ok=True)
    speak(f"Folder '{os.path.basename(folder_path)}' created on Desktop")
    open_folder(folder_path)

def delete_file(file_path):
    try:
        os.remove(file_path)
        speak(f"File '{os.path.basename(file_path)}' deleted from Desktop")
    except FileNotFoundError:
        speak(f"File '{os.path.basename(file_path)}' not found on Desktop")

def open_file(file_path):
    try:
        os.startfile(file_path)
    except FileNotFoundError:
        speak(f"File '{os.path.basename(file_path)}' not found on Desktop")

def open_folder(folder_path):
    try:
        os.startfile(folder_path)
    except FileNotFoundError:
        speak(f"Folder '{os.path.basename(folder_path)}' not found on Desktop")

def list_files_in_folder(folder_path):
    try:
        files = os.listdir(folder_path)
        if files:
            speak(f"Files in '{os.path.basename(folder_path)}' are:")
            for file in files:
                speak(file)
        else:
            speak(f"No files found in '{os.path.basename(folder_path)}'")
    except FileNotFoundError:
        speak(f"Folder '{os.path.basename(folder_path)}' not found on Desktop")

def get_weather(city):
    api_key = "ba621bd65e6bf4c55568056c986c200e"
    base_url = "http://api.openweathermap.org/data/2.5/weather?"
    complete_url = base_url + "q=" + city + "&appid=" + api_key + "&units=metric"
    response = requests.get(complete_url)
    weather_data = response.json()
    
    if weather_data["cod"] != "404":
        main = weather_data["main"]
        wind = weather_data["wind"]
        weather = weather_data["weather"][0]
        temperature = main["temp"]
        humidity = main["humidity"]
        weather_description = weather["description"]
        wind_speed = wind["speed"]

        weather_report = (f"Temperature: {temperature}°C\n"
                          f"Humidity: {humidity}%\n"
                          f"Weather description: {weather_description}\n"
                          f"Wind speed: {wind_speed} meter per second")

        print(weather_report)
        speak(weather_report)
    else:
        speak("City not found. Please try again.")

def get_news(api_key, query):
    url = f"https://newsapi.org/v2/everything?q={query}&apiKey={api_key}"
    response = requests.get(url)
    data = response.json()
    if data['status'] == 'ok':
        articles = data['articles']
        news_list = []
        for article in articles[:5]:  # Limit to the top 5 news articles
            news_list.append(f"Title: {article['title']}. Description: {article['description']}.")
        return " ".join(news_list)
    else:
        return "Could not retrieve news."

if __name__ == "__main__":
    api_key = "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"
    news_api_key = '0af8f4aac7614faa804bbccc4cf9fca2'

    wishMe()
    while True:
        query = takeCommand().lower()

        if query == "none":
            continue

        if 'wikipedia' in query:
            speak('Searching Wikipedia...')
            query = query.replace("wikipedia", "")
            results = wikipedia.summary(query, sentences=2)
            speak("According to Wikipedia")
            print(results)
            speak(results)

        elif 'open youtube' in query:
            webbrowser.open("youtube.com")

        elif 'open google' in query:
            webbrowser.open("google.com")

        elif 'search youtube' in query:
            speak("What would you like to search on YouTube?")
            search_query = takeCommand()
            searchYouTube(search_query)
            continue

        elif 'search google' in query:
            speak("What would you like to search on Google?")
            search_query = takeCommand()
            searchGoogle(search_query)
            continue

        elif 'play a song' in query:
            speak("Which song would you like to hear?")
            song = takeCommand().lower()
            playSpotifyTrack(song)

        elif 'the time' in query:
            strTime = datetime.datetime.now().strftime("%H:%M:%S")
            speak(f"The time is {strTime}")
            print(f"The time is {strTime}")

        elif 'quit' in query:
            speak("Have a nice day")
            sys.exit()

        elif 'write a code' in query:
            speak("What code would you like to write?")
            code_query = takeCommand().lower()
            code_response = query_groq(code_query, api_key)
            print("Code Response:", code_response)
            speak("I have written the code. Opening Notepad now.")
            code_only = extract_code(code_response)
            open_notepad_with_code(code_only)

        elif 'create a text file' in query:
            speak("What should be the name of the text file?")
            file_name = takeCommand()
            speak("What should be the topic of the text file?")
            topic = takeCommand()
            file_path = os.path.join(os.path.expanduser('~'), 'Desktop', file_name + '.txt')
            create_text_file_with_content(file_path, topic, api_key)

        elif 'create a word file' in query:
            speak("What should be the name of the Word file?")
            file_name = takeCommand()
            speak("What should be the topic of the Word file?")
            topic = takeCommand()
            file_path = os.path.join(os.path.expanduser('~'), 'Desktop', file_name + '.docx')
            create_word_file_with_content(file_path, topic, api_key)

        elif 'create a folder' in query:
            speak("What should be the name of the folder?")
            folder_name = takeCommand()
            folder_path = os.path.join(os.path.expanduser('~'), 'Desktop', folder_name)
            create_folder(folder_path)
            
        elif 'delete file' in query:
            speak("What is the name of the file to delete?")
            file_name = takeCommand()
            file_name = file_name.replace(" dot ", ".").replace(" ", "")
            file_path = os.path.join(os.path.expanduser('~'), 'Desktop', file_name)
            delete_file(file_path)

        elif 'open file' in query:
            speak("What is the name of the file to open?")
            file_name = takeCommand()
            file_name = file_name.replace(" dot ", ".").replace(" ", "")
            file_path = os.path.join(os.path.expanduser('~'), 'Desktop', file_name)
            open_file(file_path)

        elif 'open folder' in query:
            speak("What is the name of the folder to open?")
            folder_name = takeCommand()
            folder_path = os.path.join(os.path.expanduser('~'), 'Desktop', folder_name)
            open_folder(folder_path)

        elif 'list files' in query:
            speak("Which folder do you want to list the files of?")
            folder_name = takeCommand()
            folder_path = os.path.join(os.path.expanduser('~'), 'Desktop', folder_name)
            list_files_in_folder(folder_path)

        elif 'weather' in query:
            speak("Please tell me the city name.")
            city_name = takeCommand()
            get_weather(city_name)

        elif 'news' in query:
            speak("What topic would you like news about?")
            news_topic = takeCommand().lower()
            news_update = get_news(news_api_key, news_topic)
            speak(news_update)
        
        elif 'create a powerpoint' in query or 'create a ppt' in query or 'create ppt' in query or 'create presentation' in query:
            speak("What should be the name of the PowerPoint file?")
            file_name = takeCommand()
            speak("What should be the presentation about?")
            topic = takeCommand()
            file_path = os.path.join(os.path.expanduser('~'), 'Desktop', file_name + '.pptx')
            create_ppt_file_with_content(file_path, topic, api_key)

        else:
            response = query_groq(query, api_key)
            print("Response:", response)
            speak(response)
