import pyttsx3
import speech_recognition as sr
import datetime
import wikipedia # Changed from 'import wikipedia'
import webbrowser
import os
import sys
import spotipy
from spotipy.oauth2 import SpotifyOAuth
from groq import Groq
import subprocess
import pyperclip
import time
# import pyautogui # Commented out for now, will be conditionally imported
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
import keyboard # Replaced pynput with keyboard
import time # Ensure time is imported

# Conditional import for pyautogui
try:
    import pyautogui
    pyautogui_available = True
except Exception as e: # Catch generic exception for display issues like KeyError: 'DISPLAY' or import errors
    print(f"PyAutoGUI could not be imported or initialized: {e}. GUI automation features will be disabled.")
    pyautogui_available = False

# Global state variables
jarvis_active = False
hotkey_registered = False # To track if the hotkey is active
pyautogui_available = False # Ensure this is initialized before the try-except block for import

# Initialize the text-to-speech engine
engine = pyttsx3.init() # Removed 'sapi5' to allow auto-detection
voices = engine.getProperty('voices')
# Attempt to set a voice, but handle if no voices are available (e.g., in some minimal environments)
if voices and len(voices) > 0:
    engine.setProperty('voice', voices[0].id)
else:
    print("No TTS voices found. Speech output may not work.")

# Set up Spotify credentials
sp = spotipy.Spotify(auth_manager=SpotifyOAuth(
    client_id="bc4b741fbf23492f939365b367a95119",
    client_secret="1db92c2b6116461c8b6063925b21974c",
    redirect_uri="http://localhost:8888/callback",
    scope="user-read-playback-state,user-modify-playback-state"
))

# Activation callback function
def activate_jarvis():
    global jarvis_active, hotkey_registered
    if not jarvis_active: # Ensure it only activates once if somehow called multiple times
        print("Hotkey 'ctrl+j' pressed, activating Jarvis...")
        # speak("Jarvis is activating.") # TTS might not work in all envs
        jarvis_active = True
        
        # Remove the hotkey to prevent re-triggering while Jarvis is active
        # and to allow Ctrl+J to be used for other purposes if Jarvis is in a long task.
        if hotkey_registered:
            try:
                keyboard.remove_hotkey("ctrl+j")
                print("Ctrl+J hotkey has been de-registered.")
                hotkey_registered = False
            except KeyError:
                # This might happen if the key was already removed or never properly registered
                print("Warning: Ctrl+J hotkey could not be removed or was not found.")
            except Exception as e:
                print(f"An error occurred while removing hotkey: {e}")

def start_hotkey_listener(): # Renamed function
    global hotkey_registered
    if not hotkey_registered:
        try:
            # Register "ctrl+j" as the hotkey
            keyboard.add_hotkey("ctrl+j", activate_jarvis, suppress=False) # Changed hotkey
            hotkey_registered = True
            print("Hotkey listener started. Press Ctrl+J to activate Jarvis.") # Changed message
            # Removed note specific to FN key
        except Exception as e:
            print(f"Error setting up hotkey listener: {e}") # Changed message
            print("Hotkey detection might not be supported or may require administrator privileges.") # Changed message
            hotkey_registered = False

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
    notepad_path = "notepad.exe" # This is Windows-specific
    pyperclip.copy(code) # Copying to clipboard should still work
    speak("Code has been copied to your clipboard.")
    try:
        # Try to open Notepad, but don't fail if it's not Windows or doesn't open
        if sys.platform == "win32": # Check if running on Windows
            subprocess.Popen(notepad_path)
            time.sleep(1) # Give Notepad a moment to open
            if pyautogui_available:
                pyautogui.hotkey("ctrl", "v")
                speak("Pasted into Notepad.")
            else:
                speak("Please paste the code manually into Notepad or your preferred editor.")
        else:
            speak("Please paste the code into your preferred text editor.")
    except Exception as e:
        print(f"Could not open Notepad or paste: {e}")
        speak("Could not open Notepad. Please paste the code manually.")

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

    start_hotkey_listener() # Updated function call

    try:
        while not jarvis_active:
            time.sleep(0.1) # Wait for activation

        # The activate_jarvis function now handles hotkey removal.
        # Redundant block removed as per instructions.

        wishMe()
        # Main command loop
        while True:
            query = takeCommand().lower()

            if query == "none":
                continue

            if 'wikipedia' in query:
                speak('Searching Wikipedia...')
                query = query.replace("wikipedia", "").strip()
                # Initialize Wikipedia API with a user agent
                wiki_wiki = wikipediaapi.Wikipedia(
                    language='en',
                    user_agent="JarvisAssistant/1.0 (https://example.com/jarvis; jules@example.com)"
                )
                page_py = wiki_wiki.page(query)
                if page_py.exists():
                    # Get a summary, limit length to avoid very long outputs
                    # Taking first 2 sentences by splitting and joining.
                    # Summary often has newlines, so take a good chunk and then select sentences.
                    full_summary = page_py.summary
                    # A simple way to get first few sentences if summary is long
                    sentences = full_summary.split('. ')
                    if len(sentences) >= 2:
                        results = sentences[0] + ". " + sentences[1] + "."
                    else:
                        results = full_summary[0:500] # Fallback to char limit

                    speak("According to Wikipedia")
                    print(results)
                    speak(results)
                else:
                    speak(f"Sorry, I could not find a Wikipedia page for {query}")
                    results = "No page found."

            elif 'open youtube' in query:
                webbrowser.open("youtube.com")

            elif 'open google' in query:
                webbrowser.open("google.com")

            elif 'search youtube' in query:
                speak("What would you like to search on YouTube?")
                search_query = takeCommand()
                if search_query != "None": 
                    searchYouTube(search_query)
                continue

            elif 'search google' in query:
                speak("What would you like to search on Google?")
                search_query = takeCommand()
                if search_query != "None": 
                    searchGoogle(search_query)
                continue

            elif 'play a song' in query:
                speak("Which song would you like to hear?")
                song = takeCommand().lower()
                if song != "none": 
                    playSpotifyTrack(song)

            elif 'the time' in query:
                strTime = datetime.datetime.now().strftime("%H:%M:%S")
                speak(f"The time is {strTime}")
                print(f"The time is {strTime}")

            elif 'quit' in query or 'exit' in query or 'stop' in query:
                speak("Goodbye! Have a nice day.")
                if hotkey_registered: # Should be false, but as a safeguard
                    try:
                        keyboard.remove_hotkey("ctrl+j") # Changed hotkey
                    except KeyError:
                        pass
                sys.exit()

            elif 'write a code' in query:
                speak("What code would you like to write?")
                code_query = takeCommand().lower()
                if code_query != "none":
                    code_response = query_groq(code_query, api_key)
                    print("Code Response:", code_response)
                    speak("I have written the code. Opening Notepad now.")
                    code_only = extract_code(code_response)
                    open_notepad_with_code(code_only)

            elif 'create a text file' in query:
                speak("What should be the name of the text file?")
                file_name = takeCommand()
                if file_name != "none":
                    speak("What should be the topic of the text file?")
                    topic = takeCommand()
                    if topic != "none":
                        file_path = os.path.join(os.path.expanduser('~'), 'Desktop', file_name + '.txt')
                        create_text_file_with_content(file_path, topic, api_key)

            elif 'create a word file' in query:
                speak("What should be the name of the Word file?")
                file_name = takeCommand()
                if file_name != "none":
                    speak("What should be the topic of the Word file?")
                    topic = takeCommand()
                    if topic != "none":
                        file_path = os.path.join(os.path.expanduser('~'), 'Desktop', file_name + '.docx')
                        create_word_file_with_content(file_path, topic, api_key)
            
            elif 'create a powerpoint' in query or 'create a ppt' in query or 'create ppt' in query or 'create presentation' in query:
                speak("What should be the name of the PowerPoint file?")
                file_name = takeCommand()
                if file_name != "none":
                    speak("What should be the presentation about?")
                    topic = takeCommand()
                    if topic != "none":
                        file_path = os.path.join(os.path.expanduser('~'), 'Desktop', file_name + '.pptx')
                        create_ppt_file_with_content(file_path, topic, api_key)

            elif 'create a folder' in query:
                speak("What should be the name of the folder?")
                folder_name = takeCommand()
                if folder_name != "none":
                    folder_path = os.path.join(os.path.expanduser('~'), 'Desktop', folder_name)
                    create_folder(folder_path)
                
            elif 'delete file' in query:
                speak("What is the name of the file to delete?")
                file_name = takeCommand()
                if file_name != "none":
                    file_name = file_name.replace(" dot ", ".").replace(" ", "")
                    file_path = os.path.join(os.path.expanduser('~'), 'Desktop', file_name)
                    delete_file(file_path)

            elif 'open file' in query:
                speak("What is the name of the file to open?")
                file_name = takeCommand()
                if file_name != "none":
                    file_name = file_name.replace(" dot ", ".").replace(" ", "")
                    file_path = os.path.join(os.path.expanduser('~'), 'Desktop', file_name)
                    open_file(file_path)

            elif 'open folder' in query:
                speak("What is the name of the folder to open?")
                folder_name = takeCommand()
                if folder_name != "none":
                    folder_path = os.path.join(os.path.expanduser('~'), 'Desktop', folder_name)
                    open_folder(folder_path)

            elif 'list files' in query:
                speak("Which folder do you want to list the files of?")
                folder_name = takeCommand()
                if folder_name != "none":
                    folder_path = os.path.join(os.path.expanduser('~'), 'Desktop', folder_name)
                    list_files_in_folder(folder_path)

            elif 'weather' in query:
                speak("Please tell me the city name.")
                city_name = takeCommand()
                if city_name != "none":
                    get_weather(city_name)

            elif 'news' in query:
                speak("What topic would you like news about?")
                news_topic = takeCommand().lower()
                if news_topic != "none":
                    news_update = get_news(news_api_key, news_topic)
                    speak(news_update)
            
            else: # Default to Groq for other queries
                response = query_groq(query, api_key)
                print("Response:", response)
                speak(response)

    except KeyboardInterrupt:
        print("Program interrupted by user.")
    finally:
        if hotkey_registered:
            try:
                keyboard.remove_hotkey("ctrl+j") # Changed hotkey
                print("Ctrl+J hotkey removed on exit.") # Changed message
            except KeyError:
                pass # Already removed or never set
        print("Exiting Jarvis.")
