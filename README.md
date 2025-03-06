# Jarvis - A Virtual Assistant

Jarvis is a sophisticated virtual assistant designed to facilitate a wide range of tasks through voice commands. It uses various APIs and libraries to provide functionalities such as text-to-speech conversion, speech recognition, web browsing, file management, and integration with external services like Spotify and Google Slides.

## Features

1. **Text-to-Speech and Speech Recognition**
    - Libraries: `pyttsx3`, `speech_recognition`
    - Functionality: Greets users, takes voice commands, and responds verbally.

2. **Spotify Integration**
    - Library: `Spotipy`
    - Functionality: Plays specific songs on Spotify based on voice commands.

3. **Google Slides API Integration**
    - Library: `Google Slides API`
    - Functionality: Creates and edits presentations, adds slides with content from the Groq API.

4. **Web Browsing**
    - Library: `webbrowser`
    - Functionality: Opens websites, performs Google searches, and searches for videos on YouTube.

5. **File Management**
    - Libraries: `os`, `shutil`, `pyperclip`, `pyautogui`, `subprocess`, `time`
    - Functionality: Creates, deletes, and opens files and folders. Lists files in a specified directory.

6. **Weather Updates**
    - Library: `requests`
    - Functionality: Provides current weather updates for any city using the OpenWeatherMap API.

7. **News Updates**
    - Library: `requests`
    - Functionality: Fetches the latest news articles on specified topics using NewsAPI.

8. **Content Generation with Groq API**
    - Library: `Groq API`
    - Functionality: Generates text content, including code snippets, based on user queries. Creates and populates Google Slides presentations with relevant content.

## Setup and Installation

1. Clone the repository:
    ```bash
    git clone https://github.com/your-username/jarvis-assistant.git
    cd jarvis-assistant
    ```

2. Install the required dependencies:
    ```bash
    pip install -r requirements.txt
    ```

3. Set up your API keys and credentials:
    - Rename `config_template.py` to `config.py`.
    - Update `config.py` with your API keys and credentials.

4. Run the application:
    ```bash
    python jarvis.py
    ```

## Configuration

Update the `config.py` file with your API keys and credentials:
```python
# Spotify API credentials
SPOTIPY_CLIENT_ID = "your_spotify_client_id"
SPOTIPY_CLIENT_SECRET = "your_spotify_client_secret"
SPOTIPY_REDIRECT_URI = "http://localhost:8888/callback"

# Groq API key
GROQ_API_KEY = "your_groq_api_key"

# Google Slides API setup
SCOPES = ['https://www.googleapis.com/auth/presentations']
CREDENTIALS_FILE = "path_to_your_credentials_file.json"
TOKEN_FILE = 'token.json'

# Weather API key
WEATHER_API_KEY = "your_openweathermap_api_key"

# News API key
NEWS_API_KEY = 'your_newsapi_key'
