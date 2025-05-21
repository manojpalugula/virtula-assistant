import spotipy
from spotipy.oauth2 import SpotifyOAuth
from config import Config

sp = spotipy.Spotify(auth_manager=SpotifyOAuth(
    client_id=Config.SPOTIFY_CLIENT_ID,
    client_secret=Config.SPOTIFY_CLIENT_SECRET,
    redirect_uri="http://localhost:5000/callback",
    scope="user-read-playback-state,user-modify-playback-state"
))

def play_spotify_track(track_name):
    try:
        results = sp.search(q=track_name, type='track', limit=1)
        if results['tracks']['items']:
            track = results['tracks']['items'][0]
            track_uri = track['uri']
            
            # Start playback (requires Spotify app to be open)
            sp.start_playback(uris=[track_uri])
            
            return f"Now playing: {track['name']} by {', '.join([a['name'] for a in track['artists']])}"
        else:
            return f"Could not find '{track_name}' on Spotify"
    except Exception as e:
        return f"Spotify error: {str(e)}. Please ensure your Spotify app is open."