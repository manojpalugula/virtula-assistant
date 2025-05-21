import requests
from config import Config

def get_weather_report(city):
    try:
        base_url = "http://api.openweathermap.org/data/2.5/weather?"
        complete_url = f"{base_url}q={city}&appid={Config.OPENWEATHER_API_KEY}&units=metric"
        
        response = requests.get(complete_url)
        data = response.json()
        
        if data["cod"] != "404":
            main = data["main"]
            weather = data["weather"][0]
            
            report = (
                f"Weather in {city}:\n"
                f"- Temperature: {main['temp']}°C\n"
                f"- Feels like: {main['feels_like']}°C\n"
                f"- Conditions: {weather['description'].capitalize()}\n"
                f"- Humidity: {main['humidity']}%\n"
                f"- Pressure: {main['pressure']} hPa"
            )
            
            return report
        else:
            return f"Weather data not found for {city}"
    except Exception as e:
        return f"Weather service error: {str(e)}"