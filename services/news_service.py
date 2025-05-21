import requests
from config import Config

def get_news_updates(topic="technology"):
    try:
        url = f"https://newsapi.org/v2/everything?q={topic}&apiKey={Config.NEWS_API_KEY}&pageSize=3"
        response = requests.get(url)
        data = response.json()
        
        if data['status'] == 'ok' and data['articles']:
            news_items = []
            for article in data['articles'][:3]:  # Limit to 3 articles
                title = article['title'].split(' - ')[0]
                news_items.append(f"{title}")
            
            return "Latest news:\n" + "\n".join(news_items)
        else:
            return "No news found at this time"
    except Exception as e:
        return f"News service error: {str(e)}"