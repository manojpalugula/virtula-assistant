from groq import Groq
from config import Config

client = Groq(api_key=Config.GROQ_API_KEY)

def process_groq_query(query):
    try:
        chat_completion = client.chat.completions.create(
            messages=[{"role": "user", "content": query}],
            model="llama3-8b-8192",
        )
        return chat_completion.choices[0].message.content
    except Exception as e:
        return f"Sorry, I encountered an error: {str(e)}"