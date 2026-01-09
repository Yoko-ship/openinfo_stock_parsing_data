from google import genai
from dotenv import load_dotenv
import os
from prompt import prompt

load_dotenv()
def make_analyz(data):

    client = genai.Client(api_key=os.getenv("api_key"))
    response = client.models.generate_content(
        model="gemini-3-flash-preview",
        contents=f"Промпт:{prompt} Сам фрейм: {data}",
    )
    return response.text.replace("###","").replace("*","").replace("**","")