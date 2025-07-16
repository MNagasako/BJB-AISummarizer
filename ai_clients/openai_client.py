import openai
from ai_clients.base import AIClient

class OpenAIClient(AIClient):
    def __init__(self, api_key: str, model: str):
        openai.api_key = api_key
        self.model = model

    def process(self, prompt: str) -> str:
        response = openai.ChatCompletion.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}],
            temperature=0.3
        )
        return response.choices[0].message.content.strip()
