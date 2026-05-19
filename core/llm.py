import os
from typing import Any

from dotenv import load_dotenv
from openai import OpenAI


load_dotenv()


class LLMClient:
    def __init__(self, api_key: str | None = None, model: str | None = None):
        self.api_key = api_key or os.getenv("OPENAI_API_KEY", "")
        self.model = model or os.getenv("OPENAI_MODEL", "gpt-5.5")

        if not self.api_key:
            raise ValueError("OPENAI_API_KEY가 설정되지 않았습니다. .env 또는 사이드바에 입력하세요.")

        self.client = OpenAI(api_key=self.api_key)

    def generate(
        self,
        system_prompt: str,
        user_prompt: str,
        temperature: float = 0.25,
        max_output_tokens: int = 2000,
    ) -> str:
        try:
            response = self.client.responses.create(
                model=self.model,
                instructions=system_prompt,
                input=user_prompt,
                temperature=temperature,
                max_output_tokens=max_output_tokens,
            )
        except TypeError:
            response = self.client.responses.create(
                model=self.model,
                instructions=system_prompt,
                input=user_prompt,
                max_output_tokens=max_output_tokens,
            )

        return self._extract_text(response)

    @staticmethod
    def _extract_text(response: Any) -> str:
        text = getattr(response, "output_text", None)
        if text:
            return text.strip()

        chunks: list[str] = []
        for item in getattr(response, "output", []) or []:
            for content in getattr(item, "content", []) or []:
                content_text = getattr(content, "text", None)
                if content_text:
                    chunks.append(content_text)

        return "\n".join(chunks).strip()
