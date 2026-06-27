from openai import OpenAI
import json

client = OpenAI(
    api_key="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",
    base_url="https://openrouter.ai/api/v1"
)
def ask_qwen(prompt):

    response = client.chat.completions.create(
        model="qwen/qwen3.6-plus",
        messages=[
            {
                "role": "user",
                "content": prompt
            }
        ]
    )

    return response.choices[0].message.content
