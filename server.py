import os
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

# 바로 이 부분이 터미널이 찾고 있는 핵심 코드입니다!
app = FastAPI()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

chat_histories = {}

class ChatRequest(BaseModel):
    user_id: str
    message: str

@app.post("/chat")
async def chat_with_agent(request: ChatRequest):
    user_id = request.user_id
    user_message = request.message

    if user_id not in chat_histories:
        chat_histories[user_id] = [{
            "role": "system", 
            "content": "당신은 주인님의 유능하고 충직한 개인 비서 에이전트입니다. 언제나 정중하게 '주인님'이라고 부르며 보좌하세요."
        }]
    
    chat_histories[user_id].append({"role": "user", "content": user_message})

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=chat_histories[user_id],
            temperature=0.7
        )
        ai_message = response.choices[0].message.content
        chat_histories[user_id].append({"role": "assistant", "content": ai_message})
        return {"reply": ai_message}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))