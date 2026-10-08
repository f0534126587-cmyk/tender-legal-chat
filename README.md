# Tender Legal Chat

AI-powered legal assistant for drafting municipal tenders in Israel.

## Tech Stack

- **Backend:** Python, FastAPI, Anthropic Claude (claude-opus-4-8)
- **Frontend:** React, TypeScript, Vite

## Features

- Real-time streaming chat responses
- Guided tender drafting based on Israeli municipal law
- Export conversation to Word document (.docx)

## Getting Started

### Backend

cd backend

pip install -r requirements.txt



Create a `.env` file and add your Anthropic API key:

ANTHROPIC_API_KEY=your-api-key-here


uvicorn main:app --reload --port 8001



### Frontend

cd frontend

npm install

npm run dev



Open http://localhost:5173 in your browser.
