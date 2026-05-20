# SmartPDF AI Assistant

Offline PDF assistant built with Streamlit, SQLAlchemy, local document search, study tools, exports, and optional cloud mirroring.

## Features

- Login and signup system
- Dark and light mode UI
- ChatGPT-style multi-user chat workspace
- Multiple PDF upload with extracted text storage
- Offline answers restricted to uploaded PDFs only
- Automatic `Information not found in uploaded document` fallback
- Voice question recording with offline transcription
- Voice output for answers
- PDF summary, smart notes, important questions, and MCQ generation
- Page-based smart search with page number results
- Multi-language support for English, Telugu, and Hindi
- Chat export as PDF, TXT, and DOCX
- Clear current chat or remove all saved chats
- Analytics dashboard for PDFs, pages, chats, and questions
- Optional cloud PDF mirroring through `CLOUD_STORAGE_DIR`
- Persistent users, documents, chats, and messages
- SQLite for local development and PostgreSQL via `DATABASE_URL` for deployment

## Configuration

No API keys or backend secrets are required.

Optional database setting:

```bash
DATABASE_URL=postgresql+psycopg://user:password@host:5432/smartpdf
CLOUD_STORAGE_DIR=/path/to/cloud/mirror
```

## Local run

Install dependencies:

```powershell
pip install -r requirements.txt
```

Start with the included launcher:

```powershell
.\run_app.ps1
```

Or run Streamlit directly:

```powershell
streamlit run app.py --server.address 0.0.0.0 --server.port 8501
```

## How it works

- PDFs are parsed and split into searchable text chunks.
- Questions are matched against local PDF text only, with page-aware results.
- The app returns the most relevant matching excerpts.
- Study tools generate summaries, notes, questions, and MCQs from the extracted text.
- If nothing relevant is found, it returns `Information not found in uploaded document`.

## Notes

- Voice transcription runs offline through `SpeechRecognition` with `pocketsphinx`.
- DOCX export uses `python-docx`.
- Existing chat history, saved documents, and database records are preserved.
