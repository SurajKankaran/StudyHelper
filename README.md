# StudyMate Local — no API key

A local document study assistant using text extraction, keyword frequency, sentence matching, and simple templates. It does not call OpenAI or any cloud AI service.

## Supported formats
PDF (selectable text), Word DOCX, PowerPoint PPTX, TXT, Markdown. Convert legacy `.doc` files to `.docx`. Scanned PDFs need OCR first.

## Install and run
Requires Python 3.10+.

macOS/Linux:
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Windows:
```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

Open the local URL shown in the terminal, usually http://localhost:8501.

## Features
- Upload and select multiple files
- Summaries and key terms
- Short questions with source-based answers
- Long questions with supporting passages
- MCQ quiz and answer keys
- Flashcards
- Important topics based on frequency
- Basic document Q&A by matching question keywords
- Download results as text files

## Limitations
This is not a generative AI model. Its rule-based answers can be repetitive, miss context, or be incomplete. For a more capable API-free chatbot, a local language model can be added later using software such as Ollama; this requires downloading a model and suitable computer resources.
