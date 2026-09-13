# Sovereign AI Workbench — Prototype V2

## What changed
- Local Employee/Admin login with offline SQLite persistence.
- Company, sector and role are attached to the authenticated session.
- Persistent individual chat history and uploaded-file metadata.
- Team creation/join by invite code with same-company verification.
- Admin-only SOP upload with sector tagging and local KB rebuild.
- ChatGPT-style workspace with New Chat, history, Team Chats and file attachments.
- Uploaded files remain locally under `uploads/`; database remains `sovereign_local.db`.
- Deterministic CSV/Excel analysis: answers use real dataframe values.
- Deterministic multi-metric charts: column selection no longer depends on the LLM.
- Supports ratios such as Debt/Equity and Current Ratio when the source contains the relevant direct columns or numerator/denominator columns.
- PDF/DOCX/PPTX/MD/TXT/JSON/LOG extraction plus local vision analysis for images/scanned PDFs/P&IDs.
- Current prototype model routing matches the models used on the demo machine:
  - Text/general/engineering/finance: `llama3.2:3b`
  - Coding: `qwen2.5-coder:1.5b`
  - Vision: `llava:7b`

## Demo credentials
- Admin: `ADMIN001` / `admin123`
- Engineering: `EMP001` / `emp123`
- Finance: `EMP002` / `emp123`
- IT: `EMP003` / `emp123`
- Operations: `EMP004` / `emp123`

Change these before any real deployment.

## Run
```bat
cd /d E:\SIH_26_prototype
python -m pip install -r requirements.txt
python check_ollama.py
python build_kb.py
python app.py
```
Then open `http://127.0.0.1:8000`.

## Offline / air-gapped note
The application uses `localhost:11434` for Ollama and local SQLite/files. No cloud AI API is required. Initial package/model installation may require a connected staging machine; production air-gapped deployment should transfer approved packages/models into the isolated environment and block outbound traffic.

## Important prototype limitation
`ACTIVE_FILE_STORE` is still an in-process demo cache. SQLite persists users, chats and file metadata, but a production multi-user deployment should replace this cache with a per-user/per-chat server-side store and encrypted file storage. The UI and API structure are already separated so that can be added later.
