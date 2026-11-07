# PM SHRI KV No. 2 Armapur – Smart Classroom Teaching–Learning Portal

A Python + Streamlit + SQLite prototype with a persistent school identity on every application page.

## Implemented features
- Teacher and student sign-in; password hashes use PBKDF2-HMAC-SHA256 with per-user salts.
- First-run setup creates the initial teacher account (no default password).
- Teacher-created student/teacher accounts and student class/ID fields.
- Student-wise quiz attempt tracking and topic-wise learning-gap analytics.
- Teacher question bank with MCQ, short-answer and case-based entries.
- Editable DOCX question paper generation with school name in document header/footer.
- English/Hindi navigation labels.
- AI doubt assistant through a configurable OpenAI-compatible Chat Completions endpoint.
- Teacher doubt responses, learning resources, inclusive reading display.
- Docker deployment files and persistent SQLite volume guidance.

## Local installation (Windows)
1. Install Python 3.10 or later.
2. Extract the ZIP and open Command Prompt in this folder.
3. Run:
   ```bash
   py -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   streamlit run app.py
   ```
4. Open the local URL displayed in the terminal (normally http://localhost:8501).
5. On first launch, create the initial teacher account. Use a strong password (10+ characters, uppercase, lowercase and a number).
6. Sign in as teacher and create student accounts. Provide temporary passwords privately.

## AI assistant setup
The AI module is optional. Set `AI_API_KEY`, `AI_API_URL` and `AI_MODEL` as environment variables in the hosting provider. Never commit secrets to GitHub or paste keys into app.py. The assistant sends the student's question, class, subject and topic to the configured provider. Do not send names, roll numbers or sensitive student data. Verify AI-generated answers before classroom use.

If no AI key is configured, the rest of the portal continues to work and the AI module displays a configuration message.

## Docker deployment
Build:
```bash
docker build -t kv2-smart-classroom .
```
Run with persistent database storage:
```bash
docker run -d --name kv2-smart-classroom -p 8501:8501 -v kv2-classroom-data:/app/data kv2-smart-classroom
```
For school use, place the service behind HTTPS, restrict network access, maintain backups and obtain school/IT approval. Do not expose an unauthenticated or development server directly to the public internet.

## Suggested hosting workflow
- Push this folder to a private GitHub repository.
- Deploy the repository to a school-approved Docker-capable host.
- Configure persistent storage mounted at `/app/data`.
- Add AI environment secrets only if the school has approved an AI provider and its data-processing terms.
- Use a domain and HTTPS reverse proxy if public/remote access is required.

## Security and operational notes
- Passwords are not stored in plain text. However, this is a prototype and has not undergone independent security testing.
- Teacher role currently grants teacher-level functions to all teacher accounts. The first teacher is not a separate super-admin.
- There is no self-service password reset, email verification, CAPTCHA, audit log or account lockout yet. The school administrator must manage accounts and backups.
- For student privacy, use only minimum necessary data, student codes where practical, and school-approved retention rules.
- Analytics are based on submitted quiz attempts. Descriptive answers are not automatically graded; teacher review is required.
- Question paper generator selects available matching questions up to the chosen marks. It does not certify CBSE blueprint balance or guarantee exact section-wise distribution; teacher must review the generated paper before use.
- Hindi/English translation covers the interface labels; teacher-authored learning resources and questions remain in the language entered.
- The inclusive learning page offers readable display settings; test screen-reader/keyboard accessibility separately.

## Files
- `app.py` – complete application
- `requirements.txt` – Python dependencies
- `Project_Implementation_and_User_Manual.docx` – project documentation
- `Dockerfile`, `.streamlit/config.toml` – deployment setup
- `.env.example` – optional AI and database environment variables
