# Project Brain OS

Project Brain OS is a Google Workspace automation system that turns a project idea into an executable project plan using multiple AI Agents.

## Core stack

- Google Apps Script
- Google Sheets
- Google Docs
- Google Drive
- Telegram Bot
- Gemini 2.5 Flash API
- Claude API
- HTML dashboard

## Main flow

HTML Dashboard
→ Google Apps Script Web App
→ Google Sheet database
→ Agent_Tasks queue
→ Gemini/Claude
→ Google Docs output
→ Final_Output
→ Telegram notification

## Important files

- `apps-script/Code.gs`: main backend
- `apps-script/index.html`: Apps Script web interface
- `apps-script/appsscript.json`: Apps Script manifest
- `frontend/he-thong-tu-duy-agent-du-an.html`: local dashboard

## Security

Never commit real API keys.

Use placeholders only:

- `DAN_GEMINI_API_KEY_MOI_VAO_DAY`
- `DAN_CLAUDE_API_KEY_MOI_VAO_DAY`
- `DAN_TELEGRAM_BOT_TOKEN_MOI_VAO_DAY`
- `DAN_TELEGRAM_CHAT_ID_VAO_DAY`

Real keys must be stored in Google Apps Script Script Properties.
