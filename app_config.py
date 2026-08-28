from pathlib import Path

# Project root (same directory as this file)
APP_DIR = Path(__file__).resolve().parent
PROMPT_FILE = APP_DIR / "prompt.txt"
TEMPLATE_FILE = APP_DIR / "template-v2.docx"

ALLOWED_EMAIL_DOMAIN = "@red-publish.com"
VERTEX_MODEL = "gemini-3.5-flash-lite"
