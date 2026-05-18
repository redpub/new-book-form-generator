from pathlib import Path

# Project root (same directory as this file)
APP_DIR = Path(__file__).resolve().parent
PROMPT_FILE = APP_DIR / "prompt.txt"
TEMPLATE_FILE = APP_DIR / "template.docx"

ALLOWED_EMAIL_DOMAIN = "@red-publish.com"
STRAICO_ENDPOINT = "https://api.straico.com/v0/prompt/completion"
STRAICO_MODEL = "anthropic/claude-sonnet-4.5"
