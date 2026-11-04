import os
from dotenv import load_dotenv

load_dotenv()

ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
LIBREOFFICE_PATH = os.getenv("LIBREOFFICE_PATH", "soffice")
MAX_UPLOAD_MB = int(os.getenv("MAX_UPLOAD_MB", "50"))
CONVERSION_TIMEOUT_SECONDS = int(os.getenv("CONVERSION_TIMEOUT_SECONDS", "120"))

# Single place to change the model name (Product Rule 13)
CLAUDE_MODEL = "claude-opus-4-8"
