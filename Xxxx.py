# utils.py (client section)

from openai import AzureOpenAI
import os
import logging
from dotenv import load_dotenv

load_dotenv()

LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
logging.basicConfig(level=LOG_LEVEL, format="%(asctime)s | %(levelname)s | %(name)s | %(message)s")
logger = logging.getLogger("ai-ppt-generator")

def get_env(name, default=None, required=False):
    val = os.getenv(name, default)
    if required and (val is None or val == ""):
        logger.error(f"Missing required environment variable: {name}")
        raise EnvironmentError(f"Missing env var: {name}")
    return val

def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)

# -----------------------------
#  TEXT MODEL CLIENT  (GPT + EMBEDDINGS)
# -----------------------------
text_client = AzureOpenAI(
    azure_endpoint = get_env("OPENAI_API_BASE", required=True),
    api_key        = get_env("OPENAI_API_KEY", required=True),
    api_version    = get_env("OPENAI_API_VERSION", required=True)
)

# -----------------------------
#  IMAGE MODEL CLIENT (DALL·E / GPT-image)
# -----------------------------
image_client = AzureOpenAI(
    azure_endpoint = get_env("IMAGE_API_BASE", required=True),
    api_key        = get_env("IMAGE_API_KEY", required=True),
    api_version    = get_env("IMAGE_API_VERSION", required=True)
)
