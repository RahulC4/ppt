from openai import AzureOpenAI
import os
import json
import logging
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()
logging.basicConfig(level=LOG_LEVEL,
                    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s")
logger = logging.getLogger("ai-ppt-generator-chroma")


def ensure_dir(path):
    """Create directory if it does not exist."""
    os.makedirs(path, exist_ok=True)


def get_env(name, default=None, required=False):
    val = os.getenv(name, default)
    if required and (val is None or val == ""):
        logger.error(f"Missing required environment variable: {name}")
        raise EnvironmentError(f"Missing env var: {name}")
    return val


def safe_json_load(s):
    if not s:
        return None
    s = s.strip()
    starts = [s.find(ch) for ch in ['{', '['] if s.find(ch) != -1]
    if not starts:
        return None
    try:
        return json.loads(s[min(starts):])
    except Exception as e:
        logger.warning(f"safe_json_load failed: {e}")
        return None


def now_ts():
    return datetime.utcnow().isoformat() + "Z"


def get_embedding_dim(model_name):
    try:
        return int(get_env("EMBEDDING_DIM", 1536))
    except:
        return 1536
    

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
    api_version    = get_env("OPENAI_API_VERSION", required=True)
)
