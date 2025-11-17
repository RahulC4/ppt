# utils.py
import os
import json
import logging
from datetime import datetime
from dotenv import load_dotenv

load_dotenv()

LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO").upper()

logging.basicConfig(
    level=LOG_LEVEL,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s"
)
logger = logging.getLogger("ai-ppt-generator")

def get_env(name, default=None, required=False):
    val = os.getenv(name, default)
    if required and (val is None or val == ""):
        raise EnvironmentError(f"Missing required env var: {name}")
    return val

def now_ts():
    return datetime.utcnow().isoformat() + "Z"

def safe_json_load(s):
    if not s:
        return None
    s = s.strip()
    idx = min([i for i in [s.find("{"), s.find("[")] if i != -1], default=-1)
    if idx == -1:
        return None
    try:
        return json.loads(s[idx:])
    except:
        return None

def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)
