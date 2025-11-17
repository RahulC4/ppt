# ingestion_chroma.py
import os
from chromadb import PersistentClient
from openai import AzureOpenAI
from utils import get_env, logger

text_client = AzureOpenAI(
    azure_endpoint=get_env("OPENAI_API_BASE", required=True),
    api_key=get_env("OPENAI_API_KEY", required=True),
    api_version=get_env("OPENAI_API_VERSION", "2024-05-01-preview")
)

EMBEDDING_MODEL = get_env("EMBEDDING_MODEL", "text-embedding-3-small")
CHROMA_PERSIST_DIR = get_env("CHROMA_PERSIST_DIR", "./chroma_db")

chroma_client = PersistentClient(path=CHROMA_PERSIST_DIR)
collection = chroma_client.get_or_create_collection("ppt_slides")

def embed(text):
    resp = text_client.embeddings.create(
        model=EMBEDDING_MODEL,
        input=text,
        encoding_format="base64"
    )
    return resp.data[0].embedding

def add_slide(ppt_name, slide_id, text, tags=[]):
    try:
        emb = embed(text)
        collection.add(
            ids=[f"{ppt_name}_{slide_id}"],
            documents=[text],
            embeddings=[emb],
            metadatas=[{
                "ppt_name": ppt_name,
                "slide_id": slide_id,
                "tags": ",".join(tags)
            }]
        )
    except Exception as e:
        logger.error(f"Failed to add slide: {e}")
