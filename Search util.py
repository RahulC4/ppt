# search_utils.py  (final)

import chromadb
from chromadb.config import Settings
import numpy as np
from utils import get_env, logger
from openai import AzureOpenAI

# --------------------------------------------
# Initialize Azure embedding client
# --------------------------------------------
embed_client = AzureOpenAI(
    azure_endpoint=get_env("OPENAI_API_BASE", required=True),
    api_key=get_env("OPENAI_API_KEY", required=True),
    api_version=get_env("OPENAI_API_VERSION", "2024-05-01-preview")
)

EMBED_MODEL = get_env("EMBEDDING_MODEL", "text-embedding-3-large")
EMBED_DIM = int(get_env("EMBEDDING_DIM", 1536))

# --------------------------------------------
# Initialize Chroma persistent DB
# --------------------------------------------
chroma_client = chromadb.Client(
    Settings(
        chroma_db_impl="duckdb+parquet",
        persist_directory=get_env("CHROMA_PERSIST_DIR", "./chroma_db")
    )
)

collection = chroma_client.get_collection(
    name="ppt_index",
    embedding_function=None  # we embed manually
)

# --------------------------------------------
# Embed text using Azure OpenAI
# --------------------------------------------
def embed_text(text: str):
    try:
        resp = embed_client.embeddings.create(
            input=text,
            model=EMBED_MODEL
        )
        vec = resp.data[0].embedding
        if len(vec) != EMBED_DIM:
            logger.warning(f"⚠ Embedding dimension mismatch: expected {EMBED_DIM}, got {len(vec)}")
        return vec
    except Exception as e:
        logger.error(f"Embedding error: {e}")
        return np.zeros(EMBED_DIM).tolist()


# --------------------------------------------
# Parse Chroma results safely
# --------------------------------------------
def parse_results(results):
    if not results or "ids" not in results:
        return []

    out = []
    for i in range(len(results["ids"][0])):
        md = results["metadatas"][0][i]
        doc = results["documents"][0][i]

        out.append({
            "ppt_name": md.get("ppt_name"),
            "slide_number": md.get("slide_number"),
            "text": doc,
            "tags": md.get("tags", [])
        })
    return out


# --------------------------------------------
# Semantic Search with tag filtering
# --------------------------------------------
def semantic_search(query: str, top_k: int = 5, tags: list = None):
    logger.info(f"Semantic search: '{query}', tags={tags}")

    # Step 1 - embed query
    emb = embed_text(query)

    # Step 2 - build WHERE clause properly
    where = {}
    if tags:
        where = {
            "$and": [
                {"tags": {"$contains": tags}}
            ]
        }

    # Step 3 - perform query
    try:
        results = collection.query(
            query_embeddings=[emb],
            n_results=top_k,
            where=where,
            include=["metadatas", "documents"]
        )
        parsed = parse_results(results)
        logger.info(f"Retrieved {len(parsed)} slides from Chroma.")
        return parsed

    except Exception as e:
        logger.error(f"Chroma query failed: {e}")
        return []
