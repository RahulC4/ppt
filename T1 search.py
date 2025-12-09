import os
from openai import AzureOpenAI
from chromadb import PersistentClient
from utils import get_env, logger, get_embedding_dim

# === TEXT client (GPT + embeddings) ===
text_client = AzureOpenAI(
    azure_endpoint=get_env("OPENAI_API_BASE", required=True),
    api_key=get_env("OPENAI_API_KEY", required=True),
    api_version=get_env("OPENAI_API_VERSION", "2024-05-01-preview")
)

EMBEDDING_MODEL = get_env("EMBEDDING_MODEL", "text-embedding-3-large")
EMBEDDING_DIM = get_embedding_dim(EMBEDDING_MODEL)
CHROMA_PERSIST_DIR = get_env("CHROMA_PERSIST_DIR", "./chroma_db")


# === Chroma Initialization (Safe) ===
chroma_client = PersistentClient(path=CHROMA_PERSIST_DIR)

try:
    collection = chroma_client.get_collection("ppt_slides")
except Exception:
    collection = chroma_client.create_collection("ppt_slides")


# ------------------------------------------------------------
# GENERATE EMBEDDING
# ------------------------------------------------------------
def get_embedding(text):
    try:
        resp = text_client.embeddings.create(
            model=EMBEDDING_MODEL,
            input=text
        )
        return resp.data[0].embedding
    except Exception as e:
        logger.exception(f"Embedding failed: {e}")
        return None


# ------------------------------------------------------------
# SEMANTIC SEARCH (Chroma-compatible filtering)
# ------------------------------------------------------------
def semantic_search(query, top_k=5, tags=None):
    emb = get_embedding(query)
    if emb is None:
        return []

    # Chroma DOES NOT support $or, $contains anymore.
    # It only supports simple equality match:
    # where = { "tag": "value" }
    filters = None
    if tags and len(tags) > 0:
        filters = {"tags": tags[0]}   # pick first tag for filtering

    try:
        if filters:
            res = collection.query(
                query_embeddings=[emb],
                n_results=top_k,
                where=filters
            )
        else:
            res = collection.query(
                query_embeddings=[emb],
                n_results=top_k
            )

        ids = res.get("ids", [[]])[0]
        metas = res.get("metadatas", [[]])[0]
        docs = res.get("documents", [[]])[0]
        dists = res.get("distances", [[]])[0]

        out = []
        for i in range(len(ids)):
            out.append({
                "id": ids[i],
                "ppt_name": metas[i].get("ppt_name"),
                "slide_id": metas[i].get("slide_id"),
                "title": metas[i].get("title"),
                "text": docs[i],
                "tags": metas[i].get("tags"),
                "score": dists[i]
            })

        return out

    except Exception as e:
        logger.exception(f"Chroma query failed: {e}")
        return []
