
# ============================================================
# azure_blob_utils.py
# Central helpers for working with Azure Blob Storage
# ============================================================

import os
import tempfile
from typing import Optional, List

from azure.storage.blob import BlobServiceClient

from utils import get_env, logger


# ------------------------------------------------------------
# CONFIG / CONNECTION HELPERS
# ------------------------------------------------------------

def _get_blob_service_client() -> BlobServiceClient:
    """
    Build a BlobServiceClient from the connection string in env.
    Adjust AZURE_STORAGE_CONNECTION_STRING if you use a different name.
    """
    conn_str = get_env("AZURE_STORAGE_CONNECTION_STRING", required=True)
    return BlobServiceClient.from_connection_string(conn_str)


def _get_container_client(container_name: str):
    bsc = _get_blob_service_client()
    return bsc.get_container_client(container_name)


# These env-var names are just conventions.
# If you only use one container, you can point all three env vars
# to the same container name.
SOURCE_PPT_CONTAINER_ENV = "SOURCE_PPT_CONTAINER"      # for uploaded / sample PPTs
GENERATED_PPT_CONTAINER_ENV = "GENERATED_PPT_CONTAINER"  # for generated decks
JSON_LOG_CONTAINER_ENV = "JSON_LOG_CONTAINER"          # for logs (JSON)


def _get_source_container_name() -> str:
    return get_env(SOURCE_PPT_CONTAINER_ENV, required=True)


def _get_generated_container_name() -> str:
    # You can set this equal to SOURCE_PPT_CONTAINER if you prefer a single container
    return get_env(GENERATED_PPT_CONTAINER_ENV, required=True)


def _get_log_container_name() -> str:
    # You can set this equal to GENERATED_PPT_CONTAINER or SOURCE_PPT_CONTAINER
    return get_env(JSON_LOG_CONTAINER_ENV, required=True)


# ------------------------------------------------------------
# UPLOAD HELPERS (YOU WERE ALREADY USING THESE)
# ------------------------------------------------------------

def upload_source_ppt_to_blob(file_bytes: bytes, blob_name: str) -> None:
    """
    Upload a *sample* PPT (used as knowledge base / templates)
    into the SOURCE_PPT_CONTAINER.
    """
    container_name = _get_source_container_name()
    container_client = _get_container_client(container_name)

    logger.info(f"Uploading sample PPT '{blob_name}' to container '{container_name}'")
    try:
        blob_client = container_client.get_blob_client(blob_name)
        blob_client.upload_blob(file_bytes, overwrite=True)
    except Exception:
        logger.exception(f"Failed to upload source PPT '{blob_name}'")
        raise


def upload_ppt_to_blob(local_path: str, blob_name: Optional[str] = None) -> str:
    """
    Upload a *generated* PPT from a local path to GENERATED_PPT_CONTAINER.
    Returns the blob name that was used.

    If blob_name is None, a name is derived from the local filename.
    """
    container_name = _get_generated_container_name()
    container_client = _get_container_client(container_name)

    if blob_name is None:
        blob_name = os.path.basename(local_path)

    logger.info(f"Uploading generated PPT '{local_path}' as '{blob_name}' "
                f"to container '{container_name}'")
    try:
        with open(local_path, "rb") as f:
            data = f.read()
        blob_client = container_client.get_blob_client(blob_name)
        blob_client.upload_blob(data, overwrite=True)
        return blob_name
    except Exception:
        logger.exception(f"Failed to upload generated PPT '{blob_name}'")
        raise


def upload_json_to_blob(json_bytes: bytes, blob_name: str) -> None:
    """
    Upload a JSON log or metadata file to JSON_LOG_CONTAINER.
    """
    container_name = _get_log_container_name()
    container_client = _get_container_client(container_name)

    logger.info(f"Uploading JSON log '{blob_name}' to container '{container_name}'")
    try:
        blob_client = container_client.get_blob_client(blob_name)
        blob_client.upload_blob(json_bytes, overwrite=True)
    except Exception:
        logger.exception(f"Failed to upload JSON log '{blob_name}'")
        raise


# ------------------------------------------------------------
# ✨ NEW: LIST + DOWNLOAD SAMPLE PPT TEMPLATES
# ------------------------------------------------------------

def list_source_ppt_blobs() -> List[str]:
    """
    Return a sorted list of PPTX blob names from the *source/sample* container.

    Used to populate the Visual Template dropdown in the UI.
    """
    try:
        container_name = _get_source_container_name()
        container_client = _get_container_client(container_name)

        names: List[str] = []
        for blob in container_client.list_blobs():
            if blob.name.lower().endswith(".pptx"):
                names.append(blob.name)

        names.sort()
        logger.info(
            f"Found {len(names)} PPT templates in container '{container_name}'"
        )
        return names
    except Exception:
        logger.exception("Failed to list source PPT blobs")
        return []


def download_source_ppt_to_temp(blob_name: str) -> str:
    """
    Download a PPT from the *source/sample* container into a temp file.
    Returns the local temp file path.

    This is used by generate_ppt.py as the visual template.
    """
    container_name = _get_source_container_name()
    container_client = _get_container_client(container_name)
    blob_client = container_client.get_blob_client(blob_name)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    try:
        logger.info(
            f"Downloading template PPT '{blob_name}' from container '{container_name}'"
        )
        data = blob_client.download_blob().readall()
        tmp.write(data)
        tmp.close()
        return tmp.name
    except Exception:
        tmp.close()
        logger.exception(
            f"Failed to download template PPT '{blob_name}' "
            f"from container '{container_name}'"
        )
        raise
