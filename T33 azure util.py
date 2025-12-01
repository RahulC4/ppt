# azure_blob_utils.py

import io
import os
import tempfile
from typing import List, Optional

from azure.storage.blob import BlobServiceClient, ContentSettings

from utils import get_env, logger

# ------------------------------------------------------------
# CONFIG – matches YOUR .env
# ------------------------------------------------------------

# Connection string
BLOB_CONN_STR = get_env("AZURE_BLOB_CONN", required=True)

# Container that holds SOURCE / SAMPLE PPTs (knowledge base + templates)
SOURCE_PPT_CONTAINER = get_env("AZURE_BLOB_CONTAINER", required=True)

# Container that holds GENERATED PPTs (output of the app)
PPT_OUTPUT_CONTAINER = get_env("GENERATED_CONTAINER", required=True)

_blob_service = BlobServiceClient.from_connection_string(BLOB_CONN_STR)


def _get_container_client(container_name: str):
    return _blob_service.get_container_client(container_name)


# ------------------------------------------------------------
# UPLOAD: GENERATED PPT + LOG JSON
# ------------------------------------------------------------

def upload_ppt_to_blob(local_path: str, blob_name: str) -> None:
    """
    Upload a generated PPT file to the GENERATED_CONTAINER.
    """
    container = _get_container_client(PPT_OUTPUT_CONTAINER)
    with open(local_path, "rb") as f:
        container.upload_blob(
            name=blob_name,
            data=f,
            overwrite=True,
            content_settings=ContentSettings(
                content_type=(
                    "application/vnd.openxmlformats-officedocument."
                    "presentationml.presentation"
                )
            ),
        )
    logger.info(f"Uploaded generated PPT to blob: {blob_name}")


def upload_json_to_blob(data: bytes, blob_name: str) -> None:
    """
    Upload JSON log bytes next to generated PPTs in GENERATED_CONTAINER.
    """
    container = _get_container_client(PPT_OUTPUT_CONTAINER)
    container.upload_blob(
        name=blob_name,
        data=data,
        overwrite=True,
        content_settings=ContentSettings(content_type="application/json"),
    )
    logger.info(f"Uploaded JSON log to blob: {blob_name}")


# ------------------------------------------------------------
# UPLOAD: SOURCE / SAMPLE PPTs (knowledge base + templates)
# ------------------------------------------------------------

def upload_source_ppt_to_blob(data: bytes, blob_name: str) -> None:
    """
    Upload a source/sample PPT into AZURE_BLOB_CONTAINER.
    These files are:
      • ingested into Chroma for semantic search
      • used as template candidates in the UI
    """
    container = _get_container_client(SOURCE_PPT_CONTAINER)
    container.upload_blob(
        name=blob_name,
        data=io.BytesIO(data),
        overwrite=True,
        content_settings=ContentSettings(
            content_type=(
                "application/vnd.openxmlformats-officedocument."
                "presentationml.presentation"
            )
        ),
    )
    logger.info(f"Uploaded source PPT to blob: {blob_name}")


# ------------------------------------------------------------
# LIST + DOWNLOAD SOURCE PPTs (for Visual Template dropdown)
# ------------------------------------------------------------

def list_source_ppt_blobs() -> List[str]:
    """
    Return a list of *.pptx blob names from AZURE_BLOB_CONTAINER.
    Used to populate the 'Visual Template' dropdown in the Streamlit UI.
    """
    try:
        container = _get_container_client(SOURCE_PPT_CONTAINER)
        names: List[str] = []
        for blob in container.list_blobs():
            if blob.name.lower().endswith(".pptx"):
                names.append(blob.name)
        logger.info(
            f"Found {len(names)} source PPTs in container '{SOURCE_PPT_CONTAINER}': {names}"
        )
        return names
    except Exception:
        logger.exception("Failed to list source PPT blobs")
        return []


def download_source_ppt_to_temp(blob_name: str) -> Optional[str]:
    """
    Download a source PPT from AZURE_BLOB_CONTAINER to a temporary local file
    and return the local file path. Returns None on error.
    """
    try:
        container = _get_container_client(SOURCE_PPT_CONTAINER)
        downloader = container.download_blob(blob_name)
        data = downloader.readall()

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
        tmp.write(data)
        tmp.close()
        logger.info(f"Downloaded source PPT '{blob_name}' to temp file: {tmp.name}")
        return tmp.name
    except Exception:
        logger.exception(f"Failed to download source PPT: {blob_name}")
        return None
