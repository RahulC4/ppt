# azure_blob_utils.py

import os
from azure.storage.blob import BlobServiceClient
from utils import get_env, logger

# One connection string for the storage account
BLOB_CONN = get_env("AZURE_BLOB_CONN", required=True)

# Container where your SAMPLE / SOURCE PPTs live
SOURCE_CONTAINER = get_env("AZURE_BLOB_CONTAINER", "ppt-dataset")

# Container where GENERATED PPTs + logs are stored
GENERATED_CONTAINER = get_env("GENERATED_CONTAINER", "generated-presentations")


def _get_blob_service():
    return BlobServiceClient.from_connection_string(BLOB_CONN)


# ============================================================
# 1️⃣ GENERATED PPT + LOG UPLOADS (existing behavior)
# ============================================================

def upload_ppt_to_blob(file_path: str, file_name: str) -> str:
    """
    Upload a generated PPT file (from local path) into the GENERATED_CONTAINER.
    Returns '<container>/<blob_name>'.
    """
    blob_service = _get_blob_service()
    container_client = blob_service.get_container_client(GENERATED_CONTAINER)

    try:
        container_client.create_container()
    except Exception:
        # already exists
        pass

    with open(file_path, "rb") as data:
        container_client.upload_blob(name=file_name, data=data, overwrite=True)

    logger.info(f"Uploaded generated PPT to Azure Blob: {GENERATED_CONTAINER}/{file_name}")
    return f"{GENERATED_CONTAINER}/{file_name}"


def upload_json_to_blob(json_bytes: bytes, blob_name: str) -> str:
    """
    Upload a JSON log (bytes) into the GENERATED_CONTAINER.
    """
    blob_service = _get_blob_service()
    container_client = blob_service.get_container_client(GENERATED_CONTAINER)

    try:
        container_client.create_container()
    except Exception:
        pass

    container_client.upload_blob(name=blob_name, data=json_bytes, overwrite=True)
    logger.info(f"Uploaded log to Azure Blob: {GENERATED_CONTAINER}/{blob_name}")
    return f"{GENERATED_CONTAINER}/{blob_name}"


def list_generated_presentations():
    """
    List blobs in the GENERATED_CONTAINER.
    (You may or may not still use this; kept for compatibility.)
    """
    blob_service = _get_blob_service()
    container_client = blob_service.get_container_client(GENERATED_CONTAINER)
    try:
        return [b.name for b in container_client.list_blobs()]
    except Exception as e:
        logger.warning(f"Failed to list generated PPTs: {e}")
        return []


# ============================================================
# 2️⃣ SOURCE / SAMPLE PPT UPLOAD (Knowledge Base + Templates)
# ============================================================

def upload_source_ppt_to_blob(bytes_data: bytes, blob_name: str) -> str:
    """
    Upload a *source/sample* PPT directly from bytes into SOURCE_CONTAINER.
    This is what you call from Streamlit when the user uploads a PPT.
    """
    blob_service = _get_blob_service()
    container_client = blob_service.get_container_client(SOURCE_CONTAINER)

    try:
        container_client.create_container()
    except Exception:
        pass

    container_client.upload_blob(name=blob_name, data=bytes_data, overwrite=True)
    logger.info(f"Uploaded source PPT to Azure Blob: {SOURCE_CONTAINER}/{blob_name}")
    return f"{SOURCE_CONTAINER}/{blob_name}"


def list_source_ppts(only_pptx: bool = True):
    """
    List PPT files from SOURCE_CONTAINER (ppt-dataset).
    This is what you should use to populate the Visual Template dropdown.

    If only_pptx=True, filters to '.pptx' and '.ppt'.
    Returns a plain list of blob names.
    """
    blob_service = _get_blob_service()
    container_client = blob_service.get_container_client(SOURCE_CONTAINER)
    try:
        names = []
        for b in container_client.list_blobs():
            if not only_pptx:
                names.append(b.name)
            else:
                if b.name.lower().endswith((".pptx", ".ppt")):
                    names.append(b.name)
        return names
    except Exception as e:
        logger.warning(f"Failed to list source PPTs: {e}")
        return []
