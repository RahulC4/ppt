import os
from azure.storage.blob import BlobServiceClient
from utils import get_env, logger

BLOB_CONN = get_env("AZURE_BLOB_CONN", required=True)

# Container for generated PPTs
GENERATED_CONTAINER = get_env("GENERATED_CONTAINER", "generated-presentations")

# Container for source dataset PPTs (your existing ppt-dataset)
SOURCE_CONTAINER = get_env("AZURE_BLOB_CONTAINER", "ppt-dataset")


def _get_container_client(container_name: str):
    blob_service = BlobServiceClient.from_connection_string(BLOB_CONN)
    container_client = blob_service.get_container_client(container_name)
    try:
        container_client.create_container()
    except Exception:
        # already exists
        pass
    return container_client


# ----------------------------
# GENERATED PPT UPLOAD
# ----------------------------
def upload_ppt_to_blob(file_path, file_name):
    container_client = _get_container_client(GENERATED_CONTAINER)
    with open(file_path, "rb") as data:
        container_client.upload_blob(name=file_name, data=data, overwrite=True)
    logger.info(f"Uploaded generated PPT to Azure Blob: {GENERATED_CONTAINER}/{file_name}")
    return f"{GENERATED_CONTAINER}/{file_name}"


def upload_json_to_blob(json_bytes, blob_name):
    container_client = _get_container_client(GENERATED_CONTAINER)
    container_client.upload_blob(name=blob_name, data=json_bytes, overwrite=True)
    logger.info(f"Uploaded log to Azure Blob: {GENERATED_CONTAINER}/{blob_name}")
    return f"{GENERATED_CONTAINER}/{blob_name}"


def list_generated_presentations():
    try:
        container_client = _get_container_client(GENERATED_CONTAINER)
        return [b.name for b in container_client.list_blobs()]
    except Exception as e:
        logger.warning(f"Failed to list generated PPTs: {e}")
        return []


# ----------------------------
# SOURCE PPT UPLOAD (DATASET)
# ----------------------------
def upload_source_ppt_to_blob(file_bytes, blob_name: str):
    """
    Upload a user-provided sample PPT into the dataset container (ppt-dataset).
    file_bytes: bytes from uploaded file.
    blob_name: key to store under, usually original filename.
    """
    container_client = _get_container_client(SOURCE_CONTAINER)
    container_client.upload_blob(name=blob_name, data=file_bytes, overwrite=True)
    logger.info(f"Uploaded SOURCE PPT to Azure Blob: {SOURCE_CONTAINER}/{blob_name}")
    return f"{SOURCE_CONTAINER}/{blob_name}"


# ----------------------------
# SOURCE PPT LIST + DELETE (UI SUPPORT)
# ----------------------------

def list_source_ppt_blobs():
    """
    List all source PPT files stored in the dataset container (ppt-dataset).
    Used by UI to show available templates.
    """
    try:
        container_client = _get_container_client(SOURCE_CONTAINER)
        return [b.name for b in container_client.list_blobs() if b.name.lower().endswith(".pptx")]
    except Exception as e:
        logger.warning(f"Failed to list source PPTs: {e}")
        return []


def delete_source_ppt_from_blob(blob_name: str):
    """
    Delete a source PPT from the dataset container (ppt-dataset).
    This is triggered when user removes a template from the UI.
    """
    try:
        container_client = _get_container_client(SOURCE_CONTAINER)
        container_client.delete_blob(blob_name)
        logger.info(f"Deleted SOURCE PPT from Azure Blob: {SOURCE_CONTAINER}/{blob_name}")
    except Exception as e:
        logger.exception(f"Failed to delete SOURCE PPT from Azure Blob: {blob_name}")
        raise e
