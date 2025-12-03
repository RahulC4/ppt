from azure.storage.blob import BlobServiceClient
from utils import get_env

# Reuse same connection + container you already use for upload_source_ppt_to_blob
_BLOB_CONN_STR = get_env("AZURE_BLOB_CONNECTION_STRING")
_SOURCE_PPT_CONTAINER = get_env("SOURCE_PPT_CONTAINER", "source-ppts")

_blob_client = BlobServiceClient.from_connection_string(_BLOB_CONN_STR)
_source_container = _blob_client.get_container_client(_SOURCE_PPT_CONTAINER)


def upload_source_ppt_to_blob(bytes_data: bytes, blob_name: str) -> None:
    # If you already have this, keep your existing implementation
    _source_container.upload_blob(name=blob_name, data=bytes_data, overwrite=True)


def list_source_ppt_blobs() -> list[str]:
    """Return a list of blob names for all knowledge base PPTs."""
    return [b.name for b in _source_container.list_blobs()]


def delete_source_ppt_from_blob(blob_name: str) -> None:
    """Delete a PPT from the source PPT container."""
    _source_container.delete_blob(blob_name)



# Assuming you already have a global `collection` used in process_blob()
# and that you stored metadata like {"source": blob_name, ...}

def delete_ppt_from_chroma(blob_name: str) -> None:
    """
    Delete all Chroma indexes (documents) that came from the given PPT file.
    This assumes you stored the PPT name in metadata under key 'source'.
    """
    collection.delete(where={"source": blob_name})
